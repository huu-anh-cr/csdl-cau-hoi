import { NextResponse } from "next/server";
import { randomInt } from "node:crypto";
import { adminDb } from "@/lib/firebase-admin";
import { FieldValue } from "firebase-admin/firestore";

export const runtime = "nodejs";
// Không cache kết quả của GET (danh sách mã thay đổi liên tục)
export const dynamic = "force-dynamic";

const LEVELS = ["ez", "md", "hd"];
const CODE_LENGTH = 6;
const CODE_ALPHABET = "ABCDEFGHJKLMNPQRSTUVWXYZ23456789";
const MAX_BODY_BYTES = 1_000_000;
const MAX_SHEET_BYTES = 5_000_000;
const SHEET_TIMEOUT_MS = 4500;
const MAX_QUESTIONS = 50;
const RECENT_LIMIT = 10; // Số mã gần nhất trả về cho bảng
const PREVIEW_WORDS = 5; // Số chữ đầu của câu hỏi đầu tiên

// Tài liệu tổng hợp lưu 10 mã gần nhất: meta/recent -> { items: [...] }
const recentRef = () => adminDb.collection("meta").doc("recent");

class ApiError extends Error {
  constructor(status, message) {
    super(message);
    this.status = status;
  }
}

function getClientIp(request) {
  return request.headers.get("x-forwarded-for")?.split(",")[0] || "unknown";
}

function errorResponse(error) {
  const status = error.status || 500;
  const msg = error.status ? error.message : "Lỗi máy chủ nội bộ.";
  console.error("API Error:", error);
  return NextResponse.json({ error: msg }, { status });
}

// --- RATE LIMIT (In-memory, Best Effort) ---
const rateBuckets = new Map();
function checkRateLimit(key) {
  const now = Date.now();
  const bucket = rateBuckets.get(key);
  if (!bucket || bucket.resetAt <= now) {
    rateBuckets.set(key, { count: 1, resetAt: now + 60_000 });
    return false;
  }
  bucket.count++;
  return bucket.count > 15; // Tối đa 15 request / phút / IP (mỗi loại POST/GET)
}

// --- GOOGLE SHEET FETCH (CÓ PROXY DỰ PHÒNG) ---
async function fetchWithTimeout(url) {
  const response = await fetch(url, {
    cache: "no-store",
    signal: AbortSignal.timeout(SHEET_TIMEOUT_MS),
  });
  if (!response.ok) throw new Error("HTTP Error");

  const text = await response.text();
  if (text.length > MAX_SHEET_BYTES) throw new ApiError(413, "Sheet quá lớn!");
  return text;
}

async function fetchGoogleSheetSafe(id, gid) {
  const googleUrl = `https://docs.google.com/spreadsheets/d/${id}/gviz/tq?tqx=out:json${gid ? `&gid=${gid}` : ""}`;
  const proxyUrl = `https://api.allorigins.win/raw?url=${encodeURIComponent(googleUrl)}`;
  let text = "";

  try {
    // Phase 1: Thử gọi trực tiếp
    text = await fetchWithTimeout(googleUrl);
  } catch (err) {
    try {
      // Phase 2: Dự phòng dùng Proxy
      console.warn("Direct fetch failed, using proxy...");
      text = await fetchWithTimeout(proxyUrl);
    } catch (proxyErr) {
      throw new ApiError(
        504,
        "Google Sheets phản hồi quá chậm hoặc chặn kết nối. Hãy thử lại.",
      );
    }
  }

  const start = text.indexOf("{");
  const end = text.lastIndexOf("}");
  if (start === -1 || end <= start)
    throw new ApiError(
      400,
      'Lỗi đọc dữ liệu Sheet. Hãy chắc chắn link để chế độ "Bất kỳ ai".',
    );

  const result = JSON.parse(text.slice(start, end + 1));
  if (result?.status === "error" || !Array.isArray(result?.table?.rows))
    throw new ApiError(400, "Excel sai mẫu.");
  return result.table;
}

// --- LOGIC RÚT GỌN (MINIFY) ---
const fold = (t) =>
  t
    .toLowerCase()
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .replace(/đ/g, "d")
    .trim();
const parseType = (t) => (fold(t) === "2" || fold(t).includes("dien") ? 2 : 1);
const parseLevel = (t) => {
  const s = fold(t);
  if (s === "2" || s === "md" || s.includes("trung binh")) return "md";
  if (s === "3" || s === "hd" || /\bkho\b/.test(s)) return "hd";
  return "ez";
};
const cellText = (cells, i) => {
  const v = cells[i]?.f ?? cells[i]?.v;
  return v == null ? "" : String(v).normalize("NFC").trim().substring(0, 2000);
};

function buildBankFromSheet(table) {
  const bank = { ez: [], md: [], hd: [] };

  table.rows.forEach((row, index) => {
    const cells = row?.c;
    if (!cells) return;

    // Lấy dữ liệu từ index 1 (Cột B) đến index 7 (Cột H)
    const [tTxt, lTxt, q, ...ans] = [1, 2, 3, 4, 5, 6, 7].map((i) =>
      cellText(cells, i),
    );

    // Bỏ qua các dòng chú thích hoặc dòng tiêu đề (có chữ 'noi dung')
    if (fold(q).includes("noi dung") || fold(q).includes("luu y")) return;
    if (!q || !ans[0]) return;

    const t = parseType(tTxt);
    bank[parseLevel(lTxt)].push({
      t,
      q,
      a: t === 1 ? ans.filter(Boolean) : [ans[0]],
    });
  });
  return bank;
}

// --- LẤY NỘI DUNG XEM TRƯỚC CHO BẢNG ---
// Dữ liệu được lưu tách theo mức độ nên "câu hỏi đầu tiên" là câu đầu của
// mức Dễ, nếu không có thì lấy mức Trung bình, rồi mức Khó.
function getFirstQuestion(bank) {
  for (const level of LEVELS) {
    const q = bank?.[level]?.[0]?.q;
    if (q) return String(q);
  }
  return "";
}

const firstWords = (text, n = PREVIEW_WORDS) =>
  text.trim().split(/\s+/).filter(Boolean).slice(0, n).join(" ");

// --- FIRESTORE TRANSACTION (ADMIN SDK) ---
async function createPuzzleAdmin(bank) {
  const preview = firstWords(getFirstQuestion(bank));

  for (let i = 0; i < 5; i++) {
    let code = "";
    for (let j = 0; j < CODE_LENGTH; j++)
      code += CODE_ALPHABET[randomInt(CODE_ALPHABET.length)];

    const ref = adminDb.collection("puzzles").doc(code);
    const metaRef = recentRef();
    const created = await adminDb.runTransaction(async (t) => {
      // Firestore yêu cầu: đọc hết trước, ghi sau
      const [doc, metaDoc] = await Promise.all([t.get(ref), t.get(metaRef)]);
      if (doc.exists) return false;

      // serverTimestamp() không dùng được bên trong mảng nên dùng chuỗi ISO
      const entry = { code, createdAt: new Date().toISOString(), preview };
      const old = metaDoc.exists ? (metaDoc.data().items ?? []) : [];

      t.set(ref, {
        data: bank,
        createdAt: FieldValue.serverTimestamp(),
        views: 0,
      });
      t.set(metaRef, { items: [entry, ...old].slice(0, RECENT_LIMIT) });
      return true;
    });
    if (created) return code;
  }
  throw new ApiError(503, "Không thể tạo mã lúc này. Thử lại sau.");
}

// Chạy một lần duy nhất khi meta/recent chưa tồn tại (các mã tạo trước đây)
async function backfillRecent() {
  const snapshot = await adminDb
    .collection("puzzles")
    .orderBy("createdAt", "desc")
    .limit(RECENT_LIMIT)
    .get();

  const items = snapshot.docs.map((doc) => {
    const d = doc.data();
    return {
      code: doc.id,
      createdAt: d.createdAt?.toDate?.().toISOString() ?? null,
      preview: firstWords(getFirstQuestion(d.data)),
    };
  });

  if (items.length > 0) {
    try {
      // create() báo lỗi nếu tài liệu đã có (vừa có mã mới tạo cùng lúc) -> bỏ qua
      await recentRef().create({ items });
    } catch (e) {}
  }
  return items;
}

// --- GET: DANH SÁCH MÃ GẦN ĐÂY ---
export async function GET(request) {
  try {
    if (checkRateLimit(`list:${getClientIp(request)}`))
      throw new ApiError(429, "Bạn thao tác quá nhanh, vui lòng chờ 1 phút.");

    // Bình thường chỉ tốn 1 lượt đọc
    const metaDoc = await recentRef().get();
    const items = metaDoc.exists
      ? (metaDoc.data().items ?? [])
      : await backfillRecent();

    return NextResponse.json(
      { items },
      { headers: { "Cache-Control": "no-store" } },
    );
  } catch (error) {
    return errorResponse(error);
  }
}

// --- POST: TẠO MÃ MỚI ---
export async function POST(request) {
  try {
    if (checkRateLimit(`post:${getClientIp(request)}`))
      throw new ApiError(429, "Bạn thao tác quá nhanh, vui lòng chờ 1 phút.");

    const bodyLen = Number(request.headers.get("content-length"));
    if (bodyLen > MAX_BODY_BYTES) throw new ApiError(413, "File quá nặng!");
    const body = await request.json();

    let bank;
    if (body.sheetLink) {
      const id = body.sheetLink.match(/spreadsheets\/d\/([a-zA-Z0-9_-]+)/)?.[1];
      const gid = body.sheetLink.match(/[#&?]gid=(\d+)/)?.[1];
      if (!id) throw new ApiError(400, "Link Google Sheet không hợp lệ!");
      bank = buildBankFromSheet(await fetchGoogleSheetSafe(id, gid));
    } else if (body.fileData) {
      bank = body.fileData;
    } else {
      throw new ApiError(400, "Yêu cầu không hợp lệ.");
    }

    // --- BỘ LUẬT KIỂM TRA SỐ LƯỢNG (21 - 50, ít nhất 7 mỗi loại) ---
    const countEz = bank.ez.length;
    const countMd = bank.md.length;
    const countHd = bank.hd.length;
    const total = countEz + countMd + countHd;

    if (total < 21) {
      throw new ApiError(
        400,
        `Cần tối thiểu 21 câu hỏi. Hiện tại file của bạn chỉ có ${total} câu hợp lệ.`,
      );
    }
    if (total > MAX_QUESTIONS) {
      throw new ApiError(
        400,
        `Tối đa ${MAX_QUESTIONS} câu hỏi. File của bạn đang có ${total} câu.`,
      );
    }
    if (countEz < 7 || countMd < 7 || countHd < 7) {
      throw new ApiError(
        400,
        `Cần ít nhất 7 câu cho mỗi mức độ. Hiện tại: Dễ (${countEz}), Trung bình (${countMd}), Khó (${countHd}).`,
      );
    }

    const code = await createPuzzleAdmin(bank);
    return NextResponse.json({ success: true, code, total });
  } catch (error) {
    return errorResponse(error);
  }
}
