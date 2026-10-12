// app/page.js
"use client";

import { useState, useRef, useEffect, useCallback } from "react";
import * as XLSX from "xlsx";

// ----- Cấu hình cột dựa trên logic backend (Index 1 = Cột B) -----
const COL = {
  TYPE: 1, // Cột B
  DIFFICULTY: 2, // Cột C
  QUESTION: 3, // Cột D
  ANSWER_1: 4, // Cột E
  ANSWER_2: 5, // Cột F
  ANSWER_3: 6, // Cột G
  ANSWER_4: 7, // Cột H
};
const HEADER_ROWS = 2; // Bỏ qua 2 dòng tiêu đề
const MAX_FILE_SIZE = 10 * 1024 * 1024; // 10MB
const RELOAD_COOLDOWN_SECONDS = 20; // Thời gian chờ giữa 2 lần bấm nút tải lại
const PREVIEW_WORDS = 5; // Số chữ đầu tiên của câu hỏi đầu tiên

// Helper functions đồng bộ với Backend
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

const isValidSheetLink = (link) =>
  /^https:\/\/docs\.google\.com\/spreadsheets\/d\/[a-zA-Z0-9_-]+/i.test(link);

// Lấy N chữ đầu tiên của một đoạn văn bản
const firstWords = (text, n = PREVIEW_WORDS) =>
  (text || "")
    .toString()
    .trim()
    .split(/\s+/)
    .filter(Boolean)
    .slice(0, n)
    .join(" ");

// Chuyển giá trị thời gian (ISO string, timestamp, hoặc {seconds}) thành chuỗi hiển thị
const formatTime = (value) => {
  if (value === null || value === undefined || value === "") return "—";
  let date;
  if (typeof value === "object") {
    const sec = value.seconds ?? value._seconds;
    date = sec != null ? new Date(sec * 1000) : new Date(value);
  } else {
    date = new Date(value);
  }
  if (isNaN(date.getTime())) return "—";
  return date.toLocaleString("vi-VN", { hour12: false });
};

const readFileAsArrayBuffer = (file) => {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = (e) => resolve(e.target.result);
    reader.onerror = () =>
      reject(new Error("Không thể đọc file. Vui lòng thử lại."));
    reader.readAsArrayBuffer(file);
  });
};

const minifyData = (rawData) => {
  const bank = { ez: [], md: [], hd: [] };

  for (let i = HEADER_ROWS; i < rawData.length; i++) {
    const row = rawData[i];
    if (!row || row.length === 0) continue;

    const tTxt = (row[COL.TYPE] || "").toString();
    const lTxt = (row[COL.DIFFICULTY] || "").toString();
    const q = (row[COL.QUESTION] || "").toString().normalize("NFC").trim();

    // Logic loại trừ giống backend
    if (fold(q).includes("noi dung") || fold(q).includes("luu y")) continue;

    const ans1 = (row[COL.ANSWER_1] || "").toString().normalize("NFC").trim();
    if (!q || !ans1) continue;

    const t = parseType(tTxt);
    const answers = [ans1];

    if (t === 1) {
      [COL.ANSWER_2, COL.ANSWER_3, COL.ANSWER_4].forEach((col) => {
        if (row[col]) answers.push(row[col].toString().normalize("NFC").trim());
      });
    }

    const questionObj = { t, q, a: answers };
    bank[parseLevel(lTxt)].push(questionObj);
  }
  return bank;
};

export default function Page() {
  const [uploadType, setUploadType] = useState("file");
  const [file, setFile] = useState(null);
  const [sheetLink, setSheetLink] = useState("");

  const [isSubmitting, setIsSubmitting] = useState(false);
  const [btnText, setBtnText] = useState("Tạo Mã Dữ Liệu");

  const [error, setError] = useState(null);
  const [result, setResult] = useState(null);
  const [copyText, setCopyText] = useState("Sao chép");

  // ----- Bảng "các mã đã được tạo gần đây" -----
  const [recent, setRecent] = useState([]);
  const [isLoadingList, setIsLoadingList] = useState(true);
  const [listError, setListError] = useState(null);
  const [cooldown, setCooldown] = useState(0); // Số giây còn lại trước khi bấm tải lại được

  const fileInputRef = useRef(null);
  const listRequestId = useRef(0); // Chỉ nhận kết quả của lần gọi mới nhất

  // Tải 10 mã gần nhất từ server
  const loadRecent = useCallback(async () => {
    const requestId = ++listRequestId.current;
    setIsLoadingList(true);
    setListError(null);

    try {
      const res = await fetch("/api/save-puzzle", { cache: "no-store" });
      const data = await res.json();
      if (!res.ok)
        throw new Error(data.error || `Lỗi server (mã ${res.status}).`);

      const items = Array.isArray(data) ? data : data.items;
      if (!Array.isArray(items))
        throw new Error("Server trả về danh sách không hợp lệ.");

      if (requestId === listRequestId.current) setRecent(items.slice(0, 10));
    } catch (err) {
      if (requestId === listRequestId.current)
        setListError(err.message || "Không tải được danh sách mã.");
    } finally {
      if (requestId === listRequestId.current) setIsLoadingList(false);
    }
  }, []);

  // Tự tải danh sách khi mở trang
  useEffect(() => {
    loadRecent();
  }, [loadRecent]);

  // Đếm ngược thời gian chờ của nút tải lại
  useEffect(() => {
    if (cooldown <= 0) return;
    const timer = setTimeout(() => setCooldown((c) => c - 1), 1000);
    return () => clearTimeout(timer);
  }, [cooldown]);

  const handleReloadClick = () => {
    if (cooldown > 0 || isLoadingList) return;
    setCooldown(RELOAD_COOLDOWN_SECONDS);
    loadRecent();
  };

  const handleUploadTypeChange = (e) => {
    setUploadType(e.target.value);
    setError(null);
  };

  const copyCode = async () => {
    if (!result?.code) return;
    try {
      await navigator.clipboard.writeText(result.code);
    } catch (err) {
      const tempInput = document.createElement("textarea");
      tempInput.value = result.code;
      tempInput.style.position = "fixed";
      tempInput.style.opacity = "0";
      document.body.appendChild(tempInput);
      tempInput.select();
      try {
        document.execCommand("copy");
      } catch (e) {}
      document.body.removeChild(tempInput);
    }
    setCopyText("Đã sao chép!");
    setTimeout(() => setCopyText("Sao chép"), 1500);
  };

  const sendToServer = async (payload) => {
    setIsSubmitting(true);
    setBtnText("Đang lưu lên Cloud...");

    try {
      // Đã cập nhật đúng đường dẫn API của bạn
      const res = await fetch("/api/save-puzzle", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify(payload),
      });

      const data = await res.json();
      if (!res.ok)
        throw new Error(data.error || `Lỗi server (mã ${res.status}).`);
      if (!data.code) throw new Error("Server không trả về mã dữ liệu.");

      // Set result dựa theo format response backend trả về
      setResult({ code: data.code, total: data.total });

      // Cập nhật ngay bảng mã gần đây (không tính vào thời gian chờ của nút tải lại)
      loadRecent();
    } catch (err) {
      throw err; // Ném lỗi ra ngoài để khối try...catch bên ngoài bắt
    } finally {
      setIsSubmitting(false);
      setBtnText("Tạo Mã Dữ Liệu");
    }
  };

  const handleUpload = async () => {
    setError(null);
    setResult(null);

    try {
      if (uploadType === "link") {
        if (!sheetLink.trim())
          throw new Error("Vui lòng dán link Google Sheet!");
        if (!isValidSheetLink(sheetLink.trim()))
          throw new Error("Link có vẻ không đúng định dạng Google Sheet.");

        await sendToServer({ sheetLink: sheetLink.trim() });
      } else {
        if (!file) throw new Error("Vui lòng chọn file Excel!");
        if (!/\.(xlsx|xls)$/i.test(file.name))
          throw new Error(
            "Vui lòng chọn đúng định dạng file Excel (.xlsx hoặc .xls).",
          );
        if (file.size > MAX_FILE_SIZE)
          throw new Error("File quá lớn. Vui lòng chọn file dưới 10MB.");

        setIsSubmitting(true);
        setBtnText("Đang phân tích dữ liệu...");

        const buffer = await readFileAsArrayBuffer(file);

        let fileData;
        try {
          const workbook = XLSX.read(new Uint8Array(buffer), { type: "array" });
          const sheetName = workbook.SheetNames[0];
          if (!sheetName) throw new Error("empty-workbook");
          const rawJsonData = XLSX.utils.sheet_to_json(
            workbook.Sheets[sheetName],
            { header: 1 },
          );
          fileData = minifyData(rawJsonData);
        } catch (err) {
          throw new Error(
            "Không đọc được nội dung file. Vui lòng kiểm tra file có đúng định dạng mẫu không.",
          );
        }

        const countEz = fileData.ez.length;
        const countMd = fileData.md.length;
        const countHd = fileData.hd.length;
        const total = countEz + countMd + countHd;

        // Validation trên client giống hệt backend để báo lỗi sớm cho người dùng
        if (total < 21)
          throw new Error(
            `Cần tối thiểu 21 câu hỏi. Hiện tại file của bạn chỉ có ${total} câu hợp lệ.`,
          );
        if (total > 50)
          throw new Error(
            `Tối đa 50 câu hỏi. File của bạn đang có ${total} câu.`,
          );
        if (countEz < 7 || countMd < 7 || countHd < 7) {
          throw new Error(
            `Cần ít nhất 7 câu cho mỗi mức độ. Hiện tại: Dễ (${countEz}), Trung bình (${countMd}), Khó (${countHd}).`,
          );
        }

        await sendToServer({ fileData });
      }
    } catch (err) {
      setError(err.message || "Đã xảy ra lỗi không xác định.");
      setIsSubmitting(false);
      setBtnText("Tạo Mã Dữ Liệu");
    }
  };

  const reloadDisabled = cooldown > 0 || isLoadingList;
  const reloadLabel = isLoadingList
    ? "Đang tải..."
    : cooldown > 0
      ? `Tải lại (${cooldown}s)`
      : "Tải lại";

  return (
    <div className="container">
      <h1>Tạo Gói Dữ Liệu Game</h1>
      <p className="subtitle">Tải lên bộ câu hỏi để nhận mã dữ liệu cho game</p>

      <div className="step">
        <div className="step-title">1. Lấy Template (Mẫu)</div>
        <div className="template-options">
          <a
            href="/csdl_mau.xlsx"
            download="csdl_mau.xlsx"
            className="template-link excel-btn"
            rel="noopener noreferrer"
          >
            Tải File Excel
          </a>
          <a
            href="https://docs.google.com/spreadsheets/d/1bBihWQ3_UlYFmROB8efOwtgNvEMdwq8_J9r8MiZvGIU/copy"
            className="template-link sheet-btn"
            target="_blank"
            rel="noopener noreferrer"
          >
            Tạo bản sao Google Sheet
          </a>
        </div>
      </div>

      <div className="step">
        <div className="step-title">2. Gửi dữ liệu đã điền</div>
        <div className="input-group">
          <div className="radio-group">
            <label>
              <input
                type="radio"
                name="uploadType"
                value="file"
                checked={uploadType === "file"}
                onChange={handleUploadTypeChange}
              />
              File Excel (.xlsx)
            </label>
            <label>
              <input
                type="radio"
                name="uploadType"
                value="link"
                checked={uploadType === "link"}
                onChange={handleUploadTypeChange}
              />
              Link Google Sheet
            </label>
          </div>

          {uploadType === "file" && (
            <div>
              <label className="visually-hidden" htmlFor="excelFile">
                Chọn file Excel đã điền dữ liệu
              </label>
              <input
                type="file"
                id="excelFile"
                accept=".xlsx,.xls"
                aria-describedby="fileHint"
                ref={fileInputRef}
                onChange={(e) => setFile(e.target.files[0])}
              />
              <div className="hint-text" id="fileHint">
                Định dạng .xlsx hoặc .xls, tối đa 10MB
              </div>
            </div>
          )}

          {uploadType === "link" && (
            <div>
              <label className="visually-hidden" htmlFor="sheetLink">
                Link Google Sheet
              </label>
              <input
                type="url"
                id="sheetLink"
                placeholder="https://docs.google.com/spreadsheets/d/..."
                aria-describedby="sheetHint"
                value={sheetLink}
                onChange={(e) => setSheetLink(e.target.value)}
              />
              <div className="hint-text" id="sheetHint">
                * Đảm bảo link đã được bật "Bất kỳ ai có liên kết".
              </div>
            </div>
          )}
        </div>
      </div>

      {error && (
        <div
          className="error-box"
          role="alert"
          aria-live="assertive"
          style={{ display: "block" }}
        >
          {error}
        </div>
      )}

      <button
        className="submit-btn"
        onClick={handleUpload}
        disabled={isSubmitting}
      >
        {isSubmitting && <span className="spinner"></span>}
        <span>{btnText}</span>
      </button>

      {result && (
        <div
          className="result-box"
          aria-live="polite"
          style={{ display: "block" }}
        >
          <div className="result-title">
            Xử lý thành công! Đây là mã của bạn:
          </div>
          <div className="code-row">
            <div className="code-display">{result.code}</div>
            <button type="button" className="copy-btn" onClick={copyCode}>
              {copyText}
            </button>
          </div>
          {result.total && (
            <div className="result-breakdown" style={{ display: "block" }}>
              Đã lưu thành công <strong>{result.total}</strong> câu hỏi.
            </div>
          )}
          <div className="result-hint">(Hãy nhập mã này vào web game)</div>
        </div>
      )}

      {/* ----- Bảng các mã đã được tạo gần đây ----- */}
      <div className="recent-section">
        <div className="recent-header">
          <h2 className="recent-title">Các mã đã được tạo gần đây</h2>
          <button
            type="button"
            className="reload-btn"
            onClick={handleReloadClick}
            disabled={reloadDisabled}
          >
            {reloadLabel}
          </button>
        </div>

        {listError && (
          <div className="error-box" role="alert" style={{ display: "block" }}>
            {listError}
          </div>
        )}

        <div className="table-scroll">
          <table className="recent-table">
            <thead>
              <tr>
                <th>STT</th>
                <th>Tên mã</th>
                <th>Thời gian tạo</th>
                <th>Nội dung</th>
              </tr>
            </thead>
            <tbody>
              {recent.length === 0 ? (
                <tr>
                  <td colSpan={4} className="recent-empty">
                    {isLoadingList
                      ? "Đang tải danh sách..."
                      : "Chưa có mã nào."}
                  </td>
                </tr>
              ) : (
                recent.map((item, index) => {
                  const preview = firstWords(item.preview);
                  return (
                    <tr key={item.code ?? index}>
                      <td>{index + 1}</td>
                      <td className="recent-code">{item.code}</td>
                      <td>{formatTime(item.createdAt)}</td>
                      <td>{preview ? `"${preview}..."` : "—"}</td>
                    </tr>
                  );
                })
              )}
            </tbody>
          </table>
        </div>
      </div>
    </div>
  );
}
