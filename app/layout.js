// app/layout.js
import "./globals.css";

export const metadata = {
  title: "Lưu Trữ Dữ Dữ Game Giải Đố",
  description: "Tải lên bộ câu hỏi để nhận mã dữ liệu cho game",
  icons: {
    icon: "data:image/svg+xml,<svg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 100 100'><text y='.9em' font-size='90'>🧩</text></svg>",
  },
  themeColor: [
    { media: "(prefers-color-scheme: light)", color: "#111827" },
    { media: "(prefers-color-scheme: dark)", color: "#0e0f12" },
  ],
  colorScheme: "light dark",
};

export default function RootLayout({ children }) {
  return (
    <html lang="vi">
      <body>{children}</body>
    </html>
  );
}
