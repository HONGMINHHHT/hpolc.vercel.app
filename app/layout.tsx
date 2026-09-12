import type { Metadata } from "next";
import "./globals.css";

export const metadata: Metadata = {
  title: "HPOLC – Quản lý đào tạo",
  description: "Tổng hợp, chuẩn hóa và đối soát quyết định trúng tuyển.",
  icons: {
    icon: "/favicon.svg",
    shortcut: "/favicon.svg",
  },
};

export default function RootLayout({
  children,
}: Readonly<{
  children: React.ReactNode;
}>) {
  return (
    <html lang="vi">
      <body className="antialiased">{children}</body>
    </html>
  );
}
