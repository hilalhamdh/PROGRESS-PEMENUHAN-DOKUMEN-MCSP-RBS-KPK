import type { Metadata } from "next";
import "./globals.css";

export const metadata: Metadata = {
  title: "Progress Dokumen MCSP-RBS KPK | Kabupaten Deiyai",
  description:
    "Dashboard progress pemenuhan dokumen MCSP-RBS KPK Kabupaten Deiyai per 24 Agustus 2026",
};

export default function RootLayout({ children }: LayoutProps<"/">) {
  return (
    <html lang="id" className="h-full antialiased">
      <body className="min-h-full flex flex-col font-sans">{children}</body>
    </html>
  );
}
