import type { Metadata } from "next";
import "./styles.css";

export const metadata: Metadata = {
  title: "Krishiv Agri Genetics LLP | Krishiv Seeds",
  description:
    "Krishiv Seeds supplies quality wheat, soyabean, paddy, chana, moong, and mustard seeds from Kota, Rajasthan."
};

export default function RootLayout({
  children
}: Readonly<{ children: React.ReactNode }>) {
  return (
    <html lang="en">
      <body>{children}</body>
    </html>
  );
}
