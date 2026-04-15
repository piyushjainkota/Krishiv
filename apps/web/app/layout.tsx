import type { Metadata } from "next";
import "./styles.css";

export const metadata: Metadata = {
  title: "KRISHIV Seed Intake Platform",
  description: "Certification-safe raw seed intake and traceability platform"
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
