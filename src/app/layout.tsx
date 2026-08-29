import type { Metadata } from "next";
import { Geist, Geist_Mono } from "next/font/google";
import "./globals.css";

const geistSans = Geist({
  variable: "--font-geist-sans",
  subsets: ["latin"],
});

const geistMono = Geist_Mono({
  variable: "--font-geist-mono",
  subsets: ["latin"],
});

const siteUrl = "https://dockaro.com";

export const metadata: Metadata = {
  metadataBase: new URL(siteUrl),
  title: {
    default: "DocKaro — Word, Sheets & Slides in your browser",
    template: "%s — DocKaro",
  },
  description:
    "DocKaro is a browser-based Word, Excel and PowerPoint alternative with a full API for developers. Edit documents, spreadsheets and presentations online — no install, no license fees.",
  keywords: [
    "online word processor",
    "browser spreadsheet",
    "document editor API",
    "docx editor online",
    "excel alternative",
    "powerpoint alternative",
    "embeddable office suite",
    "DocKaro",
  ],
  authors: [{ name: "DocKaro" }],
  creator: "DocKaro",
  openGraph: {
    type: "website",
    url: siteUrl,
    siteName: "DocKaro",
    title: "DocKaro — Word, Sheets & Slides in your browser",
    description:
      "The full office suite — Docs, Sheets, Slides — in your browser, with an API to embed it in your own product.",
    locale: "en_US",
  },
  twitter: {
    card: "summary_large_image",
    title: "DocKaro — Word, Sheets & Slides in your browser",
    description:
      "The full office suite — Docs, Sheets, Slides — in your browser, with an API to embed it in your own product.",
  },
  robots: {
    index: true,
    follow: true,
    googleBot: {
      index: true,
      follow: true,
      "max-image-preview": "large",
      "max-snippet": -1,
    },
  },
  alternates: {
    canonical: siteUrl,
  },
};

export default function RootLayout({ children }: LayoutProps<"/">) {
  return (
    <html
      lang="en"
      className={`${geistSans.variable} ${geistMono.variable} h-full antialiased`}
    >
      <body className="min-h-full flex flex-col bg-[#0a0a0b] text-white">
        {children}
      </body>
    </html>
  );
}
