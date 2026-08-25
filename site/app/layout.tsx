import { Analytics } from "@vercel/analytics/next";
import { RootProvider } from "fumadocs-ui/provider/next";
import type { Metadata } from "next";
import { Inter, JetBrains_Mono } from "next/font/google";
import localFont from "next/font/local";
import type { ReactNode } from "react";
import { libraryDescription } from "@/lib/library";
import "./global.css";

const inter = Inter({
  subsets: ["latin"],
  weight: ["400", "500", "600", "700", "800"],
  variable: "--font-inter",
});

const jetBrainsMono = JetBrains_Mono({
  subsets: ["latin"],
  weight: ["400", "500"],
  variable: "--font-jetbrains-mono",
});

// Pretendard is not on Google Fonts, so it is self-hosted from the npm package.
// This is the full ~2MB variable file; the package also ships a dynamic subset,
// but that is a stylesheet of many @font-face rules, which next/font/local
// cannot consume. Worth revisiting once Korean content actually ships.
const pretendard = localFont({
  src: "../node_modules/pretendard/dist/web/variable/woff2/PretendardVariable.woff2",
  weight: "45 920",
  display: "swap",
  // Every page is English today, so the browser should fetch this only when a
  // Hangul glyph is actually asked for rather than on every first paint
  preload: false,
  variable: "--font-pretendard",
});

export const metadata: Metadata = {
  title: {
    default: "docx-editor",
    template: "%s - docx-editor",
  },
  description: libraryDescription,
};

export default function RootLayout({ children }: { children: ReactNode }) {
  return (
    <html
      className={`${inter.variable} ${jetBrainsMono.variable} ${pretendard.variable}`}
      lang="en"
      suppressHydrationWarning
    >
      <body className="flex min-h-screen flex-col font-sans">
        <RootProvider>{children}</RootProvider>
        <Analytics />
      </body>
    </html>
  );
}
