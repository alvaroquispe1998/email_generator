import "./globals.css";
import type { Metadata } from "next";

export const metadata: Metadata = {
  title: "Generador de CSV Outlook - UAI",
  description: "Convierte Excel de estudiantes a CSV compatible con Outlook."
};

export default function RootLayout({ children }: { children: React.ReactNode }) {
  return (
    <html lang="es">
      <body className="text-ink">
        {children}
      </body>
    </html>
  );
}
