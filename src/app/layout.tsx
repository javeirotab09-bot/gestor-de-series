import type { Metadata } from "next";
import type { ReactNode } from "react";
import "./globals.css";

export const metadata: Metadata = {
  title: "Plataforma de Validación de Series",
  description:
    "Sistema logístico para importación masiva, validación de series con lector y exportación de resultados.",
};

export default function RootLayout({ children }: { children: ReactNode }) {
  return (
    <html lang="es">
      <body className="bg-slate-100 text-slate-900 antialiased">{children}</body>
    </html>
  );
}
