"use client";

import { AlertTriangle, RefreshCw } from "lucide-react";

export default function Error({
  error,
  reset,
}: {
  error: Error & { digest?: string };
  reset: () => void;
}) {
  const isDb = /DATABASE_URL/i.test(error.message);

  return (
    <div className="page-stack" style={{ maxWidth: 560, margin: "4rem auto", textAlign: "center" }}>
      <AlertTriangle size={40} style={{ color: "var(--danger)", margin: "0 auto 1rem" }} />
      <h1 style={{ fontSize: "1.35rem", marginBottom: "0.5rem" }}>Sayfa yüklenemedi</h1>
      <p style={{ color: "var(--muted)", marginBottom: "1rem" }}>
        {isDb
          ? "Veritabanı bağlantısı kurulamadı. Vercel proje ayarlarında DATABASE_URL ortam değişkeninin tanımlı olduğundan emin olun."
          : "Sunucu tarafında bir hata oluştu."}
      </p>
      {error.message && (
        <pre
          style={{
            textAlign: "left",
            fontSize: "0.8rem",
            padding: "0.75rem 1rem",
            background: "var(--surface-2)",
            borderRadius: 8,
            overflow: "auto",
            marginBottom: "1.25rem",
          }}
        >
          {error.message}
        </pre>
      )}
      <button type="button" className="btn btn-primary" onClick={reset}>
        <RefreshCw size={16} />
        Tekrar dene
      </button>
    </div>
  );
}
