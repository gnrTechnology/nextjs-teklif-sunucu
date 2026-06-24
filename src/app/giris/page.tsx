"use client";

import { useRouter, useSearchParams } from "next/navigation";
import { useState, Suspense } from "react";
import { Lock } from "lucide-react";

function LoginForm() {
  const router = useRouter();
  const params = useSearchParams();
  const [password, setPassword] = useState("");
  const [error, setError] = useState("");
  const [loading, setLoading] = useState(false);

  async function onSubmit(e: React.FormEvent) {
    e.preventDefault();
    setLoading(true);
    setError("");
    const res = await fetch("/api/auth/login", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ password }),
    });
    if (res.ok) {
      router.push(params.get("next") || "/");
      router.refresh();
    } else {
      setError("Geçersiz parola");
    }
    setLoading(false);
  }

  return (
    <div className="login-page">
      <form className="login-card card" onSubmit={onSubmit}>
        <div className="login-icon">
          <Lock size={28} />
        </div>
        <h1 className="page-title">Yönetici Girişi</h1>
        <p className="page-sub">Teklif Sunucu paneline erişim</p>
        <input
          type="password"
          className="form-input"
          placeholder="Parola"
          value={password}
          onChange={(e) => setPassword(e.target.value)}
          autoFocus
        />
        {error && <p className="login-error">{error}</p>}
        <button type="submit" className="btn btn-primary login-submit" disabled={loading}>
          {loading ? "Giriş…" : "Giriş yap"}
        </button>
      </form>
    </div>
  );
}

export default function GirisPage() {
  return (
    <Suspense>
      <LoginForm />
    </Suspense>
  );
}
