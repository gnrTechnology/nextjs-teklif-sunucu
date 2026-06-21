const BASE = "https://nextjs-teklif-sunucu.vercel.app/api";

const ENDPOINTS = [
  {
    method: "GET",
    path: "/license/{mac}/",
    title: "Lisans Sorgula",
    desc: "MAC adresine göre lisans kaydını döndürür.",
    request: null,
    response: `{\n  "success": true,\n  "data": {\n    "macAdresi": "38:00:25:B0:44:6B",\n    "license": "true",\n    "firmaAdi": "EPRON",\n    "userAdi": "Onur MEMİŞ",\n    "updatedAt": "2026-06-21T12:29:43.000Z"\n  }\n}`,
    responses: [
      { code: "200", desc: "Kayıt bulundu" },
      { code: "200 (success:false)", desc: "MAC bulunamadı" },
      { code: "400", desc: "Geçersiz MAC" },
    ],
  },
  {
    method: "POST",
    path: "/license/",
    title: "Lisans Kayıt / Güncelleme",
    desc: "Yeni cihaz kaydeder veya mevcut kaydı günceller. Lisans durumu (true/false) sunucu tarafından korunur.",
    request: `{\n  "macAdresi": "38:00:25:B0:44:6B",\n  "firmaAdi": "EPRON",\n  "userAdi": "Onur MEMİŞ",\n  "ipAdresi": "192.168.1.100",\n  "dosyaAdi": "teklif.xlam"\n}`,
    response: `{\n  "success": true,\n  "message": "Yeni lisans kaydı oluşturuldu.",\n  "data": {\n    "macAdresi": "38:00:25:B0:44:6B",\n    "license": "false",\n    "firmaAdi": "EPRON",\n    "userAdi": "Onur MEMİŞ"\n  }\n}`,
    responses: [
      { code: "201", desc: "Yeni kayıt oluşturuldu (license=false)" },
      { code: "200", desc: "Mevcut kayıt güncellendi" },
      { code: "400", desc: "macAdresi eksik" },
    ],
  },
  {
    method: "PATCH",
    path: "/license/{mac}/",
    title: "Lisans Aktif/Pasif",
    desc: "Dashboard üzerinden lisansı aktif veya pasif yapar.",
    request: `{ "license": "true" }`,
    response: `{ "success": true, "data": { "license": "true" } }`,
    responses: [
      { code: "200", desc: "Güncellendi" },
      { code: "404", desc: "MAC bulunamadı" },
    ],
  },
  {
    method: "POST",
    path: "/module/",
    title: "Uzak VBA Modülü Al",
    desc: "methodName ile VBA kaynak kodunu döndürür. RunRemoteCode tarafından kullanılır.",
    request: `{ "methodName": "getLicense" }`,
    response: `{\n  "methodName": "getLicense",\n  "description": "...",\n  "code": "Public Function DynamicFunc..."\n}`,
    responses: [
      { code: "200", desc: "Modül kodu döndürüldü" },
      { code: "404", desc: "Modül bulunamadı" },
    ],
  },
  {
    method: "GET",
    path: "/auto-start/{mac}/",
    title: "Firma Otomatik Modülleri",
    desc: "MAC adresine göre Excel açılışında çalışacak modül listesini döndürür.",
    request: null,
    response: `{\n  "success": true,\n  "data": {\n    "firmaAdi": "EPRON",\n    "modules": [\n      { "methodName": "getLicense", "order": 1, "delaySeconds": 0 }\n    ]\n  }\n}`,
    responses: [
      { code: "200", desc: "Modül listesi döndürüldü" },
      { code: "200 (modules:[])", desc: "Modül tanımı yok" },
    ],
  },
  {
    method: "POST",
    path: "/download/teklif/",
    title: "Eklenti İndir (.xlam)",
    desc: "Lisanslı kullanıcıya teklif.xlam eklentisini binary olarak gönderir.",
    request: `{ "macAdresi": "38:00:25:B0:44:6B" }`,
    response: "(binary .xlam dosyası)",
    responses: [
      { code: "200", desc: "Dosya içeriği (binary)" },
      { code: "403", desc: "Lisans pasif veya bulunamadı" },
    ],
  },
];

function MethodBadge({ method }: { method: string }) {
  const cls = method === "GET" ? "method-get" : method === "POST" ? "method-post" : "method-patch";
  return (
    <span className={`badge ${cls}`} style={{ fontFamily: "monospace", fontWeight: 700, fontSize: 11 }}>
      {method}
    </span>
  );
}

export default function ApiReferansPage() {
  return (
    <div className="page-wrap">
      <div className="page-header">
        <div>
          <div className="page-title">🔌 API Referans</div>
          <div className="page-sub">
            Base URL: <span className="mono">{BASE}</span>
          </div>
        </div>
      </div>

      {ENDPOINTS.map((ep) => (
        <div
          key={`${ep.method}-${ep.path}`}
          className="card"
          style={{ marginBottom: "1.25rem" }}
        >
          <div className="card-header">
            <span className="card-title" style={{ gap: 12 }}>
              <MethodBadge method={ep.method} />
              <span className="mono" style={{ fontSize: 14 }}>{ep.path}</span>
              <span style={{ color: "var(--text-muted)", fontWeight: 400, fontSize: 13 }}>{ep.title}</span>
            </span>
          </div>

          <div style={{ padding: "14px 18px", display: "flex", flexDirection: "column", gap: 14 }}>
            <p style={{ fontSize: 13, color: "var(--text-muted)" }}>{ep.desc}</p>

            <div style={{ display: "grid", gridTemplateColumns: ep.request ? "1fr 1fr" : "1fr", gap: 14 }}>
              {ep.request && (
                <div>
                  <div style={{ fontSize: 11, fontWeight: 600, textTransform: "uppercase", letterSpacing: "0.06em", color: "var(--text-muted)", marginBottom: 6 }}>
                    Request Body (JSON)
                  </div>
                  <pre style={{ fontFamily: "var(--font-geist-mono, monospace)", fontSize: 12, lineHeight: 1.7, color: "var(--code-text)", background: "var(--code-bg)", padding: "12px 14px", borderRadius: "var(--radius-sm)", overflowX: "auto", whiteSpace: "pre-wrap" }}>
                    {ep.request}
                  </pre>
                </div>
              )}
              <div>
                <div style={{ fontSize: 11, fontWeight: 600, textTransform: "uppercase", letterSpacing: "0.06em", color: "var(--text-muted)", marginBottom: 6 }}>
                  Örnek Yanıt
                </div>
                <pre style={{ fontFamily: "var(--font-geist-mono, monospace)", fontSize: 12, lineHeight: 1.7, color: "var(--code-text)", background: "var(--code-bg)", padding: "12px 14px", borderRadius: "var(--radius-sm)", overflowX: "auto", whiteSpace: "pre-wrap" }}>
                  {ep.response}
                </pre>
              </div>
            </div>

            <div>
              <div style={{ fontSize: 11, fontWeight: 600, textTransform: "uppercase", letterSpacing: "0.06em", color: "var(--text-muted)", marginBottom: 6 }}>
                HTTP Yanıt Kodları
              </div>
              <div style={{ display: "flex", flexWrap: "wrap", gap: 8 }}>
                {ep.responses.map((r) => (
                  <span key={r.code} style={{ fontSize: 12, color: "var(--text-muted)" }}>
                    <span className="mono">{r.code}</span> {r.desc}
                  </span>
                ))}
              </div>
            </div>
          </div>
        </div>
      ))}
    </div>
  );
}
