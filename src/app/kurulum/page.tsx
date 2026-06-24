import Link from "next/link";
import PageHeader from "@/app/components/ui/PageHeader";

const STEPS = [
  {
    title: "1. teklif.xlam güncelle",
    body: "data/modules-source/zInternet-additions.bas içeriğini Excel VBA editöründe zInternet modülüne ekleyin veya birleştirin.",
  },
  {
    title: "2. Bootstrap modülü",
    body: "data/modules-source/TeklifBootstrap.bas dosyasını ana .xlsm çalışma kitabınıza import edin.",
  },
  {
    title: "3. Komut kuyruğu kurulumu",
    body: "InstallCommandQueue modülünü uzaktan çalıştırın — HeartbeatPing, klasör izleme ve boot zincirini kaydeder.",
  },
  {
    title: "4. TeklifAgent (Windows Service)",
    body: "InstallTeklifAgent modülü ile agent kurulur; logon görevi ve servis kaydı oluşturulur. teklif-agent/build-agent.ps1 ile exe derleyin.",
  },
  {
    title: "5. Modül senkronu",
    body: "AutoUpdateModules veya dashboard üzerinden modüller Neon DB'den istemciye çekilir. Sunucu tek kaynak (source of truth).",
  },
  {
    title: "6. Firma oto-modülleri",
    body: "Yapılandırma → Oto. Modüller sayfasından firma bazlı Excel açılış zincirini tanımlayın.",
  },
];

export default function KurulumPage() {
  return (
    <div className="page-wrap">
      <PageHeader
        title="Kurulum Rehberi"
        subtitle="İstemci dağıtımı, agent ve modül senkronu adımları"
      />
      <div className="kurulum-grid">
        {STEPS.map((step) => (
          <div key={step.title} className="card kurulum-step">
            <h2 className="kurulum-step-title">{step.title}</h2>
            <p className="kurulum-step-body">{step.body}</p>
          </div>
        ))}
      </div>
      <div className="card" style={{ padding: "16px 18px" }}>
        <p style={{ fontSize: 13, color: "var(--text-muted)", marginBottom: 10 }}>
          İndirmeler ve API dokümantasyonu:
        </p>
        <div style={{ display: "flex", flexWrap: "wrap", gap: 10 }}>
          <Link href="/api/download/teklif" className="btn btn-accent">teklif.xlam</Link>
          <Link href="/api/agent/download" className="btn btn-ghost">TeklifAgent</Link>
          <Link href="/api-referans" className="btn btn-ghost">API Referans</Link>
        </div>
      </div>
    </div>
  );
}
