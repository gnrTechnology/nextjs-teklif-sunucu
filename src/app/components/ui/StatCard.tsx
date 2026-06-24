import Link from "next/link";

export default function StatCard({
  label,
  value,
  hint,
  tone,
  href,
}: {
  label: string;
  value: React.ReactNode;
  hint?: string;
  tone?: "green" | "accent" | "yellow" | "red";
  href?: string;
}) {
  const inner = (
    <div className={`stat-card${href ? " stat-card--link" : ""}`}>
      <div className="stat-label">{label}</div>
      <div className={`stat-value${tone ? ` ${tone}` : ""}`}>{value}</div>
      {hint && <div className="stat-hint">{hint}</div>}
    </div>
  );
  if (href) return <Link href={href}>{inner}</Link>;
  return inner;
}
