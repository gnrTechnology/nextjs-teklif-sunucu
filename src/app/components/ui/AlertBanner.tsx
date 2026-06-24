import Link from "next/link";
import type { LucideIcon } from "lucide-react";
import { AlertTriangle, XCircle } from "lucide-react";

export type AlertItem = {
  id: string;
  tone: "warning" | "danger";
  message: string;
  href?: string;
  icon?: LucideIcon;
};

export default function AlertBanner({ items }: { items: AlertItem[] }) {
  if (items.length === 0) return null;
  return (
    <div className="alert-banner-stack">
      {items.map((item) => {
        const Icon = item.icon ?? (item.tone === "danger" ? XCircle : AlertTriangle);
        const content = (
          <div className={`alert-banner alert-banner--${item.tone}`}>
            <Icon size={16} className="alert-banner-icon" />
            <span>{item.message}</span>
          </div>
        );
        return item.href ? (
          <Link key={item.id} href={item.href} className="alert-banner-link">
            {content}
          </Link>
        ) : (
          <div key={item.id}>{content}</div>
        );
      })}
    </div>
  );
}
