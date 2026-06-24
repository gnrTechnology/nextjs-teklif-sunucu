"use client";

import Link from "next/link";
import { usePathname } from "next/navigation";
import { X, Server } from "lucide-react";
import { NAV_GROUPS } from "@/lib/nav";

export default function Sidebar({
  mobileOpen,
  onClose,
}: {
  mobileOpen?: boolean;
  onClose?: () => void;
}) {
  const pathname = usePathname();

  return (
    <aside className={`sidebar${mobileOpen ? " sidebar--open" : ""}`}>
      <div className="sidebar-logo">
        <div className="sidebar-logo-icon">
          <Server size={18} />
        </div>
        <div className="sidebar-logo-text">
          <div className="sidebar-logo-title">Teklif Sunucu</div>
          <div className="sidebar-logo-sub">VBA Lisans &amp; Modül</div>
        </div>
        {mobileOpen && (
          <button type="button" className="sidebar-close" onClick={onClose} aria-label="Kapat">
            <X size={18} />
          </button>
        )}
      </div>

      <nav className="sidebar-nav">
        {NAV_GROUPS.map((group) => (
          <div key={group.id} className="sidebar-group">
            <div className="sidebar-group-label">{group.label}</div>
            {group.items.map(({ href, label, icon: Icon }) => {
              const active = href === "/" ? pathname === "/" : pathname.startsWith(href);
              return (
                <Link
                  key={href}
                  href={href}
                  className={`sidebar-link${active ? " sidebar-link--active" : ""}`}
                >
                  <span className="sidebar-link-icon">
                    <Icon size={17} strokeWidth={2} />
                  </span>
                  <span className="sidebar-link-label">{label}</span>
                  {active && <span className="sidebar-link-bar" />}
                </Link>
              );
            })}
          </div>
        ))}
      </nav>

      <div className="sidebar-footer">
        <span className="sidebar-footer-text">v1.1 · Neon DB</span>
      </div>
    </aside>
  );
}
