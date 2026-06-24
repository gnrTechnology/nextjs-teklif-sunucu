"use client";

import Link from "next/link";
import { usePathname } from "next/navigation";
import { useEffect, useState } from "react";
import { ChevronDown, X, Server } from "lucide-react";
import { NAV_GROUPS } from "@/lib/nav";

const STORAGE_KEY = "sidebar-collapsed-sections";

function loadCollapsed(): Set<string> {
  if (typeof window === "undefined") return new Set();
  try {
    const raw = localStorage.getItem(STORAGE_KEY);
    if (!raw) return new Set();
    return new Set(JSON.parse(raw) as string[]);
  } catch {
    return new Set();
  }
}

function saveCollapsed(set: Set<string>) {
  localStorage.setItem(STORAGE_KEY, JSON.stringify([...set]));
}

export default function Sidebar({
  mobileOpen,
  onClose,
}: {
  mobileOpen?: boolean;
  onClose?: () => void;
}) {
  const pathname = usePathname();
  const [collapsed, setCollapsed] = useState<Set<string>>(new Set());

  useEffect(() => {
    setCollapsed(loadCollapsed());
  }, []);

  function toggleSection(id: string) {
    setCollapsed((prev) => {
      const next = new Set(prev);
      if (next.has(id)) next.delete(id);
      else next.add(id);
      saveCollapsed(next);
      return next;
    });
  }

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
            {group.sections.map((section) => {
              const isCollapsed = collapsed.has(section.id);
              const sectionActive = section.items.some((item) =>
                item.href === "/" ? pathname === "/" : pathname.startsWith(item.href),
              );
              return (
                <div key={section.id} className="sidebar-section">
                  <button
                    type="button"
                    className={`sidebar-section-toggle${sectionActive ? " sidebar-section-toggle--active" : ""}`}
                    onClick={() => toggleSection(section.id)}
                    aria-expanded={!isCollapsed}
                  >
                    <span>{section.label}</span>
                    <ChevronDown
                      size={14}
                      className={`sidebar-section-chevron${isCollapsed ? " sidebar-section-chevron--collapsed" : ""}`}
                    />
                  </button>
                  {!isCollapsed && (
                    <div className="sidebar-section-items">
                      {section.items.map(({ href, label, icon: Icon }) => {
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
                  )}
                </div>
              );
            })}
          </div>
        ))}
      </nav>

      <div className="sidebar-footer">
        <span className="sidebar-footer-text">v1.2 · Neon DB</span>
      </div>
    </aside>
  );
}
