"use client";

import Link from "next/link";
import { usePathname } from "next/navigation";

const NAV = [
  { href: "/",                  icon: "⬛", label: "Dashboard"         },
  { href: "/lisanslar",         icon: "🔑", label: "Lisanslar"         },
  { href: "/loglar",            icon: "📋", label: "Loglar"            },
  { href: "/moduller",          icon: "📦", label: "Uzak Modüller"     },
  { href: "/firma-modulleri",   icon: "⚡", label: "Oto. Modüller"    },
  { href: "/heartbeats",        icon: "📡", label: "Nabız İzleme"      },
  { href: "/cihazlar",          icon: "🖥", label: "Cihaz Bilgileri"   },
  { href: "/komutlar",          icon: "🎮", label: "Uzak Komutlar"     },
  { href: "/modul-ciktilari",   icon: "📤", label: "Modül Çıktıları"  },
  { href: "/analitik",          icon: "📊", label: "Analitik"          },
  { href: "/api-referans",      icon: "🔌", label: "API Referans"      },
];

export default function Sidebar() {
  const pathname = usePathname();

  return (
    <aside className="sidebar">
      <div className="sidebar-logo">
        <div className="sidebar-logo-icon">🖥️</div>
        <div className="sidebar-logo-text">
          <div className="sidebar-logo-title">Teklif Sunucu</div>
          <div className="sidebar-logo-sub">VBA Lisans &amp; Modül</div>
        </div>
      </div>

      <nav className="sidebar-nav">
        {NAV.map(({ href, icon, label }) => {
          const active = href === "/" ? pathname === "/" : pathname.startsWith(href);
          return (
            <Link
              key={href}
              href={href}
              className={`sidebar-link${active ? " sidebar-link--active" : ""}`}
            >
              <span className="sidebar-link-icon">{icon}</span>
              <span className="sidebar-link-label">{label}</span>
              {active && <span className="sidebar-link-bar" />}
            </Link>
          );
        })}
      </nav>

      <div className="sidebar-footer">
        <span className="sidebar-footer-text">v1.0 · Neon DB</span>
      </div>
    </aside>
  );
}
