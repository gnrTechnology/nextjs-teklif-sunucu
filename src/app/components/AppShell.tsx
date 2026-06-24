"use client";

import { useEffect, useMemo, useState } from "react";
import Link from "next/link";
import { usePathname, useRouter } from "next/navigation";
import { Menu, Search, Server } from "lucide-react";
import { ALL_NAV_ITEMS, breadcrumbTrail } from "@/lib/nav";
import Sidebar from "./Sidebar";
import ThemeToggle from "./ThemeToggle";

export default function AppShell({ children }: { children: React.ReactNode }) {
  const pathname = usePathname();
  const router = useRouter();
  const [mobileOpen, setMobileOpen] = useState(false);
  const [query, setQuery] = useState("");

  const crumbs = useMemo(() => breadcrumbTrail(pathname), [pathname]);

  useEffect(() => {
    setMobileOpen(false);
  }, [pathname]);

  function onSearchSubmit(e: React.FormEvent) {
    e.preventDefault();
    const q = query.trim().toLowerCase();
    if (!q) return;
    const macHit = q.match(/([0-9a-f]{2}[:-]){5}[0-9a-f]{2}/i);
    if (macHit) {
      router.push(`/heartbeats?q=${encodeURIComponent(macHit[0])}`);
      return;
    }
    const nav = ALL_NAV_ITEMS.find(
      (i) =>
        i.label.toLowerCase().includes(q) ||
        i.keywords?.some((k) => k.includes(q) || q.includes(k)),
    );
    if (nav) {
      router.push(nav.href);
      return;
    }
    router.push(`/moduller?q=${encodeURIComponent(query.trim())}`);
  }

  if (pathname === "/giris") {
    return <>{children}</>;
  }

  return (
    <div className="app-shell">
      <div
        className={`sidebar-backdrop${mobileOpen ? " sidebar-backdrop--open" : ""}`}
        onClick={() => setMobileOpen(false)}
        aria-hidden={!mobileOpen}
      />
      <Sidebar mobileOpen={mobileOpen} onClose={() => setMobileOpen(false)} />
      <div className="app-main">
        <header className="topbar">
          <button
            type="button"
            className="topbar-menu-btn"
            onClick={() => setMobileOpen(true)}
            aria-label="Menüyü aç"
          >
            <Menu size={20} />
          </button>
          <nav className="breadcrumb" aria-label="Breadcrumb">
            {crumbs.map((c, i) => (
              <span key={c.href} className="breadcrumb-item">
                {i > 0 && <span className="breadcrumb-sep">/</span>}
                <Link href={c.href}>{c.label}</Link>
              </span>
            ))}
          </nav>
          <form className="topbar-search" onSubmit={onSearchSubmit}>
            <Search size={16} className="topbar-search-icon" />
            <input
              value={query}
              onChange={(e) => setQuery(e.target.value)}
              placeholder="MAC, modül veya sayfa ara…"
              aria-label="Global arama"
            />
          </form>
          <div className="topbar-actions">
            <ThemeToggle />
            <span className="topbar-badge" title="Neon PostgreSQL">
              <Server size={14} />
              <span className="topbar-badge-text">Neon</span>
            </span>
          </div>
        </header>
        <main className="app-content">{children}</main>
      </div>
    </div>
  );
}
