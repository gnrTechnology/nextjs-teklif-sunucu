import type { LucideIcon } from "lucide-react";
import {
  LayoutDashboard,
  Radio,
  Terminal,
  Upload,
  Monitor,
  ScrollText,
  BarChart3,
  KeyRound,
  Zap,
  FolderSearch,
  Package,
  Lightbulb,
  Plug,
  BookOpen,
} from "lucide-react";

export type NavItem = {
  href: string;
  label: string;
  icon: LucideIcon;
  keywords?: string[];
};

export type NavSection = {
  id: string;
  label: string;
  items: NavItem[];
};

export type NavGroup = {
  id: string;
  label: string;
  sections: NavSection[];
};

export const NAV_GROUPS: NavGroup[] = [
  {
    id: "ops",
    label: "Operasyon",
    sections: [
      {
        id: "ops-overview",
        label: "Genel Bakış",
        items: [
          { href: "/", label: "Dashboard", icon: LayoutDashboard, keywords: ["ana", "özet"] },
          { href: "/analitik", label: "Analitik", icon: BarChart3, keywords: ["istatistik", "grafik"] },
        ],
      },
      {
        id: "ops-live",
        label: "Canlı İzleme",
        items: [
          { href: "/heartbeats", label: "Nabız İzleme", icon: Radio, keywords: ["heartbeat", "mac"] },
          { href: "/loglar", label: "Loglar", icon: ScrollText, keywords: ["aktivite", "denetim"] },
        ],
      },
      {
        id: "ops-device",
        label: "Cihaz & Komut",
        items: [
          { href: "/cihazlar", label: "Cihaz Bilgileri", icon: Monitor, keywords: ["donanım", "snapshot"] },
          { href: "/komutlar", label: "Uzak Komutlar", icon: Terminal, keywords: ["komut", "kuyruk"] },
          { href: "/modul-ciktilari", label: "Modül Çıktıları", icon: Upload, keywords: ["output", "screenshot"] },
        ],
      },
    ],
  },
  {
    id: "config",
    label: "Yapılandırma",
    sections: [
      {
        id: "cfg-license",
        label: "Lisans & Erişim",
        items: [
          { href: "/lisanslar", label: "Lisanslar", icon: KeyRound, keywords: ["license", "mac"] },
        ],
      },
      {
        id: "cfg-auto",
        label: "Otomasyon",
        items: [
          { href: "/firma-modulleri", label: "Oto. Modüller", icon: Zap, keywords: ["firma", "auto-start"] },
          { href: "/klasor-izleme", label: "Klasör İzleme", icon: FolderSearch, keywords: ["watch", "folder"] },
        ],
      },
    ],
  },
  {
    id: "dev",
    label: "Geliştirici",
    sections: [
      {
        id: "dev-modules",
        label: "Modül Yönetimi",
        items: [
          { href: "/moduller", label: "Uzak Modüller", icon: Package, keywords: ["vba", "modül"] },
          { href: "/oneriler", label: "Modül Önerileri", icon: Lightbulb, keywords: ["proposal"] },
        ],
      },
      {
        id: "dev-docs",
        label: "Kaynaklar",
        items: [
          { href: "/api-referans", label: "API Referans", icon: Plug, keywords: ["endpoint", "rest"] },
          { href: "/kurulum", label: "Kurulum Rehberi", icon: BookOpen, keywords: ["setup", "xlam", "agent"] },
        ],
      },
    ],
  },
];

export const ALL_NAV_ITEMS = NAV_GROUPS.flatMap((g) => g.sections.flatMap((s) => s.items));

export function findNavItem(pathname: string): NavItem | undefined {
  if (pathname === "/") return ALL_NAV_ITEMS.find((i) => i.href === "/");
  return ALL_NAV_ITEMS.find((i) => i.href !== "/" && pathname.startsWith(i.href));
}

export function breadcrumbTrail(pathname: string): { href: string; label: string }[] {
  const trail = [{ href: "/", label: "Dashboard" }];
  const current = findNavItem(pathname);
  if (current && current.href !== "/") trail.push({ href: current.href, label: current.label });
  return trail;
}
