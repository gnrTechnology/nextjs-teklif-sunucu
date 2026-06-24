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

export type NavGroup = {
  id: string;
  label: string;
  items: NavItem[];
};

export const NAV_GROUPS: NavGroup[] = [
  {
    id: "ops",
    label: "Operasyon",
    items: [
      { href: "/", label: "Dashboard", icon: LayoutDashboard, keywords: ["ana", "özet"] },
      { href: "/heartbeats", label: "Nabız İzleme", icon: Radio, keywords: ["heartbeat", "mac"] },
      { href: "/komutlar", label: "Uzak Komutlar", icon: Terminal, keywords: ["komut", "kuyruk"] },
      { href: "/modul-ciktilari", label: "Modül Çıktıları", icon: Upload, keywords: ["output"] },
      { href: "/cihazlar", label: "Cihaz Bilgileri", icon: Monitor, keywords: ["donanım", "snapshot"] },
      { href: "/loglar", label: "Loglar", icon: ScrollText, keywords: ["aktivite", "denetim"] },
      { href: "/analitik", label: "Analitik", icon: BarChart3, keywords: ["istatistik"] },
    ],
  },
  {
    id: "config",
    label: "Yapılandırma",
    items: [
      { href: "/lisanslar", label: "Lisanslar", icon: KeyRound, keywords: ["license", "mac"] },
      { href: "/firma-modulleri", label: "Oto. Modüller", icon: Zap, keywords: ["firma", "auto-start"] },
      { href: "/klasor-izleme", label: "Klasör İzleme", icon: FolderSearch, keywords: ["watch", "folder"] },
    ],
  },
  {
    id: "dev",
    label: "Geliştirici",
    items: [
      { href: "/moduller", label: "Uzak Modüller", icon: Package, keywords: ["vba", "modül"] },
      { href: "/oneriler", label: "Modül Önerileri", icon: Lightbulb, keywords: ["proposal"] },
      { href: "/api-referans", label: "API Referans", icon: Plug, keywords: ["endpoint", "rest"] },
      { href: "/kurulum", label: "Kurulum Rehberi", icon: BookOpen, keywords: ["setup", "xlam", "agent"] },
    ],
  },
];

export const ALL_NAV_ITEMS = NAV_GROUPS.flatMap((g) => g.items);

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
