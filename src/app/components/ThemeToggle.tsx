"use client";

import { Moon, Sun, Monitor } from "lucide-react";
import { useTheme } from "./ThemeProvider";

export default function ThemeToggle() {
  const { theme, setTheme } = useTheme();
  const cycle = () => {
    setTheme(theme === "dark" ? "light" : theme === "light" ? "system" : "dark");
  };
  const Icon = theme === "dark" ? Moon : theme === "light" ? Sun : Monitor;
  const label = theme === "dark" ? "Koyu" : theme === "light" ? "Açık" : "Sistem";
  return (
    <button type="button" className="theme-toggle" onClick={cycle} title={`Tema: ${label}`}>
      <Icon size={16} />
    </button>
  );
}
