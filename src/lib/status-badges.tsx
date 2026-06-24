import Badge from "@/app/components/ui/Badge";

export function isLicenseActive(value: string) {
  return ["true", "1", "active", "evet"].includes(value.toLowerCase());
}

export function LicenseBadge({ value }: { value: string }) {
  const active = isLicenseActive(value);
  return (
    <Badge variant={active ? "green" : "red"} dot>
      {active ? "Aktif" : "Pasif"}
    </Badge>
  );
}

const CMD_LABELS: Record<string, string> = {
  pending: "Bekliyor",
  running: "Çalışıyor",
  done: "Tamam",
  error: "Hata",
};

const CMD_VARIANT: Record<string, "yellow" | "blue" | "green" | "red"> = {
  pending: "yellow",
  running: "blue",
  done: "green",
  error: "red",
};

export function CommandStatusBadge({ status }: { status: string }) {
  return (
    <Badge variant={CMD_VARIANT[status] ?? "blue"}>
      {CMD_LABELS[status] ?? status}
    </Badge>
  );
}

export function HeartbeatDot({ status }: { status: "online" | "idle" | "offline" }) {
  const color =
    status === "online" ? "var(--green)" : status === "idle" ? "var(--yellow)" : "var(--red)";
  return (
    <span
      className="hb-dot"
      style={{ background: color }}
      title={status === "online" ? "Çevrimiçi" : status === "idle" ? "Boşta" : "Çevrimdışı"}
    />
  );
}
