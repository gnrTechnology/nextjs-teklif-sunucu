type BadgeVariant = "green" | "red" | "yellow" | "blue" | "neutral";

const VARIANT_CLASS: Record<BadgeVariant, string> = {
  green: "badge-green",
  red: "badge-red",
  yellow: "badge-yellow",
  blue: "badge-blue",
  neutral: "",
};

export default function Badge({
  children,
  variant = "neutral",
  dot = false,
  title,
  className = "",
}: {
  children: React.ReactNode;
  variant?: BadgeVariant;
  dot?: boolean;
  title?: string;
  className?: string;
}) {
  return (
    <span className={`badge ${VARIANT_CLASS[variant]} ${className}`.trim()} title={title}>
      {dot && <span className="badge-dot" />}
      {children}
    </span>
  );
}
