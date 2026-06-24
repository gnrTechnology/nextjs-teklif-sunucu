import Refresher from "../Refresher";

export default function PageHeader({
  title,
  subtitle,
  actions,
  live = false,
}: {
  title: string;
  subtitle?: string;
  actions?: React.ReactNode;
  live?: boolean;
}) {
  return (
    <div className="page-header">
      <div>
        <h1 className="page-title">{title}</h1>
        {subtitle && <p className="page-sub">{subtitle}</p>}
      </div>
      <div className="page-header-actions">
        {live && <Refresher />}
        {actions}
      </div>
    </div>
  );
}
