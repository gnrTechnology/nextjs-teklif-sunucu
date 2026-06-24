import { Skeleton } from "@/app/components/ui/Skeleton";

export default function Loading() {
  return (
    <div className="page-loading">
      <Skeleton className="page-loading-title" />
      <Skeleton className="page-loading-sub" />
      <div className="device-grid">
        {[1, 2, 3].map((i) => (
          <Skeleton key={i} style={{ height: 140, borderRadius: 10 }} />
        ))}
      </div>
    </div>
  );
}
