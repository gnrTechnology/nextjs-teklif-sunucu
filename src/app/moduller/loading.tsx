import { Skeleton } from "@/app/components/ui/Skeleton";

export default function Loading() {
  return (
    <div className="page-loading">
      <Skeleton className="page-loading-title" />
      {[1, 2, 3, 4, 5, 6].map((i) => (
        <Skeleton key={i} style={{ height: 56, borderRadius: 6 }} />
      ))}
    </div>
  );
}
