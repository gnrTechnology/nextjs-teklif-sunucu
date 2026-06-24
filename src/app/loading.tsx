import { Skeleton } from "@/app/components/ui/Skeleton";

export default function Loading() {
  return (
    <div className="page-loading">
      <Skeleton className="page-loading-title" />
      <Skeleton className="page-loading-sub" />
      <Skeleton style={{ height: 120, borderRadius: 10 }} />
      <Skeleton style={{ height: 280, borderRadius: 10 }} />
    </div>
  );
}
