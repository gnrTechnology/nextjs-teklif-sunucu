import { TableSkeleton } from "@/app/components/ui/Skeleton";

export default function Loading() {
  return (
    <div className="page-loading">
      <TableSkeleton rows={8} cols={5} />
    </div>
  );
}
