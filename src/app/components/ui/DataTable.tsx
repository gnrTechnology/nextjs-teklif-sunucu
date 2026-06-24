"use client";

import { useMemo, useState } from "react";
import { ChevronDown, ChevronUp, ChevronsUpDown } from "lucide-react";
import EmptyState from "./EmptyState";

export type Column<T> = {
  key: string;
  header: string;
  sortable?: boolean;
  width?: string;
  render: (row: T) => React.ReactNode;
  sortValue?: (row: T) => string | number;
};

type SortDir = "asc" | "desc";

export default function DataTable<T>({
  columns,
  data,
  pageSize = 25,
  emptyTitle = "Kayıt yok",
  emptyDescription,
  rowKey,
  rowClassName,
}: {
  columns: Column<T>[];
  data: T[];
  pageSize?: number;
  emptyTitle?: string;
  emptyDescription?: string;
  rowKey?: (row: T, index: number) => string | number;
  rowClassName?: (row: T) => string | undefined;
}) {
  const [sortKey, setSortKey] = useState<string | null>(null);
  const [sortDir, setSortDir] = useState<SortDir>("asc");
  const [page, setPage] = useState(0);

  const sorted = useMemo(() => {
    if (!sortKey) return data;
    const col = columns.find((c) => c.key === sortKey);
    if (!col?.sortValue) return data;
    const copy = [...data];
    copy.sort((a, b) => {
      const av = col.sortValue!(a);
      const bv = col.sortValue!(b);
      if (av < bv) return sortDir === "asc" ? -1 : 1;
      if (av > bv) return sortDir === "asc" ? 1 : -1;
      return 0;
    });
    return copy;
  }, [data, columns, sortKey, sortDir]);

  const totalPages = Math.max(1, Math.ceil(sorted.length / pageSize));
  const safePage = Math.min(page, totalPages - 1);
  const pageData = sorted.slice(safePage * pageSize, safePage * pageSize + pageSize);

  function toggleSort(key: string) {
    if (sortKey === key) setSortDir((d) => (d === "asc" ? "desc" : "asc"));
    else {
      setSortKey(key);
      setSortDir("asc");
    }
    setPage(0);
  }

  if (data.length === 0) {
    return <EmptyState title={emptyTitle} description={emptyDescription} />;
  }

  return (
    <div className="data-table-root">
      <div className="table-wrap">
        <table className="data-table">
          <thead>
            <tr>
              {columns.map((col) => (
                <th key={col.key} style={col.width ? { width: col.width } : undefined}>
                  {col.sortable ? (
                    <button type="button" className="th-sort" onClick={() => toggleSort(col.key)}>
                      {col.header}
                      {sortKey === col.key ? (
                        sortDir === "asc" ? <ChevronUp size={14} /> : <ChevronDown size={14} />
                      ) : (
                        <ChevronsUpDown size={14} className="th-sort-idle" />
                      )}
                    </button>
                  ) : (
                    col.header
                  )}
                </th>
              ))}
            </tr>
          </thead>
          <tbody>
            {pageData.map((row, i) => (
              <tr
                key={rowKey ? rowKey(row, i) : i}
                className={rowClassName?.(row)}
              >
                {columns.map((col) => (
                  <td key={col.key}>{col.render(row)}</td>
                ))}
              </tr>
            ))}
          </tbody>
        </table>
      </div>
      {sorted.length > pageSize && (
        <div className="table-pagination">
          <span className="table-pagination-info">
            {safePage * pageSize + 1}–{Math.min((safePage + 1) * pageSize, sorted.length)} / {sorted.length}
          </span>
          <div className="table-pagination-btns">
            <button type="button" className="btn btn-ghost" disabled={safePage === 0} onClick={() => setPage(0)}>
              İlk
            </button>
            <button type="button" className="btn btn-ghost" disabled={safePage === 0} onClick={() => setPage((p) => p - 1)}>
              Önceki
            </button>
            <span className="table-pagination-page">{safePage + 1} / {totalPages}</span>
            <button type="button" className="btn btn-ghost" disabled={safePage >= totalPages - 1} onClick={() => setPage((p) => p + 1)}>
              Sonraki
            </button>
            <button type="button" className="btn btn-ghost" disabled={safePage >= totalPages - 1} onClick={() => setPage(totalPages - 1)}>
              Son
            </button>
          </div>
        </div>
      )}
    </div>
  );
}
