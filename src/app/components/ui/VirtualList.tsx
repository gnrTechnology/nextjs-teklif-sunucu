"use client";

import { useCallback, useRef, useState } from "react";

/** Sabit yükseklikli satırlar için basit pencereleme */
export default function VirtualList<T>({
  items,
  rowHeight,
  height,
  renderRow,
  className = "",
}: {
  items: T[];
  rowHeight: number;
  height: number;
  renderRow: (item: T, index: number) => React.ReactNode;
  className?: string;
}) {
  const [scrollTop, setScrollTop] = useState(0);
  const ref = useRef<HTMLDivElement>(null);

  const onScroll = useCallback(() => {
    if (ref.current) setScrollTop(ref.current.scrollTop);
  }, []);

  const totalHeight = items.length * rowHeight;
  const start = Math.max(0, Math.floor(scrollTop / rowHeight) - 2);
  const visible = Math.ceil(height / rowHeight) + 4;
  const end = Math.min(items.length, start + visible);
  const slice = items.slice(start, end);
  const offsetY = start * rowHeight;

  return (
    <div
      ref={ref}
      className={`virtual-list ${className}`.trim()}
      style={{ height, overflow: "auto" }}
      onScroll={onScroll}
    >
      <div style={{ height: totalHeight, position: "relative" }}>
        <div style={{ transform: `translateY(${offsetY}px)` }}>
          {slice.map((item, i) => (
            <div key={start + i} style={{ height: rowHeight }}>
              {renderRow(item, start + i)}
            </div>
          ))}
        </div>
      </div>
    </div>
  );
}
