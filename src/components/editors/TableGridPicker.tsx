"use client";

import { useState } from "react";

// Word's Insert > Table opens a grid you drag across to pick the table's
// size, with a live "4 x 3 Table" caption above it. Univer only ships a
// dialog with two number inputs, so this is the missing half.

const MAX_COLUMNS = 10;
const MAX_ROWS = 8;

/** The value handed back to the menu, parsed by the ribbon into params. */
export function formatTableSize(rows: number, columns: number): string {
  return `${rows}x${columns}`;
}

export function parseTableSize(value: unknown): { rowCount: number; colCount: number } | null {
  const match = /^(\d+)x(\d+)$/.exec(String(value ?? ""));
  if (!match) return null;
  return { rowCount: Number(match[1]), colCount: Number(match[2]) };
}

export default function TableGridPicker({ onChange }: { onChange?: (value: string) => void }) {
  const [hover, setHover] = useState<{ row: number; column: number } | null>(null);

  return (
    <div className="w-max select-none p-2" onMouseLeave={() => setHover(null)}>
      <div className="mb-2 text-center text-xs text-word-muted">
        {hover ? `${hover.column + 1} × ${hover.row + 1} table` : "Insert table"}
      </div>
      <div className="flex flex-col gap-[3px]">
        {Array.from({ length: MAX_ROWS }, (_, row) => (
          <div key={row} className="flex gap-[3px]">
            {Array.from({ length: MAX_COLUMNS }, (_, column) => {
              const selected = hover !== null && row <= hover.row && column <= hover.column;
              return (
                <button
                  key={column}
                  type="button"
                  aria-label={`${column + 1} by ${row + 1} table`}
                  onMouseEnter={() => setHover({ row, column })}
                  onFocus={() => setHover({ row, column })}
                  onClick={() => onChange?.(formatTableSize(row + 1, column + 1))}
                  className={`h-4 w-4 rounded-[2px] border ${
                    selected ? "border-word-accent bg-word-accent/30" : "border-word-border bg-white"
                  }`}
                />
              );
            })}
          </div>
        ))}
      </div>
    </div>
  );
}
