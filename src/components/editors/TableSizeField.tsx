"use client";

import { useState } from "react";

// Word's Table Layout tab sets row height and column width with a number
// box in inches, not a list of presets. This is that box, used inside the
// ribbon's Row height and Column width dropdowns.

const PX_PER_INCH = 96;

export function inchesToPixels(inches: number): number {
  return Math.round(inches * PX_PER_INCH);
}

export default function TableSizeField({
  onChange,
  title,
}: {
  onChange?: (value: string) => void;
  /** "Height" or "Width", as Word labels them. */
  title?: string;
}) {
  const [text, setText] = useState("");

  const commit = () => {
    const inches = Number.parseFloat(text.replace(/["\s]/g, ""));
    if (!Number.isFinite(inches) || inches <= 0) return;
    onChange?.(String(inchesToPixels(inches)));
  };

  return (
    <div className="flex items-center gap-2 px-2 py-1.5 text-xs text-word-text">
      <span className="text-word-muted">{title ?? "Size"}</span>
      <input
        value={text}
        autoFocus
        placeholder="0.5"
        aria-label={`${title ?? "Size"} in inches`}
        onChange={(event) => setText(event.target.value)}
        onKeyDown={(event) => {
          if (event.key === "Enter") {
            event.preventDefault();
            commit();
          }
          // Typing inside a menu must not be read as menu navigation.
          event.stopPropagation();
        }}
        className="w-16 rounded border border-word-border px-1.5 py-1 text-right outline-none focus:border-word-accent"
      />
      <span className="text-word-muted">in</span>
      <button
        type="button"
        onClick={commit}
        className="rounded bg-word-accent px-2 py-1 text-white transition-opacity hover:opacity-90"
      >
        Set
      </button>
    </div>
  );
}
