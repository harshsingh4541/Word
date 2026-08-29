"use client";

import { Minus, Plus } from "lucide-react";

// Word's status bar, in Word's order: page position on the left, word
// count next to it, zoom controls on the right. Univer's own footer is
// switched off in DocsEditor so this is the only one on screen.
const ZOOM_STEPS = [50, 75, 100, 125, 150, 200];

export default function WordStatusBar({
  currentPage,
  pageCount,
  wordCount,
  zoom,
  onZoomChange,
}: {
  currentPage: number;
  pageCount: number;
  wordCount: number;
  zoom: number;
  onZoomChange: (zoom: number) => void;
}) {
  const step = (direction: -1 | 1) => {
    const index = ZOOM_STEPS.findIndex((preset) => preset >= zoom);
    const current = index === -1 ? ZOOM_STEPS.length - 1 : index;
    const next = ZOOM_STEPS[Math.min(ZOOM_STEPS.length - 1, Math.max(0, current + direction))];
    if (next !== zoom) onZoomChange(next);
  };

  return (
    <div className="flex h-6 shrink-0 items-center gap-4 border-t border-word-border bg-word-chrome px-3 text-[11px] text-word-muted">
      <span>
        Page {currentPage} of {pageCount}
      </span>
      <span>
        {wordCount} {wordCount === 1 ? "word" : "words"}
      </span>
      <div className="ml-auto flex items-center gap-1.5">
        <button
          type="button"
          aria-label="Zoom out"
          onClick={() => step(-1)}
          className="flex h-4 w-4 items-center justify-center rounded hover:bg-black/10"
        >
          <Minus size={11} />
        </button>
        <input
          type="range"
          aria-label="Zoom"
          min={50}
          max={200}
          step={5}
          value={zoom}
          onChange={(event) => onZoomChange(Number(event.target.value))}
          className="h-1 w-28 accent-word-accent"
        />
        <button
          type="button"
          aria-label="Zoom in"
          onClick={() => step(1)}
          className="flex h-4 w-4 items-center justify-center rounded hover:bg-black/10"
        >
          <Plus size={11} />
        </button>
        <span className="w-10 text-right tabular-nums">{zoom}%</span>
      </div>
    </div>
  );
}
