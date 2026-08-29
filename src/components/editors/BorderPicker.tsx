"use client";

import { useEffect, useRef, useState } from "react";
import type { RefObject } from "react";
import {
  AllBorderIcon,
  DownBorderDoubleIcon,
  LeftBorderDoubleIcon,
  NoBorderIcon,
  RightBorderDoubleIcon,
  UpBorderDoubleIcon,
} from "@univerjs/icons";
import { BORDER_WEIGHTS, pointsToPixels } from "@/lib/univer/border-pen";
import type { BorderSide } from "@/lib/univer/table-style-commands";

// Word's Borders dropdown arms a pen (weight, colour) that only takes effect
// on a later Borders click, which here meant reopening the menu three times
// to draw one coloured border. This panel collapses that into a single visit:
// weight and colour are picked in place, and the side you finish on is the
// one click that dispatches the border command.

/** Word's "All borders" — every side of every selected cell. */
const ALL_SIDES: BorderSide[] = ["Top", "Bottom", "Left", "Right"];

const SIDE_CHOICES: {
  key: string;
  label: string;
  sides: BorderSide[];
  Icon: typeof AllBorderIcon;
}[] = [
  { key: "All", label: "All borders", sides: ALL_SIDES, Icon: AllBorderIcon },
  { key: "Top", label: "Top border", sides: ["Top"], Icon: UpBorderDoubleIcon },
  { key: "Bottom", label: "Bottom border", sides: ["Bottom"], Icon: DownBorderDoubleIcon },
  { key: "Left", label: "Left border", sides: ["Left"], Icon: LeftBorderDoubleIcon },
  { key: "Right", label: "Right border", sides: ["Right"], Icon: RightBorderDoubleIcon },
];

/**
 * The palette Univer's own colour picker shows: a tint-to-shade ramp per hue,
 * lightest row first. Kept here rather than reusing COLOR_PICKER_COMPONENT
 * because that one is registered as a whole menu entry — dropping it in would
 * bring back the separate trip through the ribbon this panel exists to avoid.
 */
const SWATCH_ROWS = [
  ["#FFFFFF", "#DEEBFF", "#FFE8E8", "#FFEEDD", "#FFF8D6", "#E4F7EC", "#DFF5F5", "#ECE8FD", "#FDE8F3"],
  ["#C9CDD4", "#A6C8FF", "#FFB3B3", "#FFCFA3", "#FFE066", "#93E0B7", "#A8E5E5", "#C7BFFA", "#F9B8DA"],
  ["#8F959E", "#3B82F6", "#F05252", "#FB6514", "#D9A406", "#22A06B", "#1D9C9C", "#8B5CF6", "#EC4899"],
  ["#4E5969", "#1D4ED8", "#C81E1E", "#B23C0E", "#8A6A04", "#16794C", "#146C6C", "#6D28D9", "#BE185D"],
  ["#000000", "#1E3A8A", "#7F1D1D", "#7C2D12", "#533F03", "#14532D", "#0F4C4C", "#4C1D95", "#831843"],
];

const DELIMITER = "|";

/** The single value handed to the menu, parsed by the ribbon into params. */
function formatBorderChoice(sideKey: string, widthPx: number, color: string): string {
  return [sideKey, widthPx, color].join(DELIMITER);
}

export function parseBorderChoice(
  value: unknown,
): { sides: BorderSide[]; width: number; color: string | null } | null {
  const [sideKey, width, color] = String(value ?? "").split(DELIMITER);

  // "No border" still has to write a border: the renderer falls back to a
  // default grey line for cells with no border property at all, so clearing
  // is a transparent zero-width line. Mirrors SetTableCellBorderCommand.
  if (sideKey === "None") return { sides: ALL_SIDES, width: 1, color: null };

  const choice = SIDE_CHOICES.find((candidate) => candidate.key === sideKey);
  const widthPx = Number(width);
  if (!choice || !Number.isFinite(widthPx) || !color) return null;
  return { sides: choice.sides, width: widthPx, color };
}

/**
 * Renders each weight as a line of that actual thickness, the way Word's
 * Line Weight gallery does — "0.5 pt" alone gives no sense of the result.
 * Hairlines round up to 1px so the thinnest option is still visible.
 */
function previewHeight(points: number): number {
  return Math.max(1, Math.round(pointsToPixels(points)));
}

/**
 * Keeps the dropdown open while weight and colour are being chosen.
 *
 * The panel renders inside a Radix menu item, and Radix closes the menu on
 * select unless its `menu.itemSelect` event is cancelled. Cancelling has to
 * happen on the item element itself: the event is dispatched there and
 * bubbles upward, so a listener inside the panel never sees it. Worse, Radix
 * synthesises the click by calling `.click()` on that same item element, so
 * the click doesn't originate in the panel either — which is why swallowing
 * pointer events inside the panel can't work.
 *
 * Applying still closes the menu, by a different route: calling `onChange`
 * runs Univer's own option handler, which hides the dropdown itself.
 */
function useKeepMenuOpen(ref: RefObject<HTMLDivElement | null>) {
  useEffect(() => {
    const item = ref.current?.closest('[role="menuitem"]');
    if (!item) return;
    const cancelSelect = (event: Event) => event.preventDefault();
    item.addEventListener("menu.itemSelect", cancelSelect);
    return () => item.removeEventListener("menu.itemSelect", cancelSelect);
  }, [ref]);
}

export default function BorderPicker({
  onChange,
  initialWidth,
  initialColor,
}: {
  onChange?: (value: string) => void;
  /** Current pen width in pixels, so the panel opens on the armed weight. */
  initialWidth?: number;
  initialColor?: string;
}) {
  const [points, setPoints] = useState(() => {
    const match = BORDER_WEIGHTS.find(
      (candidate) => Math.abs(pointsToPixels(candidate) - (initialWidth ?? 0)) < 0.01,
    );
    return match ?? 1;
  });
  const [color, setColor] = useState(initialColor ?? "#000000");
  const rootRef = useRef<HTMLDivElement>(null);
  useKeepMenuOpen(rootRef);

  const apply = (sideKey: string) => {
    onChange?.(formatBorderChoice(sideKey, pointsToPixels(points), color));
  };

  return (
    <div ref={rootRef} className="w-[228px] select-none py-1.5 text-xs text-word-text">
      <div>
        <div className="px-2.5 pb-1 text-word-muted">Line weight</div>
        <div className="flex flex-col px-1.5">
          {BORDER_WEIGHTS.map((candidate) => (
            <button
              key={candidate}
              type="button"
              aria-pressed={candidate === points}
              aria-label={`${candidate} point border`}
              onClick={() => setPoints(candidate)}
              className={`flex items-center gap-2 rounded px-1 py-[3px] text-left hover:bg-word-canvas ${
                candidate === points ? "bg-word-canvas" : ""
              }`}
            >
              <span className="w-10 shrink-0 whitespace-nowrap text-right text-word-muted">
                {candidate} pt
              </span>
              <span className="flex h-3.5 flex-1 items-center">
                <span
                  className="w-full rounded-full"
                  style={{ height: previewHeight(candidate), backgroundColor: color }}
                />
              </span>
            </button>
          ))}
        </div>

        <div className="mt-1.5 px-2.5 pb-1 text-word-muted">Pen color</div>
        <div className="flex flex-col gap-1 px-2.5">
          {SWATCH_ROWS.map((row, index) => (
            <div key={index} className="flex gap-1">
              {row.map((swatch) => (
                <button
                  key={swatch}
                  type="button"
                  aria-pressed={swatch.toLowerCase() === color.toLowerCase()}
                  aria-label={`Border color ${swatch}`}
                  title={swatch}
                  onClick={() => setColor(swatch)}
                  className={`h-[18px] w-[18px] rounded-full border ${
                    swatch.toLowerCase() === color.toLowerCase()
                      ? "border-word-accent ring-1 ring-word-accent"
                      : "border-word-border"
                  }`}
                  style={{ backgroundColor: swatch }}
                />
              ))}
            </div>
          ))}
        </div>
        <label className="mt-1.5 flex cursor-pointer items-center gap-2 px-2.5 py-1 text-word-muted">
          More colors
          <input
            type="color"
            value={color}
            aria-label="Custom border color"
            onChange={(event) => setColor(event.target.value)}
            className="h-[18px] w-7 cursor-pointer rounded border border-word-border bg-white"
          />
          <span className="ml-auto font-mono text-word-text">{color.toUpperCase()}</span>
        </label>
      </div>

      <div className="mt-1.5 border-t border-word-border px-2.5 pb-1 pt-2 text-word-muted">
        Apply to
      </div>
      <div className="flex flex-col px-1.5">
        {SIDE_CHOICES.map(({ key, label, Icon }) => (
          <button
            key={key}
            type="button"
            onClick={() => apply(key)}
            className="flex items-center gap-2 rounded px-1 py-1 text-left hover:bg-word-canvas"
          >
            <Icon className="shrink-0" />
            {label}
          </button>
        ))}
        <button
          type="button"
          onClick={() => apply("None")}
          className="flex items-center gap-2 rounded px-1 py-1 text-left hover:bg-word-canvas"
        >
          <NoBorderIcon className="shrink-0" />
          No border
        </button>
      </div>
    </div>
  );
}
