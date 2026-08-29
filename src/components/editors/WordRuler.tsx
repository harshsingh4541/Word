"use client";

import { useEffect, useState } from "react";

// Word's horizontal ruler: the page's margins as shaded bands, inch
// markings measured out from the text area, and draggable markers for the
// paragraph's first-line, hanging and left/right indents. Univer draws no
// ruler at all, so this sits above its canvas and is kept aligned with it
// by reading the live page geometry every frame.

const PX_PER_INCH = 96;
/** Word snaps ruler drags to eighths of an inch. */
const SNAP_PX = PX_PER_INCH / 8;

export interface RulerGeometry {
  /** Left edge of the page in container pixels. */
  pageLeft: number;
  /** Top edge of the first page in container pixels. */
  pageTop: number;
  /** On-screen page width (already zoomed). */
  pageWidth: number;
  /** On-screen page height (already zoomed). */
  pageHeight: number;
  /** Page margins in document pixels. */
  marginTop: number;
  marginBottom: number;
  /** Page margins in document pixels. */
  marginLeft: number;
  marginRight: number;
  /** Paragraph indents at the cursor, in document pixels. */
  indentStart: number;
  indentEnd: number;
  indentFirstLine: number;
  /** Current zoom, as a multiplier. */
  scale: number;
}

export interface RulerHandlers {
  onIndentChange: (indents: { indentStart?: number; indentEnd?: number; indentFirstLine?: number }) => void;
  onMarginChange: (margins: { marginLeft?: number; marginRight?: number }) => void;
}

type Marker = "firstLine" | "hanging" | "leftIndent" | "rightIndent" | "leftMargin" | "rightMargin";

function snap(px: number): number {
  return Math.round(px / SNAP_PX) * SNAP_PX;
}

export default function WordRuler({
  getGeometry,
  handlers,
}: {
  getGeometry: () => RulerGeometry | null;
  handlers: RulerHandlers;
}) {
  const [geometry, setGeometry] = useState<RulerGeometry | null>(null);
  const [drag, setDrag] = useState<{ marker: Marker; startX: number; delta: number; base: RulerGeometry } | null>(null);

  // The page moves with scrolling, zooming, window resizes and page-setup
  // changes; sampling per frame keeps the ruler locked to it without
  // subscribing to four different Univer signals.
  useEffect(() => {
    let frame = 0;
    let previous = "";
    const tick = () => {
      const next = getGeometry();
      const key = next ? JSON.stringify(next) : "";
      if (key !== previous) {
        previous = key;
        setGeometry(next);
      }
      frame = requestAnimationFrame(tick);
    };
    frame = requestAnimationFrame(tick);
    return () => cancelAnimationFrame(frame);
  }, [getGeometry]);

  // While a marker is held, the window owns the drag so it survives the
  // pointer leaving the ruler.
  useEffect(() => {
    if (!drag) return;
    const onMove = (event: PointerEvent) => {
      setDrag((current) => (current ? { ...current, delta: event.clientX - current.startX } : current));
    };
    const onUp = (event: PointerEvent) => {
      setDrag(null);
      const { marker, base, startX } = drag;
      // Screen pixels are zoomed document pixels.
      const delta = (event.clientX - startX) / (base.scale || 1);
      if (Math.abs(delta) < 1) return;
      switch (marker) {
        case "firstLine":
          handlers.onIndentChange({ indentFirstLine: snap(base.indentFirstLine + delta) });
          break;
        case "hanging":
          handlers.onIndentChange({ indentStart: snap(base.indentStart + delta) });
          break;
        case "leftIndent":
          handlers.onIndentChange({
            indentStart: snap(base.indentStart + delta),
            indentFirstLine: snap(base.indentFirstLine + delta),
          });
          break;
        case "rightIndent":
          handlers.onIndentChange({ indentEnd: snap(base.indentEnd - delta) });
          break;
        case "leftMargin":
          handlers.onMarginChange({ marginLeft: snap(base.marginLeft + delta) });
          break;
        case "rightMargin":
          handlers.onMarginChange({ marginRight: snap(base.marginRight - delta) });
          break;
      }
    };
    window.addEventListener("pointermove", onMove);
    window.addEventListener("pointerup", onUp);
    return () => {
      window.removeEventListener("pointermove", onMove);
      window.removeEventListener("pointerup", onUp);
    };
  }, [drag, handlers]);

  if (!geometry) return <div className="h-6 shrink-0 border-b border-word-border bg-word-chrome" />;

  const { pageLeft, pageWidth, marginLeft, marginRight, indentStart, indentEnd, indentFirstLine, scale } = geometry;
  const toScreen = (documentPx: number) => pageLeft + documentPx * scale;
  const textLeft = toScreen(marginLeft);
  const textRight = toScreen(pageWidth / scale - marginRight);
  /** Live offset for the marker being dragged, so it follows the pointer. */
  const offsetFor = (marker: Marker) => (drag?.marker === marker ? drag.delta : 0);

  const startDrag = (marker: Marker) => (event: React.PointerEvent) => {
    event.preventDefault();
    setDrag({ marker, startX: event.clientX, delta: 0, base: geometry });
  };

  // Inch numbers count outwards from the text area, exactly as Word's do.
  const ticks: { x: number; label?: string }[] = [];
  const textWidth = textRight - textLeft;
  for (let px = 0; px <= textWidth; px += (PX_PER_INCH / 8) * scale) {
    const inches = px / (PX_PER_INCH * scale);
    const isInch = Math.abs(inches - Math.round(inches)) < 0.01;
    ticks.push({ x: textLeft + px, label: isInch && Math.round(inches) > 0 ? String(Math.round(inches)) : undefined });
  }

  return (
    <div className="relative h-6 shrink-0 overflow-hidden border-b border-word-border bg-word-chrome text-[9px] text-word-muted">
      {/* The page, with its margins shaded like Word's ruler. */}
      <div className="absolute inset-y-1 rounded-sm bg-[#c8c6c4]" style={{ left: pageLeft, width: pageWidth }} />
      <div className="absolute inset-y-1 bg-white" style={{ left: textLeft, width: Math.max(0, textRight - textLeft) }} />

      {ticks.map((tick, index) => (
        <div key={index} className="absolute" style={{ left: tick.x }}>
          {tick.label ? (
            <span className="absolute -translate-x-1/2 select-none" style={{ top: 5 }}>
              {tick.label}
            </span>
          ) : (
            <span className="absolute block h-1 w-px bg-word-muted/60" style={{ top: 10 }} />
          )}
        </div>
      ))}

      {/* Margin handles */}
      <button
        type="button"
        aria-label="Left margin"
        onPointerDown={startDrag("leftMargin")}
        className="absolute inset-y-0 w-2 cursor-col-resize"
        style={{ left: textLeft - 4 + offsetFor("leftMargin") }}
      />
      <button
        type="button"
        aria-label="Right margin"
        onPointerDown={startDrag("rightMargin")}
        className="absolute inset-y-0 w-2 cursor-col-resize"
        style={{ left: textRight - 4 + offsetFor("rightMargin") }}
      />

      {/* Indent markers: first line on top, hanging and left indent below. */}
      <button
        type="button"
        aria-label="First line indent"
        onPointerDown={startDrag("firstLine")}
        title="First line indent"
        className="absolute h-0 w-0 -translate-x-1/2 cursor-col-resize border-x-[5px] border-t-[6px] border-x-transparent border-t-word-accent"
        style={{ left: textLeft + indentFirstLine * scale + offsetFor("firstLine"), top: 2 }}
      />
      <button
        type="button"
        aria-label="Hanging indent"
        onPointerDown={startDrag("hanging")}
        title="Hanging indent"
        className="absolute h-0 w-0 -translate-x-1/2 cursor-col-resize border-x-[5px] border-b-[6px] border-x-transparent border-b-word-accent"
        style={{ left: textLeft + indentStart * scale + offsetFor("hanging"), top: 10 }}
      />
      <button
        type="button"
        aria-label="Left indent"
        onPointerDown={startDrag("leftIndent")}
        title="Left indent"
        className="absolute h-1.5 w-2.5 -translate-x-1/2 cursor-col-resize bg-word-accent"
        style={{ left: textLeft + indentStart * scale + offsetFor("leftIndent"), top: 17 }}
      />
      <button
        type="button"
        aria-label="Right indent"
        onPointerDown={startDrag("rightIndent")}
        title="Right indent"
        className="absolute h-0 w-0 -translate-x-1/2 cursor-col-resize border-x-[5px] border-b-[6px] border-x-transparent border-b-word-accent"
        style={{ left: textRight - indentEnd * scale + offsetFor("rightIndent"), top: 10 }}
      />
    </div>
  );
}
