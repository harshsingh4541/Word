"use client";

import { useEffect, useState } from "react";
import type { RulerGeometry } from "./WordRuler";

// Word's vertical ruler: the page's top and bottom margins as shaded bands
// down the left of the document view, with inch markings measured from the
// top margin and both margins draggable.

const PX_PER_INCH = 96;
const SNAP_PX = PX_PER_INCH / 8;

type Marker = "topMargin" | "bottomMargin";

function snap(px: number): number {
  return Math.round(px / SNAP_PX) * SNAP_PX;
}

export default function WordVerticalRuler({
  getGeometry,
  onMarginChange,
}: {
  getGeometry: () => RulerGeometry | null;
  onMarginChange: (margins: { marginTop?: number; marginBottom?: number }) => void;
}) {
  const [geometry, setGeometry] = useState<RulerGeometry | null>(null);
  const [drag, setDrag] = useState<{ marker: Marker; startY: number; delta: number; base: RulerGeometry } | null>(null);

  // Same per-frame sampling as the horizontal ruler: the page moves with
  // scrolling, zooming and page setup.
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

  useEffect(() => {
    if (!drag) return;
    const onMove = (event: PointerEvent) => {
      setDrag((current) => (current ? { ...current, delta: event.clientY - current.startY } : current));
    };
    const onUp = (event: PointerEvent) => {
      setDrag(null);
      const { marker, base, startY } = drag;
      const delta = (event.clientY - startY) / (base.scale || 1);
      if (Math.abs(delta) < 1) return;
      if (marker === "topMargin") onMarginChange({ marginTop: snap(base.marginTop + delta) });
      else onMarginChange({ marginBottom: snap(base.marginBottom - delta) });
    };
    window.addEventListener("pointermove", onMove);
    window.addEventListener("pointerup", onUp);
    return () => {
      window.removeEventListener("pointermove", onMove);
      window.removeEventListener("pointerup", onUp);
    };
  }, [drag, onMarginChange]);

  if (!geometry) return null;

  const { pageTop, pageHeight, marginTop, marginBottom, scale } = geometry;
  const textTop = pageTop + marginTop * scale;
  const textBottom = pageTop + pageHeight - marginBottom * scale;
  const offsetFor = (marker: Marker) => (drag?.marker === marker ? drag.delta : 0);

  const startDrag = (marker: Marker) => (event: React.PointerEvent) => {
    event.preventDefault();
    setDrag({ marker, startY: event.clientY, delta: 0, base: geometry });
  };

  // Inch numbers count from the top margin down, as Word's do.
  const ticks: { y: number; label?: string }[] = [];
  for (let px = 0; px <= textBottom - textTop; px += (PX_PER_INCH / 8) * scale) {
    const inches = px / (PX_PER_INCH * scale);
    const isInch = Math.abs(inches - Math.round(inches)) < 0.01;
    ticks.push({ y: textTop + px, label: isInch && Math.round(inches) > 0 ? String(Math.round(inches)) : undefined });
  }

  return (
    <div className="pointer-events-none absolute inset-y-0 left-0 z-10 w-6 overflow-hidden text-[9px] text-word-muted">
      <div className="absolute inset-x-1 rounded-sm bg-[#c8c6c4]" style={{ top: pageTop, height: pageHeight }} />
      <div className="absolute inset-x-1 bg-white" style={{ top: textTop, height: Math.max(0, textBottom - textTop) }} />

      {ticks.map((tick, index) => (
        <div key={index} className="absolute inset-x-0" style={{ top: tick.y }}>
          {tick.label ? (
            <span className="absolute left-1/2 -translate-x-1/2 -translate-y-1/2 select-none bg-white px-px">
              {tick.label}
            </span>
          ) : (
            <span className="absolute left-1/2 block h-px w-1 -translate-x-1/2 bg-word-muted/60" />
          )}
        </div>
      ))}

      <button
        type="button"
        aria-label="Top margin"
        onPointerDown={startDrag("topMargin")}
        className="pointer-events-auto absolute inset-x-0 h-2 cursor-row-resize"
        style={{ top: textTop - 4 + offsetFor("topMargin") }}
      />
      <button
        type="button"
        aria-label="Bottom margin"
        onPointerDown={startDrag("bottomMargin")}
        className="pointer-events-auto absolute inset-x-0 h-2 cursor-row-resize"
        style={{ top: textBottom - 4 + offsetFor("bottomMargin") }}
      />
    </div>
  );
}
