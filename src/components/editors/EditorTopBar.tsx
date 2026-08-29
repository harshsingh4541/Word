import Link from "next/link";
import { FileText, Sheet, Presentation, ArrowLeft } from "lucide-react";
import clsx from "clsx";

const tabs = [
  { href: "/editor/docs", label: "Docs", icon: FileText },
  { href: "/editor/sheets", label: "Sheets", icon: Sheet },
  { href: "/editor/slides", label: "Slides", icon: Presentation, soon: true },
];

/**
 * The editor title bar, in the same light Word chrome as the ribbon below
 * it: app switcher on the left, document name in the middle, actions on the
 * right. (It used to be a dark bar sitting directly on top of Univer's
 * light toolbar, which is exactly the seam this rework removes.)
 */
export default function EditorTopBar({
  active,
  center,
  right,
}: {
  active: "docs" | "sheets" | "slides";
  center?: React.ReactNode;
  right?: React.ReactNode;
}) {
  return (
    <div className="relative shrink-0 border-b border-word-border bg-word-chrome">
      <header className="flex h-9 items-center gap-3 overflow-x-auto px-2 text-sm">
        <Link
          href="/"
          title="Back to DocKaro"
          className="flex shrink-0 items-center gap-1.5 rounded px-1 py-1 text-word-muted transition-colors hover:bg-black/5 hover:text-word-text"
        >
          <ArrowLeft size={15} />
          <span className="flex h-5 w-5 items-center justify-center rounded bg-word-accent text-[11px] font-bold text-white">
            D
          </span>
        </Link>
        <div className="h-4 w-px shrink-0 bg-word-border" />
        <div className="flex shrink-0 items-center gap-0.5">
          {tabs.map((t) => (
            <Link
              key={t.href}
              href={t.soon ? "/pricing" : t.href}
              className={clsx(
                "flex shrink-0 items-center gap-1.5 rounded px-2 py-1 text-xs transition-colors",
                active === t.href.split("/").pop()
                  ? "bg-white text-word-text shadow-sm"
                  : "text-word-muted hover:bg-black/5 hover:text-word-text",
              )}
            >
              <t.icon size={13} />
              {t.label}
              {t.soon && <span className="text-[10px] text-word-accent">soon</span>}
            </Link>
          ))}
        </div>
        {center && <div className="mx-auto flex min-w-0 shrink items-center justify-center">{center}</div>}
        <div className={clsx("flex shrink-0 items-center gap-2", !center && "ml-auto")}>{right}</div>
      </header>
    </div>
  );
}
