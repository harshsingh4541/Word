"use client";

import { useEffect, useRef, useState } from "react";

/**
 * The document name in the title bar, renamed in place the way Word's own
 * title does. The name is stored on the document itself, so it also becomes
 * the exported file's name and its Word document title.
 */
export default function DocumentTitle({
  name,
  onRename,
}: {
  name: string;
  onRename: (name: string) => void;
}) {
  const [draft, setDraft] = useState(name);
  const [editing, setEditing] = useState(false);
  const inputRef = useRef<HTMLInputElement>(null);

  // The draft is seeded when editing starts, so the field always opens on
  // the document's current name without shadowing what is being typed.
  const startEditing = () => {
    setDraft(name);
    setEditing(true);
  };

  useEffect(() => {
    if (editing) inputRef.current?.select();
  }, [editing]);

  const commit = () => {
    setEditing(false);
    const next = draft.trim();
    if (!next) {
      setDraft(name);
      return;
    }
    if (next !== name) onRename(next);
  };

  if (!editing) {
    return (
      <button
        type="button"
        onClick={startEditing}
        title="Rename document"
        className="max-w-[40vw] truncate rounded px-2 py-0.5 text-sm text-word-text transition-colors hover:bg-black/5"
      >
        {name}
      </button>
    );
  }

  return (
    <input
      ref={inputRef}
      value={draft}
      onChange={(event) => setDraft(event.target.value)}
      onBlur={commit}
      onKeyDown={(event) => {
        if (event.key === "Enter") commit();
        if (event.key === "Escape") {
          setDraft(name);
          setEditing(false);
        }
      }}
      className="w-56 rounded border border-word-accent bg-white px-2 py-0.5 text-sm text-word-text outline-none"
    />
  );
}
