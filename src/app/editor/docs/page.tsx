"use client";

import { useRef, useState } from "react";
import dynamic from "next/dynamic";
import { Cloud } from "lucide-react";
import EditorTopBar from "@/components/editors/EditorTopBar";
import DocumentTitle from "@/components/editors/DocumentTitle";
import WordStatusBar from "@/components/editors/WordStatusBar";
import type { DocsEditorHandle, WordDocumentStatus } from "@/components/editors/DocsEditor";

const DocsEditor = dynamic(() => import("@/components/editors/DocsEditor"), {
  ssr: false,
  loading: () => (
    <div className="flex flex-1 items-center justify-center text-sm text-word-muted">Loading editor…</div>
  ),
});

const INITIAL_STATUS: WordDocumentStatus = {
  name: "Untitled document",
  wordCount: 0,
  pageCount: 1,
  currentPage: 1,
  zoom: 100,
};

export default function DocsEditorPage() {
  const apiRef = useRef<DocsEditorHandle | null>(null);
  const [status, setStatus] = useState<WordDocumentStatus>(INITIAL_STATUS);

  // Reflect a zoom change immediately: the document's own value only comes
  // back on the next (debounced) status refresh, and a slider that lags
  // behind the drag feels broken.
  const handleZoomChange = (zoom: number) => {
    setStatus((current) => ({ ...current, zoom }));
    apiRef.current?.setZoom(zoom);
  };

  return (
    <>
      <EditorTopBar
        active="docs"
        center={<DocumentTitle name={status.name} onRename={(name) => apiRef.current?.setName(name)} />}
        right={
          <span className="flex items-center gap-1.5 rounded px-2 py-1 text-xs text-word-muted" title="Saved in this browser">
            <Cloud size={13} /> Saved
          </span>
        }
      />
      <div className="word-docs flex min-h-0 flex-1 flex-col overflow-hidden bg-word-canvas">
        <DocsEditor apiRef={apiRef} onStatusChange={setStatus} />
      </div>
      <WordStatusBar
        currentPage={status.currentPage}
        pageCount={status.pageCount}
        wordCount={status.wordCount}
        zoom={status.zoom}
        onZoomChange={handleZoomChange}
      />
    </>
  );
}
