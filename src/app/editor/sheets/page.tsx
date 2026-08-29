"use client";

import dynamic from "next/dynamic";
import EditorTopBar from "@/components/editors/EditorTopBar";

const SheetsEditor = dynamic(() => import("@/components/editors/SheetsEditor"), {
  ssr: false,
  loading: () => (
    <div className="flex flex-1 items-center justify-center text-sm text-word-muted">
      Loading editor…
    </div>
  ),
});

export default function SheetsEditorPage() {
  return (
    <>
      <EditorTopBar active="sheets" />
      <div className="word-sheets flex-1 overflow-hidden">
        <SheetsEditor />
      </div>
    </>
  );
}
