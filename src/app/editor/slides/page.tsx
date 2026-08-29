import type { Metadata } from "next";
import Link from "next/link";
import { Presentation } from "lucide-react";
import EditorTopBar from "@/components/editors/EditorTopBar";

export const metadata: Metadata = {
  title: "Slides — Coming soon",
};

export default function SlidesEditorPage() {
  return (
    <>
      <EditorTopBar active="slides" />
      <div className="flex flex-1 flex-col items-center justify-center gap-4 text-center">
        <div className="flex h-14 w-14 items-center justify-center rounded-2xl bg-accent/15 text-accent">
          <Presentation size={26} />
        </div>
        <h1 className="text-2xl font-semibold tracking-tight">Slides is in the oven</h1>
        <p className="max-w-sm text-sm text-muted">
          The presentation editor is under active development. Join the
          waitlist and we&apos;ll email you the moment it&apos;s ready.
        </p>
        <Link
          href="/pricing"
          className="mt-2 rounded-lg bg-white px-5 py-2.5 text-sm font-medium text-black transition-opacity hover:opacity-90"
        >
          View plans
        </Link>
      </div>
    </>
  );
}
