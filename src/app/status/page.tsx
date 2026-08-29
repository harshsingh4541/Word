import type { Metadata } from "next";
import { CheckCircle2 } from "lucide-react";
import Navbar from "@/components/Navbar";
import Footer from "@/components/Footer";

export const metadata: Metadata = {
  title: "System Status",
  alternates: { canonical: "/status" },
};

const systems = [
  "Docs editor",
  "Sheets editor",
  "REST API",
  "Document export (docx/xlsx)",
  "Website",
];

export default function StatusPage() {
  return (
    <>
      <Navbar />
      <main className="flex-1">
        <section className="mx-auto max-w-2xl px-6 py-20">
          <h1 className="text-3xl font-semibold tracking-tight">System status</h1>
          <p className="mt-2 text-sm text-muted">
            No incidents reported. This page is a placeholder — wire it to
            your real uptime monitor before launch.
          </p>

          <div className="mt-10 divide-y divide-border rounded-2xl border border-border">
            {systems.map((s) => (
              <div key={s} className="flex items-center justify-between p-4">
                <span className="text-sm">{s}</span>
                <span className="flex items-center gap-1.5 text-xs text-emerald-400">
                  <CheckCircle2 size={14} /> Operational
                </span>
              </div>
            ))}
          </div>
        </section>
      </main>
      <Footer />
    </>
  );
}
