import type { Metadata } from "next";
import Navbar from "@/components/Navbar";
import Footer from "@/components/Footer";

export const metadata: Metadata = {
  title: "API Reference",
  description:
    "DocKaro REST API reference — create, read, update and delete documents and spreadsheets, with embeddable editor URLs.",
  alternates: { canonical: "/api-docs" },
};

const endpoints = [
  {
    method: "POST",
    path: "/v1/documents",
    desc: "Create a new document or spreadsheet.",
    body: `{
  "type": "docx",       // "docx" | "xlsx"
  "title": "Invoice #1042"
}`,
    response: `{
  "id": "doc_9f2a1b7c3d4e5f60718",
  "type": "docx",
  "title": "Invoice #1042",
  "createdAt": "2026-08-24T09:12:00.000Z",
  "updatedAt": "2026-08-24T09:12:00.000Z",
  "editUrl": "https://dockaro.com/e/doc_9f2a1b7c3d4e5f60718"
}`,
  },
  {
    method: "GET",
    path: "/v1/documents",
    desc: "List all documents on your account.",
    response: `{ "data": [ { "id": "doc_9f2a...", "title": "Invoice #1042", ... } ] }`,
  },
  {
    method: "GET",
    path: "/v1/documents/:id",
    desc: "Retrieve a single document by ID.",
    response: `{ "id": "doc_9f2a...", "title": "Invoice #1042", ... }`,
  },
  {
    method: "PATCH",
    path: "/v1/documents/:id",
    desc: "Rename a document.",
    body: `{ "title": "Invoice #1042 (revised)" }`,
  },
  {
    method: "DELETE",
    path: "/v1/documents/:id",
    desc: "Permanently delete a document.",
    response: "204 No Content",
  },
];

const methodColor: Record<string, string> = {
  GET: "text-emerald-400",
  POST: "text-accent",
  PATCH: "text-amber-400",
  DELETE: "text-red-400",
};

export default function ApiDocsPage() {
  return (
    <>
      <Navbar />
      <main className="flex-1">
        <section className="bg-grid border-b border-border px-6 py-20 text-center">
          <h1 className="text-4xl font-semibold tracking-tight">API reference</h1>
          <p className="mx-auto mt-4 max-w-xl text-muted">
            Base URL: <code className="text-foreground">https://api.dockaro.com</code>
            {" · "}Authenticate with{" "}
            <code className="text-foreground">Authorization: Bearer dk_live_...</code>
          </p>
        </section>

        <section className="mx-auto max-w-3xl divide-y divide-border px-6 py-16">
          {endpoints.map((e) => (
            <div key={e.method + e.path} className="py-8 first:pt-0">
              <div className="flex items-center gap-3">
                <span className={`text-xs font-mono font-semibold ${methodColor[e.method]}`}>
                  {e.method}
                </span>
                <code className="text-sm">{e.path}</code>
              </div>
              <p className="mt-2 text-sm text-muted">{e.desc}</p>
              {e.body && (
                <div className="mt-4">
                  <p className="mb-1 text-xs uppercase tracking-wide text-muted">Body</p>
                  <pre className="overflow-x-auto rounded-lg border border-border bg-surface p-4 text-xs text-muted">
                    {e.body}
                  </pre>
                </div>
              )}
              {e.response && (
                <div className="mt-4">
                  <p className="mb-1 text-xs uppercase tracking-wide text-muted">Response</p>
                  <pre className="overflow-x-auto rounded-lg border border-border bg-surface p-4 text-xs text-muted">
                    {e.response}
                  </pre>
                </div>
              )}
            </div>
          ))}
        </section>

        <section className="mx-auto max-w-3xl px-6 pb-24">
          <div className="rounded-2xl border border-border bg-surface p-8 text-center">
            <h2 className="text-lg font-medium">Need a key?</h2>
            <p className="mt-2 text-sm text-muted">
              API access ships with the API plan.{" "}
              <a href="/pricing" className="text-accent">
                See pricing →
              </a>
            </p>
          </div>
        </section>
      </main>
      <Footer />
    </>
  );
}
