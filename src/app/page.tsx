import Link from "next/link";
import {
  FileText,
  Sheet,
  Presentation,
  Code2,
  Check,
  X,
  Zap,
  Globe2,
  ShieldCheck,
  ArrowRight,
} from "lucide-react";
import Navbar from "@/components/Navbar";
import Footer from "@/components/Footer";

const products = [
  {
    icon: FileText,
    name: "Docs",
    tagline: "A full Word replacement",
    desc: "Rich formatting, deep table control, headers & footers, styles, real .docx import and export — not a stripped-down text box.",
    href: "/editor/docs",
  },
  {
    icon: Sheet,
    name: "Sheets",
    tagline: "A full Excel replacement",
    desc: "Formulas, conditional formatting, data validation, filters and sort — spreadsheets that behave like the ones you already know.",
    href: "/editor/sheets",
  },
  {
    icon: Presentation,
    name: "Slides",
    tagline: "Coming soon",
    desc: "A full PowerPoint replacement for decks and presentations. In active development — join the waitlist to get early access.",
    href: "/pricing",
    soon: true,
  },
];

const comparison = [
  { feature: "Real-time editing in the browser", us: true, google: true, ms: false, oo: true },
  { feature: "Embeddable via API", us: true, google: false, ms: false, oo: true },
  { feature: "India pricing", us: true, google: false, ms: false, oo: false },
  { feature: "No per-seat license required", us: true, google: false, ms: false, oo: false },
  { feature: "Self-serve signup, no sales call", us: true, google: true, ms: true, oo: false },
  { feature: "Full .docx / .xlsx fidelity", us: true, google: false, ms: true, oo: true },
];

const faqs = [
  {
    q: "Is DocKaro really free to start?",
    a: "Yes. The Free plan gives you unlimited documents with core editing in Docs and Sheets, no credit card required.",
  },
  {
    q: "Can I open my existing Word and Excel files?",
    a: "Yes — upload a .docx or .xlsx file, edit it in the browser, and export back to the same format with formatting preserved.",
  },
  {
    q: "Do you have an API?",
    a: "Yes. The Business and API plans include a REST API and embeddable editor SDK so you can add Docs, Sheets and Slides to your own product.",
  },
  {
    q: "How is this cheaper than Microsoft 365?",
    a: "There's no per-seat desktop license and no India-specific markup — you pay for the plan tier, not per install.",
  },
];

export default function Home() {
  return (
    <>
      <Navbar />
      <main className="flex-1">
        {/* Hero */}
        <section className="bg-grid relative overflow-hidden border-b border-border">
          <div className="pointer-events-none absolute inset-0 bg-gradient-to-b from-accent/10 via-transparent to-background" />
          <div className="relative mx-auto max-w-6xl px-6 py-28 text-center">
            <span className="inline-flex items-center gap-2 rounded-full border border-border bg-surface px-3 py-1 text-xs text-muted">
              <Zap size={14} className="text-accent" />
              Now in beta — free for early users
            </span>
            <h1 className="text-balance mx-auto mt-6 max-w-3xl text-5xl font-semibold tracking-tight sm:text-6xl">
              Word, Sheets and Slides.
              <br />
              <span className="text-muted">Right in your browser.</span>
            </h1>
            <p className="text-balance mx-auto mt-6 max-w-xl text-lg text-muted">
              A complete office suite with a developer API — build documents,
              spreadsheets and presentations into your own product, or just use
              ours.
            </p>
            <div className="mt-10 flex flex-col items-center justify-center gap-3 sm:flex-row">
              <Link
                href="/editor/docs"
                className="flex items-center gap-2 rounded-lg bg-white px-6 py-3 text-sm font-medium text-black transition-opacity hover:opacity-90"
              >
                Start free <ArrowRight size={16} />
              </Link>
              <Link
                href="/pricing"
                className="rounded-lg border border-border px-6 py-3 text-sm font-medium text-foreground transition-colors hover:bg-surface"
              >
                See pricing
              </Link>
            </div>
            <p className="mt-4 text-xs text-muted">
              No credit card required · Free plan forever
            </p>
          </div>
        </section>

        {/* Product grid */}
        <section id="product" className="mx-auto max-w-6xl px-6 py-24">
          <div className="mb-14 text-center">
            <h2 className="text-3xl font-semibold tracking-tight">
              Three editors. One account.
            </h2>
            <p className="mt-3 text-muted">
              Everything opens and saves in the format you already use.
            </p>
          </div>
          <div className="grid gap-6 md:grid-cols-3">
            {products.map((p) => (
              <Link
                key={p.name}
                href={p.href}
                className="group relative flex flex-col rounded-2xl border border-border bg-surface p-7 transition-colors hover:border-accent/50"
              >
                {p.soon && (
                  <span className="absolute right-5 top-5 rounded-full border border-border px-2 py-0.5 text-[10px] uppercase tracking-wide text-muted">
                    Soon
                  </span>
                )}
                <div className="flex h-11 w-11 items-center justify-center rounded-lg bg-accent/15 text-accent">
                  <p.icon size={22} />
                </div>
                <h3 className="mt-5 text-lg font-medium">{p.name}</h3>
                <p className="text-sm text-accent">{p.tagline}</p>
                <p className="mt-3 text-sm leading-relaxed text-muted">{p.desc}</p>
                <span className="mt-6 flex items-center gap-1 text-sm text-foreground opacity-0 transition-opacity group-hover:opacity-100">
                  Open editor <ArrowRight size={14} />
                </span>
              </Link>
            ))}
          </div>
        </section>

        {/* API section */}
        <section id="api" className="border-y border-border bg-surface/40">
          <div className="mx-auto grid max-w-6xl gap-12 px-6 py-24 md:grid-cols-2 md:items-center">
            <div>
              <div className="flex h-11 w-11 items-center justify-center rounded-lg bg-accent/15 text-accent">
                <Code2 size={22} />
              </div>
              <h2 className="mt-5 text-3xl font-semibold tracking-tight">
                Ship it inside your own app
              </h2>
              <p className="mt-4 text-muted">
                Every editor is available as an API and an embeddable SDK.
                Generate documents server-side, or drop a full editor into your
                product with one script tag.
              </p>
              <ul className="mt-6 space-y-3 text-sm">
                {[
                  "REST API for create, edit, export (.docx / .xlsx / .pdf)",
                  "Embeddable editor via iframe or JS SDK",
                  "Webhooks for document change events",
                  "Per-key usage limits and analytics",
                ].map((f) => (
                  <li key={f} className="flex items-start gap-2 text-muted">
                    <Check size={16} className="mt-0.5 shrink-0 text-accent" />
                    {f}
                  </li>
                ))}
              </ul>
              <Link
                href="/pricing"
                className="mt-8 inline-flex items-center gap-2 text-sm font-medium text-foreground"
              >
                View API pricing <ArrowRight size={14} />
              </Link>
            </div>
            <div className="overflow-hidden rounded-xl border border-border bg-[#0c0c0e] shadow-2xl">
              <div className="flex items-center gap-1.5 border-b border-border px-4 py-3">
                <span className="h-2.5 w-2.5 rounded-full bg-red-500/70" />
                <span className="h-2.5 w-2.5 rounded-full bg-yellow-500/70" />
                <span className="h-2.5 w-2.5 rounded-full bg-green-500/70" />
              </div>
              <pre className="overflow-x-auto p-5 text-xs leading-relaxed text-muted">
{`curl -X POST https://api.dockaro.com/v1/documents \\
  -H "Authorization: Bearer dk_live_xxx" \\
  -H "Content-Type: application/json" \\
  -d '{
    "type": "docx",
    "title": "Invoice #1042",
    "template": "invoice-basic"
  }'

# → 201 Created
{
  "id": "doc_9f2a...",
  "editUrl": "https://dockaro.com/e/doc_9f2a",
  "downloadUrl": "https://api.dockaro.com/v1/documents/doc_9f2a/export?format=docx"
}`}
              </pre>
            </div>
          </div>
        </section>

        {/* Comparison table */}
        <section className="mx-auto max-w-6xl px-6 py-24">
          <div className="mb-14 text-center">
            <h2 className="text-3xl font-semibold tracking-tight">
              How DocKaro compares
            </h2>
            <p className="mt-3 text-muted">
              Same core features, without the desktop license or the sales call.
            </p>
          </div>
          <div className="overflow-x-auto rounded-2xl border border-border">
            <table className="w-full min-w-[640px] text-sm">
              <thead>
                <tr className="border-b border-border bg-surface text-left text-muted">
                  <th className="px-5 py-4 font-medium">Feature</th>
                  <th className="px-5 py-4 text-center font-medium text-foreground">
                    DocKaro
                  </th>
                  <th className="px-5 py-4 text-center font-medium">
                    Google Workspace
                  </th>
                  <th className="px-5 py-4 text-center font-medium">
                    Microsoft 365
                  </th>
                  <th className="px-5 py-4 text-center font-medium">
                    ONLYOFFICE
                  </th>
                </tr>
              </thead>
              <tbody>
                {comparison.map((row, i) => (
                  <tr
                    key={row.feature}
                    className={i % 2 === 0 ? "bg-transparent" : "bg-surface/50"}
                  >
                    <td className="px-5 py-4 text-muted">{row.feature}</td>
                    {[row.us, row.google, row.ms, row.oo].map((v, idx) => (
                      <td key={idx} className="px-5 py-4 text-center">
                        {v ? (
                          <Check
                            size={16}
                            className={`mx-auto ${idx === 0 ? "text-accent" : "text-muted"}`}
                          />
                        ) : (
                          <X size={16} className="mx-auto text-muted/40" />
                        )}
                      </td>
                    ))}
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        </section>

        {/* Trust strip */}
        <section className="border-y border-border bg-surface/40">
          <div className="mx-auto grid max-w-6xl gap-8 px-6 py-16 text-center sm:grid-cols-3">
            <div className="flex flex-col items-center gap-2">
              <Globe2 size={20} className="text-accent" />
              <p className="text-sm text-muted">Global pricing in USD, INR pricing for India</p>
            </div>
            <div className="flex flex-col items-center gap-2">
              <ShieldCheck size={20} className="text-accent" />
              <p className="text-sm text-muted">Your documents, exportable anytime — no lock-in</p>
            </div>
            <div className="flex flex-col items-center gap-2">
              <Zap size={20} className="text-accent" />
              <p className="text-sm text-muted">Loads in the browser in under a second</p>
            </div>
          </div>
        </section>

        {/* FAQ */}
        <section id="faq" className="mx-auto max-w-3xl px-6 py-24">
          <h2 className="text-center text-3xl font-semibold tracking-tight">
            Frequently asked questions
          </h2>
          <div className="mt-10 divide-y divide-border rounded-2xl border border-border">
            {faqs.map((f) => (
              <div key={f.q} className="p-6">
                <h3 className="font-medium">{f.q}</h3>
                <p className="mt-2 text-sm leading-relaxed text-muted">{f.a}</p>
              </div>
            ))}
          </div>
        </section>

        {/* Final CTA */}
        <section className="mx-auto max-w-6xl px-6 pb-24">
          <div className="bg-grid relative overflow-hidden rounded-3xl border border-border px-8 py-16 text-center">
            <div className="pointer-events-none absolute inset-0 bg-gradient-to-b from-accent/15 via-transparent to-transparent" />
            <div className="relative">
              <h2 className="text-3xl font-semibold tracking-tight">
                Start writing in the next 10 seconds
              </h2>
              <p className="mx-auto mt-3 max-w-md text-muted">
                No install, no card, no waiting. Open the editor and go.
              </p>
              <Link
                href="/editor/docs"
                className="mt-8 inline-flex items-center gap-2 rounded-lg bg-white px-6 py-3 text-sm font-medium text-black transition-opacity hover:opacity-90"
              >
                Start free <ArrowRight size={16} />
              </Link>
            </div>
          </div>
        </section>
      </main>
      <Footer />
    </>
  );
}
