import Link from "next/link";

const columns = [
  {
    title: "Product",
    links: [
      { href: "/editor/docs", label: "Docs" },
      { href: "/editor/sheets", label: "Sheets" },
      { href: "/#api", label: "API" },
      { href: "/pricing", label: "Pricing" },
    ],
  },
  {
    title: "Developers",
    links: [
      { href: "/api-docs", label: "API reference" },
      { href: "/#api", label: "Embed guide" },
      { href: "/status", label: "Status" },
    ],
  },
  {
    title: "Company",
    links: [
      { href: "/#faq", label: "FAQ" },
      { href: "mailto:hello@dockaro.com", label: "Contact" },
      { href: "/legal/privacy", label: "Privacy" },
      { href: "/legal/terms", label: "Terms" },
    ],
  },
];

export default function Footer() {
  return (
    <footer className="border-t border-border">
      <div className="mx-auto max-w-6xl px-6 py-14">
        <div className="grid grid-cols-2 gap-8 md:grid-cols-5">
          <div className="col-span-2">
            <Link href="/" className="flex items-center gap-2 font-semibold tracking-tight">
              <span className="flex h-7 w-7 items-center justify-center rounded-md bg-accent text-sm font-bold text-white">
                D
              </span>
              <span>DocKaro</span>
            </Link>
            <p className="mt-3 max-w-xs text-sm text-muted">
              Word, Sheets and Slides in your browser — and an API to put them
              inside your own product.
            </p>
          </div>
          {columns.map((col) => (
            <div key={col.title}>
              <h4 className="text-sm font-medium text-foreground">{col.title}</h4>
              <ul className="mt-3 space-y-2">
                {col.links.map((l) => (
                  <li key={l.label}>
                    <Link
                      href={l.href}
                      className="text-sm text-muted transition-colors hover:text-foreground"
                    >
                      {l.label}
                    </Link>
                  </li>
                ))}
              </ul>
            </div>
          ))}
        </div>
        <div className="mt-12 flex flex-col items-center justify-between gap-4 border-t border-border pt-6 text-xs text-muted sm:flex-row">
          <p>© {new Date().getFullYear()} DocKaro. All rights reserved.</p>
          <p>Made for creators, agencies and developers.</p>
        </div>
      </div>
    </footer>
  );
}
