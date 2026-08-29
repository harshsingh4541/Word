"use client";

import { useState } from "react";
import Link from "next/link";
import { Menu, X } from "lucide-react";

const links = [
  { href: "/#product", label: "Product" },
  { href: "/#api", label: "API" },
  { href: "/pricing", label: "Pricing" },
  { href: "/#faq", label: "FAQ" },
];

export default function Navbar() {
  const [open, setOpen] = useState(false);

  return (
    <header className="sticky top-0 z-50 border-b border-border/80 bg-background/80 backdrop-blur-md">
      <nav className="mx-auto flex h-16 max-w-6xl items-center justify-between px-6">
        <Link href="/" className="flex items-center gap-2 font-semibold tracking-tight">
          <span className="flex h-7 w-7 items-center justify-center rounded-md bg-accent text-sm font-bold text-white">
            D
          </span>
          <span>DocKaro</span>
        </Link>

        <div className="hidden items-center gap-8 text-sm text-muted md:flex">
          {links.map((l) => (
            <Link key={l.href} href={l.href} className="transition-colors hover:text-foreground">
              {l.label}
            </Link>
          ))}
        </div>

        <div className="flex items-center gap-3">
          <Link
            href="/editor/docs"
            className="hidden text-sm text-muted transition-colors hover:text-foreground sm:block"
          >
            Sign in
          </Link>
          <Link
            href="/editor/docs"
            className="rounded-lg bg-white px-4 py-2 text-sm font-medium text-black transition-opacity hover:opacity-90"
          >
            Start free
          </Link>
          <button
            onClick={() => setOpen((v) => !v)}
            aria-label={open ? "Close menu" : "Open menu"}
            aria-expanded={open}
            className="flex h-9 w-9 items-center justify-center rounded-md text-muted transition-colors hover:text-foreground md:hidden"
          >
            {open ? <X size={20} /> : <Menu size={20} />}
          </button>
        </div>
      </nav>

      {open && (
        <div className="border-t border-border/80 px-6 py-3 md:hidden">
          <div className="flex flex-col gap-1 text-sm text-muted">
            {links.map((l) => (
              <Link
                key={l.href}
                href={l.href}
                onClick={() => setOpen(false)}
                className="rounded-md px-2 py-2.5 transition-colors hover:bg-white/5 hover:text-foreground"
              >
                {l.label}
              </Link>
            ))}
            <Link
              href="/editor/docs"
              onClick={() => setOpen(false)}
              className="rounded-md px-2 py-2.5 transition-colors hover:bg-white/5 hover:text-foreground"
            >
              Sign in
            </Link>
          </div>
        </div>
      )}
    </header>
  );
}
