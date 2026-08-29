"use client";

import { useState } from "react";
import Link from "next/link";
import { Check } from "lucide-react";
import clsx from "clsx";

type Currency = "inr" | "usd";
type Period = "monthly" | "yearly";

const plans = [
  {
    id: "free",
    name: "Free",
    tagline: "Try it with no commitment",
    price: { inr: 0, usd: 0 },
    unit: "",
    cta: "Start free",
    href: "/editor/docs",
    features: [
      "Docs + Sheets editors",
      "Up to 3 active documents",
      "Import & export .docx / .xlsx",
      "DocKaro watermark on PDF export",
      "Community support",
    ],
  },
  {
    id: "pro",
    name: "Pro",
    tagline: "For individuals and freelancers",
    price: { inr: 499, usd: 19 },
    unit: "/mo",
    cta: "Start Pro trial",
    href: "/editor/docs",
    highlight: true,
    features: [
      "Everything in Free",
      "Unlimited documents & spreadsheets",
      "No watermark on exports",
      "10 GB cloud storage",
      "Version history (30 days)",
      "Priority email support",
    ],
  },
  {
    id: "business",
    name: "Business",
    tagline: "For teams and agencies",
    price: { inr: 399, usd: 15 },
    unit: "/user/mo",
    note: "3 user minimum",
    cta: "Start Business trial",
    href: "/editor/docs",
    features: [
      "Everything in Pro",
      "Shared team workspace",
      "Admin roles & permissions",
      "Unlimited version history",
      "Priority chat support",
    ],
  },
  {
    id: "api",
    name: "API",
    tagline: "For developers embedding DocKaro",
    price: { inr: 1999, usd: 49 },
    unit: "/mo",
    note: "+ usage above included quota",
    cta: "Get API access",
    href: "/api-docs",
    features: [
      "REST API — create, edit, export",
      "Embeddable editor SDK",
      "10,000 API calls / month included",
      "Webhooks for document events",
      "Dedicated developer support",
    ],
  },
];

const inr = (n: number) => `₹${n.toLocaleString("en-IN")}`;
const usd = (n: number) => `$${n}`;

export default function PricingSection() {
  const [currency, setCurrency] = useState<Currency>("inr");
  const [period, setPeriod] = useState<Period>("monthly");
  const yearlyMultiplier = 10; // 2 months free on yearly

  return (
    <div>
      <div className="flex flex-col items-center justify-center gap-4 sm:flex-row sm:gap-6">
        <div className="inline-flex rounded-lg border border-border bg-surface p-1 text-sm">
          {(["monthly", "yearly"] as Period[]).map((p) => (
            <button
              key={p}
              onClick={() => setPeriod(p)}
              className={clsx(
                "rounded-md px-4 py-1.5 capitalize transition-colors",
                period === p ? "bg-white text-black" : "text-muted hover:text-foreground",
              )}
            >
              {p}
              {p === "yearly" && (
                <span className="ml-1.5 text-[10px] text-accent">2 mo free</span>
              )}
            </button>
          ))}
        </div>
        <div className="inline-flex rounded-lg border border-border bg-surface p-1 text-sm">
          {(["inr", "usd"] as Currency[]).map((c) => (
            <button
              key={c}
              onClick={() => setCurrency(c)}
              className={clsx(
                "rounded-md px-4 py-1.5 uppercase transition-colors",
                currency === c ? "bg-white text-black" : "text-muted hover:text-foreground",
              )}
            >
              {c}
            </button>
          ))}
        </div>
      </div>

      <div className="mt-10 grid gap-6 lg:grid-cols-4">
        {plans.map((plan) => {
          const rawPrice = plan.price[currency];
          const shown =
            period === "yearly" && rawPrice > 0
              ? Math.round((rawPrice * yearlyMultiplier) / 12)
              : rawPrice;
          const format = currency === "inr" ? inr : usd;

          return (
            <div
              key={plan.id}
              className={clsx(
                "flex flex-col rounded-2xl border p-7",
                plan.highlight
                  ? "border-accent bg-accent/[0.06] ring-1 ring-accent"
                  : "border-border bg-surface",
              )}
            >
              {plan.highlight && (
                <span className="mb-3 w-fit rounded-full bg-accent px-2.5 py-0.5 text-[10px] font-medium uppercase tracking-wide text-white">
                  Most popular
                </span>
              )}
              <h3 className="text-lg font-medium">{plan.name}</h3>
              <p className="mt-1 text-sm text-muted">{plan.tagline}</p>

              <div className="mt-6 flex items-baseline gap-1">
                <span className="text-3xl font-semibold tracking-tight">
                  {format(shown)}
                </span>
                {plan.unit && <span className="text-sm text-muted">{plan.unit}</span>}
              </div>
              {plan.note && <p className="mt-1 text-xs text-muted">{plan.note}</p>}
              {period === "yearly" && rawPrice > 0 && (
                <p className="mt-1 text-xs text-muted">billed annually</p>
              )}

              <Link
                href={plan.href}
                className={clsx(
                  "mt-6 rounded-lg px-4 py-2.5 text-center text-sm font-medium transition-opacity hover:opacity-90",
                  plan.highlight ? "bg-white text-black" : "border border-border text-foreground",
                )}
              >
                {plan.cta}
              </Link>

              <ul className="mt-7 space-y-3 text-sm">
                {plan.features.map((f) => (
                  <li key={f} className="flex items-start gap-2 text-muted">
                    <Check size={16} className="mt-0.5 shrink-0 text-accent" />
                    {f}
                  </li>
                ))}
              </ul>
            </div>
          );
        })}
      </div>
    </div>
  );
}
