import type { Metadata } from "next";
import Navbar from "@/components/Navbar";
import Footer from "@/components/Footer";
import PricingSection from "@/components/PricingSection";

export const metadata: Metadata = {
  title: "Pricing",
  description:
    "Simple, transparent pricing for DocKaro — Free, Pro, Business and API plans. India pricing in INR, global pricing in USD.",
  alternates: { canonical: "/pricing" },
};

const faqs = [
  {
    q: "Can I switch plans later?",
    a: "Yes, upgrade or downgrade anytime from account settings. Changes are prorated automatically.",
  },
  {
    q: "Is there a free trial on paid plans?",
    a: "Pro and Business include a 14-day trial, no card required to start.",
  },
  {
    q: "What happens to my documents if I downgrade?",
    a: "Your documents stay yours and remain exportable. Some Pro-only features (like watermark-free export) simply turn back on if you upgrade again.",
  },
  {
    q: "Do you offer nonprofit or student pricing?",
    a: "Yes — email hello@dockaro.com with proof of status for 50% off Pro.",
  },
];

export default function PricingPage() {
  return (
    <>
      <Navbar />
      <main className="flex-1">
        <section className="bg-grid border-b border-border px-6 py-20 text-center">
          <h1 className="text-4xl font-semibold tracking-tight">
            Simple pricing, no per-seat traps
          </h1>
          <p className="mx-auto mt-4 max-w-xl text-muted">
            Start free. Upgrade when you need more documents, a team, or API
            access.
          </p>
        </section>

        <section className="mx-auto max-w-6xl px-6 py-20">
          <PricingSection />
        </section>

        <section className="mx-auto max-w-3xl px-6 pb-24">
          <h2 className="text-center text-2xl font-semibold tracking-tight">
            Pricing FAQ
          </h2>
          <div className="mt-8 divide-y divide-border rounded-2xl border border-border">
            {faqs.map((f) => (
              <div key={f.q} className="p-6">
                <h3 className="font-medium">{f.q}</h3>
                <p className="mt-2 text-sm leading-relaxed text-muted">{f.a}</p>
              </div>
            ))}
          </div>
        </section>
      </main>
      <Footer />
    </>
  );
}
