import type { Metadata } from "next";
import Navbar from "@/components/Navbar";
import Footer from "@/components/Footer";

export const metadata: Metadata = {
  title: "Terms of Service",
  alternates: { canonical: "/legal/terms" },
};

export default function TermsPage() {
  return (
    <>
      <Navbar />
      <main className="flex-1">
        <section className="mx-auto max-w-2xl px-6 py-20">
          <h1 className="text-3xl font-semibold tracking-tight">Terms of Service</h1>
          <p className="mt-2 text-sm text-muted">Draft — last updated 24 August 2026</p>

          <div className="prose prose-invert mt-10 max-w-none space-y-6 text-sm leading-relaxed text-muted">
            <p>
              This is a starting template, not a reviewed legal document — have
              a lawyer adapt it to your entity, jurisdiction and actual billing
              terms before relying on it publicly.
            </p>
            <div>
              <h2 className="text-base font-medium text-foreground">Using DocKaro</h2>
              <p>
                You may use DocKaro to create, edit and store documents,
                spreadsheets and presentations, and to access the API on plans
                that include it. You&apos;re responsible for content you upload
                and for keeping your API keys secret.
              </p>
            </div>
            <div>
              <h2 className="text-base font-medium text-foreground">Billing</h2>
              <p>
                Paid plans renew automatically for the billing period you
                selected (monthly or yearly) until cancelled. You can cancel
                anytime from account settings; access continues until the end
                of the paid period.
              </p>
            </div>
            <div>
              <h2 className="text-base font-medium text-foreground">Acceptable use</h2>
              <p>
                No uploading unlawful content, no attempting to break the
                service or other users&apos; accounts, no reselling API access
                without a partner agreement.
              </p>
            </div>
            <div>
              <h2 className="text-base font-medium text-foreground">Contact</h2>
              <p>
                Questions about these terms: <a href="mailto:hello@dockaro.com" className="text-accent">hello@dockaro.com</a>
              </p>
            </div>
          </div>
        </section>
      </main>
      <Footer />
    </>
  );
}
