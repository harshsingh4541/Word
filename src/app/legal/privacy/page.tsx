import type { Metadata } from "next";
import Navbar from "@/components/Navbar";
import Footer from "@/components/Footer";

export const metadata: Metadata = {
  title: "Privacy Policy",
  alternates: { canonical: "/legal/privacy" },
};

export default function PrivacyPage() {
  return (
    <>
      <Navbar />
      <main className="flex-1">
        <section className="mx-auto max-w-2xl px-6 py-20">
          <h1 className="text-3xl font-semibold tracking-tight">Privacy Policy</h1>
          <p className="mt-2 text-sm text-muted">Draft — last updated 24 August 2026</p>

          <div className="prose prose-invert mt-10 max-w-none space-y-6 text-sm leading-relaxed text-muted">
            <p>
              This is a starting template, not a reviewed legal document — have
              a lawyer check it against your actual data flows (storage
              provider, region, sub-processors, payment provider) before you
              rely on it publicly.
            </p>
            <div>
              <h2 className="text-base font-medium text-foreground">What we collect</h2>
              <p>
                Account details you provide (name, email), the documents,
                spreadsheets and presentations you create in DocKaro, and
                basic usage data (pages visited, API calls made) needed to
                operate and secure the service.
              </p>
            </div>
            <div>
              <h2 className="text-base font-medium text-foreground">How we use it</h2>
              <p>
                To provide the editor and API, to bill for paid plans, to
                respond to support requests, and to improve reliability. We
                do not sell your data or your documents&apos; contents to third
                parties.
              </p>
            </div>
            <div>
              <h2 className="text-base font-medium text-foreground">Your documents</h2>
              <p>
                Documents you create remain yours. You can export or delete
                them at any time. Deleting a document removes it from active
                storage; backups are purged on a rolling schedule.
              </p>
            </div>
            <div>
              <h2 className="text-base font-medium text-foreground">Contact</h2>
              <p>
                Questions about this policy: <a href="mailto:hello@dockaro.com" className="text-accent">hello@dockaro.com</a>
              </p>
            </div>
          </div>
        </section>
      </main>
      <Footer />
    </>
  );
}
