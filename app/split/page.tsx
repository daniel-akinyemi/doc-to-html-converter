import type { Metadata } from "next";
import { SiteFooter, SiteHeader } from "../components/SiteChrome";
import { Splitter } from "../components/Splitter";

export const metadata: Metadata = {
  title: "Folio · Split — section a structured doc into HTML blocks",
  description:
    "Drop in a structured .docx and get it split into Product overview, Specifications, Sizing & selection, Care & storage, and FAQ — each as clean, copyable HTML. Runs entirely in your browser.",
};

const SECTIONS = [
  "Product overview",
  "Specifications",
  "Sizing & selection",
  "Care & storage",
  "Frequently asked questions",
];

export default function SplitPage() {
  return (
    <div className="flex flex-col flex-1">
      <SiteHeader active="split" />

      <main className="flex-1">
        <section className="relative overflow-hidden">
          <div
            aria-hidden
            className="hairline-grid absolute inset-0 pointer-events-none opacity-60"
          />
          <div className="relative mx-auto max-w-5xl px-6 sm:px-10 pt-20 sm:pt-28 pb-12 sm:pb-16">
            <div className="font-mono text-[10px] tracking-[0.3em] uppercase text-muted-foreground mb-8">
              § Document → Sections
            </div>

            <h1 className="font-serif text-5xl sm:text-7xl leading-[0.95] tracking-tight max-w-3xl">
              Split a doc into <em className="text-muted-foreground">sections.</em>
            </h1>

            <p className="mt-8 max-w-xl text-base sm:text-lg text-muted-foreground leading-relaxed">
              Drop in a structured{" "}
              <code className="font-mono text-sm">.docx</code> — like a product
              buyer&apos;s guide — and get it broken into five labelled HTML
              sections you can copy one at a time. If the file holds several
              products, each one is detected and split on its own.
            </p>

            <ol className="mt-8 grid gap-1 font-mono text-[10px] tracking-[0.2em] uppercase text-muted-foreground sm:grid-cols-5">
              {SECTIONS.map((s, i) => (
                <li key={s}>
                  {String(i + 1).padStart(2, "0")} — {s}
                </li>
              ))}
            </ol>
          </div>
        </section>

        <section className="border-t border-border">
          <div className="mx-auto max-w-5xl px-6 sm:px-10 py-16 sm:py-24">
            <Splitter />
          </div>
        </section>
      </main>

      <SiteFooter />
    </div>
  );
}
