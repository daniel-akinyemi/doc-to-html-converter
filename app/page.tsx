import Link from "next/link";
import { Converter } from "./components/Converter";
import { SiteFooter, SiteHeader } from "./components/SiteChrome";

export default function Home() {
  return (
    <div className="flex flex-col flex-1">
      <SiteHeader active="convert" />

      <main className="flex-1">
        <Hero />

        <section id="convert" className="border-t border-border">
          <div className="mx-auto max-w-5xl px-6 sm:px-10 py-16 sm:py-24">
            <Converter />
            <p className="mt-6 font-mono text-[10px] tracking-[0.2em] uppercase text-muted-foreground">
              Got a structured doc?{" "}
              <Link
                href="/split"
                className="underline underline-offset-4 decoration-foreground/30 hover:decoration-foreground transition-colors"
              >
                Split it into sections →
              </Link>
            </p>
          </div>
        </section>

        <Features />
      </main>

      <SiteFooter />
    </div>
  );
}

function Hero() {
  return (
    <section className="relative overflow-hidden">
      <div
        aria-hidden
        className="hairline-grid absolute inset-0 pointer-events-none opacity-60"
      />
      <div className="relative mx-auto max-w-5xl px-6 sm:px-10 pt-20 sm:pt-28 pb-16 sm:pb-24">
        <div className="font-mono text-[10px] tracking-[0.3em] uppercase text-muted-foreground mb-8">
          § Document → HTML · v1
        </div>

        <h1 className="font-serif text-5xl sm:text-7xl md:text-[5.5rem] leading-[0.95] tracking-tight max-w-4xl">
          Beautiful HTML
          <br />
          from your <em className="text-muted-foreground">Word documents.</em>
        </h1>

        <p className="mt-8 max-w-xl text-base sm:text-lg text-muted-foreground leading-relaxed">
          Drop in a <code className="font-mono text-sm">.docx</code> from
          Microsoft Word or Google Docs. Get clean, semantic HTML — preview it,
          copy it, ship it.
        </p>

        <div className="mt-10 flex flex-wrap items-center gap-x-8 gap-y-3 font-mono text-[10px] tracking-[0.25em] uppercase text-muted-foreground">
          <span className="flex items-center gap-2">
            <Dot /> 100% in your browser
          </span>
          <span className="flex items-center gap-2">
            <Dot /> No upload, no account
          </span>
          <span className="flex items-center gap-2">
            <Dot /> Semantic markup
          </span>
        </div>
      </div>
    </section>
  );
}

function Dot() {
  return (
    <span
      aria-hidden
      className="inline-block h-1.5 w-1.5 rounded-full bg-foreground"
    />
  );
}

function Features() {
  const items = [
    {
      n: "01",
      title: "Stays on your device",
      body: "Conversion runs entirely in the browser. Your documents are never uploaded — we couldn’t see them if we wanted to.",
    },
    {
      n: "02",
      title: "Semantic, not visual",
      body: "Headings, paragraphs, lists, tables, images, links, bold, italic, underline — converted to the right HTML tags so the structure holds up anywhere you paste it.",
    },
    {
      n: "03",
      title: "Convert, or split",
      body: "Turn a whole doc into one HTML block — or split a structured doc into labelled sections (overview, specs, sizing, care, FAQ), each ready to copy on its own.",
    },
  ];
  return (
    <section id="about" className="border-t border-border">
      <div className="mx-auto max-w-5xl px-6 sm:px-10 py-20 sm:py-28">
        <div className="font-mono text-[10px] tracking-[0.3em] uppercase text-muted-foreground mb-12">
          § How it works
        </div>
        <div className="grid gap-12 sm:gap-10 sm:grid-cols-3">
          {items.map((it) => (
            <article key={it.n} className="space-y-3">
              <div className="font-mono text-[10px] tracking-[0.25em] uppercase text-muted-foreground">
                — {it.n}
              </div>
              <h3 className="font-serif text-2xl leading-tight tracking-tight">
                {it.title}
              </h3>
              <p className="text-sm text-muted-foreground leading-relaxed">
                {it.body}
              </p>
            </article>
          ))}
        </div>
      </div>
    </section>
  );
}
