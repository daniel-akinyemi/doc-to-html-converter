import Link from "next/link";

export function SiteHeader({ active }: { active?: "convert" | "split" }) {
  const link = (href: string, label: string, key: "convert" | "split") =>
    (
      <Link
        href={href}
        className={`hover:text-foreground transition-colors ${
          active === key ? "text-foreground" : ""
        }`}
      >
        {label}
      </Link>
    );
  return (
    <header className="border-b border-border">
      <div className="mx-auto max-w-5xl px-6 sm:px-10 h-14 flex items-center justify-between">
        <Link href="/" className="font-mono text-[11px] tracking-[0.3em] uppercase">
          Folio
        </Link>
        <nav className="flex items-center gap-6 font-mono text-[10px] tracking-[0.25em] uppercase text-muted-foreground">
          {link("/", "Convert", "convert")}
          {link("/split", "Split", "split")}
        </nav>
      </div>
    </header>
  );
}

export function SiteFooter() {
  return (
    <footer className="border-t border-border">
      <div className="mx-auto max-w-5xl px-6 sm:px-10 h-14 flex items-center justify-between font-mono text-[10px] tracking-[0.25em] uppercase text-muted-foreground">
        <span>© Folio</span>
        <span>Made for writers, editors & engineers</span>
      </div>
    </footer>
  );
}
