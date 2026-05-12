"use client";

export type View = "preview" | "source";

export function ViewToggle({
  view,
  setView,
}: {
  view: View;
  setView: (v: View) => void;
}) {
  const items: { id: View; label: string }[] = [
    { id: "preview", label: "Preview" },
    { id: "source", label: "HTML" },
  ];
  return (
    <div className="flex border border-border">
      {items.map((it) => (
        <button
          key={it.id}
          type="button"
          onClick={() => setView(it.id)}
          aria-pressed={view === it.id}
          className={`font-mono text-[10px] tracking-[0.25em] uppercase px-3 py-2 transition-colors ${
            view === it.id ? "bg-foreground text-background" : "hover:bg-subtle"
          }`}
        >
          {it.label}
        </button>
      ))}
    </div>
  );
}
