import Link from "next/link";

export default function HomePage() {
  return (
    <div className="flex flex-1 flex-col items-center justify-center gap-4 text-center">
      <h1 className="text-3xl font-bold">pptx</h1>
      <p className="text-fd-muted-foreground">Open-source PowerPoint renderer for the web.</p>
      <div className="flex gap-3">
        <Link
          href="/docs"
          className="rounded-md bg-fd-primary px-4 py-2 text-sm font-medium text-fd-primary-foreground"
        >
          Documentation
        </Link>
        <Link
          href="/playground"
          className="rounded-md border border-fd-border px-4 py-2 text-sm font-medium"
        >
          Playground
        </Link>
      </div>
    </div>
  );
}
