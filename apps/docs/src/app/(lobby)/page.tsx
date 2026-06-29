import { siteConfig } from "@/lib/site";
import { Button } from "@/components/ui/button";
import Link from "next/link";

export default function IndexPage() {
  return (
    <section className="container flex flex-col items-center justify-center gap-6 pt-6 pb-8 md:py-10">
      <div className="flex max-w-5xl flex-col items-center gap-4">
        <h1
          className="animate-fade-up bg-linear-to-br from-foreground/80 to-muted-foreground bg-clip-text text-center text-4xl/tight leading-tight font-bold tracking-tighter text-balance text-transparent opacity-0 drop-shadow-xs md:text-5xl/tight"
          style={{ animationDelay: "0.20s", animationFillMode: "forwards" }}
        >
          {siteConfig.name}
        </h1>
        <p
          className="max-w-2xl animate-fade-up text-center text-lg font-light text-balance text-muted-foreground opacity-0 md:text-xl"
          style={{ animationDelay: "0.30s", animationFillMode: "forwards" }}
        >
          {siteConfig.description}
        </p>
        <div
          className="flex animate-fade-up justify-center gap-4 pt-2 opacity-0"
          style={{ animationDelay: "0.40s", animationFillMode: "forwards" }}
        >
          <Button nativeButton={false} render={<Link href="/docs">Docs</Link>} />
          <Button
            variant="outline"
            nativeButton={false}
            render={
              <Link target="_blank" rel="noreferrer" href={siteConfig.links.github}>
                GitHub
              </Link>
            }
          />
        </div>
      </div>
    </section>
  );
}
