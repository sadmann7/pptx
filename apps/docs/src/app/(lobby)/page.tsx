import { Button } from "@pptx/ui/components/button";
import Link from "next/link";

import { siteConfig } from "@/lib/site";

export default function IndexPage() {
  return (
    <section className="container flex flex-col items-center justify-center gap-6 pt-6 pb-8 md:py-10">
      <div className="flex max-w-4xl flex-col items-center gap-4 text-balance">
        <h1
          className="animate-fade-up text-center text-4xl/tight leading-tight font-bold tracking-tighter opacity-0 md:text-5xl/tight"
          style={{ animationDelay: "0.20s", animationFillMode: "forwards" }}
        >
          {siteConfig.name}
        </h1>
        <p
          className="animate-fade-up text-center text-lg font-light opacity-0 md:text-xl"
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
