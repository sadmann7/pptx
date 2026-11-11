"use client";

import * as React from "react";

import { Button } from "@pptx/ui/components/button";
import { Kbd } from "@pptx/ui/components/kbd";
import { cn } from "@pptx/ui/lib/utils";
import { useSearchContext } from "fumadocs-ui/contexts/search";
import { Moon, Search, Sun } from "lucide-react";
import { useTheme } from "next-themes";
import Link from "next/link";

import { GitHubIcon } from "@/components/icons";
import { SiteLogo } from "@/lib/layout.shared";
import { siteConfig } from "@/lib/site";

const LINKS = [{ text: "Docs", href: "/docs" }];

export function SiteHeader({ className, ...props }: React.ComponentProps<"header">) {
  const { enabled, hotKey, setOpenSearch } = useSearchContext();
  const { resolvedTheme, setTheme } = useTheme();

  return (
    <header
      className={cn(
        "sticky top-0 z-40 flex h-14 items-center justify-between gap-6 border-b bg-background px-4 md:h-16 md:px-8",
        className,
      )}
      {...props}
    >
      <div className="flex items-center gap-10">
        <Link href="/" className="flex items-center gap-2">
          <SiteLogo />
        </Link>
        <nav className="flex items-center gap-7 text-sm text-muted-foreground max-sm:hidden">
          {LINKS.map((link) => (
            <Link
              key={link.href}
              href={link.href}
              className="transition-colors hover:text-foreground"
            >
              {link.text}
            </Link>
          ))}
        </nav>
      </div>
      <div className="flex items-center gap-1 md:gap-2">
        {enabled && (
          <>
            <button
              type="button"
              className="flex h-8.5 w-55 items-center gap-2 rounded-md border bg-card px-2.5 text-[13px] text-muted-foreground transition-colors hover:text-foreground max-lg:hidden"
              onClick={() => setOpenSearch(true)}
            >
              <Search className="size-3.5" />
              <span className="flex-1 text-left">Search docs…</span>
              <Kbd className="h-auto border bg-accent px-1.5 py-0.5 font-mono text-[11px] font-normal">
                {hotKey.map((key, index) => (
                  <React.Fragment key={index}>{key.display}</React.Fragment>
                ))}
              </Kbd>
            </button>
            <Button
              variant="ghost"
              size="icon-lg"
              aria-label="Search docs"
              className="lg:hidden [&_svg:not([class*='size-'])]:size-4.5"
              onClick={() => setOpenSearch(true)}
            >
              <Search />
            </Button>
          </>
        )}
        <Button
          variant="ghost"
          size="icon-lg"
          aria-label="Toggle theme"
          className="[&_svg:not([class*='size-'])]:size-4.5"
          onClick={() => setTheme(resolvedTheme === "dark" ? "light" : "dark")}
        >
          <Sun className="dark:hidden" />
          <Moon className="hidden dark:block" />
        </Button>
        <Button
          variant="ghost"
          size="icon-lg"
          aria-label="GitHub"
          nativeButton={false}
          render={
            <Link href={siteConfig.links.github} target="_blank" rel="noreferrer">
              <GitHubIcon className="size-4.5" />
            </Link>
          }
        />
      </div>
    </header>
  );
}
