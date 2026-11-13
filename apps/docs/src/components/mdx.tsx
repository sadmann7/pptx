import { Kbd, KbdGroup } from "@pptx/ui/components/kbd";
import { cn } from "@pptx/ui/lib/utils";
import {
  createFileSystemGeneratorCache,
  createGenerator,
} from "fumadocs-typescript";
import { AutoTypeTable, type AutoTypeTableProps } from "fumadocs-typescript/ui";
import defaultMdxComponents from "fumadocs-ui/mdx";
import type { MDXComponents } from "mdx/types";
import dynamic from "next/dynamic";

import { PresentationEditingDemo } from "@/components/presentation-editing-demo";

const Changelogs = dynamic(() =>
  import("@/components/changelogs").then((mod) => ({
    default: mod.Changelogs,
  })),
);

const generator = createGenerator({
  cache: createFileSystemGeneratorCache(".next/fumadocs-typescript"),
});

export function getMDXComponents(components?: MDXComponents) {
  return {
    ...defaultMdxComponents,
    AutoTypeTable: (props: Partial<AutoTypeTableProps>) => (
      <AutoTypeTable {...props} generator={generator} />
    ),
    // Prose styles every <kbd>, and KbdGroup is one too, so both opt out.
    Kbd: (props: React.ComponentProps<typeof Kbd>) => (
      <Kbd {...props} className={cn("not-prose", props.className)} />
    ),
    KbdGroup: (props: React.ComponentProps<typeof KbdGroup>) => (
      <KbdGroup {...props} className={cn("not-prose", props.className)} />
    ),
    Changelogs: () => <Changelogs />,
    PresentationEditingDemo,
    ...components,
  } satisfies MDXComponents;
}

export const useMDXComponents = getMDXComponents;

declare global {
  type MDXProvidedComponents = ReturnType<typeof getMDXComponents>;
}
