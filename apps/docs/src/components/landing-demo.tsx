"use client";

import * as React from "react";

import { useSelection, useSlide } from "@diceui/pptx";
import type { PresentationStore } from "@diceui/pptx";
import { cn } from "@pptx/ui/lib/utils";
import { FileArchive } from "lucide-react";

import { useDemoDeckStore } from "@/components/presentation-editing-demo";
import { PresentationEditor } from "@/components/presentation-editor";

export interface DeckPart {
  path: string;
  size: number;
}

type PartStatus = "unchanged" | "rewritten" | "added" | "removed";

type SourcePackage = NonNullable<
  NonNullable<ReturnType<PresentationStore["getState"]>["presentation"]>["sourcePackage"]
>;

const CONTENT_TYPES_PATH = "[Content_Types].xml";

const FEATURED_PARTS = new Set([
  "ppt/presentation.xml",
  "ppt/slides/slide1.xml",
  "ppt/slides/slide2.xml",
  "ppt/slides/slide3.xml",
  "ppt/theme/theme1.xml",
]);

const NO_CHANGES = new Map<string, PartChange>();

interface PartChange {
  status: Exclude<PartStatus, "unchanged">;
  size: number;
}

const DemoDeckContext = React.createContext<PresentationStore | null>(null);

function useDemoDeck() {
  const store = React.useContext(DemoDeckContext);
  if (!store) throw new Error("`useDemoDeck` must be used within `DemoDeckProvider`.");
  return store;
}

export function DemoDeckProvider({ children }: { children: React.ReactNode }) {
  const store = useDemoDeckStore();

  return <DemoDeckContext.Provider value={store}>{children}</DemoDeckContext.Provider>;
}

export function DemoDeckEditor({
  className,
  defaultSelectedNodeName,
}: {
  className?: string;
  defaultSelectedNodeName?: string;
}) {
  const store = useDemoDeck();

  return (
    <PresentationEditor store={store} className={className}>
      {defaultSelectedNodeName && <DefaultSelection nodeName={defaultSelectedNodeName} />}
    </PresentationEditor>
  );
}

function DefaultSelection({ nodeName }: { nodeName: string }) {
  const { slide } = useSlide();
  const { nodeIds, select } = useSelection();
  const isDoneRef = React.useRef(false);
  const nodeId = slide?.nodes.find((node) => node.name === nodeName)?.id;

  React.useEffect(() => {
    if (isDoneRef.current || !nodeId) return;
    if (nodeIds.includes(nodeId)) {
      isDoneRef.current = true;
      return;
    }
    select([nodeId]);
  }, [nodeId, nodeIds, select]);

  return null;
}

export function DemoDeckParts({ fileName, parts }: { fileName: string; parts: DeckPart[] }) {
  const store = useDemoDeck();
  const changes = usePartChanges(store, parts);
  const sourceRows = parts.filter(
    (part) => FEATURED_PARTS.has(part.path) || changes.has(part.path),
  );
  const rows = sourceRows
    .concat(
      [...changes]
        .filter(([, change]) => change.status === "added")
        .map(([path, change]) => ({ path, size: change.size })),
    )
    .sort((a, b) => comparePaths(a.path, b.path));
  const hiddenCount = parts.length - sourceRows.length;

  return (
    <div className="overflow-hidden rounded-xl border bg-card font-mono">
      <div className="flex items-center justify-between gap-4 border-b px-3.5 py-3.5 sm:px-5">
        <div className="flex min-w-0 items-center gap-2.5 text-[13px]">
          <FileArchive className="size-4 shrink-0 text-muted-foreground" />
          <span className="truncate">{fileName}</span>
        </div>
        <div className="hidden items-center gap-4 text-xs text-muted-foreground sm:flex">
          <span className="flex items-center gap-1.5">
            <span className="size-1.5 rounded-full bg-green-500" />
            rewritten
          </span>
          <span className="flex items-center gap-1.5">
            <span className="size-1.5 rounded-full bg-muted-foreground" />
            byte-identical
          </span>
        </div>
      </div>
      <ul className="py-2.5">
        {rows.map((row) => {
          const change = changes.get(row.path);
          const status = change?.status ?? "unchanged";
          const isChanged = status !== "unchanged";

          return (
            <li
              key={row.path}
              className={cn(
                "flex items-center gap-3 px-3.5 py-2 text-[13px] sm:px-5",
                isChanged && "bg-green-500/5",
              )}
            >
              <span
                className={cn(
                  "size-1.5 shrink-0 rounded-full",
                  isChanged ? "bg-green-500" : "bg-muted-foreground",
                )}
              />
              <span
                className={cn(
                  "min-w-0 flex-1 truncate",
                  isChanged ? "text-foreground" : "text-muted-foreground",
                  status === "removed" && "line-through",
                )}
              >
                {row.path}
              </span>
              <span
                className={cn(
                  "text-xs",
                  isChanged ? "text-green-600 dark:text-green-400" : "text-muted-foreground",
                )}
              >
                {status === "unchanged" ? (
                  <>
                    <span className="sm:hidden">same</span>
                    <span className="max-sm:hidden">unchanged</span>
                  </>
                ) : (
                  status
                )}
              </span>
              <span className="w-16 text-right text-xs text-muted-foreground max-sm:hidden">
                {formatSize(change && status !== "removed" ? change.size : row.size)}
              </span>
            </li>
          );
        })}
        <li className="px-3.5 py-2 text-xs text-muted-foreground sm:px-5">
          <span className="sm:hidden">+ {hiddenCount} more, unchanged</span>
          <span className="max-sm:hidden">+ {hiddenCount} more parts, unchanged</span>
        </li>
      </ul>
      <div className="flex items-center justify-between gap-4 border-t bg-background px-3.5 py-3.5 text-xs sm:px-5">
        <span className="text-muted-foreground">
          {changes.size} of {parts.length} parts changed
        </span>
        <span>
          <span className="max-sm:hidden">await </span>store.save()
        </span>
      </div>
    </div>
  );
}

function usePartChanges(store: PresentationStore, parts: DeckPart[]) {
  const sourcePackage = React.useSyncExternalStore(
    store.subscribe,
    () => store.getState().presentation?.sourcePackage,
    () => undefined,
  );
  const revision = React.useSyncExternalStore(
    store.subscribe,
    () => store.getState().revision,
    () => 0,
  );
  const [baseline, setBaseline] = React.useState<{
    sourcePackage: SourcePackage;
    bytes: Map<string, Uint8Array>;
  } | null>(null);
  const [result, setResult] = React.useState<{
    sourcePackage: SourcePackage;
    changes: Map<string, PartChange>;
  } | null>(null);

  React.useEffect(() => {
    if (!sourcePackage) return;
    let cancelled = false;

    void readReplaceableParts(sourcePackage).then((bytes) => {
      if (!cancelled) setBaseline({ sourcePackage, bytes });
    });

    return () => {
      cancelled = true;
    };
  }, [sourcePackage]);

  React.useEffect(() => {
    if (!sourcePackage || baseline?.sourcePackage !== sourcePackage || revision === 0) return;
    let cancelled = false;

    void diffParts(sourcePackage, parts, baseline.bytes).then((changes) => {
      if (!cancelled) setResult({ sourcePackage, changes });
    });

    return () => {
      cancelled = true;
    };
  }, [sourcePackage, parts, baseline, revision]);

  if (revision === 0 || !result || result.sourcePackage !== sourcePackage) return NO_CHANGES;
  return result.changes;
}

async function readReplaceableParts(sourcePackage: SourcePackage) {
  const bytes = new Map<string, Uint8Array>();
  for (const path of sourcePackage.paths()) {
    if (!isReplaceable(path)) continue;
    const content = await sourcePackage.readBytes(path);
    if (content) bytes.set(path, content);
  }
  return bytes;
}

async function diffParts(
  sourcePackage: SourcePackage,
  parts: DeckPart[],
  baseline: Map<string, Uint8Array>,
) {
  const changes = new Map<string, PartChange>();
  const sourcePaths = new Set(parts.map((part) => part.path));
  const paths = sourcePackage.paths();

  for (const path of paths) {
    const isDirty = sourcePackage.isDirty(path);
    const isAdded = !sourcePaths.has(path);
    if (!isDirty && !isAdded && !isReplaceable(path)) continue;

    const content = await sourcePackage.readBytes(path);
    if (!content) continue;
    if (isAdded) {
      changes.set(path, { status: "added", size: content.length });
    } else if (isDirty || !getIsSameBytes(content, baseline.get(path))) {
      changes.set(path, { status: "rewritten", size: content.length });
    }
  }

  const currentPaths = new Set(paths);
  for (const part of parts) {
    if (!currentPaths.has(part.path)) {
      changes.set(part.path, { status: "removed", size: part.size });
    }
  }

  return changes;
}

// Edits replace these parts wholesale rather than mark them dirty.
function isReplaceable(path: string) {
  return path === CONTENT_TYPES_PATH || path.endsWith(".rels");
}

function comparePaths(a: string, b: string) {
  if (a === CONTENT_TYPES_PATH) return -1;
  if (b === CONTENT_TYPES_PATH) return 1;
  return a.localeCompare(b, "en", { numeric: true });
}

function formatSize(bytes: number) {
  if (bytes < 1024) return `${bytes} B`;
  return `${(bytes / 1024).toFixed(1)} KB`;
}

function getIsSameBytes(a: Uint8Array, b: Uint8Array | undefined) {
  if (!b || a.length !== b.length) return false;
  return a.every((byte, index) => byte === b[index]);
}
