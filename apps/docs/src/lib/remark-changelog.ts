import { replaceSource } from "fumadocs-core/mdx-plugins/stringifier";
import type { Heading, Nodes, Paragraph, Root } from "mdast";

type File = Parameters<typeof replaceSource>[0];

const COMMIT_PREFIX = /^[0-9a-f]{7,40}: /;

export function remarkChangelog() {
  return (tree: Root, file: File) => {
    if (file.basename === "changelog.mdx") transform(tree, file);
  };
}

function transform(node: Nodes, file: File) {
  if (node.type === "heading") shiftHeading(node, file);
  if (node.type === "listItem") stripCommit(node.children[0], file);
  if (!("children" in node)) return;

  for (const child of node.children) transform(child, file);
}

// Changesets writes the package as h1, versions as h2, and change types as h3.
// Nodes are edited in place, as the stringifier tracks included ones by identity.
function shiftHeading(heading: Heading, file: File) {
  if (heading.depth < 3) {
    const depth = (heading.depth + 1) as Heading["depth"];
    heading.depth = depth;
    replaceSource(
      file,
      heading,
      (s) => `${"#".repeat(depth)} ${s.inner(heading)}`,
    );
    return;
  }

  const { children } = heading;
  const paragraph: Paragraph = Object.assign(heading, {
    type: "paragraph" as const,
    children: [{ type: "strong" as const, children }],
  });
  replaceSource(
    file,
    paragraph,
    (s) => `**${children.map((child) => s.stringify(child)).join("")}**`,
  );
}

function stripCommit(node: Nodes | undefined, file: File) {
  if (node?.type !== "paragraph") return;

  const first = node.children[0];
  if (first?.type !== "text" || !COMMIT_PREFIX.test(first.value)) return;

  first.value = first.value.replace(COMMIT_PREFIX, "");
  replaceSource(file, node, (s) => s.inner(node).replace(COMMIT_PREFIX, ""));
}
