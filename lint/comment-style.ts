interface Comment {
  type: "Line" | "Block" | "Shebang";
  value: string;
  range: [number, number];
}

interface Context {
  sourceCode: {
    text: string;
    getAllComments: () => Comment[];
  };
  report: (diagnostic: {
    node: Pick<Comment, "type" | "range">;
    message: string;
    fix?: () => { range: [number, number]; text: string };
  }) => void;
}

const DIRECTIVE =
  /^\s*(eslint|oxlint|@ts-|biome-|prettier-|istanbul|c8 |#region|#endregion|\/ <reference|@vitest-)/;
const IMPERATIVE =
  /^(Add|Allow|Append|Apply|Assign|Attach|Bail|Build|Cache|Calculate|Call|Cancel|Capture|Check|Clamp|Clear|Close|Collect|Compare|Compensate|Compute|Convert|Copy|Count|Create|Decode|Delete|Derive|Detect|Determine|Disable|Dispatch|Draw|Emit|Enable|Encode|Ensure|Exclude|Expand|Extract|Fall|Fetch|Filter|Find|Fix|Flush|Force|Format|Generate|Get|Handle|Hide|Ignore|Include|Inherit|Initialize|Insert|Invalidate|Iterate|Keep|Let|Load|Loop|Make|Mark|Measure|Merge|Mirror|Move|Normalize|Notify|Open|Parse|Pick|Prefer|Prevent|Process|Push|Read|Register|Remove|Render|Replace|Reset|Resolve|Restore|Return|Reuse|Rotate|Run|Save|Scale|Schedule|Scroll|Select|Set|Show|Simulate|Skip|Sort|Split|Start|Stop|Store|Strip|Subscribe|Swap|Sync|Toggle|Track|Treat|Trigger|Trim|Try|Unsubscribe|Update|Use|Validate|Wait|Walk|Wrap|Write) [a-z`]/;
const FILLER =
  /^(Note|NOTE|Important|IMPORTANT|Basically|Simply|Just|Here we|Now we|This (function|helper|component|hook|method|file)|Helper (to|for|function))\b/;
const FIRST_PERSON = /\b(we|we're|we'll|our|us|let's|Let's|I|I'm)\b/;
const BANNER = /^\s*([-=*#~_/])\1{3,}|([-=*#~_])\2{3,}\s*$/;
const END_PUNCTUATION = /([.?!:]|\.\.\.|https?:\/\/\S+)$/;
const CODE =
  /^(const|let|var|return|await|import|export|if|for|while)\b.*[;{]$|^[\w$.]+\([^)]*\);$|^[\w$.]+ = .+;$|^[{}]\)?;?$/;

export const commentStyle = {
  meta: {
    type: "suggestion",
    docs: {
      description:
        "Keeps `//` comments to short reasons: no narrating what the next line does, filler, first person, banners or commented-out code, and every comment ends with a full stop",
    },
    fixable: "code",
  },
  create(context: Context) {
    return {
      "Program:exit"() {
        for (const block of getLineCommentBlocks(context)) {
          checkBlock(context, block);
        }
      },
    };
  },
};

function getLineCommentBlocks(context: Context) {
  const { text } = context.sourceCode;
  const blocks: Comment[][] = [];
  let previous: Comment | undefined;
  for (const comment of context.sourceCode.getAllComments()) {
    if (comment.type !== "Line") {
      previous = undefined;
      continue;
    }
    const between = previous && text.slice(previous.range[1], comment.range[0]);
    const block = blocks.at(-1);
    // Lines right under each other read as one comment.
    if (block && between !== undefined && /^[ \t]*\n[ \t]*$/.test(between)) {
      block.push(comment);
    } else {
      blocks.push([comment]);
    }
    previous = comment;
  }
  return blocks.filter((block) => !DIRECTIVE.test(block[0]?.value ?? ""));
}

function checkBlock(context: Context, block: Comment[]) {
  const first = block[0];
  const last = block.at(-1);
  if (!first || !last) return;
  const lines = block.map((comment) => comment.value.trim());
  const opening = lines[0] ?? "";
  const report = (message: string, fix?: () => { range: [number, number]; text: string }) =>
    context.report({
      node: { type: "Line", range: [first.range[0], last.range[1]] },
      message,
      fix,
    });

  if (lines.some((line) => BANNER.test(line))) {
    report("Drop banner comments; let the code's structure show the sections.");
    return;
  }
  if (lines.every((line) => CODE.test(line))) {
    report("Delete commented-out code; git keeps the history.");
    return;
  }
  const filler = FILLER.exec(opening)?.[0];
  if (filler) {
    report(`Drop the filler \`${filler}\` and say the reason directly.`);
  }
  if (IMPERATIVE.test(opening)) {
    report(
      "Say why, not what the next line does: `Keeps …` rather than `Keep …`, or drop it if the code says it already.",
    );
  }
  if (lines.some((line) => FIRST_PERSON.test(line))) {
    report("Describe the code instead of `we`/`our`.");
  }
  const closing = lines.at(-1) ?? "";
  if (!isTrailing(context, first) && closing && !END_PUNCTUATION.test(closing)) {
    const end = last.range[0] + 2 + last.value.trimEnd().length;
    report("End the comment with a full stop.", () => ({ range: [end, end], text: "." }));
  }
}

/** A comment after code on the same line, which reads as a label rather than a sentence. */
function isTrailing(context: Context, comment: Comment) {
  const { text } = context.sourceCode;
  const lineStart = text.lastIndexOf("\n", comment.range[0] - 1) + 1;
  return text.slice(lineStart, comment.range[0]).trim() !== "";
}
