import { Sequence, interpolate, useCurrentFrame } from "remotion";

export interface GlassCodeBlockProps {
  code?: string;
  title?: string;
  width?: number;
  height?: number;
  fontSize?: number;
  glassColor?: string;
  staggerFrames?: number;
  showTrafficLights?: boolean;
  aura?: boolean;
  className?: string;
}

const FONT_MONO = "var(--font-geist-mono), ui-monospace, SFMono-Regular, monospace";

const DEFAULT_CODE = `import { Presentation } from "@diceui/pptx";

export function Viewer({ file }) {
  return (
    <Presentation.Root file={file}>
      <Presentation.Viewport autoFit>
        <Presentation.Slide />
      </Presentation.Viewport>
    </Presentation.Root>
  );
}`;

const KEYWORDS = new Set([
  "import",
  "from",
  "export",
  "function",
  "const",
  "let",
  "var",
  "return",
  "if",
  "else",
  "for",
  "while",
  "new",
  "class",
  "extends",
  "default",
  "true",
  "false",
  "null",
  "undefined",
]);

type Token = {
  text: string;
  kind: "code" | "comment" | "string" | "keyword" | "number" | "tag" | "attribute";
};

function tokenizeLine(line: string): Token[] {
  const trimmed = line.trimStart();
  if (trimmed.startsWith("//")) {
    return [{ text: line, kind: "comment" }];
  }

  const tokens: Token[] = [];
  const re =
    /("[^"]*"|'[^']*'|`[^`]*`|<\/?[A-Z][\w.]*|\/?>|\b\d+\b|\b[A-Za-z_$][\w$]*\b|[^\w"'<>]+|[<>])/g;
  let match: RegExpExecArray | null;
  let inTag = false;
  while ((match = re.exec(line)) !== null) {
    const t = match[0];
    const first = t[0];
    if (first === '"' || first === "'" || first === "`") {
      tokens.push({ text: t, kind: inTag ? "string" : "string" });
    } else if (t.startsWith("<") && /^<\/?[A-Z]/.test(t)) {
      inTag = true;
      tokens.push({ text: t, kind: "tag" });
    } else if (t === "/>" || t === ">") {
      tokens.push({ text: t, kind: "tag" });
      if (t === "/>") inTag = false;
      if (t === ">") inTag = false;
    } else if (/^\d+$/.test(t)) {
      tokens.push({ text: t, kind: "number" });
    } else if (/^[A-Za-z_$][\w$]*$/.test(t) && KEYWORDS.has(t)) {
      tokens.push({ text: t, kind: "keyword" });
    } else if (inTag && /^[a-z][A-Za-z]*$/.test(t)) {
      tokens.push({ text: t, kind: "attribute" });
    } else {
      tokens.push({ text: t, kind: "code" });
    }
  }
  return tokens;
}

const TOKEN_COLORS: Record<Token["kind"], string> = {
  code: "#e4e4e7",
  comment: "#6b7280",
  string: "#4ade80",
  keyword: "#a78bfa",
  number: "#fbbf24",
  tag: "#7dd3fc",
  attribute: "#fca5a5",
};

export function GlassCodeBlock({
  code = DEFAULT_CODE,
  title = "viewer.tsx",
  width = 760,
  height = 460,
  fontSize = 16,
  glassColor = "rgba(10, 10, 10, 0.6)",
  staggerFrames = 4,
  showTrafficLights = true,
  aura = false,
  className,
}: GlassCodeBlockProps) {
  const lines = code.split("\n");

  return (
    <div
      className={className}
      style={{
        position: "relative",
        width,
        height,
        display: "flex",
        alignItems: "stretch",
        justifyContent: "center",
      }}
    >
      {aura && <BackdropAura />}

      <div
        style={{
          position: "relative",
          width: "100%",
          borderRadius: 16,
          padding: 1,
          background:
            "linear-gradient(180deg, rgba(255,255,255,0.12) 0%, rgba(255,255,255,0.04) 100%)",
        }}
      >
        <div
          style={{
            display: "flex",
            flexDirection: "column",
            width: "100%",
            height: "100%",
            borderRadius: 15,
            background: glassColor,
            backdropFilter: "blur(40px)",
            overflow: "hidden",
          }}
        >
          {/* Chrome */}
          <div
            style={{
              display: "flex",
              alignItems: "center",
              gap: 8,
              padding: "12px 16px",
              borderBottom: "1px solid rgba(255,255,255,0.06)",
            }}
          >
            {showTrafficLights && (
              <div style={{ display: "flex", gap: 6 }}>
                <Light color="#ff5f57" />
                <Light color="#febc2e" />
                <Light color="#28c840" />
              </div>
            )}
            <span
              style={{
                fontSize: fontSize * 0.75,
                color: "#71717a",
                fontFamily: FONT_MONO,
                marginLeft: showTrafficLights ? 6 : 0,
              }}
            >
              {title}
            </span>
          </div>

          {/* Code body */}
          <div
            style={{
              flex: 1,
              overflow: "hidden",
              padding: "16px 20px",
            }}
          >
            {lines.map((line, i) => (
              <Sequence key={i} from={i * staggerFrames} layout="none">
                <CodeLine line={line} index={i} fontSize={fontSize} />
              </Sequence>
            ))}
          </div>
        </div>
      </div>
    </div>
  );
}

function Light({ color }: { color: string }) {
  return (
    <div
      style={{
        width: 11,
        height: 11,
        borderRadius: "50%",
        backgroundColor: color,
      }}
    />
  );
}

function CodeLine({ line, index, fontSize }: { line: string; index: number; fontSize: number }) {
  const frame = useCurrentFrame();
  const opacity = interpolate(frame, [0, 8], [0, 1], {
    extrapolateLeft: "clamp",
    extrapolateRight: "clamp",
  });
  const ty = interpolate(frame, [0, 8], [4, 0], {
    extrapolateLeft: "clamp",
    extrapolateRight: "clamp",
  });

  const tokens = tokenizeLine(line);
  if (tokens.length === 0) {
    return <div style={{ height: fontSize * 0.8 }} />;
  }
  return (
    <div
      style={{
        display: "flex",
        gap: 0,
        opacity,
        transform: `translateY(${ty}px)`,
        fontFamily: FONT_MONO,
        fontSize,
        lineHeight: 1.7,
        whiteSpace: "pre",
      }}
    >
      <span
        style={{
          color: "#3f3f46",
          width: "2.5ch",
          textAlign: "right",
          marginRight: "1.5ch",
          userSelect: "none",
        }}
      >
        {String(index + 1).padStart(2, " ")}
      </span>
      <span>
        {tokens.map((t, i) => (
          <span key={i} style={{ color: TOKEN_COLORS[t.kind] }}>
            {t.text}
          </span>
        ))}
      </span>
    </div>
  );
}

function BackdropAura() {
  const frame = useCurrentFrame();
  const t = frame / 60;
  const x = 50 + Math.sin(t) * 20;
  const y = 50 + Math.cos(t * 0.7) * 15;
  return (
    <div
      style={{
        position: "absolute",
        inset: -60,
        borderRadius: "50%",
        background: `radial-gradient(circle at ${x}% ${y}%, rgba(99,102,241,0.25) 0%, transparent 60%)`,
        filter: "blur(60px)",
        zIndex: -1,
      }}
    />
  );
}
