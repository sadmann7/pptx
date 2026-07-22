import type { CSSProperties, ReactNode } from "react";

import { geistSans } from "@/fonts";
import { theme } from "@/theme";

const C = {
  panel: "rgba(10, 10, 10, 0.85)",
  panel2: "#181b15",
  line: "rgba(255,255,255,0.08)",
  accent: "#3b82f6",
  muted: "#71717a",
};

const font: CSSProperties = {
  fontFamily: geistSans,
};

function ToolbarIcon({ wide = false }: { wide?: boolean }) {
  return (
    <div
      style={{
        width: wide ? 64 : 26,
        height: 26,
        borderRadius: 6,
        border: `1px solid ${C.line}`,
        background: C.panel2,
      }}
    />
  );
}

export function EditorWindow({
  children,
  filename = "presentation.pptx",
  cursor,
  width = 1320,
  height = 800,
}: {
  children: ReactNode;
  filename?: string;
  cursor?: { x: number; y: number };
  width?: number;
  height?: number;
}) {
  const headerH = 48;
  const bodyH = height - headerH;

  return (
    <div
      style={{
        ...font,
        width,
        height,
        border: `1px solid rgba(255,255,255,0.06)`,
        borderRadius: 16,
        overflow: "hidden",
        background: theme.background,
        boxShadow: "0 80px 160px rgba(0,0,0,.55)",
        color: theme.text,
      }}
    >
      {/* Title bar */}
      <div
        style={{
          height: headerH,
          padding: "0 18px",
          display: "flex",
          alignItems: "center",
          gap: 8,
          background: "rgba(0,0,0,0.4)",
          borderBottom: `1px solid ${C.line}`,
        }}
      >
        {["#ff665f", "#f4bf4f", "#62c554"].map((color) => (
          <i key={color} style={{ width: 11, height: 11, borderRadius: 99, background: color }} />
        ))}
        <div style={{ marginLeft: 16, fontSize: 13, color: C.muted }}>{filename}</div>
        <div style={{ marginLeft: "auto", display: "flex", gap: 8 }}>
          <ToolbarIcon />
          <ToolbarIcon />
          <ToolbarIcon wide />
        </div>
      </div>

      {/* Main canvas area */}
      <main
        style={{
          position: "relative",
          width: "100%",
          height: bodyH,
          display: "grid",
          placeItems: "center",
          background: "radial-gradient(circle at 50% 42%, #1a1a2e, #0f0f1a 58%, #09090b)",
        }}
      >
        {children}

        {/* Zoom indicator */}
        <div
          style={{
            position: "absolute",
            right: 18,
            bottom: 14,
            display: "flex",
            alignItems: "center",
            gap: 8,
            color: C.muted,
            fontSize: 12,
          }}
        >
          <div style={{ width: 80, height: 3, background: "rgba(255,255,255,0.08)" }}>
            <div style={{ width: 55, height: 3, background: C.accent }} />
          </div>
          100%
        </div>

        {/* Cursor */}
        {cursor && (
          <svg
            viewBox="0 0 32 40"
            style={{
              position: "absolute",
              left: cursor.x,
              top: cursor.y,
              width: 28,
              filter: "drop-shadow(0 4px 5px rgba(0,0,0,.4))",
            }}
          >
            <path
              d="M3 2 28 24 16 26 10 38Z"
              fill={theme.text}
              stroke={theme.background}
              strokeWidth="2"
            />
          </svg>
        )}
      </main>
    </div>
  );
}
