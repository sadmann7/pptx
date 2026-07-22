import type { CSSProperties, ReactNode } from "react";

import { AbsoluteFill, Easing, interpolate, Sequence, useCurrentFrame } from "remotion";

import { SoftBlurIn } from "@/components/remocn/soft-blur-in";
import { geistSans } from "@/fonts";

const C = {
  ink: "#070806",
  panel: "#11130f",
  panel2: "#181b15",
  line: "#30342a",
  acid: "#d9ff43",
  white: "#f5f6ef",
  muted: "#9ba18f",
};

const clamp = {
  extrapolateLeft: "clamp",
  extrapolateRight: "clamp",
} as const;
const ease = Easing.bezier(0.22, 1, 0.36, 1);

const progress = (frame: number, from: number, to: number) =>
  interpolate(frame, [from, to], [0, 1], { ...clamp, easing: ease });

const fadeWindow = (frame: number, duration: number, edge = 16) =>
  interpolate(frame, [0, edge, duration - edge, duration], [0, 1, 1, 0], clamp);

const font: CSSProperties = {
  fontFamily: geistSans,
};

function Noise() {
  return (
    <AbsoluteFill
      style={{
        opacity: 0.045,
        backgroundImage:
          'url("data:image/svg+xml,%3Csvg viewBox=%270 0 180 180%27 xmlns=%27http://www.w3.org/2000/svg%27%3E%3Cfilter id=%27n%27%3E%3CfeTurbulence type=%27fractalNoise%27 baseFrequency=%271.1%27 numOctaves=%273%27 stitchTiles=%27stitch%27/%3E%3C/filter%3E%3Crect width=%27100%25%27 height=%27100%25%27 filter=%27url(%23n)%27 opacity=%27.8%27/%3E%3C/svg%3E")',
        pointerEvents: "none",
      }}
    />
  );
}

function Backdrop({ label }: { label?: string }) {
  return (
    <AbsoluteFill
      style={{
        ...font,
        color: C.white,
        overflow: "hidden",
        background: `radial-gradient(circle at 73% 26%, rgba(217,255,67,.13), transparent 29%), ${C.ink}`,
      }}
    >
      <AbsoluteFill
        style={{
          opacity: 0.16,
          backgroundImage:
            "linear-gradient(#43483a 1px, transparent 1px), linear-gradient(90deg, #43483a 1px, transparent 1px)",
          backgroundSize: "72px 72px",
          maskImage: "linear-gradient(to bottom, transparent, black 20%, black 80%, transparent)",
        }}
      />
      {label ? (
        <div
          style={{
            position: "absolute",
            left: 64,
            top: 48,
            color: C.acid,
            fontSize: 20,
            fontWeight: 700,
            letterSpacing: 3,
            textTransform: "uppercase",
          }}
        >
          {label}
        </div>
      ) : null}
      <Noise />
    </AbsoluteFill>
  );
}

function DiceMark({ dark = false }: { dark?: boolean }) {
  return (
    <div style={{ display: "flex", alignItems: "center", gap: 12 }}>
      <div
        style={{
          width: 28,
          height: 28,
          borderRadius: 7,
          background: dark ? C.ink : C.acid,
          display: "grid",
          gridTemplateColumns: "repeat(2, 5px)",
          placeContent: "center",
          gap: 4,
        }}
      >
        {[0, 1, 2, 3].map((dot) => (
          <i
            key={dot}
            style={{
              width: 5,
              height: 5,
              borderRadius: 99,
              background: dark ? C.acid : C.ink,
            }}
          />
        ))}
      </div>
      <strong style={{ fontSize: 22, letterSpacing: -0.6 }}>DiceUI</strong>
    </div>
  );
}

function MiniChart() {
  return (
    <svg viewBox="0 0 420 170" style={{ width: "100%", height: "100%" }}>
      {[35, 75, 115, 155].map((y) => (
        <line key={y} x1="8" y1={y} x2="412" y2={y} stroke="#d9ff43" opacity=".16" />
      ))}
      <path
        d="M 10 145 C 55 131, 76 139, 112 107 S 174 85, 209 93 S 264 66, 302 70 S 361 31, 410 24"
        fill="none"
        stroke={C.acid}
        strokeWidth="7"
        strokeLinecap="round"
      />
      {[10, 112, 209, 302, 410].map((x, index) => (
        <circle
          key={x}
          cx={x}
          cy={[145, 107, 93, 70, 24][index]}
          r="7"
          fill={C.ink}
          stroke={C.acid}
          strokeWidth="4"
        />
      ))}
    </svg>
  );
}

type ArtworkVariant = "cover" | "chart" | "editor" | "architecture";

function SlideArtwork({ variant }: { variant: ArtworkVariant }) {
  const shell: CSSProperties = {
    ...font,
    width: "100%",
    height: "100%",
    boxSizing: "border-box",
    position: "relative",
    overflow: "hidden",
    color: C.white,
    background: C.ink,
    border: `1px solid ${C.line}`,
  };

  if (variant === "cover") {
    return (
      <div style={{ ...shell, padding: "8% 8%" }}>
        <div style={{ display: "flex", justifyContent: "space-between" }}>
          <DiceMark />
          <span style={{ color: C.muted, fontSize: 18 }}>OPEN SOURCE · REACT</span>
        </div>
        <h1
          style={{
            margin: "12% 0 0",
            width: "80%",
            fontSize: 82,
            lineHeight: 0.92,
            letterSpacing: -5,
          }}
        >
          PRESENTATIONS,
          <br />
          <span style={{ color: C.acid }}>COMPOSED.</span>
        </h1>
        <p
          style={{
            fontSize: 25,
            color: C.muted,
            width: "61%",
            lineHeight: 1.35,
          }}
        >
          A headless, themeable PPTX viewer and editor built for React.
        </p>
        <div
          style={{
            position: "absolute",
            right: "-5%",
            bottom: "-20%",
            width: "35%",
            aspectRatio: "1",
            borderRadius: "50%",
            border: `30px solid ${C.acid}`,
            opacity: 0.85,
          }}
        />
      </div>
    );
  }

  if (variant === "chart") {
    return (
      <div
        style={{
          ...shell,
          background: C.acid,
          color: C.ink,
          padding: "7% 8%",
        }}
      >
        <div style={{ display: "flex", justifyContent: "space-between" }}>
          <DiceMark dark />
          <span style={{ fontSize: 18, fontWeight: 700 }}>NATIVE CHARTS</span>
        </div>
        <h2 style={{ fontSize: 60, letterSpacing: -3, margin: "8% 0 0" }}>
          Fidelity without the lock-in.
        </h2>
        <div style={{ height: "35%", marginTop: "2%" }}>
          <MiniChart />
        </div>
        <div style={{ display: "flex", gap: 36, fontSize: 17, fontWeight: 700 }}>
          <span>CATEGORIES</span>
          <span>SERIES</span>
          <span>LEGENDS</span>
          <span>AXES</span>
        </div>
      </div>
    );
  }

  if (variant === "editor") {
    return (
      <div style={{ ...shell, padding: "6% 7%" }}>
        <div style={{ display: "flex", justifyContent: "space-between" }}>
          <DiceMark />
          <span style={{ fontSize: 18, color: C.acid }}>EDIT → SAVE → REOPEN</span>
        </div>
        <h2 style={{ fontSize: 58, letterSpacing: -3, margin: "7% 0 5%" }}>
          Editing that stays composable.
        </h2>
        <div
          style={{
            display: "grid",
            gridTemplateColumns: "1fr 1fr",
            gap: 18,
          }}
        >
          {["Text runs", "Solid fills", "Transforms", "Slide order"].map((item, i) => (
            <div
              key={item}
              style={{
                padding: "20px 24px",
                border: `1px solid ${i === 1 ? C.acid : C.line}`,
                background: i === 1 ? C.acid : C.panel,
                color: i === 1 ? C.ink : C.white,
                fontSize: 20,
                fontWeight: 700,
              }}
            >
              0{i + 1} — {item}
            </div>
          ))}
        </div>
      </div>
    );
  }

  // architecture
  return (
    <div style={{ ...shell, padding: "6% 7%" }}>
      <div style={{ display: "flex", justifyContent: "space-between" }}>
        <DiceMark />
        <span style={{ fontSize: 18, color: C.muted }}>YOUR UI, YOUR RULES</span>
      </div>
      <h2 style={{ fontSize: 58, letterSpacing: -3, margin: "7% 0 4%" }}>
        Primitive by primitive.
      </h2>
      <div
        style={{
          padding: 28,
          border: `1px solid ${C.line}`,
          background: "#0c0e0b",
          color: "#c9cfba",
          fontFamily: "ui-monospace, SFMono-Regular, Menlo, monospace",
          fontSize: 19,
          lineHeight: 1.65,
        }}
      >
        <span style={{ color: C.acid }}>&lt;Presentation.Root&gt;</span>
        <br />
        &nbsp;&nbsp;&lt;Presentation.Toolbar /&gt;
        <br />
        &nbsp;&nbsp;&lt;Presentation.Sidebar /&gt;
        <br />
        &nbsp;&nbsp;&lt;Presentation.Canvas /&gt;
        <br />
        <span style={{ color: C.acid }}>&lt;/Presentation.Root&gt;</span>
      </div>
    </div>
  );
}

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

function EditorWindowChrome({
  active = 0,
  children,
  cursor,
}: {
  active?: number;
  children: ReactNode;
  cursor?: { x: number; y: number };
}) {
  const variants: ArtworkVariant[] = ["cover", "chart", "editor", "architecture"];
  return (
    <div
      style={{
        ...font,
        width: 1320,
        height: 800,
        border: "1px solid #3a3e33",
        borderRadius: 20,
        overflow: "hidden",
        background: C.panel,
        boxShadow: "0 80px 160px rgba(0,0,0,.55)",
        color: C.white,
      }}
    >
      <div
        style={{
          height: 54,
          padding: "0 20px",
          display: "flex",
          alignItems: "center",
          gap: 9,
          background: "#0c0e0b",
          borderBottom: `1px solid ${C.line}`,
        }}
      >
        {["#ff665f", "#f4bf4f", "#62c554"].map((color) => (
          <i
            key={color}
            style={{
              width: 11,
              height: 11,
              borderRadius: 99,
              background: color,
            }}
          />
        ))}
        <div style={{ marginLeft: 18, fontSize: 15, color: C.muted }}>launch-deck.pptx</div>
        <div style={{ marginLeft: "auto", display: "flex", gap: 10 }}>
          <ToolbarIcon />
          <ToolbarIcon />
          <ToolbarIcon wide />
        </div>
      </div>
      <div style={{ display: "flex", height: 746 }}>
        <aside
          style={{
            width: 210,
            padding: 18,
            boxSizing: "border-box",
            borderRight: `1px solid ${C.line}`,
            background: "#0d0f0c",
          }}
        >
          {variants.map((variant, index) => (
            <div
              key={variant}
              style={{
                position: "relative",
                width: 172,
                height: 97,
                marginBottom: 16,
                border: `2px solid ${index === active ? C.acid : "#2b2f26"}`,
                borderRadius: 5,
                overflow: "hidden",
                opacity: index === active ? 1 : 0.62,
              }}
            >
              <div
                style={{
                  width: 960,
                  height: 540,
                  transform: "scale(.179)",
                  transformOrigin: "top left",
                }}
              >
                <SlideArtwork variant={variant} />
              </div>
              <span
                style={{
                  position: "absolute",
                  left: 5,
                  bottom: 4,
                  padding: "2px 5px",
                  borderRadius: 4,
                  background: "rgba(0,0,0,.7)",
                  fontSize: 10,
                }}
              >
                {index + 1}
              </span>
            </div>
          ))}
        </aside>
        <main
          style={{
            flex: 1,
            position: "relative",
            display: "grid",
            placeItems: "center",
            background: "radial-gradient(circle at 50% 42%, #262a20, #171a15 58%, #10120f)",
          }}
        >
          <div
            style={{
              width: 960,
              height: 540,
              boxShadow: "0 28px 70px rgba(0,0,0,.48)",
            }}
          >
            {children}
          </div>
          <div
            style={{
              position: "absolute",
              right: 22,
              bottom: 18,
              display: "flex",
              alignItems: "center",
              gap: 10,
              color: C.muted,
              fontSize: 13,
            }}
          >
            <div style={{ width: 100, height: 3, background: C.line }}>
              <div style={{ width: 68, height: 3, background: C.acid }} />
            </div>
            100%
          </div>
          {cursor ? (
            <svg
              viewBox="0 0 32 40"
              style={{
                position: "absolute",
                left: cursor.x,
                top: cursor.y,
                width: 32,
                filter: "drop-shadow(0 4px 5px rgba(0,0,0,.4))",
              }}
            >
              <path d="M3 2 28 24 16 26 10 38Z" fill={C.white} stroke={C.ink} strokeWidth="2" />
            </svg>
          ) : null}
        </main>
      </div>
    </div>
  );
}

function FeatureTag({ children }: { children: ReactNode }) {
  return (
    <div
      style={{
        ...font,
        padding: "13px 18px",
        borderRadius: 999,
        border: `1px solid ${C.acid}`,
        background: "rgba(7,8,6,.86)",
        color: C.acid,
        fontWeight: 700,
        fontSize: 17,
        letterSpacing: 0.2,
        boxShadow: "0 14px 40px rgba(0,0,0,.35)",
      }}
    >
      {children}
    </div>
  );
}

function PerspectiveDriftScene({ duration = 180 }: { duration?: number }) {
  const frame = useCurrentFrame();
  const t = progress(frame, 0, duration - 1);
  const intro = progress(frame, 0, 28);
  const cursorT = progress(frame, 42, 108);
  return (
    <AbsoluteFill>
      <Backdrop label="01 / Perspective drift" />
      <div
        style={{
          position: "absolute",
          left: 300,
          top: 165,
          opacity: intro * fadeWindow(frame, duration),
          transformStyle: "preserve-3d",
          transform: `perspective(1700px) translate3d(${interpolate(t, [0, 1], [110, -105])}px, ${interpolate(t, [0, 1], [60, -42])}px, 0) rotateX(${interpolate(t, [0, 1], [4, -1.5])}deg) rotateY(${interpolate(t, [0, 1], [-12, 5])}deg) scale(${interpolate(intro, [0, 1], [0.88, 1.03])})`,
        }}
      >
        <EditorWindowChrome
          active={0}
          cursor={{
            x: interpolate(cursorT, [0, 1], [760, 490]),
            y: interpolate(cursorT, [0, 1], [580, 360]),
          }}
        >
          <SlideArtwork variant="cover" />
        </EditorWindowChrome>
      </div>
      <div
        style={{
          ...font,
          position: "absolute",
          left: 66,
          bottom: 54,
          color: C.muted,
          fontSize: 18,
          opacity: progress(frame, 20, 40),
        }}
      >
        CSS perspective · deterministic camera easing
      </div>
    </AbsoluteFill>
  );
}

function LayeredParallaxScene({ duration = 180 }: { duration?: number }) {
  const frame = useCurrentFrame();
  const t = progress(frame, 0, duration - 1);
  const opacity = fadeWindow(frame, duration);
  const cards: Array<{
    variant: ArtworkVariant;
    x: number;
    y: number;
    depth: number;
    rotate: number;
  }> = [
    { variant: "architecture", x: 120, y: 182, depth: 0.55, rotate: -7 },
    { variant: "chart", x: 640, y: 110, depth: 0.86, rotate: 4 },
    { variant: "editor", x: 1090, y: 330, depth: 1.15, rotate: -3 },
  ];
  return (
    <AbsoluteFill>
      <Backdrop label="02 / Layered parallax" />
      <div style={{ opacity }}>
        {cards.map((card, index) => (
          <div
            key={card.variant}
            style={{
              position: "absolute",
              left: card.x,
              top: card.y,
              width: 760,
              height: 428,
              transform: `translate3d(${interpolate(t, [0, 1], [120 * card.depth, -150 * card.depth])}px, ${interpolate(t, [0, 1], [70 * card.depth, -55 * card.depth])}px, 0) rotate(${card.rotate + interpolate(t, [0, 1], [-1.3, 1.3])}deg) scale(${0.88 + index * 0.03})`,
              boxShadow: "0 60px 100px rgba(0,0,0,.48)",
            }}
          >
            <SlideArtwork variant={card.variant} />
          </div>
        ))}
        <div
          style={{
            position: "absolute",
            left: 410 + interpolate(t, [0, 1], [90, -190]),
            top: 740,
          }}
        >
          <FeatureTag>headless primitives</FeatureTag>
        </div>
        <div
          style={{
            position: "absolute",
            left: 960 + interpolate(t, [0, 1], [170, -240]),
            top: 220,
          }}
        >
          <FeatureTag>native charts</FeatureTag>
        </div>
        <div
          style={{
            position: "absolute",
            left: 1310 + interpolate(t, [0, 1], [240, -320]),
            top: 820,
          }}
        >
          <FeatureTag>round-trip editing</FeatureTag>
        </div>
      </div>
    </AbsoluteFill>
  );
}

function FlyPanel({ variant, x, y }: { variant: ArtworkVariant; x: number; y: number }) {
  return (
    <div
      style={{
        position: "absolute",
        left: x,
        top: y,
        width: 960,
        height: 540,
        boxShadow: "0 80px 140px rgba(0,0,0,.5)",
      }}
    >
      <SlideArtwork variant={variant} />
    </div>
  );
}

function CanvasFlythroughScene({ duration = 240 }: { duration?: number }) {
  const frame = useCurrentFrame();
  const stops = [0, duration * 0.3, duration * 0.63, duration - 1];
  const cameraX = interpolate(frame, stops, [60, -940, -1400, -580], {
    ...clamp,
    easing: ease,
  });
  const cameraY = interpolate(frame, stops, [0, -110, -670, -1000], {
    ...clamp,
    easing: ease,
  });
  const zoom = interpolate(frame, stops, [1, 0.84, 0.72, 0.86], {
    ...clamp,
    easing: ease,
  });
  const tilt = interpolate(frame, stops, [0, -1.1, 1, -0.7], {
    ...clamp,
    easing: ease,
  });
  return (
    <AbsoluteFill>
      <Backdrop label="03 / Canvas fly-through" />
      <div
        style={{
          position: "absolute",
          left: 420,
          top: 270,
          width: 3300,
          height: 1900,
          transformOrigin: "0 0",
          transform: `perspective(1800px) translate3d(${cameraX}px, ${cameraY}px, 0) rotateX(${tilt * 0.5}deg) rotateZ(${tilt}deg) scale(${zoom})`,
        }}
      >
        <FlyPanel variant="cover" x={0} y={0} />
        <FlyPanel variant="chart" x={1280} y={180} />
        <FlyPanel variant="editor" x={2220} y={1030} />
        <FlyPanel variant="architecture" x={820} y={1210} />
        <svg width="3100" height="1800" style={{ position: "absolute", inset: 0, zIndex: -1 }}>
          <path
            d="M960 270 C1100 270 1120 370 1280 420"
            fill="none"
            stroke={C.acid}
            strokeWidth="3"
            strokeDasharray="10 12"
            opacity=".55"
          />
          <path
            d="M2240 720 C2380 820 2450 880 2600 1030"
            fill="none"
            stroke={C.acid}
            strokeWidth="3"
            strokeDasharray="10 12"
            opacity=".55"
          />
          <path
            d="M1850 700 C1710 900 1570 1080 1520 1210"
            fill="none"
            stroke={C.acid}
            strokeWidth="3"
            strokeDasharray="10 12"
            opacity=".55"
          />
        </svg>
      </div>
      <div
        style={{
          ...font,
          position: "absolute",
          right: 70,
          bottom: 54,
          fontSize: 18,
          color: C.muted,
        }}
      >
        One continuous camera · four product moments
      </div>
    </AbsoluteFill>
  );
}

function TitleCard({ outro = false, duration }: { outro?: boolean; duration: number }) {
  const frame = useCurrentFrame();
  const opacity = fadeWindow(frame, duration, 14);

  const taglineOpacity = progress(frame, 25, 45);
  const taglineY = interpolate(frame, [25, 45], [16, 0], {
    ...clamp,
    easing: Easing.bezier(0.16, 1, 0.3, 1),
  });

  return (
    <AbsoluteFill>
      <Backdrop />
      <AbsoluteFill
        style={{
          ...font,
          alignItems: "center",
          justifyContent: "center",
          flexDirection: "column",
          gap: 28,
          opacity,
        }}
      >
        <SoftBlurIn
          text={outro ? "The web can open .pptx" : "PowerPoint in the Browser"}
          fontSize={96}
          fontWeight={700}
          color={C.white}
        />

        <div
          style={{
            fontSize: 24,
            fontFamily: geistSans,
            fontWeight: 500,
            color: C.muted,
            letterSpacing: "0.04em",
            opacity: taglineOpacity,
            transform: `translateY(${taglineY}px)`,
          }}
        >
          {outro
            ? "Viewer · Editor · React primitives"
            : "@diceui/pptx — parse, render, and edit .pptx in React"}
        </div>
      </AbsoluteFill>
    </AbsoluteFill>
  );
}

export function LaunchFilm() {
  return (
    <AbsoluteFill style={{ background: C.ink }}>
      <Sequence from={0} durationInFrames={72} premountFor={20}>
        <TitleCard duration={72} />
      </Sequence>
      <Sequence from={64} durationInFrames={132} premountFor={20}>
        <PerspectiveDriftScene duration={132} />
      </Sequence>
      <Sequence from={188} durationInFrames={128} premountFor={20}>
        <LayeredParallaxScene duration={128} />
      </Sequence>
      <Sequence from={308} durationInFrames={116} premountFor={20}>
        <CanvasFlythroughScene duration={116} />
      </Sequence>
      <Sequence from={416} durationInFrames={64} premountFor={20}>
        <TitleCard outro duration={64} />
      </Sequence>
    </AbsoluteFill>
  );
}
