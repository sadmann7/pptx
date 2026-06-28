import type {
  ColorMap,
  Paragraph,
  ParagraphContent,
  RunStyle,
  ThemeColors,
  ThemeFonts,
} from "@pptx/parser";
import { toCSS } from "../../render/color";
import { paragraphStyle, runStyle } from "../../render/text";

interface ParagraphElementProps {
  paragraph: Paragraph;
  theme: ThemeColors;
  colorMap?: ColorMap;
  themeFonts?: ThemeFonts;
  fontScale?: number;
  lnSpcReduction?: number;
}

export function ParagraphElement({
  paragraph,
  theme,
  colorMap,
  themeFonts,
  fontScale,
  lnSpcReduction,
}: ParagraphElementProps) {
  const style = paragraphStyle(paragraph.style, theme, colorMap, lnSpcReduction);
  const defaultRunStyle = paragraph.style.defaultRunStyle;

  const effectiveFontSize = defaultRunStyle?.fontSize ?? 12;
  const scale = fontScale ?? 1;
  const paraFontSize = `${effectiveFontSize * scale}pt`;

  const hasBullet =
    paragraph.style.bullet &&
    paragraph.style.bullet.type !== "none" &&
    paragraph.runs.some((r) => r.type === "run" && r.text !== "");

  if (!hasBullet) {
    return (
      <p style={{ ...style, fontSize: paraFontSize }}>
        {paragraph.runs.map((run, i) => (
          <RunElement
            key={i}
            run={run}
            theme={theme}
            defaultRunStyle={defaultRunStyle}
            colorMap={colorMap}
            themeFonts={themeFonts}
            fontScale={fontScale}
          />
        ))}
      </p>
    );
  }

  const bullet = paragraph.style.bullet!;
  const bulletChar =
    bullet.type === "char" ? bullet.char : bullet.type === "auto" ? (bullet.char ?? "•") : "•";
  const bulletColor =
    bullet.type !== "none" && "color" in bullet && bullet.color
      ? toCSS(bullet.color, theme, colorMap)
      : undefined;

  return (
    <p style={{ ...style, fontSize: paraFontSize, display: "flex", gap: "4pt" }}>
      <span style={{ flexShrink: 0, color: bulletColor }}>{bulletChar}</span>
      <span style={{ flex: 1 }}>
        {paragraph.runs.map((run, i) => (
          <RunElement
            key={i}
            run={run}
            theme={theme}
            defaultRunStyle={defaultRunStyle}
            colorMap={colorMap}
            themeFonts={themeFonts}
            fontScale={fontScale}
          />
        ))}
      </span>
    </p>
  );
}

function RunElement({
  run,
  theme,
  defaultRunStyle,
  colorMap,
  themeFonts,
  fontScale,
}: {
  run: ParagraphContent;
  theme: ThemeColors;
  defaultRunStyle?: RunStyle;
  colorMap?: ColorMap;
  themeFonts?: ThemeFonts;
  fontScale?: number;
}) {
  if (run.type === "lineBreak") return <br />;

  const style = runStyle(run.style, theme, defaultRunStyle, colorMap, themeFonts, fontScale);

  if (run.type === "field") {
    return <span style={style}>{run.text}</span>;
  }

  if (!run.text) return null;

  if (run.style.link) {
    return (
      <a href={run.style.link} style={style} target="_blank" rel="noopener noreferrer">
        {run.text}
      </a>
    );
  }

  return <span style={style}>{run.text}</span>;
}
