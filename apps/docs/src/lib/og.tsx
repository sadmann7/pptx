import { LogoMark } from "@/components/icons";
import { SITE_NAME, SITE_URL } from "@/lib/site";

// Rendered at 2x the 1200×630 design so previews stay sharp on high-density screens.
const SCALE = 2;

function px(value: number) {
  return value * SCALE;
}

export const OG_IMAGE_SIZE = {
  width: px(1200),
  height: px(630),
};

const SLIDE_LINES = [
  { x: 40, y: 84, width: 180 },
  { x: 40, y: 104, width: 150 },
  { x: 40, y: 148, width: 120 },
  { x: 40, y: 168, width: 160 },
  { x: 40, y: 188, width: 100 },
];

const SELECTION_HANDLES = [
  { x: 28, y: 36 },
  { x: 264, y: 36 },
  { x: 28, y: 70 },
  { x: 264, y: 70 },
];

interface OgImageProps {
  title: string;
  description?: string;
}

export function OgImage({ title, description }: OgImageProps) {
  return (
    <div
      tw="flex h-full w-full flex-col justify-between overflow-hidden text-[#fafafa]"
      style={{
        padding: `${px(64)}px ${px(72)}px ${px(72)}px`,
        fontFamily: "Geist",
        backgroundColor: "#0a0a0a",
        backgroundImage:
          "radial-gradient(110% 120% at 70% 95%, rgba(26, 115, 232, 0.12), rgba(10, 10, 10, 0))",
      }}
    >
      <div
        tw="absolute flex"
        style={{
          top: px(150),
          left: px(688),
          width: px(512),
          height: px(420),
        }}
      >
        <SlideCard x={128} y={0} fill="#121212" />
        <SlideCard x={64} y={48} fill="#151515" />
        <SlideCard x={0} y={96} fill="#171717" front />
      </div>
      <div tw="flex w-full items-center justify-between">
        <div tw="flex items-center">
          <div
            tw="flex items-center justify-center bg-[#e5e5e5]"
            style={{ width: px(36), height: px(36), borderRadius: px(8) }}
          >
            <LogoMark width={px(26)} height={px(26)} color="#171717" />
          </div>
          <p
            tw="m-0 font-semibold"
            style={{
              marginLeft: px(12),
              fontSize: px(28),
              letterSpacing: px(-0.6),
            }}
          >
            {SITE_NAME}
          </p>
        </div>
        <p
          tw="m-0 text-[#a1a1a1]"
          style={{ fontFamily: "Geist Mono", fontSize: px(20) }}
        >
          {new URL(SITE_URL).host}
        </p>
      </div>
      <div tw="flex flex-col" style={{ width: px(600) }}>
        <p
          tw="m-0 text-[#a1a1a1]"
          style={{
            fontFamily: "Geist Mono",
            fontSize: px(20),
            letterSpacing: px(0.4),
          }}
        >
          Docs
        </p>
        <p
          tw="m-0 font-semibold"
          style={{
            marginTop: px(20),
            fontSize: px(64),
            letterSpacing: px(-2),
            lineHeight: 1.08,
          }}
        >
          {title}
        </p>
        {description ? (
          <p
            tw="m-0 text-[#a1a1a1]"
            style={{ marginTop: px(20), fontSize: px(28), lineHeight: 1.35 }}
          >
            {description}
          </p>
        ) : null}
      </div>
    </div>
  );
}

interface SlideCardProps {
  x: number;
  y: number;
  fill: string;
  front?: boolean;
}

function SlideCard({ x, y, fill, front }: SlideCardProps) {
  return (
    <div
      tw="absolute flex border-[rgba(255,255,255,0.2)]"
      style={{
        left: px(x),
        top: px(y),
        width: px(480),
        height: px(270),
        borderWidth: px(1),
        borderRadius: px(12),
        backgroundColor: fill,
        ...(front && {
          boxShadow: `0 ${px(24)}px ${px(64)}px rgba(0, 0, 0, 0.6)`,
        }),
      }}
    >
      {front ? (
        <>
          <div
            tw="absolute bg-[rgba(250,250,250,0.85)]"
            style={{
              left: px(40),
              top: px(48),
              width: px(220),
              height: px(18),
              borderRadius: px(3),
            }}
          />
          {SLIDE_LINES.map((line) => (
            <div
              key={`${line.x}-${line.y}`}
              tw="absolute bg-[#262626]"
              style={{
                left: px(line.x),
                top: px(line.y),
                width: px(line.width),
                height: px(10),
                borderRadius: px(3),
              }}
            />
          ))}
          <div
            tw="absolute"
            style={{
              left: px(280),
              top: px(48),
              width: px(160),
              height: px(174),
              borderRadius: px(8),
              backgroundImage:
                "linear-gradient(200deg, rgba(26, 115, 232, 0.4), rgba(26, 115, 232, 0.06))",
            }}
          />
          <div
            tw="absolute border-[#1a73e8]"
            style={{
              left: px(32),
              top: px(40),
              width: px(236),
              height: px(34),
              borderWidth: px(1.5),
            }}
          />
          {SELECTION_HANDLES.map((handle) => (
            <div
              key={`${handle.x}-${handle.y}`}
              tw="absolute border-[#1a73e8] bg-white"
              style={{
                left: px(handle.x),
                top: px(handle.y),
                width: px(8),
                height: px(8),
                borderWidth: px(1.5),
                borderRadius: px(2),
              }}
            />
          ))}
        </>
      ) : null}
    </div>
  );
}
