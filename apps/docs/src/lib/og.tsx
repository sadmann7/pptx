import { LogoMark } from "@/components/icons";
import { SITE_NAME, SITE_URL } from "@/lib/site";

export const OG_IMAGE_SIZE = {
  width: 1200,
  height: 630,
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
      tw="flex h-full w-full flex-col justify-between overflow-hidden px-[72px] pt-16 pb-[72px] text-[#fafafa]"
      style={{
        fontFamily: "Geist",
        backgroundColor: "#0a0a0a",
        backgroundImage:
          "radial-gradient(110% 120% at 70% 95%, rgba(26, 115, 232, 0.12), rgba(10, 10, 10, 0))",
      }}
    >
      <div tw="absolute top-[150px] left-[688px] flex h-[420px] w-[512px]">
        <SlideCard x={128} y={0} fill="#121212" />
        <SlideCard x={64} y={48} fill="#151515" />
        <SlideCard x={0} y={96} fill="#171717" front />
      </div>
      <div tw="flex w-full items-center justify-between">
        <div tw="flex items-center">
          <div tw="flex h-9 w-9 items-center justify-center rounded-lg bg-[#e5e5e5]">
            <LogoMark width={26} height={26} color="#171717" />
          </div>
          <p
            tw="m-0 ml-3 text-[28px] font-semibold"
            style={{ letterSpacing: -0.6 }}
          >
            {SITE_NAME}
          </p>
        </div>
        <p
          tw="m-0 text-[20px] text-[#a1a1a1]"
          style={{ fontFamily: "Geist Mono" }}
        >
          {new URL(SITE_URL).host}
        </p>
      </div>
      <div tw="flex w-[600px] flex-col">
        <p
          tw="m-0 text-[20px] text-[#a1a1a1]"
          style={{ fontFamily: "Geist Mono", letterSpacing: 0.4 }}
        >
          Docs
        </p>
        <p
          tw="m-0 mt-5 text-[64px] font-semibold"
          style={{ letterSpacing: -2, lineHeight: 1.08 }}
        >
          {title}
        </p>
        {description ? (
          <p
            tw="m-0 mt-5 text-[28px] text-[#a1a1a1]"
            style={{ lineHeight: 1.35 }}
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
      tw="absolute flex h-[270px] w-[480px] rounded-xl border border-[rgba(255,255,255,0.2)]"
      style={{
        left: x,
        top: y,
        backgroundColor: fill,
        ...(front && { boxShadow: "0 24px 64px rgba(0, 0, 0, 0.6)" }),
      }}
    >
      {front ? (
        <>
          <div tw="absolute top-12 left-10 h-[18px] w-[220px] rounded-[3px] bg-[rgba(250,250,250,0.85)]" />
          {SLIDE_LINES.map((line) => (
            <div
              key={`${line.x}-${line.y}`}
              tw="absolute h-2.5 rounded-[3px] bg-[#262626]"
              style={{ left: line.x, top: line.y, width: line.width }}
            />
          ))}
          <div
            tw="absolute top-12 left-[280px] h-[174px] w-40 rounded-lg"
            style={{
              backgroundImage:
                "linear-gradient(200deg, rgba(26, 115, 232, 0.4), rgba(26, 115, 232, 0.06))",
            }}
          />
          <div tw="absolute top-10 left-8 h-[34px] w-[236px] border-[1.5px] border-[#1a73e8]" />
          {SELECTION_HANDLES.map((handle) => (
            <div
              key={`${handle.x}-${handle.y}`}
              tw="absolute h-2 w-2 rounded-[2px] border-[1.5px] border-[#1a73e8] bg-white"
              style={{ left: handle.x, top: handle.y }}
            />
          ))}
        </>
      ) : null}
    </div>
  );
}
