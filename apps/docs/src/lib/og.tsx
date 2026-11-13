import { LogoMark } from "@/components/icons";
import { SITE_NAME, SITE_URL } from "@/lib/site";

export const OG_IMAGE_SIZE = {
  width: 1200,
  height: 630,
};

interface OgImageProps {
  title: string;
  description?: string;
}

export function OgImage({ title, description }: OgImageProps) {
  return (
    <div
      tw="flex h-full w-full flex-col px-[72px] pt-16 pb-[72px] text-[#fafafa]"
      style={{
        fontFamily: "Geist",
        backgroundColor: "#0a0a0a",
        backgroundImage:
          "radial-gradient(110% 120% at 70% 95%, rgba(26, 115, 232, 0.12), rgba(10, 10, 10, 0))",
      }}
    >
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
      <p
        tw="m-0 mt-auto text-[60px] font-semibold"
        style={{ letterSpacing: -2, lineHeight: 1.08 }}
      >
        {title}
      </p>
      {description ? (
        <p tw="m-0 mt-6 text-[32px] text-[#a1a1a1]" style={{ lineHeight: 1.3 }}>
          {description}
        </p>
      ) : null}
    </div>
  );
}
