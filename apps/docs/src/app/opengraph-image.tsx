import { ImageResponse } from "next/og";

import { loadFonts } from "@/lib/fonts";
import { OG_IMAGE_SIZE, OgImage } from "@/lib/og";
import { siteConfig } from "@/lib/site";

export const alt = siteConfig.name;
export const size = OG_IMAGE_SIZE;
export const contentType = "image/png";

export default async function Image() {
  const fonts = await loadFonts();

  return new ImageResponse(
    <OgImage title={siteConfig.tagline} description={siteConfig.description} />,
    {
      ...size,
      fonts,
    },
  );
}
