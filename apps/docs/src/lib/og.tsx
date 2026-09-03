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
      tw="flex h-full w-full flex-col p-16 text-white"
      style={{
        fontFamily: "Geist",
        background: "linear-gradient(to bottom right, #111827, #000000)",
      }}
    >
      <p tw="m-0 text-[82px] font-semibold">{title}</p>
      <p tw="m-0 mt-4 text-[44px] text-[rgba(240,240,240,0.8)]">{description}</p>
      <p tw="mt-auto mb-0 self-end text-[36px] text-[rgba(240,240,240,0.8)]">pptx</p>
    </div>
  );
}
