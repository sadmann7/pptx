import { createMDX } from "fumadocs-mdx/next";

const withMDX = createMDX();

/** @type {import('next').NextConfig} */
const config = {
  reactStrictMode: true,
  // Already doing typechecking as separate tasks in CI
  typescript: { ignoreBuildErrors: true },
  // `PROFILE=1 pnpm build` keeps source maps, so perf traces name functions.
  productionBrowserSourceMaps: process.env.PROFILE === "1",
};

export default withMDX(config);
