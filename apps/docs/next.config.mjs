import { createMDX } from "fumadocs-mdx/next";

const withMDX = createMDX();

/** @type {import('next').NextConfig} */
const config = {
  reactStrictMode: true,
  // Already doing typechecking as separate tasks in CI
  typescript: { ignoreBuildErrors: true },
};

export default withMDX(config);
