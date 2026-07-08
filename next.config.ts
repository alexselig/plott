import type { NextConfig } from "next";

// Static export for VibeHub hosting is opt-in via env so local dev keeps the
// API route (AI) and normal server rendering.
const isExport = process.env.VIBEHUB_EXPORT === "1";
const basePath = process.env.VIBEHUB_BASE_PATH || "";

const nextConfig: NextConfig = isExport
  ? {
      output: "export",
      trailingSlash: true,
      basePath: basePath || undefined,
      images: { unoptimized: true },
      env: { NEXT_PUBLIC_BASE_PATH: basePath },
    }
  : {};

export default nextConfig;
