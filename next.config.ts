import type { NextConfig } from "next";

const nextConfig: NextConfig = {
  // VBA POST_LICENSE_URL sondaki slash ile istek atıyor: /api/license/
  trailingSlash: true,
};

export default nextConfig;
