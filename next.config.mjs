/** @type {import('next').NextConfig} */
const nextConfig = {
  serverExternalPackages: ['pdf-parse', '@napi-rs/canvas'],
  experimental: {
    serverActions: {
      bodySizeLimit: '50mb', // Allow larger uploads
    },
  },
};

export default nextConfig;
