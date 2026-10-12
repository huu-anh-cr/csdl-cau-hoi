/** @type {import('next').NextConfig} */
const nextConfig = {
  async headers() {
    return [
      {
        source: "/csdl_mau.xlsx",
        headers: [
          {
            key: "Content-Disposition",
            value: 'attachment; filename="csdl_mau.xlsx"',
          },
        ],
      },
    ];
  },
};

export default nextConfig;
