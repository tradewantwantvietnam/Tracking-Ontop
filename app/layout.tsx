import "./globals.css";

export const metadata = {
  title: "Theo dõi phân bổ chương trình",
  description:
    "Dashboard theo dõi suất phân bổ, tiền chiết khấu và NPP"
};

export default function RootLayout({
  children
}: {
  children: React.ReactNode;
}) {
  return (
    <html lang="vi">
      <body>{children}</body>
    </html>
  );
}
