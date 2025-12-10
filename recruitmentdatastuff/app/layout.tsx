/***************************************************************
 * File: login.tsx
 *
 * Date: 11/30/2025
 *
 * Description: login for website
 *
 * Author: Bella Mann
 ***************************************************************/
// app/layout.tsx
import './globals.css'; // This keeps your styling working

export default function RootLayout({
  children,
}: {
  children: React.ReactNode;
}) {
  return (
    <html lang="en">
      <body>{children}</body>
    </html>
  );
}