/***************************************************************
 * File: layout.tsx
 *
 * Date: 11/30/2025
 *
 * Description: layout for website
 *
 * Author: Bella Mann
 ***************************************************************/
import './globals.css';

export default function RootLayout({
  children,
}: {
  children: React.ReactNode;
}) {
  return (
    <html lang="en">
      <body>
        {/* This {children} is where your App.tsx gets injected */}
        {children}
      </body>
    </html>
  );
}