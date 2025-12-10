/***************************************************************
 * File: login.tsx
 *
 * Date: 11/30/2025
 *
 * Description: login for website
 *
 * Author: Bella Mann
 ***************************************************************/
// app/page.tsx
'use client'; // This is required because your App.tsx uses React Router
import dynamic from 'next/dynamic';

const App = dynamic(() => import('./App'), { ssr: false });

export default function Home() {
  return <App />;
}