/***************************************************************
 * File: pages.tsx
 *
 * Date: 11/30/2025
 *
 * Description: mainpage for website
 *
 * Author: Bella Mann
 ***************************************************************/
'use client'

import React, { useEffect } from "react";
import { BrowserRouter as Router, Routes, Route, Navigate } from 'react-router-dom';
import { useRouter } from 'next/navigation';
import { AuthProvider, useAuth } from './AuthContext';
import ProtectedRoute from './ProtectedRoute';
import Login from './Login';

const Dashboard: React.FC = () => {
  const { user, logout } = useAuth();
  return(
    <div className="flex min-h-screen items-center justify-center bg-zinc-50 font-sans dark:bg-black">
      <main className="flex min-h-screen w-full max-w-3xl flex-col items-center justify-between py-32 px-16 bg-white dark:bg-black sm:items-start">
        <h1>Welcome, {user?.email}</h1>
        <button onClick={logout}>Logout</button>
      </main>
    </div>
  );
};

const App: React.FC = () => {
  return (
    <AuthProvider>
      <Router>
        <Routes>
          <Route path="/login" element={<Login />} />

          <Route path="/dashboard" element={<ProtectedRoute><Dashboard /></ProtectedRoute>} />
        </Routes>
      </Router>
    </AuthProvider>
  );
};

export default App;