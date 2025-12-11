/***************************************************************
 * File: app.tsx
 *
 * Date: 12/10/25
 *
 * Description: mainpage for website
 *
 * Author: Bella Mann
 ***************************************************************/
'use client'

import React, { useEffect } from "react";
import { BrowserRouter as Router, Routes, Route, Navigate } from 'react-router-dom';
import { useRouter } from 'next/navigation';
import { AuthProvider, useAuth } from './authContext';
import ProtectedRoute from './protectedRoute';
import Login from './login';

const Dashboard: React.FC = () => {
  const { user, logout } = useAuth();
  const [selectedFile, setSelectedFile] = useState<File | null>(null);

  const handleFileSelection = (event: React.ChangeEvent<HTMLInputElement>) => {
    if(event.target.files && event.target.files.length >0 )
      setSelectedFile(event.target.files[0])
  };

  const handleUpload = () => {
    if(selectedFile)
      alert(`File "${selectedFile.name}" selected for upload!`);
    else
      alert("please select a file first");
  }


  return(
    <div className="flex min-h-screen items-center justify-center bg-zinc-50 font-sans dark:bg-black">
      <main className="flex min-h-screen w-full max-w-3xl flex-col items-center justify-between py-32 px-16 bg-white dark:bg-black sm:items-start">
        <h1>Welcome, {user?.email}</h1>
        <input type="file" onChange={handleFileSelection}/>
        <br></br>
        <button onClick={handleUPload}>Upload File</button>
        {selectedFile && <p>Selected: {selectedFile.name}</p>}
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
          <Route path="/" element={<Navigate to="/login" replace />} />
          <Route path="/login" element={<Login />} />

          <Route path="/dashboard" element={<ProtectedRoute><Dashboard /></ProtectedRoute>} />
          <Route path="/" element={<Navigate to="/login" replace />} />
        </Routes>
      </Router>
    </AuthProvider>
  );
};

export default App;