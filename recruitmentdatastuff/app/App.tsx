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

import React, { useEffect, useState, useRef } from "react";
import { BrowserRouter as Router, Routes, Route, Navigate } from 'react-router-dom';
import { useRouter } from 'next/navigation';
import { AuthProvider, useAuth } from './authContext';
import ProtectedRoute from './protectedRoute';
import Login from './login';

const Dashboard: React.FC = () => {
  const dataTypes = ["Sisterhood Data", "Philanthropy Data", "House Tours Data", "Preference Data"];

  const { user, logout } = useAuth();
  const [selectedFiles, setSelectedFiles] = useState<File | null>(null);
  const [checkedItems, setCheckedItems] = useState<Record<String, boolean>>({});
  const fileInputRef = useRef<HTMLInputElement>(null); 

  const handleFileSelection = (event: React.ChangeEvent<HTMLInputElement>) => {
    if(event.target.files && event.target.files.length >0 )
      setSelectedFiles(event.target.files[0])
  };

  const handleClear = () => {
    setSelectedFiles(null);
    if(fileInputRef.current)
      fileInputRef.current.value = '';
  };

  const handleUpload = () => {
    if (selectedFiles) {
      alert(`File "${selectedFiles.name}" uploading...`);
    } else {
      alert("Please select a file first.");
    }
  };

  const handleCheckboxChange = (category: string) => {
    setCheckedItems(prev => {
      const newState = {...prev, [category]: !prev[category]};
      if(!newState[category]) {
        setSelectedFiles(prevFiles => {
          const newFiles = { ...prevFiles};
          delete newFiles[category];
          return newFiles;
        });
      }
      return newState;
    });
  };

  return(
    <div className="flex min-h-screen items-center justify-center bg-zinc-50 font-sans dark:bg-black">
      <main className="flex min-h-screen w-full max-w-3xl flex-col items-center justify-between py-32 px-16 bg-white dark:bg-black sm:items-start">
        <h1>Welcome, {user?.email}</h1>
        
        <div className="flex flex-col gap-6 w-full">
          {/* 4. Map over the list to generate independent sections */}
          {dataTypes.map((category) => (
            <label key={category} className="flex flex-col gap-2 p-4 border rounded shadow-sm">
              <span className="font-bold">{category}</span>
              
              <div className="flex items-center gap-4">
                {/* CHECKBOX */}
                <input 
                  type="checkbox" 
                  checked={!!checkedItems[category]} 
                  onChange={() => handleCheckboxChange(category)}
                  className="w-5 h-5"
                />
                
                {/* FILE INPUT */}
                <input 
                  type="file" 
                  disabled={!checkedItems[category]} 
                  onChange={(e) => handleFileSelection(category, e)}
                  // Use 'key' to reset input when unchecked (clears the browser text)
                  key={checkedItems[category] ? "active" : "disabled"}
                />
              </div>

              {/* Show selected file name if exists */}
              {selectedFiles[category] && (
                <p className="text-sm text-green-600">
                  Ready to upload: {selectedFiles[category]?.name}
                </p>
              )}
            </label>
          ))}
        </div>
        
        <br></br>
        <button onClick={handleUpload} disabled={Object.keys(selectedFiles).length === 0}>Upload All Files</button>
        {selectedFiles && (<button onClick={handleClear}>Clear</button>)}
        {selectedFiles && <p>Selected: {selectedFiles?.name}</p>}
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