/***************************************************************
 * File: login.tsx
 *
 * Date: 11/30/2025
 *
 * Description: login for website
 *
 * Author: Bella Mann
 ***************************************************************/

import React, { useState } from 'react';
import { useNavigate, useLocation} from 'react-router-dom';
import { useAuth } from './authContext';

interface LocationState {
  from?: {
    pathname: string;
  }
}

const Login = () => {
  const [email, setEmail] = useState<string>('');
  const [password, setPassword] = useState<string>('');
  const [error, setError] = useState<string>('');

  const { login } = useAuth();
  const navigate = useNavigate();
  const location = useLocation();

  const state = location.state as LocationState;
  const from = state?.from?.pathname || '/dashboard';

  const handleSubmit = async (e: React.FormEvent<HTMLFormElement>) => {
    e.preventDefault();
    setError('');

    const MASTER_EMAIL: "recruitment.aoii.alpharho@gmail.com"
    const MASTER_PASSWORD: "Passypass7."

    if(email === MASTER_EMAIL && password === MASTER_PASSWORD) {
      await login({email});
      navigate(from, {replace: true});
    } else {
      setError('Invalid email or password');
    }
  };

  return (
    <div className="flex min-h-screen items-center justify-center bg-zinc-50 font-sans dark:bg-black">
      <h2>Login Page</h2>
      <main className="flex min-h-screen w-full max-w-3xl flex-col items-center justify-between py-32 px-16 bg-white dark:bg-black sm:items-start">
        <form onSubmit={handleSubmit}>
          <label>Email:</label>
          <input type="email" value={email} onChange={(e) => setEmail(e.target.value)} required/>
          <label>Password:</label>
          <input type="password" value={password} onChange={(e) => setPassword(e.target.value)} required/>
          {error && <p style={{ color: 'red' }}>{error}</p>}
          <button type="submit">Login</button>
        </form>
      </main>
    </div>
  )
}