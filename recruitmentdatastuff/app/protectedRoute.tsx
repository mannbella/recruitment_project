/***************************************************************
 * File: protectedRoute.tsx
 *
 * Date: 11/30/2025
 *
 * Description: route
 *
 * Author: Bella Mann
 ***************************************************************/

import React from 'react';
import { Navigate, useLocation } from 'react-router-dom';
import { useAuth } from './authContext';

interface ProtectedRouteProps {
  children: JSX.Element;
}

const protectedRoute = ({ children }: protectedRouteProps) => {
  const { user } = useAuth();
  const location = useLocation();

  if(!user)
    return <Navigate to="/login" state={{ from: location }} replace />;

  return children;
};

export default protectedRoute;
