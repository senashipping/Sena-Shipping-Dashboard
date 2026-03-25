"use client";

import React from "react";
import { Navigate, useLocation } from "react-router-dom";
import { useAuth } from "../contexts/AuthContext";
import { Loader2 } from "lucide-react";

interface ProtectedRouteProps {
  requiredRole?: "user" | "admin" | "super_admin";
  children: React.ReactNode;
}

const ProtectedRoute: React.FC<ProtectedRouteProps> = ({ requiredRole, children }) => {
  const { auth } = useAuth();
  const location = useLocation();

  if (auth.loading) {
    return (
      <div className="flex items-center justify-center min-h-screen">
        <Loader2 className="h-8 w-8 animate-spin" />
      </div>
    );
  }

  if (!auth.isAuthenticated) {
    return <Navigate to="/login" state={{ from: location }} replace />;
  }

  if (requiredRole) {
    // Check role hierarchy - super_admin can access admin routes, admin can access user routes
    const hasPermission = 
      (requiredRole === "user" && ["user", "admin", "super_admin"].includes(auth.user?.role || "")) ||
      (requiredRole === "admin" && ["admin", "super_admin"].includes(auth.user?.role || "")) ||
      (requiredRole === "super_admin" && auth.user?.role === "super_admin");

    if (!hasPermission) {
      // Redirect to appropriate dashboard based on user role
      if (auth.user?.role === "admin" || auth.user?.role === "super_admin") {
        return <Navigate to="/admin" replace />;
      } else {
        return <Navigate to="/dashboard" replace />;
      }
    }
  }

  return <>{children}</>;
};

export default ProtectedRoute;