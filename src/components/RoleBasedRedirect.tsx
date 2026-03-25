import React from "react";
import { Navigate } from "react-router-dom";
import { useAuth } from "../contexts/AuthContext";
import { Loader2 } from "lucide-react";

const RoleBasedRedirect: React.FC = () => {
  const { auth } = useAuth();

  if (auth.loading) {
    return (
      <div className="flex items-center justify-center min-h-screen">
        <Loader2 className="w-8 h-8 animate-spin" />
      </div>
    );
  }

  if (!auth.isAuthenticated) {
    return <Navigate to="/login" replace />;
  }

  // Redirect based on user role
  if (auth.user?.role === "admin" || auth.user?.role === "super_admin") {
    return <Navigate to="/admin" replace />;
  } else {
    return <Navigate to="/dashboard" replace />;
  }
};

export default RoleBasedRedirect;