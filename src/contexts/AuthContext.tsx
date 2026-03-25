"use client";

import React, { createContext, useContext, useState, useEffect, ReactNode } from "react";
import { useNavigate } from "react-router-dom";
import api from "../api";
import { User, AuthState } from "../types";

interface AuthContextType {
  auth: AuthState;
  login: (email: string, password: string) => Promise<void>;
  logout: () => void;
  register: (userData: any) => Promise<void>;
  updateProfile: (profileData: any) => Promise<void>;
  changePassword: (passwordData: any) => Promise<void>;
  refreshUser: () => Promise<void>;
  isUser: () => boolean;
  isAdmin: () => boolean;
  isSuperAdmin: () => boolean;
  isDeckUser: () => boolean;
  isEngineUser: () => boolean;
  getUserType: () => string | null;
  canAccessCategory: (category: string) => boolean;
  getAccessibleCategories: () => string[];
  getCategoryDisplayInfo: () => any;
}

const AuthContext = createContext<AuthContextType | undefined>(undefined);

export const AuthProvider: React.FC<{ children: ReactNode }> = ({ children }) => {
  const [auth, setAuth] = useState<AuthState>({
    user: null,
    token: null,
    refreshToken: null,
    isAuthenticated: false,
    loading: true,
  });
  
  const navigate = useNavigate();

  useEffect(() => {
    // Check for existing token in localStorage
    const token = localStorage.getItem("token");
    const refreshToken = localStorage.getItem("refreshToken");
    const user = localStorage.getItem("user");

    if (token && refreshToken && user) {
      try {
        const userData: User = JSON.parse(user);
        setAuth({
          user: userData,
          token,
          refreshToken,
          isAuthenticated: true,
          loading: false,
        });
        api.setAuthState({
          user: userData,
          token,
          refreshToken,
          isAuthenticated: true,
          loading: false,
        });
      } catch (error) {
        console.error("Error parsing user data:", error);
        setAuth({
          user: null,
          token: null,
          refreshToken: null,
          isAuthenticated: false,
          loading: false,
        });
        api.setAuthState({
          user: null,
          token: null,
          refreshToken: null,
          isAuthenticated: false,
          loading: false,
        });
      }
    } else {
      setAuth({
        user: null,
        token: null,
        refreshToken: null,
        isAuthenticated: false,
        loading: false,
      });
      api.setAuthState({
        user: null,
        token: null,
        refreshToken: null,
        isAuthenticated: false,
        loading: false,
      });
    }
  }, []);

  const login = async (email: string, password: string) => {
    try {
      const response = await api.login(email, password);
      const { user, token, refreshToken } = response.data.data;

      setAuth({
        user,
        token,
        refreshToken,
        isAuthenticated: true,
        loading: false,
      });

      // Save to localStorage
      localStorage.setItem("token", token);
      localStorage.setItem("refreshToken", refreshToken);
      localStorage.setItem("user", JSON.stringify(user));

      api.setAuthState({
        user,
        token,
        refreshToken,
        isAuthenticated: true,
        loading: false,
      });
      
      // Navigate to appropriate dashboard based on role
      if (user.role === "admin" || user.role === "super_admin") {
        navigate("/admin");
      } else {
        navigate("/dashboard");
      }
    } catch (error) {
      throw error;
    }
  };

  const logout = () => {
    setAuth({
      user: null,
      token: null,
      refreshToken: null,
      isAuthenticated: false,
      loading: false,
    });

    // Clear localStorage
    localStorage.removeItem("token");
    localStorage.removeItem("refreshToken");
    localStorage.removeItem("user");

    api.logout();
    navigate("/login");
  };

  const register = async (userData: any) => {
    try {
      const response = await api.register(userData);
      const { user, token, refreshToken } = response.data;

      setAuth({
        user,
        token,
        refreshToken,
        isAuthenticated: true,
        loading: false,
      });

      // Save to localStorage
      localStorage.setItem("token", token);
      localStorage.setItem("refreshToken", refreshToken);
      localStorage.setItem("user", JSON.stringify(user));

      api.setAuthState({
        user,
        token,
        refreshToken,
        isAuthenticated: true,
        loading: false,
      });
      
      // Navigate based on role (users typically register as "user" role)
      if (user.role === "admin" || user.role === "super_admin") {
        navigate("/admin");
      } else {
        navigate("/dashboard");
      }
    } catch (error) {
      throw error;
    }
  };

  const updateProfile = async (profileData: any) => {
    try {
      const response = await api.updateProfile(profileData);
      const updatedUser = response.data;

      setAuth(prev => ({
        ...prev,
        user: updatedUser,
      }));

      // Update localStorage
      localStorage.setItem("user", JSON.stringify(updatedUser));

      api.setAuthState({
        ...auth,
        user: updatedUser,
      });
    } catch (error) {
      throw error;
    }
  };

  const changePassword = async (passwordData: any) => {
    try {
      await api.changePassword(passwordData);
    } catch (error) {
      throw error;
    }
  };

  const refreshUser = async () => {
    if (!auth.token) return;

    try {
      const response = await api.getProfile();
      const user = response.data;

      setAuth(prev => ({
        ...prev,
        user,
      }));

      // Update localStorage
      localStorage.setItem("user", JSON.stringify(user));

      api.setAuthState({
        ...auth,
        user,
      });
    } catch (error) {
      console.error("Error refreshing user:", error);
    }
  };

  // Role helper functions
  const isUser = (): boolean => {
    return auth.user?.role === "user" || false;
  };

  const isAdmin = (): boolean => {
    return auth.user?.role === "admin" || auth.user?.role === "super_admin" || false;
  };

  const isSuperAdmin = (): boolean => {
    return auth.user?.role === "super_admin" || false;
  };

  // User type helper functions
  const isDeckUser = (): boolean => {
    return auth.user?.role === "user" && auth.user?.userType === "deck" || false;
  };

  const isEngineUser = (): boolean => {
    return auth.user?.role === "user" && auth.user?.userType === "engine" || false;
  };

  const getUserType = (): string | null => {
    return auth.user?.userType || null;
  };

  // Form access helper functions
  const canAccessCategory = (category: string): boolean => {
    // Admins can access everything
    if (isAdmin()) return true;
    
    // For regular users, check based on userType
    if (isEngineUser()) {
      // Engine Officers can access Engine forms and Deck+Engine forms
      return ["eng", "deck_engine"].includes(category);
    }
    
    if (isDeckUser()) {
      // Deck Officers can access all forms except Engine-only, including Deck+Engine forms
      return ["deck", "mlc", "isps", "drill", "deck_engine"].includes(category);
    }
    
    return false;
  };

  const getAccessibleCategories = (): string[] => {
    // Admins can access everything
    if (isAdmin()) return ["deck", "eng", "mlc", "isps", "drill", "deck_engine"];
    
    // For regular users, return based on userType
    if (isEngineUser()) {
      // Engine Officers can access Engine forms and Deck+Engine forms
      return ["eng", "deck_engine"];
    }
    
    if (isDeckUser()) {
      // Deck Officers can access all forms except Engine-only, including Deck+Engine forms
      return ["deck", "mlc", "isps", "drill", "deck_engine"];
    }
    
    return [];
  };

  const getCategoryDisplayInfo = () => {
    return {
      eng: { name: "Engine Forms", description: "Forms related to engine operations and maintenance", icon: "🔧" },
      deck: { name: "Deck Forms", description: "Forms for deck operations and equipment", icon: "⚓" },
      mlc: { name: "MLC Forms", description: "Maritime Labour Convention compliance forms", icon: "👥" },
      isps: { name: "ISPS Forms", description: "International Ship and Port Facility Security forms", icon: "🛡️" },
      drill: { name: "Drill Forms", description: "Safety drill and emergency response forms", icon: "🚨" },
      deck_engine: { name: "Deck + Engine Forms", description: "Forms for both Deck and Engine departments", icon: "🤝" }
    };
  };

  return (
    <AuthContext.Provider
      value={{
        auth,
        login,
        logout,
        register,
        updateProfile,
        changePassword,
        refreshUser,
        isUser,
        isAdmin,
        isSuperAdmin,
        isDeckUser,
        isEngineUser,
        getUserType,
        canAccessCategory,
        getAccessibleCategories,
        getCategoryDisplayInfo,
      }}
    >
      {children}
    </AuthContext.Provider>
  );
};

export const useAuth = () => {
  const context = useContext(AuthContext);
  if (context === undefined) {
    throw new Error("useAuth must be used within an AuthProvider");
  }
  return context;
};