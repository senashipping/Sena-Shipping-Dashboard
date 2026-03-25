"use client";

import React from "react";
import { Outlet } from "react-router-dom";
import { useAuth } from "../contexts/AuthContext";
import { useNotifications } from "../contexts/NotificationContext";
import Sidebar from "../components/layout/Sidebar";
import Header from "../components/layout/Header";

const AdminLayout: React.FC = () => {
  const { auth } = useAuth();
  const { unreadCount } = useNotifications();
  const [isMobileMenuOpen, setIsMobileMenuOpen] = React.useState(false);

  const adminNavItems = [
    {
      title: "Dashboard",
      href: "/admin/dashboard",
      icon: "LayoutDashboard",
    },
    {
      title: "Users",
      href: "/admin/users",
      icon: "Users",
    },
    {
      title: "Ships",
      href: "/admin/ships",
      icon: "Ship",
    },
    {
      title: "Forms",
      href: "/admin/forms",
      icon: "FileText",
    },
    {
      title: "Submissions",
      href: "/admin/submissions",
      icon: "ClipboardList",
    },
    {
      title: "Notifications",
      href: "/admin/notifications",
      icon: "Bell",
      badge: unreadCount,
    },
  ];

  return (
    <div className="flex h-screen bg-background">
      <Sidebar 
        navItems={adminNavItems}
        isMobileMenuOpen={isMobileMenuOpen}
        onMobileMenuClose={() => setIsMobileMenuOpen(false)}
      />
      <div className="flex-1 flex flex-col overflow-hidden">
        <Header 
          user={auth.user} 
          onMobileMenuToggle={() => setIsMobileMenuOpen(!isMobileMenuOpen)}
        />
        <main className="flex-1 overflow-x-hidden overflow-y-auto p-3 sm:p-6 bg-background">
          <Outlet />
        </main>
      </div>
    </div>
  );
};

export default AdminLayout;