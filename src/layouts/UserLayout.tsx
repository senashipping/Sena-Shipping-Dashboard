"use client";

import React from "react";
import { Outlet } from "react-router-dom";
import { useAuth } from "../contexts/AuthContext";
import { useNotifications } from "../contexts/NotificationContext";
import Sidebar from "../components/layout/Sidebar";
import Header from "../components/layout/Header";

const UserLayout: React.FC = () => {
  const { auth, canAccessCategory } = useAuth();
  const { unreadCount } = useNotifications();
  const [isMobileMenuOpen, setIsMobileMenuOpen] = React.useState(false);

  // Define all possible category navigation items
  const allCategoryItems = [
    {
      title: "Engine Forms",
      href: "/dashboard/forms/category/eng",
      icon: "Cog",
      category: "eng"
    },
    {
      title: "Deck Forms",
      href: "/dashboard/forms/category/deck",
      icon: "Anchor",
      category: "deck"
    },
    {
      title: "MLC Forms",
      href: "/dashboard/forms/category/mlc",
      icon: "Users",
      category: "mlc"
    },
    {
      title: "ISPS Forms",
      href: "/dashboard/forms/category/isps",
      icon: "Shield",
      category: "isps"
    },
    {
      title: "Drill Forms",
      href: "/dashboard/forms/category/drill",
      icon: "AlarmClock",
      category: "drill"
    },
  ];

  // Filter category items based on user access
  const accessibleCategoryItems = allCategoryItems.filter(item => 
    canAccessCategory(item.category)
  );

  const userNavItems = [
    {
      title: "Dashboard",
      href: "/dashboard",
      icon: "LayoutDashboard",
    },
    ...accessibleCategoryItems,
    {
      title: "Notifications",
      href: "/dashboard/notifications",
      icon: "Bell",
      badge: unreadCount,
    },
    {
      title: "Profile",
      href: "/dashboard/profile",
      icon: "User",
    },
  ];

  return (
    <div className="flex h-screen bg-background">
      <Sidebar 
        navItems={userNavItems}
        isMobileMenuOpen={isMobileMenuOpen}
        onMobileMenuClose={() => setIsMobileMenuOpen(false)}
      />
      <div className="flex flex-col flex-1 overflow-hidden">
        <Header 
          user={auth.user} 
          onMobileMenuToggle={() => setIsMobileMenuOpen(!isMobileMenuOpen)}
        />
        <main className="flex-1 p-3 sm:p-6 overflow-x-hidden overflow-y-auto bg-background">
          <Outlet />
        </main>
      </div>
    </div>
  );
};

export default UserLayout;