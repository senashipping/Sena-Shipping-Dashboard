"use client";

import React from "react";
import { Link, useLocation } from "react-router-dom";
import { cn } from "../../lib/utils";
import { Button } from "../ui/button";
import { ScrollArea } from "../ui/scroll-area";
import {
  LayoutDashboard,
  Users,
  Ship,
  FileText,
  ClipboardList,
  Bell,
  Cog,
  Anchor,
  Shield,
  AlarmClock,
  User,
  ChevronLeft,
  ChevronRight,
} from "lucide-react";

interface NavItem {
  title: string;
  href: string;
  icon: string;
  badge?: number;
}

interface SidebarProps {
  navItems: NavItem[];
  isMobileMenuOpen?: boolean;
  onMobileMenuClose?: () => void;
}

const iconMap: { [key: string]: React.ComponentType<any> } = {
  LayoutDashboard,
  Users,
  Ship,
  FileText,
  ClipboardList,
  Bell,
  Cog,
  Anchor,
  Shield,
  AlarmClock,
  User,
};

const Sidebar: React.FC<SidebarProps> = ({ navItems, isMobileMenuOpen = false, onMobileMenuClose }) => {
  const [isCollapsed, setIsCollapsed] = React.useState(false);
  const location = useLocation();

  const toggleSidebar = () => {
    setIsCollapsed(!isCollapsed);
  };

  return (
    <>
      {/* Mobile Overlay */}
      {isMobileMenuOpen && (
        <div 
          className="fixed inset-0 z-40 bg-black/50 lg:hidden"
          onClick={onMobileMenuClose}
        />
      )}

      {/* Sidebar */}
      <div
        className={cn(
          "border-r bg-gradient-to-b from-sena-navy to-sena-darkNavy dark:from-sena-darkBg dark:to-sena-darkCard transition-all duration-300 shadow-lg",
          // Desktop behavior
          "hidden lg:block lg:relative",
          isCollapsed && "lg:w-16",
          !isCollapsed && "lg:w-64",
          // Mobile behavior - always show when menu is open
          isMobileMenuOpen && "block fixed inset-y-0 left-0 z-50 w-64",
          !isMobileMenuOpen && "lg:translate-x-0 -translate-x-full lg:block"
        )}
      >
      <div className="flex items-center justify-between h-16 px-4 border-b border-sena-lightBlue/20 dark:border-sena-gold/20">
        {/* Mobile always shows full header, desktop respects collapsed state */}
        {(!isCollapsed || isMobileMenuOpen) && (
          <div className="flex items-center space-x-2">
            <img 
              src="/sena_logo.png" 
              alt="Sena" 
              className="w-auto h-8 brightness-0 invert"
            />
            <div className="flex flex-col">
              <h1 className="text-sm font-bold text-white dark:text-white">Sena Shipping</h1>
              <p className="text-xs text-sena-lightBlue dark:text-white/90">Dashboard</p>
            </div>
          </div>
        )}
        {isCollapsed && !isMobileMenuOpen && (
          <img 
            src="/sena_logo.png" 
            alt="Sena" 
            className="w-auto h-8 mx-auto brightness-0 invert"
          />
        )}
        {/* Hide collapse button on mobile */}
        <Button
          variant="ghost"
          size="icon"
          onClick={toggleSidebar}
          className="hidden w-8 h-8 text-white lg:flex dark:text-white hover:bg-sena-lightBlue/20 dark:hover:bg-sena-gold/20"
        >
          {isCollapsed ? <ChevronRight size={16} /> : <ChevronLeft size={16} />}
        </Button>
      </div>

      <ScrollArea className="h-[calc(100vh-4rem)]">
        <nav className="p-2 space-y-1">
          {navItems.map((item) => {
            const IconComponent = iconMap[item.icon];
            const isActive = location.pathname === item.href;

            return (
              <Link
                key={item.href}
                to={item.href}
                onClick={onMobileMenuClose}
                className={cn(
                  "flex items-center rounded-md px-3 py-2 text-sm font-medium transition-all duration-200 text-white/90 dark:text-white/95 hover:text-white dark:hover:text-sena-gold hover:bg-sena-lightBlue/20 dark:hover:bg-sena-gold/20",
                  isActive ? "bg-sena-gold dark:bg-sena-gold text-sena-navy font-semibold shadow-md" : "transparent",
                  isCollapsed && !isMobileMenuOpen ? "justify-center" : "justify-start"
                )}
              >
                {IconComponent && <IconComponent className="w-4 h-4" />}
                {/* Show text on mobile (isMobileMenuOpen) or desktop when not collapsed */}
                {(!isCollapsed || isMobileMenuOpen) && (
                  <>
                    <span className="ml-3">{item.title}</span>
                    {item.badge !== undefined && item.badge > 0 && (
                      <span className="flex items-center justify-center w-6 h-6 ml-auto text-xs font-bold rounded-full bg-sena-gold text-sena-navy animate-pulse">
                        {item.badge}
                      </span>
                    )}
                  </>
                )}
              </Link>
            );
          })}
        </nav>
      </ScrollArea>
    </div>
    </>
  );
};

export default Sidebar;