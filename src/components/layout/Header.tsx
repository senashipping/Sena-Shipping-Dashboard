"use client";

import React from "react";
import { useAuth } from "../../contexts/AuthContext";
import { Button } from "../ui/button";
import {
  DropdownMenu,
  DropdownMenuContent,
  DropdownMenuItem,
  DropdownMenuLabel,
  DropdownMenuSeparator,
  DropdownMenuTrigger,
} from "../ui/dropdown-menu";
import { LogOut, User as UserIcon, Menu } from "lucide-react";
import ThemeToggle from "./ThemeToggle";

interface HeaderProps {
  user: any;
  onMobileMenuToggle?: () => void;
}

const Header: React.FC<HeaderProps> = ({ user, onMobileMenuToggle }) => {
  const { logout } = useAuth();

  return (
    <header className="flex items-center justify-between h-16 px-3 sm:px-6 transition-colors border-b shadow-sm bg-background dark:shadow-lg">
      <div className="flex items-center space-x-2 sm:space-x-3 min-w-0">
        {/* Mobile Menu Button */}
        {onMobileMenuToggle && (
          <Button
            variant="ghost"
            size="icon"
            onClick={onMobileMenuToggle}
            className="lg:hidden w-8 h-8 p-0 text-sena-navy dark:text-white hover:bg-sena-lightBlue/20 dark:hover:bg-sena-gold/20"
          >
            <Menu className="w-5 h-5" />
          </Button>
        )}
        <img 
          src="/sena_logo.png" 
          alt="Logo" 
          className="w-auto h-12 sm:h-14 flex-shrink-0 dark:brightness-0 dark:invert"
        />
      </div>

      <div className="flex items-center space-x-2 sm:space-x-4 flex-shrink-0">
        <ThemeToggle />
        <DropdownMenu>
          <DropdownMenuTrigger asChild>
            <Button variant="ghost" className="relative w-9 h-9 sm:w-10 sm:h-10 p-0 transition-colors rounded-full hover:bg-muted/50">
              <div className="flex items-center justify-center w-9 h-9 sm:w-10 sm:h-10 transition-all rounded-full shadow-md bg-gradient-to-r from-sena-navy to-sena-lightBlue dark:from-sena-gold dark:to-sena-lightBlue hover:shadow-lg">
                <UserIcon className="w-4 h-4 sm:w-5 sm:h-5 text-white dark:text-sena-navy" strokeWidth={1.5} />
              </div>
            </Button>
          </DropdownMenuTrigger>
          <DropdownMenuContent className="w-48 sm:w-56" align="end" forceMount>
            <DropdownMenuLabel className="font-normal">
              <div className="flex flex-col space-y-1">
                <p className="text-sm font-medium leading-none">
                  {user?.name}
                </p>
                <p className="text-xs leading-none text-muted-foreground">
                  {user?.email}
                </p>
                <p className="text-xs leading-none text-muted-foreground">
                  {user?.role?.charAt(0).toUpperCase() + user?.role?.slice(1)}
                </p>
                {user?.ship && (
                  <p className="text-xs leading-none text-muted-foreground">
                    Ship: {user.ship.name}
                  </p>
                )}
              </div>
            </DropdownMenuLabel>
            <DropdownMenuSeparator />
            {/* <DropdownMenuItem>
              <Settings className="w-4 h-4 mr-2" />
              <span>Settings</span>
            </DropdownMenuItem> */}
            <DropdownMenuSeparator />
            <DropdownMenuItem onClick={logout}>
              <LogOut className="w-4 h-4 mr-2" />
              <span>Log out</span>
            </DropdownMenuItem>
          </DropdownMenuContent>
        </DropdownMenu>
      </div>
    </header>
  );
};

export default Header;