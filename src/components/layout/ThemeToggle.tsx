import React from 'react';
import { Moon, Sun } from 'lucide-react';
import { Button } from '../ui/button';
import { useTheme } from '../../contexts/ThemeContext';

const ThemeToggle: React.FC = () => {
  const { theme, toggleTheme } = useTheme();

  return (
    <Button
      variant="ghost"
      size="icon"
      onClick={toggleTheme}
      className="relative h-9 w-9 rounded-full hover:bg-sena-lightBlue/10 dark:hover:bg-sena-gold/10 transition-colors"
      aria-label={`Switch to ${theme === 'light' ? 'dark' : 'light'} mode`}
    >
      <Sun className="h-4 w-4 rotate-0 scale-100 transition-all dark:-rotate-90 dark:scale-0 text-sena-gold" />
      <Moon className="absolute h-4 w-4 rotate-90 scale-0 transition-all dark:rotate-0 dark:scale-100 text-sena-lightBlue" />
    </Button>
  );
};

export default ThemeToggle;