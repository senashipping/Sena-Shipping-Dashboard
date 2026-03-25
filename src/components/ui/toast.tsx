import React, { createContext, useContext, useState } from 'react';
import { X, CheckCircle, AlertCircle, Info } from 'lucide-react';

interface Toast {
  id: string;
  title: string;
  description?: string;
  variant?: 'default' | 'destructive' | 'success';
  duration?: number;
}

interface ToastContextValue {
  toast: (toast: Omit<Toast, 'id'>) => void;
}

const ToastContext = createContext<ToastContextValue | undefined>(undefined);

export const useToast = () => {
  const context = useContext(ToastContext);
  if (!context) {
    throw new Error('useToast must be used within a ToastProvider');
  }
  return context;
};

export const ToastProvider: React.FC<{ children: React.ReactNode }> = ({ children }) => {
  const [toasts, setToasts] = useState<Toast[]>([]);

  const toast = (newToast: Omit<Toast, 'id'>) => {
    const id = Math.random().toString(36).substr(2, 9);
    const toastWithId = { ...newToast, id };
    
    setToasts(prev => [...prev, toastWithId]);

    // Auto remove after duration
    setTimeout(() => {
      setToasts(prev => prev.filter(t => t.id !== id));
    }, newToast.duration || 5000);
  };

  const removeToast = (id: string) => {
    setToasts(prev => prev.filter(t => t.id !== id));
  };

  const getIcon = (variant?: string) => {
    switch (variant) {
      case 'success':
        return <CheckCircle className="w-5 h-5 text-green-600" />;
      case 'destructive':
        return <AlertCircle className="w-5 h-5 text-red-600" />;
      default:
        return <Info className="w-5 h-5 text-blue-600" />;
    }
  };

  const getStyles = (variant?: string) => {
    switch (variant) {
      case 'success':
        return 'bg-green-50 border-green-200 text-green-800';
      case 'destructive':
        return 'bg-red-50 border-red-200 text-red-800';
      default:
        return 'bg-blue-50 border-blue-200 text-blue-800';
    }
  };

  return (
    <ToastContext.Provider value={{ toast }}>
      {children}
      
      {/* Toast Container */}
      <div className="fixed top-4 right-4 z-50 flex flex-col gap-2 max-w-md">
        {toasts.map((toast) => (
          <div
            key={toast.id}
            className={`
              flex items-start gap-3 p-4 rounded-lg border shadow-lg
              animate-in slide-in-from-right-full duration-300
              ${getStyles(toast.variant)}
            `}
          >
            {getIcon(toast.variant)}
            <div className="flex-1 min-w-0">
              <div className="font-semibold text-sm">{toast.title}</div>
              {toast.description && (
                <div className="text-sm opacity-90 mt-1">{toast.description}</div>
              )}
            </div>
            <button
              onClick={() => removeToast(toast.id)}
              className="text-gray-400 hover:text-gray-600 transition-colors"
            >
              <X className="w-4 h-4" />
            </button>
          </div>
        ))}
      </div>
    </ToastContext.Provider>
  );
};