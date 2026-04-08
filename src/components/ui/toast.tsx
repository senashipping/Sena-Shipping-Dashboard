import React, { createContext, useCallback, useContext } from "react";
import { ToastContainer, toast as rtToast, type ToastOptions } from "react-toastify";
import { useTheme } from "../../contexts/ThemeContext";

interface Toast {
  id?: string;
  title: string;
  description?: string;
  variant?: "default" | "destructive" | "success";
  duration?: number;
}

interface ToastContextValue {
  toast: (toast: Omit<Toast, "id">) => void;
}

const ToastContext = createContext<ToastContextValue | undefined>(undefined);

export const useToast = () => {
  const context = useContext(ToastContext);
  if (!context) {
    throw new Error("useToast must be used within a ToastProvider");
  }
  return context;
};

export const ToastProvider: React.FC<{ children: React.ReactNode }> = ({ children }) => {
  const { theme } = useTheme();

  const toast = useCallback((newToast: Omit<Toast, "id">) => {
    const duration = newToast.duration ?? 5000;
    const toastId = `${newToast.variant || "default"}:${newToast.title}:${newToast.description || ""}`;
    const options: ToastOptions = { autoClose: duration, toastId };
    const content = (
      <div>
        <div className="font-semibold">{newToast.title}</div>
        {newToast.description ? (
          <div className="mt-1 text-sm opacity-90">{newToast.description}</div>
        ) : null}
      </div>
    );
    switch (newToast.variant) {
      case "success":
        rtToast.success(content, options);
        break;
      case "destructive":
        rtToast.error(content, options);
        break;
      default:
        rtToast.info(content, options);
    }
  }, []);

  return (
    <ToastContext.Provider value={{ toast }}>
      {children}
      <ToastContainer
        position="top-right"
        theme={theme}
        newestOnTop
        closeOnClick
        rtl={false}
        pauseOnFocusLoss
        draggable
        pauseOnHover
        limit={5}
        style={{ zIndex: 99999 }}
      />
    </ToastContext.Provider>
  );
};
