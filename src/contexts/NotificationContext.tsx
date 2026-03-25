"use client";

import React, { createContext, useContext, useState, ReactNode, useMemo } from "react";
import { useQuery, useQueryClient } from "@tanstack/react-query";
import api from "../api";
import { Notification } from "../types";
import { useAuth } from "./AuthContext";

interface NotificationContextType {
  notifications: Notification[];
  unreadCount: number;
  markAsRead: (id: string) => Promise<void>;
  markAllAsRead: () => Promise<void>;
  deleteNotification: (id: string) => Promise<void>;
  refreshNotifications: () => void;
  checkUnfilledForms: () => Promise<void>;
}

const NotificationContext = createContext<NotificationContextType | undefined>(undefined);

export const NotificationProvider: React.FC<{ children: ReactNode }> = ({ children }) => {
  const { auth, canAccessCategory } = useAuth();
  const queryClient = useQueryClient();
  const [unreadCount, setUnreadCount] = useState(0);

  const {
    data: allNotifications = [],
    refetch,
  } = useQuery<Notification[]>({
    queryKey: ["notifications"],
    queryFn: async () => {
      const response = await api.getNotifications({ limit: 50 });
      return response.data.data.notifications || [];
    },
    enabled: auth.isAuthenticated, // This is the fix!
    refetchInterval: 30000, // Refetch every 30 seconds
  });

  // Filter notifications based on user access to form categories
  const notifications = useMemo(() => {
    if (!Array.isArray(allNotifications)) {
      return [];
    }
    
    return allNotifications.filter((notification: Notification) => {
      // If notification has no related form, show it (system notifications, etc.)
      if (!notification.relatedForm) {
        return true;
      }

      // Check if form and category exist before accessing properties
      if (!notification.relatedForm.category || !notification.relatedForm.category.name) {
        return true; // Show notifications without proper category info
      }

      // Check if user can access the form's category
      return canAccessCategory(notification.relatedForm.category.name);
    });
  }, [allNotifications, canAccessCategory]);

  // Update unread count based on filtered notifications
  React.useEffect(() => {
    const filteredUnreadCount = notifications.filter((n: Notification) => !n.isRead).length;
    setUnreadCount(filteredUnreadCount);
  }, [notifications]);

  React.useEffect(() => {
    if (auth.isAuthenticated) {
      checkUnfilledForms();
    }
  }, [auth.isAuthenticated]);

  const checkUnfilledForms = async () => {
    try {
      await api.checkUnfilledForms();
      await refetch();
    } catch (error) {
      console.error("Error checking for unfilled forms:", error);
    }
  };

  const markAsRead = async (id: string) => {
    try {
      await api.markNotificationAsRead(id);
      await refetch();
      queryClient.invalidateQueries({ queryKey: ["notificationStats"] });
    } catch (error) {
      console.error("Error marking notification as read:", error);
    }
  };

  const markAllAsRead = async () => {
    try {
      await api.markAllNotificationsAsRead();
      await refetch();
      setUnreadCount(0);
      queryClient.invalidateQueries({ queryKey: ["notificationStats"] });
    } catch (error) {
      console.error("Error marking all notifications as read:", error);
    }
  };

  const deleteNotification = async (id: string) => {
    try {
      await api.deleteNotification(id);
      await refetch();
      queryClient.invalidateQueries({ queryKey: ["notificationStats"] });
    } catch (error) {
      console.error("Error deleting notification:", error);
    }
  };

  const refreshNotifications = () => {
    refetch();
  };

  return (
    <NotificationContext.Provider
      value={{
        notifications,
        unreadCount,
        markAsRead,
        markAllAsRead,
        deleteNotification,
        refreshNotifications,
        checkUnfilledForms,
      }}
    >
      {children}
    </NotificationContext.Provider>
  );
};

export const useNotifications = () => {
  const context = useContext(NotificationContext);
  if (context === undefined) {
    throw new Error("useNotifications must be used within a NotificationProvider");
  }
  return context;
};