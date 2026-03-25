import React, { useState, useMemo } from "react";
import { useNavigate } from "react-router-dom";
import { Button } from "../../components/ui/button";
import { Card, CardContent } from "../../components/ui/card";
import { Badge } from "../../components/ui/badge";
import { Select, SelectContent, SelectItem, SelectTrigger, SelectValue } from "../../components/ui/select";
import { Bell, Check, Trash2, ChevronLeft, ChevronRight, Filter } from "lucide-react";
import { useNotifications } from "../../contexts/NotificationContext";
import { useAuth } from "../../contexts/AuthContext";
import { formatDateTime } from "../../lib/utils";

const Notifications: React.FC = () => {
  const { notifications, unreadCount, markAsRead, markAllAsRead, deleteNotification } = useNotifications();
  const { auth } = useAuth();
  const navigate = useNavigate();

  // Pagination state
  const [currentPage, setCurrentPage] = useState(1);
  const [itemsPerPage, setItemsPerPage] = useState(10);
  const [filterType, setFilterType] = useState<'all' | 'unread' | 'read'>('all');

  // Helper function to mark notification as read and navigate
  const handleNotificationAction = async (notificationId: string, path: string) => {
    try {
      await markAsRead(notificationId);
      navigate(path);
    } catch (error) {
      console.error('Failed to mark notification as read:', error);
      // Still navigate even if marking as read fails
      navigate(path);
    }
  };



  // Filter and pagination logic
  const filteredNotifications = useMemo(() => {
    let filtered = [...notifications];
    
    // Apply filter
    if (filterType === 'unread') {
      filtered = filtered.filter(n => !n.isRead);
    } else if (filterType === 'read') {
      filtered = filtered.filter(n => n.isRead);
    }
    
    // Sort by date (newest first)
    filtered.sort((a, b) => new Date(b.createdAt).getTime() - new Date(a.createdAt).getTime());
    
    return filtered;
  }, [notifications, filterType]);

  const totalPages = Math.ceil(filteredNotifications.length / itemsPerPage);
  const paginatedNotifications = filteredNotifications.slice(
    (currentPage - 1) * itemsPerPage,
    currentPage * itemsPerPage
  );

  // Reset to page 1 when filter changes
  React.useEffect(() => {
    setCurrentPage(1);
  }, [filterType]);

  return (
    <div className="space-y-6">
      <div className="flex flex-col gap-4 sm:flex-row sm:items-center sm:justify-between">
        <div>
          <h1 className="text-3xl font-bold">Notifications</h1>
          <p className="mt-1 ">
            Showing {paginatedNotifications.length} of {filteredNotifications.length} notifications
          </p>
        </div>
        <div className="flex items-center gap-2">
          {unreadCount > 0 && (
            <Button variant="outline" onClick={markAllAsRead}>
              <Check className="w-4 h-4 mr-2" /> Mark All as Read
            </Button>
          )}
          <Button 
            variant="destructive" 
            onClick={() => notifications.filter(n => n.isRead).forEach(n => deleteNotification(n._id))}
            disabled={notifications.filter(n => n.isRead).length === 0}
          >
            <Trash2 className="w-4 h-4 mr-2" /> Clear Read
          </Button>
        </div>
      </div>

      {/* Filters and Controls */}
      <Card>
        <CardContent className="p-4">
          <div className="flex flex-col gap-4 sm:flex-row sm:items-center sm:justify-between">
            <div className="flex items-center gap-4">
              <div className="flex items-center gap-2">
                <Filter className="w-4 h-4 " />
                <Select value={filterType} onValueChange={(value: 'all' | 'unread' | 'read') => setFilterType(value)}>
                  <SelectTrigger className="w-32">
                    <SelectValue />
                  </SelectTrigger>
                  <SelectContent>
                    <SelectItem value="all">All</SelectItem>
                    <SelectItem value="unread">Unread</SelectItem>
                    <SelectItem value="read">Read</SelectItem>
                  </SelectContent>
                </Select>
              </div>
              
              <div className="flex items-center gap-2">
                <span className="text-sm ">Show:</span>
                <Select 
                  value={itemsPerPage.toString()} 
                  onValueChange={(value) => {
                    setItemsPerPage(Number(value));
                    setCurrentPage(1);
                  }}
                >
                  <SelectTrigger className="w-20">
                    <SelectValue />
                  </SelectTrigger>
                  <SelectContent>
                    <SelectItem value="5">5</SelectItem>
                    <SelectItem value="10">10</SelectItem>
                    <SelectItem value="20">20</SelectItem>
                    <SelectItem value="50">50</SelectItem>
                  </SelectContent>
                </Select>
              </div>
            </div>

            <div className="flex items-center gap-2">
              <Badge variant="secondary">
                {unreadCount} unread
              </Badge>
              <Badge variant="outline">
                {filteredNotifications.length} total
              </Badge>
            </div>
          </div>
        </CardContent>
      </Card>

      {/* Notifications List */}
      <Card>
        <CardContent className="p-6">
          <div className="space-y-4">
            {filteredNotifications.length === 0 ? (
              <div className="py-12 text-center ">
                <Bell className="w-12 h-12 mx-auto mb-4" />
                <h3 className="mb-2 text-lg font-medium">No notifications found</h3>
                <p>
                  {filterType === 'all' 
                    ? "You have no notifications yet." 
                    : `No ${filterType} notifications found.`
                  }
                </p>
              </div>
            ) : (
              paginatedNotifications.map((notification) => (
                <Card 
                  key={notification._id} 
                  className={`transition-all duration-200 hover:shadow-md ${
                    !notification.isRead 
                      ? "bg-blue-50 border-blue-200 shadow-sm" 
                      : "hover:bg-gray-50"
                  }`}
                >
                  <CardContent className="p-4">
                    <div className="flex items-start justify-between gap-4">
                      <div className="flex-1">
                        <div className="flex items-center gap-2 mb-2">
                          <h3 className="font-semibold text-gray-900">{notification.title}</h3>
                          {!notification.isRead && (
                            <Badge variant="secondary" className="text-blue-800 bg-blue-100 hover:bg-blue-100">
                              New
                            </Badge>
                          )}
                        </div>
                        <p className="text-sm ">{notification.message}</p>
                        {/* Action buttons based on notification type and user role */}
                        {(() => {
                          const isAdmin = auth.user?.role === 'admin' || auth.user?.role === 'super_admin';
                          
                          // Handle unfilled form notifications (new forms to fill)
                          if (notification.type === 'unfilled_form' && notification.relatedForm) {
                            const formId = typeof notification.relatedForm === 'object' ? notification.relatedForm._id : notification.relatedForm;
                            if (isAdmin) {
                              // Admin can manage the form
                              return (
                                <Button 
                                  variant="outline" 
                                  size="sm"
                                  className="mt-3 hover:bg-blue-50 hover:border-blue-300"
                                  onClick={() => handleNotificationAction(notification._id, `/admin/forms/${formId}`)}
                                >
                                  Manage Form
                                </Button>
                              );
                            } else {
                              // User can fill the form
                              return (
                                <Button 
                                  variant="outline" 
                                  size="sm"
                                  className="mt-3 hover:bg-green-50 hover:border-green-300"
                                  onClick={() => handleNotificationAction(notification._id, `/dashboard/forms/${formId}`)}
                                >
                                  Fill Form Now
                                </Button>
                              );
                            }
                          }
                          
                          // Handle form expiry notifications (both submission-based and form status-based)
                          if (['form_expiring', 'form_expiring_2_days', 'form_expiring_today', 'form_expired', 'form_status_expired', 'form_status_expiring_soon', 'user_form_expired', 'user_form_expiring_today'].includes(notification.type) && notification.relatedForm) {
                            const formId = typeof notification.relatedForm === 'object' ? notification.relatedForm._id : notification.relatedForm;
                            const isAdminAlert = notification.type === 'user_form_expired' || notification.type === 'user_form_expiring_today';
                            
                            if (isAdmin) {
                              // Admin can manage the form or view user submissions
                              if (isAdminAlert) {
                                return (
                                  <div className="flex gap-2 mt-3">
                                    <Button 
                                      variant="outline" 
                                      size="sm"
                                      className="hover:bg-red-50 hover:border-red-300"
                                      onClick={() => handleNotificationAction(notification._id, `/admin/forms/${formId}`)}
                                    >
                                      Manage Form
                                    </Button>
                                    <Button 
                                      variant="outline" 
                                      size="sm"
                                      className="hover:bg-blue-50 hover:border-blue-300"
                                      onClick={() => handleNotificationAction(notification._id, `/admin/submissions?form=${formId}`)}
                                    >
                                      View User Submissions
                                    </Button>
                                  </div>
                                );
                              } else {
                                return (
                                  <Button 
                                    variant="outline" 
                                    size="sm"
                                    className="mt-3 hover:bg-orange-50 hover:border-orange-300"
                                    onClick={() => handleNotificationAction(notification._id, `/admin/forms/${formId}`)}
                                  >
                                    Manage Form
                                  </Button>
                                );
                              }
                            } else {
                              // User can fill the form (including expired forms to refill them)
                              const isExpired = notification.type === 'form_expired' || notification.type === 'form_status_expired';
                              const isStatusBased = notification.type === 'form_status_expired' || notification.type === 'form_status_expiring_soon';
                              
                              return (
                                <Button 
                                  variant="outline"
                                  size="sm"
                                  className={`mt-3 ${isExpired 
                                    ? 'bg-red-50 border-red-200 text-red-700 hover:bg-red-100 hover:border-red-300' 
                                    : 'hover:bg-yellow-50 hover:border-yellow-300'
                                  }`}
                                  onClick={() => handleNotificationAction(notification._id, `/dashboard/forms/${formId}`)}
                                >
                                  {isExpired 
                                    ? (isStatusBased ? 'Refill Expired Form' : 'Refill Expired Form') 
                                    : (isStatusBased ? 'Review Form Status' : 'Fill Form Now')
                                  }
                                </Button>
                              );
                            }
                          }
                          
                          // Handle form submission notifications (for admins)
                          if (notification.type === 'form_submitted' && isAdmin && notification.relatedSubmission) {
                            const submissionId = typeof notification.relatedSubmission === 'object' ? notification.relatedSubmission._id : notification.relatedSubmission;
                            return (
                              <Button 
                                variant="outline"
                                size="sm"
                                className="mt-3 hover:bg-purple-50 hover:border-purple-300"
                                onClick={() => handleNotificationAction(notification._id, `/admin/submissions?submission=${submissionId}`)}
                              >
                                View Submission
                              </Button>
                            );
                          }
                          
                          // Handle form approved notifications (for users)
                          if (notification.type === 'form_approved' && !isAdmin && notification.relatedSubmission) {
                            const submissionId = typeof notification.relatedSubmission === 'object' ? notification.relatedSubmission._id : notification.relatedSubmission;
                            return (
                              <Button 
                                variant="outline"
                                size="sm"
                                className="mt-3 bg-green-50 border-green-200 text-green-700 hover:bg-green-100 hover:border-green-300"
                                onClick={() => handleNotificationAction(notification._id, `/dashboard/submissions/${submissionId}`)}
                              >
                                View Approved Submission
                              </Button>
                            );
                          }
                          
                          // Handle form rejected notifications (for users) - with refill button
                          if (notification.type === 'form_rejected' && !isAdmin && notification.relatedForm) {
                            const formId = typeof notification.relatedForm === 'object' ? notification.relatedForm._id : notification.relatedForm;
                            return (
                              <div className="flex gap-2 mt-3">
                                <Button 
                                  variant="outline"
                                  size="sm"
                                  className="bg-red-50 border-red-200 text-red-700 hover:bg-red-100 hover:border-red-300"
                                  onClick={() => handleNotificationAction(notification._id, `/dashboard/forms/${formId}`)}
                                >
                                  Refill Form
                                </Button>
                                {notification.relatedSubmission && (
                                  <Button 
                                    variant="outline"
                                    size="sm"
                                    className="hover:bg-gray-50 hover:border-gray-300"
                                    onClick={() => handleNotificationAction(notification._id, `/dashboard/submissions/${typeof notification.relatedSubmission === 'object' ? notification.relatedSubmission._id : notification.relatedSubmission}`)}
                                  >
                                    View Rejected Submission
                                  </Button>
                                )}
                              </div>
                            );
                          }
                          
                          return null;
                        })()}
                        <p className="mt-2 text-xs ">{formatDateTime(notification.createdAt)}</p>
                      </div>
                      <div className="flex items-center">
                        {!notification.isRead && (
                          <Button variant="ghost" size="icon" onClick={() => markAsRead(notification._id)} title="Mark as read">
                            <Check className="w-4 h-4" />
                          </Button>
                        )}
                        <Button variant="ghost" size="icon" onClick={() => deleteNotification(notification._id)} title="Delete notification">
                          <Trash2 className="w-4 h-4 text-destructive" />
                        </Button>
                      </div>
                    </div>
                  </CardContent>
                </Card>
              ))
            )}
          </div>
        </CardContent>
      </Card>

      {/* Pagination Controls */}
      {totalPages > 1 && (
        <Card>
          <CardContent className="p-4">
            <div className="flex flex-col gap-4 sm:flex-row sm:items-center sm:justify-between">
              <div className="text-sm ">
                Showing {((currentPage - 1) * itemsPerPage) + 1} to {Math.min(currentPage * itemsPerPage, filteredNotifications.length)} of {filteredNotifications.length} notifications
              </div>
              
              <div className="flex items-center gap-2">
                <Button
                  variant="outline"
                  size="sm"
                  onClick={() => setCurrentPage(Math.max(1, currentPage - 1))}
                  disabled={currentPage === 1}
                >
                  <ChevronLeft className="w-4 h-4" />
                  Previous
                </Button>
                
                <div className="flex items-center gap-1">
                  {Array.from({ length: Math.min(5, totalPages) }, (_, i) => {
                    let pageNum;
                    if (totalPages <= 5) {
                      pageNum = i + 1;
                    } else if (currentPage <= 3) {
                      pageNum = i + 1;
                    } else if (currentPage >= totalPages - 2) {
                      pageNum = totalPages - 4 + i;
                    } else {
                      pageNum = currentPage - 2 + i;
                    }
                    
                    return (
                      <Button
                        key={pageNum}
                        variant={currentPage === pageNum ? "default" : "outline"}
                        size="sm"
                        className="w-8 h-8"
                        onClick={() => setCurrentPage(pageNum)}
                      >
                        {pageNum}
                      </Button>
                    );
                  })}
                </div>
                
                <Button
                  variant="outline"
                  size="sm"
                  onClick={() => setCurrentPage(Math.min(totalPages, currentPage + 1))}
                  disabled={currentPage === totalPages}
                >
                  Next
                  <ChevronRight className="w-4 h-4" />
                </Button>
              </div>
            </div>
          </CardContent>
        </Card>
      )}

    </div>
  );
};

export default Notifications;