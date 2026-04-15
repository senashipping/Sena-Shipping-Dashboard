import React from "react";
import { BrowserRouter, Routes, Route } from "react-router-dom";
import { AuthProvider } from "./contexts/AuthContext";
import { NotificationProvider } from "./contexts/NotificationContext";
import { ThemeProvider } from "./contexts/ThemeContext";
import { ToastProvider } from "./components/ui/toast";
import ProtectedRoute from "./components/ProtectedRoute";
import RoleBasedRedirect from "./components/RoleBasedRedirect";
const AdminLayout = React.lazy(() => import("./layouts/AdminLayout"));
const UserLayout = React.lazy(() => import("./layouts/UserLayout"));
const Login = React.lazy(() => import("./pages/auth/Login"));
const Register = React.lazy(() => import("./pages/auth/Register"));
const AdminDashboard = React.lazy(() => import("./pages/admin/Dashboard"));
const UserDashboard = React.lazy(() => import("./pages/user/Dashboard"));
const FormList = React.lazy(() => import("./pages/forms/FormList"));
const FormBuilder = React.lazy(() => import("./pages/forms/FormBuilder"));
const CategoryForms = React.lazy(() => import("./pages/forms/CategoryForms"));
const FormSubmission = React.lazy(() => import("./pages/submissions/FormSubmission"));
const UserProfile = React.lazy(() => import("./pages/user/Profile"));
const SubmissionView = React.lazy(() => import("./pages/user/SubmissionView"));
const AdminUsers = React.lazy(() => import("./pages/admin/Users"));
const AdminShips = React.lazy(() => import("./pages/admin/Ships"));
const AdminSubmissions = React.lazy(() => import("./pages/admin/Submissions"));
const UserSubmissions = React.lazy(() => import("./pages/admin/UserSubmissions"));
const Notifications = React.lazy(() => import("./pages/notifications/Notifications"));
const NotFound = React.lazy(() => import("./pages/NotFound"));
const Unauthorized = React.lazy(() => import("./pages/Unauthorized"));

function App() {
  return (
    <BrowserRouter>
      <ThemeProvider>
        <AuthProvider>
          <NotificationProvider>
            <ToastProvider>
              <React.Suspense
                fallback={
                  <div className="flex items-center justify-center min-h-[40vh]">
                    Loading...
                  </div>
                }
              >
                <Routes>
                  {/* Root route that redirects based on role */}
                  <Route path="/" element={<RoleBasedRedirect />} />

                  <Route path="/login" element={<Login />} />
                  <Route path="/register" element={<Register />} />
                  <Route path="/unauthorized" element={<Unauthorized />} />

                  {/* Admin routes */}
                  <Route path="/admin" element={<ProtectedRoute requiredRole="admin"><AdminLayout /></ProtectedRoute>}>
                    <Route index element={<AdminDashboard />} />
                    <Route path="dashboard" element={<AdminDashboard />} />
                    <Route path="users" element={<AdminUsers />} />
                    <Route path="users/:userId/submissions" element={<UserSubmissions />} />
                    <Route path="ships" element={<AdminShips />} />
                    <Route path="forms" element={<FormList />} />
                    <Route path="forms/new" element={<FormBuilder />} />
                    <Route path="forms/:id" element={<FormBuilder />} />
                    <Route path="submissions" element={<AdminSubmissions />} />
                    <Route path="notifications" element={<Notifications />} />
                  </Route>

                  {/* User routes */}
                  <Route path="/dashboard" element={<ProtectedRoute requiredRole="user"><UserLayout /></ProtectedRoute>}>
                    <Route index element={<UserDashboard />} />
                    <Route path="forms" element={<FormList />} />
                    <Route path="forms/category/:category" element={<CategoryForms />} />
                    <Route path="forms/:id" element={<FormSubmission />} />
                    <Route path="submissions/:submissionId" element={<SubmissionView />} />
                    <Route path="profile" element={<UserProfile />} />
                    <Route path="notifications" element={<Notifications />} />
                  </Route>

                  <Route path="*" element={<NotFound />} />
                </Routes>
              </React.Suspense>
            </ToastProvider>
          </NotificationProvider>
        </AuthProvider>
      </ThemeProvider>
    </BrowserRouter>
  );
}

export default App;