import { BrowserRouter, Routes, Route } from "react-router-dom";
import { AuthProvider } from "./contexts/AuthContext";
import { NotificationProvider } from "./contexts/NotificationContext";
import { ThemeProvider } from "./contexts/ThemeContext";
import { ToastProvider } from "./components/ui/toast";
import ProtectedRoute from "./components/ProtectedRoute";
import RoleBasedRedirect from "./components/RoleBasedRedirect";
import AdminLayout from "./layouts/AdminLayout";
import UserLayout from "./layouts/UserLayout";
import Login from "./pages/auth/Login";
import Register from "./pages/auth/Register";
import AdminDashboard from "./pages/admin/Dashboard";
import UserDashboard from "./pages/user/Dashboard";
import FormList from "./pages/forms/FormList";
import FormBuilder from "./pages/forms/FormBuilder";
import CategoryForms from "./pages/forms/CategoryForms";
import FormSubmission from "./pages/submissions/FormSubmission";
import UserProfile from "./pages/user/Profile";
import SubmissionView from "./pages/user/SubmissionView";
import AdminUsers from "./pages/admin/Users";
import AdminShips from "./pages/admin/Ships";
import AdminSubmissions from "./pages/admin/Submissions";
import UserSubmissions from "./pages/admin/UserSubmissions";
import Notifications from "./pages/notifications/Notifications";
import NotFound from "./pages/NotFound";
import Unauthorized from "./pages/Unauthorized";

function App() {
  return (
    <BrowserRouter>
      <ThemeProvider>
        <AuthProvider>
          <NotificationProvider>
            <ToastProvider>
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
            </ToastProvider>
          </NotificationProvider>
        </AuthProvider>
      </ThemeProvider>
    </BrowserRouter>
  );
}

export default App;