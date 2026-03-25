import axios, { AxiosInstance } from "axios";
import { AuthState } from "../types";

class ApiService {
  private api: AxiosInstance;
  private auth: AuthState = {
    user: null,
    token: null,
    refreshToken: null,
    isAuthenticated: false,
    loading: false,
  };

  constructor() {
    // Use environment variable or fallback to localhost for development
    const apiUrl = (import.meta.env.VITE_API_URL as string) || "http://localhost:8080/api";
    
    // console.log('API Base URL:', apiUrl); // Debug log to see which URL is being used
    
    this.api = axios.create({
      baseURL: apiUrl,
      timeout: 10000,
      headers: { "Content-Type": "application/json" },
    });

    this.api.interceptors.request.use(
      (config) => {
        if (this.auth.token) {
          config.headers.Authorization = `Bearer ${this.auth.token}`;
        }
        return config;
      },
      (error) => Promise.reject(error)
    );

    this.api.interceptors.response.use(
      (response) => response,
      async (error) => {
        const originalRequest = error.config;
        if (error.response?.status === 401 && !originalRequest._retry) {
          originalRequest._retry = true;
          try {
            const newToken = await this.refreshToken();
            if (newToken) {
              originalRequest.headers.Authorization = `Bearer ${newToken}`;
              return this.api(originalRequest);
            }
          } catch (refreshError) {
            this.logout();
            window.location.href = "/login";
          }
        }
        return Promise.reject(error);
      }
    );
  }

  setAuthState(auth: AuthState) { this.auth = auth; }

  async refreshToken(): Promise<string | null> {
    if (!this.auth.refreshToken) return null;
    const response = await this.api.post("/auth/refresh", { refreshToken: this.auth.refreshToken });
    const { token, refreshToken } = response.data.data;
    this.auth.token = token;
    this.auth.refreshToken = refreshToken;
    return token;
  }

  logout() {
    this.auth = { user: null, token: null, refreshToken: null, isAuthenticated: false, loading: false };
  }

  // Auth endpoints
  async login(email: string, password: string) { return this.api.post("/auth/login", { email, password }); }
  async register(userData: any) { return this.api.post("/auth/register", userData); }
  async getProfile() { return this.api.get("/auth/profile"); }
  async updateProfile(profileData: any) { return this.api.put("/auth/profile", profileData); }
  async changePassword(passwordData: any) { return this.api.put("/auth/change-password", passwordData); }
  async forgotPassword(email: string) { return this.api.post("/auth/forgot-password", { email }); }

  // Admin endpoints
  async getAdminDashboard() { return this.api.get("/admin/dashboard"); }
  async getUsers(params?: any) { return this.api.get("/admin/users", { params }); }
  async getUser(id: string) { return this.api.get(`/admin/users/${id}`); }
  async createUser(userData: any) { return this.api.post("/admin/users", userData); }
  async updateUser(id: string, userData: any) { return this.api.put(`/admin/users/${id}`, userData); }
  async deleteUser(id: string) { return this.api.delete(`/admin/users/${id}`); }
  async getShips(params?: any) { return this.api.get("/admin/ships", { params }); }
  async createShip(shipData: any) { return this.api.post("/admin/ships", shipData); }
  async updateShip(id: string, shipData: any) { return this.api.put(`/admin/ships/${id}`, shipData); }
  async deleteShip(id: string) { return this.api.delete(`/admin/ships/${id}`); }
  async getPendingSubmissions(params?: any) { return this.api.get("/admin/submissions/pending", { params }); }
  async getSubmissionStats(params?: any) { return this.api.get("/admin/submissions/stats", { params }); }

  // Form endpoints
  async getCategories(params?: any) { return this.api.get("/forms/categories", { params }); }
  async createCategory(categoryData: any) { return this.api.post("/forms/categories", categoryData); }
  async updateCategory(id: string, categoryData: any) { return this.api.put(`/forms/categories/${id}`, categoryData); }
  async deleteCategory(id: string) { return this.api.delete(`/forms/categories/${id}`); }
  async getForms(params?: any) { return this.api.get("/forms", { params }); }
  async getFormsWithUserStatus(params?: any) { return this.api.get("/forms/user-status", { params }); }
  async getFormsByCategory(categoryName: string, params?: any) { return this.api.get(`/forms/category/${categoryName}`, { params }); }
  async getFormById(id: string) { return this.api.get(`/forms/${id}`); }
  async createForm(formData: any) { return this.api.post("/forms", formData); }
  async updateForm(id: string, formData: any) { return this.api.put(`/forms/${id}`, formData); }
  async getFormStats(id: string) { return this.api.get(`/forms/${id}/stats`); }
  async deleteForm(id: string) { return this.api.delete(`/forms/${id}`); }
  async toggleFormStatus(id: string) { return this.api.patch(`/forms/${id}/toggle-status`); }
  async duplicateForm(id: string, duplicateData: any) { return this.api.post(`/forms/${id}/duplicate`, duplicateData); }

  // Submission endpoints
  async getSubmissions(params?: any) { return this.api.get("/submissions", { params }); }
  async getSubmissionById(id: string) { return this.api.get(`/submissions/${id}`); }
  async createSubmission(submissionData: any) { return this.api.post("/submissions", submissionData); }
  async updateSubmission(id: string, submissionData: any) { return this.api.put(`/submissions/${id}`, submissionData); }
  async deleteSubmission(id: string) { return this.api.delete(`/submissions/${id}`); }
  async reviewSubmission(id: string, reviewData: any) { return this.api.put(`/submissions/${id}/review`, reviewData); }
  async getSubmissionDashboard() { return this.api.get("/submissions/dashboard"); }
  async getStatusSummary() { return this.api.get("/submissions/status-summary"); }

  // Notification endpoints
  async getNotifications(params?: any) { return this.api.get("/notifications", { params }); }
  async markNotificationAsRead(id: string) { return this.api.put(`/notifications/${id}/read`); }
  async markAllNotificationsAsRead() { return this.api.put("/notifications/read-all"); }
  async deleteNotification(id: string) { return this.api.delete(`/notifications/${id}`); }
  async getNotificationStats() { return this.api.get("/notifications/stats"); }
  async createSystemNotification(notificationData: any) { return this.api.post("/notifications/system", notificationData); }
  async checkUnfilledForms() { return this.api.post("/notifications/check-unfilled"); }
  async triggerFormStatusCheck() { return this.api.post("/notifications/trigger-form-status-check"); }
}

export default new ApiService();