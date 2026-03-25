import React, { useState } from "react";
import { useQuery } from "@tanstack/react-query";
import api from "../../api";
import { Button } from "../../components/ui/button";
import { Card, CardContent, CardDescription, CardHeader, CardTitle } from "../../components/ui/card";
import { Input } from "../../components/ui/input";
import { Label } from "../../components/ui/label";
import { Alert, AlertDescription } from "../../components/ui/alert";
import { useAuth } from "../../contexts/AuthContext";
import { User, Mail, Building } from "lucide-react";

const UserProfile: React.FC = () => {
  const { auth, updateProfile, changePassword } = useAuth();
  const [profileForm, setProfileForm] = useState({
    name: auth.user?.name || "",
    email: auth.user?.email || "",
  });
  const [passwordForm, setPasswordForm] = useState({
    currentPassword: "",
    newPassword: "",
    confirmPassword: "",
  });
  const [profileError, setProfileError] = useState("");
  const [passwordError, setPasswordError] = useState("");
  const [profileSuccess, setProfileSuccess] = useState("");
  const [passwordSuccess, setPasswordSuccess] = useState("");
  const [isUpdatingProfile, setIsUpdatingProfile] = useState(false);
  const [isChangingPassword, setIsChangingPassword] = useState(false);

  useQuery({
    queryKey: ["userProfile"],
    queryFn: () => api.getProfile(),
    enabled: !!auth.user,
  });

  const handleProfileSubmit = async (e: React.FormEvent) => {
    e.preventDefault();
    setProfileError("");
    setProfileSuccess("");
    setIsUpdatingProfile(true);

    try {
      await updateProfile(profileForm);
      setProfileSuccess("Profile updated successfully");
    } catch (err: any) {
      setProfileError(err.response?.data?.message || "Failed to update profile");
    } finally {
      setIsUpdatingProfile(false);
    }
  };

  const handlePasswordSubmit = async (e: React.FormEvent) => {
    e.preventDefault();
    
    if (passwordForm.newPassword !== passwordForm.confirmPassword) {
      setPasswordError("New passwords do not match");
      return;
    }
    
    setPasswordError("");
    setPasswordSuccess("");
    setIsChangingPassword(true);

    try {
      await changePassword({
        currentPassword: passwordForm.currentPassword,
        newPassword: passwordForm.newPassword,
      });
      setPasswordSuccess("Password changed successfully");
      setPasswordForm({
        currentPassword: "",
        newPassword: "",
        confirmPassword: "",
      });
    } catch (err: any) {
      setPasswordError(err.response?.data?.message || "Failed to change password");
    } finally {
      setIsChangingPassword(false);
    }
  };

  return (
    <div className="space-y-6">
      <h1 className="text-3xl font-bold">User Profile</h1>

      <div className="grid grid-cols-1 md:grid-cols-2 gap-6">
        {/* Profile Information */}
        <Card>
          <CardHeader>
            <CardTitle className="flex items-center">
              <User className="mr-2 h-5 w-5" />
              Profile Information
            </CardTitle>
            <CardDescription>Update your personal information</CardDescription>
          </CardHeader>
          <CardContent>
            {profileError && (
              <Alert variant="destructive" className="mb-4">
                <AlertDescription>{profileError}</AlertDescription>
              </Alert>
            )}
            
            {profileSuccess && (
              <Alert className="mb-4">
                <AlertDescription>{profileSuccess}</AlertDescription>
              </Alert>
            )}
            
            <form onSubmit={handleProfileSubmit} className="space-y-4">
              <div className="space-y-2">
                <Label htmlFor="name">Full Name</Label>
                <Input
                  id="name"
                  value={profileForm.name}
                  onChange={(e) => setProfileForm({...profileForm, name: e.target.value})}
                  required
                />
              </div>
              
              <div className="space-y-2">
                <Label htmlFor="email">Email</Label>
                <Input
                  id="email"
                  type="email"
                  value={profileForm.email}
                  onChange={(e) => setProfileForm({...profileForm, email: e.target.value})}
                  required
                />
              </div>
              
              <div className="space-y-2">
                <Label>Role</Label>
                <Input
                  value={auth.user?.role ? auth.user.role.charAt(0).toUpperCase() + auth.user.role.slice(1) : ""}
                  disabled
                />
              </div>
              
              {auth.user?.ship && (
                <div className="space-y-2">
                  <Label className="flex items-center">
                    <Building className="mr-2 h-4 w-4" />
                    Assigned Ship
                  </Label>
                  <Input
                    value={auth.user.ship.name}
                    disabled
                  />
                </div>
              )}
              
              <Button type="submit" disabled={isUpdatingProfile}>
                {isUpdatingProfile ? "Updating..." : "Update Profile"}
              </Button>
            </form>
          </CardContent>
        </Card>

        {/* Change Password */}
        <Card>
          <CardHeader>
            <CardTitle className="flex items-center">
              <Mail className="mr-2 h-5 w-5" />
              Change Password
            </CardTitle>
            <CardDescription>Update your account password</CardDescription>
          </CardHeader>
          <CardContent>
            {passwordError && (
              <Alert variant="destructive" className="mb-4">
                <AlertDescription>{passwordError}</AlertDescription>
              </Alert>
            )}
            
            {passwordSuccess && (
              <Alert className="mb-4">
                <AlertDescription>{passwordSuccess}</AlertDescription>
              </Alert>
            )}
            
            <form onSubmit={handlePasswordSubmit} className="space-y-4">
              <div className="space-y-2">
                <Label htmlFor="currentPassword">Current Password</Label>
                <Input
                  id="currentPassword"
                  type="password"
                  value={passwordForm.currentPassword}
                  onChange={(e) => setPasswordForm({...passwordForm, currentPassword: e.target.value})}
                  required
                />
              </div>
              
              <div className="space-y-2">
                <Label htmlFor="newPassword">New Password</Label>
                <Input
                  id="newPassword"
                  type="password"
                  value={passwordForm.newPassword}
                  onChange={(e) => setPasswordForm({...passwordForm, newPassword: e.target.value})}
                  required
                  minLength={6}
                />
              </div>
              
              <div className="space-y-2">
                <Label htmlFor="confirmPassword">Confirm New Password</Label>
                <Input
                  id="confirmPassword"
                  type="password"
                  value={passwordForm.confirmPassword}
                  onChange={(e) => setPasswordForm({...passwordForm, confirmPassword: e.target.value})}
                  required
                />
              </div>
              
              <Button type="submit" disabled={isChangingPassword}>
                {isChangingPassword ? "Changing..." : "Change Password"}
              </Button>
            </form>
          </CardContent>
        </Card>
      </div>
    </div>
  );
};

export default UserProfile;