"use client";

import React from "react";
import { useAuth } from "../../contexts/AuthContext";
import { useToast } from "../../components/ui/toast";
import { getApiErrorMessage } from "../../lib/utils";
import { Button } from "../../components/ui/button";
import { Input } from "../../components/ui/input";
import { Card, CardContent, CardDescription, CardHeader, CardTitle } from "../../components/ui/card";
import { Label } from "../../components/ui/label";
import { Alert, AlertDescription } from "../../components/ui/alert";
import { AlertCircle } from "lucide-react";

const Register: React.FC = () => {
  const [formData, setFormData] = React.useState({
    name: "",
    email: "",
    password: "",
    confirmPassword: "",
  });
  const [error, setError] = React.useState("");
  const [loading, setLoading] = React.useState(false);
  
  const { register } = useAuth();
  const { toast } = useToast();

  const handleSubmit = async (e: React.FormEvent) => {
    e.preventDefault();
    setError("");

    if (formData.password !== formData.confirmPassword) {
      const msg = "Passwords do not match";
      setError(msg);
      toast({ title: "Check your password", description: msg, variant: "destructive" });
      return;
    }

    setLoading(true);

    try {
      await register({
        name: formData.name,
        email: formData.email,
        password: formData.password,
        role: "user",
      });
      // Navigation + success toast are handled in AuthContext.register
    } catch (err: unknown) {
      const message = getApiErrorMessage(err, "Registration failed");
      setError(message);
      toast({ title: "Registration failed", description: message, variant: "destructive" });
    } finally {
      setLoading(false);
    }
  };

  const handleChange = (e: React.ChangeEvent<HTMLInputElement>) => {
    setFormData({
      ...formData,
      [e.target.name]: e.target.value,
    });
  };

  return (
    <div className="flex items-center justify-center min-h-screen px-4 py-6 sm:py-12 bg-gradient-to-br from-white via-sena-skyBlue/5 to-sena-lightBlue/10 dark:from-sena-darkBg dark:via-sena-darkCard/30 dark:to-sena-navy/20">
      <div className="w-full max-w-md">
        {/* Logo Header */}
        <div className="mb-6 text-center sm:mb-8">
          <img 
            src="/sena_logo.png" 
            alt="Logo" 
            className="w-auto h-20 mx-auto mb-3 sm:h-24 sm:mb-4 dark:brightness-0 dark:invert"
          />
        </div>

        <Card className="w-full border-0 shadow-xl bg-white/95 dark:bg-sena-darkCard/95 backdrop-blur-sm">
          <CardHeader className="px-4 pt-6 pb-4 space-y-1 sm:px-6">
            <CardTitle className="text-xl font-bold text-center sm:text-2xl text-sena-navy dark:text-white">
              Create Account
            </CardTitle>
            <CardDescription className="text-center text-sena-lightBlue dark:text-white/90">
              Sign up for a new Sena Shipping account
            </CardDescription>
          </CardHeader>
          <CardContent className="px-4 pb-6 sm:px-6">
            {error && (
              <Alert variant="destructive" className="mb-4 border-red-200 dark:border-red-800">
                <AlertCircle className="w-4 h-4" />
                <AlertDescription>{error}</AlertDescription>
              </Alert>
            )}
            
            <form onSubmit={handleSubmit} className="space-y-4">
              <div className="space-y-2">
                <Label htmlFor="name" className="text-sena-navy dark:text-white">Full Name</Label>
                <Input
                  id="name"
                  name="name"
                  placeholder="Enter your full name"
                  value={formData.name}
                  onChange={handleChange}
                  required
                  className="h-10 border-sena-lightBlue/30 focus:border-sena-gold focus:ring-sena-gold/20 sm:h-11"
                />
              </div>
              
              <div className="space-y-2">
                <Label htmlFor="email" className="text-sena-navy dark:text-white">Email</Label>
                <Input
                  id="email"
                  name="email"
                  type="email"
                  placeholder="Enter your email"
                  value={formData.email}
                  onChange={handleChange}
                  required
                  className="h-10 border-sena-lightBlue/30 focus:border-sena-gold focus:ring-sena-gold/20 sm:h-11"
                />
              </div>
              
              <div className="space-y-2">
                <Label htmlFor="password" className="text-sena-navy dark:text-white">Password</Label>
                <Input
                  id="password"
                  name="password"
                  type="password"
                  placeholder="Enter your password"
                  value={formData.password}
                  onChange={handleChange}
                  required
                  className="h-10 border-sena-lightBlue/30 focus:border-sena-gold focus:ring-sena-gold/20 sm:h-11"
                />
              </div>
              
              <div className="space-y-2">
                <Label htmlFor="confirmPassword" className="text-sena-navy dark:text-white">Confirm Password</Label>
                <Input
                  id="confirmPassword"
                  name="confirmPassword"
                  type="password"
                  placeholder="Confirm your password"
                  value={formData.confirmPassword}
                  onChange={handleChange}
                  required
                  className="h-10 border-sena-lightBlue/30 focus:border-sena-gold focus:ring-sena-gold/20 sm:h-11"
                />
              </div>
              
              <Button 
                type="submit" 
                className="w-full py-2.5 sm:py-3 font-semibold text-white transition-all duration-200 rounded-lg shadow-lg bg-sena-navy hover:bg-sena-darkNavy hover:shadow-xl h-10 sm:h-11" 
                disabled={loading}
              >
                {loading ? "Creating account..." : "Create Account"}
              </Button>
            </form>
            
            <div className="mt-4 text-center">
              <p className="text-sm text-sena-lightBlue dark:text-white/90">
                Join our maritime management platform
              </p>
            </div>
          </CardContent>
        </Card>
      </div>
    </div>
  );
};

export default Register;