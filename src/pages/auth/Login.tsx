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

const Login: React.FC = () => {
  const [email, setEmail] = React.useState("");
  const [password, setPassword] = React.useState("");
  const [error, setError] = React.useState("");
  const [loading, setLoading] = React.useState(false);
  
  const { login } = useAuth();
  const { toast } = useToast();

  const handleSubmit = async (e: React.FormEvent) => {
    e.preventDefault();
    setError("");
    setLoading(true);

    try {
      await login(email, password);
      // Success toast + navigation are handled in AuthContext
    } catch (err: unknown) {
      const message = getApiErrorMessage(err, "Login failed");
      setError(message);
      toast({ title: "Sign in failed", description: message, variant: "destructive" });
    } finally {
      setLoading(false);
    }
  };

  return (
    <div className="flex items-center justify-center min-h-screen px-4 py-6 sm:py-12 bg-gradient-to-br from-white via-sena-skyBlue/5 to-sena-lightBlue/10 dark:from-sena-darkBg dark:via-sena-darkCard/30 dark:to-sena-navy/20">
      <div className="w-full max-w-md">
        {/* Logo Header */}
        <div className="mb-6 sm:mb-8 text-center">
          <img 
            src="/sena_logo.png" 
            alt="Logo" 
            className="w-auto h-20 sm:h-24 mx-auto mb-3 sm:mb-4 dark:brightness-0 dark:invert"
          />
        </div>

        <Card className="w-full border-0 shadow-xl bg-white/95 dark:bg-sena-darkCard/95 backdrop-blur-sm">
          <CardHeader className="pb-4 space-y-1 px-4 sm:px-6">
            <CardTitle className="text-xl sm:text-2xl font-bold text-center text-sena-navy dark:text-white">
              Access Your Dashboard
            </CardTitle>
            <CardDescription className="text-center text-sena-lightBlue dark:text-white/90">
              Enter your credentials to access your dashboard
            </CardDescription>
          </CardHeader>
        <CardContent className="px-4 sm:px-6 pb-6">
          {error && (
            <Alert variant="destructive" className="mb-4">
              <AlertCircle className="w-4 h-4" />
              <AlertDescription>{error}</AlertDescription>
            </Alert>
          )}
          
          <form onSubmit={handleSubmit} className="space-y-4">
            <div className="space-y-2">
              <Label htmlFor="email" className="text-sena-navy dark:text-white">Email</Label>
              <Input
                id="email"
                type="email"
                placeholder="Enter your email"
                value={email}
                onChange={(e) => setEmail(e.target.value)}
                required
                className="border-sena-lightBlue/30 focus:border-sena-gold focus:ring-sena-gold/20 h-10 sm:h-11"
              />
            </div>
            
            <div className="space-y-2">
              <Label htmlFor="password" className="text-sena-navy dark:text-white">Password</Label>
              <Input
                id="password"
                type="password"
                placeholder="Enter your password"
                value={password}
                onChange={(e) => setPassword(e.target.value)}
                required
                className="border-sena-lightBlue/30 focus:border-sena-gold focus:ring-sena-gold/20 h-10 sm:h-11"
              />
            </div>
            
            <Button 
              type="submit" 
              className="w-full py-2.5 sm:py-3 font-semibold text-white transition-all duration-200 rounded-lg shadow-lg bg-sena-navy hover:bg-sena-darkNavy hover:shadow-xl h-10 sm:h-11" 
              disabled={loading}
            >
              {loading ? "Signing in..." : "Sign in to Dashboard"}
            </Button>
          </form>
          
          <div className="mt-4 text-center">
            <p className="text-sm text-sena-lightBlue dark:text-white/90">
              Professional maritime management at your fingertips
            </p>
          </div>
        </CardContent>
      </Card>
      </div>
    </div>
  );
};

export default Login;