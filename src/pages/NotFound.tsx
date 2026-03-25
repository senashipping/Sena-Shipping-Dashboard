"use client";

import React from "react";
import { Link } from "react-router-dom";
import { Button } from "../components/ui/button";
import { Card, CardContent, CardDescription, CardHeader, CardTitle } from "../components/ui/card";
import { AlertTriangle } from "lucide-react";

const NotFound: React.FC = () => {
  return (
    <div className="flex items-center justify-center min-h-screen bg-background">
      <Card className="w-full max-w-md">
        <CardHeader className="text-center">
          <div className="flex items-center justify-center w-16 h-16 mx-auto rounded-full bg-destructive/10">
            <AlertTriangle className="w-8 h-8 text-destructive" />
          </div>
          <CardTitle className="mt-4 text-2xl">Page Not Found</CardTitle>
          <CardDescription>
            The page you're looking for doesn't exist or has been moved.
          </CardDescription>
        </CardHeader>
        <CardContent className="text-center">
          <p className="mb-6">
            Sorry, we couldn't find the page you were looking for. Please check the URL or return to the dashboard.
          </p>
          <Link to="/">
            <Button>Go to Dashboard</Button>
          </Link>
        </CardContent>
      </Card>
    </div>
  );
};

export default NotFound;