import React from "react";
import { Link } from "react-router-dom";
import { Button } from "../components/ui/button";
import { Card, CardContent, CardHeader, CardTitle } from "../components/ui/card";
import { Shield, ArrowLeft } from "lucide-react";

const Unauthorized: React.FC = () => {
  return (
    <div className="min-h-screen flex items-center justify-center bg-gradient-to-br from-white/30 to-sena-lightBlue/10">
      <Card className="w-full max-w-md mx-4">
        <CardHeader className="text-center">
          <div className="flex justify-center mb-4">
            <div className="p-3 rounded-full bg-red-100">
              <Shield className="w-8 h-8 text-red-600" />
            </div>
          </div>
          <CardTitle className="text-2xl text-sena-navy">Access Denied</CardTitle>
        </CardHeader>
        <CardContent className="text-center space-y-4">
          <p className="text-gray-600">
            You don't have permission to access this section. 
            Please contact your administrator if you believe this is an error.
          </p>
          <div className="space-y-2">
            <Link to="/dashboard" className="block">
              <Button className="w-full bg-sena-gold hover:bg-sena-gold/90">
                <ArrowLeft className="w-4 h-4 mr-2" />
                Go to Dashboard
              </Button>
            </Link>
          </div>
        </CardContent>
      </Card>
    </div>
  );
};

export default Unauthorized;