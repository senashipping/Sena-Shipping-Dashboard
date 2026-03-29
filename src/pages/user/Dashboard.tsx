import React from "react";
import { useQuery } from "@tanstack/react-query";
import { Link } from "react-router-dom";
import api from "../../api";
import { Card, CardContent, CardHeader, CardTitle } from "../../components/ui/card";
import { Button } from "../../components/ui/button";
import { Badge } from "../../components/ui/badge";
import { AlertTriangle, CheckCircle, FileText, Calendar, ArrowRight, Clock, XCircle } from "lucide-react";
import { getExpirationStatus, formatDate } from "../../lib/utils";
import { Table, TableBody, TableCell, TableHead, TableHeader, TableRow } from "../../components/ui/table";
import { useAuth } from "../../contexts/AuthContext";

const UserDashboard: React.FC = () => {
  const { canAccessCategory } = useAuth();

  // Helper function to get status badge
  const getStatusBadge = (status: string) => {
    switch (status) {
      case 'pending':
        return (
          <Badge variant="secondary" className="flex items-center gap-1">
            <Clock className="w-3 h-3" />
            Pending
          </Badge>
        );
      case 'approved':
        return (
          <Badge variant="default" className="flex items-center gap-1 text-green-800 bg-green-100 hover:bg-green-100">
            <CheckCircle className="w-3 h-3" />
            Approved
          </Badge>
        );
      case 'rejected':
        return (
          <Badge variant="destructive" className="flex items-center gap-1">
            <XCircle className="w-3 h-3" />
            Rejected
          </Badge>
        );
      default:
        return (
          <Badge variant="outline">
            {status}
          </Badge>
        );
    }
  };
  
  const { data: dashboardData, isLoading } = useQuery({
    queryKey: ["userDashboard"],
    queryFn: () => api.getSubmissionDashboard(),
  });

  // Get all active forms with user-specific status for upcoming forms table
  const { data: formsData } = useQuery({
    queryKey: ["activeForms"],
    queryFn: () => api.getFormsWithUserStatus({ isActive: true }),
  });



  const stats = dashboardData?.data?.data?.stats || {};
  const expiring = dashboardData?.data?.data?.expiring || [];
  const recent = dashboardData?.data?.data?.recent || [];

  // Process upcoming forms from dashboard data
  const upcomingForms = React.useMemo(() => {
    if (!formsData?.data?.data) return [];

    // Filter forms based on user access to categories
    const allForms = formsData.data.data;
    const forms = allForms.filter((form: any) => 
      canAccessCategory(form.category.name)
    );
    const submissions = dashboardData?.data?.data?.recent || [];
    const expiring = dashboardData?.data?.data?.expiring || [];
    const allSubmissions = [...submissions, ...expiring];

    const formsWithStatus = forms.map((form: any) => {
      // Find latest submission for this form (all are auto-approved)
      const formSubmissions = allSubmissions.filter((sub: any) => 
        sub.form && sub.form._id === form._id
      );
      
      let latestSubmission = null;
      if (formSubmissions.length > 0) {
        // Sort by submission date, most recent first
        latestSubmission = formSubmissions.sort((a: any, b: any) => {
          const dateA = new Date(a.submittedAt || a.createdAt);
          const dateB = new Date(b.submittedAt || b.createdAt);
          return dateB.getTime() - dateA.getTime();
        })[0];
      }
      
      // Determine form lifecycle state
      let formLifecycleState: 'unsubmitted' | 'active' | 'expired';
      let displayDate: Date;
      let displayText: string;
      let colorClass: string;
      
      if (!latestSubmission) {
        // No submissions at all - unsubmitted
        formLifecycleState = 'unsubmitted';
        displayDate = new Date();
        displayText = "Not submitted";
        colorClass = "text-red-600 bg-red-100";
      } else {
        // Has submission - check if expired
        const submissionDate = new Date(latestSubmission.submittedAt || latestSubmission.createdAt);
        const validityPeriod = form.validityPeriod || 30;
        const expiryDate = new Date(submissionDate);
        expiryDate.setDate(expiryDate.getDate() + validityPeriod);
        
        // Set time to start of day for accurate date comparison
        const today = new Date();
        today.setHours(0, 0, 0, 0);
        const expiryDateOnly = new Date(expiryDate);
        expiryDateOnly.setHours(0, 0, 0, 0);
        
        // Calculate difference in days using Math.floor for more accurate results
        const daysUntilExpiry = Math.floor((expiryDateOnly.getTime() - today.getTime()) / (1000 * 60 * 60 * 24));
        
        if (daysUntilExpiry < 0) {
          // Expired (past the expiry date)
          formLifecycleState = 'expired';
          displayDate = expiryDate;
          displayText = `Expired ${Math.abs(daysUntilExpiry)} day${Math.abs(daysUntilExpiry) === 1 ? '' : 's'} ago`;
          colorClass = "text-red-600 bg-red-100";
        } else if (daysUntilExpiry === 0) {
          // Expires today
          formLifecycleState = 'active';
          displayDate = expiryDate;
          displayText = "Expires Today";
          colorClass = "text-red-600 bg-red-100"; // Critical - expires today
        } else {
          // Active - future expiry
          formLifecycleState = 'active';
          displayDate = expiryDate;
          displayText = `Expires in ${daysUntilExpiry} day${daysUntilExpiry === 1 ? '' : 's'}`;
          
          // Color based on days remaining
          if (daysUntilExpiry <= 7) {
            colorClass = "text-red-600 bg-red-100"; // Critical
          } else if (daysUntilExpiry <= 15) {
            colorClass = "text-yellow-600 bg-yellow-100"; // Warning
          } else {
            colorClass = "text-green-600 bg-green-100"; // Good
          }
        }
      }
      
      return {
        ...form,
        formLifecycleState,
        displayDate,
        displayText,
        colorClass,
        latestSubmission,
        hasSubmission: !!latestSubmission
      };
    });

    // Sort by priority: unsubmitted/expired first, then active by expiry date
    return formsWithStatus
      .sort((a: any, b: any) => {
        // Priority order: unsubmitted -> expired -> active
        const priority: Record<string, number> = { 'unsubmitted': 1, 'expired': 2, 'active': 3 };
        const aPriority = priority[a.formLifecycleState] || 5;
        const bPriority = priority[b.formLifecycleState] || 5;
        
        if (aPriority !== bPriority) {
          return aPriority - bPriority;
        }
        
        // Within same priority, sort by display date (earliest first for active forms)
        return a.displayDate.getTime() - b.displayDate.getTime();
      })
      .slice(0, 10);
  }, [formsData?.data?.data, dashboardData?.data?.data, canAccessCategory]);

  const statCards = [
    {
      title: "Total Submissions",
      value: stats.total || 0,
      icon: FileText,
      color: "text-sena-navy",
      bgColor: "bg-sena-lightBlue/20",
      borderColor: "border-sena-lightBlue/30",
    },
    {
      title: "Pending Review", 
      value: stats.pending || 0,
      icon: Calendar,
      color: "text-yellow-700",
      bgColor: "bg-yellow-50",
      borderColor: "border-yellow-200",
    },
    {
      title: "Approved",
      value: stats.approved || 0,
      icon: CheckCircle,
      color: "text-green-700",
      bgColor: "bg-green-50",
      borderColor: "border-green-200",
    },
    {
      title: "Rejected",
      value: stats.rejected || 0,
      icon: AlertTriangle,
      color: "text-red-700",
      bgColor: "bg-red-50",
      borderColor: "border-red-200",
    },
  ];

  if (isLoading) {
    return (
      <div className="flex items-center justify-center h-64">
        <div className="w-12 h-12 border-b-2 rounded-full animate-spin border-primary"></div>
      </div>
    );
  }

  return (
    <div className="p-4 space-y-4 sm:space-y-6 bg-gradient-to-br from-white/30 to-sena-lightBlue/10 dark:from-sena-darkBg/30 dark:to-sena-darkCard/10 sm:p-6 rounded-xl">
      <div className="flex flex-col gap-4 sm:flex-row sm:items-center sm:justify-between">
        <div className="flex items-center space-x-3">
          <div className="flex-shrink-0 p-2 rounded-full sm:p-3 bg-sena-navy/10 dark:bg-sena-gold/10">
            <FileText className="w-6 h-6 sm:h-8 sm:w-8 text-sena-navy dark:text-white" />
          </div>
          <div className="min-w-0">
            <h1 className="text-2xl font-bold sm:text-3xl text-sena-navy dark:text-white">My Dashboard</h1>
            <p className="text-sm sm:text-base text-sena-lightBlue dark:text-white/90">Manage your maritime submissions</p>
          </div>
        </div>
        <Link to="/dashboard/forms" className="w-full sm:w-auto">
          <Button variant="gold" size="sm" className="w-full">
            <span className="sm:hidden">Forms</span>
            <span className="hidden sm:inline">View All Forms</span>
            <ArrowRight className="w-4 h-4 ml-2" />
          </Button>
        </Link>
      </div>

      {/* Statistics Grid */}
      <div className="grid grid-cols-1 gap-4 sm:gap-6 sm:grid-cols-2 lg:grid-cols-4">
        {statCards.map((card, index) => {
          const IconComponent = card.icon;
          return (
            <Card key={index} className={`border-2 ${card.borderColor} hover:shadow-xl transition-all duration-200 hover:scale-105`}>
              <CardHeader className="flex flex-row items-center justify-between pb-2 space-y-0">
                <CardTitle className="text-sm font-medium text-sena-navy">
                  {card.title}
                </CardTitle>
                <div className={`p-3 rounded-full ${card.bgColor} shadow-sm`}>
                  <IconComponent className={`h-5 w-5 ${card.color}`} />
                </div>
              </CardHeader>
              <CardContent>
                <div className={`text-3xl font-bold ${card.color}`}>{card.value}</div>
                <p className="mt-1 text-xs text-sena-lightBlue">Last updated today</p>
              </CardContent>
            </Card>
          );
        })}
      </div>

      {/* Expiring Forms */}
      {expiring.length > 0 && (
        <Card>
          <CardHeader>
            <CardTitle className="text-yellow-600">Expiring Soon</CardTitle>
          </CardHeader>
          <CardContent>
            <div className="space-y-3">
              {expiring.slice(0, 5).map((submission: any) => (
                <div key={submission._id} className="flex flex-col gap-2 p-3 rounded-lg bg-yellow-50 dark:bg-yellow-900/20 sm:flex-row sm:items-center sm:justify-between">
                  <div className="space-y-1">
                    <p className="font-medium">{submission.form.title}</p>
                    <p className="text-sm ">
                      Expires: {new Date(submission.expiryDate).toLocaleDateString()}
                    </p>
                  </div>
                  <div className="flex flex-wrap items-center gap-2">
                    <Badge variant="warning">
                      {getExpirationStatus(submission.expiryDate)}
                    </Badge>
                    <Link to={`/dashboard/submissions/${submission._id}`}>
                      <Button variant="outline" size="sm">
                        View Submission
                      </Button>
                    </Link>
                  </div>
                </div>
              ))}
            </div>
          </CardContent>
        </Card>
      )}

      {/* Upcoming Forms */}
      <Card>
        <CardHeader>
          <div className="flex items-center justify-between">
            <CardTitle className="flex items-center gap-2">
              <Calendar className="w-5 h-5" />
              Upcoming Forms
            </CardTitle>
            <Link to="/dashboard/forms">
              <Button variant="outline" size="sm">
                View All <ArrowRight className="w-4 h-4 ml-1" />
              </Button>
            </Link>
          </div>
        </CardHeader>
        <CardContent>
          {upcomingForms.length === 0 ? (
            <p className="">No upcoming forms</p>
          ) : (
            <Table>
              <TableHeader>
                <TableRow>
                  <TableHead>Form Name</TableHead>
                  <TableHead>Category</TableHead>
                  <TableHead>Due Date</TableHead>
                  <TableHead>Status</TableHead>
                  <TableHead>Action</TableHead>
                </TableRow>
              </TableHeader>
              <TableBody>
                {upcomingForms.map((form: any) => (
                  <TableRow key={form._id}>
                    <TableCell className="font-medium">{form.title}</TableCell>
                    <TableCell>
                      <Badge variant="outline">{form.category.displayName}</Badge>
                    </TableCell>
                    <TableCell>
                      {form.formLifecycleState === 'unsubmitted' ? 
                        <span className="text-red-600">Not submitted</span> :
                        <span className={form.colorClass.split(' ')[0]}>
                          {formatDate(form.displayDate)}
                        </span>
                      }
                    </TableCell>
                    <TableCell>
                      <Badge 
                        variant="outline" 
                        className={`${form.colorClass} border-current`}
                      >
                        {form.displayText}
                      </Badge>
                    </TableCell>
                    <TableCell>
                      <Link to={`/dashboard/forms/${form._id}`}>
                        <Button size="sm" variant="outline">
                          {form.formLifecycleState === 'unsubmitted' || form.formLifecycleState === 'expired' ? 
                            'Submit' : 'View'}
                        </Button>
                      </Link>
                    </TableCell>
                  </TableRow>
                ))}
              </TableBody>
            </Table>
          )}
        </CardContent>
      </Card>

      {/* Recent Submissions */}
      <Card>
        <CardHeader>
          <CardTitle>Recent Submissions</CardTitle>
        </CardHeader>
        <CardContent>
          {recent.length === 0 ? (
            <p className="">No recent submissions</p>
          ) : (
            <div className="space-y-3">
              {recent.map((submission: any) => (
                <div key={submission._id} className="flex flex-col gap-2 pb-3 border-b sm:flex-row sm:items-center sm:justify-between">
                  <div className="space-y-1">
                    <p className="font-medium">{submission.form.title}</p>
                    <p className="text-sm ">
                      Submitted: {new Date(submission.submittedAt).toLocaleDateString()}
                    </p>
                  </div>
                  <div className="flex flex-wrap items-center gap-2">
                    {getStatusBadge(submission.status)}
                    <Link to={`/dashboard/submissions/${submission._id}`}>
                      <Button variant="outline" size="sm">
                        View Submission
                      </Button>
                    </Link>
                  </div>
                </div>
              ))}
            </div>
          )}
        </CardContent>
      </Card>
    </div>
  );
};

export default UserDashboard;