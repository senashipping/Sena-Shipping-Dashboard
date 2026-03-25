import React, { useState } from "react";
import { useQuery } from "@tanstack/react-query";
import { Link } from "react-router-dom";
import api from "../../api";
import { Card, CardContent, CardHeader, CardTitle } from "../../components/ui/card";
import { Button } from "../../components/ui/button";
import { Badge } from "../../components/ui/badge";
import { Users, Ship, FileText, ClipboardList, Activity, TrendingUp, Calendar, Eye } from "lucide-react";
import { BarChart, Bar, XAxis, YAxis, CartesianGrid, Tooltip, ResponsiveContainer, PieChart, Pie, Cell, Area, AreaChart } from 'recharts';
import SubmissionViewModal from "../../components/SubmissionViewModal";

const AdminDashboard: React.FC = () => {
  const [selectedSubmissionId, setSelectedSubmissionId] = useState<string | null>(null);
  const [isModalOpen, setIsModalOpen] = useState(false);

  const { data: dashboardData, isLoading } = useQuery({
    queryKey: ["adminDashboard"],
    queryFn: () => api.getAdminDashboard(),
  });

  // Get recent submissions for activity section
  const { data: recentSubmissions } = useQuery({
    queryKey: ["recent-submissions"],
    queryFn: () => api.getSubmissions({ page: 1, limit: 10 }),
  });

  const handleViewSubmission = (submissionId: string) => {
    setSelectedSubmissionId(submissionId);
    setIsModalOpen(true);
  };

  // Category mapping for better display names
  const getCategoryDisplayName = (category: any) => {
    if (!category) return 'Unknown';
    
    // If it already has displayName, use it
    if (category.displayName) return category.displayName;
    
    // Map category names to display names
    const categoryMap: { [key: string]: string } = {
      'eng': 'Engine Forms',
      'deck': 'Deck Forms', 
      'mlc': 'MLC Forms',
      'isps': 'ISPS Forms',
      'drill': 'Drill Forms'
    };
    
    return categoryMap[category.name] || category.name || 'Unknown';
  };

  // Get all submissions for chart calculations
  const { data: allSubmissions } = useQuery({
    queryKey: ["all-submissions-chart"],
    queryFn: () => api.getSubmissions({ page: 1, limit: 1000 }),
  });

  const stats = dashboardData?.data?.data?.overview || {};
  const submissions = recentSubmissions?.data?.data || [];
  const allSubmissionsData = allSubmissions?.data?.data || [];

  const statCards = [
    {
      title: "Total Users",
      value: stats.totalUsers || 0,
      icon: Users,
      color: "text-blue-600",
      bgColor: "bg-blue-100",
    },
    {
      title: "Total Ships",
      value: stats.totalShips || 0,
      icon: Ship,
      color: "text-green-600",
      bgColor: "bg-green-100",
    },
    {
      title: "Total Forms",
      value: stats.totalForms || 0,
      icon: FileText,
      color: "text-purple-600",
      bgColor: "bg-purple-100",
    },
    {
      title: "Total Submissions",
      value: stats.totalSubmissions || 0,
      icon: ClipboardList,
      color: "text-orange-600",
      bgColor: "bg-orange-100",
    },
  ];

  // Generate dynamic chart data from real submissions
  const generateSubmissionTrends = () => {
    if (!allSubmissionsData || allSubmissionsData.length === 0) {
      return [
        { month: 'Jan', submissions: 0 },
        { month: 'Feb', submissions: 0 },
        { month: 'Mar', submissions: 0 },
        { month: 'Apr', submissions: 0 },
        { month: 'May', submissions: 0 },
        { month: 'Jun', submissions: 0 },
      ];
    }

    // Get last 6 months of data
    const monthlyData: { [key: string]: number } = {};
    const months = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
    
    // Initialize last 6 months with 0
    const now = new Date();
    for (let i = 5; i >= 0; i--) {
      const date = new Date(now.getFullYear(), now.getMonth() - i, 1);
      const monthName = months[date.getMonth()];
      monthlyData[monthName] = 0;
    }

    // Count submissions by month
    allSubmissionsData.forEach((submission: any) => {
      if (submission.submittedAt || submission.createdAt) {
        const date = new Date(submission.submittedAt || submission.createdAt);
        const monthName = months[date.getMonth()];
        if (monthlyData.hasOwnProperty(monthName)) {
          monthlyData[monthName]++;
        }
      }
    });

    return Object.entries(monthlyData).map(([month, count]) => ({
      month,
      submissions: count
    }));
  };

  const generateCategoryData = () => {
    if (!allSubmissionsData || allSubmissionsData.length === 0) {
      return [
        { name: 'Engine Forms', value: 0, color: '#3B82F6' },
        { name: 'Deck Forms', value: 0, color: '#10B981' },
        { name: 'MLC Forms', value: 0, color: '#8B5CF6' },
        { name: 'ISPS Forms', value: 0, color: '#F59E0B' },
        { name: 'Drill Forms', value: 0, color: '#EF4444' },
      ];
    }

    const categoryCount: { [key: string]: number } = {
      'eng': 0,
      'deck': 0,
      'mlc': 0,
      'isps': 0,
      'drill': 0
    };

    const total = allSubmissionsData.length;

    allSubmissionsData.forEach((submission: any) => {
      const category = submission.form?.category?.name || submission.form?.category;
      if (categoryCount.hasOwnProperty(category)) {
        categoryCount[category]++;
      }
    });

    return [
      { name: 'Engine Forms', value: total > 0 ? Math.round((categoryCount.eng / total) * 100) : 0, color: '#3B82F6' },
      { name: 'Deck Forms', value: total > 0 ? Math.round((categoryCount.deck / total) * 100) : 0, color: '#10B981' },
      { name: 'MLC Forms', value: total > 0 ? Math.round((categoryCount.mlc / total) * 100) : 0, color: '#8B5CF6' },
      { name: 'ISPS Forms', value: total > 0 ? Math.round((categoryCount.isps / total) * 100) : 0, color: '#F59E0B' },
      { name: 'Drill Forms', value: total > 0 ? Math.round((categoryCount.drill / total) * 100) : 0, color: '#EF4444' },
    ];
  };

  const generateFormTypeData = () => {
    if (!allSubmissionsData || allSubmissionsData.length === 0) {
      return [
        { name: 'Regular', submissions: 0 },
        { name: 'Table', submissions: 0 },
        { name: 'Mixed', submissions: 0 },
      ];
    }

    const typeCount: { [key: string]: number } = {
      'regular': 0,
      'table': 0,
      'mixed': 0
    };

    allSubmissionsData.forEach((submission: any) => {
      let formType = submission.form?.formType || submission.form?.type;
      
      // If no form type, try to infer from submission data
      if (!formType && submission.data) {
        const hasTableData = Object.keys(submission.data).some(key => 
          key.startsWith('table_') || key === 'tableData' || 
          (Array.isArray(submission.data[key]) && typeof submission.data[key][0] === 'object')
        );
        const hasRegularFields = Object.keys(submission.data).some(key => 
          !key.startsWith('table_') && key !== 'tableData' && 
          !Array.isArray(submission.data[key])
        );
        
        if (hasTableData && hasRegularFields) {
          formType = 'mixed';
        } else if (hasTableData) {
          formType = 'table';
        } else {
          formType = 'regular';
        }
      }
      
      formType = formType || 'regular';
      if (typeCount.hasOwnProperty(formType)) {
        typeCount[formType]++;
      }
    });

    return [
      { name: 'Regular', submissions: typeCount.regular },
      { name: 'Table', submissions: typeCount.table },
      { name: 'Mixed', submissions: typeCount.mixed },
    ];
  };

  const submissionTrends = generateSubmissionTrends();
  const categoryData = generateCategoryData();
  const formTypeData = generateFormTypeData();

  const formatDate = (dateString: string) => {
    return new Date(dateString).toLocaleDateString("en-US", {
      month: "short",
      day: "numeric",
      hour: "2-digit",
      minute: "2-digit",
    });
  };

  if (isLoading) {
    return (
      <div className="flex items-center justify-center h-64">
        <div className="w-12 h-12 border-b-2 rounded-full animate-spin border-primary"></div>
      </div>
    );
  }

  const hasSubmissionData = allSubmissionsData && allSubmissionsData.length > 0;

  return (
    <div className="space-y-4 sm:space-y-6">
      <div className="flex flex-col gap-4 sm:flex-row sm:items-center sm:justify-between">
        <h1 className="text-2xl font-bold sm:text-3xl">Admin Dashboard</h1>
        <div className="flex items-center gap-2 text-sm ">
          <Calendar className="w-4 h-4" />
          <span className="hidden sm:inline">{new Date().toLocaleDateString("en-US", { 
            weekday: 'long', 
            year: 'numeric', 
            month: 'long', 
            day: 'numeric' 
          })}</span>
          <span className="sm:hidden">{new Date().toLocaleDateString("en-US", { 
            month: 'short', 
            day: 'numeric',
            year: 'numeric'
          })}</span>
        </div>
      </div>

      {/* Statistics Grid */}
      <div className="grid grid-cols-1 gap-4 sm:gap-6 sm:grid-cols-2 lg:grid-cols-4">
        {statCards.map((card, index) => {
          const IconComponent = card.icon;
          return (
            <Card key={index} className="transition-shadow hover:shadow-md">
              <CardHeader className="flex flex-row items-center justify-between pb-2 space-y-0">
                <CardTitle className="text-sm font-medium">
                  {card.title}
                </CardTitle>
                <div className={`p-2 rounded-full ${card.bgColor}`}>
                  <IconComponent className={`h-4 w-4 ${card.color}`} />
                </div>
              </CardHeader>
              <CardContent>
                <div className="text-2xl font-bold">{card.value}</div>
              </CardContent>
            </Card>
          );
        })}
      </div>

      {/* Charts Row */}
      <div className="grid grid-cols-1 gap-4 sm:gap-6 xl:grid-cols-2">
        {/* Submission Trends Chart */}
        <Card>
          <CardHeader>
            <CardTitle className="flex items-center gap-2 text-lg sm:text-xl">
              <TrendingUp className="w-4 h-4 sm:w-5 sm:h-5" />
              Submission Trends
            </CardTitle>
          </CardHeader>
          <CardContent>
            {hasSubmissionData ? (
              <ResponsiveContainer width="100%" height={250} className="sm:h-[300px]">
                <AreaChart data={submissionTrends}>
                  <CartesianGrid strokeDasharray="3 3" />
                  <XAxis dataKey="month" />
                  <YAxis />
                  <Tooltip />
                  <Area 
                    type="monotone" 
                    dataKey="submissions" 
                    stroke="#3B82F6" 
                    fill="#3B82F6" 
                    fillOpacity={0.1}
                  />
                </AreaChart>
              </ResponsiveContainer>
            ) : (
              <div className="flex items-center justify-center h-[300px] ">
                <div className="text-center">
                  <TrendingUp className="w-12 h-12 mx-auto mb-4 opacity-50" />
                  <p>No submission data available</p>
                  <p className="text-sm">Charts will appear when forms are submitted</p>
                </div>
              </div>
            )}
          </CardContent>
        </Card>

        {/* Category Distribution */}
        <Card>
          <CardHeader>
            <CardTitle className="text-lg sm:text-xl">Form Categories</CardTitle>
          </CardHeader>
          <CardContent>
            {hasSubmissionData ? (
              <>
                <ResponsiveContainer width="100%" height={250} className="sm:h-[300px]">
                  <PieChart>
                    <Pie
                      data={categoryData}
                      cx="50%"
                      cy="50%"
                      innerRadius={60}
                      outerRadius={100}
                      paddingAngle={5}
                      dataKey="value"
                    >
                      {categoryData.map((entry, index) => (
                        <Cell key={`cell-${index}`} fill={entry.color} />
                      ))}
                    </Pie>
                    <Tooltip />
                  </PieChart>
                </ResponsiveContainer>
                <div className="grid grid-cols-2 gap-2 mt-4">
                  {categoryData.map((item, index) => (
                    <div key={index} className="flex items-center gap-2 text-sm">
                      <div 
                        className="w-3 h-3 rounded-full" 
                        style={{ backgroundColor: item.color }}
                      />
                      <span>{item.name}</span>
                      <span className="">({item.value}%)</span>
                    </div>
                  ))}
                </div>
              </>
            ) : (
              <div className="flex items-center justify-center h-[300px] ">
                <div className="text-center">
                  <FileText className="w-12 h-12 mx-auto mb-4 opacity-50" />
                  <p>No category data available</p>
                  <p className="text-sm">Distribution will show when forms are submitted</p>
                </div>
              </div>
            )}
          </CardContent>
        </Card>
      </div>

      {/* Form Types Chart */}
      <Card>
        <CardHeader>
          <CardTitle className="text-lg sm:text-xl">Form Type Distribution</CardTitle>
        </CardHeader>
        <CardContent>
          {hasSubmissionData ? (
            <ResponsiveContainer width="100%" height={250} className="sm:h-[300px]">
              <BarChart data={formTypeData}>
                <CartesianGrid strokeDasharray="3 3" />
                <XAxis dataKey="name" />
                <YAxis />
                <Tooltip />
                <Bar dataKey="submissions" fill="#8B5CF6" radius={[4, 4, 0, 0]} />
              </BarChart>
            </ResponsiveContainer>
          ) : (
            <div className="flex items-center justify-center h-[300px] ">
              <div className="text-center">
                <ClipboardList className="w-12 h-12 mx-auto mb-4 opacity-50" />
                <p>No form type data available</p>
                <p className="text-sm">Distribution will show when forms are submitted</p>
              </div>
            </div>
          )}
        </CardContent>
      </Card>

      {/* Recent Activity */}
      <Card>
        <CardHeader className="flex flex-col gap-4 sm:flex-row sm:items-center sm:justify-between">
          <CardTitle className="flex items-center gap-2 text-lg sm:text-xl">
            <Activity className="w-4 h-4 sm:w-5 sm:h-5" />
            Recent Activity
          </CardTitle>
          <Button variant="outline" size="sm" asChild>
            <Link to="/admin/submissions">View All</Link>
          </Button>
        </CardHeader>
        <CardContent>
          {submissions.length === 0 ? (
            <p className="py-8 text-center ">
              No recent submissions found.
            </p>
          ) : (
            <div className="space-y-3 sm:space-y-4">
              {submissions.slice(0, 10).map((submission: any) => (
                <div key={submission._id} className="flex flex-col gap-3 p-3 transition-colors border rounded-lg sm:flex-row sm:items-center sm:justify-between sm:p-4 bg-card hover:bg-accent/50 sm:gap-4">
                  <div className="flex items-center min-w-0 gap-3 sm:gap-4">
                    <div className="flex-shrink-0 p-2 bg-blue-100 rounded-full">
                      <FileText className="w-4 h-4 text-blue-600" />
                    </div>
                    <div className="min-w-0">
                      <div className="font-medium truncate">{submission.form?.title || 'Unknown Form'}</div>
                      <div className="text-sm truncate">
                        <span className="hidden sm:inline">Submitted by </span>
                        {submission.user?.name}
                        <span className="hidden sm:inline"> • {submission.ship?.name}</span>
                      </div>
                    </div>
                  </div>
                  <div className="flex flex-col items-start flex-shrink-0 gap-2 sm:flex-row sm:items-center sm:gap-3">
                    <Badge variant="outline" className="text-xs">
                      {getCategoryDisplayName(submission.form?.category)}
                    </Badge>
                    <span className="text-xs sm:text-sm ">
                      {formatDate(submission.submittedAt || submission.createdAt)}
                    </span>
                    <Button 
                      variant="ghost" 
                      size="sm" 
                      className="w-8 h-8 sm:h-9 sm:w-9"
                      onClick={() => handleViewSubmission(submission._id)}
                    >
                      <Eye className="w-3 h-3 sm:w-4 sm:h-4" />
                    </Button>
                  </div>
                </div>
              ))}
            </div>
          )}
        </CardContent>
      </Card>

      {/* Submission View Modal */}
      <SubmissionViewModal
        submissionId={selectedSubmissionId}
        isOpen={isModalOpen}
        onClose={() => {
          setIsModalOpen(false);
          setSelectedSubmissionId(null);
        }}
      />
    </div>
  );
};

export default AdminDashboard;