"use client";

import React, { useState } from "react";
import { useQuery } from "@tanstack/react-query";
import { useParams, Link, Navigate } from "react-router-dom";
import api from "../../api";
import { Button } from "../../components/ui/button";
import { Card, CardContent, CardDescription, CardHeader, CardTitle } from "../../components/ui/card";
import { Badge } from "../../components/ui/badge";
import { Input } from "../../components/ui/input";
import { Select, SelectContent, SelectItem, SelectTrigger, SelectValue } from "../../components/ui/select";
import { Search, Filter, Calendar, Clock, ChevronLeft, ChevronRight } from "lucide-react";
import { formatDate } from "../../lib/utils";
import { useClientSearch } from "../../hooks/useDebounce";
import { useAuth } from "../../contexts/AuthContext";

const CategoryForms: React.FC = () => {
  const { category } = useParams<{ category: string }>();
  const { canAccessCategory } = useAuth();
  const [searchTerm, setSearchTerm] = useState("");
  const [currentPage, setCurrentPage] = useState(1);
  const [pageSize] = useState(6); // Forms per page
  
  // Check if user can access this category
  if (category && !canAccessCategory(category)) {
    return <Navigate to="/unauthorized" replace />;
  }
  
  // Filter states
  const [statusFilter, setStatusFilter] = useState<'all' | 'not-submitted' | 'soon' | 'later'>('all');
  const [formTypeFilter, setFormTypeFilter] = useState<'all' | 'regular' | 'table' | 'mixed'>('all');
  const [startDate, setStartDate] = useState("");
  const [endDate, setEndDate] = useState("");

  // Category mapping for display names and descriptions
  const categoryInfo = {
    eng: { name: "Engine Forms", description: "Forms related to engine operations and maintenance", icon: "🔧" },
    deck: { name: "Deck Forms", description: "Forms for deck operations and equipment", icon: "⚓" },
    mlc: { name: "MLC Forms", description: "Maritime Labour Convention compliance forms", icon: "👥" },
    isps: { name: "ISPS Forms", description: "International Ship and Port Facility Security forms", icon: "🛡️" },
    drill: { name: "Drill Forms", description: "Safety drill and emergency response forms", icon: "🚨" },
    deck_engine: { name: "Deck + Engine Forms", description: "Forms for both Deck and Engine departments", icon: "🤝" }
  };

  const currentCategory = category ? categoryInfo[category as keyof typeof categoryInfo] : null;

  // Fetch forms for this category with user-specific status
  const { data: forms, isLoading } = useQuery({
    queryKey: ["categoryForms", category],
    queryFn: () => api.getFormsWithUserStatus({
      category: category,
      isActive: true,
    }),
    enabled: !!category,
  });





  // Process and sort forms by due date
  const processedForms = React.useMemo(() => {
    if (!forms?.data?.data) return [];

    const formsWithStatus = forms.data.data.map((form: any) => {
      // Use the calculated status from the API
      const status = form.status;
      const calculatedStatus = form.calculatedStatus;
      
      // Determine display information based on API status
      let formLifecycleState: 'unsubmitted' | 'active' | 'expired';
      let displayDate: Date;
      let displayText: string;
      let colorClass: string;
      
      switch (status) {
        case 'not-submitted':
          formLifecycleState = 'unsubmitted';
          displayDate = new Date();
          displayText = "Not submitted";
          colorClass = "text-gray-600 bg-gray-100";
          break;
          
        case 'active':
          formLifecycleState = 'active';
          displayDate = calculatedStatus?.expiryDate ? new Date(calculatedStatus.expiryDate) : new Date();
          displayText = calculatedStatus?.daysLeft ? `Expires in ${calculatedStatus.daysLeft} days` : "Active";
          colorClass = "text-green-600 bg-green-100"; // Good
          break;
          
        case 'expiring-soon':
          formLifecycleState = 'active';
          displayDate = calculatedStatus?.expiryDate ? new Date(calculatedStatus.expiryDate) : new Date();
          displayText = calculatedStatus?.daysLeft ? `Expires in ${calculatedStatus.daysLeft} days` : "Expiring soon";
          colorClass = "text-yellow-600 bg-yellow-100"; // Warning
          break;
          
        case 'expired':
          formLifecycleState = 'expired';
          displayDate = calculatedStatus?.expiryDate ? new Date(calculatedStatus.expiryDate) : new Date();
          displayText = calculatedStatus?.daysLeft ? `Expired ${Math.abs(calculatedStatus.daysLeft)} days ago` : "Expired";
          colorClass = "text-red-600 bg-red-100";
          break;
          
        default:
          formLifecycleState = 'unsubmitted';
          displayDate = new Date();
          displayText = "Unknown status";
          colorClass = "text-gray-600 bg-gray-100";
          break;
      }
      
      return {
        ...form,
        formLifecycleState,
        displayDate,
        displayText,
        colorClass,
        hasSubmission: status !== 'not-submitted'
      };
    });

    return formsWithStatus;
  }, [forms]);

  // Apply client-side search
  const searchFields = ['title', 'description'];
  const searchedForms = useClientSearch(processedForms, searchTerm, searchFields, 300);

  // Apply other filters
  const filteredForms = React.useMemo(() => {
    let filtered = searchedForms;

    // Filter by status
    if (statusFilter !== 'all') {
      filtered = filtered.filter((form: any) => {
        if (statusFilter === 'not-submitted') {
          return form.formLifecycleState === 'unsubmitted';
        } else if (statusFilter === 'soon') {
          if (form.formLifecycleState === 'active') {
            const today = new Date();
            today.setHours(0, 0, 0, 0);
            const expiryDate = new Date(form.displayDate);
            expiryDate.setHours(0, 0, 0, 0);
            const daysUntilExpiry = Math.floor((expiryDate.getTime() - today.getTime()) / (1000 * 60 * 60 * 24));
            return daysUntilExpiry <= 15; // Expires soon
          }
          return false;
        } else if (statusFilter === 'later') {
          if (form.formLifecycleState === 'active') {
            const today = new Date();
            today.setHours(0, 0, 0, 0);
            const expiryDate = new Date(form.displayDate);
            expiryDate.setHours(0, 0, 0, 0);
            const daysUntilExpiry = Math.floor((expiryDate.getTime() - today.getTime()) / (1000 * 60 * 60 * 24));
            return daysUntilExpiry > 15; // Expires later
          }
          return false;
        }
        return true;
      });
    }

    // Filter by form type
    if (formTypeFilter !== 'all') {
      filtered = filtered.filter((form: any) => form.formType === formTypeFilter);
    }

    // Filter by date range
    if (startDate || endDate) {
      filtered = filtered.filter((form: any) => {
        if (!form.latestSubmission) return !startDate; // Include unsubmitted if no start date
        
        const submissionDate = new Date(form.latestSubmission.submittedAt || form.latestSubmission.createdAt);
        const start = startDate ? new Date(startDate) : new Date('1970-01-01');
        const end = endDate ? new Date(endDate) : new Date('2099-12-31');
        
        return submissionDate >= start && submissionDate <= end;
      });
    }

    // Sort by priority: unsubmitted/expired first, then active by expiry date
    const sortedForms = filtered.sort((a: any, b: any) => {
      // Priority order: unsubmitted -> expired -> active
      const priority: Record<string, number> = { 'unsubmitted': 1, 'expired': 2, 'active': 3 };
      const aPriority = priority[a.formLifecycleState] || 5;
      const bPriority = priority[b.formLifecycleState] || 5;
      
      if (aPriority !== bPriority) {
        return aPriority - bPriority;
      }
      
      // Within same priority, sort by display date (earliest first for active forms)
      return a.displayDate.getTime() - b.displayDate.getTime();
    });

    return sortedForms;
  }, [searchedForms, statusFilter, formTypeFilter, startDate, endDate]);

  // Pagination calculations
  const totalForms = filteredForms.length;
  const totalPages = Math.ceil(totalForms / pageSize);
  const startIndex = (currentPage - 1) * pageSize;
  const endIndex = startIndex + pageSize;
  const paginatedForms = filteredForms.slice(startIndex, endIndex);

  // Reset to first page when filters change
  React.useEffect(() => {
    setCurrentPage(1);
  }, [statusFilter, formTypeFilter, startDate, endDate, searchTerm]);

  if (!currentCategory) {
    return (
      <div className="flex items-center justify-center h-64">
        <p className="">Category not found</p>
      </div>
    );
  }

  return (
    <div className="space-y-6">
      {/* Header */}
      <div className="flex items-center justify-between">
        <div>
          <h1 className="flex items-center gap-3 text-3xl font-bold">
            <span className="text-2xl">{currentCategory.icon}</span>
            {currentCategory.name}
          </h1>
          <p className="mt-1 ">{currentCategory.description}</p>
        </div>
      </div>

      {/* Filters */}
      <Card>
        <CardHeader>
          <CardTitle className="flex items-center gap-2">
            <Filter className="w-5 h-5" />
            Filters & Search
          </CardTitle>
        </CardHeader>
        <CardContent className="space-y-4">
          {/* Search */}
          <div className="relative max-w-md">
            <Search className="absolute w-4 h-4 left-3 top-3 " />
            <Input
              placeholder="Search forms..."
              value={searchTerm}
              onChange={(e) => setSearchTerm(e.target.value)}
              className="pl-9"
            />
          </div>

          {/* Filter Controls */}
          <div className="grid grid-cols-1 gap-4 md:grid-cols-2 lg:grid-cols-4">
            {/* Status Filter */}
            <div>
              <label className="block mb-2 text-sm font-medium">Status</label>
              <Select value={statusFilter} onValueChange={(value: any) => setStatusFilter(value)}>
                <SelectTrigger>
                  <SelectValue placeholder="Filter by status" />
                </SelectTrigger>
                <SelectContent>
                  <SelectItem value="all">All Forms</SelectItem>
                  <SelectItem value="not-submitted">Not Submitted</SelectItem>
                  <SelectItem value="soon">Expires Soon (≤15 days)</SelectItem>
                  <SelectItem value="later">Expires Later (&gt;15 days)</SelectItem>
                </SelectContent>
              </Select>
            </div>

            {/* Form Type Filter */}
            <div>
              <label className="block mb-2 text-sm font-medium">Form Type</label>
              <Select value={formTypeFilter} onValueChange={(value: any) => setFormTypeFilter(value)}>
                <SelectTrigger>
                  <SelectValue placeholder="Filter by type" />
                </SelectTrigger>
                <SelectContent>
                  <SelectItem value="all">All Types</SelectItem>
                  <SelectItem value="regular">Regular Forms</SelectItem>
                  <SelectItem value="table">Table Forms</SelectItem>
                  <SelectItem value="mixed">Mixed Forms</SelectItem>
                </SelectContent>
              </Select>
            </div>

            {/* Start Date Filter */}
            <div>
              <label className="block mb-2 text-sm font-medium">From Date</label>
              <Input
                type="date"
                value={startDate}
                onChange={(e) => setStartDate(e.target.value)}
              />
            </div>

            {/* End Date Filter */}
            <div>
              <label className="block mb-2 text-sm font-medium">To Date</label>
              <Input
                type="date"
                value={endDate}
                onChange={(e) => setEndDate(e.target.value)}
              />
            </div>
          </div>

          {/* Results Summary */}
          <div className="flex items-center justify-between text-sm ">
            <span>Showing {Math.min(startIndex + 1, totalForms)}-{Math.min(endIndex, totalForms)} of {totalForms} forms</span>
            <Button 
              variant="outline" 
              size="sm" 
              onClick={() => {
                setSearchTerm("");
                setStatusFilter('all');
                setFormTypeFilter('all');
                setStartDate("");
                setEndDate("");
                setCurrentPage(1);
              }}
            >
              Clear Filters
            </Button>
          </div>
        </CardContent>
      </Card>

      {/* Forms List */}
      {isLoading ? (
        <div className="flex items-center justify-center h-64">
          <div className="w-12 h-12 border-b-2 rounded-full animate-spin border-primary"></div>
        </div>
      ) : processedForms.length === 0 ? (
        <Card>
          <CardContent className="flex flex-col items-center justify-center h-32 space-y-2">
            <p className="">No forms found matching your criteria</p>
            <Button 
              variant="outline" 
              onClick={() => {
                setSearchTerm("");
                setStatusFilter('all');
                setFormTypeFilter('all');
                setStartDate("");
                setEndDate("");
                setCurrentPage(1);
              }}
            >
              Clear all filters
            </Button>
          </CardContent>
        </Card>
      ) : (
        <div className="space-y-4">
          {paginatedForms.map((form: any) => (
            <Card 
              key={form._id} 
              className="transition-shadow border-l-4 hover:shadow-md border-l-gray-200"
            >
              <CardHeader>
                <div className="flex items-center justify-between">
                  <div className="flex-1">
                    <CardTitle className="flex items-center gap-2 text-lg">
                      {form.title}
                      <Badge 
                        variant="outline" 
                        className={`${form.colorClass} border-current`}
                      >
                        {form.displayText}
                      </Badge>
                    </CardTitle>
                    <CardDescription className="mt-1">{form.description}</CardDescription>
                  </div>
                </div>
              </CardHeader>
              <CardContent>
                <div className="grid grid-cols-1 gap-4 mb-4 md:grid-cols-3">
                  <div className="flex items-center gap-2 text-sm">
                    <Calendar className="w-4 h-4 " />
                    <span className="">
                      {form.formLifecycleState === 'unsubmitted' ? 'Status:' : 'Due Date:'}
                    </span>
                    <span className={form.colorClass.split(' ')[0]}>
                      {form.formLifecycleState === 'unsubmitted' ? 
                        'Not submitted' : 
                        formatDate(form.displayDate)
                      }
                    </span>
                  </div>
                  
                  <div className="flex items-center gap-2 text-sm">
                    <Clock className="w-4 h-4 text-card-foreground" />
                    <span className="font-medium text-card-foreground">Validity:</span>
                    <span>{form.validityPeriod} days</span>
                  </div>

                  <div className="flex items-center gap-2 text-sm">
                    <span className="font-medium text-card-foreground">Type:</span>
                    <Badge variant="outline">{form.formType}</Badge>
                  </div>
                </div>

                <div className="flex justify-end">
                  <Link to={`/dashboard/forms/${form._id}`}>
                    <Button size="sm">
                      {form.formLifecycleState === 'unsubmitted' || form.formLifecycleState === 'expired' ? 
                        (form.formType === "table" ? "Submit Table" : "Submit Form") :
                        "View Submission"
                      }
                    </Button>
                  </Link>
                </div>
              </CardContent>
            </Card>
          ))}
        </div>
      )}

      {/* Pagination */}
      {totalPages > 1 && (
        <Card>
          <CardContent className="pt-6">
            <div className="flex items-center justify-between">
              <div className="text-sm ">
                Page {currentPage} of {totalPages}
              </div>
              <div className="flex items-center space-x-2">
                <Button
                  variant="outline"
                  size="sm"
                  onClick={() => setCurrentPage(Math.max(1, currentPage - 1))}
                  disabled={currentPage === 1}
                >
                  <ChevronLeft className="w-4 h-4" />
                  Previous
                </Button>
                
                {/* Page numbers */}
                <div className="flex space-x-1">
                  {Array.from({ length: Math.min(5, totalPages) }, (_, i) => {
                    let pageNum;
                    if (totalPages <= 5) {
                      pageNum = i + 1;
                    } else if (currentPage <= 3) {
                      pageNum = i + 1;
                    } else if (currentPage >= totalPages - 2) {
                      pageNum = totalPages - 4 + i;
                    } else {
                      pageNum = currentPage - 2 + i;
                    }
                    
                    return (
                      <Button
                        key={pageNum}
                        variant={currentPage === pageNum ? "default" : "outline"}
                        size="sm"
                        onClick={() => setCurrentPage(pageNum)}
                      >
                        {pageNum}
                      </Button>
                    );
                  })}
                </div>
                
                <Button
                  variant="outline"
                  size="sm"
                  onClick={() => setCurrentPage(Math.min(totalPages, currentPage + 1))}
                  disabled={currentPage === totalPages}
                >
                  Next
                  <ChevronRight className="w-4 h-4" />
                </Button>
              </div>
            </div>
          </CardContent>
        </Card>
      )}
    </div>
  );
};

export default CategoryForms;