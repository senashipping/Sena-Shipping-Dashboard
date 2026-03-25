import React, { useState } from "react";
import { useQuery } from "@tanstack/react-query";
import { useParams, useNavigate } from "react-router-dom";
import api from "../../api";
import { Button } from "../../components/ui/button";
import { Card, CardContent, CardHeader, CardTitle, CardDescription } from "../../components/ui/card";
import { Input } from "../../components/ui/input";
import { Select, SelectContent, SelectItem, SelectTrigger, SelectValue } from "../../components/ui/select";
import { Badge } from "../../components/ui/badge";
import { Table, TableBody, TableCell, TableHead, TableHeader, TableRow } from "../../components/ui/table";
import { Search, Eye, ArrowLeft, Filter, ChevronLeft, ChevronRight } from "lucide-react";
import { formatDate } from "../../lib/utils";
import SubmissionViewModal from "../../components/SubmissionViewModal";
import { useClientSearch } from "../../hooks/useDebounce";

const UserSubmissions: React.FC = () => {
  const { userId } = useParams<{ userId: string }>();
  const navigate = useNavigate();
  const [search, setSearch] = useState("");
  const [categoryFilter, setCategoryFilter] = useState("all");
  const [formTypeFilter, setFormTypeFilter] = useState("all");
  const [startDate, setStartDate] = useState("");
  const [endDate, setEndDate] = useState("");
  const [currentPage, setCurrentPage] = useState(1);
  const [pageSize] = useState(10);
  const [selectedSubmissionId, setSelectedSubmissionId] = useState<string | null>(null);
  const [isModalOpen, setIsModalOpen] = useState(false);

  const { data: userData, isLoading: userLoading } = useQuery({
    queryKey: ["user", userId],
    queryFn: () => api.getUser(userId!),
    enabled: !!userId,
  });

  const { data: submissionsData, isLoading: submissionsLoading } = useQuery({
    queryKey: ["user-submissions", userId],
    queryFn: () => api.getSubmissions({
      user: userId,
      page: 1,
      limit: 1000, // Get all results for client-side filtering
    }),
    enabled: !!userId,
  });

  // Get unique form IDs from submissions to fetch full form details
  const formIds = React.useMemo(() => {
    if (!submissionsData?.data?.data) return [];
    const ids = submissionsData.data.data.map((sub: any) => sub.form?._id).filter(Boolean) as string[];
    return [...new Set(ids)]; // Remove duplicates
  }, [submissionsData]);

  // Fetch full form details for all forms used in submissions
  const { data: formsData, isLoading: formsLoading } = useQuery({
    queryKey: ["forms-details", formIds],
    queryFn: async () => {
      if (formIds.length === 0) return [];
      // Make individual API calls for each form
      const promises = formIds.map((formId: string) => api.getFormById(formId));
      const results = await Promise.all(promises);
      return results.map((result: any) => result.data?.data || result.data).filter(Boolean);
    },
    enabled: formIds.length > 0,
  });

  const user = userData?.data?.data;
  const rawSubmissions = Array.isArray(submissionsData?.data?.data) ? submissionsData.data.data : [];
  
  // Create a map of form IDs to full form data
  const formsMap = React.useMemo(() => {
    if (!formsData) return new Map();
    const map = new Map();
    formsData.forEach((form: any) => {
      map.set(form._id, form);
    });
    return map;
  }, [formsData]);

  // Category mapping for filtering
  const categoryInfo = {
    eng: "Engine Forms",
    deck: "Deck Forms", 
    mlc: "MLC Forms",
    isps: "ISPS Forms",
    drill: "Drill Forms"
  };

  // Process submissions with status calculation and enhanced form data
  const processedSubmissions = React.useMemo(() => {
    return rawSubmissions.map((sub: any) => {
      let actualStatus = sub.status;
      
      // Get full form data from the forms map
      const fullForm = formsMap.get(sub.form?._id) || sub.form;
      
      // Calculate if submission is expired based on form validity or existing expiry date
      if (sub.status === 'approved' && sub.submittedAt) {
        if (sub.expiryDate) {
          // Use existing expiry date from API
          const expiryDate = new Date(sub.expiryDate);
          const today = new Date();
          if (today > expiryDate) {
            actualStatus = 'expired';
          }
        } else if (fullForm) {
          // Calculate from validity period
          const submissionDate = new Date(sub.submittedAt);
          const validityPeriod = fullForm.validityPeriod || 30;
          const expiryDate = new Date(submissionDate);
          expiryDate.setDate(expiryDate.getDate() + validityPeriod);
          
          const today = new Date();
          if (today > expiryDate) {
            actualStatus = 'expired';
          }
        }
      }
      
      return {
        ...sub,
        form: fullForm, // Replace with full form data
        actualStatus,
        isExpired: actualStatus === 'expired'
      };
    });
  }, [rawSubmissions, formsMap]);

  // Apply client-side search
  const searchFields = ['form.title', 'ship.name', 'user.name', 'user.email'];
  const searchedSubmissions = useClientSearch(processedSubmissions, search, searchFields);



  // Apply other filters
  const filteredSubmissions = React.useMemo(() => {
    let filtered = searchedSubmissions;

    // Apply category filter
    if (categoryFilter !== 'all') {
      filtered = filtered.filter((sub: any) => 
        sub.form && (sub.form.category === categoryFilter || sub.form.category?.name === categoryFilter)
      );
    }

    // Apply form type filter - Enhanced with intelligent type detection
    if (formTypeFilter !== 'all') {
      filtered = filtered.filter((sub: any) => {
        if (!sub.form && !sub.data) return false;
        
        // Get the form type using the same logic as display
        let formType = sub.form?.formType || sub.form?.type || sub.form?.formStyle;
        
        // If no form type specified, infer from submission data
        if (!formType && sub.data) {
          const hasTableData = Object.keys(sub.data).some(key => 
            key.startsWith('table_') || key === 'tableData' || 
            (Array.isArray(sub.data[key]) && typeof sub.data[key][0] === 'object')
          );
          
          const hasRegularFields = Object.keys(sub.data).some(key => 
            !key.startsWith('table_') && key !== 'tableData' && 
            !Array.isArray(sub.data[key])
          );
          
          if (hasTableData && hasRegularFields) {
            formType = 'mixed';
          } else if (hasTableData) {
            formType = 'table';
          } else {
            formType = 'regular';
          }
        }
        
        // Default to regular if still no type
        formType = formType || 'regular';
        
        // Normalize for comparison
        const normalizedFormType = String(formType).toLowerCase().trim();
        const normalizedFilter = formTypeFilter.toLowerCase().trim();
        
        return normalizedFormType === normalizedFilter;
      });
    }

    // Apply date range filter
    if (startDate || endDate) {
      filtered = filtered.filter((sub: any) => {
        if (!sub.submittedAt) return false;
        
        const subDate = new Date(sub.submittedAt);
        const start = startDate ? new Date(startDate) : new Date('1970-01-01');
        const end = endDate ? new Date(endDate) : new Date('2099-12-31');
        
        return subDate >= start && subDate <= end;
      });
    }

    return filtered;
  }, [searchedSubmissions, categoryFilter, formTypeFilter, startDate, endDate]);

  // Pagination for filtered results
  const totalItems = filteredSubmissions.length;
  const totalPages = Math.ceil(totalItems / pageSize);
  const startIndex = (currentPage - 1) * pageSize;
  const endIndex = Math.min(startIndex + pageSize, totalItems);
  const submissions = filteredSubmissions.slice(startIndex, endIndex);

  const isLoading = userLoading || submissionsLoading || formsLoading;

  if (isLoading) {
    return <div className="flex items-center justify-center h-64"><div className="w-12 h-12 border-b-2 rounded-full animate-spin border-primary"></div></div>;
  }

  return (
    <div className="space-y-6">
      <div className="flex items-center gap-4">
        <Button variant="outline" size="icon" onClick={() => navigate("/admin/users")}>
          <ArrowLeft className="w-4 h-4" />
        </Button>
        <div>
          <h1 className="text-3xl font-bold">Submissions for {user?.name}</h1>
          <p className="">{user?.email}</p>
        </div>
      </div>

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
              placeholder="Search by form or ship..."
              value={search}
              onChange={(e) => setSearch(e.target.value)}
              className="pl-9"
            />
          </div>

          {/* Filter Controls */}
          <div className="grid grid-cols-1 gap-4 md:grid-cols-2 lg:grid-cols-4">
            {/* Category Filter */}
            <div>
              <label className="block mb-2 text-sm font-medium">Category</label>
              <Select value={categoryFilter} onValueChange={setCategoryFilter}>
                <SelectTrigger>
                  <SelectValue placeholder="Filter by category" />
                </SelectTrigger>
                <SelectContent>
                  <SelectItem value="all">All Categories</SelectItem>
                  <SelectItem value="eng">Engine Forms</SelectItem>
                  <SelectItem value="deck">Deck Forms</SelectItem>
                  <SelectItem value="mlc">MLC Forms</SelectItem>
                  <SelectItem value="isps">ISPS Forms</SelectItem>
                  <SelectItem value="drill">Drill Forms</SelectItem>
                </SelectContent>
              </Select>
            </div>

            {/* Form Type Filter */}
            <div>
              <label className="block mb-2 text-sm font-medium">Form Type</label>
              <Select value={formTypeFilter} onValueChange={setFormTypeFilter}>
                <SelectTrigger>
                  <SelectValue placeholder="Filter by type" />
                </SelectTrigger>
                <SelectContent>
                  <SelectItem value="all">All Types</SelectItem>
                  <SelectItem value="regular">Regular</SelectItem>
                  <SelectItem value="table">Table</SelectItem>
                  <SelectItem value="mixed">Mixed</SelectItem>
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

          {/* Results Summary and Clear Filters */}
          <div className="flex items-center justify-between text-sm ">
            <span>Showing {Math.min(startIndex + 1, totalItems)}-{Math.min(endIndex, totalItems)} of {totalItems} submissions</span>
            <Button 
              variant="outline" 
              size="sm" 
              onClick={() => {
                setSearch("");
                setCategoryFilter('all');
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

      <Card>
        <CardHeader>
          <CardTitle>Submissions</CardTitle>
          <CardDescription>A total of {totalItems} submissions found for this user.</CardDescription>
        </CardHeader>
        <CardContent>
          <Table>
            <TableHeader>
              <TableRow>
                <TableHead>Form Title</TableHead>
                <TableHead>Category</TableHead>
                <TableHead>Form Type</TableHead>
                <TableHead>Ship</TableHead>
                <TableHead>Status</TableHead>
                <TableHead>Submitted At</TableHead>
                <TableHead>Expires At</TableHead>
                <TableHead>Actions</TableHead>
              </TableRow>
            </TableHeader>
            <TableBody>
              {submissions.length > 0 ? submissions.map((sub: any) => {
                // Calculate expiry info for display
                let expiryInfo = "N/A";
                let statusColor = "default";
                
                // Use existing expiry date if available, otherwise calculate
                if (sub.expiryDate) {
                  expiryInfo = formatDate(sub.expiryDate);
                } else if (sub.submittedAt && sub.form) {
                  const submissionDate = new Date(sub.submittedAt);
                  const validityPeriod = sub.form.validityPeriod || 30;
                  const expiryDate = new Date(submissionDate);
                  expiryDate.setDate(expiryDate.getDate() + validityPeriod);
                  expiryInfo = formatDate(expiryDate);
                }
                
                // Status color based on actual status
                if (sub.actualStatus === 'expired') {
                  statusColor = "destructive";
                } else if (sub.actualStatus === 'approved') {
                  statusColor = "secondary";
                } else if (sub.actualStatus === 'submitted') {
                  statusColor = "default";
                } else if (sub.actualStatus === 'rejected') {
                  statusColor = "destructive";
                } else if (sub.actualStatus === 'draft') {
                  statusColor = "outline";
                }
                
                return (
                  <TableRow key={sub._id}>
                    <TableCell className="font-medium">{sub.form?.title || "N/A"}</TableCell>
                    <TableCell>
                      <Badge variant="outline">
                        {(() => {
                          const category = sub.form?.category?.name || sub.form?.category;
                          return categoryInfo[category as keyof typeof categoryInfo] || category || "N/A";
                        })()}
                      </Badge>
                    </TableCell>
                    <TableCell>
                      <Badge variant="outline">
                        {(() => {
                          // First try to get form type from form data
                          let formType = sub.form?.formType || sub.form?.type || sub.form?.formStyle;
                          
                          // If no form type specified, try to infer from submission data
                          if (!formType && sub.data) {
                            const hasTableData = Object.keys(sub.data).some(key => 
                              key.startsWith('table_') || key === 'tableData' || 
                              (Array.isArray(sub.data[key]) && typeof sub.data[key][0] === 'object')
                            );
                            
                            const hasRegularFields = Object.keys(sub.data).some(key => 
                              !key.startsWith('table_') && key !== 'tableData' && 
                              !Array.isArray(sub.data[key])
                            );
                            
                            if (hasTableData && hasRegularFields) {
                              formType = 'mixed';
                            } else if (hasTableData) {
                              formType = 'table';
                            } else {
                              formType = 'regular';
                            }
                          }
                          
                          // Default to regular if still no type
                          formType = formType || 'regular';
                          
                          // Capitalize first letter for display
                          return String(formType).charAt(0).toUpperCase() + String(formType).slice(1).toLowerCase();
                        })()}
                      </Badge>
                    </TableCell>
                    <TableCell>{sub.ship?.name || "N/A"}</TableCell>
                    <TableCell>
                      <Badge variant={statusColor as any}>
                        {sub.actualStatus || sub.status}
                      </Badge>
                    </TableCell>
                    <TableCell>{sub.submittedAt ? formatDate(sub.submittedAt) : "N/A"}</TableCell>
                    <TableCell>{expiryInfo}</TableCell>
                    <TableCell>
                      <Button 
                        variant="outline" 
                        size="sm"
                        onClick={() => {
                          setSelectedSubmissionId(sub._id);
                          setIsModalOpen(true);
                        }}
                      >
                        <Eye className="w-4 h-4 mr-2" /> View
                      </Button>
                    </TableCell>
                  </TableRow>
                );
              }) : (
                <TableRow>
                  <TableCell colSpan={8} className="h-24 text-center">No submissions found.</TableCell>
                </TableRow>
              )}
            </TableBody>
          </Table>

          {/* Pagination */}
          {totalPages > 1 && (
            <div className="flex items-center justify-between mt-6">
              <div className="text-sm ">
                Page {currentPage} of {totalPages}
              </div>
              <div className="flex items-center gap-2">
                <Button
                  variant="outline"
                  size="sm"
                  onClick={() => setCurrentPage(prev => Math.max(1, prev - 1))}
                  disabled={currentPage === 1}
                >
                  <ChevronLeft className="w-4 h-4 mr-1" />
                  Previous
                </Button>
                <Button
                  variant="outline"
                  size="sm"
                  onClick={() => setCurrentPage(prev => Math.min(totalPages, prev + 1))}
                  disabled={currentPage === totalPages}
                >
                  Next
                  <ChevronRight className="w-4 h-4 ml-1" />
                </Button>
              </div>
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

export default UserSubmissions;