"use client";

import React, { useState, useMemo } from "react";
import { useQuery, useMutation, useQueryClient } from "@tanstack/react-query";
import { Link } from "react-router-dom";
import api from "../../api";
import { Button } from "../../components/ui/button";
import { Card, CardContent, CardDescription, CardHeader, CardTitle } from "../../components/ui/card";
import { Badge } from "../../components/ui/badge";
import { Input } from "../../components/ui/input";
import { Select, SelectContent, SelectItem, SelectTrigger, SelectValue } from "../../components/ui/select";
import { Plus, Search, ChevronLeft, ChevronRight, Trash2 } from "lucide-react";
import { useAuth } from "../../contexts/AuthContext";
import { useClientSearch } from "../../hooks/useDebounce";
import { ConfirmationDialog } from "../../components/ui/confirmation-dialog";
import { useToast } from "../../components/ui/toast";

const FormList: React.FC = () => {
  const [searchTerm, setSearchTerm] = useState("");
  const [categoryFilter, setCategoryFilter] = useState("all");
  const [currentPage, setCurrentPage] = useState(1);
  const [itemsPerPage, setItemsPerPage] = useState(9); // 3x3 grid
  
  // New filter states matching CategoryForms
  const [formTypeFilter, setFormTypeFilter] = useState<'all' | 'regular' | 'table' | 'mixed'>('all');
  const [startDate, setStartDate] = useState("");
  const [endDate, setEndDate] = useState("");
  
  // Delete confirmation state
  const [deleteConfirmation, setDeleteConfirmation] = useState<{
    isOpen: boolean;
    form: any | null;
    submissionCount?: number;
    isLoadingStats?: boolean;
  }>({ isOpen: false, form: null, submissionCount: 0, isLoadingStats: false });
  
  // const [searchParams] = useSearchParams(); // Future use
  const { auth, isSuperAdmin, canAccessCategory } = useAuth();
  const queryClient = useQueryClient();
  const { toast } = useToast();

  // const initialCategory = searchParams.get("category"); // Future use
  const isAdmin = auth.user?.role === "admin" || auth.user?.role === "super_admin";

  const { data: categories, isLoading: categoriesLoading } = useQuery({
    queryKey: ["categories"],
    queryFn: () => api.getCategories({ isActive: true }),
  });

  const { data: forms, isLoading } = useQuery({
    queryKey: ["forms"],
    queryFn: () => api.getForms({}),
  });

  // Delete form mutation
  const deleteFormMutation = useMutation({
    mutationFn: (id: string) => api.deleteForm(id),
    onSuccess: () => {
      toast({
        title: "Form Deleted",
        description: "Form deleted successfully",
        variant: "success",
      });
      queryClient.invalidateQueries({ queryKey: ["forms"] });
      setDeleteConfirmation({ isOpen: false, form: null });
    },
    onError: (error: any) => {
      toast({
        title: "Delete Failed",
        description: error.response?.data?.message || "Failed to delete form",
        variant: "destructive",
      });
    },
  });

  const handleDeleteForm = async (form: any) => {
    setDeleteConfirmation({ isOpen: true, form, submissionCount: 0, isLoadingStats: true });
    
    try {
      const statsResponse = await api.getFormStats(form._id);
      const submissionCount = statsResponse.data?.data?.submissionCount || 0;
      setDeleteConfirmation({ isOpen: true, form, submissionCount, isLoadingStats: false });
    } catch (error) {
      console.error('Failed to fetch form stats:', error);
      setDeleteConfirmation({ isOpen: true, form, submissionCount: 0, isLoadingStats: false });
    }
  };

  const confirmDeleteForm = () => {
    if (deleteConfirmation.form) {
      deleteFormMutation.mutate(deleteConfirmation.form._id);
    }
  };

  const rawForms = Array.isArray(forms?.data?.data) ? forms.data.data : [];
  
  // Apply client-side search and filtering
  const searchFields = ['title', 'description', 'category.displayName'];
  const searchedForms = useClientSearch(rawForms, searchTerm, searchFields);
  
  // Apply all filters
  const filteredForms = searchedForms.filter((form: any) => {
    // Category access filter - check if user can access this category
    const categoryAccessMatch = canAccessCategory(form.category.name);
    
    // Category filter
    const categoryMatch = categoryFilter === "all" || form.category._id === categoryFilter;
    
    // Form type filter
    const formTypeMatch = formTypeFilter === "all" || form.formType === formTypeFilter;
    
    // Date range filter (using form creation date)
    let dateMatch = true;
    if (startDate || endDate) {
      const formDate = new Date(form.createdAt);
      const start = startDate ? new Date(startDate) : new Date('1970-01-01');
      const end = endDate ? new Date(endDate) : new Date('2099-12-31');
      dateMatch = formDate >= start && formDate <= end;
    }
    
    return categoryAccessMatch && categoryMatch && formTypeMatch && dateMatch;
  });

  // Pagination logic
  const paginationInfo = useMemo(() => {
    const totalItems = filteredForms.length;
    const totalPages = Math.ceil(totalItems / itemsPerPage);
    const startIndex = (currentPage - 1) * itemsPerPage;
    const endIndex = startIndex + itemsPerPage;
    const paginatedForms = filteredForms.slice(startIndex, endIndex);

    return {
      totalItems,
      totalPages,
      currentPage,
      paginatedForms,
      hasNextPage: currentPage < totalPages,
      hasPrevPage: currentPage > 1,
    };
  }, [filteredForms, currentPage, itemsPerPage]);

  // Reset to first page when filters change
  React.useEffect(() => {
    setCurrentPage(1);
  }, [searchTerm, categoryFilter, formTypeFilter, startDate, endDate]);

  return (
    <div className="space-y-4 sm:space-y-6">
      <div className="flex flex-col gap-4 sm:flex-row sm:items-center sm:justify-between">
        <h1 className="text-2xl font-bold sm:text-3xl text-sena-navy dark:text-white">
          {isAdmin ? "Form Management" : "Available Forms"}
        </h1>
        {isAdmin && (
          <Link to="/admin/forms/new" className="self-start">
            <Button className="bg-sena-gold hover:bg-sena-gold/90">
              <Plus className="w-4 h-4 mr-2" />
              <span className="hidden sm:inline">Create Form</span>
              <span className="sm:hidden">Create</span>
            </Button>
          </Link>
        )}
      </div>

      {/* Filters */}
      <Card className="border-sena-lightBlue/20">
        <CardHeader>
          <CardTitle className="text-sena-navy">Filters</CardTitle>
        </CardHeader>
        <CardContent className="space-y-4">
          {/* Search */}
          <div className="relative">
            <Search className="absolute w-4 h-4 left-3 top-3 " />
            <Input
              placeholder="Search forms..."
              value={searchTerm}
              onChange={(e) => setSearchTerm(e.target.value)}
              className="pl-9 border-sena-lightBlue/30 focus:border-sena-gold focus:ring-sena-gold/20"
            />
          </div>

          {/* Filter Controls */}
          <div className="grid grid-cols-1 gap-4 sm:grid-cols-2 lg:grid-cols-4">
            {/* Category Filter */}
            <div>
              <label className="block mb-2 text-sm font-medium text-sena-navy">Category</label>
              <Select
                value={categoryFilter}
                onValueChange={setCategoryFilter}
                disabled={categoriesLoading}
              >
                <SelectTrigger className="border-sena-lightBlue/30 focus:border-sena-gold focus:ring-sena-gold/20">
                  <SelectValue placeholder={categoriesLoading ? "Loading..." : "All Categories"} />
                </SelectTrigger>
                <SelectContent>
                  <SelectItem value="all">All Categories</SelectItem>
                  {Array.isArray(categories?.data?.data) ? categories.data.data
                    .filter((category: any) => canAccessCategory(category.name))
                    .map((category: any) => (
                    <SelectItem key={category._id} value={category._id}>
                      {category.displayName}
                    </SelectItem>
                  )) : []}
                </SelectContent>
              </Select>
            </div>

            {/* Form Type Filter */}
            <div>
              <label className="block mb-2 text-sm font-medium text-sena-navy">Form Type</label>
              <Select value={formTypeFilter} onValueChange={(value: any) => setFormTypeFilter(value)}>
                <SelectTrigger className="border-sena-lightBlue/30 focus:border-sena-gold focus:ring-sena-gold/20">
                  <SelectValue placeholder="All Types" />
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
              <label className="block mb-2 text-sm font-medium text-sena-navy">From Date</label>
              <Input
                type="date"
                value={startDate}
                onChange={(e) => setStartDate(e.target.value)}
                className="border-sena-lightBlue/30 focus:border-sena-gold focus:ring-sena-gold/20"
              />
            </div>

            {/* End Date Filter */}
            <div>
              <label className="block mb-2 text-sm font-medium text-sena-navy">To Date</label>
              <Input
                type="date"
                value={endDate}
                onChange={(e) => setEndDate(e.target.value)}
                className="border-sena-lightBlue/30 focus:border-sena-gold focus:ring-sena-gold/20"
              />
            </div>
          </div>

          {/* Results Summary */}
          <div className="flex flex-col gap-2 text-sm sm:flex-row sm:items-center sm:justify-between ">
            <span>Found {filteredForms.length} forms</span>
            <Button 
              variant="outline" 
              size="sm" 
              onClick={() => {
                setSearchTerm("");
                setCategoryFilter("all");
                setFormTypeFilter("all");
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

      {/* Forms Grid */}
      {isLoading ? (
        <div className="flex items-center justify-center h-64">
          <div className="w-12 h-12 border-b-2 rounded-full animate-spin border-primary"></div>
        </div>
      ) : filteredForms.length === 0 ? (
        <Card>
          <CardContent className="flex items-center justify-center h-32">
            <p className="">No forms found</p>
          </CardContent>
        </Card>
      ) : (
        <>
          {/* Results Summary */}
          <div className="flex items-center justify-between mb-4 text-sm ">
            <div className="flex items-center space-x-4">
              <span>
                Showing {((currentPage - 1) * itemsPerPage) + 1} to {Math.min(currentPage * itemsPerPage, paginationInfo.totalItems)} of {paginationInfo.totalItems} forms
              </span>
              {(searchTerm || categoryFilter !== "all" || formTypeFilter !== "all" || startDate || endDate) && (
                <div className="flex items-center space-x-2">
                  <span>•</span>
                  <span>Filters applied:</span>
                  {searchTerm && <Badge variant="outline" className="text-xs">Search: "{searchTerm}"</Badge>}
                  {categoryFilter !== "all" && (
                    <Badge variant="outline" className="text-xs">
                      Category: {categories?.data?.data?.find((c: any) => c._id === categoryFilter)?.displayName || categoryFilter}
                    </Badge>
                  )}
                  {formTypeFilter !== "all" && (
                    <Badge variant="outline" className="text-xs">
                      Type: {formTypeFilter}
                    </Badge>
                  )}
                  {(startDate || endDate) && (
                    <Badge variant="outline" className="text-xs">
                      Date: {startDate || '...'} to {endDate || '...'}
                    </Badge>
                  )}
                </div>
              )}
            </div>
            <Select value={itemsPerPage.toString()} onValueChange={(value) => {
              setItemsPerPage(parseInt(value));
              setCurrentPage(1);
            }}>
              <SelectTrigger className="w-32">
                <SelectValue />
              </SelectTrigger>
              <SelectContent>
                <SelectItem value="6">6 per page</SelectItem>
                <SelectItem value="9">9 per page</SelectItem>
                <SelectItem value="12">12 per page</SelectItem>
                <SelectItem value="18">18 per page</SelectItem>
              </SelectContent>
            </Select>
          </div>

          <div className="grid grid-cols-1 gap-4 sm:gap-6 sm:grid-cols-2 lg:grid-cols-3">
            {paginationInfo.paginatedForms.map((form: any) => (
            <Card key={form._id} className="transition-shadow hover:shadow-md border-sena-lightBlue/20">
              <CardHeader className="pb-3">
                <div className="flex items-start justify-between gap-2">
                  <CardTitle className="text-base sm:text-lg text-sena-navy line-clamp-2">{form.title}</CardTitle>
                  <Badge variant={form.isActive ? "default" : "secondary"} className="text-xs shrink-0">
                    {form.isActive ? "Active" : "Inactive"}
                  </Badge>
                </div>
                <CardDescription className="text-sm text-sena-lightBlue line-clamp-2">{form.description}</CardDescription>
              </CardHeader>
              <CardContent className="pt-3">
                <div className="space-y-2">
                  <div className="flex items-center justify-between text-sm">
                    <span className="font-medium text-gray-600 dark:text-gray-300">Category:</span>
                    <Badge variant="outline" className="text-xs">{form.category.displayName}</Badge>
                  </div>
                  
                  <div className="flex items-center justify-between text-sm">
                    <span className="font-medium text-gray-600 dark:text-gray-300">Type:</span>
                    <Badge variant="outline" className="text-xs capitalize">{form.formType}</Badge>
                  </div>

                  <div className="flex items-center justify-between text-sm">
                    <span className="font-medium text-gray-600 dark:text-gray-300">Validity:</span>
                    <span className="text-sena-navy">{form.validityPeriod} days</span>
                  </div>

                  {form.formType === "regular" && (
                    <div className="flex items-center justify-between text-sm">
                      <span className="font-medium text-gray-600 dark:text-gray-300">Fields:</span>
                      <span className="text-sena-navy">{form.fields?.length || 0}</span>
                    </div>
                  )}

                  {form.formType === "table" && (
                    <div className="flex items-center justify-between text-sm">
                      <span className="font-medium text-gray-600 dark:text-gray-300">Columns:</span>
                      <span className="text-sena-navy">{form.tableConfig?.columns?.length || 0}</span>
                    </div>
                  )}
                </div>

                <div className="flex gap-2 mt-4">
                  {isAdmin ? (
                    <>
                      <Link to={`/admin/forms/${form._id}`} className="flex-1">
                        <Button variant="outline" size="sm" className="w-full">
                          <span className="hidden sm:inline">Edit</span>
                          <span className="sm:hidden">Edit</span>
                        </Button>
                      </Link>
                      <Link to={`/admin/submissions?form=${form._id}`} className="flex-1">
                        <Button size="sm" className="w-full bg-sena-gold hover:bg-sena-gold/90">
                          <span className="hidden sm:inline">View Submissions</span>
                          <span className="sm:hidden">Submissions</span>
                        </Button>
                      </Link>
                      {isSuperAdmin() && (
                        <Button 
                          variant="destructive" 
                          size="sm" 
                          onClick={(e) => {
                            e.preventDefault();
                            e.stopPropagation();
                            handleDeleteForm(form);
                          }}
                          title="Delete form (Super Admin Only)"
                          className="px-3"
                        >
                          <Trash2 className="w-4 h-4" />
                        </Button>
                      )}
                    </>
                  ) : (
                    <Link to={isAdmin ? `/admin/forms/${form._id}` : `/dashboard/forms/${form._id}`} className="w-full">
                      <Button size="sm" className="w-full bg-sena-gold hover:bg-sena-gold/90">
                        {form.formType === "table" ? "Fill Table" : "Fill Form"}
                      </Button>
                    </Link>
                  )}
                </div>
              </CardContent>
            </Card>
          ))}
          </div>

          {/* Pagination Controls */}
          {paginationInfo.totalPages > 1 && (
            <div className="flex flex-col items-center justify-center gap-4 mt-6 sm:flex-row">
              <div className="flex items-center space-x-2">
                <Button
                  variant="outline"
                  size="sm"
                  onClick={() => setCurrentPage(prev => Math.max(prev - 1, 1))}
                  disabled={!paginationInfo.hasPrevPage}
                  className="px-3"
                >
                  <ChevronLeft className="w-4 h-4" />
                  <span className="hidden ml-1 sm:inline">Previous</span>
                </Button>
                
                <div className="flex items-center space-x-1">
                  {Array.from({ length: Math.min(3, paginationInfo.totalPages) }, (_, i) => {
                    const pageNumber = Math.max(1, Math.min(
                      paginationInfo.totalPages - 2,
                      Math.max(1, currentPage - 1)
                    )) + i;
                    
                    if (pageNumber > paginationInfo.totalPages) return null;
                    
                    return (
                      <Button
                        key={pageNumber}
                        variant={pageNumber === currentPage ? "default" : "outline"}
                        size="sm"
                        onClick={() => setCurrentPage(pageNumber)}
                        className="w-8 px-0 sm:w-10"
                      >
                        {pageNumber}
                      </Button>
                    );
                  }).filter(Boolean)}
                </div>
                
                <Button
                  variant="outline"
                  size="sm"
                  onClick={() => setCurrentPage(prev => Math.min(prev + 1, paginationInfo.totalPages))}
                  disabled={!paginationInfo.hasNextPage}
                  className="px-3"
                >
                  <span className="hidden mr-1 sm:inline">Next</span>
                  <ChevronRight className="w-4 h-4" />
                </Button>
              </div>
              <div className="text-sm ">
                Page {currentPage} of {paginationInfo.totalPages}
              </div>
            </div>
          )}
        </>
      )}

      {/* Delete Confirmation Dialog */}
      <ConfirmationDialog
        isOpen={deleteConfirmation.isOpen}
        onClose={() => setDeleteConfirmation({ isOpen: false, form: null })}
        onConfirm={confirmDeleteForm}
        title="Delete Form"
        description={
          deleteConfirmation.isLoadingStats 
            ? `Loading form statistics...`
            : (deleteConfirmation.submissionCount || 0) > 0
              ? `Are you sure you want to delete "${deleteConfirmation.form?.title}"? This action cannot be undone and will permanently delete ${deleteConfirmation.submissionCount} submission${(deleteConfirmation.submissionCount || 0) > 1 ? 's' : ''} associated with this form.`
              : `Are you sure you want to delete "${deleteConfirmation.form?.title}"? This action cannot be undone.`
        }
        confirmText={
          (deleteConfirmation.submissionCount || 0) > 0 
            ? `Delete Form & ${deleteConfirmation.submissionCount} Submission${(deleteConfirmation.submissionCount || 0) > 1 ? 's' : ''}` 
            : "Delete Form"
        }
        isLoading={deleteFormMutation.isPending || deleteConfirmation.isLoadingStats}
      />
    </div>
  );
};

export default FormList;