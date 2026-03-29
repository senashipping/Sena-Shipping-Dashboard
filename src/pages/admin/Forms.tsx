import React, { useState } from "react";
import { useQuery, useMutation } from "@tanstack/react-query";
import api from "../../api";
import { Card, CardContent, CardHeader, CardTitle } from "../../components/ui/card";
import { Button } from "../../components/ui/button";
import { Badge } from "../../components/ui/badge";
import { Select, SelectContent, SelectItem, SelectTrigger, SelectValue } from "../../components/ui/select";
import { Table, TableBody, TableCell, TableHead, TableHeader, TableRow } from "../../components/ui/table";
import { Pagination } from "../../components/ui/pagination";
import { FileText, Settings, AlertTriangle, CheckCircle, Clock, XCircle, Trash2 } from "lucide-react";
import { ConfirmationDialog } from "../../components/ui/confirmation-dialog";
import { Form } from "../../types";
import { formatDateTime } from "../../lib/utils";
import { useAuth } from "../../contexts/AuthContext";
import { useToast } from "../../components/ui/toast";

const AdminForms: React.FC = () => {
  const { isSuperAdmin } = useAuth();
  const { toast } = useToast();
  const [selectedCategory, setSelectedCategory] = useState<string>("all");
  const [currentPage, setCurrentPage] = useState<number>(1);
  const [itemsPerPage] = useState<number>(25);
  const [deleteConfirmation, setDeleteConfirmation] = useState<{
    isOpen: boolean;
    form: any | null;
  }>({ isOpen: false, form: null });

  const { data: forms, isLoading: formsLoading, refetch: refetchForms } = useQuery({
    queryKey: ["admin-forms", selectedCategory, currentPage, itemsPerPage],
    queryFn: () => {
      const params = {
        page: currentPage,
        limit: itemsPerPage,
      };
      
      if (selectedCategory === "all") {
        return api.getForms(params);
      }
      return api.getFormsByCategory(selectedCategory, params);
    },
  });

  const { data: categories } = useQuery({
    queryKey: ["categories"],
    queryFn: () => api.getCategories(),
  });

  // Reset to first page when category changes
  const handleCategoryChange = (value: string) => {
    setSelectedCategory(value);
    setCurrentPage(1);
  };



  const triggerFormStatusNotificationsMutation = useMutation({
    mutationFn: () => api.triggerFormStatusCheck(),
    onSuccess: () => {
      toast({
        title: "Success",
        description: "Form status notifications triggered successfully",
        variant: "success",
      });
    },
    onError: (error: any) => {
      toast({
        title: "Error",
        description: error.response?.data?.message || "Failed to trigger notifications",
        variant: "destructive",
      });
    },
  });

  const deleteFormMutation = useMutation({
    mutationFn: (id: string) => api.deleteForm(id),
    onSuccess: () => {
      toast({
        title: "Success",
        description: "Form deleted successfully",
        variant: "success",
      });
      setDeleteConfirmation({ isOpen: false, form: null });
      refetchForms(); // Refetch forms after deletion
    },
    onError: (error: any) => {
      toast({
        title: "Error",
        description: error.response?.data?.message || "Failed to delete form",
        variant: "destructive",
      });
    },
  });

  const handleDeleteForm = (form: any) => {
    setDeleteConfirmation({ isOpen: true, form });
  };

  const confirmDeleteForm = () => {
    if (deleteConfirmation.form) {
      deleteFormMutation.mutate(deleteConfirmation.form._id);
    }
  };



  const getStatusIcon = (status: string) => {
    switch (status) {
      case "active":
        return <CheckCircle className="w-4 h-4 text-green-600" />;
      case "expiring-soon":
        return <Clock className="w-4 h-4 text-yellow-600" />;
      case "expired":
        return <AlertTriangle className="w-4 h-4 text-red-600" />;
      case "inactive":
        return <XCircle className="w-4 h-4 text-gray-600" />;
      default:
        return <FileText className="w-4 h-4" />;
    }
  };

  const getStatusBadgeVariant = (status: string) => {
    switch (status) {
      case "active":
        return "default";
      case "expiring-soon":
        return "secondary";
      case "expired":
        return "destructive";
      case "inactive":
        return "outline";
      default:
        return "outline";
    }
  };

  const formsData = forms?.data?.data || [];
  const pagination = forms?.data?.pagination || {
    currentPage: 1,
    totalPages: 1,
    totalItems: 0,
    itemsPerPage: itemsPerPage,
    hasNextPage: false,
    hasPrevPage: false,
  };
  const categoriesData = categories?.data?.data || [];

  if (formsLoading) {
    return (
      <div className="space-y-6">
        <div className="animate-pulse">
          <div className="w-1/4 h-8 mb-4 bg-gray-200 rounded"></div>
          <div className="space-y-4">
            {[...Array(5)].map((_, i) => (
              <div key={i} className="h-16 bg-gray-200 rounded"></div>
            ))}
          </div>
        </div>
      </div>
    );
  }

  return (
    <div className="space-y-6">
      <div className="flex flex-col gap-4 sm:flex-row sm:items-center sm:justify-between">
        <div>
          <h1 className="text-3xl font-bold">Form Management</h1>
          <p className="mt-1">
            Manage form statuses and trigger notifications
          </p>
        </div>
        <div className="flex items-center gap-2">
          <Button
            onClick={() => triggerFormStatusNotificationsMutation.mutate()}
            disabled={triggerFormStatusNotificationsMutation.isPending}
            variant="outline"
          >
            <Settings className="w-4 h-4 mr-2" />
            {triggerFormStatusNotificationsMutation.isPending ? "Triggering..." : "Trigger Notifications"}
          </Button>
        </div>
      </div>

      {/* Filters */}
      <Card>
        <CardContent className="p-4">
          <div className="flex items-center gap-4">
            <label className="text-sm font-medium">Category:</label>
            <Select value={selectedCategory} onValueChange={handleCategoryChange}>
              <SelectTrigger className="w-48">
                <SelectValue placeholder="Select category" />
              </SelectTrigger>
              <SelectContent>
                <SelectItem value="all">All Categories</SelectItem>
                {categoriesData.map((category: any) => (
                  <SelectItem key={category.name} value={category.name}>
                    {category.displayName}
                  </SelectItem>
                ))}
              </SelectContent>
            </Select>
          </div>
        </CardContent>
      </Card>

      {/* Forms Table */}
      <Card>
        <CardHeader>
          <CardTitle>Forms ({pagination.totalItems})</CardTitle>
        </CardHeader>
        <CardContent>
          {formsData.length === 0 ? (
            <div className="py-8 text-center">
              <FileText className="w-12 mx-auto mb-4 -h-12 text-sena-lightBlue" />
              <p className="text-sena-lightBlue">No forms found</p>
            </div>
          ) : (
            <>
              <div className="overflow-x-auto">
                <Table>
                  <TableHeader>
                    <TableRow>
                      <TableHead>Form Name</TableHead>
                      <TableHead>Category</TableHead>
                      <TableHead>Current Status</TableHead>
                      <TableHead>Validity Period</TableHead>
                      <TableHead>Created</TableHead>
                      <TableHead>Note</TableHead>
                      {isSuperAdmin() && <TableHead>Actions</TableHead>}
                    </TableRow>
                  </TableHeader>
                  <TableBody>
                    {formsData.map((form: Form) => (
                      <TableRow key={form._id}>
                        <TableCell className="font-medium">{form.title}</TableCell>
                        <TableCell>
                          <Badge variant="outline" className="text-xs">
                            {typeof form.category === 'object' ? form.category.displayName : form.category}
                          </Badge>
                        </TableCell>
                        <TableCell>
                          <div className="flex items-center gap-2">
                            {getStatusIcon(form.status)}
                            <Badge variant={getStatusBadgeVariant(form.status)}>
                              {form.status.replace('-', ' ')}
                            </Badge>
                          </div>
                        </TableCell>
                        <TableCell>{form.validityPeriod} days</TableCell>
                        <TableCell className="text-sm">
                          {formatDateTime(form.createdAt)}
                        </TableCell>
                        <TableCell>
                          <div className="text-sm">
                            Status calculated automatically based on submissions
                          </div>
                        </TableCell>
                        {isSuperAdmin() && (
                          <TableCell>
                            <Button 
                              variant="destructive" 
                              size="sm" 
                              onClick={() => handleDeleteForm(form)}
                              title="Delete form (Super Admin Only)"
                            >
                              <Trash2 className="w-4 h-4" />
                            </Button>
                          </TableCell>
                        )}
                      </TableRow>
                    ))}
                  </TableBody>
                </Table>
              </div>
              
              {/* Pagination */}
              {pagination.totalPages > 1 && (
                <div className="pt-4 mt-4 border-t">
                  <Pagination
                    currentPage={pagination.currentPage}
                    totalPages={pagination.totalPages}
                    totalItems={pagination.totalItems}
                    itemsPerPage={pagination.itemsPerPage}
                    onPageChange={setCurrentPage}
                    hasNextPage={pagination.hasNextPage}
                    hasPrevPage={pagination.hasPrevPage}
                  />
                </div>
              )}
            </>
          )}
        </CardContent>
      </Card>

      {/* Delete Confirmation Dialog */}
      <ConfirmationDialog
        isOpen={deleteConfirmation.isOpen}
        onClose={() => setDeleteConfirmation({ isOpen: false, form: null })}
        onConfirm={confirmDeleteForm}
        title="Delete Form"
        description={`Are you sure you want to delete "${deleteConfirmation.form?.title}"? This action cannot be undone. Note: Forms with existing submissions cannot be deleted.`}
        confirmText="Delete Form"
        isLoading={deleteFormMutation.isPending}
      />
    </div>
  );
};

export default AdminForms;