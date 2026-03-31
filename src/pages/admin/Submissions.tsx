import React, { useState, useMemo, useEffect } from "react";
import { useQuery, useMutation, useQueryClient } from "@tanstack/react-query";
import { useSearchParams } from "react-router-dom";
import api from "../../api";
import { Button } from "../../components/ui/button";
import {
  Card,
  CardContent,
  CardHeader,
  CardTitle,
} from "../../components/ui/card";
import { Input } from "../../components/ui/input";
import {
  Select,
  SelectContent,
  SelectItem,
  SelectTrigger,
  SelectValue,
} from "../../components/ui/select";
import { Badge } from "../../components/ui/badge";
import {
  Table,
  TableBody,
  TableCell,
  TableHead,
  TableHeader,
  TableRow,
} from "../../components/ui/table";
import {
  Search,
  Eye,
  FileText,
  Clock,
  CheckCircle,
  XCircle,
  AlertCircle,
  Filter,
  ChevronLeft,
  ChevronRight,
  Check,
  X,
} from "lucide-react";
import SubmissionViewModal from "../../components/SubmissionViewModal";
import { useClientSearch } from "../../hooks/useDebounce";
import { useToast } from "../../components/ui/toast";

const AdminSubmissions: React.FC = () => {
  const [searchParams] = useSearchParams();
  const selectedFormId = searchParams.get("form");

  const [search, setSearch] = useState("");
  const [categoryFilter, setCategoryFilter] = useState("all");
  const [formTypeFilter, setFormTypeFilter] = useState("all");
  const [statusFilter, setStatusFilter] = useState("all");
  const [startDate, setStartDate] = useState("");
  const [endDate, setEndDate] = useState("");
  const [currentPage, setCurrentPage] = useState(1);
  const [pageSize, setPageSize] = useState(20);
  const [selectedSubmissionId, setSelectedSubmissionId] = useState<
    string | null
  >(null);
  const [isModalOpen, setIsModalOpen] = useState(false);

  const queryClient = useQueryClient();
  const { toast } = useToast();

  const { data: submissionsData, isLoading } = useQuery({
    queryKey: ["admin-submissions"],
    queryFn: () =>
      api.getSubmissions({
        page: 1,
        limit: 1000, // Get all submissions for client-side filtering
      }),
  });

  const reviewMutation = useMutation({
    mutationFn: ({
      submissionId,
      status,
      reviewComments,
    }: {
      submissionId: string;
      status: string;
      reviewComments?: string;
    }) => api.reviewSubmission(submissionId, { status, reviewComments }),
    onSuccess: () => {
      queryClient.invalidateQueries({ queryKey: ["admin-submissions"] });
      toast({
        title: "Success",
        description: "Submission status updated successfully",
      });
    },
    onError: () => {
      toast({
        title: "Error",
        description: "Failed to update submission status",
        variant: "destructive",
      });
    },
  });

  // Apply client-side search - must be called before any conditional returns
  const rawSubmissions = submissionsData?.data?.data || [];
  const searchFields = ["form.title", "user.name", "user.email", "ship.name"];
  const searchedSubmissions = useClientSearch(
    rawSubmissions,
    search,
    searchFields,
  );

  // Category mapping for filtering
  const categoryInfo = {
    eng: "Engine Forms",
    deck: "Deck Forms",
    mlc: "MLC Forms",
    isps: "ISPS Forms",
    drill: "Drill Forms",
  };

  // Apply filters
  const filteredSubmissions = useMemo(() => {
    let filtered = searchedSubmissions;

    // Scope to a specific form when opened from "View Submissions"
    if (selectedFormId) {
      filtered = filtered.filter((sub: any) => {
        const formId = typeof sub.form === "string" ? sub.form : sub.form?._id;
        return formId === selectedFormId;
      });
    }

    // Apply category filter
    if (categoryFilter !== "all") {
      filtered = filtered.filter(
        (sub: any) =>
          sub.form &&
          (sub.form.category === categoryFilter ||
            sub.form.category?.name === categoryFilter),
      );
    }

    // Apply status filter
    if (statusFilter !== "all") {
      filtered = filtered.filter((sub: any) => sub.status === statusFilter);
    }

    // Apply form type filter - Enhanced with intelligent type detection
    if (formTypeFilter !== "all") {
      filtered = filtered.filter((sub: any) => {
        if (!sub.form && !sub.data) return false;

        // Get the form type using intelligent detection
        let formType =
          sub.form?.formType || sub.form?.type || sub.form?.formStyle;

        // If no form type specified, infer from submission data
        if (!formType && sub.data) {
          const hasTableData = Object.keys(sub.data).some(
            (key) =>
              key.startsWith("table_") ||
              key === "tableData" ||
              (Array.isArray(sub.data[key]) &&
                typeof sub.data[key][0] === "object"),
          );

          const hasRegularFields = Object.keys(sub.data).some(
            (key) =>
              !key.startsWith("table_") &&
              key !== "tableData" &&
              !Array.isArray(sub.data[key]),
          );

          if (hasTableData && hasRegularFields) {
            formType = "mixed";
          } else if (hasTableData) {
            formType = "table";
          } else {
            formType = "regular";
          }
        }

        // Default to regular if still no type
        formType = formType || "regular";

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
        const start = startDate ? new Date(startDate) : new Date("1970-01-01");
        const end = endDate ? new Date(endDate) : new Date("2099-12-31");

        return subDate >= start && subDate <= end;
      });
    }

    return filtered;
  }, [
    searchedSubmissions,
    selectedFormId,
    categoryFilter,
    formTypeFilter,
    startDate,
    endDate,
    statusFilter,
  ]);

  // Reset to first page when filters change
  useEffect(() => {
    setCurrentPage(1);
  }, [
    search,
    categoryFilter,
    formTypeFilter,
    statusFilter,
    startDate,
    endDate,
  ]);

  const getStatusBadge = (status: string) => {
    const variants = {
      draft: { variant: "secondary" as const, icon: Clock, label: "Draft" },
      pending: { variant: "secondary" as const, icon: Clock, label: "Pending" },
      submitted: {
        variant: "default" as const,
        icon: FileText,
        label: "Submitted",
      },
      approved: {
        variant: "default" as const,
        icon: CheckCircle,
        label: "Approved",
      },
      rejected: {
        variant: "destructive" as const,
        icon: XCircle,
        label: "Rejected",
      },
      expired: {
        variant: "destructive" as const,
        icon: AlertCircle,
        label: "Expired",
      },
    };

    const config =
      variants[status as keyof typeof variants] || variants.pending;
    const Icon = config.icon;

    return (
      <Badge variant={config.variant} className="flex items-center gap-1">
        <Icon className="w-3 h-3" />
        {config.label}
      </Badge>
    );
  };

  const handleApproveSubmission = (submissionId: string) => {
    reviewMutation.mutate({ submissionId, status: "approved" });
  };

  const handleRejectSubmission = (submissionId: string) => {
    const reviewComments = prompt(
      "Please provide a reason for rejection (optional):",
    );
    reviewMutation.mutate({
      submissionId,
      status: "rejected",
      reviewComments: reviewComments || undefined,
    });
  };

  const formatDate = (dateString: string) => {
    return new Date(dateString).toLocaleDateString("en-US", {
      year: "numeric",
      month: "short",
      day: "numeric",
      hour: "2-digit",
      minute: "2-digit",
    });
  };

  const handleViewSubmission = (submissionId: string) => {
    setSelectedSubmissionId(submissionId);
    setIsModalOpen(true);
  };

  if (isLoading) {
    return (
      <div className="flex items-center justify-center h-64">
        <div className="w-8 h-8 border-b-2 border-blue-600 rounded-full animate-spin"></div>
      </div>
    );
  }

  // Pagination for filtered results
  const totalItems = filteredSubmissions.length;
  const totalPages = Math.ceil(totalItems / pageSize);
  const startIndex = (currentPage - 1) * pageSize;
  const endIndex = Math.min(startIndex + pageSize, totalItems);
  const submissions = filteredSubmissions.slice(startIndex, endIndex);

  return (
    <div className="space-y-4 sm:space-y-6">
      <div className="flex flex-col gap-4 sm:flex-row sm:items-center sm:justify-between">
        <div>
          <h1 className="text-xl font-bold sm:text-2xl text-sena-navy dark:text-white">
            Form Submissions
          </h1>
          <p className="text-sena-lightBlue dark:text-white/90">
            Manage and review all form submissions
          </p>
        </div>
      </div>

      {/* Filters */}
      <Card className="border-sena-lightBlue/20 dark:border-gray-700">
        <CardHeader>
          <CardTitle className="flex items-center gap-2 text-sena-navy">
            <Filter className="w-5 h-5" />
            Filters & Search
          </CardTitle>
        </CardHeader>
        <CardContent className="space-y-4">
          {/* Search */}
          <div className="relative">
            <Search className="absolute w-4 h-4 left-3 top-3" />
            <Input
              placeholder="Search by form, user, or ship..."
              value={search}
              onChange={(e) => setSearch(e.target.value)}
              className="pl-9 border-sena-lightBlue/30 focus:border-sena-gold focus:ring-sena-gold/20"
            />
          </div>

          {/* Filter Controls */}
          <div className="grid grid-cols-1 gap-4 sm:grid-cols-2 lg:grid-cols-5">
            {/* Category Filter */}
            <div>
              <label className="block mb-2 text-sm font-medium text-sena-navy">
                Category
              </label>
              <Select value={categoryFilter} onValueChange={setCategoryFilter}>
                <SelectTrigger className="border-sena-lightBlue/30 focus:border-sena-gold focus:ring-sena-gold/20">
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
              <label className="block mb-2 text-sm font-medium text-sena-navy">
                Form Type
              </label>
              <Select value={formTypeFilter} onValueChange={setFormTypeFilter}>
                <SelectTrigger className="border-sena-lightBlue/30 focus:border-sena-gold focus:ring-sena-gold/20">
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

            {/* Status Filter */}
            <div>
              <label className="block mb-2 text-sm font-medium text-sena-navy">
                Status
              </label>
              <Select value={statusFilter} onValueChange={setStatusFilter}>
                <SelectTrigger className="border-sena-lightBlue/30 focus:border-sena-gold focus:ring-sena-gold/20">
                  <SelectValue placeholder="Filter by status" />
                </SelectTrigger>
                <SelectContent>
                  <SelectItem value="all">All Status</SelectItem>
                  <SelectItem value="pending">Pending</SelectItem>
                  <SelectItem value="approved">Approved</SelectItem>
                  <SelectItem value="rejected">Rejected</SelectItem>
                  <SelectItem value="expired">Expired</SelectItem>
                </SelectContent>
              </Select>
            </div>

            {/* Start Date Filter */}
            <div>
              <label className="block mb-2 text-sm font-medium text-sena-navy">
                From Date
              </label>
              <Input
                type="date"
                value={startDate}
                onChange={(e) => setStartDate(e.target.value)}
                className="border-sena-lightBlue/30 focus:border-sena-gold focus:ring-sena-gold/20"
              />
            </div>

            {/* End Date Filter */}
            <div>
              <label className="block mb-2 text-sm font-medium text-sena-navy">
                To Date
              </label>
              <Input
                type="date"
                value={endDate}
                onChange={(e) => setEndDate(e.target.value)}
                className="border-sena-lightBlue/30 focus:border-sena-gold focus:ring-sena-gold/20"
              />
            </div>
          </div>

          {/* Results Summary and Clear Filters */}
          <div className="flex flex-col gap-2 text-sm sm:flex-row sm:items-center sm:justify-between">
            <span>
              Showing {Math.min(startIndex + 1, totalItems)}-
              {Math.min(endIndex, totalItems)} of {totalItems} submissions
            </span>
            <Button
              variant="outline"
              size="sm"
              onClick={() => {
                setSearch("");
                setCategoryFilter("all");
                setFormTypeFilter("all");
                setStatusFilter("all");
                setStartDate("");
                setEndDate("");
                setCurrentPage(1);
              }}
              className="self-start sm:self-auto"
            >
              Clear Filters
            </Button>
          </div>
        </CardContent>
      </Card>

      {/* Results Summary */}
      {totalItems > 0 && (
        <div className="flex items-center justify-between mb-4 text-sm">
          <div className="flex items-center space-x-4">
            <span>
              Showing {startIndex + 1} to {Math.min(endIndex, totalItems)} of{" "}
              {totalItems} submissions
            </span>
            {(search ||
              categoryFilter !== "all" ||
              formTypeFilter !== "all" ||
              statusFilter !== "all" ||
              startDate ||
              endDate) && (
              <div className="flex items-center space-x-2">
                <span>•</span>
                <span>Filters applied:</span>
                {search && (
                  <Badge variant="outline" className="text-xs">
                    Search: "{search}"
                  </Badge>
                )}
                {categoryFilter !== "all" && (
                  <Badge variant="outline" className="text-xs">
                    Category:{" "}
                    {categoryInfo[
                      categoryFilter as keyof typeof categoryInfo
                    ] || categoryFilter}
                  </Badge>
                )}
                {formTypeFilter !== "all" && (
                  <Badge variant="outline" className="text-xs">
                    Type: {formTypeFilter}
                  </Badge>
                )}
                {statusFilter !== "all" && (
                  <Badge variant="outline" className="text-xs">
                    Status: {statusFilter}
                  </Badge>
                )}
                {(startDate || endDate) && (
                  <Badge variant="outline" className="text-xs">
                    Date: {startDate || "..."} to {endDate || "..."}
                  </Badge>
                )}
              </div>
            )}
          </div>
          <Select
            value={pageSize.toString()}
            onValueChange={(value) => {
              setPageSize(parseInt(value));
              setCurrentPage(1);
            }}
          >
            <SelectTrigger className="w-32">
              <SelectValue />
            </SelectTrigger>
            <SelectContent>
              <SelectItem value="10">10 per page</SelectItem>
              <SelectItem value="20">20 per page</SelectItem>
              <SelectItem value="50">50 per page</SelectItem>
              <SelectItem value="100">100 per page</SelectItem>
            </SelectContent>
          </Select>
        </div>
      )}

      {/* Submissions Table */}
      <Card className="border-sena-lightBlue/20 dark:border-gray-700">
        <CardHeader>
          <CardTitle className="text-sena-navy">
            Submissions ({totalItems})
          </CardTitle>
        </CardHeader>
        <CardContent className="p-0">
          {submissions.length === 0 ? (
            <div className="px-6 py-12 text-center">
              <FileText className="w-12 h-12 mx-auto text-gray-400" />
              <h3 className="mt-4 text-lg font-medium text-sena-navy ">
                No submissions found
              </h3>
              <p className="mt-2 text-sena-lightBlue">
                {search ||
                categoryFilter !== "all" ||
                formTypeFilter !== "all" ||
                startDate ||
                endDate
                  ? "Try adjusting your filters"
                  : "No form submissions have been created yet."}
              </p>
            </div>
          ) : (
            <>
              {/* Mobile Cards View */}
              <div className="block lg:hidden">
                <div className="p-4 space-y-4">
                  {submissions.map((submission: any) => (
                    <Card
                      key={submission._id}
                      className="border-sena-lightBlue/20 dark:border-gray-600"
                    >
                      <CardContent className="p-4">
                        <div className="space-y-3">
                          <div className="flex items-start justify-between gap-2">
                            <h3 className="flex-1 font-medium text-sena-navy line-clamp-2">
                              {submission.form?.title || "N/A"}
                            </h3>
                            <Badge
                              variant={
                                submission.status === "completed"
                                  ? "default"
                                  : submission.status === "pending"
                                    ? "secondary"
                                    : "destructive"
                              }
                              className="shrink-0"
                            >
                              {submission.status}
                            </Badge>
                          </div>

                          <div className="grid grid-cols-2 gap-3 text-sm">
                            <div>
                              <span className="text-gray-600 dark:text-gray-300">
                                Category:
                              </span>
                              <div className="font-medium text-sena-navy ">
                                {submission.form?.category?.displayName ||
                                  "N/A"}
                              </div>
                            </div>
                            <div>
                              <span className="text-gray-600 dark:text-gray-300">
                                Type:
                              </span>
                              <div className="font-medium capitalize text-sena-navy ">
                                {submission.form?.formType || "N/A"}
                              </div>
                            </div>
                            <div>
                              <span className="text-gray-600 dark:text-gray-300">
                                User:
                              </span>
                              <div className="font-medium text-sena-navy ">
                                {submission.user?.name || "N/A"}
                              </div>
                            </div>
                            <div>
                              <span className="text-gray-600 dark:text-gray-300">
                                Ship:
                              </span>
                              <div className="font-medium text-sena-navy ">
                                {submission.ship?.name || "N/A"}
                              </div>
                            </div>
                          </div>

                          <div className="text-xs text-gray-500 dark:text-gray-400">
                            Submitted:{" "}
                            {new Date(
                              submission.submittedAt,
                            ).toLocaleDateString()}
                          </div>

                          <div className="space-y-2">
                            <Button
                              size="sm"
                              onClick={() => {
                                setSelectedSubmissionId(submission._id);
                                setIsModalOpen(true);
                              }}
                              className="w-full bg-sena-gold hover:bg-sena-gold/90"
                            >
                              <Eye className="w-4 h-4 mr-2" />
                              View Details
                            </Button>
                            {submission.status === "pending" && (
                              <div className="flex gap-2">
                                <Button
                                  size="sm"
                                  onClick={() =>
                                    handleApproveSubmission(submission._id)
                                  }
                                  disabled={reviewMutation.isPending}
                                  className="flex-1 bg-green-600 hover:bg-green-700 text-white"
                                >
                                  <Check className="w-4 h-4 mr-1" />
                                  Approve
                                </Button>
                                <Button
                                  size="sm"
                                  variant="destructive"
                                  onClick={() =>
                                    handleRejectSubmission(submission._id)
                                  }
                                  disabled={reviewMutation.isPending}
                                  className="flex-1"
                                >
                                  <X className="w-4 h-4 mr-1" />
                                  Reject
                                </Button>
                              </div>
                            )}
                          </div>
                        </div>
                      </CardContent>
                    </Card>
                  ))}
                </div>
              </div>

              {/* Desktop Table View */}
              <div className="hidden overflow-x-auto lg:block">
                <Table>
                  <TableHeader>
                    <TableRow>
                      <TableHead className="text-sena-navy">Form</TableHead>
                      <TableHead className="text-sena-navy">Category</TableHead>
                      <TableHead className="text-sena-navy">
                        Form Type
                      </TableHead>
                      <TableHead className="text-sena-navy">
                        Submitted By
                      </TableHead>
                      <TableHead className="text-sena-navy">Ship</TableHead>
                      <TableHead className="text-sena-navy">Status</TableHead>
                      <TableHead className="text-sena-navy">
                        Submitted
                      </TableHead>
                      <TableHead className="text-sena-navy">Expires</TableHead>
                      <TableHead className="text-sena-navy">Actions</TableHead>
                    </TableRow>
                  </TableHeader>
                  <TableBody>
                    {submissions.map((submission: any) => (
                      <TableRow key={submission._id}>
                        <TableCell className="font-medium">
                          {submission.form?.title || "N/A"}
                        </TableCell>
                        <TableCell>
                          <Badge variant="outline">
                            {(() => {
                              const category =
                                submission.form?.category?.name ||
                                submission.form?.category;
                              return (
                                categoryInfo[
                                  category as keyof typeof categoryInfo
                                ] ||
                                category ||
                                "N/A"
                              );
                            })()}
                          </Badge>
                        </TableCell>
                        <TableCell>
                          <Badge variant="outline">
                            {(() => {
                              // First try to get form type from form data
                              let formType =
                                submission.form?.formType ||
                                submission.form?.type ||
                                submission.form?.formStyle;

                              // If no form type specified, try to infer from submission data
                              if (!formType && submission.data) {
                                const hasTableData = Object.keys(
                                  submission.data,
                                ).some(
                                  (key) =>
                                    key.startsWith("table_") ||
                                    key === "tableData" ||
                                    (Array.isArray(submission.data[key]) &&
                                      typeof submission.data[key][0] ===
                                        "object"),
                                );

                                const hasRegularFields = Object.keys(
                                  submission.data,
                                ).some(
                                  (key) =>
                                    !key.startsWith("table_") &&
                                    key !== "tableData" &&
                                    !Array.isArray(submission.data[key]),
                                );

                                if (hasTableData && hasRegularFields) {
                                  formType = "mixed";
                                } else if (hasTableData) {
                                  formType = "table";
                                } else {
                                  formType = "regular";
                                }
                              }

                              // Default to regular if still no type
                              formType = formType || "regular";

                              // Capitalize first letter for display
                              return (
                                String(formType).charAt(0).toUpperCase() +
                                String(formType).slice(1).toLowerCase()
                              );
                            })()}
                          </Badge>
                        </TableCell>
                        <TableCell>
                          <div>
                            <div className="font-medium">
                              {submission.user?.name}
                            </div>
                            <div className="text-sm text-gray-500">
                              {submission.user?.email}
                            </div>
                          </div>
                        </TableCell>
                        <TableCell>
                          <div>
                            <div className="font-medium">
                              {submission.ship?.name}
                            </div>
                            <div className="text-sm text-gray-500">
                              IMO: {submission.ship?.imoNumber}
                            </div>
                          </div>
                        </TableCell>
                        <TableCell>
                          {getStatusBadge(submission.status)}
                        </TableCell>
                        <TableCell>
                          {submission.submittedAt
                            ? formatDate(submission.submittedAt)
                            : "Not submitted"}
                        </TableCell>
                        <TableCell>
                          <div className="text-sm">
                            {formatDate(submission.expiryDate)}
                            {submission.isExpired && (
                              <div className="font-medium text-red-500">
                                Expired
                              </div>
                            )}
                          </div>
                        </TableCell>
                        <TableCell>
                          <div className="flex items-center gap-2">
                            <Button
                              variant="outline"
                              size="sm"
                              onClick={() =>
                                handleViewSubmission(submission._id)
                              }
                              className="flex items-center gap-1"
                            >
                              <Eye className="w-3 h-3" />
                              View
                            </Button>
                            {submission.status === "pending" && (
                              <>
                                <Button
                                  size="sm"
                                  onClick={() =>
                                    handleApproveSubmission(submission._id)
                                  }
                                  disabled={reviewMutation.isPending}
                                  className="bg-green-600 hover:bg-green-700 text-white"
                                >
                                  <Check className="w-3 h-3 mr-1" />
                                  Approve
                                </Button>
                                <Button
                                  size="sm"
                                  variant="destructive"
                                  onClick={() =>
                                    handleRejectSubmission(submission._id)
                                  }
                                  disabled={reviewMutation.isPending}
                                >
                                  <X className="w-3 h-3 mr-1" />
                                  Reject
                                </Button>
                              </>
                            )}
                          </div>
                        </TableCell>
                      </TableRow>
                    ))}
                  </TableBody>
                </Table>
              </div>
            </>
          )}
        </CardContent>
      </Card>

      {/* Pagination */}
      {totalPages > 1 && (
        <div className="flex flex-col items-center justify-center gap-4 mt-6 sm:flex-row">
          <div className="flex items-center space-x-2">
            <Button
              variant="outline"
              size="sm"
              onClick={() => setCurrentPage((prev) => Math.max(prev - 1, 1))}
              disabled={currentPage === 1}
              className="px-3"
            >
              <ChevronLeft className="w-4 h-4" />
              <span className="hidden ml-1 sm:inline">Previous</span>
            </Button>

            <div className="flex items-center space-x-1">
              {Array.from({ length: Math.min(3, totalPages) }, (_, i) => {
                const pageNumber =
                  Math.max(
                    1,
                    Math.min(totalPages - 2, Math.max(1, currentPage - 1)),
                  ) + i;

                if (pageNumber > totalPages) return null;

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
              })}
            </div>

            <Button
              variant="outline"
              size="sm"
              onClick={() =>
                setCurrentPage((prev) => Math.min(prev + 1, totalPages))
              }
              disabled={currentPage === totalPages}
              className="px-3"
            >
              <span className="hidden mr-1 sm:inline">Next</span>
              <ChevronRight className="w-4 h-4" />
            </Button>
          </div>
          <div className="text-sm">
            Page {currentPage} of {totalPages}
          </div>
        </div>
      )}

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

export default AdminSubmissions;
