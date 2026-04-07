import React from "react";
import { useQuery } from "@tanstack/react-query";
import api from "../api";
import {
  Dialog,
  DialogContent,
  DialogDescription,
  DialogHeader,
  DialogTitle,
} from "./ui/dialog";
import { Badge } from "./ui/badge";
import { Card, CardContent, CardDescription, CardHeader, CardTitle } from "./ui/card";
import { Table, TableBody, TableCell, TableHead, TableHeader, TableRow } from "./ui/table";
import { ScrollArea } from "./ui/scroll-area";
import { formatDate } from "../lib/utils";
import { FileText, User, Ship, Calendar, Clock, CheckCircle, XCircle, AlertCircle, FileDown } from "lucide-react";
import { Button } from "./ui/button";
import SubmissionPdfDocument from "./pdf/SubmissionPdfDocument";
import { downloadPdfDocument } from "../lib/pdfDownload";
import { useToast } from "./ui/toast";
import EmbeddedExcelWorkbook from "./form-builder/EmbeddedExcelWorkbook";

interface SubmissionViewModalProps {
  submissionId: string | null;
  isOpen: boolean;
  onClose: () => void;
}

const SubmissionViewModal: React.FC<SubmissionViewModalProps> = ({
  submissionId,
  isOpen,
  onClose,
}) => {
  const { toast } = useToast();
  const { data: submissionData, isLoading } = useQuery({
    queryKey: ["submission-detail", submissionId],
    queryFn: () => api.getSubmissionById(submissionId!),
    enabled: !!submissionId && isOpen,
  });

  const submission = submissionData?.data?.data;

  if (!submission) {
    return (
      <Dialog open={isOpen} onOpenChange={onClose}>
        <DialogContent className="max-w-4xl max-h-[90vh]">
          <DialogHeader>
            <DialogTitle>View Submission</DialogTitle>
            <DialogDescription>
              {isLoading ? "Loading submission details..." : "Submission not found"}
            </DialogDescription>
          </DialogHeader>
          {isLoading && (
            <div className="flex items-center justify-center h-32">
              <div className="w-8 h-8 border-b-2 rounded-full animate-spin border-primary"></div>
            </div>
          )}
        </DialogContent>
      </Dialog>
    );
  }

  // Calculate status info
  let statusInfo = {
    text: submission.status,
    color: "default" as any,
    icon: <AlertCircle className="w-4 h-4" />
  };

  switch (submission.status) {
    case 'pending':
      statusInfo = {
        text: "pending review",
        color: "secondary",
        icon: <Clock className="w-4 h-4" />
      };
      break;
    case 'approved':
      // Check if expired
      if (submission.submittedAt && submission.form) {
        const submissionDate = new Date(submission.submittedAt);
        const validityPeriod = submission.form.validityPeriod || 30;
        const expiryDate = new Date(submissionDate);
        expiryDate.setDate(expiryDate.getDate() + validityPeriod);
        
        const today = new Date();
        if (today > expiryDate) {
          statusInfo = {
            text: "expired",
            color: "destructive",
            icon: <XCircle className="w-4 h-4" />
          };
        } else {
          statusInfo = {
            text: "approved",
            color: "default",
            icon: <CheckCircle className="w-4 h-4" />
          };
        }
      } else {
        statusInfo = {
          text: "approved",
          color: "default",
          icon: <CheckCircle className="w-4 h-4" />
        };
      }
      break;
    case 'rejected':
      statusInfo = {
        text: "rejected",
        color: "destructive",
        icon: <XCircle className="w-4 h-4" />
      };
      break;
    default:
      statusInfo = {
        text: submission.status || "unknown",
        color: "secondary",
        icon: <AlertCircle className="w-4 h-4" />
      };
  }

  // Helper function to render form data
  const renderFormData = (data: any, fields: any[]) => {
    if (!data || !fields) return null;

    return fields.map((field: any) => {
      let value = data[field.name];
      
      // Handle different field types
      switch (field.type) {
        case 'checkbox':
          if (Array.isArray(value) && value.length > 0) {
            // Map the selected values back to their labels
            const selectedOptions = field.options?.filter((option: any) => 
              value.includes(option.value)
            ).map((option: any) => option.label) || [];
            
            value = selectedOptions.length > 0 ? (
              <ul className="list-disc ml-4">
                {selectedOptions.map((label: string, index: number) => (
                  <li key={index} className="text-sm">{label}</li>
                ))}
              </ul>
            ) : 'No options selected';
          } else {
            value = 'No options selected';
          }
          break;
        case 'date':
          value = value ? formatDate(value) : 'Not specified';
          break;
        case 'signature':
          value = value ? (
            <div className="flex justify-start">
              <img 
                src={value} 
                alt="Signature" 
                className="max-w-xs border border-gray-300 rounded shadow-sm max-h-32"
              />
            </div>
          ) : 'No signature uploaded';
          break;
        case 'file':
          value = value ? (
            <a href={value} target="_blank" rel="noopener noreferrer" className="text-blue-600 underline">
              View File
            </a>
          ) : 'No file uploaded';
          break;
        case 'embedded_excel': {
          const excelSrc = field.excelFileDataUrl || field.excelFileUrl;
          value = excelSrc ? (
            <div className="max-h-[min(70vh,520px)] overflow-auto border rounded-md bg-background">
              <EmbeddedExcelWorkbook
                excelSource={excelSrc}
                value={data[field.name]}
                onChange={() => {}}
                readOnly
              />
            </div>
          ) : (
            <span className="text-muted-foreground">No Excel template configured</span>
          );
          break;
        }
        default:
          value = value || 'Not specified';
      }

      return (
        <div key={field.name} className="space-y-1">
          <label className="text-sm font-medium text-gray-700">{field.label}</label>
          <div className="p-2 text-sm text-gray-900 border rounded bg-gray-50">
            {value}
          </div>
        </div>
      );
    });
  };

  // Helper function to render table data with pre-filled cells
  const renderTableData = (tableData: any[], tableFields: any[], preFilledData: any[] = []) => {
    if (!tableData || !tableFields || tableData.length === 0) {
      return <p className="text-sm ">No table data available</p>;
    }

    // Helper function to check if a cell is pre-filled
    const getPreFilledValue = (rowIndex: number, columnName: string) => {
      return preFilledData.find(
        (cell) => cell.rowIndex === rowIndex && cell.columnName === columnName
      );
    };

    return (
      <div className="space-y-2">
        {/* Legend */}
        {preFilledData && preFilledData.length > 0 && (
          <div className="flex items-center gap-4 p-2 text-xs border rounded bg-blue-50">
            <div className="flex items-center gap-2">
              <div className="w-3 h-3 rounded bg-blue-100 border-blue-300"></div>
              <span>Admin Pre-filled (Read-only)</span>
            </div>
            <div className="flex items-center gap-2">
              <div className="w-3 h-3 bg-white border border-gray-300 rounded"></div>
              <span>User Filled</span>
            </div>
          </div>
        )}
        
        <Table>
          <TableHeader>
            <TableRow>
              {tableFields.map((field: any) => (
                <TableHead key={field.name}>{field.label}</TableHead>
              ))}
            </TableRow>
          </TableHeader>
          <TableBody>
            {tableData.map((row: any, index: number) => (
              <TableRow key={index}>
                {tableFields.map((field: any) => {
                  const preFilledCell = getPreFilledValue(index, field.name);
                  const isPreFilled = !!preFilledCell;
                  
                  // Use pre-filled value if available, otherwise use submitted value
                  let cellValue = isPreFilled ? preFilledCell.value : (row[field.name] || 'N/A');
                  
                  // Handle signature type in table cells
                  if (field.type === 'signature' && cellValue && cellValue !== 'N/A') {
                    cellValue = (
                      <img 
                        src={cellValue} 
                        alt="Signature" 
                        className="max-w-xs border border-gray-300 rounded max-h-20"
                      />
                    );
                  }
                  
                  return (
                    <TableCell 
                      key={field.name}
                      className={isPreFilled ? "bg-blue-50 border-blue-200" : ""}
                    >
                      <div className="flex items-center gap-2">
                        {isPreFilled && (
                          <div className="w-2 h-2 rounded-full bg-blue-400 flex-shrink-0" title="Pre-filled by admin" />
                        )}
                        <div className="flex-1">
                          {cellValue}
                        </div>
                      </div>
                    </TableCell>
                  );
                })}
              </TableRow>
            ))}
          </TableBody>
        </Table>
      </div>
    );
  };

  const exportSubmissionPdf = async () => {
    try {
      const sanitize = (raw: string, fallback: string) => {
        const s = (raw || fallback)
          .replace(/[/\\?%*:|"<>]/g, "-")
          .replace(/\s+/g, "-")
          .replace(/-+/g, "-")
          .replace(/^-|-$/g, "");
        return s || fallback;
      };
      const formPart = sanitize(String(submission.form?.title || ""), "form");
      const shipPart = sanitize(
        submission.ship?.name ? String(submission.ship.name) : "",
        "No-ship"
      );
      const filename = `${formPart}-${shipPart}.pdf`;
      await downloadPdfDocument(
        <SubmissionPdfDocument submission={submission} />,
        filename
      );
      toast({ title: "PDF downloaded", variant: "success" });
    } catch (e: unknown) {
      const msg = e instanceof Error ? e.message : "Could not generate PDF";
      toast({ title: "PDF export failed", description: msg, variant: "destructive" });
    }
  };

  return (
    <Dialog open={isOpen} onOpenChange={onClose}>
      <DialogContent className="max-w-6xl max-h-[90vh]">
        <DialogHeader className="flex flex-col gap-3 sm:flex-row sm:items-start sm:justify-start sm:space-y-0 sm:gap-4">
          <div className="space-y-1.5">
            <DialogTitle className="flex items-center gap-2">
              <FileText className="w-5 h-5" />
              Submission Details
            </DialogTitle>
            <DialogDescription>
              Viewing submission for "{submission.form?.title}"
            </DialogDescription>
          </div>
          <Button
            type="button"
            variant="outline"
            size="sm"
            className="shrink-0"
            onClick={exportSubmissionPdf}
          >
            <FileDown className="w-4 h-4 mr-2" />
            Export PDF
          </Button>
        </DialogHeader>

        <ScrollArea className="max-h-[calc(90vh-120px)]">
          <div className="p-1 space-y-6">
            {/* Header Info */}
            <div className="grid grid-cols-1 gap-4 md:grid-cols-2 lg:grid-cols-4">
              <Card>
                <CardHeader className="pb-2">
                  <CardTitle className="flex items-center gap-2 text-sm">
                    <User className="w-4 h-4" />
                    Submitted By
                  </CardTitle>
                </CardHeader>
                <CardContent>
                  <p className="font-medium">{submission.user?.name}</p>
                  <p className="text-sm ">{submission.user?.email}</p>
                </CardContent>
              </Card>

              <Card>
                <CardHeader className="pb-2">
                  <CardTitle className="flex items-center gap-2 text-sm">
                    <Ship className="w-4 h-4" />
                    Ship
                  </CardTitle>
                </CardHeader>
                <CardContent>
                  <p className="font-medium">{submission.ship?.name || 'N/A'}</p>
                </CardContent>
              </Card>

              <Card>
                <CardHeader className="pb-2">
                  <CardTitle className="flex items-center gap-2 text-sm">
                    <Calendar className="w-4 h-4" />
                    Submitted At
                  </CardTitle>
                </CardHeader>
                <CardContent>
                  <p className="font-medium">
                    {submission.submittedAt ? formatDate(submission.submittedAt) : 'N/A'}
                  </p>
                </CardContent>
              </Card>

              <Card>
                <CardHeader className="pb-2">
                  <CardTitle className="flex items-center gap-2 text-sm">
                    <Clock className="w-4 h-4" />
                    Status
                  </CardTitle>
                </CardHeader>
                <CardContent>
                  <Badge variant={statusInfo.color} className="flex items-center gap-1 w-fit">
                    {statusInfo.icon}
                    {statusInfo.text}
                  </Badge>
                </CardContent>
              </Card>
            </div>

            {/* Review Information - only show if reviewed */}
            {(submission.reviewedAt || submission.reviewedBy || submission.reviewComments) && (
              <>
                <hr className="my-4" />
                <Card>
                  <CardHeader>
                    <CardTitle className="flex items-center gap-2">
                      <CheckCircle className="w-4 h-4" />
                      Review Information
                    </CardTitle>
                  </CardHeader>
                  <CardContent>
                    <div className="grid grid-cols-1 gap-4 md:grid-cols-2">
                      {submission.reviewedBy && (
                        <div>
                          <label className="text-sm font-medium text-gray-700">Reviewed By</label>
                          <p className="text-sm text-gray-900">{submission.reviewedBy.name || submission.reviewedBy.email}</p>
                        </div>
                      )}
                      {submission.reviewedAt && (
                        <div>
                          <label className="text-sm font-medium text-gray-700">Reviewed At</label>
                          <p className="text-sm text-gray-900">{formatDate(submission.reviewedAt)}</p>
                        </div>
                      )}
                    </div>
                    {submission.reviewComments && (
                      <div className="mt-4">
                        <label className="text-sm font-medium text-gray-700">Review Comments</label>
                        <div className="p-3 mt-1 text-sm text-gray-900 border rounded bg-gray-50">
                          {submission.reviewComments}
                        </div>
                      </div>
                    )}
                  </CardContent>
                </Card>
              </>
            )}

            <hr className="my-4" />

            {/* Form Information */}
            <Card>
              <CardHeader>
                <CardTitle>Form Information</CardTitle>
              </CardHeader>
              <CardContent className="space-y-4">
                <div className="grid grid-cols-1 gap-4 md:grid-cols-3">
                  <div>
                    <label className="text-sm font-medium text-gray-700">Form Title</label>
                    <p className="text-sm text-gray-900">{submission.form?.title}</p>
                  </div>
                  <div>
                    <label className="text-sm font-medium text-gray-700">Category</label>
                    <p className="text-sm text-gray-900">{submission.form?.category?.name || submission.form?.category}</p>
                  </div>
                  <div>
                    <label className="text-sm font-medium text-gray-700">Form Type</label>
                    <p className="text-sm text-gray-900">{submission.form?.formType}</p>
                  </div>
                </div>
                {submission.form?.description && (
                  <div>
                    <label className="text-sm font-medium text-gray-700">Description</label>
                    <p className="p-2 text-sm text-gray-900 border rounded bg-gray-50">
                      {submission.form.description}
                    </p>
                  </div>
                )}
              </CardContent>
            </Card>

            {/* Submission Data */}
            {(submission.form?.formType === 'regular' || !submission.form?.formType) && submission.data && (
              <Card>
                <CardHeader>
                  <CardTitle>Form Responses</CardTitle>
                  <CardDescription>Regular form field responses</CardDescription>
                </CardHeader>
                <CardContent className="space-y-4">
                  {renderFormData(submission.data, submission.form?.fields || [])}
                </CardContent>
              </Card>
            )}

            {submission.form?.formType === 'table' && submission.data && (
              <Card>
                <CardHeader>
                  <CardTitle>Table Data</CardTitle>
                  <CardDescription>Submitted table entries</CardDescription>
                </CardHeader>
                <CardContent>
                  {renderTableData(
                    submission.data.tableData || [], 
                    submission.form?.tableConfig?.columns || [],
                    submission.form?.tableConfig?.preFilledData || []
                  )}
                </CardContent>
              </Card>
            )}

            {submission.form?.formType === 'mixed' && submission.data && submission.form?.sections && (
              <>
                {submission.form.sections.map((section: any) => (
                  <Card key={section.id}>
                    <CardHeader>
                      <CardTitle>{section.title}</CardTitle>
                      {section.description && (
                        <CardDescription>{section.description}</CardDescription>
                      )}
                    </CardHeader>
                    <CardContent className="space-y-4">
                      {section.type === 'fields' && section.fields && (
                        <div className="space-y-4">
                          {renderFormData(submission.data, section.fields)}
                        </div>
                      )}
                      {section.type === 'table' && section.tableConfig && (
                        <div>
                          {renderTableData(
                            submission.data[`table_${section.id}`] || [],
                            section.tableConfig.columns || [],
                            section.tableConfig.preFilledData || []
                          )}
                        </div>
                      )}
                    </CardContent>
                  </Card>
                ))}
              </>
            )}

            {/* Fallback: Raw Data Display */}
            {submission.data && (!submission.form?.fields && !submission.form?.tableConfig && !submission.form?.sections) && (
              <Card>
                <CardHeader>
                  <CardTitle>Submission Data</CardTitle>
                  <CardDescription>Raw submission data</CardDescription>
                </CardHeader>
                <CardContent>
                  <div className="space-y-4">
                    {Object.entries(submission.data).map(([key, value]) => {
                      // Handle table data specially
                      if (key.includes('table_') && Array.isArray(value)) {
                        return (
                          <div key={key} className="space-y-2">
                            <label className="text-sm font-medium text-gray-700">Table: {key}</label>
                            <div className="p-2 border rounded bg-gray-50">
                              <Table>
                                <TableHeader>
                                  <TableRow>
                                    {value.length > 0 && Object.keys(value[0]).map((col) => (
                                      <TableHead key={col}>{col}</TableHead>
                                    ))}
                                  </TableRow>
                                </TableHeader>
                                <TableBody>
                                  {value.map((row: any, index: number) => (
                                    <TableRow key={index}>
                                      {Object.values(row).map((cellValue: any, cellIndex: number) => (
                                        <TableCell key={cellIndex}>{String(cellValue)}</TableCell>
                                      ))}
                                    </TableRow>
                                  ))}
                                </TableBody>
                              </Table>
                            </div>
                          </div>
                        );
                      }
                      
                      // Handle regular fields
                      return (
                        <div key={key} className="space-y-1">
                          <label className="text-sm font-medium text-gray-700">{key}</label>
                          <div className="p-2 text-sm text-gray-900 border rounded bg-gray-50">
                            {typeof value === 'object' ? JSON.stringify(value, null, 2) : String(value)}
                          </div>
                        </div>
                      );
                    })}
                  </div>
                </CardContent>
              </Card>
            )}

            {/* Expiry Information */}
            {submission.submittedAt && submission.form?.validityPeriod && (
              <Card>
                <CardHeader>
                  <CardTitle>Validity Information</CardTitle>
                </CardHeader>
                <CardContent>
                  <div className="grid grid-cols-1 gap-4 md:grid-cols-2">
                    <div>
                      <label className="text-sm font-medium text-gray-700">Validity Period</label>
                      <p className="text-sm text-gray-900">{submission.form.validityPeriod} days</p>
                    </div>
                    <div>
                      <label className="text-sm font-medium text-gray-700">Expires On</label>
                      <p className="text-sm text-gray-900">
                        {(() => {
                          const submissionDate = new Date(submission.submittedAt);
                          const expiryDate = new Date(submissionDate);
                          expiryDate.setDate(expiryDate.getDate() + submission.form.validityPeriod);
                          return formatDate(expiryDate);
                        })()}
                      </p>
                    </div>
                  </div>
                </CardContent>
              </Card>
            )}
          </div>
        </ScrollArea>
      </DialogContent>
    </Dialog>
  );
};

export default SubmissionViewModal;