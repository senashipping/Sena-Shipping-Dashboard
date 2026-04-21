import React, { useState } from "react";
import { useParams, useNavigate, useSearchParams } from "react-router-dom";
import { useQuery, useMutation, useQueryClient } from "@tanstack/react-query";
import api from "../../api";
import { Button } from "../../components/ui/button";
import { Alert, AlertDescription } from "../../components/ui/alert";
import { ArrowLeft, FileText } from "lucide-react";
import SharedFormRenderer, {
  SharedFormRendererRef,
} from "../../components/form-builder/SharedFormRenderer";

const FormSubmission: React.FC = () => {
  const { id: formId } = useParams<{ id: string }>();
  const navigate = useNavigate();
  const [searchParams] = useSearchParams();
  const isAdminView = searchParams.get("admin") === "true";
  const submissionId = searchParams.get("submissionId"); // Get from query params if editing existing
  const queryClient = useQueryClient();

  const [formData, setFormData] = useState<Record<string, any>>({});
  const [tableData, setTableData] = useState<any[]>([]);
  const latestResolvedFormDataRef = React.useRef<Record<string, any>>({});
  const formRendererRef = React.useRef<SharedFormRendererRef | null>(null);

  const { data: form, isLoading: formLoading } = useQuery({
    queryKey: ["form", formId],
    queryFn: () => api.getFormById(formId!),
    enabled: !!formId,
  });

  const { data: submission, isLoading: submissionLoading } = useQuery({
    queryKey: ["submission", submissionId],
    queryFn: () => api.getSubmissionById(submissionId as string),
    enabled: !!submissionId,
  });

  const createMutation = useMutation({
    mutationFn: (data: any) => api.createSubmission(data),
    onSuccess: () => {
      // Invalidate relevant queries to refresh dashboard data
      queryClient.invalidateQueries({ queryKey: ["userDashboard"] });
      queryClient.invalidateQueries({ queryKey: ["activeForms"] });
      queryClient.invalidateQueries({ queryKey: ["submissions"] });
      navigate(isAdminView ? "/admin/submissions" : "/dashboard");
    },
  });



  React.useEffect(() => {
    if (submission?.data) {
      if (form?.data?.data?.formType === "table") {
        setTableData(submission.data.data?.tableData || []);
      } else {
        setFormData(submission.data.data || {});
      }
    }
  }, [submission, form]);

  React.useEffect(() => {
    latestResolvedFormDataRef.current = formData;
  }, [formData]);

  // Initialize mixed form table data
  React.useEffect(() => {
    if (form?.data?.data?.formType === "mixed" && form?.data?.data?.sections) {
      const initialData = { ...formData };
      
      form.data.data.sections.forEach((section: any) => {
        if (section.type === "table") {
          const tableKey = `table_${section.id}`;
          if (!initialData[tableKey]) {
            initialData[tableKey] = [{}]; // Initialize with one empty row
          }
        }
      });
      
      setFormData(initialData);
    }
  }, [form]);

  const handleInputChange = (name: string, value: any) => {
    setFormData(prev => ({
      ...prev,
      [name]: value
    }));
  };

  const handleTableChange = (rowIndex: number, columnName: string, value: any) => {
    setTableData(prev => {
      const newData = [...prev];
      if (!newData[rowIndex]) newData[rowIndex] = {};
      newData[rowIndex][columnName] = value;
      return newData;
    });
  };



  // Mixed form table functions
  const handleMixedTableChange = (sectionId: string, rowIndex: number, columnName: string, value: any) => {
    const tableKey = `table_${sectionId}`;
    setFormData(prev => {
      const tableData = prev[tableKey] || [];
      const newTableData = [...tableData];
      if (!newTableData[rowIndex]) {
        newTableData[rowIndex] = {};
      }
      newTableData[rowIndex][columnName] = value;
      return {
        ...prev,
        [tableKey]: newTableData
      };
    });
  };

  const handleAddTableRow = (tableId?: string) => {
    if (tableId) {
      // Mixed form table
      const tableKey = `table_${tableId}`;
      setFormData(prev => ({
        ...prev,
        [tableKey]: [...(prev[tableKey] || [{}]), {}]
      }));
    } else {
      // Regular table form
      setTableData(prev => [...prev, {}]);
    }
  };

  const handleRemoveTableRow = (rowIndex: number, tableId?: string) => {
    if (tableId) {
      // Mixed form table
      const tableKey = `table_${tableId}`;
      setFormData(prev => ({
        ...prev,
        [tableKey]: (prev[tableKey] || []).filter((_: any, i: number) => i !== rowIndex)
      }));
    } else {
      // Regular table form
      setTableData(prev => prev.filter((_, i) => i !== rowIndex));
    }
  };

  const handleSubmit = (e: React.FormEvent) => {
    e.preventDefault();
    
    let submissionFormData;
    if (form?.data?.data?.formType === "table") {
      submissionFormData = { tableData };
    } else if (form?.data?.data?.formType === "mixed") {
      submissionFormData =
        formRendererRef.current?.getResolvedFormData() ??
        latestResolvedFormDataRef.current;
    } else {
      submissionFormData =
        formRendererRef.current?.getResolvedFormData() ??
        latestResolvedFormDataRef.current;
    }

    const submissionData = {
      formId: formId,
      data: submissionFormData
    };

    // Only create new submissions (no updates since forms can't be modified)
    createMutation.mutate(submissionData);
  };



  // Show loading only when form is actually loading or when submission is loading (if submissionId exists)
  if (formLoading || (submissionId && submissionLoading)) {
    return (
      <div className="flex items-center justify-center h-64">
        <div className="w-12 h-12 border-b-2 rounded-full animate-spin border-primary"></div>
      </div>
    );
  }

  const getErrorMessage = (err: unknown): string => {
    if (typeof err === 'object' && err !== null && 'response' in err) {
      const response = (err as any).response;
      if (typeof response === 'object' && response !== null && 'data' in response) {
        const data = response.data;
        if (typeof data === 'object' && data !== null && 'message' in data) {
          return data.message;
        }
      }
    }
    return "An error occurred";
  };

  const error = createMutation.error;
  const errorMessage = error ? getErrorMessage(error) : null;

  return (
    <div className="space-y-6">
      <div className="flex flex-col gap-4 lg:flex-row lg:items-center lg:justify-between">
        <div className="flex items-center space-x-4">
          <Button 
            variant="outline" 
            size="icon" 
            onClick={() => navigate(isAdminView ? "/admin/submissions" : "/dashboard")}
          >
            <ArrowLeft className="w-4 h-4" />
          </Button>
          <h1 className="text-3xl font-bold">
            {form?.data?.data?.title || "Form Submission"}
          </h1>
        </div>
      </div>

      {errorMessage && (
        <Alert variant="destructive">
          <AlertDescription>
            {errorMessage}
          </AlertDescription>
        </Alert>
      )}

      {form?.data?.data && (
        <form onSubmit={handleSubmit}>
          <SharedFormRenderer
            ref={formRendererRef}
            formState={form.data.data}
            formData={formData}
            tableData={tableData}
            onFieldChange={handleInputChange}
            onTableChange={handleTableChange}
            onMixedTableChange={handleMixedTableChange}
            onAddTableRow={handleAddTableRow}
            onRemoveTableRow={handleRemoveTableRow}
            onResolvedFormDataChange={(data) => {
              latestResolvedFormDataRef.current = data;
            }}
            useLocalExcelState
            excelReadOnly
            allowReadOnlyWorkbookActions
            submitButton={
              <Button
                type="submit"
                disabled={createMutation.isPending}
                className="w-full"
              >
                <FileText className="w-4 h-4 mr-2" />
                {createMutation.isPending ? "Submitting..." : "Submit Form"}
              </Button>
            }
          />
        </form>
      )}
    </div>
  );
};

export default FormSubmission;