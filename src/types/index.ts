export interface User {
  _id: string;
  email: string;
  name: string;
  role: "user" | "admin" | "super_admin";
  userType?: "deck" | "engine";
  ship?: Ship;
  isActive: boolean;
  lastLogin?: string;
  createdAt: string;
  updatedAt: string;
}

export interface Ship {
  _id: string;
  name: string;
  imoNumber: string;
  vesselType?: string;
  flag?: string;
  grossTonnage?: number;
  buildYear?: number;
  owner?: string;
  operator?: string;
  user?: string;
  isActive: boolean;
  documents?: Document[];
  createdAt: string;
  updatedAt: string;
}

export interface Document {
  name: string;
  type: string;
  expiryDate: string;
  status: "valid" | "expiring" | "expired";
}

export interface Category {
  _id: string;
  name: "eng" | "deck" | "mlc" | "isps" | "drill" | "deck_engine";
  displayName: string;
  description?: string;
  isActive: boolean;
  color: string;
  icon: string;
  createdAt: string;
  updatedAt: string;
}

export interface FormField {
  id: string;
  name: string;
  label: string;
  type: "text" | "email" | "number" | "date" | "datetime-local" | "time" | "textarea" | "select" | "checkbox" | "radio" | "file" | "phone" | "url" | "signature" | "embedded_excel";
  required: boolean;
  /** Legacy: public path to `.xlsx` (older forms). New forms use `excelFileDataUrl` only. */
  excelFileUrl?: string;
  /** Uploaded `.xlsx` as a data URL (takes precedence over `excelFileUrl` when set). */
  excelFileDataUrl?: string;
  excelDisplayName?: string;
  /** In-app created spreadsheet template (preferred for new forms). */
  excelTemplate?: {
    sheets: Array<{
      name: string;
      grid: string[][];
      mergeCells?: Array<{ row: number; col: number; rowspan: number; colspan: number }>;
      cellMeta?: Array<{
        row: number;
        col: number;
        className?: string;
        readOnly?: boolean;
        formula?: string;
        formulaCachedValue?: string;
        formulaWarning?: boolean;
      }>;
    }>;
  };
  placeholder?: string;
  description?: string;
  options?: { label: string; value: string }[];
  validation?: {
    min?: number;
    max?: number;
    minLength?: number;
    maxLength?: number;
    pattern?: string;
    message?: string;
  };
  layout?: {
    width: "full" | "half" | "third" | "quarter" | "auto";
    order: number;
    row: number;
    column: number;
  };
  style?: {
    className?: string;
    customCSS?: string;
  };
}

export interface TableColumn {
  id: string;
  name: string;
  label: string;
  type: "text" | "number" | "date" | "select" | "checkbox" | "email" | "phone" | "signature";
  width?: number;
  required: boolean;
  options?: string[];
  validation?: {
    min?: number;
    max?: number;
    minLength?: number;
    maxLength?: number;
    pattern?: string;
    message?: string;
  };
  order: number;
}

export interface PreFilledCell {
  rowIndex: number;
  columnName: string;
  value: any;
  isReadOnly: boolean;
}

export interface TableConfig {
  id: string;
  name: string;
  title: string;
  description?: string;
  columns: TableColumn[];
  preFilledData?: PreFilledCell[];
  minRows: number;
  maxRows: number;
  allowAddRows: boolean;
  allowDeleteRows: boolean;
  defaultRows: number;
  layout?: {
    order: number;
    width: "full" | "half" | "third" | "quarter";
  };
}

export interface FormSection {
  id: string;
  name: string;
  title: string;
  description?: string;
  type: "fields" | "table";
  fields?: FormField[];
  tableConfig?: TableConfig;
  layout?: {
    order: number;
    columnsPerRow: number;
  };
  style?: {
    backgroundColor?: string;
    borderColor?: string;
    className?: string;
  };
}

export interface FormLayout {
  columnsPerRow: number;
  spacing: "compact" | "normal" | "relaxed";
  theme: "default" | "modern" | "minimal" | "professional";
}

export interface Form {
  _id: string;
  title: string;
  description?: string;
  category: Category;
  formType: "regular" | "table" | "mixed";
  fields?: FormField[];
  tableConfig?: TableConfig;
  sections?: FormSection[];
  layout?: FormLayout;
  validityPeriod: number;
  status: "not-submitted" | "active" | "expiring-soon" | "expired";
  isActive: boolean;
  createdBy: User;
  updatedBy?: User;
  createdAt: string;
  updatedAt: string;
}

export interface FormSubmission {
  _id: string;
  form: Form;
  user: User;
  ship?: Ship;
  data: Record<string, any>;
  status: "draft" | "submitted" | "approved" | "rejected";
  submittedAt?: string;
  approvedAt?: string;
  reviewedBy?: User;
  feedback?: string;
  createdAt: string;
  updatedAt: string;
}

export interface Notification {
  _id: string;
  recipient: User;
  type: "unfilled_form" | "form_submitted" | "form_approved" | "form_rejected" | "form_expiring_2_days" | "form_expiring_today" | "form_expired" | "form_status_expired" | "form_status_expiring_soon" | "user_form_expiring_today" | "user_form_expired" | "system";
  title: string;
  message: string;
  relatedForm?: Form;
  relatedSubmission?: FormSubmission;
  isRead: boolean;
  priority: "low" | "medium" | "high" | "urgent";
  createdAt: string;
  updatedAt: string;
}

// Form Builder Types
export interface FieldTemplate {
  id: string;
  type: FormField["type"];
  label: string;
  icon: string;
  description: string;
  defaultProps: Partial<FormField>;
}

export interface DraggedItem {
  id: string;
  type: "field" | "table" | "section";
  data: FieldTemplate | TableConfig | FormSection;
}

export interface DropZoneProps {
  onDrop: (item: DraggedItem) => void;
  accepts: Array<"field" | "table" | "section">;
  isOver?: boolean;
}

export interface FormBuilderState {
  form: Partial<Form>;
  selectedItem: string | null;
  draggedItem: DraggedItem | null;
  history: Partial<Form>[];
  historyIndex: number;
}

export interface ApiResponse<T> {
  success: boolean;
  message: string;
  data: T;
  pagination?: {
    currentPage: number;
    totalPages: number;
    totalItems: number;
    itemsPerPage: number;
    hasNextPage: boolean;
    hasPrevPage: boolean;
    nextPage: number | null;
    prevPage: number | null;
  };
}

export interface AuthState {
  user: User | null;
  token: string | null;
  refreshToken: string | null;
  isAuthenticated: boolean;
  loading: boolean;
}