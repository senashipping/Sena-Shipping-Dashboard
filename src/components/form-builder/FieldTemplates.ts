import { FieldTemplate } from "../../types";
import { 
  Type, 
  Mail, 
  Hash, 
  Calendar, 
  Clock,
  AlignLeft, 
  ChevronDown, 
  CheckSquare, 
  Circle, 
  FileUp, 
  Phone, 
  Globe,
  Table,
  PenTool
} from "lucide-react";

export const FIELD_TEMPLATES: FieldTemplate[] = [
  {
    id: "text",
    type: "text",
    label: "Text Input",
    icon: "Type",
    description: "Single line text input",
    defaultProps: {
      type: "text",
      required: false,
      placeholder: "Enter text...",
      layout: { width: "full", order: 0, row: 0, column: 0 }
    }
  },
  {
    id: "email",
    type: "email",
    label: "Email",
    icon: "Mail",
    description: "Email address input",
    defaultProps: {
      type: "email",
      required: false,
      placeholder: "Enter email address...",
      validation: { pattern: "^[^\\s@]+@[^\\s@]+\\.[^\\s@]+$", message: "Please enter a valid email address" },
      layout: { width: "full", order: 0, row: 0, column: 0 }
    }
  },
  {
    id: "number",
    type: "number",
    label: "Number",
    icon: "Hash",
    description: "Numeric input",
    defaultProps: {
      type: "number",
      required: false,
      placeholder: "Enter number...",
      layout: { width: "half", order: 0, row: 0, column: 0 }
    }
  },
  {
    id: "date",
    type: "date",
    label: "Date",
    icon: "Calendar",
    description: "Date picker",
    defaultProps: {
      type: "date",
      required: false,
      layout: { width: "half", order: 0, row: 0, column: 0 }
    }
  },
  {
    id: "datetime-local",
    type: "datetime-local",
    label: "Date & Time",
    icon: "Clock",
    description: "Date and time picker",
    defaultProps: {
      type: "datetime-local",
      required: false,
      layout: { width: "half", order: 0, row: 0, column: 0 }
    }
  },
  {
    id: "time",
    type: "time",
    label: "Time",
    icon: "Clock",
    description: "Time picker",
    defaultProps: {
      type: "time",
      required: false,
      layout: { width: "half", order: 0, row: 0, column: 0 }
    }
  },
  {
    id: "textarea",
    type: "textarea",
    label: "Text Area",
    icon: "AlignLeft",
    description: "Multi-line text input",
    defaultProps: {
      type: "textarea",
      required: false,
      placeholder: "Enter detailed text...",
      layout: { width: "full", order: 0, row: 0, column: 0 }
    }
  },
  {
    id: "select",
    type: "select",
    label: "Dropdown",
    icon: "ChevronDown",
    description: "Select from dropdown options",
    defaultProps: {
      type: "select",
      required: false,
      options: [
        { label: "Option 1", value: "option1" },
        { label: "Option 2", value: "option2" },
        { label: "Option 3", value: "option3" }
      ],
      layout: { width: "half", order: 0, row: 0, column: 0 }
    }
  },
  {
    id: "checkbox",
    type: "checkbox",
    label: "Checkbox",
    icon: "CheckSquare",
    description: "Single or multiple checkboxes",
    defaultProps: {
      type: "checkbox",
      required: false,
      layout: { width: "full", order: 0, row: 0, column: 0 }
    }
  },
  {
    id: "radio",
    type: "radio",
    label: "Radio Button",
    icon: "Circle",
    description: "Choose one from multiple options",
    defaultProps: {
      type: "radio",
      required: false,
      options: [
        { label: "Option 1", value: "option1" },
        { label: "Option 2", value: "option2" },
        { label: "Option 3", value: "option3" }
      ],
      layout: { width: "full", order: 0, row: 0, column: 0 }
    }
  },
  {
    id: "signature",
    type: "signature",
    label: "Signature Upload",
    icon: "PenTool",
    description: "Upload signature image",
    defaultProps: {
      type: "signature",
      required: false,
      placeholder: "Upload your signature",
      layout: { width: "full", order: 0, row: 0, column: 0 }
    }
  },
  {
    id: "phone",
    type: "phone",
    label: "Phone Number",
    icon: "Phone",
    description: "Phone number input",
    defaultProps: {
      type: "phone",
      required: false,
      placeholder: "Enter phone number...",
      validation: { pattern: "^[+]?[\\d\\s\\-\\(\\)]+$", message: "Please enter a valid phone number" },
      layout: { width: "half", order: 0, row: 0, column: 0 }
    }
  },
  {
    id: "url",
    type: "url",
    label: "URL",
    icon: "Globe",
    description: "Website URL input",
    defaultProps: {
      type: "url",
      required: false,
      placeholder: "https://example.com",
      validation: { pattern: "^https?:\\/\\/.+", message: "Please enter a valid URL starting with http:// or https://" },
      layout: { width: "full", order: 0, row: 0, column: 0 }
    }
  }
];

export const ICON_MAP = {
  Type,
  Mail,
  Hash,
  Calendar,
  Clock,
  AlignLeft,
  ChevronDown,
  CheckSquare,
  Circle,
  FileUp,
  Phone,
  Globe,
  Table,
  PenTool
};