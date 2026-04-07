import React, { useState, useCallback, useEffect } from "react";
import { useParams, useNavigate } from "react-router-dom";
import { useQuery, useMutation } from "@tanstack/react-query";
import {
  DndContext,
  DragEndEvent,
  DragStartEvent,
  closestCenter,
  PointerSensor,
  useSensor,
  useSensors,
} from "@dnd-kit/core";
import { arrayMove } from "@dnd-kit/sortable";
import { v4 as uuidv4 } from "uuid";
import api from "../../api";
import { Button } from "../../components/ui/button";
import { Card, CardContent, CardDescription, CardHeader, CardTitle } from "../../components/ui/card";
import { Input } from "../../components/ui/input";
import { Textarea } from "../../components/ui/textarea";
import { Label } from "../../components/ui/label";
import { Select, SelectContent, SelectItem, SelectTrigger, SelectValue } from "../../components/ui/select";
import { Dialog, DialogContent, DialogHeader, DialogTitle, DialogTrigger } from "../../components/ui/dialog";

import { ArrowLeft, Eye, Save, Undo, Redo, Plus } from "lucide-react";
import { FormField, FormSection, TableConfig, Category, FieldTemplate, FormLayout } from "../../types";
import FieldPalette from "../../components/form-builder/FieldPalette";
import FormCanvas from "../../components/form-builder/FormCanvas";
import PropertiesPanel from "../../components/form-builder/PropertiesPanel";
import PreviewForm from "../../components/form-builder/PreviewForm";

interface FormBuilderState {
  title: string;
  description: string;
  category: string;
  formType: "regular" | "table" | "mixed";
  validityPeriod: number;
  layout: FormLayout;
  fields: FormField[];
  sections: FormSection[];
  tableConfig?: TableConfig;
}

const ALLOWED_FIELD_TYPES = new Set([
  "text",
  "email",
  "number",
  "date",
  "datetime-local",
  "time",
  "textarea",
  "select",
  "checkbox",
  "radio",
  "file",
  "phone",
  "url",
  "signature",
]);

const sanitizeFieldForSubmit = (field: FormField): FormField => {
  const normalizedType = field.type === "embedded_excel" ? "file" : field.type;
  const safeType = ALLOWED_FIELD_TYPES.has(normalizedType) ? normalizedType : "text";
  return { ...field, type: safeType as FormField["type"] };
};

const FormBuilder: React.FC = () => {
  const { id } = useParams();
  const navigate = useNavigate();
  const isEditing = !!id;

  const [formState, setFormState] = useState<FormBuilderState>({
    title: "",
    description: "",
    category: "",
    formType: "regular",
    validityPeriod: 30,
    layout: {
      columnsPerRow: 1,
      spacing: "normal",
      theme: "default",
    },
    fields: [],
    sections: [],
    tableConfig: undefined,
  });

  const [selectedItem, setSelectedItem] = useState<{
    id: string;
    type: "field" | "section" | "table";
    data: FormField | FormSection | TableConfig;
  } | null>(null);

  const [history, setHistory] = useState<FormBuilderState[]>([formState]);
  const [historyIndex, setHistoryIndex] = useState(0);

  const sensors = useSensors(
    useSensor(PointerSensor, {
      activationConstraint: {
        distance: 8,
      },
    })
  );

  const { data: categories } = useQuery({
    queryKey: ["categories"],
    queryFn: () => api.getCategories({ isActive: true }),
  });

  const { data: form } = useQuery({
    queryKey: ["form", id],
    queryFn: () => api.getFormById(id!),
    enabled: !!id,
  });

  // Load form data when editing
  useEffect(() => {
    if (form?.data?.data) {
      const formData = form.data.data;
      
      setFormState({
        title: formData.title || "",
        description: formData.description || "",
        category: formData.category._id || "",
        formType: formData.formType || "regular",
        validityPeriod: formData.validityPeriod || 30,
        layout: formData.layout || {
          columnsPerRow: 1,
          spacing: "normal",
          theme: "default",
        },
        fields: formData.fields || [],
        sections: formData.sections || [],
        tableConfig: formData.tableConfig,
      });
    }
  }, [form]);

  const createMutation = useMutation({
    mutationFn: (data: any) => api.createForm(data),
    onSuccess: () => {
      navigate("/admin/forms");
    },
  });

  const updateMutation = useMutation({
    mutationFn: (data: any) => api.updateForm(id!, data),
    onSuccess: () => {
      navigate("/admin/forms");
    },
  });

  const saveToHistory = useCallback((newState: FormBuilderState) => {
    setHistory(prev => {
      const newHistory = prev.slice(0, historyIndex + 1);
      newHistory.push(newState);
      return newHistory.slice(-50); // Keep only last 50 states
    });
    setHistoryIndex(prev => prev + 1);
  }, [historyIndex]);

  const updateFormState = useCallback((updates: Partial<FormBuilderState>) => {
    const newState = { ...formState, ...updates };
    setFormState(newState);
    saveToHistory(newState);
  }, [formState, saveToHistory]);

  const undo = () => {
    if (historyIndex > 0) {
      setHistoryIndex(prev => prev - 1);
      setFormState(history[historyIndex - 1]);
    }
  };

  const redo = () => {
    if (historyIndex < history.length - 1) {
      setHistoryIndex(prev => prev + 1);
      setFormState(history[historyIndex + 1]);
    }
  };

  const handleDragStart = (_event: DragStartEvent) => {
    // Store the dragged item data for potential use
  };

  const addFieldToMixedSection = (newField: FormField, preferredSectionId?: string) => {
    const preferredSectionIndex = preferredSectionId
      ? formState.sections.findIndex(s => s.id === preferredSectionId && s.type === "fields")
      : -1;

    if (preferredSectionIndex !== -1) {
      const sectionsCopy = [...formState.sections];
      const target = sectionsCopy[preferredSectionIndex];
      sectionsCopy[preferredSectionIndex] = {
        ...target,
        fields: [...(target.fields || []), newField],
      };
      updateFormState({ sections: sectionsCopy });
      return;
    }

    const selectedSectionIndex =
      selectedItem?.type === "section"
        ? formState.sections.findIndex(s => s.id === selectedItem.id && s.type === "fields")
        : -1;

    if (selectedSectionIndex !== -1) {
      const sectionsCopy = [...formState.sections];
      const target = sectionsCopy[selectedSectionIndex];
      sectionsCopy[selectedSectionIndex] = {
        ...target,
        fields: [...(target.fields || []), newField],
      };
      updateFormState({ sections: sectionsCopy });
      return;
    }

    // Fallback: append to the last fields section, or create one if none exists
    const lastFieldsSectionIndex = formState.sections.map((s, i) => ({ s, i })).filter(({ s }) => s.type === "fields").pop()?.i;
    if (lastFieldsSectionIndex === undefined) {
      const defaultSection: FormSection = {
        id: uuidv4(),
        name: `section_${Date.now()}`,
        title: `Section ${formState.sections.length + 1}`,
        type: "fields",
        fields: [newField],
        layout: { order: formState.sections.length, columnsPerRow: 1 },
      };
      updateFormState({
        sections: [...formState.sections, defaultSection],
      });
      return;
    }

    const sectionsCopy = [...formState.sections];
    const target = sectionsCopy[lastFieldsSectionIndex];
    sectionsCopy[lastFieldsSectionIndex] = {
      ...target,
      fields: [...(target.fields || []), newField],
    };
    updateFormState({ sections: sectionsCopy });
  };

  const handleDragEnd = (event: DragEndEvent) => {
    const { active, over } = event;
    
    if (!over) return;

    const isNewElement = active.data.current?.type === "field" || active.data.current?.type === "table";
    const isReorderingElement = (!isNewElement) && (
      formState.fields.some(f => f.id === active.id) ||
      formState.sections.some(s => s.id === active.id)
    );

    // Handle dropping new elements from palette
    if (isNewElement && (over.id === "form-canvas" || over.data.current?.accepts?.includes(active.data.current?.type))) {
      if (active.data.current?.type === "field") {
        const template = active.data.current.template as FieldTemplate;
        const fieldId = uuidv4();
        const newField: FormField = {
          id: fieldId,
          name: fieldId, // Use unique ID as name to prevent conflicts
          label: template.label,
          ...template.defaultProps,
        } as FormField;

        if (formState.formType === "regular") {
          // Insert at the appropriate position if dropping over an existing field
          let insertIndex = formState.fields.length;
          if (over.id !== "form-canvas") {
            const targetIndex = formState.fields.findIndex(f => f.id === over.id);
            if (targetIndex !== -1) {
              insertIndex = targetIndex + 1;
            }
          }
          
          const newFields = [...formState.fields];
          newFields.splice(insertIndex, 0, newField);
          
          updateFormState({
            fields: newFields
          });
        } else if (formState.formType === "mixed") {
          const targetSectionId = over.id !== "form-canvas" ? String(over.id) : undefined;
          addFieldToMixedSection(newField, targetSectionId);
        }
      } else if (active.data.current?.type === "table") {
        // Handle table section drop
        const newTableConfig: TableConfig = {
          id: uuidv4(),
          name: `table_${Date.now()}`,
          title: "New Table",
          description: "",
          columns: [
            {
              id: uuidv4(),
              name: "column1",
              label: "Column 1",
              type: "text",
              required: false,
              order: 0,
            }
          ],
          preFilledData: [], // Initialize empty pre-filled data array
          minRows: 1,
          maxRows: 100,
          allowAddRows: true,
          allowDeleteRows: true,
          defaultRows: 5,
          layout: { order: 0, width: "full" }
        };

        if (formState.formType === "table") {
          updateFormState({ tableConfig: newTableConfig });
        } else if (formState.formType === "mixed") {
          const newSection: FormSection = {
            id: uuidv4(),
            name: `table_section_${Date.now()}`,
            title: "Table Section",
            type: "table",
            tableConfig: newTableConfig,
            layout: { order: formState.sections.length, columnsPerRow: 1 },
          };
          
          updateFormState({
            sections: [...formState.sections, newSection]
          });
        }
      }
    }
    // Handle reordering existing elements
    else if (isReorderingElement && active.id !== over.id) {
      if (formState.formType === "regular") {
        const oldIndex = formState.fields.findIndex(field => field.id === active.id);
        const newIndex = formState.fields.findIndex(field => field.id === over.id);
        
        if (oldIndex !== -1 && newIndex !== -1 && oldIndex !== newIndex) {
          updateFormState({
            fields: arrayMove(formState.fields, oldIndex, newIndex)
          });
        }
      } else if (formState.formType === "mixed") {
        const oldIndex = formState.sections.findIndex(section => section.id === active.id);
        const newIndex = formState.sections.findIndex(section => section.id === over.id);
        
        if (oldIndex !== -1 && newIndex !== -1 && oldIndex !== newIndex) {
          updateFormState({
            sections: arrayMove(formState.sections, oldIndex, newIndex)
          });
        }
      }
    }
  };

  // Add field quickly without drag from palette
  const handleAddFieldFromTemplate = (template: FieldTemplate) => {
    const fieldId = uuidv4();
    const newField: FormField = {
      id: fieldId,
      name: fieldId, // Use unique ID as name to prevent conflicts
      label: template.label,
      ...template.defaultProps,
    } as FormField;

    if (formState.formType === "regular") {
      updateFormState({ fields: [...formState.fields, newField] });
    } else if (formState.formType === "mixed") {
      addFieldToMixedSection(newField);
    }
  };

  // Add a table section quickly without drag
  const handleAddTableSection = () => {
    const newTableConfig: TableConfig = {
      id: uuidv4(),
      name: `table_${Date.now()}`,
      title: "New Table",
      description: "",
      columns: [
        {
          id: uuidv4(),
          name: "column1",
          label: "Column 1",
          type: "text",
          required: false,
          order: 0,
        }
      ],
      minRows: 1,
      maxRows: 100,
      allowAddRows: true,
      allowDeleteRows: true,
      defaultRows: 5,
      layout: { order: 0, width: "full" }
    };

    if (formState.formType === "table") {
      updateFormState({ tableConfig: newTableConfig });
    } else if (formState.formType === "mixed") {
      const newSection: FormSection = {
        id: uuidv4(),
        name: `table_section_${Date.now()}`,
        title: "Table Section",
        type: "table",
        tableConfig: newTableConfig,
        layout: { order: formState.sections.length, columnsPerRow: 1 },
      };
      updateFormState({ sections: [...formState.sections, newSection] });
    }
  };

  // Add an empty fields section (useful for mixed form composition)
  const handleAddMixedSection = () => {
    if (formState.formType !== "mixed") return;

    const newSection: FormSection = {
      id: uuidv4(),
      name: `section_${Date.now()}`,
      title: `Section ${formState.sections.length + 1}`,
      type: "fields",
      fields: [],
      layout: { order: formState.sections.length, columnsPerRow: 1 },
    };

    updateFormState({ sections: [...formState.sections, newSection] });
  };

  const handleUpdateField = (fieldId: string, updates: Partial<FormField>) => {
    if (formState.formType === "regular") {
      const updatedFields = formState.fields.map(field =>
        field.id === fieldId ? { ...field, ...updates } : field
      );
      updateFormState({ fields: updatedFields });
      
      // Update selected item if it's the same field
      if (selectedItem?.id === fieldId) {
        setSelectedItem({
          ...selectedItem,
          data: { ...selectedItem.data, ...updates } as FormField
        });
      }
    } else if (formState.formType === "mixed") {
      // Handle fields within sections
      const updatedSections = formState.sections.map(section => {
        if (section.fields?.some(field => field.id === fieldId)) {
          return {
            ...section,
            fields: section.fields?.map(field =>
              field.id === fieldId ? { ...field, ...updates } : field
            )
          };
        }
        return section;
      });
      updateFormState({ sections: updatedSections });
      
      // Update selected item if it's the same field
      if (selectedItem?.id === fieldId) {
        setSelectedItem({
          ...selectedItem,
          data: { ...selectedItem.data, ...updates } as FormField
        });
      }
    }
  };

  const handleDeleteField = (fieldId: string) => {
    if (formState.formType === "regular") {
      updateFormState({
        fields: formState.fields.filter(field => field.id !== fieldId)
      });
    } else if (formState.formType === "mixed") {
      // Handle fields within sections
      const updatedSections = formState.sections.map(section => {
        if (section.fields?.some(field => field.id === fieldId)) {
          return {
            ...section,
            fields: section.fields?.filter(field => field.id !== fieldId)
          };
        }
        return section;
      });
      updateFormState({ sections: updatedSections });
    }
    
    if (selectedItem?.id === fieldId) {
      setSelectedItem(null);
    }
  };

  const handleUpdateSection = (sectionId: string, updates: Partial<FormSection>) => {
    const updatedSections = formState.sections.map(section => {
      if (section.id === sectionId) {
        const updated = { ...section, ...updates };
        // Clean up: table sections should not have fields property
        if (updated.type === "table" && "fields" in updated) {
          delete updated.fields;
        }
        return updated;
      }
      return section;
    });
    updateFormState({ sections: updatedSections });
    
    if (selectedItem?.id === sectionId) {
      setSelectedItem({
        ...selectedItem,
        data: { ...selectedItem.data, ...updates } as FormSection
      });
    }
  };

  const handleDeleteSection = (sectionId: string) => {
    updateFormState({
      sections: formState.sections.filter(section => section.id !== sectionId)
    });
    
    if (selectedItem?.id === sectionId) {
      setSelectedItem(null);
    }
  };

  const handleUpdateTable = (updates: Partial<TableConfig>) => {
    if (formState.tableConfig) {
      const updatedTable = { ...formState.tableConfig, ...updates };
      updateFormState({ tableConfig: updatedTable });
      
      if (selectedItem?.id === formState.tableConfig.id) {
        setSelectedItem({
          ...selectedItem,
          data: updatedTable
        });
      }
    }
  };

  const handleUpdateTableById = (tableId: string, updates: Partial<TableConfig>) => {
    // Handle main table config (table form)
    if (formState.tableConfig?.id === tableId) {
      handleUpdateTable(updates);
      return;
    }

    // Handle tables within mixed form sections
    const updatedSections = formState.sections.map(section => {
      if (section.tableConfig?.id === tableId) {
        const updatedTableConfig = { ...section.tableConfig, ...updates };
        return { ...section, tableConfig: updatedTableConfig };
      }
      return section;
    });

    updateFormState({ sections: updatedSections });

    // Update selected item if it's the same table
    if (selectedItem?.id === tableId) {
      const updatedTable = formState.sections.find(s => s.tableConfig?.id === tableId)?.tableConfig;
      if (updatedTable) {
        setSelectedItem({
          ...selectedItem,
          data: { ...updatedTable, ...updates }
        });
      }
    }
  };

  const handleMoveField = (fieldId: string, direction: "up" | "down", context?: { sectionId?: string }) => {
    if (formState.formType === "regular") {
      const idx = formState.fields.findIndex(f => f.id === fieldId);
      if (idx === -1) return;
      const newIndex = direction === "up" ? Math.max(0, idx - 1) : Math.min(formState.fields.length - 1, idx + 1);
      if (newIndex === idx) return;
      updateFormState({ fields: arrayMove(formState.fields, idx, newIndex) });
    } else if (formState.formType === "mixed" && context?.sectionId) {
      const sectionsCopy = formState.sections.map(s => ({ ...s }));
      const sectionIndex = sectionsCopy.findIndex(s => s.id === context.sectionId);
      if (sectionIndex === -1) return;
      const section = sectionsCopy[sectionIndex];
      if (section.type !== "fields" || !section.fields) return;
      const idx = section.fields.findIndex(f => f.id === fieldId);
      if (idx === -1) return;
      const newIndex = direction === "up" ? Math.max(0, idx - 1) : Math.min(section.fields.length - 1, idx + 1);
      if (newIndex === idx) return;
      section.fields = arrayMove(section.fields, idx, newIndex);
      sectionsCopy[sectionIndex] = section;
      updateFormState({ sections: sectionsCopy });
    }
  };

  const handleSelectItem = (itemId: string, itemType: "field" | "section" | "table") => {
    let itemData: FormField | FormSection | TableConfig | undefined;
    
    if (itemType === "field") {
      itemData = formState.fields.find(f => f.id === itemId) || 
                 formState.sections.flatMap(s => s.fields || []).find(f => f.id === itemId);
    } else if (itemType === "section") {
      itemData = formState.sections.find(s => s.id === itemId);
    } else if (itemType === "table") {
      // Check if it's the main table config (table form)
      if (formState.tableConfig?.id === itemId) {
        itemData = formState.tableConfig;
      } else {
        // Check if it's a table within a mixed form section
        itemData = formState.sections.find(s => s.tableConfig?.id === itemId)?.tableConfig;
      }
    }
    
    if (itemData) {
      setSelectedItem({ id: itemId, type: itemType, data: itemData });
    }
  };

  const handleSubmit = (e: React.FormEvent) => {
    e.preventDefault();

    if (!formState.title.trim() || !formState.category) {
      return;
    }

    const submitData: any = {
      title: formState.title,
      description: formState.description,
      category: formState.category,
      formType: formState.formType,
      validityPeriod: formState.validityPeriod,
      layout: formState.layout,
    };

    if (formState.formType === "regular") {
      submitData.fields = formState.fields.map(sanitizeFieldForSubmit);
    } else if (formState.formType === "table") {
      submitData.tableConfig = formState.tableConfig;
    } else if (formState.formType === "mixed") {
      // Clean sections before submitting - ensure table sections don't have fields property
      const cleanedSections = formState.sections.map(section => {
        if (section.type === "table") {
          const { fields, ...cleanSection } = section;
          return cleanSection;
        }
        if (section.type === "fields") {
          return {
            ...section,
            fields: (section.fields || []).map(sanitizeFieldForSubmit),
          };
        }
        return section;
      });
      submitData.sections = cleanedSections;
    }

    if (isEditing) {
      updateMutation.mutate(submitData);
    } else {
      createMutation.mutate(submitData);
    }
  };

  const error = createMutation.error || updateMutation.error;

  return (
    <div className="min-h-screen bg-gray-50 dark:bg-sena-darkBg">
      <DndContext
        sensors={sensors}
        collisionDetection={closestCenter}
        onDragStart={handleDragStart}
        onDragEnd={handleDragEnd}
      >
        {/* Header */}
        <div className="bg-white border-b dark:bg-sena-darkCard dark:border-gray-700">
          <div className="px-3 py-4 mx-auto sm:px-4 max-w-7xl">
            <div className="flex flex-wrap items-center justify-between gap-4">
              <div className="flex items-center space-x-2 sm:space-x-4">
                <Button variant="outline" size="icon" onClick={() => navigate("/admin/forms")}>
                  <ArrowLeft className="w-4 h-4" />
                </Button>
                <h1 className="text-lg font-bold sm:text-2xl text-sena-navy dark:text-white">
                  {isEditing ? "Edit Form" : "Create New Form"}
                </h1>
              </div>
              
              <div className="flex items-center space-x-1 sm:space-x-2">
                <Button
                  variant="outline"
                  size="sm"
                  onClick={undo}
                  disabled={historyIndex === 0}
                  className="p-2 sm:px-3"
                >
                  <Undo className="w-4 h-4" />
                </Button>
                <Button
                  variant="outline"
                  size="sm"
                  onClick={redo}
                  disabled={historyIndex >= history.length - 1}
                  className="p-2 sm:px-3"
                >
                  <Redo className="w-4 h-4" />
                </Button>
                <Dialog>
                  <DialogTrigger asChild>
                    <Button variant="outline" size="sm" className="hidden sm:flex">
                      <Eye className="w-4 h-4 mr-2" />
                      Preview
                    </Button>
                  </DialogTrigger>
                  <DialogTrigger asChild>
                    <Button variant="outline" size="sm" className="p-2 sm:hidden">
                      <Eye className="w-4 h-4" />
                    </Button>
                  </DialogTrigger>
                  <DialogContent className="max-w-4xl max-h-[80vh] overflow-y-auto mx-4">
                    <DialogHeader>
                      <DialogTitle>Form Preview - {formState.title}</DialogTitle>
                    </DialogHeader>
                    <div className="mt-4">
                      <PreviewForm formState={formState} />
                    </div>
                  </DialogContent>
                </Dialog>
                <Button 
                  onClick={handleSubmit} 
                  disabled={createMutation.isLoading || updateMutation.isLoading}
                  size="sm"
                  className="text-xs sm:text-sm"
                >
                  <Save className="w-4 h-4 mr-1 sm:mr-2" />
                  <span className="hidden sm:inline">{isEditing ? "Update Form" : "Create Form"}</span>
                  <span className="sm:hidden">Save</span>
                </Button>
              </div>
            </div>
            
            {!!error && (
              <div className="p-3 mt-4 text-red-700 bg-red-100 border border-red-400 rounded">
                <div className="font-semibold">
                  {(error as any)?.response?.data?.message || "An error occurred while saving the form."}
                </div>
                {(error as any)?.response?.data?.errors && Array.isArray((error as any).response.data.errors) && (error as any).response.data.errors.length > 0 && (
                  <ul className="mt-2 list-disc list-inside">
                    {(error as any).response.data.errors.map((err: { field?: string; message: string }, index: number) => (
                      <li key={index} className="text-sm">
                        {err.field ? `${err.field}: ` : ""}{err.message}
                      </li>
                    ))}
                  </ul>
                )}
              </div>
            )}
          </div>
        </div>

        {/* Main Content */}
        <div className="p-3 mx-auto sm:p-4 max-w-7xl">
          <div className="grid grid-cols-1 gap-4 lg:grid-cols-12 lg:gap-6">
            {/* Left Sidebar - Field Palette (Hidden on mobile, can be made into a modal) */}
            <div className="hidden lg:block lg:col-span-3">
              <div className="sticky top-4">
                <FieldPalette onAddTemplate={handleAddFieldFromTemplate} onAddTableSection={handleAddTableSection} />
              </div>
            </div>

            {/* Center - Form Settings & Canvas */}
            <div className="space-y-4 lg:col-span-6 lg:space-y-6">
              {/* Mobile Field Palette Button */}
              <div className="lg:hidden">
                <Dialog>
                  <DialogTrigger asChild>
                    <Button variant="outline" className="w-full">
                      Add Fields to Form
                    </Button>
                  </DialogTrigger>
                  <DialogContent className="max-w-md max-h-[80vh] overflow-y-auto">
                    <DialogHeader>
                      <DialogTitle>Add Form Fields</DialogTitle>
                    </DialogHeader>
                    <div className="mt-4">
                      <FieldPalette onAddTemplate={handleAddFieldFromTemplate} onAddTableSection={handleAddTableSection} />
                    </div>
                  </DialogContent>
                </Dialog>
              </div>

              {/* Form Settings */}
              <Card className="border-sena-lightBlue/20 dark:border-gray-700 dark:bg-sena-darkCard">
                <CardHeader>
                  <CardTitle className="text-sena-navy dark:text-white">Form Settings</CardTitle>
                </CardHeader>
                <CardContent className="space-y-4">
                  <div className="grid grid-cols-1 gap-4 sm:grid-cols-2">
                    <div>
                      <Label htmlFor="title" className="text-sena-navy dark:text-white">Form Title</Label>
                      <Input
                        id="title"
                        value={formState.title}
                        onChange={(e) => updateFormState({ title: e.target.value })}
                        required
                        className="border-sena-lightBlue/30 focus:border-sena-gold focus:ring-sena-gold/20"
                      />
                    </div>
                    <div>
                      <Label htmlFor="category" className="text-sena-navy dark:text-white">Category</Label>
                      <Select
                        value={formState.category}
                        onValueChange={(value) => updateFormState({ category: value })}
                      >
                        <SelectTrigger className="border-sena-lightBlue/30 focus:border-sena-gold focus:ring-sena-gold/20">
                          <SelectValue placeholder="Select category" />
                        </SelectTrigger>
                        <SelectContent>
                          {Array.isArray(categories?.data?.data) ? categories.data.data.map((category: Category) => (
                            <SelectItem key={category._id} value={category._id}>
                              {category.displayName}
                            </SelectItem>
                          )) : []}
                        </SelectContent>
                      </Select>
                    </div>
                  </div>

                  <div>
                    <Label htmlFor="description" className="text-sena-navy dark:text-white">Description</Label>
                    <Textarea
                      id="description"
                      value={formState.description}
                      onChange={(e) => updateFormState({ description: e.target.value })}
                      rows={2}
                      className="border-sena-lightBlue/30 focus:border-sena-gold focus:ring-sena-gold/20"
                    />
                  </div>

                  <div className="grid grid-cols-1 gap-4 sm:grid-cols-3">
                    <div>
                      <Label htmlFor="formType" className="text-sena-navy dark:text-white">Form Type</Label>
                      <Select
                        value={formState.formType}
                        onValueChange={(value: "regular" | "table" | "mixed") => 
                          updateFormState({ 
                            formType: value,
                            fields: value === "regular" ? formState.fields : [],
                            sections: value === "mixed" ? formState.sections : [],
                            tableConfig: value === "table" ? formState.tableConfig : undefined,
                          })
                        }
                      >
                        <SelectTrigger className="border-sena-lightBlue/30 focus:border-sena-gold focus:ring-sena-gold/20">
                          <SelectValue />
                        </SelectTrigger>
                        <SelectContent>
                          <SelectItem value="regular">Regular Form</SelectItem>
                          <SelectItem value="table">Table Form</SelectItem>
                          <SelectItem value="mixed">Mixed Form</SelectItem>
                        </SelectContent>
                      </Select>
                    </div>
                    <div>
                      <Label htmlFor="validityPeriod" className="text-sena-navy dark:text-white">Validity Period (days)</Label>
                      <Input
                        id="validityPeriod"
                        type="number"
                        min="1"
                        max="365"
                        value={formState.validityPeriod}
                        onChange={(e) => updateFormState({ validityPeriod: parseInt(e.target.value) })}
                        required
                        className="border-sena-lightBlue/30 focus:border-sena-gold focus:ring-sena-gold/20"
                      />
                    </div>
                    <div>
                      <Label htmlFor="theme" className="text-sena-navy dark:text-white">Theme</Label>
                      <Select
                        value={formState.layout.theme}
                        onValueChange={(value: FormLayout["theme"]) => 
                          updateFormState({ 
                            layout: { ...formState.layout, theme: value }
                          })
                        }
                      >
                        <SelectTrigger className="border-sena-lightBlue/30 focus:border-sena-gold focus:ring-sena-gold/20">
                          <SelectValue />
                        </SelectTrigger>
                        <SelectContent>
                          <SelectItem value="default">Default</SelectItem>
                          <SelectItem value="modern">Modern</SelectItem>
                          <SelectItem value="minimal">Minimal</SelectItem>
                          <SelectItem value="professional">Professional</SelectItem>
                        </SelectContent>
                      </Select>
                    </div>
                  </div>
                </CardContent>
              </Card>

              {/* Form Canvas */}
              <Card className="border-sena-lightBlue/20 dark:border-gray-700 dark:bg-sena-darkCard">
                <CardHeader>
                  <div className="flex items-center justify-between gap-2">
                    <CardTitle className="text-sena-navy dark:text-white">Form Builder</CardTitle>
                    {formState.formType === "mixed" && (
                      <Button type="button" variant="outline" size="sm" onClick={handleAddMixedSection}>
                        <Plus className="w-4 h-4 mr-1" />
                        Add Section
                      </Button>
                    )}
                  </div>
                  <CardDescription className="text-sena-lightBlue dark:text-white/90">
                    Drag and drop elements {window.innerWidth >= 1024 ? 'from the left panel' : 'using the Add Fields button'} to build your form
                  </CardDescription>
                </CardHeader>
                <CardContent>
                  <FormCanvas
                    formType={formState.formType}
                    fields={formState.fields}
                    sections={formState.sections}
                    tableConfig={formState.tableConfig}
                    onUpdateField={handleUpdateField}
                    onDeleteField={handleDeleteField}
                    onUpdateSection={handleUpdateSection}
                    onDeleteSection={handleDeleteSection}
                    onUpdateTable={handleUpdateTable}
                    onSelectItem={handleSelectItem}
                    selectedItemId={selectedItem?.id}
                    onMoveField={handleMoveField}
                  />
                </CardContent>
              </Card>
            </div>

            {/* Right Sidebar - Properties Panel (Hidden on mobile, shown in a separate section) */}
            <div className="hidden lg:block lg:col-span-3">
              <div className="sticky top-4">
                <PropertiesPanel
                  selectedItem={selectedItem}
                  onUpdate={(updates) => {
                    if (selectedItem) {
                      if (selectedItem.type === "field") {
                        handleUpdateField(selectedItem.id, updates);
                      } else if (selectedItem.type === "section") {
                        handleUpdateSection(selectedItem.id, updates);
                      } else if (selectedItem.type === "table") {
                        handleUpdateTableById(selectedItem.id, updates);
                      }
                    }
                  }}
                />
              </div>
            </div>
          </div>

          {/* Mobile Properties Panel */}
          <div className="mt-4 lg:hidden">
            {selectedItem && (
              <Card className="border-sena-lightBlue/20 dark:border-gray-700 dark:bg-sena-darkCard">
                <CardHeader>
                  <CardTitle className="text-sena-navy dark:text-white">Properties</CardTitle>
                  <CardDescription className="text-sena-lightBlue dark:text-white/90">
                    Modify the selected element
                  </CardDescription>
                </CardHeader>
                <CardContent>
                  <PropertiesPanel
                    selectedItem={selectedItem}
                    onUpdate={(updates) => {
                      if (selectedItem) {
                        if (selectedItem.type === "field") {
                          handleUpdateField(selectedItem.id, updates);
                        } else if (selectedItem.type === "section") {
                          handleUpdateSection(selectedItem.id, updates);
                        } else if (selectedItem.type === "table") {
                          handleUpdateTableById(selectedItem.id, updates);
                        }
                      }
                    }}
                  />
                </CardContent>
              </Card>
            )}
          </div>
        </div>
      </DndContext>
    </div>
  );
};

export default FormBuilder;