import type { FormTemplatePdfProps } from "./FormTemplatePdfDocument";

function resolveCategoryLabel(form: { category?: unknown }): string | undefined {
  const cat = form.category;
  if (cat == null) return undefined;
  if (typeof cat === "object" && cat !== null) {
    if ("displayName" in cat && (cat as { displayName?: string }).displayName) {
      return String((cat as { displayName: string }).displayName);
    }
    if ("name" in cat && (cat as { name?: string }).name) {
      return String((cat as { name: string }).name);
    }
  }
  return String(cat);
}

/** Map API submission + form into props for FormTemplatePdfDocument (filled). */
export function submissionToFormTemplatePdfProps(submission: {
  form?: Record<string, any> | null;
  data?: Record<string, any> | null;
}): FormTemplatePdfProps | null {
  const form = submission.form;
  if (!form) return null;

  const raw = submission.data;
  const data =
    raw && typeof raw === "object" && !Array.isArray(raw) ? { ...raw } : {};

  const formType = (form.formType as FormTemplatePdfProps["formType"]) || "regular";

  let formData: Record<string, any> = {};
  let tableData: any[] = [];

  if (formType === "table") {
    tableData = Array.isArray(data.tableData) ? data.tableData : [];
  } else {
    formData = data;
  }

  return {
    title: String(form.title || "Form"),
    description: form.description,
    categoryLabel: resolveCategoryLabel(form),
    formType,
    fields: form.fields,
    sections: form.sections,
    tableConfig: form.tableConfig,
    variant: "filled",
    formData,
    tableData,
  };
}
