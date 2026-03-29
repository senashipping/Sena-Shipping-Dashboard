import React from "react";
import {
  Document,
  Page,
  Text,
  View,
  StyleSheet,
  Image,
} from "@react-pdf/renderer";
import type { FormField, FormSection, TableColumn, TableConfig } from "../../types";

const styles = StyleSheet.create({
  page: {
    padding: 36,
    fontSize: 10,
    fontFamily: "Helvetica",
  },
  title: { fontSize: 16, marginBottom: 8, fontFamily: "Helvetica-Bold" },
  subtitle: { fontSize: 9, color: "#444", marginBottom: 12 },
  sectionTitle: { fontSize: 12, marginTop: 12, marginBottom: 6, fontFamily: "Helvetica-Bold" },
  label: { fontSize: 9, fontFamily: "Helvetica-Bold", marginBottom: 2 },
  valueBox: {
    borderWidth: 1,
    borderColor: "#ccc",
    padding: 6,
    minHeight: 18,
    marginBottom: 8,
  },
  row: { flexDirection: "row", borderBottomWidth: 1, borderBottomColor: "#ddd" },
  cell: { flex: 1, padding: 4, fontSize: 8 },
  headerCell: {
    flex: 1,
    padding: 4,
    fontSize: 8,
    fontFamily: "Helvetica-Bold",
    backgroundColor: "#f0f0f0",
  },
  tableWrap: { marginTop: 6, borderWidth: 1, borderColor: "#ccc" },
  meta: { flexDirection: "row", flexWrap: "wrap", marginBottom: 8 },
  metaItem: { fontSize: 9, marginRight: 12, marginBottom: 4 },
});

export type FormTemplatePdfVariant = "blank" | "filled";

export interface FormTemplatePdfProps {
  title: string;
  description?: string;
  categoryLabel?: string;
  formType: "regular" | "table" | "mixed";
  fields?: FormField[];
  sections?: FormSection[];
  tableConfig?: TableConfig;
  formData?: Record<string, any>;
  tableData?: any[];
  variant: FormTemplatePdfVariant;
}

function formatFieldValue(
  field: FormField,
  raw: any,
  variant: FormTemplatePdfVariant
): { text?: string; imageSrc?: string } {
  if (variant === "blank" && field.type !== "signature") {
    return { text: "" };
  }
  if (raw === undefined || raw === null || raw === "") {
    if (variant === "blank") return { text: "" };
    return { text: "\u2014" };
  }

  switch (field.type) {
    case "checkbox": {
      if (!Array.isArray(raw)) return { text: String(raw) };
      const labels =
        field.options?.filter((o) => raw.includes(o.value)).map((o) => o.label) || [];
      return { text: labels.length ? labels.join(", ") : "\u2014" };
    }
    case "signature": {
      const s = String(raw);
      if (s.startsWith("data:image")) return { imageSrc: s };
      return { text: s || "\u2014" };
    }
    case "file":
      return { text: String(raw) };
    default:
      return { text: String(raw) };
  }
}

function FieldBlock({
  field,
  variant,
  valueOverride,
}: {
  field: FormField;
  variant: FormTemplatePdfVariant;
  valueOverride?: any;
}) {
  const raw = valueOverride;
  const { text, imageSrc } = formatFieldValue(field, raw, variant);
  const showEmpty = variant === "blank" || !text?.trim();

  return (
    <View wrap={false}>
      <Text style={styles.label}>
        {field.label}
        {field.required ? " *" : ""}
      </Text>
      {field.description ? (
        <Text style={{ fontSize: 8, color: "#666", marginBottom: 2 }}>{field.description}</Text>
      ) : null}
      {field.type === "signature" && imageSrc ? (
        <Image src={imageSrc} style={{ width: 140, height: 48, marginBottom: 8 }} />
      ) : (
        <View style={styles.valueBox}>
          <Text style={{ fontSize: 9 }}>
            {showEmpty && variant === "blank" ? " " : text || " "}
          </Text>
        </View>
      )}
    </View>
  );
}

function TablePdf({
  columns,
  rows,
  preFilledData,
  variant,
}: {
  columns: TableColumn[];
  preFilledData?: { rowIndex: number; columnName: string; value: any; isReadOnly: boolean }[];
  rows: any[];
  variant: FormTemplatePdfVariant;
}) {
  const sortedCols = [...columns].sort((a, b) => (a.order ?? 0) - (b.order ?? 0));

  const cellDisplay = (rowIndex: number, col: TableColumn, row: any) => {
    const pf = preFilledData?.find(
      (p) => p.rowIndex === rowIndex && p.columnName === col.name
    );
    const raw = pf ? pf.value : row?.[col.name];
    if (variant === "blank" && !pf) {
      return { text: "" };
    }
    if (col.type === "signature" && raw && String(raw).startsWith("data:image")) {
      return { imageSrc: String(raw) };
    }
    if (raw === undefined || raw === null || raw === "") {
      return { text: variant === "blank" ? "" : "\u2014" };
    }
    return { text: String(raw) };
  };

  return (
    <View style={styles.tableWrap}>
      <View style={[styles.row, { borderBottomWidth: 1 }]}>
        {sortedCols.map((c) => (
          <Text key={c.name} style={styles.headerCell}>
            {c.label}
          </Text>
        ))}
      </View>
      {rows.map((row, ri) => (
        <View key={ri} style={styles.row} wrap={false}>
          {sortedCols.map((col) => {
            const disp = cellDisplay(ri, col, row);
            return (
              <View key={col.name} style={styles.cell}>
                {"imageSrc" in disp && disp.imageSrc ? (
                  <Image src={disp.imageSrc!} style={{ width: 60, height: 22 }} />
                ) : (
                  <Text style={{ fontSize: 8 }}>{disp.text ?? ""}</Text>
                )}
              </View>
            );
          })}
        </View>
      ))}
    </View>
  );
}

function ensureTableRows(
  config: TableConfig,
  existing: any[] | undefined,
  variant: FormTemplatePdfVariant
): any[] {
  const n = Math.max(
    config.defaultRows ?? 1,
    config.minRows ?? 1,
    Array.isArray(existing) ? existing.length : 0
  );
  const base = Array.isArray(existing) ? [...existing] : [];
  while (base.length < n) base.push({});
  if (variant === "blank") {
    return base.map(() => ({}));
  }
  return base;
}

/** Inner form layout (title, fields, tables) for embedding in one or more PDF pages. */
export const FormTemplatePdfPageBody: React.FC<FormTemplatePdfProps> = (props) => {
  const {
    title,
    description,
    categoryLabel,
    formType,
    fields = [],
    sections = [],
    tableConfig,
    formData = {},
    tableData = [],
    variant,
  } = props;

  return (
    <>
      <Text style={styles.title}>{title || "Form"}</Text>
      <View style={styles.meta}>
        {categoryLabel ? (
          <Text style={styles.metaItem}>Category: {categoryLabel}</Text>
        ) : null}
        <Text style={styles.metaItem}>Type: {formType}</Text>
      </View>
      {description ? <Text style={styles.subtitle}>{description}</Text> : null}

      {formType === "regular" && (
        <View>
          <Text style={styles.sectionTitle}>Fields</Text>
          {[...fields]
            .sort((a, b) => (a.layout?.order ?? 0) - (b.layout?.order ?? 0))
            .map((field) => (
              <FieldBlock
                key={field.id}
                field={field}
                variant={variant}
                valueOverride={formData[field.name]}
              />
            ))}
        </View>
      )}

      {formType === "table" && tableConfig && (
        <View>
          <Text style={styles.sectionTitle}>{tableConfig.title || "Table"}</Text>
          {tableConfig.description ? (
            <Text style={{ fontSize: 9, marginBottom: 4 }}>{tableConfig.description}</Text>
          ) : null}
          <TablePdf
            columns={tableConfig.columns || []}
            rows={ensureTableRows(tableConfig, tableData, variant)}
            preFilledData={tableConfig.preFilledData}
            variant={variant}
          />
        </View>
      )}

      {formType === "mixed" &&
        sections
          .slice()
          .sort((a, b) => (a.layout?.order ?? 0) - (b.layout?.order ?? 0))
          .map((section) => (
            <View key={section.id} wrap={false}>
              <Text style={styles.sectionTitle}>{section.title}</Text>
              {section.description ? (
                <Text style={{ fontSize: 9, marginBottom: 4 }}>{section.description}</Text>
              ) : null}
              {section.type === "fields" && section.fields && (
                <View>
                  {section.fields.map((field) => (
                    <FieldBlock
                      key={field.id}
                      field={field}
                      variant={variant}
                      valueOverride={formData[field.name]}
                    />
                  ))}
                </View>
              )}
              {section.type === "table" && section.tableConfig && (
                <TablePdf
                  columns={section.tableConfig.columns || []}
                  rows={ensureTableRows(
                    section.tableConfig,
                    formData[`table_${section.id}`],
                    variant
                  )}
                  preFilledData={section.tableConfig.preFilledData}
                  variant={variant}
                />
              )}
            </View>
          ))}
    </>
  );
};

const FormTemplatePdfDocument: React.FC<FormTemplatePdfProps> = (props) => (
  <Document>
    <Page size="A4" style={styles.page}>
      <FormTemplatePdfPageBody {...props} />
    </Page>
  </Document>
);

export default FormTemplatePdfDocument;