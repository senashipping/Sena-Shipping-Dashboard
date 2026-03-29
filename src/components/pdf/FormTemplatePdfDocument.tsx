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

// ─── Colour palette ────────────────────────────────────────────────────────
const C = {
  navy:       "#0f2341",
  navyMid:    "#1a3558",
  teal:       "#0d7fa5",
  tealLight:  "#e6f4f9",
  tealBorder: "#7ac8df",
  silver:     "#f4f6f8",
  silver2:    "#e8ecf0",
  border:     "#cdd5dc",
  text:       "#1e2a38",
  muted:      "#5a6a7a",
  white:      "#ffffff",
  success:    "#166534",
  successBg:  "#dcfce7",
  warning:    "#92400e",
  warningBg:  "#fef3c7",
  accent:     "#e85d26",   // orange stripe
  danger:     "#b91c1c",
  dangerBg:   "#fee2e2",
};

// ─── Styles ────────────────────────────────────────────────────────────────
const s = StyleSheet.create({
  page: {
    fontFamily: "Helvetica",
    fontSize: 9,
    color: C.text,
    backgroundColor: C.white,
  },

  // ── Header band ──
  headerBand: {
    backgroundColor: C.navy,
    paddingHorizontal: 36,
    paddingTop: 22,
    paddingBottom: 18,
  },
  accentStripe: {
    position: "absolute",
    top: 0,
    left: 0,
    right: 0,
    height: 4,
    backgroundColor: C.teal,
  },
  /** Stacked layout: title block full width, category badge on next row (no horizontal clash). */
  headerBandStack: {
    flexDirection: "column",
    width: "100%",
  },
  headerTitleBlock: {
    width: "100%",
    paddingRight: 0,
  },
  headerCategoryRow: {
    marginTop: 10,
    flexDirection: "row",
    flexWrap: "wrap",
    alignItems: "center",
  },
  headerTitle: {
    fontSize: 20,
    fontFamily: "Helvetica-Bold",
    color: C.white,
    letterSpacing: 0.5,
  },
  headerSubtitle: {
    fontSize: 9,
    color: "#8eaec9",
    marginTop: 3,
  },
  headerBadge: {
    borderWidth: 1,
    borderColor: C.teal,
    borderRadius: 3,
    paddingHorizontal: 8,
    paddingVertical: 3,
  },
  headerBadgeText: {
    fontSize: 8,
    color: C.teal,
    fontFamily: "Helvetica-Bold",
    textTransform: "uppercase",
    letterSpacing: 1,
  },

  // ── Body wrapper (bottom clearance comes mainly from Page `paddingBottom`; keep this modest) ──
  body: {
    paddingHorizontal: 36,
    paddingTop: 12,
    paddingBottom: 12,
  },

  // ── Meta cards row ──
  metaRow: {
    flexDirection: "row",
    gap: 8,
    marginBottom: 18,
  },
  metaCard: {
    flex: 1,
    backgroundColor: C.silver,
    borderWidth: 1,
    borderColor: C.border,
    borderRadius: 4,
    padding: 10,
  },
  metaCardAccent: {
    flex: 1,
    backgroundColor: C.tealLight,
    borderWidth: 1,
    borderColor: C.tealBorder,
    borderRadius: 4,
    padding: 10,
  },
  metaLabel: {
    fontSize: 7,
    fontFamily: "Helvetica-Bold",
    color: C.muted,
    textTransform: "uppercase",
    letterSpacing: 0.5,
    marginBottom: 4,
  },
  metaValue: {
    fontSize: 9,
    fontFamily: "Helvetica-Bold",
    color: C.text,
  },

  // ── Status badge ──
  statusBadge: {
    alignSelf: "flex-start",
    borderRadius: 3,
    paddingHorizontal: 7,
    paddingVertical: 3,
  },
  statusText: {
    fontSize: 8,
    fontFamily: "Helvetica-Bold",
    textTransform: "uppercase",
    letterSpacing: 0.5,
  },

  // ── Divider ──
  divider: {
    height: 1,
    backgroundColor: C.border,
    marginVertical: 14,
  },
  dividerTeal: {
    height: 2,
    backgroundColor: C.teal,
    marginBottom: 14,
  },

  // ── Section heading ──
  sectionHeading: {
    flexDirection: "row",
    alignItems: "center",
    marginBottom: 10,
    marginTop: 6,
  },
  sectionPill: {
    width: 4,
    height: 16,
    backgroundColor: C.teal,
    borderRadius: 2,
    marginRight: 8,
  },
  sectionTitle: {
    fontSize: 11,
    fontFamily: "Helvetica-Bold",
    color: C.navy,
  },
  sectionDesc: {
    fontSize: 8,
    color: C.muted,
    marginBottom: 10,
    marginLeft: 12,
  },

  // ── Field blocks ──
  fieldLabel: {
    fontSize: 8,
    fontFamily: "Helvetica-Bold",
    color: C.navyMid,
    marginBottom: 3,
    textTransform: "uppercase",
    letterSpacing: 0.3,
  },
  fieldDesc: {
    fontSize: 7.5,
    color: C.muted,
    marginBottom: 3,
  },
  fieldBox: {
    borderWidth: 1,
    borderColor: C.border,
    borderRadius: 3,
    backgroundColor: C.silver,
    padding: 7,
    minHeight: 20,
    marginBottom: 10,
  },
  fieldBoxFilled: {
    borderWidth: 1,
    borderColor: C.tealBorder,
    borderRadius: 3,
    backgroundColor: C.tealLight,
    padding: 7,
    minHeight: 20,
    marginBottom: 10,
  },
  fieldValue: {
    fontSize: 9,
    color: C.text,
  },

  // ── Table ──
  tableWrap: {
    borderWidth: 1,
    borderColor: C.border,
    borderRadius: 4,
    overflow: "hidden",
    marginBottom: 10,
  },
  tableHeaderRow: {
    flexDirection: "row",
    backgroundColor: C.navy,
  },
  tableHeaderCell: {
    flex: 1,
    padding: 6,
    fontSize: 7.5,
    fontFamily: "Helvetica-Bold",
    color: C.white,
    textTransform: "uppercase",
    letterSpacing: 0.4,
  },
  tableRow: {
    flexDirection: "row",
    borderTopWidth: 1,
    borderTopColor: C.border,
    backgroundColor: C.white,
  },
  tableRowAlt: {
    flexDirection: "row",
    borderTopWidth: 1,
    borderTopColor: C.border,
    backgroundColor: C.silver,
  },
  tableCell: {
    flex: 1,
    padding: 5,
    fontSize: 8,
    color: C.text,
  },
  tableCellMuted: {
    flex: 1,
    padding: 5,
    fontSize: 8,
    color: "#aab4c0",
  },

  // ── Footer ──
  footer: {
    position: "absolute",
    bottom: 0,
    left: 0,
    right: 0,
    height: 28,
    backgroundColor: C.navy,
    flexDirection: "row",
    alignItems: "center",
    justifyContent: "space-between",
    paddingHorizontal: 36,
  },
  footerText: {
    fontSize: 7,
    color: "#8eaec9",
  },
  footerAccent: {
    fontSize: 7,
    color: C.teal,
    fontFamily: "Helvetica-Bold",
  },
});

// ─── Types ────────────────────────────────────────────────────────────────
export type FormTemplatePdfVariant = "blank" | "filled";

/** Space below fixed header so flow text starts under the band (tune if title wraps to 3+ lines). */
export const PDF_FORM_HEADER_RESERVE_PT = 108;
/** Space above bottom so flow does not draw under fixed footer (footer bar is 28pt + small buffer). */
export const PDF_FOOTER_RESERVE_PT = 38;

export interface FormTemplatePdfProps {
  title: string;
  description?: string;
  /** Shown under the title in the header (e.g. vessel name). */
  documentSubtitle?: string;
  categoryLabel?: string;
  formType: "regular" | "table" | "mixed";
  fields?: FormField[];
  sections?: FormSection[];
  tableConfig?: TableConfig;
  formData?: Record<string, any>;
  tableData?: any[];
  variant: FormTemplatePdfVariant;
  /** When true, header is omitted (use with `FormTemplatePdfFixedHeader` on the same Page). */
  omitHeader?: boolean;
}

// ─── Helpers ─────────────────────────────────────────────────────────────
function formatFieldValue(field: FormField, raw: any, variant: FormTemplatePdfVariant) {
  if (variant === "blank" && field.type !== "signature") return { text: "" };
  if (raw === undefined || raw === null || raw === "") {
    return { text: variant === "blank" ? "" : "\u2014" };
  }
  switch (field.type) {
    case "checkbox": {
      if (!Array.isArray(raw)) return { text: String(raw) };
      const labels = field.options?.filter((o) => raw.includes(o.value)).map((o) => o.label) || [];
      return { text: labels.length ? labels.join(", ") : "\u2014" };
    }
    case "signature": {
      const s2 = String(raw);
      if (s2.startsWith("data:image")) return { imageSrc: s2 };
      return { text: s2 || "\u2014" };
    }
    default:
      return { text: String(raw) };
  }
}

function statusStyle(status?: string) {
  if (!status) return null;
  const lower = status.toLowerCase();
  if (lower.includes("reject") || lower.includes("denied"))
    return { bg: C.dangerBg, fg: C.danger };
  if (lower.includes("expir"))
    return { bg: C.warningBg, fg: C.warning };
  if (lower.includes("approv") || lower.includes("complet"))
    return { bg: C.successBg, fg: C.success };
  if (lower.includes("pending") || lower.includes("review"))
    return { bg: C.warningBg, fg: C.warning };
  return { bg: C.silver2, fg: C.muted };
}

// ─── FieldBlock ───────────────────────────────────────────────────────────
function FieldBlock({
  field,
  variant,
  valueOverride,
}: {
  field: FormField;
  variant: FormTemplatePdfVariant;
  valueOverride?: any;
}) {
  const { text, imageSrc } = formatFieldValue(field, valueOverride, variant);
  const isFilled = variant === "filled" && text && text !== "\u2014";

  return (
    <View wrap style={{ marginBottom: 12 }}>
      <Text style={s.fieldLabel}>
        {field.label}
        {field.required ? " *" : ""}
      </Text>
      {field.description ? <Text style={s.fieldDesc}>{field.description}</Text> : null}
      {field.type === "signature" && imageSrc ? (
        <View style={{ borderWidth: 1, borderColor: C.border, borderRadius: 3, padding: 6, marginBottom: 10, backgroundColor: C.white }}>
          <Image src={imageSrc} style={{ width: 180, height: 56, objectFit: "contain" }} />
        </View>
      ) : (
        <View style={isFilled ? s.fieldBoxFilled : s.fieldBox}>
          <Text style={s.fieldValue}>{text || " "}</Text>
        </View>
      )}
    </View>
  );
}

// ─── TablePdf ─────────────────────────────────────────────────────────────
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
  const sorted = [...columns].sort((a, b) => (a.order ?? 0) - (b.order ?? 0));

  const cellDisplay = (ri: number, col: TableColumn, row: any) => {
    const pf = preFilledData?.find((p) => p.rowIndex === ri && p.columnName === col.name);
    const raw = pf ? pf.value : row?.[col.name];
    if (variant === "blank" && !pf) return { text: "" };
    if (col.type === "signature" && raw && String(raw).startsWith("data:image"))
      return { imageSrc: String(raw) };
    if (raw === undefined || raw === null || raw === "")
      return { text: variant === "blank" ? "" : "\u2014" };
    return { text: String(raw) };
  };

  return (
    <View style={s.tableWrap}>
      {/* Header */}
      <View style={s.tableHeaderRow}>
        {sorted.map((c) => (
          <Text key={c.name} style={s.tableHeaderCell}>
            {c.label}
          </Text>
        ))}
      </View>
      {/* Rows */}
      {rows.map((row, ri) => (
        <View key={ri} style={ri % 2 === 0 ? s.tableRow : s.tableRowAlt} wrap={false}>
          {sorted.map((col) => {
            const disp = cellDisplay(ri, col, row);
            const isEmpty = !disp.text || disp.text === "\u2014";
            return (
              <View key={col.name} style={isEmpty ? s.tableCellMuted : s.tableCell}>
                {"imageSrc" in disp && disp.imageSrc ? (
                  <Image src={disp.imageSrc!} style={{ width: 60, height: 22 }} />
                ) : (
                  <Text>{disp.text ?? ""}</Text>
                )}
              </View>
            );
          })}
        </View>
      ))}
    </View>
  );
}

function ensureTableRows(config: TableConfig, existing: any[] | undefined, variant: FormTemplatePdfVariant): any[] {
  const n = Math.max(config.defaultRows ?? 1, config.minRows ?? 1, Array.isArray(existing) ? existing.length : 0);
  const base = Array.isArray(existing) ? [...existing] : [];
  while (base.length < n) base.push({});
  return variant === "blank" ? base.map(() => ({})) : base;
}

// ─── Section heading helper ───────────────────────────────────────────────
function SectionHeading({ title }: { title: string }) {
  return (
    <View style={s.sectionHeading}>
      <View style={s.sectionPill} />
      <Text style={s.sectionTitle}>{title}</Text>
    </View>
  );
}

// ─── Shared footer (repeats on every page when placed inside Page) ────────
export const PdfPageFooter: React.FC = () => (
  <View style={s.footer} fixed>
    <Text style={s.footerText}>SENA Ship Management — Confidential</Text>
    <Text style={s.footerText}>
      Page <Text render={({ pageNumber, totalPages }) => `${pageNumber} / ${totalPages}`} />
    </Text>
    <Text style={s.footerAccent}>senashipping.com</Text>
  </View>
);

/** Repeats on every printed page (use with Page paddingTop ≈ PDF_FORM_HEADER_RESERVE_PT). */
export const FormTemplatePdfFixedHeader: React.FC<
  Pick<FormTemplatePdfProps, "title" | "description" | "documentSubtitle" | "categoryLabel">
> = ({ title, description, documentSubtitle, categoryLabel }) => (
  <View style={s.headerBand} fixed>
    <View style={s.accentStripe} />
    <View style={s.headerBandStack}>
      <View style={s.headerTitleBlock}>
        <Text style={s.headerTitle}>{title || "Form"}</Text>
        {documentSubtitle ? (
          <Text style={s.headerSubtitle}>Vessel: {documentSubtitle}</Text>
        ) : null}
        {description ? <Text style={s.headerSubtitle}>{description}</Text> : null}
      </View>
      {categoryLabel ? (
        <View style={s.headerCategoryRow}>
          <View style={s.headerBadge}>
            <Text style={s.headerBadgeText}>{categoryLabel}</Text>
          </View>
        </View>
      ) : null}
    </View>
  </View>
);

// ─── Page 1: Submission Record ────────────────────────────────────────────
interface SubmissionRecordProps {
  submittedBy?: string;
  ship?: string;
  submittedAt?: string;
  /** Used with statusStyle (raw or display). */
  status?: string;
  /** Human-readable label for the status badge. */
  statusLabel?: string;
  formTitle?: string;
  reviewedBy?: string;
  reviewedAt?: string;
  reviewComments?: string;
}

export const SubmissionRecordPage: React.FC<SubmissionRecordProps> = ({
  submittedBy,
  ship,
  submittedAt,
  status,
  statusLabel,
  formTitle,
  reviewedBy,
  reviewedAt,
  reviewComments,
}) => {
  const st = statusStyle(statusLabel || status);
  const badgeText = statusLabel || status || "\u2014";
  return (
    <Page size="A4" style={[s.page, { paddingBottom: PDF_FOOTER_RESERVE_PT }]}>
      {/* Header */}
      <View style={s.headerBand}>
        <View style={s.accentStripe} />
        <View style={s.headerBandStack}>
          <View style={s.headerTitleBlock}>
            <Text style={s.headerTitle}>Submission Record</Text>
            {formTitle && <Text style={s.headerSubtitle}>{formTitle}</Text>}
          </View>
          <View style={s.headerCategoryRow}>
            <View style={s.headerBadge}>
              <Text style={s.headerBadgeText}>SENA Ship Management</Text>
            </View>
          </View>
        </View>
      </View>

      {/* Body */}
      <View style={s.body}>
        <View style={s.dividerTeal} />

        {/* Meta cards */}
        <View style={s.metaRow}>
          <View style={s.metaCardAccent}>
            <Text style={s.metaLabel}>Submitted By</Text>
            <Text style={s.metaValue}>{submittedBy || "\u2014"}</Text>
          </View>
          <View style={s.metaCard}>
            <Text style={s.metaLabel}>Ship / Vessel</Text>
            <Text style={s.metaValue}>{ship || "\u2014"}</Text>
          </View>
        </View>

        <View style={s.metaRow}>
          <View style={s.metaCard}>
            <Text style={s.metaLabel}>Submitted At</Text>
            <Text style={s.metaValue}>{submittedAt || "\u2014"}</Text>
          </View>
          <View style={s.metaCard}>
            <Text style={s.metaLabel}>Status</Text>
            {st ? (
              <View style={[s.statusBadge, { backgroundColor: st.bg, marginTop: 2 }]}>
                <Text style={[s.statusText, { color: st.fg }]}>{badgeText}</Text>
              </View>
            ) : (
              <Text style={s.metaValue}>{badgeText}</Text>
            )}
          </View>
        </View>

        {(reviewedBy || reviewedAt || reviewComments) && (
          <>
            <View style={s.divider} />
            <View style={s.sectionHeading}>
              <View style={s.sectionPill} />
              <Text style={s.sectionTitle}>Review</Text>
            </View>
            <View style={s.metaRow}>
              {reviewedBy ? (
                <View style={s.metaCard}>
                  <Text style={s.metaLabel}>Reviewed By</Text>
                  <Text style={s.metaValue}>{reviewedBy}</Text>
                </View>
              ) : null}
              {reviewedAt ? (
                <View style={s.metaCard}>
                  <Text style={s.metaLabel}>Reviewed At</Text>
                  <Text style={s.metaValue}>{reviewedAt}</Text>
                </View>
              ) : null}
            </View>
            {reviewComments ? (
              <View
                style={{
                  marginTop: 10,
                  backgroundColor: C.silver,
                  borderWidth: 1,
                  borderColor: C.border,
                  borderRadius: 4,
                  padding: 10,
                }}
              >
                <Text style={s.metaLabel}>Comments</Text>
                <Text style={{ fontSize: 9, color: C.text, marginTop: 4 }}>{reviewComments}</Text>
              </View>
            ) : null}
          </>
        )}

        <View style={s.divider} />

        {/* Decorative info note */}
        <View
          style={{
            backgroundColor: C.silver,
            borderLeftWidth: 3,
            borderLeftColor: C.teal,
            borderRadius: 3,
            padding: 10,
            marginTop: 4,
          }}
        >
          <Text style={{ fontSize: 8, color: C.muted }}>
            This submission record was automatically generated by the SENA Ship Management
            platform. Please retain this document for your records and compliance audits.
          </Text>
        </View>
      </View>

      <PdfPageFooter />
    </Page>
  );
};

// ─── Page 2+: Form Body ───────────────────────────────────────────────────
export const FormTemplatePdfPageBody: React.FC<FormTemplatePdfProps> = (props) => {
  const {
    omitHeader,
    title,
    description,
    documentSubtitle,
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
      {!omitHeader ? (
        <View style={s.headerBand}>
          <View style={s.accentStripe} />
          <View style={s.headerBandStack}>
            <View style={s.headerTitleBlock}>
              <Text style={s.headerTitle}>{title || "Form"}</Text>
              {documentSubtitle ? (
                <Text style={s.headerSubtitle}>Vessel: {documentSubtitle}</Text>
              ) : null}
              {description ? <Text style={s.headerSubtitle}>{description}</Text> : null}
            </View>
            {categoryLabel ? (
              <View style={s.headerCategoryRow}>
                <View style={s.headerBadge}>
                  <Text style={s.headerBadgeText}>{categoryLabel}</Text>
                </View>
              </View>
            ) : null}
          </View>
        </View>
      ) : null}

      {/* Body */}
      <View style={s.body}>
        <View style={s.dividerTeal} />

        {/* Regular fields */}
        {formType === "regular" && (
          <View>
            <SectionHeading title="Fields" />
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

        {/* Table form */}
        {formType === "table" && tableConfig && (
          <View>
            <SectionHeading title={tableConfig.title || "Table"} />
            {tableConfig.description && (
              <Text style={s.sectionDesc}>{tableConfig.description}</Text>
            )}
            <TablePdf
              columns={tableConfig.columns || []}
              rows={ensureTableRows(tableConfig, tableData, variant)}
              preFilledData={tableConfig.preFilledData}
              variant={variant}
            />
          </View>
        )}

        {/* Mixed form */}
        {formType === "mixed" &&
          sections
            .slice()
            .sort((a, b) => (a.layout?.order ?? 0) - (b.layout?.order ?? 0))
            .map((section, idx) => (
              <View key={section.id} wrap={false}>
                {idx > 0 && <View style={s.divider} />}
                <SectionHeading title={section.title} />
                {section.description && (
                  <Text style={s.sectionDesc}>{section.description}</Text>
                )}
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
      </View>
    </>
  );
};

// ─── Full Document ────────────────────────────────────────────────────────
const FormTemplatePdfDocument: React.FC<FormTemplatePdfProps> = (props) => (
  <Document>
    <Page size="A4" style={[s.page, { paddingBottom: PDF_FOOTER_RESERVE_PT }]}>
      <FormTemplatePdfPageBody {...props} />
      <PdfPageFooter />
    </Page>
  </Document>
);

/** Shared StyleSheet for `<Page style={pdfStyles.page}>` in submission exports. */
export const pdfStyles = s;

export default FormTemplatePdfDocument;