import React from "react";
import { Document, Page, Text, View, StyleSheet } from "@react-pdf/renderer";
import {
  FormTemplatePdfPageBody,
  PdfPageFooter,
  SubmissionRecordPage,
  pdfStyles,
} from "./FormTemplatePdfDocument";
import { submissionToFormTemplatePdfProps } from "./submissionToFormTemplatePdfProps";
import { formatDateTime } from "../../lib/utils";

// ─── Colour tokens (kept in sync with FormTemplatePdfDocument) ─────────────
const C = {
  navy:      "#0f2341",
  teal:      "#0d7fa5",
  tealLight: "#e6f4f9",
  tealBorder:"#7ac8df",
  silver:    "#f4f6f8",
  border:    "#cdd5dc",
  text:      "#1e2a38",
  muted:     "#5a6a7a",
  white:     "#ffffff",
};

// ─── Local styles for pieces specific to this file ─────────────────────────
const ls = StyleSheet.create({
  // ── Fallback page ──
  fallbackBody: {
    paddingHorizontal: 36,
    paddingTop: 24,
    paddingBottom: 48,
  },
  fallbackCard: {
    borderWidth: 1,
    borderColor: C.tealBorder,
    borderRadius: 4,
    backgroundColor: C.tealLight,
    padding: 16,
    flexDirection: "row",
    alignItems: "flex-start",
    gap: 10,
  },
  fallbackIconWrap: {
    width: 20,
    height: 20,
    borderRadius: 10,
    backgroundColor: C.teal,
    alignItems: "center",
    justifyContent: "center",
    flexShrink: 0,
  },
  fallbackIcon: {
    fontSize: 11,
    color: C.white,
    fontFamily: "Helvetica-Bold",
  },
  fallbackText: {
    fontSize: 9,
    color: C.muted,
    lineHeight: 1.5,
    flex: 1,
  },
  fallbackTitle: {
    fontSize: 10,
    fontFamily: "Helvetica-Bold",
    color: C.navy,
    marginBottom: 4,
  },

  // ── Continuation header (shown on form pages 2, 3…) ──
  contHeader: {
    backgroundColor: C.navy,
    paddingHorizontal: 36,
    paddingVertical: 10,
    flexDirection: "row",
    justifyContent: "space-between",
    alignItems: "center",
    borderBottomWidth: 2,
    borderBottomColor: C.teal,
  },
  contHeaderLeft: {
    fontSize: 9,
    fontFamily: "Helvetica-Bold",
    color: C.white,
    letterSpacing: 0.3,
  },
  contHeaderRight: {
    fontSize: 7.5,
    color: "#8eaec9",
  },

  // ── Divider ──
  divider: {
    height: 1,
    backgroundColor: C.border,
    marginVertical: 14,
  },
  tealRule: {
    height: 2,
    backgroundColor: C.teal,
    marginBottom: 14,
  },
});

// ─── Helpers ───────────────────────────────────────────────────────────────
function submissionStatusLine(submission: {
  status?: string;
  submittedAt?: string;
  form?: { validityPeriod?: number };
}): string {
  const st = submission.status || "";
  if (st === "pending") return "Pending review";
  if (st === "rejected") return "Rejected";
  if (st === "approved") {
    if (submission.submittedAt && submission.form?.validityPeriod != null) {
      const submissionDate = new Date(submission.submittedAt);
      const validityPeriod = submission.form.validityPeriod || 30;
      const expiryDate = new Date(submissionDate);
      expiryDate.setDate(expiryDate.getDate() + validityPeriod);
      if (new Date() > expiryDate) return "Expired (approved past validity)";
    }
    return "Approved";
  }
  return st || "Unknown";
}

function formatSubmittedBy(submission: Record<string, any>): string | undefined {
  const name  = submission.user?.name;
  const email = submission.user?.email;
  if (name && email) return `${name} (${email})`;
  if (name)  return String(name);
  if (email) return String(email);
  return undefined;
}

function submissionRecordProps(submission: Record<string, any>) {
  const statusLabel = submissionStatusLine(submission);
  const reviewedByStr =
    submission.reviewedBy?.name ||
    submission.reviewedBy?.email ||
    (typeof submission.reviewedBy === "string" ? submission.reviewedBy : undefined);

  return {
    submittedBy:    formatSubmittedBy(submission),
    ship:           submission.ship?.name ? String(submission.ship.name) : undefined,
    submittedAt:    submission.submittedAt ? formatDateTime(submission.submittedAt) : undefined,
    status:         submission.status,
    statusLabel,
    formTitle:      submission.form?.title ? String(submission.form.title) : undefined,
    reviewedBy:     reviewedByStr,
    reviewedAt:     submission.reviewedAt ? formatDateTime(submission.reviewedAt) : undefined,
    reviewComments: submission.reviewComments ? String(submission.reviewComments) : undefined,
  };
}

// ─── Header height constant — must match contHeader style ──────────────────
// paddingVertical:10 + borderBottom:2 + text line ~14 = ~36pt header
// + 16pt breathing room below it on every page
const FORM_HEADER_H = 36;
const FORM_BODY_TOP_PAD = FORM_HEADER_H + 16;

// ─── Continuation header ───────────────────────────────────────────────────
// Rendered as fixed inside <Page> so it appears on EVERY overflow page.
// The body wrapper gets paddingTop: FORM_HEADER_H so content never hides under it.
function ContinuationHeader({ title, vesselName }: { title?: string; vesselName?: string }) {
  return (
    <View style={ls.contHeader} fixed>
      <Text style={ls.contHeaderLeft}>{title || "Form"}</Text>
      {vesselName ? (
        <Text style={ls.contHeaderRight}>Vessel: {vesselName}</Text>
      ) : null}
    </View>
  );
}

// ─── Styled fallback page body ──────────────────────────────────────────────
function FallbackPageBody() {
  return (
    <View style={ls.fallbackBody}>
      <View style={ls.tealRule} />
      <View style={ls.fallbackCard}>
        {/* Icon circle */}
        <View style={ls.fallbackIconWrap}>
          <Text style={ls.fallbackIcon}>!</Text>
        </View>
        {/* Text block */}
        <View style={{ flex: 1 }}>
          <Text style={ls.fallbackTitle}>Form definition unavailable</Text>
          <Text style={ls.fallbackText}>
            The form definition associated with this submission could not be retrieved.
            The submission record above reflects all metadata captured at the time of filing.
            Please contact your system administrator or refer to the original form entry
            in the SENA Ship Management platform for full field details.
          </Text>
        </View>
      </View>
    </View>
  );
}

// ─── Document ──────────────────────────────────────────────────────────────
export interface SubmissionPdfDocumentProps {
  submission: Record<string, any>;
}

const SubmissionPdfDocument: React.FC<SubmissionPdfDocumentProps> = ({ submission }) => {
  const formProps = submissionToFormTemplatePdfProps(submission);
  const record   = submissionRecordProps(submission);

  // ── Fallback: no form definition ──
  if (!formProps) {
    return (
      <Document>
        <SubmissionRecordPage {...record} />
        <Page size="A4" style={pdfStyles.page}>
          <ContinuationHeader
            title={record.formTitle || "Submission Detail"}
            vesselName={record.ship}
          />
          <View style={{ paddingTop: FORM_BODY_TOP_PAD }}>
            <FallbackPageBody />
          </View>
          <PdfPageFooter />
        </Page>
      </Document>
    );
  }

  // ── Normal: submission + form content ──
  return (
    <Document>
      {/* Page 1 — Submission record (meta + review section) */}
      <SubmissionRecordPage {...record} />

      {/* Page 2+ — Form body */}
      <Page size="A4" style={pdfStyles.page}>
        {/* fixed=true → stamped at top of EVERY page this <Page> flows onto */}
        <ContinuationHeader
          title={formProps.title || record.formTitle}
          vesselName={record.ship}
        />
        {/* paddingTop offsets content below the fixed header + breathing room on all pages */}
        <View style={{ paddingTop: FORM_BODY_TOP_PAD }}>
          <FormTemplatePdfPageBody {...formProps} />
        </View>
        <PdfPageFooter />
      </Page>
    </Document>
  );
};

export default SubmissionPdfDocument;