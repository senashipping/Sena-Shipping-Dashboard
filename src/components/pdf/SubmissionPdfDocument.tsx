import React from "react";
import { Document, Page, Text, View, StyleSheet } from "@react-pdf/renderer";
import {
  FormTemplatePdfFixedHeader,
  FormTemplatePdfPageBody,
  PDF_FOOTER_RESERVE_PT,
  PDF_FORM_HEADER_RESERVE_PT,
  PdfPageFooter,
  SubmissionRecordPage,
  pdfStyles,
} from "./FormTemplatePdfDocument";
import { submissionToFormTemplatePdfProps } from "./submissionToFormTemplatePdfProps";
import { formatDateTime } from "../../lib/utils";

// ─── Colour tokens (aligned with FormTemplatePdfDocument) ───────────────────
const C = {
  navy: "#0f2341",
  teal: "#0d7fa5",
  tealLight: "#e6f4f9",
  tealBorder: "#7ac8df",
  silver: "#f4f6f8",
  border: "#cdd5dc",
  text: "#1e2a38",
  muted: "#5a6a7a",
  white: "#ffffff",
};

// ─── Local styles (fallback page only) ─────────────────────────────────────
const ls = StyleSheet.create({
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
  tealRule: {
    height: 2,
    backgroundColor: C.teal,
    marginBottom: 14,
  },
});

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
  const name = submission.user?.name;
  const email = submission.user?.email;
  if (name && email) return `${name} (${email})`;
  if (name) return String(name);
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
    submittedBy: formatSubmittedBy(submission),
    ship: submission.ship?.name ? String(submission.ship.name) : undefined,
    submittedAt: submission.submittedAt ? formatDateTime(submission.submittedAt) : undefined,
    status: submission.status,
    statusLabel,
    formTitle: submission.form?.title ? String(submission.form.title) : undefined,
    reviewedBy: reviewedByStr,
    reviewedAt: submission.reviewedAt ? formatDateTime(submission.reviewedAt) : undefined,
    reviewComments: submission.reviewComments ? String(submission.reviewComments) : undefined,
  };
}

function FallbackPageBody() {
  return (
    <View style={ls.fallbackBody}>
      <View style={ls.tealRule} />
      <View style={ls.fallbackCard}>
        <View style={ls.fallbackIconWrap}>
          <Text style={ls.fallbackIcon}>!</Text>
        </View>
        <View style={{ flex: 1 }}>
          <Text style={ls.fallbackTitle}>Form definition unavailable</Text>
          <Text style={ls.fallbackText}>
            The form definition associated with this submission could not be retrieved. The submission
            record above reflects all metadata captured at the time of filing. Please contact your system
            administrator or refer to the original form entry in the SENA Ship Management platform for full
            field details.
          </Text>
        </View>
      </View>
    </View>
  );
}

export interface SubmissionPdfDocumentProps {
  submission: Record<string, any>;
}

/**
 * PDF wiring:
 * - Page 1: `SubmissionRecordPage` (cover + review).
 * - Form section: fixed header/footer + `Page` padding so text clears both (see `PDF_*_RESERVE_PT`).
 */
const SubmissionPdfDocument: React.FC<SubmissionPdfDocumentProps> = ({ submission }) => {
  const formProps = submissionToFormTemplatePdfProps(submission);
  const record = submissionRecordProps(submission);

  const formPageStyle = [
    pdfStyles.page,
    {
      paddingTop: PDF_FORM_HEADER_RESERVE_PT,
      paddingBottom: PDF_FOOTER_RESERVE_PT,
    },
  ];

  if (!formProps) {
    return (
      <Document>
        <SubmissionRecordPage {...record} />
        <Page size="A4" style={[pdfStyles.page, { paddingBottom: PDF_FOOTER_RESERVE_PT }]}>
          <FallbackPageBody />
          <PdfPageFooter />
        </Page>
      </Document>
    );
  }

  return (
    <Document>
      <SubmissionRecordPage {...record} />
      <Page size="A4" style={formPageStyle}>
        <FormTemplatePdfFixedHeader
          title={formProps.title}
          description={formProps.description}
          documentSubtitle={formProps.documentSubtitle}
          categoryLabel={formProps.categoryLabel}
        />
        <FormTemplatePdfPageBody {...formProps} omitHeader />
        <PdfPageFooter />
      </Page>
    </Document>
  );
};

export default SubmissionPdfDocument;
