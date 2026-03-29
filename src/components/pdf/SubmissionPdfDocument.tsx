import React from "react";
import { Document, Page, Text, View, StyleSheet } from "@react-pdf/renderer";
import {
  FormTemplatePdfPageBody,
  PDF_CONTINUATION_HEADER_RESERVE_PT,
  PDF_FOOTER_RESERVE_PT,
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

  // ── Continuation header (fixed, every wrapped page) ──
  // `Page` `paddingTop` insets the content box; without `absolute`, fixed nodes are
  // laid out *below* that inset — a blank band above the bar. Pin to the real page top.
  contHeaderFixedWrap: {
    position: "absolute",
    top: 0,
    left: 0,
    right: 0,
    width: "100%",
  },
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
  /** Space after the navy bar (repeats with the fixed continuation strip). */
  contHeaderAfterGap: {
    height: 20,
    backgroundColor: C.white,
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

// ─── Continuation header ───────────────────────────────────────────────────
/** Thin navy strip: every slice of this wrapping page except the first (first form page keeps only the large header). */
function ContinuationHeader({
  title,
  vesselName,
  showOnFirstSlice = true,
}: {
  title?: string;
  vesselName?: string;
  /** When false, only render from the 2nd wrapped slice onward (`subPageNumber` > 1). */
  showOnFirstSlice?: boolean;
}) {
  if (showOnFirstSlice) {
    return (
      <View style={ls.contHeaderFixedWrap} fixed>
        <View style={ls.contHeader}>
          <Text style={ls.contHeaderLeft}>{title || "Form"}</Text>
          {vesselName ? (
            <Text style={ls.contHeaderRight}>Vessel: {vesselName}</Text>
          ) : null}
        </View>
        <View style={ls.contHeaderAfterGap} />
      </View>
    );
  }

  return (
    <View
      style={ls.contHeaderFixedWrap}
      fixed
      render={({ subPageNumber }) =>
        subPageNumber > 1 ? (
          <View>
            <View style={ls.contHeader}>
              <Text style={ls.contHeaderLeft}>{title || "Form"}</Text>
              {vesselName ? (
                <Text style={ls.contHeaderRight}>Vessel: {vesselName}</Text>
              ) : null}
            </View>
            <View style={ls.contHeaderAfterGap} />
          </View>
        ) : null
      }
    />
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
        <Page
          size="A4"
          style={[
            pdfStyles.page,
            {
              paddingTop: PDF_CONTINUATION_HEADER_RESERVE_PT,
              paddingBottom: PDF_FOOTER_RESERVE_PT,
            },
          ]}
        >
          <ContinuationHeader
            title={record.formTitle || "Submission Detail"}
            vesselName={record.ship}
          />
          <FallbackPageBody />
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

      {/* Page 2+ — Form body (auto-wraps across as many pages as needed) */}
      <Page
        size="A4"
        style={[
          pdfStyles.page,
          {
            paddingTop: PDF_CONTINUATION_HEADER_RESERVE_PT,
            paddingBottom: PDF_FOOTER_RESERVE_PT,
          },
        ]}
      >
        <ContinuationHeader
          title={formProps.title || record.formTitle}
          vesselName={record.ship}
          showOnFirstSlice={false}
        />
        <FormTemplatePdfPageBody
          {...formProps}
          pullUpTopReservePt={PDF_CONTINUATION_HEADER_RESERVE_PT}
        />
        <PdfPageFooter />
      </Page>
    </Document>
  );
};

export default SubmissionPdfDocument;