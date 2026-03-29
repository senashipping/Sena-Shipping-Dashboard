import React from "react";
import { Document, Page, Text, View } from "@react-pdf/renderer";
import {
  FormTemplatePdfPageBody,
  PdfPageFooter,
  SubmissionRecordPage,
  pdfStyles,
} from "./FormTemplatePdfDocument";
import { submissionToFormTemplatePdfProps } from "./submissionToFormTemplatePdfProps";
import { formatDateTime } from "../../lib/utils";

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

export interface SubmissionPdfDocumentProps {
  submission: Record<string, any>;
}

const submissionRecordProps = (submission: Record<string, any>) => {
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
};

const SubmissionPdfDocument: React.FC<SubmissionPdfDocumentProps> = ({ submission }) => {
  const formProps = submissionToFormTemplatePdfProps(submission);
  const record = submissionRecordProps(submission);

  if (!formProps) {
    return (
      <Document>
        <SubmissionRecordPage {...record} />
        <Page size="A4" style={pdfStyles.page}>
          <View style={{ paddingHorizontal: 36, paddingTop: 20, paddingBottom: 40 }}>
            <Text style={{ fontSize: 10, color: "#5a6a7a" }}>
              Form definition was not available for this submission.
            </Text>
          </View>
          <PdfPageFooter />
        </Page>
      </Document>
    );
  }

  return (
    <Document>
      <SubmissionRecordPage {...record} />
      <Page size="A4" style={pdfStyles.page}>
        <FormTemplatePdfPageBody {...formProps} />
        <PdfPageFooter />
      </Page>
    </Document>
  );
};

export default SubmissionPdfDocument;
