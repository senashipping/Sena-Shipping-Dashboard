import React from "react";
import { Document, Page, Text, View, StyleSheet } from "@react-pdf/renderer";
import { FormTemplatePdfPageBody } from "./FormTemplatePdfDocument";
import { submissionToFormTemplatePdfProps } from "./submissionToFormTemplatePdfProps";
import { formatDateTime } from "../../lib/utils";

const page = StyleSheet.create({
  page: {
    padding: 36,
    fontSize: 10,
    fontFamily: "Helvetica",
  },
  heading: {
    fontSize: 14,
    marginBottom: 12,
    fontFamily: "Helvetica-Bold",
  },
  row: { marginBottom: 8 },
  label: { fontSize: 9, fontFamily: "Helvetica-Bold", marginBottom: 2 },
  value: { fontSize: 9 },
});

function submissionStatusLine(submission: {
  status?: string;
  submittedAt?: string;
  form?: { validityPeriod?: number };
}): string {
  const s = submission.status || "";
  if (s === "pending") return "Pending review";
  if (s === "rejected") return "Rejected";
  if (s === "approved") {
    if (submission.submittedAt && submission.form?.validityPeriod != null) {
      const submissionDate = new Date(submission.submittedAt);
      const validityPeriod = submission.form.validityPeriod || 30;
      const expiryDate = new Date(submissionDate);
      expiryDate.setDate(expiryDate.getDate() + validityPeriod);
      if (new Date() > expiryDate) return "Expired (approved past validity)";
    }
    return "Approved";
  }
  return s || "Unknown";
}

export interface SubmissionPdfDocumentProps {
  submission: Record<string, any>;
}

const SubmissionPdfDocument: React.FC<SubmissionPdfDocumentProps> = ({ submission }) => {
  const formProps = submissionToFormTemplatePdfProps(submission);

  const meta = (
    <>
      <Text style={page.heading}>Submission record</Text>
      <View style={page.row}>
        <Text style={page.label}>Submitted by</Text>
        <Text style={page.value}>
          {submission.user?.name || "—"}
          {submission.user?.email ? ` (${submission.user.email})` : ""}
        </Text>
      </View>
      <View style={page.row}>
        <Text style={page.label}>Ship</Text>
        <Text style={page.value}>{submission.ship?.name || "—"}</Text>
      </View>
      <View style={page.row}>
        <Text style={page.label}>Submitted at</Text>
        <Text style={page.value}>
          {submission.submittedAt ? formatDateTime(submission.submittedAt) : "—"}
        </Text>
      </View>
      <View style={page.row}>
        <Text style={page.label}>Status</Text>
        <Text style={page.value}>{submissionStatusLine(submission)}</Text>
      </View>
      {(submission.reviewedAt || submission.reviewedBy || submission.reviewComments) && (
        <>
          <Text style={[page.heading, { marginTop: 12 }]}>Review</Text>
          {submission.reviewedBy && (
            <View style={page.row}>
              <Text style={page.label}>Reviewed by</Text>
              <Text style={page.value}>
                {submission.reviewedBy.name || submission.reviewedBy.email || "—"}
              </Text>
            </View>
          )}
          {submission.reviewedAt && (
            <View style={page.row}>
              <Text style={page.label}>Reviewed at</Text>
              <Text style={page.value}>{formatDateTime(submission.reviewedAt)}</Text>
            </View>
          )}
          {submission.reviewComments ? (
            <View style={page.row}>
              <Text style={page.label}>Comments</Text>
              <Text style={page.value}>{String(submission.reviewComments)}</Text>
            </View>
          ) : null}
        </>
      )}
    </>
  );

  if (!formProps) {
    return (
      <Document>
        <Page size="A4" style={page.page}>
          {meta}
          <Text style={{ marginTop: 12 }}>Form definition was not available.</Text>
        </Page>
      </Document>
    );
  }

  return (
    <Document>
      <Page size="A4" style={page.page}>
        {meta}
      </Page>
      <Page size="A4" style={page.page}>
        <FormTemplatePdfPageBody {...formProps} />
      </Page>
    </Document>
  );
};

export default SubmissionPdfDocument;
