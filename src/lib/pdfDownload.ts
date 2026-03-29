import type { ReactElement } from "react";
import { pdf } from "@react-pdf/renderer";

export async function downloadPdfDocument(doc: ReactElement, filename: string): Promise<void> {
  const blob = await pdf(doc).toBlob();
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = filename.replace(/[^\w.\-]+/g, "_");
  a.click();
  URL.revokeObjectURL(url);
}
