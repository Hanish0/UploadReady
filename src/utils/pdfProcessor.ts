import * as pdfjsLib from 'pdfjs-dist';
import { PDFDocument } from 'pdf-lib';

// Set up the worker source using Vite's native URL resolution
pdfjsLib.GlobalWorkerOptions.workerSrc = new URL(
  'pdfjs-dist/build/pdf.worker.min.mjs',
  import.meta.url
).toString();

export interface ExtractedFields {
  invoiceNumber: string;
  vendorGst: string;
  cgst: string;
  sgst: string;
  invoiceAmount: string;
  paidAmount: string;
  billingPeriod: string;
  mobileNumber: string;
}

/**
 * Extracts raw text content page-by-page from a PDF File in the browser.
 */
export async function extractTextFromPdf(file: File): Promise<string> {
  const arrayBuffer = await file.arrayBuffer();
  const loadingTask = pdfjsLib.getDocument({ data: new Uint8Array(arrayBuffer) });
  const pdf = await loadingTask.promise;
  let fullText = '';

  for (let i = 1; i <= pdf.numPages; i++) {
    const page = await pdf.getPage(i);
    const textContent = await page.getTextContent();
    const pageText = textContent.items.map((item: any) => item.str).join('\n');
    fullText += `\n--- Page ${i} ---\n` + pageText;
  }
  return fullText;
}

/**
 * Applies regex patterns on the bill and receipt texts to extract fields.
 */
export function extractFields(billText: string, receiptText: string): ExtractedFields {
  const fields: ExtractedFields = {
    invoiceNumber: '',
    vendorGst: '',
    cgst: '',
    sgst: '',
    invoiceAmount: '',
    paidAmount: '',
    billingPeriod: '',
    mobileNumber: '',
  };

  // 1. Mobile Number
  const mobileMatch = billText.match(/(?:Phone\s*Number|PhoneNo|Mobile)\s*:?\s*(\d{10})/i);
  if (mobileMatch) {
    fields.mobileNumber = mobileMatch[1];
  } else {
    const receiptMobileMatch = receiptText.match(/(?:POSTPAID|Customer\s+Number\s+91)\s*(\d{10})/i);
    if (receiptMobileMatch) fields.mobileNumber = receiptMobileMatch[1];
  }

  // 2. Invoice Number
  const invoiceMatch = billText.match(/\b(MF[A-Z0-9]{10,20})\b/);
  if (invoiceMatch) {
    fields.invoiceNumber = invoiceMatch[1];
  }

  // 3. Vendor GST
  const gstMatch = billText.match(/GST\s*registration\s*no\.?:\s*([0-9]{2}[A-Z]{5}[0-9]{4}[A-Z]{1}[A-Z0-9]{1}[Z]{1}[A-Z0-9]{1})/i);
  if (gstMatch) {
    fields.vendorGst = gstMatch[1];
  }

  // 4. Billing Period
  const periodMatch = billText.match(/(?:Statement|Bill)\s*Period\s*:?\s*([0-9]{1,2}\s+[a-zA-Z]{3}\s+[0-9]{4}\s*-\s*[0-9]{1,2}\s+[a-zA-Z]{3}\s+[0-9]{4})/i);
  if (periodMatch) {
    fields.billingPeriod = periodMatch[1].trim();
  } else {
    const periodAltMatch = billText.match(/(?:Statement|Bill)\s*Period\s*:?\s*([0-9]{1,2}\/[0-9]{1,2}\/[0-9]{2,4}\s*-\s*[0-9]{1,2}\/[0-9]{1,2}\/[0-9]{2,4})/i);
    if (periodAltMatch) fields.billingPeriod = periodAltMatch[1].trim();
  }

  // 5. Invoice Amount (This Month's Charges)
  const amountMatch = billText.match(/Total\s+Amount\s*(?:Payable:?)?\s*[\`\₹\s\n]*([0-9,]+\.[0-9]{2})/i);
  const chargesMatch = billText.match(/This\s+Month's\s+Charges\s*[\+\`\₹\s\n]*([0-9,]+\.[0-9]{2})/i);
  if (chargesMatch) {
    fields.invoiceAmount = chargesMatch[1];
  } else if (amountMatch) {
    fields.invoiceAmount = amountMatch[1];
  }

  // 6. Paid Amount (From receipt)
  const paidMatch = receiptText.match(/Paid\s+amount\s*[\`\₹\s\n]*([0-9,]+\.[0-9]{2})/i);
  if (paidMatch) {
    fields.paidAmount = paidMatch[1];
  } else {
    const billPayableMatch = billText.match(/Amount\s+Payable\s*[\=\`\₹\s\n]*([0-9,]+\.[0-9]{2})/i);
    if (billPayableMatch) fields.paidAmount = billPayableMatch[1];
  }

  // 7. CGST and SGST
  const cgstMatch = billText.match(/CGST[\s\S]*?9%[\s\S]*?(\d+\.\d{2})/i);
  if (cgstMatch) {
    fields.cgst = cgstMatch[1];
  }
  const sgstMatch = billText.match(/SGST\/UTGST[\s\S]*?9%[\s\S]*?(\d+\.\d{2})/i);
  if (sgstMatch) {
    fields.sgst = sgstMatch[1];
  }

  if (!fields.cgst || !fields.sgst) {
    const totalGstMatch = billText.match(/Taxes\s*\(GST\)\s*[\`\₹\s\n]*([0-9,]+\.[0-9]{2})/i);
    if (totalGstMatch) {
      const totalGst = parseFloat(totalGstMatch[1].replace(/,/g, ''));
      const half = (totalGst / 2).toFixed(2);
      fields.cgst = fields.cgst || half;
      fields.sgst = fields.sgst || half;
    }
  }

  return fields;
}

/**
 * Creates the merged PDF package ordered as Bill Page 3, Bill Page 1, Bill Page 4, and Receipt Page.
 */
export async function mergeAndCleanPdfs(
  billFile: File,
  receiptFile: File
): Promise<Blob> {
  const mergedPdf = await PDFDocument.create();

  // Load files
  const billBytes = await billFile.arrayBuffer();
  const receiptBytes = await receiptFile.arrayBuffer();

  const billPdf = await PDFDocument.load(billBytes);
  const receiptPdf = await PDFDocument.load(receiptBytes);

  const billPageCount = billPdf.getPageCount();
  const receiptPageCount = receiptPdf.getPageCount();

  // Selected pages to copy: Bill Page 3 (index 2), Page 1 (index 0), Page 4 (index 3)
  const billPagesToCopy: number[] = [];
  if (billPageCount >= 3) billPagesToCopy.push(2);
  if (billPageCount >= 1) billPagesToCopy.push(0);
  if (billPageCount >= 4) billPagesToCopy.push(3);

  // Copy and add Bill pages
  if (billPagesToCopy.length > 0) {
    const copiedBillPages = await mergedPdf.copyPages(billPdf, billPagesToCopy);
    copiedBillPages.forEach((page) => mergedPdf.addPage(page));
  }

  // Copy and add Receipt page(s) - usually page 1 (index 0)
  if (receiptPageCount >= 1) {
    const copiedReceiptPages = await mergedPdf.copyPages(receiptPdf, [0]);
    copiedReceiptPages.forEach((page) => mergedPdf.addPage(page));
  }

  const mergedPdfBytes = await mergedPdf.save();
  return new Blob([mergedPdfBytes.buffer as ArrayBuffer], { type: 'application/pdf' });
}

/**
 * Computes Month and Year for naming the output PDF.
 * Format: Airtel_Reimbursement_<MonthYear>.pdf (e.g. Airtel_Reimbursement_May2026.pdf)
 */
export function getReimbursementFilename(billingPeriod: string): string {
  // Billing Period sample: "17 Apr 2026-16 May 2026"
  // We parse the second date to extract the month and year
  const endMonthYearMatch = billingPeriod.match(/-\s*\d{1,2}\s+([a-zA-Z]{3,9})\s+(\d{4})/);
  if (endMonthYearMatch) {
    const month = endMonthYearMatch[1];
    const year = endMonthYearMatch[2];
    // Capitalize month if not already (e.g. May)
    const formattedMonth = month.charAt(0).toUpperCase() + month.slice(1).toLowerCase();
    return `Airtel_Reimbursement_${formattedMonth}${year}.pdf`;
  }

  // Fallback to current date
  const now = new Date();
  const months = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
  return `Airtel_Reimbursement_${months[now.getMonth()]}${now.getFullYear()}.pdf`;
}
