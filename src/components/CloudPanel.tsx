import { useState } from 'react';
import { 
  Copy, 
  Check, 
  Mail, 
  Calendar, 
  Hash, 
  ExternalLink, 
  Zap,
  Download,
  Clock,
  ChevronDown,
  ChevronUp,
  Play
} from 'lucide-react';

interface CloudPanelProps {
  defaultWorkEmail?: string;
}

export function CloudPanel({ defaultWorkEmail = '' }: CloudPanelProps) {
  const [workEmail, setWorkEmail] = useState(defaultWorkEmail);
  const [billGenDay, setBillGenDay] = useState<number>(17);
  const [scanWindow, setScanWindow] = useState<number>(10);
  const [scheduleDay, setScheduleDay] = useState<string>('25');
  const [mobileNumber, setMobileNumber] = useState('');
  const [copied, setCopied] = useState(false);
  const [showSourceCode, setShowSourceCode] = useState(false);
  const [scriptCopied, setScriptCopied] = useState(false);

  // Email format validation helper
  const isEmailValid = (email: string) => {
    return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email);
  };

  // Optional 10-digit mobile check
  const isMobileValid = (mobile: string) => {
    if (!mobile) return true;
    return /^\d{10}$/.test(mobile);
  };

  const isConfigValid = isEmailValid(workEmail) && isMobileValid(mobileNumber);

  // Dynamic configuration states
  const configState = !isConfigValid ? 'invalid' : scriptCopied ? 'copied' : 'ready';

  // Generate the Google Apps Script text dynamically
  const generateScript = () => {
    const formattedScheduleDay = isNaN(Number(scheduleDay)) ? `'${scheduleDay}'` : scheduleDay;
    
    return `/**
 * Airtel Reimbursement Cloud Auto-Pilot Daemon
 * -------------------------------------------------------------
 * This script runs in Google Apps Script (script.google.com).
 * It runs on a daily time-based trigger, scans Gmail for Airtel
 * postpaid bills and payment receipts within your billing window,
 * extracts fields using Google Docs OCR, merges pages, and emails
 * the reimbursement package to your work Outlook inbox.
 */

// ==========================================
// USER SETTINGS (Configured via UploadReady UI)
// ==========================================
var WORK_EMAIL = "${workEmail || 'yourname@work.com'}";
var BILL_GENERATION_DAY = ${billGenDay};
var SCAN_WINDOW_DAYS = ${scanWindow};
var SCHEDULE_DAY = ${formattedScheduleDay};
var AIRTEL_MOBILE = "${mobileNumber}";
var GMAIL_LABEL_NAME = "Reimbursement-Processed";
var DRIVE_QUEUE_FOLDER = "UploadReady_Queue";

// ==========================================
// MAIN AUTOMATION ENTRY POINT
// ==========================================
async function runAirtelAutomation() {
  Logger.log("Airtel Automation run started: " + new Date().toString());
  
  // 1. Process anything in the queue first if schedule conditions are met
  processQueue();
  
  // 2. Check if we are inside the active scanning billing window
  var today = new Date();
  var currentDay = today.getDate();
  var windowStart = BILL_GENERATION_DAY;
  var windowEnd = BILL_GENERATION_DAY + SCAN_WINDOW_DAYS;
  
  Logger.log("Current day of month: " + currentDay + " (Scan window: " + windowStart + " to " + windowEnd + ")");
  
  if (currentDay < windowStart || currentDay > windowEnd) {
    Logger.log("Current date is outside the billing window. Skipping scan to conserve Gmail queries.");
    return;
  }
  
  // 3. Check if we have already processed this month's reimbursement
  var currentMonthYearStr = Utilities.formatDate(today, Session.getScriptTimeZone(), "MMM-yyyy");
  if (isMonthAlreadyProcessed(currentMonthYearStr)) {
    Logger.log("Airtel reimbursement for " + currentMonthYearStr + " has already been processed. Exiting.");
    return;
  }
  
  // 4. Scan Gmail for the Statement and the Receipt
  Logger.log("Active scanning window open. Searching Gmail...");
  var billSearchQuery = 'subject:(Airtel Mobile Bill) has:attachment filename:pdf -label:' + GMAIL_LABEL_NAME;
  var receiptSearchQuery = 'subject:(Airtel (Payment OR Receipt OR Transaction)) has:attachment filename:pdf -label:' + GMAIL_LABEL_NAME;
  
  if (AIRTEL_MOBILE) {
    billSearchQuery += ' AND "' + AIRTEL_MOBILE + '"';
    receiptSearchQuery += ' AND "' + AIRTEL_MOBILE + '"';
  }
  
  // Search within the last 30 days
  var dateSearchQuery = ' after:' + Utilities.formatDate(new Date(today.getTime() - 30 * 24 * 60 * 60 * 1000), Session.getScriptTimeZone(), "yyyy/MM/dd");
  billSearchQuery += dateSearchQuery;
  receiptSearchQuery += dateSearchQuery;
  
  var billThreads = GmailApp.search(billSearchQuery, 0, 1);
  var receiptThreads = GmailApp.search(receiptSearchQuery, 0, 1);
  
  if (billThreads.length === 0 || receiptThreads.length === 0) {
    Logger.log("Waiting for files. Found Bills: " + billThreads.length + ", Found Receipts: " + receiptThreads.length);
    return;
  }
  
  Logger.log("Found both Bill and Receipt emails! Initiating download...");
  
  var billThread = billThreads[0];
  var receiptThread = receiptThreads[0];
  
  var billBlob = getPdfAttachment(billThread);
  var receiptBlob = getPdfAttachment(receiptThread);
  
  if (!billBlob || !receiptBlob) {
    Logger.log("Failed to find PDF attachments in the found threads. Exiting.");
    return;
  }
  
  // 5. OCR text extraction from the statement
  Logger.log("Running Google Drive OCR to extract fields...");
  var billText = extractTextFromPdf(billBlob);
  var receiptText = extractTextFromPdf(receiptBlob);
  
  // Parse fields
  var extractedFields = parseAirtelFields(billText, receiptText);
  Logger.log("Extracted details: " + JSON.stringify(extractedFields));
  
  // 6. Load pdf-lib and merge the documents
  Logger.log("Loading pdf-lib to merge document pages...");
  var mergedBlob = await mergePdfs(billBlob, receiptBlob, extractedFields.billingPeriod);
  
  // 7. Determine transmission and save/queue
  var isImmediate = typeof SCHEDULE_DAY === 'string' && SCHEDULE_DAY.toLowerCase() === 'immediate';
  
  if (isImmediate || currentDay >= SCHEDULE_DAY) {
    Logger.log("Target transmission date met. Emailing to work email directly...");
    sendReimbursementEmail(WORK_EMAIL, extractedFields, mergedBlob);
  } else {
    Logger.log("Saving processed package to Google Drive queue (waiting until day " + SCHEDULE_DAY + ")...");
    saveToQueue(mergedBlob, extractedFields);
  }
  
  // 8. Tag email threads so they are marked as processed
  var processedLabel = getOrCreateGmailLabel(GMAIL_LABEL_NAME);
  billThread.addLabel(processedLabel);
  receiptThread.addLabel(processedLabel);
  
  Logger.log("Workflow completed successfully for cycle: " + extractedFields.billingPeriod);
}

function getPdfAttachment(thread) {
  var messages = thread.getMessages();
  for (var i = messages.length - 1; i >= 0; i--) {
    var attachments = messages[i].getAttachments();
    for (var j = 0; j < attachments.length; j++) {
      var contentType = attachments[j].getContentType();
      var name = attachments[j].getName().toLowerCase();
      if (contentType === 'application/pdf' || name.indexOf('.pdf') !== -1) {
        return attachments[j];
      }
    }
  }
  return null;
}

function extractTextFromPdf(pdfBlob) {
  var metadata = {
    name: 'UploadReady_Temp_OCR_' + Utilities.getUuid(),
    mimeType: 'application/vnd.google-apps.document'
  };
  
  var boundary = 'multipart_boundary_uploadready';
  var requestBody = 
    '--' + boundary + '\\r\\n' +
    'Content-Type: application/json; charset=UTF-8\\r\\n\\r\\n' +
    JSON.stringify(metadata) + '\\r\\n' +
    '--' + boundary + '\\r\\n' +
    'Content-Type: ' + pdfBlob.getContentType() + '\\r\\n\\r\\n';
  
  var requestBytes = Utilities.newBlob(requestBody).getBytes();
  var pdfBytes = pdfBlob.getBytes();
  var footerBytes = Utilities.newBlob('\\r\\n--' + boundary + '--\\r\\n').getBytes();
  
  var payload = requestBytes.concat(pdfBytes).concat(footerBytes);
  
  var options = {
    method: 'post',
    contentType: 'multipart/related; boundary=' + boundary,
    headers: {
      Authorization: 'Bearer ' + ScriptApp.getOAuthToken()
    },
    payload: payload,
    muteHttpExceptions: true
  };
  
  var response = UrlFetchApp.fetch('https://www.googleapis.com/upload/drive/v3/files?uploadType=multipart', options);
  var json = JSON.parse(response.getContentText());
  
  if (!json.id) {
    throw new Error('Google Drive OCR Upload failed: ' + response.getContentText());
  }
  
  var fileId = json.id;
  var extractedText = '';
  
  try {
    var doc = DocumentApp.openById(fileId);
    extractedText = doc.getBody().getText();
  } catch (err) {
    Logger.log("Error reading OCR document: " + err.toString());
  } finally {
    try {
      DriveApp.getFileById(fileId).setTrashed(true);
    } catch (e) {
      Logger.log("Failed to clean up temporary OCR file: " + e.toString());
    }
  }
  
  return extractedText;
}

function parseAirtelFields(billText, receiptText) {
  var fields = {
    invoiceNumber: "",
    vendorGst: "",
    cgst: "",
    sgst: "",
    invoiceAmount: "",
    paidAmount: "",
    billingPeriod: "",
    mobileNumber: ""
  };
  
  var mobileMatch = billText.match(/(?:Phone\\s*Number|PhoneNo|Mobile)\\s*:?\\s*(\\d{10})/i);
  if (mobileMatch) {
    fields.mobileNumber = mobileMatch[1];
  } else {
    var receiptMobileMatch = receiptText.match(/(?:POSTPAID|Customer\\s+Number\\s+91)\\s*(\\d{10})/i);
    if (receiptMobileMatch) fields.mobileNumber = receiptMobileMatch[1];
  }
  
  var invoiceMatch = billText.match(/\\b(MF[A-Z0-9]{10,20})\\b/);
  if (invoiceMatch) {
    fields.invoiceNumber = invoiceMatch[1];
  }
  
  var gstMatch = billText.match(/GST\\s*registration\\s*no\\.?:\\s*([0-9]{2}[A-Z]{5}[0-9]{4}[A-Z]{1}[A-Z0-9]{1}[Z]{1}[A-Z0-9]{1})/i);
  if (gstMatch) {
    fields.vendorGst = gstMatch[1];
  }
  
  var periodMatch = billText.match(/(?:Statement|Bill)\\s*Period\\s*:?\\s*([0-9]{1,2}\\s+[a-zA-Z]{3}\\s+[0-9]{4}\\s*-\\s*[0-9]{1,2}\\s+[a-zA-Z]{3}\\s+[0-9]{4})/i);
  if (periodMatch) {
    fields.billingPeriod = periodMatch[1].trim();
  } else {
    var periodAltMatch = billText.match(/(?:Statement|Bill)\\s*Period\\s*:?\\s*([0-9]{1,2}\\/[0-9]{1,2}\\/[0-9]{2,4}\\s*-\\s*[0-9]{1,2}\\/[0-9]{1,2}\\/[0-9]{2,4})/i);
    if (periodAltMatch) fields.billingPeriod = periodAltMatch[1].trim();
  }
  
  var amountMatch = billText.match(/Total\\s+Amount\\s*(?:Payable:?)?\\s*[\\\`\\₹\\s\\n]*([0-9,]+\\.[0-9]{2})/i);
  var chargesMatch = billText.match(/This\\s+Month's\\s+Charges\\s*[\\+\\\`\\₹\\s\\n]*([0-9,]+\\.[0-9]{2})/i);
  if (chargesMatch) {
    fields.invoiceAmount = chargesMatch[1];
  } else if (amountMatch) {
    fields.invoiceAmount = amountMatch[1];
  }
  
  var paidMatch = receiptText.match(/Paid\\s+amount\\s*[\\\`\\₹\\s\\n]*([0-9,]+\\.[0-9]{2})/i);
  if (paidMatch) {
    fields.paidAmount = paidMatch[1];
  } else {
    var billPayableMatch = billText.match(/Amount\\s+Payable\\s*[\\=\\\`\\₹\\s\\n]*([0-9,]+\\.[0-9]{2})/i);
    if (billPayableMatch) fields.paidAmount = billPayableMatch[1];
  }
  
  var cgstMatch = billText.match(/CGST[\\s\\S]*?9%[\\s\\S]*?(\\d+\\.\\d{2})/i);
  if (cgstMatch) {
    fields.cgst = cgstMatch[1];
  }
  var sgstMatch = billText.match(/SGST\\/UTGST[\\s\\S]*?9%[\\s\\S]*?(\\d+\\.\\d{2})/i);
  if (sgstMatch) {
    fields.sgst = sgstMatch[1];
  }
  
  if (!fields.cgst || !fields.sgst) {
    var totalGstMatch = billText.match(/Taxes\\s*\\(GST\\)\\s*[\\\`\\₹\\s\\n]*([0-9,]+\\.[0-9]{2})/i);
    if (totalGstMatch) {
      var totalGst = parseFloat(totalGstMatch[1].replace(/,/g, ''));
      var half = (totalGst / 2).toFixed(2);
      fields.cgst = fields.cgst || half;
      fields.sgst = fields.sgst || half;
    }
  }
  
  return fields;
}

async function mergePdfs(billBlob, receiptBlob, billingPeriod) {
  // Polyfill global timer functions for libraries (like pdf-lib) that expect them in GAS
  if (typeof setTimeout === 'undefined') {
    globalThis.setTimeout = function(cb, ms) {
      Utilities.sleep(ms || 0);
      try { cb(); } catch (e) {}
      return 0;
    };
    globalThis.clearTimeout = function() {};
  }

  var pdfLibUrl = "https://unpkg.com/pdf-lib@1.17.1/dist/pdf-lib.min.js";
  var pdfLibCode = UrlFetchApp.fetch(pdfLibUrl).getContentText();
  eval(pdfLibCode);
  
  var PDFDocument = PDFLib.PDFDocument;
  var billBytes = new Uint8Array(billBlob.getBytes());
  var receiptBytes = new Uint8Array(receiptBlob.getBytes());
  
  var mergedPdf = await PDFDocument.create();
  var billDoc = await PDFDocument.load(billBytes);
  var receiptDoc = await PDFDocument.load(receiptBytes);
  
  var billPageCount = billDoc.getPageCount();
  var receiptPageCount = receiptDoc.getPageCount();
  
  var billPagesToCopy = [];
  if (billPageCount >= 3) billPagesToCopy.push(2);
  if (billPageCount >= 1) billPagesToCopy.push(0);
  if (billPageCount >= 4) billPagesToCopy.push(3);
  
  if (billPagesToCopy.length > 0) {
    var copiedBillPages = await mergedPdf.copyPages(billDoc, billPagesToCopy);
    copiedBillPages.forEach(function(page) { mergedPdf.addPage(page); });
  }
  
  if (receiptPageCount >= 1) {
    var copiedReceiptPages = await mergedPdf.copyPages(receiptDoc, [0]);
    copiedReceiptPages.forEach(function(page) { mergedPdf.addPage(page); });
  }
  
  var mergedBytes = await mergedPdf.save();
  
  var filename = "Airtel_Reimbursement.pdf";
  var endMonthYearMatch = billingPeriod.match(/-\\s*\\d{1,2}\\s+([a-zA-Z]{3,9})\\s+(\\d{2,4})/);
  if (endMonthYearMatch) {
    var month = endMonthYearMatch[1];
    var formattedMonth = month.charAt(0).toUpperCase() + month.slice(1).toLowerCase();
    filename = "Airtel_Reimbursement_" + formattedMonth + endMonthYearMatch[2] + ".pdf";
  } else {
    var now = new Date();
    var months = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
    filename = "Airtel_Reimbursement_" + months[now.getMonth()] + now.getFullYear() + ".pdf";
  }
  
  return Utilities.newBlob(mergedBytes, "application/pdf", filename);
}

function sendReimbursementEmail(recipientEmail, fields, pdfBlob) {
  var subject = "Airtel Reimbursement Package - " + pdfBlob.getName().replace("Airtel_Reimbursement_", "").replace(".pdf", "");
  
  var htmlBody = 
    '<div style="font-family: \\'Segoe UI\\', Tahoma, Geneva, Verdana, sans-serif; max-width: 600px; margin: 0 auto; border: 1px solid #e2e8f0; border-radius: 16px; overflow: hidden; box-shadow: 0 10px 25px -5px rgba(0, 0, 0, 0.05);">' +
    '  <div style="background: linear-gradient(135deg, #6366f1 0%, #4f46e5 100%); padding: 32px 24px; text-align: center; color: white;">' +
    '    <h2 style="margin: 0; font-size: 24px; font-weight: 700; letter-spacing: -0.5px;">Airtel Telecom Reimbursement</h2>' +
    '    <p style="margin: 8px 0 0 0; opacity: 0.9; font-size: 14px;">Your compiled monthly reimbursement package is ready</p>' +
    '  </div>' +
    '  <div style="padding: 32px 24px; background-color: #ffffff; color: #1e293b;">' +
    '    <p style="margin: 0 0 24px 0; font-size: 15px; line-height: 1.6;">Hello,</p>' +
    '    <p style="margin: 0 0 24px 0; font-size: 15px; line-height: 1.6;">Here are the parsed details and the cleaned, merged PDF invoice for your monthly Airtel telecom reimbursement. You can copy the values directly below into your corporate portal:</p>' +
    '    ' +
    '    <table style="width: 100%; border-collapse: collapse; margin-bottom: 28px;">' +
    '      <tbody>' +
    '        <tr style="border-bottom: 1px solid #f1f5f9;">' +
    '          <td style="padding: 12px 8px; font-weight: 600; color: #64748b; width: 40%; font-size: 14px;">Mobile Number</td>' +
    '          <td style="padding: 12px 8px; font-weight: 700; color: #0f172a; font-size: 14px; font-family: monospace;">' + fields.mobileNumber + '</td>' +
    '        </tr>' +
    '        <tr style="border-bottom: 1px solid #f1f5f9;">' +
    '          <td style="padding: 12px 8px; font-weight: 600; color: #64748b; font-size: 14px;">Invoice Number</td>' +
    '          <td style="padding: 12px 8px; font-weight: 700; color: #0f172a; font-size: 14px; font-family: monospace;">' + fields.invoiceNumber + '</td>' +
    '        </tr>' +
    '        <tr style="border-bottom: 1px solid #f1f5f9;">' +
    '          <td style="padding: 12px 8px; font-weight: 600; color: #64748b; font-size: 14px;">Vendor GSTIN</td>' +
    '          <td style="padding: 12px 8px; font-weight: 700; color: #0f172a; font-size: 14px; font-family: monospace;">' + fields.vendorGst + '</td>' +
    '        </tr>' +
    '        <tr style="border-bottom: 1px solid #f1f5f9;">' +
    '          <td style="padding: 12px 8px; font-weight: 600; color: #64748b; font-size: 14px;">Billing Period</td>' +
    '          <td style="padding: 12px 8px; font-weight: 700; color: #0f172a; font-size: 14px;">' + fields.billingPeriod + '</td>' +
    '        </tr>' +
    '        <tr style="border-bottom: 1px solid #f1f5f9;">' +
    '          <td style="padding: 12px 8px; font-weight: 600; color: #64748b; font-size: 14px;">Invoice Amount</td>' +
    '          <td style="padding: 12px 8px; font-weight: 700; color: #4f46e5; font-size: 15px;">₹ ' + fields.invoiceAmount + '</td>' +
    '        </tr>' +
    '        <tr style="border-bottom: 1px solid #f1f5f9;">' +
    '          <td style="padding: 12px 8px; font-weight: 600; color: #64748b; font-size: 14px;">Paid Amount</td>' +
    '          <td style="padding: 12px 8px; font-weight: 700; color: #16a34a; font-size: 15px;">₹ ' + fields.paidAmount + '</td>' +
    '        </tr>' +
    '        <tr style="border-bottom: 1px solid #f1f5f9;">' +
    '          <td style="padding: 12px 8px; font-weight: 600; color: #64748b; font-size: 14px;">CGST (9%)</td>' +
    '          <td style="padding: 12px 8px; font-weight: 700; color: #0f172a; font-size: 14px;">₹ ' + fields.cgst + '</td>' +
    '        </tr>' +
    '        <tr>' +
    '          <td style="padding: 12px 8px; font-weight: 600; color: #64748b; font-size: 14px;">SGST (9%)</td>' +
    '          <td style="padding: 12px 8px; font-weight: 700; color: #0f172a; font-size: 14px;">₹ ' + fields.sgst + '</td>' +
    '        </tr>' +
    '      </tbody>' +
    '    </table>' +
    '    ' +
    '    <div style="background-color: #f8fafc; border-left: 4px solid #6366f1; padding: 16px; border-radius: 0 8px 8px 0; margin-bottom: 24px;">' +
    '      <span style="font-size: 13px; font-weight: 600; color: #475569; display: block; margin-bottom: 4px;">Attachment</span>' +
    '      <span style="font-size: 14px; font-weight: 700; color: #0f172a; word-break: break-all;">' + pdfBlob.getName() + '</span>' +
    '    </div>' +
    '    ' +
    '    <p style="margin: 0; font-size: 13px; color: #94a3b8; text-align: center; border-top: 1px solid #f1f5f9; padding-top: 24px;">' +
    '      Generated automatically by UploadReady Cloud Auto-Pilot.' +
    '    </p>' +
    '  </div>' +
    '</div>';

  MailApp.sendEmail({
    to: recipientEmail,
    subject: subject,
    htmlBody: htmlBody,
    attachments: [pdfBlob]
  });
  
  Logger.log("Email successfully sent to " + recipientEmail);
}

function saveToQueue(pdfBlob, fields) {
  var parentFolder = getOrCreateFolder(DRIVE_QUEUE_FOLDER);
  var file = parentFolder.createFile(pdfBlob);
  file.setDescription(JSON.stringify(fields));
  Logger.log("Saved file ID " + file.getId() + " in folder " + DRIVE_QUEUE_FOLDER);
}

function processQueue() {
  var today = new Date();
  var currentDay = today.getDate();
  var isImmediate = typeof SCHEDULE_DAY === 'string' && SCHEDULE_DAY.toLowerCase() === 'immediate';
  if (!isImmediate && currentDay < SCHEDULE_DAY) {
    Logger.log("Checked queue: Today (day " + currentDay + ") is before scheduled transmission day (" + SCHEDULE_DAY + "). Keeping files in queue.");
    return;
  }
  
  var parentFolder = getOrCreateFolder(DRIVE_QUEUE_FOLDER);
  var files = parentFolder.getFiles();
  var dispatchCount = 0;
  
  while (files.hasNext()) {
    var file = files.next();
    try {
      var pdfBlob = file.getBlob();
      var description = file.getDescription();
      
      if (!description) {
        Logger.log("Warning: File found in queue without fields metadata description. Skipping.");
        continue;
      }
      
      var fields = JSON.parse(description);
      Logger.log("Dispatching queued item from: " + fields.billingPeriod);
      
      sendReimbursementEmail(WORK_EMAIL, fields, pdfBlob);
      file.setTrashed(true);
      dispatchCount++;
    } catch (e) {
      Logger.log("Error processing queued file ID " + file.getId() + ": " + e.toString());
    }
  }
  
  if (dispatchCount > 0) {
    Logger.log("Queue checked. Dispatched " + dispatchCount + " pending packages.");
  } else {
    Logger.log("Checked queue. No pending packages to dispatch.");
  }
}

function isMonthAlreadyProcessed(monthYearStr) {
  var label = getOrCreateGmailLabel(GMAIL_LABEL_NAME);
  var query = 'label:' + GMAIL_LABEL_NAME + ' after:' + Utilities.formatDate(new Date(new Date().getTime() - 45 * 24 * 60 * 60 * 1000), Session.getScriptTimeZone(), "yyyy/MM/dd");
  var threads = GmailApp.search(query, 0, 10);
  
  for (var i = 0; i < threads.length; i++) {
    var messages = threads[i].getMessages();
    for (var j = 0; j < messages.length; j++) {
      var date = messages[j].getDate();
      var msgMonthYear = Utilities.formatDate(date, Session.getScriptTimeZone(), "MMM-yyyy");
      if (msgMonthYear === monthYearStr) {
        return true;
      }
    }
  }
  return false;
}

function getOrCreateGmailLabel(name) {
  var label = GmailApp.getUserLabelByName(name);
  if (!label) {
    label = GmailApp.createLabel(name);
  }
  return label;
}

function getOrCreateFolder(name) {
  var folders = DriveApp.getFoldersByName(name);
  if (folders.hasNext()) {
    return folders.next();
  }
  return DriveApp.createFolder(name);
}
`;
  };

  const handleCopyCode = async () => {
    if (!isConfigValid) return;
    try {
      await navigator.clipboard.writeText(generateScript());
      setCopied(true);
      setScriptCopied(true);
      setTimeout(() => setCopied(false), 2500);
    } catch (err) {
      console.error('Failed to copy code: ', err);
    }
  };

  const handleDownloadCode = () => {
    if (!isConfigValid) return;
    try {
      const element = document.createElement("a");
      const file = new Blob([generateScript()], { type: 'text/javascript' });
      element.href = URL.createObjectURL(file);
      element.download = "AirtelReimbursement.js";
      document.body.appendChild(element);
      element.click();
      document.body.removeChild(element);
      setScriptCopied(true);
    } catch (err) {
      console.error('Failed to download code: ', err);
    }
  };

  const highlightCode = (code: string) => {
    // Escape HTML first to prevent code injection issues
    const escaped = code
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;');
    
    // Single regex scan ensures that string literals and comments are not double-highlighted
    const tokenRegex = /(\/\*\*[\s\S]*?\*\/|\/\/.*)|('(?:\\.|[^'])*'|"(?:\\.|[^"])*")|\b(var|let|const|function|return|if|else|for|new|typeof)\b|\b(Logger|GmailApp|DriveApp|DocumentApp|Session|Utilities|UrlFetchApp|MailApp)\b|\b(\d+)\b/g;
    
    const highlighted = escaped.replace(tokenRegex, (match, comment, string, keyword, api, number) => {
      if (comment) return `<span class="code-comment">${comment}</span>`;
      if (string) return `<span class="code-string">${string}</span>`;
      if (keyword) return `<span class="code-keyword">${keyword}</span>`;
      if (api) return `<span class="code-api">${api}</span>`;
      if (number) return `<span class="code-number">${number}</span>`;
      return match;
    });

    return <code dangerouslySetInnerHTML={{ __html: highlighted }} />;
  };

  return (
    <div className="cloud-panel-container animate-fade-in">
      {/* Premium Informative Header Card */}
      <div className="glass-panel cloud-header-card">
        <div className="card-accent-gradient" />
        <div className="cloud-header-layout">
          <div className="cloud-icon-bg animate-pulse-glow">
            <Zap className="accent-icon" />
          </div>
          <div>
            <h3>Cloud Auto-Pilot Setup</h3>
            <p>
              Deploy this lightweight background script to your <strong>Google Apps Script</strong> environment. It operates securely in your Google Account to scan for Airtel statements in Gmail, extract invoice metrics, compile the reimbursement PDF, and deliver it directly to your corporate inbox on your monthly schedule.
            </p>
          </div>
        </div>
      </div>

      {/* 2-Column Grid: Settings on Left and Code Editor Panel on Right */}
      <div className="cloud-config-grid">
        {/* Left Column: Form Card */}
        <div className="cloud-left-col">
          <div className="glass-panel cloud-form-card">
            <div>
              <h2 className="panel-title">Configure Settings</h2>
              <p className="panel-subtitle">
                Tailor the automation script to your personal Gmail and billing cycle.
              </p>
            </div>

            <div className="form-fields-container">
              {/* Input 1: Work Outlook Email */}
              <div className="form-group">
                <label htmlFor="work-email">
                  <Mail size={14} /> Work Outlook Email
                </label>
                <input
                  id="work-email"
                  type="email"
                  placeholder="yourname@company.com"
                  value={workEmail}
                  onChange={(e) => setWorkEmail(e.target.value)}
                  className="clay-input-box"
                />
                <span className="input-helper">Corporate email where reimbursement requests will be sent.</span>
              </div>
              
              {/* Row with Bill Gen Day & Scan Window */}
              <div className="form-row-2">
                {/* Input 2: Bill Generation Day */}
                <div className="form-group">
                  <label htmlFor="bill-gen-day">
                    <Calendar size={14} /> Bill Generated Day
                  </label>
                  <input
                    id="bill-gen-day"
                    type="number"
                    min={1}
                    max={31}
                    value={billGenDay}
                    onChange={(e) => setBillGenDay(Math.min(31, Math.max(1, parseInt(e.target.value) || 1)))}
                    className="clay-input-box"
                  />
                  <span className="input-helper">Day Airtel issues the bill.</span>
                </div>

                {/* Input 3: Scan Window Duration */}
                <div className="form-group">
                  <label htmlFor="scan-window">
                    <Zap size={14} /> Scan Window (Days)
                  </label>
                  <input
                    id="scan-window"
                    type="number"
                    min={1}
                    max={25}
                    value={scanWindow}
                    onChange={(e) => setScanWindow(Math.min(25, Math.max(1, parseInt(e.target.value) || 1)))}
                    className="clay-input-box"
                  />
                  <span className="input-helper">Active checking window.</span>
                </div>
              </div>

              {/* Input 4: Scheduled Delivery Day */}
              <div className="form-group">
                <label htmlFor="schedule-day">
                  <Clock size={14} /> Send to Work Email Day
                </label>
                <div className="select-container">
                  <select
                    id="schedule-day"
                    value={scheduleDay}
                    onChange={(e) => setScheduleDay(e.target.value)}
                    className="clay-input-box clay-select"
                  >
                    <option value="immediate">Immediate (As soon as paid)</option>
                    {Array.from({ length: 28 }, (_, i) => String(i + 1)).map((d) => (
                      <option key={d} value={d}>
                        Day {d} of month
                      </option>
                    ))}
                  </select>
                </div>
                <span className="input-helper">Choose when to dispatch the compiled package to your inbox.</span>
              </div>

              {/* Input 5: Mobile Number Filter */}
              <div className="form-group">
                <label htmlFor="mobile-number">
                  <Hash size={14} /> Filter by Mobile (Optional)
                </label>
                <input
                  id="mobile-number"
                  type="text"
                  placeholder="10-digit number"
                  maxLength={10}
                  value={mobileNumber}
                  onChange={(e) => setMobileNumber(e.target.value.replace(/\D/g, ''))}
                  className="clay-input-box"
                />
                <span className="input-helper">Filters statement emails to only process this mobile number.</span>
              </div>
            </div>
          </div>
        </div>

        {/* Right Side: Generated Code Editor Card with Dynamic States */}
        <div className="glass-panel cloud-code-card">
          <div className="preview-header">
            <div>
              <div className="code-status-header">
                <h2 className="panel-title">Google Apps Script</h2>
                {configState === 'invalid' && (
                  <span className="status-badge badge-invalid">● Incomplete Configuration</span>
                )}
                {configState === 'ready' && (
                  <span className="status-badge badge-ready">● Ready</span>
                )}
                {configState === 'copied' && (
                  <span className="status-badge badge-copied">● Script Generated & Copied</span>
                )}
              </div>
              <p className="panel-subtitle">
                {configState === 'invalid' 
                  ? 'Please configure a valid work email above to enable copy and download.' 
                  : 'Deploy this automation script inside your Google Account.'}
              </p>
            </div>
            <div className="code-actions-group">
              <button
                onClick={handleDownloadCode}
                className="download-code-btn clay-button"
                aria-label="Download script file"
                disabled={!isConfigValid}
              >
                <Download size={14} /> Download Script
              </button>
              <button
                onClick={handleCopyCode}
                className={`copy-code-btn ${copied ? 'copied' : ''} clay-button`}
                aria-label="Copy code to clipboard"
                disabled={!isConfigValid}
              >
                {copied ? (
                  <>
                    <Check size={14} className="animate-scale-up text-success" /> Copied!
                  </>
                ) : (
                  <>
                    <Copy size={14} /> Copy Script
                  </>
                )}
              </button>
            </div>
          </div>

          <div className="code-accordion-section">
            <button
              onClick={() => isConfigValid && setShowSourceCode(!showSourceCode)}
              className={`source-toggle-btn ${!isConfigValid ? 'disabled' : ''}`}
              disabled={!isConfigValid}
            >
              <span>{showSourceCode ? 'Hide Javascript Source' : 'Show Javascript Source'}</span>
              {showSourceCode ? <ChevronUp size={16} /> : <ChevronDown size={16} />}
            </button>

            {showSourceCode && isConfigValid && (
              <div className="code-pre-wrapper animate-scale-up">
                <pre className="code-pre">
                  {highlightCode(generateScript())}
                </pre>
              </div>
            )}
          </div>
        </div>
      </div>

      {/* Bottom Section: Actionable Deployment Guide revealed when config is valid */}
      {isConfigValid && (
        <div className="glass-panel setup-guide-card animate-fade-in">
          <div className="card-accent-gradient" />
          
          {scriptCopied && (
            <div className="copy-success-banner animate-scale-up">
              <span className="banner-icon">✓</span>
              <span><strong>Script copied to clipboard!</strong> Next, open the Google Apps Script portal and paste it in.</span>
            </div>
          )}

          <div>
            <h2 className="panel-title">Deployment Guide</h2>
            <p className="panel-subtitle">
              Follow these four actionable steps to deploy and run the background automation script inside your secure Google Account.
            </p>
          </div>

          <div className="deploy-cards-grid">
            {/* Step 1 Card */}
            <div className="deploy-card">
              <div className="step-node-connector">1</div>
              <div className="deploy-card-body">
                <h3>Open Apps Script</h3>
                <p>
                  Access the Google Apps Script developer dashboard. It runs scripts natively on secure Google Cloud servers.
                </p>
                <a 
                  href="https://script.google.com" 
                  target="_blank" 
                  rel="noopener noreferrer" 
                  className="deploy-action-btn primary-link-btn"
                >
                  Open script.google.com <ExternalLink size={14} />
                </a>
              </div>
            </div>

            {/* Step 2 Card */}
            <div className="deploy-card">
              <div className="step-node-connector">2</div>
              <div className="deploy-card-body">
                <h3>Paste & Save Code</h3>
                <p>
                  Click <strong>New Project</strong>, replace all default code with your customized script, and click the <strong>Save</strong> icon.
                </p>
                <button 
                  onClick={handleCopyCode} 
                  className="deploy-action-btn secondary-btn"
                >
                  {copied ? (
                    <>
                      <Check size={14} className="text-success" /> Copied!
                    </>
                  ) : (
                    <>
                      <Copy size={14} /> Copy Script Code
                    </>
                  )}
                </button>
              </div>
            </div>

            {/* Step 3 Card */}
            <div className="deploy-card">
              <div className="step-node-connector">3</div>
              <div className="deploy-card-body">
                <h3>Run & Authorize</h3>
                <p>
                  Click the <strong>Run</strong> button at the top (play icon). Click <strong>Review Permissions</strong>, select your Google Account, click <strong>Advanced</strong> &rarr; <strong>Go to Untitled project (unsafe)</strong>, and select <strong>Allow</strong> to grant required Gmail/Drive access.
                </p>
                <div className="trigger-badge">
                  <Play size={12} /> Run manually once
                </div>
              </div>
            </div>

            {/* Step 4 Card */}
            <div className="deploy-card">
              <div className="step-node-connector">4</div>
              <div className="deploy-card-body">
                <h3>Create Daily Trigger</h3>
                <p>
                  Click the clock icon (<strong>Triggers</strong>) on the left sidebar. Add a trigger to run <code>runAirtelAutomation</code> <strong>Time-driven daily</strong>.
                </p>
                <div className="trigger-badge">
                  <Clock size={12} /> Time-driven daily trigger
                </div>
              </div>
            </div>
          </div>

          <div className="deploy-tip-banner animate-scale-up">
            <span className="banner-icon">💡</span>
            <div>
              <strong>Immediate Testing Tip:</strong> When running the script manually, if your execution log says <em>"Current date is outside the billing window. Skipping scan..."</em> (normal behavior outside the billing window), temporarily change <code>BILL_GENERATION_DAY</code> to today's date in your script settings to verify it runs and processes your emails.
            </div>
          </div>
        </div>
      )}
    </div>
  );
}
