/**
 * Airtel Reimbursement Cloud Auto-Pilot Daemon
 * -------------------------------------------------------------
 * This script runs in Google Apps Script (script.google.com).
 * It runs on a daily time-based trigger, scans Gmail for Airtel
 * postpaid bills and payment receipts within your billing window,
 * extracts fields using Google Docs OCR, merges pages, and emails
 * the reimbursement package to your work Outlook inbox.
 * -------------------------------------------------------------
 * 
 * SETUP INSTRUCTIONS:
 * 1. Go to https://script.google.com and log in with your personal Gmail.
 * 2. Click "New Project".
 * 3. Copy-paste this entire file into the editor, replacing any default code.
 * 4. Adjust the USER SETTINGS section below as needed.
 * 5. Save the project (Ctrl+S) and rename it to "Airtel Reimbursement".
 * 6. Click the clock icon on the left sidebar (Triggers) and add a new trigger:
 *    - Choose which function to run: "runAirtelAutomation"
 *    - Select event source: "Time-driven"
 *    - Select type of time based trigger: "Day timer"
 *    - Select time of day: "8 AM to 9 AM"
 * 7. Click "Save", and authorize the permissions requested (safe because it runs in your account).
 * 8. Run "runAirtelAutomation" manually once to test permissions and label creation.
 */

// ==========================================
// USER SETTINGS (Configured via UploadReady UI)
// ==========================================
var WORK_EMAIL = "yourname@work.com";        // Your Outlook work email
var BILL_GENERATION_DAY = 17;                 // Day of month Airtel bill is generated
var SCAN_WINDOW_DAYS = 10;                    // Days after bill to actively check for receipt
var SCHEDULE_DAY = 25;                       // Day of month to email Outlook (or 'immediate')
var AIRTEL_MOBILE = "";                      // Optional: Filter by 10-digit mobile number if multiple bills exist
var GMAIL_LABEL_NAME = "Reimbursement-Processed";
var DRIVE_QUEUE_FOLDER = "UploadReady_Queue";

// ==========================================
// MAIN AUTOMATION ENTRY POINT
// ==========================================
function runAirtelAutomation() {
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
  var currentMonthYearStr = Utilities.formatDate(today, Session.getScriptTimeZone(), "MMM-yyyy"); // e.g. "May-2026"
  if (isMonthAlreadyProcessed(currentMonthYearStr)) {
    Logger.log("Airtel reimbursement for " + currentMonthYearStr + " has already been processed. Exiting.");
    return;
  }
  
  // 4. Scan Gmail for the Statement and the Receipt
  Logger.log("Active scanning window open. Searching Gmail...");
  var billSearchQuery = 'subject:("Airtel Postpaid Bill" OR "Airtel e-Bill" OR "Airtel Bill") has:attachment filename:pdf -label:' + GMAIL_LABEL_NAME;
  var receiptSearchQuery = 'subject:("Airtel Payment" OR "Airtel Receipt" OR "Airtel Transaction") has:attachment filename:pdf -label:' + GMAIL_LABEL_NAME;
  
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
  var mergedBlob = mergePdfs(billBlob, receiptBlob, extractedFields.billingPeriod);
  
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

// ==========================================
// HELPER FUNCTIONS
// ==========================================

/**
 * Searches a thread for a PDF attachment (starts with the most recent message)
 */
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

/**
 * Extracts text content from a PDF using native Google Drive OCR
 */
function extractTextFromPdf(pdfBlob) {
  // Setup request body metadata for Drive API v3
  var metadata = {
    name: 'UploadReady_Temp_OCR_' + Utilities.getUuid(),
    mimeType: 'application/vnd.google-apps.document'
  };
  
  var boundary = 'multipart_boundary_uploadready';
  var requestBody = 
    '--' + boundary + '\r\n' +
    'Content-Type: application/json; charset=UTF-8\r\n\r\n' +
    JSON.stringify(metadata) + '\r\n' +
    '--' + boundary + '\r\n' +
    'Content-Type: ' + pdfBlob.getContentType() + '\r\n\r\n';
  
  var requestBytes = Utilities.newBlob(requestBody).getBytes();
  var pdfBytes = pdfBlob.getBytes();
  var footerBytes = Utilities.newBlob('\r\n--' + boundary + '--\r\n').getBytes();
  
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
  
  // Call Google Drive API to upload & convert (does OCR automatically)
  var response = UrlFetchApp.fetch('https://www.googleapis.com/upload/drive/v3/files?uploadType=multipart', options);
  var json = JSON.parse(response.getContentText());
  
  if (!json.id) {
    throw new Error('Google Drive OCR Upload failed: ' + response.getContentText());
  }
  
  var fileId = json.id;
  var extractedText = '';
  
  try {
    // Open the Google Doc and read its body text
    var doc = DocumentApp.openById(fileId);
    extractedText = doc.getBody().getText();
  } catch (err) {
    Logger.log("Error reading OCR document: " + err.toString());
  } finally {
    // Delete the temporary file immediately so Drive remains clean
    try {
      DriveApp.getFileById(fileId).setTrashed(true);
    } catch (e) {
      Logger.log("Failed to clean up temporary OCR file: " + e.toString());
    }
  }
  
  return extractedText;
}

/**
 * Regex matching functions for Airtel Postpaid Statements
 */
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
  
  // 1. Mobile Number
  var mobileMatch = billText.match(/(?:Phone\s*Number|PhoneNo|Mobile)\s*:?\s*(\d{10})/i);
  if (mobileMatch) {
    fields.mobileNumber = mobileMatch[1];
  } else {
    var receiptMobileMatch = receiptText.match(/(?:POSTPAID|Customer\s+Number\s+91)\s*(\d{10})/i);
    if (receiptMobileMatch) fields.mobileNumber = receiptMobileMatch[1];
  }
  
  // 2. Invoice Number
  var invoiceMatch = billText.match(/\b(MF[A-Z0-9]{10,20})\b/);
  if (invoiceMatch) {
    fields.invoiceNumber = invoiceMatch[1];
  }
  
  // 3. Vendor GST
  var gstMatch = billText.match(/GST\s*registration\s*no\.?:\s*([0-9]{2}[A-Z]{5}[0-9]{4}[A-Z]{1}[A-Z0-9]{1}[Z]{1}[A-Z0-9]{1})/i);
  if (gstMatch) {
    fields.vendorGst = gstMatch[1];
  }
  
  // 4. Billing Period
  var periodMatch = billText.match(/(?:Statement|Bill)\s*Period\s*:?\s*([0-9]{1,2}\s+[a-zA-Z]{3}\s+[0-9]{4}\s*-\s*[0-9]{1,2}\s+[a-zA-Z]{3}\s+[0-9]{4})/i);
  if (periodMatch) {
    fields.billingPeriod = periodMatch[1].trim();
  } else {
    var periodAltMatch = billText.match(/(?:Statement|Bill)\s*Period\s*:?\s*([0-9]{1,2}\/[0-9]{1,2}\/[0-9]{2,4}\s*-\s*[0-9]{1,2}\/[0-9]{1,2}\/[0-9]{2,4})/i);
    if (periodAltMatch) fields.billingPeriod = periodAltMatch[1].trim();
  }
  
  // 5. Invoice Amount (This Month's Charges)
  var amountMatch = billText.match(/Total\s+Amount\s*(?:Payable:?)?\s*[\`\₹\s\n]*([0-9,]+\.[0-9]{2})/i);
  var chargesMatch = billText.match(/This\s+Month's\s+Charges\s*[\+\`\₹\s\n]*([0-9,]+\.[0-9]{2})/i);
  if (chargesMatch) {
    fields.invoiceAmount = chargesMatch[1];
  } else if (amountMatch) {
    fields.invoiceAmount = amountMatch[1];
  }
  
  // 6. Paid Amount (From receipt)
  var paidMatch = receiptText.match(/Paid\s+amount\s*[\`\₹\s\n]*([0-9,]+\.[0-9]{2})/i);
  if (paidMatch) {
    fields.paidAmount = paidMatch[1];
  } else {
    var billPayableMatch = billText.match(/Amount\s+Payable\s*[\=\`\₹\s\n]*([0-9,]+\.[0-9]{2})/i);
    if (billPayableMatch) fields.paidAmount = billPayableMatch[1];
  }
  
  // 7. CGST and SGST
  var cgstMatch = billText.match(/CGST[\s\S]*?9%[\s\S]*?(\d+\.\d{2})/i);
  if (cgstMatch) {
    fields.cgst = cgstMatch[1];
  }
  var sgstMatch = billText.match(/SGST\/UTGST[\s\S]*?9%[\s\S]*?(\d+\.\d{2})/i);
  if (sgstMatch) {
    fields.sgst = sgstMatch[1];
  }
  
  if (!fields.cgst || !fields.sgst) {
    var totalGstMatch = billText.match(/Taxes\s*\(GST\)\s*[\`\₹\s\n]*([0-9,]+\.[0-9]{2})/i);
    if (totalGstMatch) {
      var totalGst = parseFloat(totalGstMatch[1].replace(/,/g, ''));
      var half = (totalGst / 2).toFixed(2);
      fields.cgst = fields.cgst || half;
      fields.sgst = fields.sgst || half;
    }
  }
  
  return fields;
}

/**
 * Loads pdf-lib and merges Bill pages [3, 1, 4] with Receipt page [1]
 */
function mergePdfs(billBlob, receiptBlob, billingPeriod) {
  // Dynamically load pdf-lib UMD library
  var pdfLibUrl = "https://unpkg.com/pdf-lib@1.17.1/dist/pdf-lib.min.js";
  var pdfLibCode = UrlFetchApp.fetch(pdfLibUrl).getContentText();
  
  // Inject into global script environment
  eval(pdfLibCode);
  
  // Since eval runs globally, PDFLib is now loaded
  var PDFDocument = PDFLib.PDFDocument;
  
  // Get binary arrays
  var billBytes = new Uint8Array(billBlob.getBytes());
  var receiptBytes = new Uint8Array(receiptBlob.getBytes());
  
  // Run async merger mapping to a synchronous-looking execution (GAS supports modern promises)
  var mergePromise = (async function() {
    var mergedPdf = await PDFDocument.create();
    var billDoc = await PDFDocument.load(billBytes);
    var receiptDoc = await PDFDocument.load(receiptBytes);
    
    var billPageCount = billDoc.getPageCount();
    var receiptPageCount = receiptDoc.getPageCount();
    
    // Copy Bill Page 3 (index 2), Page 1 (index 0), Page 4 (index 3)
    var billPagesToCopy = [];
    if (billPageCount >= 3) billPagesToCopy.push(2);
    if (billPageCount >= 1) billPagesToCopy.push(0);
    if (billPageCount >= 4) billPagesToCopy.push(3);
    
    if (billPagesToCopy.length > 0) {
      var copiedBillPages = await mergedPdf.copyPages(billDoc, billPagesToCopy);
      copiedBillPages.forEach(function(page) { mergedPdf.addPage(page); });
    }
    
    // Copy Receipt Page 1 (index 0)
    if (receiptPageCount >= 1) {
      var copiedReceiptPages = await mergedPdf.copyPages(receiptDoc, [0]);
      copiedReceiptPages.forEach(function(page) { mergedPdf.addPage(page); });
    }
    
    var mergedBytes = await mergedPdf.save();
    return mergedBytes;
  })();
  
  // Wait for the asynchronous promise to resolve
  var outputBytes = mergePromise.valueOf();
  
  // Parse filename
  var filename = "Airtel_Reimbursement.pdf";
  var endMonthYearMatch = billingPeriod.match(/-\s*\d{1,2}\s+([a-zA-Z]{3,9})\s+(\d{4})/);
  if (endMonthYearMatch) {
    var month = endMonthYearMatch[1];
    var formattedMonth = month.charAt(0).toUpperCase() + month.slice(1).toLowerCase();
    filename = "Airtel_Reimbursement_" + formattedMonth + endMonthYearMatch[2] + ".pdf";
  } else {
    var now = new Date();
    var months = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
    filename = "Airtel_Reimbursement_" + months[now.getMonth()] + now.getFullYear() + ".pdf";
  }
  
  return Utilities.newBlob(outputBytes, "application/pdf", filename);
}

/**
 * Emails the compiled reimbursement package to work Outlook email
 */
function sendReimbursementEmail(recipientEmail, fields, pdfBlob) {
  var subject = "Airtel Reimbursement Package - " + pdfBlob.getName().replace("Airtel_Reimbursement_", "").replace(".pdf", "");
  
  // Build a beautiful claymorphic/modern HTML email template
  var htmlBody = 
    '<div style="font-family: \'Segoe UI\', Tahoma, Geneva, Verdana, sans-serif; max-width: 600px; margin: 0 auto; border: 1px solid #e2e8f0; border-radius: 16px; overflow: hidden; box-shadow: 0 10px 25px -5px rgba(0, 0, 0, 0.05);">' +
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

/**
 * Saves processed files in a Google Drive folder for scheduled delivery
 */
function saveToQueue(pdfBlob, fields) {
  var parentFolder = getOrCreateFolder(DRIVE_QUEUE_FOLDER);
  
  // Save PDF
  var file = parentFolder.createFile(pdfBlob);
  
  // Store the fields JSON in the file's description
  file.setDescription(JSON.stringify(fields));
  
  Logger.log("Saved file ID " + file.getId() + " in folder " + DRIVE_QUEUE_FOLDER);
}

/**
 * Processes files waiting in the Google Drive Queue
 */
function processQueue() {
  var today = new Date();
  var currentDay = today.getDate();
  
  // Queue processing conditions:
  // If schedule day is 'immediate', the queue shouldn't really have files,
  // but if it does, or if today is >= SCHEDULE_DAY, process them!
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
      
      // Move to trash
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

/**
 * Checks if this month has already been processed based on processed labels
 */
function isMonthAlreadyProcessed(monthYearStr) {
  var label = getOrCreateGmailLabel(GMAIL_LABEL_NAME);
  // Search if any message is labeled and matches the month
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

/**
 * Gets or creates a Gmail label securely
 */
function getOrCreateGmailLabel(name) {
  var label = GmailApp.getUserLabelByName(name);
  if (!label) {
    label = GmailApp.createLabel(name);
  }
  return label;
}

/**
 * Gets or creates a Google Drive folder securely
 */
function getOrCreateFolder(name) {
  var folders = DriveApp.getFoldersByName(name);
  if (folders.hasNext()) {
    return folders.next();
  }
  return DriveApp.createFolder(name);
}
