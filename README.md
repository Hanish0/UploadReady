# UploadReady — Airtel Reimbursement Automation

**UploadReady** is a client-side web application designed to simplify the monthly telecom reimbursement process for Airtel users. It automates the tedious steps of manually reading PDFs, extracting tax and invoice values, trimming down statement files, and merging receipts.

👉 **Live Demo:** [https://Hanish0.github.io/UploadReady/](https://Hanish0.github.io/UploadReady/)

---

## 🔒 Privacy & Security

Reimbursement files contain sensitive personal details (bills, payment types, phone numbers, and home addresses). **UploadReady runs 100% in your browser or secure personal account.**

- **Interactive Mode**: Runs 100% inside your local web browser. No files are ever uploaded to a backend server. Text parsing, page trimming, and merging happen client-side using JavaScript.
- **Cloud Auto-Pilot**: Runs inside your own Google Account via a secure Google Apps Script. All emails and PDFs are processed natively in Google Cloud, and delivery is managed via secure Google APIs. No third-party servers ever touch your files.

---

## ✨ Features

### 💻 Interactive App Mode
- **Automatic Extraction**: Parses Airtel PDF statements and payment receipts to retrieve Invoice Number, Vendor GSTIN, CGST/SGST splits (9% / 9%), total charges, paid amounts, and billing periods in real-time.
- **Sleek Click-to-Copy**: Cards with one-click copy buttons and micro-interaction animations to quickly fill out corporate expense reports.
- **Smart PDF Cleaning**: Removes unnecessary pages (tariff rules, ISD charges, details of other numbers) and generates a compact 4-page reimbursement package ordered as:
  1. **Bill Page 3** (Summarized Tax Invoice details)
  2. **Bill Page 1** (Overview and total amounts)
  3. **Bill Page 4** (Tax and HSN breakdown)
  4. **Payment Receipt Page** (Proof of payment)
- **Automatic Naming**: Dynamically names output files based on the billing period, e.g., `Airtel_Reimbursement_May2026.pdf`.

### ☁️ Cloud Auto-Pilot Mode (Guided Workflow)
- **Background Automation Script**: A customizable script that runs inside your Google Account to check Gmail for statement files on your billing day, parse them using Google Drive OCR, merge PDF pages, and email the reimbursement package to your work Outlook inbox.
- **Guided Single-Page Dashboard**:
  - **Dynamic Background Generation**: Settings inputs are validated instantly, generating the customized script in the background without mandatory page transitions or step-locked wizards.
  - **Progressive State Feedback**: Displays color-coded badges indicating current state (`● Incomplete Configuration`, `● Ready`, and `● Script Generated & Copied`).
  - **Source Code Accordion**: Hides the 500+ lines JavaScript code block by default under a clean toggle button, keeping the interface uncluttered.
  - **Actionable Deployment Cards**: Features interactive steps with quick-access buttons (e.g. "Open script.google.com" and secondary "Copy Script Code" buttons) to guide setup.

---

## 🛠️ Technology Stack

- **Framework**: [Vite](https://vite.dev/) + [React](https://react.dev/) + [TypeScript](https://www.typescriptlang.org/)
- **Styling**: [Tailwind CSS v4](https://tailwindcss.com/) + [shadcn/ui](https://ui.shadcn.com/) + Vanilla CSS design tokens
- **Fonts**: Plus Jakarta Sans, Geist, Roboto Mono
- **PDF Extraction**: [pdfjs-dist](https://github.com/mozilla/pdf.js) (Mozilla PDF Reader)
- **PDF Manipulation**: [pdf-lib](https://pdf-lib.js.org/) (Page-level assembly & merging)
- **Hosting & CI/CD**: [GitHub Pages](https://pages.github.com/) + [GitHub Actions](https://github.com/features/actions)
