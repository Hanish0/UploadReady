# UploadReady — Airtel Reimbursement Automation

**UploadReady** is a client-side web application designed to simplify the monthly telecom reimbursement process for Airtel users. It automates the tedious steps of manually reading PDFs, extracting tax and invoice values, trimming down statement files, and merging receipts.

👉 **Live Demo:** [https://Hanish0.github.io/UploadReady/](https://Hanish0.github.io/UploadReady/)

---

## 🔒 Privacy & Security

Reimbursement files contain sensitive personal details (bills, payment types, phone numbers, and home addresses). **UploadReady runs 100% in your browser.**

- No files are ever uploaded to a backend server.
- All PDF text parsing, metadata extraction, page deletion, and merging happen client-side using WebAssembly and Javascript.
- Your personal statements remain private to your local computer.

---

## ✨ Features

- **Automatic Extraction**: Parses Airtel PDF statements and payment receipts to retrieve:
  - Mobile Number
  - Invoice Number & Vendor GSTIN
  - CGST / SGST splits (9% / 9% dynamic)
  - Invoice Amount (This Month's Charges)
  - Paid Amount (Receipt payment confirmation)
  - Billing Period range
- **Sleek Click-to-Copy**: Cards with one-click copy buttons and micro-interaction animations to quickly fill out corporate expense reports.
- **Smart PDF Cleaning**: Removes unnecessary pages (tariff rules, ISD charges, details of other numbers) and generates a compact 4-page reimbursement package ordered as:
  1. **Bill Page 3** (Summarized Tax Invoice details)
  2. **Bill Page 1** (Overview and total amounts)
  3. **Bill Page 4** (Tax and HSN breakdown)
  4. **Payment Receipt Page** (Proof of payment)
- **Automatic Naming**: Dynamically names output files based on the billing period, e.g., `Airtel_Reimbursement_May2026.pdf`.

---

## 🛠️ Technology Stack

- **Framework**: [Vite](https://vite.dev/) + [React](https://react.dev/) + [TypeScript](https://www.typescriptlang.org/)
- **Styling**: [Tailwind CSS v4](https://tailwindcss.com/) + [shadcn/ui](https://ui.shadcn.com/)
- **Fonts**: Plus Jakarta Sans, Lora, Roboto Mono
- **PDF Extraction**: [pdfjs-dist](https://github.com/mozilla/pdf.js) (Mozilla PDF Reader)
- **PDF Manipulation**: [pdf-lib](https://pdf-lib.js.org/) (Page-level assembly & merging)
- **Hosting & CI/CD**: [GitHub Pages](https://pages.github.com/) + [GitHub Actions](https://github.com/features/actions)

---
