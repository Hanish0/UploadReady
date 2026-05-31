import { useState, useEffect } from 'react';
import './App.css';
import { ThemeToggle } from './components/ThemeToggle';
import { UploadArea } from './components/UploadArea';
import { ExtractedFields } from './components/ExtractedFields';
import { PdfPreview } from './components/PdfPreview';
import { CloudPanel } from './components/CloudPanel';
import {
  extractTextFromPdf,
  extractFields,
  mergeAndCleanPdfs,
  getReimbursementFilename,
} from './utils/pdfProcessor';
import type { ExtractedFields as FieldsType } from './utils/pdfProcessor';
import { LayoutGrid, CloudLightning } from 'lucide-react';

function App() {
  const [activeTab, setActiveTab] = useState<'web' | 'cloud'>('web');
  const [billFile, setBillFile] = useState<File | null>(null);
  const [receiptFile, setReceiptFile] = useState<File | null>(null);
  
  const [isExtracting, setIsExtracting] = useState<boolean>(false);
  const [isGenerating, setIsGenerating] = useState<boolean>(false);
  
  const [extractedFields, setExtractedFields] = useState<FieldsType>({
    invoiceNumber: '',
    vendorGst: '',
    cgst: '',
    sgst: '',
    invoiceAmount: '',
    paidAmount: '',
    billingPeriod: '',
    mobileNumber: '',
  });

  const [pdfUrl, setPdfUrl] = useState<string | null>(null);
  const [filename, setFilename] = useState<string>('Airtel_Reimbursement.pdf');

  // Trigger processing automatically when both files are selected
  useEffect(() => {
    let active = true;

    async function processPdfs() {
      if (!billFile || !receiptFile) return;

      // Clean up previous PDF preview URL if it exists
      if (pdfUrl) {
        URL.revokeObjectURL(pdfUrl);
        setPdfUrl(null);
      }

      setIsExtracting(true);
      setExtractedFields({
        invoiceNumber: '',
        vendorGst: '',
        cgst: '',
        sgst: '',
        invoiceAmount: '',
        paidAmount: '',
        billingPeriod: '',
        mobileNumber: '',
      });

      try {
        // Step 1: Text extraction & regex parsing
        const [billText, receiptText] = await Promise.all([
          extractTextFromPdf(billFile),
          extractTextFromPdf(receiptFile),
        ]);

        if (!active) return;

        const parsedFields = extractFields(billText, receiptText);
        setExtractedFields(parsedFields);
        setIsExtracting(false);

        // Step 2: PDF Page Merging & Cleaning
        setIsGenerating(true);
        const mergedBlob = await mergeAndCleanPdfs(billFile, receiptFile);
        
        if (!active) return;
        
        const objectUrl = URL.createObjectURL(mergedBlob);
        setPdfUrl(objectUrl);
        
        const outputFilename = getReimbursementFilename(parsedFields.billingPeriod);
        setFilename(outputFilename);
      } catch (err) {
        console.error('Error processing PDFs:', err);
        setIsExtracting(false);
      } finally {
        if (active) {
          setIsGenerating(false);
        }
      }
    }

    processPdfs();

    return () => {
      active = false;
    };
  }, [billFile, receiptFile]);

  // Clean up object URLs on unmount
  useEffect(() => {
    return () => {
      if (pdfUrl) {
        URL.revokeObjectURL(pdfUrl);
      }
    };
  }, [pdfUrl]);

  const handleClearBill = () => {
    setBillFile(null);
    if (pdfUrl) {
      URL.revokeObjectURL(pdfUrl);
      setPdfUrl(null);
    }
    setExtractedFields({
      invoiceNumber: '',
      vendorGst: '',
      cgst: '',
      sgst: '',
      invoiceAmount: '',
      paidAmount: '',
      billingPeriod: '',
      mobileNumber: '',
    });
  };

  const handleClearReceipt = () => {
    setReceiptFile(null);
    if (pdfUrl) {
      URL.revokeObjectURL(pdfUrl);
      setPdfUrl(null);
    }
    setExtractedFields({
      invoiceNumber: '',
      vendorGst: '',
      cgst: '',
      sgst: '',
      invoiceAmount: '',
      paidAmount: '',
      billingPeriod: '',
      mobileNumber: '',
    });
  };

  return (
    <div className="app-wrapper">
      {/* Sticky Navigation Header bar */}
      <header className="sticky-nav animate-fade-in">
        <div className="nav-content">
          <div className="brand-section">
            <div className="logo-container">
              <svg
                xmlns="http://www.w3.org/2000/svg"
                viewBox="0 0 24 24"
                fill="none"
                stroke="currentColor"
                strokeWidth="2.5"
                strokeLinecap="round"
                strokeLinejoin="round"
                className="logo-svg"
              >
                <path d="M14.7 6.3a1 1 0 0 0 0 1.4l1.6 1.6a1 1 0 0 0 1.4 0l3.77-3.77a6 6 0 0 1-7.94 7.94l-6.91 6.91a2.12 2.12 0 0 1-3-3l6.91-6.91a6 6 0 0 1 7.94-7.94l-3.76 3.76z" />
              </svg>
            </div>
            <div className="app-title-group">
              <h1 className="nav-title">UploadReady</h1>
              <p className="nav-subtitle">Secure monthly reimbursement compiler</p>
            </div>
          </div>
          <ThemeToggle />
        </div>
      </header>

      <div className="app-container">
        {/* Overhauled Hero Section */}
        <section className="hero-section">
          {activeTab === 'web' ? (
            <>
              <span className="hero-badge">Cloud Auto-Pilot Available</span>
              <h1>Automate Airtel <span className="text-gradient">Reimbursements</span></h1>
              <p>
                Upload your monthly statement and payment receipt. We will merge them and extract key fields in-browser.
              </p>
            </>
          ) : (
            <>
              <span className="hero-badge">Gmail & Email Automation</span>
              <h1>Cloud Auto-Pilot <span className="text-gradient">Automation</span></h1>
              <p>
                Deploy a secure Google Apps Script to scan Gmail, extract metrics, and email reimbursement packages automatically.
              </p>
            </>
          )}
        </section>

      {/* Redesigned Mode Switcher (Segmented Control) */}
      <div className="mode-switcher-container">
        <button
          className={`mode-switcher-btn ${activeTab === 'web' ? 'active' : ''}`}
          onClick={() => setActiveTab('web')}
        >
          <LayoutGrid className="tab-btn-icon" />
          Interactive App
        </button>
        <button
          className={`mode-switcher-btn ${activeTab === 'cloud' ? 'active' : ''}`}
          onClick={() => setActiveTab('cloud')}
        >
          <CloudLightning className="tab-btn-icon" />
          Cloud Auto-Pilot
        </button>
      </div>

      {activeTab === 'web' ? (
        <>
          {/* Upload Zone Row */}
          <section className="upload-section animate-fade-in">
            <UploadArea
              id="bill-upload"
              label="Airtel Monthly Statement"
              description="Drag & drop your statement PDF here"
              file={billFile}
              onFileSelect={setBillFile}
              onFileClear={handleClearBill}
            />
            <UploadArea
              id="receipt-upload"
              label="Payment Receipt"
              description="Drag & drop your receipt PDF here"
              file={receiptFile}
              onFileSelect={setReceiptFile}
              onFileClear={handleClearReceipt}
            />
          </section>

          {/* Output Grid (Fields Card on Left, PDF Preview Card on Right) - revealed only when both files are loaded */}
          {billFile && receiptFile && (
            <main className="dashboard-grid animate-fade-in">
              <ExtractedFields fields={extractedFields} isLoading={isExtracting} />
              <PdfPreview pdfUrl={pdfUrl} filename={filename} isProcessing={isGenerating} />
            </main>
          )}
        </>
      ) : (
        <div className="animate-fade-in">
          <CloudPanel defaultWorkEmail="" />
        </div>
      )}

      {/* Footer bar */}
      <footer className="app-footer">
        <p>
          UploadReady &bull; 100% Private and Secure (Processed in-browser or inside your Google Account)
        </p>
      </footer>
      </div>
    </div>
  );
}

export default App;
