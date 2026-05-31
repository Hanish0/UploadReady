import React from 'react';

interface PdfPreviewProps {
  pdfUrl: string | null;
  filename: string;
  isProcessing: boolean;
}

export const PdfPreview: React.FC<PdfPreviewProps> = ({ pdfUrl, filename, isProcessing }) => {
  const triggerDownload = () => {
    if (!pdfUrl) return;
    const a = document.createElement('a');
    a.href = pdfUrl;
    a.download = filename;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
  };

  return (
    <div className="glass-panel pdf-preview-panel">
      <div className="preview-header">
        <div>
          <h2 className="panel-title">Reimbursement Package</h2>
          <p className="panel-subtitle">Cleaned, merged, and ready for submission.</p>
        </div>
        {pdfUrl && (
          <button onClick={triggerDownload} className="clay-button" style={{ padding: '0 16px', height: '38px', fontSize: '13px', borderRadius: 'var(--radius-md)' }} title="Download final PDF">
            <svg
              xmlns="http://www.w3.org/2000/svg"
              viewBox="0 0 24 24"
              fill="none"
              stroke="currentColor"
              strokeWidth="2.5"
              strokeLinecap="round"
              strokeLinejoin="round"
              className="download-icon"
              style={{ width: '14px', height: '14px', marginRight: '6px' }}
            >
              <path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4" />
              <polyline points="7 10 12 15 17 10" />
              <line x1="12" y1="15" x2="12" y2="3" />
            </svg>
            Download Package
          </button>
        )}
      </div>

      {isProcessing ? (
        <div className="preview-loading-container">
          <div className="processing-loader">
            <div className="progress-bar-fill"></div>
          </div>
          <p className="text-sm font-medium">Assembling reimbursement package...</p>
          <span className="loader-detail-text">Extracting: Bill Pg 3, Pg 1, Pg 4 + Payment Receipt</span>
        </div>
      ) : !pdfUrl ? (
        <div className="preview-empty-container">
          <svg
            xmlns="http://www.w3.org/2000/svg"
            viewBox="0 0 24 24"
            fill="none"
            stroke="currentColor"
            strokeWidth="1.5"
            strokeLinecap="round"
            strokeLinejoin="round"
            className="empty-preview-icon"
          >
            <path d="M4 19.5A2.5 2.5 0 0 1 6.5 17H20" />
            <path d="M6.5 2H20v20H6.5A2.5 2.5 0 0 1 4 19.5v-15A2.5 2.5 0 0 1 6.5 2z" />
          </svg>
          <p className="text-sm font-medium">Generated PDF preview will appear here.</p>
        </div>
      ) : (
        <div className="pdf-frame-wrapper">
          <div className="pdf-meta-ribbon">
            <span className="pdf-filename">{filename}</span>
            <span className="badge-pages">4 Pages</span>
          </div>
          <div className="pdf-iframe-container">
            <iframe
              src={`${pdfUrl}#toolbar=0&navpanes=0`}
              title="Merged PDF Preview"
              className="pdf-iframe"
            />
          </div>
        </div>
      )}
    </div>
  );
};
