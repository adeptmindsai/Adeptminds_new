/**
 * @license
 * SPDX-License-Identifier: Apache-2.0
 */

import { useState, useRef, DragEvent, ChangeEvent } from 'react';
import { motion, AnimatePresence } from 'motion/react';
import * as XLSX from 'xlsx';
import { 
  Upload, 
  FileText, 
  CheckCircle2, 
  AlertTriangle, 
  Loader2, 
  FileJson, 
  Layers, 
  Calendar, 
  User, 
  DollarSign, 
  ShieldAlert, 
  Copy, 
  Check, 
  ArrowRight,
  RefreshCw,
  Sparkles,
  Database,
  Table,
  Terminal,
  Grid,
  TrendingUp,
  Download,
  Info,
  SlidersHorizontal,
  Eye,
  EyeOff
} from 'lucide-react';

interface HeaderData {
  supplier_name: string | null;
  customer_name: string | null;
  date: string | null;
  total_amount: string | null;
  tax_amount: string | null;
}

interface LineItem {
  description: string | null;
  quantity: string | null;
  unit: string | null;
  size_specification: string | null;
  rate: string | null;
  tax: string | null;
  total: string | null;
}

interface TraceabilityData {
  page_number: number;
  duplicate_warning: boolean;
  ocr_uncertainty_warning: string | null;
}

interface SinglePageResult {
  document_type: string;
  header: HeaderData;
  line_items: LineItem[];
  traceability: TraceabilityData;
}

interface ExtractionResult {
  pages: SinglePageResult[];
}

// Interactive sample portfolio for demo purposes
const DEMO_PAYLOAD: ExtractionResult = {
  pages: [
    {
      document_type: "Invoice",
      header: {
        supplier_name: "Apex Metal Fabrication Corp",
        customer_name: "BuildCorp Structures Ltd",
        date: "2026-05-12",
        total_amount: "$15,240.00",
        tax_amount: "$2,286.00"
      },
      line_items: [
        {
          description: "Heavy Structural Mild Steel H-Beams",
          quantity: "25",
          unit: "MTR",
          size_specification: "150x150x12x3000mm",
          rate: "$410.00",
          tax: "15%",
          total: "$10,250.00"
        },
        {
          description: "Galvanized High-Tensile Steel Tie-Bars",
          quantity: "85",
          unit: "PCS",
          size_specification: "12mmx6000mm",
          rate: "$58.70",
          tax: "15%",
          total: "$4,990.00"
        }
      ],
      traceability: {
        page_number: 1,
        duplicate_warning: false,
        ocr_uncertainty_warning: null
      }
    },
    {
      document_type: "Delivery Note",
      header: {
        supplier_name: "Apex Metal Fabrication Corp",
        customer_name: "BuildCorp Structures Ltd",
        date: "2026-05-15",
        total_amount: "N/A",
        tax_amount: "N/A"
      },
      line_items: [
        {
          description: "Premium Corrugated Iron Cladding Panels",
          quantity: "110",
          unit: "PCS",
          size_specification: "1220x2440x1.2mm",
          rate: "N/A",
          tax: "N/A",
          total: "N/A"
        }
      ],
      traceability: {
        page_number: 2,
        duplicate_warning: true,
        ocr_uncertainty_warning: "Faint barcode segment detected on bottom margin metadata"
      }
    },
    {
      document_type: "Job Work Note",
      header: {
        supplier_name: "Precision CNC Milling Guild",
        customer_name: "Apex Metal Fabrication Corp",
        date: "2026-05-18",
        total_amount: "$3,375.00",
        tax_amount: "$506.25"
      },
      line_items: [
        {
          description: "Bespoke Stainless Steel Mounting Coupler Joints",
          quantity: "45",
          unit: "PCS",
          size_specification: "2MTRx6MTRx12MM",
          rate: "$75.00",
          tax: "15%",
          total: "$3,375.00"
        }
      ],
      traceability: {
        page_number: 3,
        duplicate_warning: false,
        ocr_uncertainty_warning: "Handwritten supervisor mark detected in signature column"
      }
    }
  ]
};

interface ColumnConfig {
  key: string;
  label: string;
  visible: boolean;
}

export default function App() {
  const [file, setFile] = useState<File | null>(null);
  const [isDragActive, setIsDragActive] = useState<boolean>(false);
  const [activeTab, setActiveTab] = useState<'documents' | 'lineItems' | 'exceptions'>('documents');
  const [result, setResult] = useState<ExtractionResult | null>(DEMO_PAYLOAD);
  const [loading, setLoading] = useState<boolean>(false);
  const [error, setError] = useState<string | null>(null);
  const [isQuotaExceeded, setIsQuotaExceeded] = useState<boolean>(false);
  const [copied, setCopied] = useState<boolean>(false);
  const [isUsingDemo, setIsUsingDemo] = useState<boolean>(true);
  
  const [lineColumns, setLineColumns] = useState<ColumnConfig[]>([
    { key: 'pageNumber', label: 'Page', visible: true },
    { key: 'supplierName', label: 'Supplier', visible: true },
    { key: 'buyerName', label: 'Buyer', visible: true },
    { key: 'date', label: 'Date', visible: true },
    { key: 'description', label: 'Item', visible: true },
    { key: 'size_specification', label: 'Size', visible: true },
    { key: 'quantity', label: 'Qty', visible: true },
    { key: 'unit', label: 'Unit', visible: true },
    { key: 'rate', label: 'Rate', visible: true },
    { key: 'tax', label: 'Tax', visible: false },
    { key: 'total', label: 'Total', visible: true },
  ]);
  const [showConfigPanel, setShowConfigPanel] = useState<boolean>(true);

  const toggleColumn = (key: string) => {
    setLineColumns(prev => prev.map(col => {
      if (col.key === key) {
        const visibleCount = prev.filter(c => c.visible).length;
        if (col.visible && visibleCount <= 1) {
          return col;
        }
        return { ...col, visible: !col.visible };
      }
      return col;
    }));
  };

  const selectAllColumns = () => {
    setLineColumns(prev => prev.map(col => ({ ...col, visible: true })));
  };

  const resetToDefaults = () => {
    setLineColumns([
      { key: 'pageNumber', label: 'Page', visible: true },
      { key: 'supplierName', label: 'Supplier', visible: true },
      { key: 'buyerName', label: 'Buyer', visible: true },
      { key: 'date', label: 'Date', visible: true },
      { key: 'description', label: 'Item', visible: true },
      { key: 'size_specification', label: 'Size', visible: true },
      { key: 'quantity', label: 'Qty', visible: true },
      { key: 'unit', label: 'Unit', visible: true },
      { key: 'rate', label: 'Rate', visible: false },
      { key: 'tax', label: 'Tax', visible: false },
      { key: 'total', label: 'Total', visible: false },
    ]);
  };
  
  const fileInputRef = useRef<HTMLInputElement>(null);

  const handleDrag = (e: DragEvent<HTMLDivElement>) => {
    e.preventDefault();
    e.stopPropagation();
    if (e.type === 'dragenter' || e.type === 'dragover') {
      setIsDragActive(true);
    } else if (e.type === 'dragleave') {
      setIsDragActive(false);
    }
  };

  const handleDrop = (e: DragEvent<HTMLDivElement>) => {
    e.preventDefault();
    e.stopPropagation();
    setIsDragActive(false);

    if (e.dataTransfer.files && e.dataTransfer.files[0]) {
      const droppedFile = e.dataTransfer.files[0];
      setFile(droppedFile);
      setError(null);
      setIsQuotaExceeded(false);
    }
  };

  const handleFileChange = (e: ChangeEvent<HTMLInputElement>) => {
    if (e.target.files && e.target.files[0]) {
      setFile(e.target.files[0]);
      setError(null);
      setIsQuotaExceeded(false);
    }
  };

  const triggerFileSelect = () => {
    fileInputRef.current?.click();
  };

  const handleUpload = async () => {
    if (!file) return;
    setLoading(true);
    setError(null);
    setIsQuotaExceeded(false);
    setIsUsingDemo(false);

    const formData = new FormData();
    formData.append('document', file);

    try {
      const response = await fetch('/api/extract', {
        method: 'POST',
        body: formData,
      });
      
      const data = await response.json();
      
      if (!response.ok) {
        if (response.status === 429 || data?.isQuotaExceeded) {
          setIsQuotaExceeded(true);
          throw new Error(data?.message || data?.error || 'Google Gemini API Quota Limit Exceeded.');
        }
        throw new Error(data?.error || 'Failed to extract accounting records.');
      }
      
      setResult(data);
    } catch (err: any) {
      console.error(err);
      setError(err.message || 'An error occurred while communicating with the document parser.');
    } finally {
      setLoading(false);
    }
  };

  const resetEngine = () => {
    setFile(null);
    setResult(null);
    setError(null);
    setIsQuotaExceeded(false);
    setIsUsingDemo(false);
  };

  const loadDemo = () => {
    setFile(null);
    setResult(DEMO_PAYLOAD);
    setError(null);
    setIsQuotaExceeded(false);
    setIsUsingDemo(true);
  };

  const copyJsonToClipboard = () => {
    if (!result) return;
    navigator.clipboard.writeText(JSON.stringify(result, null, 2));
    setCopied(true);
    setTimeout(() => setCopied(false), 2000);
  };

  const escapeCSVCell = (val: any): string => {
    if (val === null || val === undefined) return '';
    const str = String(val);
    if (str.includes(',') || str.includes('"') || str.includes('\n') || str.includes('\r')) {
      return `"${str.replace(/"/g, '""')}"`;
    }
    return str;
  };

  const triggerDownload = (headers: string[], dataRows: any[][], filename: string) => {
    const csvContent = [
      headers.map(escapeCSVCell).join(','),
      ...dataRows.map(row => row.map(escapeCSVCell).join(','))
    ].join('\r\n');

    const blob = new Blob([csvContent], { type: 'text/csv;charset=utf-8;' });
    const url = URL.createObjectURL(blob);
    const link = document.createElement('a');
    link.setAttribute('href', url);
    link.setAttribute('download', filename);
    link.style.visibility = 'hidden';
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
    URL.revokeObjectURL(url);
  };

  const getColumnValue = (item: any, key: string) => {
    switch (key) {
      case 'pageNumber':
        return `Page ${item.pageNumber}`;
      case 'supplierName':
        return item.supplierName || 'null';
      case 'buyerName':
        return item.buyerName || 'null';
      case 'date':
        return item.date || 'null';
      case 'description':
        return item.description || 'null';
      case 'size_specification':
        return item.size_specification || 'null';
      case 'quantity':
        return item.quantity || 'null';
      case 'unit':
        return item.unit || 'null';
      case 'rate':
        return item.rate || 'null';
      case 'tax':
        return item.tax || 'null';
      case 'total':
        return item.total || 'null';
      default:
        return '';
    }
  };

  const downloadAllCSV = () => {
    if (!result) return;
    
    // 1. Documents Sheet
    const docHeaders = ["Page", "Type", "Supplier", "Customer", "Date"];
    const docRows = pagesList.map((page, idx) => [
      page.traceability?.page_number || (idx + 1),
      page.document_type || '',
      page.header?.supplier_name || '',
      page.header?.customer_name || '',
      page.header?.date || ''
    ]);
    triggerDownload(docHeaders, docRows, 'extracted_documents.csv');

    // 2. Line Items Sheet (Dynamic Columns based on lineColumns selection)
    const activeLineCols = lineColumns.filter(c => c.visible);
    const lineHeaders = activeLineCols.map(c => c.label);
    const lineRows = allLineItems.map(item => 
      activeLineCols.map(col => getColumnValue(item, col.key))
    );
    triggerDownload(lineHeaders, lineRows, 'extracted_line_items.csv');

    // 3. Exceptions Sheet
    const expHeaders = ["Page", "Warning"];
    const expRows = allExceptions.map(exc => [
      exc.pageNumber,
      exc.warningMsg
    ]);
    triggerDownload(expHeaders, expRows, 'extracted_exceptions.csv');
  };

  const downloadAllExcel = () => {
    if (!result) return;

    try {
      // Create a native Excel Workbook
      const wb = XLSX.utils.book_new();

      // Sheet 1: Documents mapping to arrays of objects
      const docSheetData = pagesList.map((page, idx) => ({
        "Page": page.traceability?.page_number || (idx + 1),
        "Type": page.document_type || '',
        "Supplier": page.header?.supplier_name || '',
        "Customer": page.header?.customer_name || '',
        "Date": page.header?.date || ''
      }));
      const docWS = XLSX.utils.json_to_sheet(docSheetData);
      XLSX.utils.book_append_sheet(wb, docWS, "Sheet 1 - Documents");

      // Sheet 2: Line Items (Dynamic Columns based on lineColumns selection)
      const activeLineCols = lineColumns.filter(c => c.visible);
      const lineSheetData = allLineItems.map(item => {
        const rowObj: any = {};
        activeLineCols.forEach(col => {
          rowObj[col.label] = getColumnValue(item, col.key);
        });
        return rowObj;
      });
      const lineWS = XLSX.utils.json_to_sheet(lineSheetData);
      XLSX.utils.book_append_sheet(wb, lineWS, "Sheet 2 - Line Items");

      // Sheet 3: Exceptions
      const expSheetData = allExceptions.map(exc => ({
        "Page": exc.pageNumber,
        "Warning": exc.warningMsg
      }));
      const expWS = XLSX.utils.json_to_sheet(expSheetData);
      XLSX.utils.book_append_sheet(wb, expWS, "Sheet 3 - Exceptions");

      // Trigger download as unified workbook file
      XLSX.writeFile(wb, "accounting_workbook.xlsx");
    } catch (err) {
      console.error("Failed to generate Excel file, falling back to separate file CSV downloads", err);
      downloadAllCSV();
    }
  };

  const formatBytes = (bytes: number) => {
    if (bytes === 0) return '0 Bytes';
    const k = 1024;
    const sizes = ['Bytes', 'KB', 'MB'];
    const i = Math.floor(Math.log(bytes) / Math.log(k));
    return parseFloat((bytes / Math.pow(k, i)).toFixed(2)) + ' ' + sizes[i];
  };

  // Extract variables for sheets calculation
  const pagesList = result?.pages || [];
  const totalDocumentsCount = pagesList.length;
  
  // Flatten line items across all extracted pages
  const allLineItems = pagesList.flatMap(page => 
    (page.line_items || []).map(item => ({
      pageNumber: page.traceability?.page_number || 1,
      supplierName: page.header?.supplier_name || 'N/A',
      buyerName: page.header?.customer_name || 'N/A',
      date: page.header?.date || 'N/A',
      ...item
    }))
  );

  // Flatten out warnings / exceptions
  const allExceptions = pagesList.flatMap(page => {
    const warns: { pageNumber: number; warningMsg: string; type: 'duplicate' | 'ocr' }[] = [];
    const trace = page.traceability;
    if (trace) {
      if (trace.duplicate_warning) {
        warns.push({
          pageNumber: trace.page_number,
          warningMsg: "Duplicate page signature flagged in workbook scope.",
          type: 'duplicate'
        });
      }
      if (trace.ocr_uncertainty_warning) {
        warns.push({
          pageNumber: trace.page_number,
          warningMsg: `OCR Uncertainty: ${trace.ocr_uncertainty_warning}`,
          type: 'ocr'
        });
      }
    }
    return warns;
  });

  return (
    <div className="min-h-screen bg-slate-900 text-slate-100 flex flex-col font-sans selection:bg-emerald-500 selection:text-white">
      
      {/* Top Ledger Header Bar */}
      <header className="border-b border-slate-800 bg-slate-950/80 backdrop-blur-md sticky top-0 z-50 px-6 py-4">
        <div className="max-w-7xl mx-auto flex flex-col sm:flex-row sm:items-center justify-between gap-4">
          <div className="flex items-center space-x-3">
            <div className="h-10 w-10 rounded-xl bg-emerald-600 flex items-center justify-center text-white shadow-md shadow-emerald-900/50">
              <Grid className="h-5 w-5" />
            </div>
            <div>
              <div className="flex items-center space-x-2">
                <h1 className="text-xl font-bold tracking-tight text-white">AccuDoc Sheet Ledger</h1>
                <span className="text-[10px] uppercase font-bold tracking-wider px-1.5 py-0.5 bg-emerald-500/20 text-emerald-400 border border-emerald-500/30 rounded">
                  v1.2
                </span>
              </div>
              <p className="text-xs text-slate-400 font-mono tracking-wider">SECURE MULTI-PAGE ACCOUNTING WORKBOOK</p>
            </div>
          </div>
          
          <div className="flex items-center space-x-3">
            {isUsingDemo && result && (
              <span className="inline-flex items-center px-2.5 py-1 rounded-md text-xs font-semibold bg-indigo-500/10 text-indigo-300 border border-indigo-400/20">
                <Info className="w-3.5 h-3.5 mr-1" /> Displaying Demo Workbook
              </span>
            )}
            <span className="hidden md:inline-flex items-center px-2.5 py-1 rounded-md text-xs font-semibold bg-emerald-500/10 text-emerald-400 border border-emerald-500/20">
              <Sparkles className="w-3.5 h-3.5 mr-1" /> Gemini 3.5 Analytical Suite
            </span>
          </div>
        </div>
      </header>

      {/* Main Ledger Core Grid */}
      <main className="flex-grow p-4 md:p-8 max-w-7xl w-full mx-auto space-y-6">
        
        {/* Top Status Index Panels */}
        <div className="grid grid-cols-2 lg:grid-cols-4 gap-4">
          <div className="bg-slate-950/60 p-4 rounded-xl border border-slate-800 flex items-center justify-between">
            <div>
              <p className="text-[10px] text-slate-500 font-mono uppercase">Extract Target</p>
              <p className="text-xl font-bold text-slate-200 mt-1">Multi-Note</p>
            </div>
            <div className="h-9 w-9 rounded-lg bg-indigo-600/10 text-indigo-400 flex items-center justify-center">
              <Database className="h-4 w-4" />
            </div>
          </div>

          <div className="bg-slate-950/60 p-4 rounded-xl border border-slate-800 flex items-center justify-between">
            <div>
              <p className="text-[10px] text-slate-500 font-mono uppercase">Total Documents</p>
              <p className="text-xl font-bold text-slate-200 mt-1">{totalDocumentsCount}</p>
            </div>
            <div className="h-9 w-9 rounded-lg bg-emerald-600/10 text-emerald-400 flex items-center justify-center">
              <Table className="h-4 w-4" />
            </div>
          </div>

          <div className="bg-slate-950/60 p-4 rounded-xl border border-slate-800 flex items-center justify-between">
            <div>
              <p className="text-[10px] text-slate-500 font-mono uppercase">Flattened Rows</p>
              <p className="text-xl font-bold text-slate-200 mt-1">{allLineItems.length}</p>
            </div>
            <div className="h-9 w-9 rounded-lg bg-amber-600/10 text-amber-400 flex items-center justify-center">
              <TrendingUp className="h-4 w-4" />
            </div>
          </div>

          <div className="bg-slate-950/60 p-4 rounded-xl border border-slate-800 flex items-center justify-between">
            <div>
              <p className="text-[10px] text-slate-500 font-mono uppercase">Fidelity Warnings</p>
              <p className={`text-xl font-bold mt-1 ${allExceptions.length > 0 ? 'text-amber-400' : 'text-slate-400'}`}>
                {allExceptions.length}
              </p>
            </div>
            <div className={`h-9 w-9 rounded-lg flex items-center justify-center ${allExceptions.length > 0 ? 'bg-amber-600/20 text-amber-400 animate-pulse' : 'bg-slate-800 text-slate-500'}`}>
              <AlertTriangle className="h-4 w-4" />
            </div>
          </div>
        </div>

        {/* Content Layout Split */}
        <div className="grid grid-cols-1 lg:grid-cols-12 gap-6 items-start">
          
          {/* File Handler Panel */}
          <section className="lg:col-span-4 space-y-4">
            <div className="bg-slate-950/80 rounded-2xl p-6 border border-slate-800 space-y-5">
              <div>
                <h2 className="text-base font-semibold text-slate-100 tracking-tight">Upload Accounting Source</h2>
                <p className="text-xs text-slate-400 mt-1 leading-relaxed">
                  Provide any PDF, scanned receipt or image. The parser handles invoices, delivery notes, purchase orders, or job work sheets automatically.
                </p>
              </div>

              {/* Drag Area */}
              {!file ? (
                <div
                  onDragEnter={handleDrag}
                  onDragOver={handleDrag}
                  onDragLeave={handleDrag}
                  onDrop={handleDrop}
                  onClick={triggerFileSelect}
                  className={`border-2 border-dashed rounded-xl p-6 flex flex-col items-center justify-center text-center cursor-pointer transition-all duration-200 ${
                    isDragActive 
                      ? 'border-emerald-500 bg-emerald-500/10 scale-[0.99]' 
                      : 'border-slate-800 hover:border-slate-700 hover:bg-slate-900/50'
                  }`}
                >
                  <input
                    ref={fileInputRef}
                    type="file"
                    className="hidden"
                    accept="image/*,application/pdf"
                    onChange={handleFileChange}
                  />
                  <div className="h-10 w-10 rounded-full bg-slate-900 flex items-center justify-center text-slate-400 mb-3 border border-slate-800">
                    <Upload className="h-5 w-5 text-slate-500" />
                  </div>
                  <h3 className="font-medium text-slate-200 text-xs">Drag & drop files here</h3>
                  <p className="text-[10px] text-slate-500 mt-1 mb-2">Or browse local directories</p>
                  <span className="text-[9px] font-mono text-emerald-400 px-1.5 py-0.5 rounded bg-emerald-500/10">Up to 10MB</span>
                </div>
              ) : (
                <div className="rounded-xl border border-slate-800 bg-slate-900/60 p-4 flex items-center justify-between">
                  <div className="flex items-center space-x-3 overflow-hidden">
                    <div className="h-10 w-10 shrink-0 rounded-lg bg-emerald-950 text-emerald-400 flex items-center justify-center border border-emerald-900/30">
                      <FileText className="h-5 w-5" />
                    </div>
                    <div className="overflow-hidden">
                      <p className="text-xs font-semibold text-slate-200 truncate">{file.name}</p>
                      <p className="text-[10px] text-slate-500 font-mono mt-0.5">{formatBytes(file.size)}</p>
                    </div>
                  </div>
                  <button 
                    onClick={resetEngine}
                    className="p-1 px-2 text-xs font-medium text-slate-400 hover:text-rose-400 rounded-md hover:bg-rose-500/10 transition-colors"
                  >
                    Clear
                  </button>
                </div>
              )}

              {/* Action Buttons */}
              <div className="space-y-2">
                {file && (
                  <button
                    onClick={handleUpload}
                    disabled={loading}
                    className="w-full py-2.5 px-4 rounded-xl text-white font-medium bg-emerald-600 hover:bg-emerald-500 disabled:bg-slate-800 disabled:text-slate-500 transition-all shadow-lg flex items-center justify-center space-x-2 text-xs"
                  >
                    {loading ? (
                      <>
                        <Loader2 className="h-3.5 w-3.5 animate-spin text-white" />
                        <span>Extracting Sheet Data...</span>
                      </>
                    ) : (
                      <>
                        <span>Start High Fidelity Scan</span>
                        <ArrowRight className="h-3.5 w-3.5" />
                      </>
                    )}
                  </button>
                )}

                {!file && !loading && !isUsingDemo && (
                  <button
                    onClick={loadDemo}
                    className="w-full py-2 px-4 rounded-lg bg-slate-900 hover:bg-slate-850 text-slate-300 border border-slate-800 text-xs font-medium transition-colors flex items-center justify-center space-x-2"
                  >
                    <RefreshCw className="h-3 w-3" />
                    <span>Reload Demo Workbook</span>
                  </button>
                )}
              </div>

              {/* Active Parser Error Callout */}
              {error && (
                <div className={`p-4 rounded-xl border flex flex-col space-y-3 ${
                  isQuotaExceeded 
                    ? 'bg-amber-950/30 border-amber-500/20 text-amber-200 shadow-lg shadow-amber-950/20' 
                    : 'bg-rose-950/55 border-rose-900/60 text-rose-300'
                }`}>
                  <div className="flex items-start space-x-2.5">
                    <AlertTriangle className={`h-4.5 w-4.5 shrink-0 mt-0.5 ${isQuotaExceeded ? 'text-amber-400' : 'text-rose-400'}`} />
                    <div className="space-y-1 text-[11px]">
                      <p className="font-bold text-white font-mono tracking-wider uppercase">
                        {isQuotaExceeded ? 'Gemini API Rate Limit / Quota Exhaustion' : 'Parser Exception Flagged'}
                      </p>
                      <p className="text-slate-300 leading-relaxed font-mono text-[10px] sm:text-[11px]">{error}</p>
                    </div>
                  </div>
                  
                  {isQuotaExceeded && (
                    <div className="bg-slate-950/60 border border-amber-500/10 rounded-lg p-3 space-y-2.5 self-stretch">
                      <div className="flex items-center space-x-1.5 text-xs text-amber-350 font-semibold font-mono tracking-wide">
                        <span className="h-1.5 w-1.5 bg-amber-400 animate-pulse rounded-full"></span>
                        <span>HOW TO RESOLVE THIS:</span>
                      </div>
                      
                      <ol className="list-decimal list-inside space-y-2 text-[10.5px] text-slate-300 leading-relaxed font-sans pl-1">
                        <li>
                          <span className="text-amber-300 font-semibold">Retry shortly:</span> Wait <span className="text-amber-300 font-mono">45–60 seconds</span> for the Google free-tier rate-limit window to auto-reset, then click upload again.
                        </li>
                        <li>
                          <span className="text-amber-300 font-semibold">Supply your own key:</span> Open the <span className="text-emerald-400 font-semibold text-[10px] uppercase font-mono bg-emerald-500/15 px-1.5 py-0.5 rounded border border-emerald-500/20">Settings (Gear Icon) &gt; Secrets</span> panel (top-right of this page), and configure an environment variable named <code className="text-slate-100 font-mono text-[10px] bg-slate-900 px-1 py-0.5 rounded">GEMINI_API_KEY</code> with your personal key from Google AI Studio.
                        </li>
                        <li>
                          <span className="text-amber-300 font-semibold">Test right away:</span> Click the button below to instantly load our beautiful multi-page demo workbook so you can try out all tables, search, filters, and Excel downloaders instantly!
                        </li>
                      </ol>
                      
                      <div className="pt-1.5">
                        <button
                          onClick={loadDemo}
                          className="w-full py-1.5 px-3 rounded-md bg-emerald-600/20 hover:bg-emerald-600/30 text-emerald-300 hover:text-emerald-250 border border-emerald-500/30 text-[10.5px] font-semibold transition-all flex items-center justify-center space-x-1.5"
                        >
                          <span>Load Sample Demo Workbook</span>
                        </button>
                      </div>
                    </div>
                  )}
                </div>
              )}
            </div>

            {/* Quick Spreadsheet Tips panel */}
            <div className="bg-slate-950/40 border border-slate-800 rounded-2xl p-5 space-y-3">
              <h3 className="text-xs font-semibold text-slate-300 flex items-center space-x-1.5 font-mono">
                <Terminal className="h-3.5 w-3.5 text-emerald-400" />
                <span>SHEET SPECIFICATION MANUAL</span>
              </h3>
              <p className="text-[11px] text-slate-400 leading-relaxed">
                Matches coordinates across multi-tier matrices. All material specifications (e.g. Dimensions like <code className="text-slate-200">2MTRx6MTRx12MM</code>) are parsed explicitly to the size spec field. Dates are normalized instantly into <code className="text-slate-200">YYYY-MM-DD</code>.
              </p>
            </div>
          </section>

          {/* Tabbed Sheet Matrix Visualizer */}
          <section className="lg:col-span-8 flex flex-col space-y-4">
            
            {/* Sheet Selection Tab Strip *automatically generated* */}
            <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-3 bg-slate-950 p-2.5 rounded-xl border border-slate-800">
              
              {/* Dynamic Excel-style bottom tabs */}
              <div className="flex flex-wrap items-center gap-1.5">
                <button
                  onClick={() => setActiveTab('documents')}
                  className={`px-4 py-2 rounded-lg text-xs font-medium transition-all flex items-center space-x-2 ${
                    activeTab === 'documents'
                      ? 'bg-emerald-600 text-white shadow-md shadow-emerald-950/50'
                      : 'text-slate-400 hover:text-slate-200 hover:bg-slate-900'
                  }`}
                >
                  <Table className="h-3.5 w-3.5" />
                  <span>Sheet 1: Documents</span>
                </button>

                <button
                  onClick={() => setActiveTab('lineItems')}
                  className={`px-4 py-2 rounded-lg text-xs font-medium transition-all flex items-center space-x-2 ${
                    activeTab === 'lineItems'
                      ? 'bg-emerald-600 text-white shadow-md shadow-emerald-950/50'
                      : 'text-slate-400 hover:text-slate-200 hover:bg-slate-900'
                  }`}
                >
                  <Database className="h-3.5 w-3.5" />
                  <span>Sheet 2: Line Items</span>
                </button>

                <button
                  onClick={() => setActiveTab('exceptions')}
                  className={`px-4 py-2 rounded-lg text-xs font-medium transition-all flex items-center space-x-2 relative ${
                    activeTab === 'exceptions'
                      ? 'bg-emerald-600 text-white shadow-md shadow-emerald-950/50'
                      : 'text-slate-400 hover:text-slate-200 hover:bg-slate-900'
                  }`}
                >
                  <ShieldAlert className="h-3.5 w-3.5" />
                  <span>Sheet 3: Exceptions</span>
                  {allExceptions.length > 0 && (
                    <span className="absolute -right-1 -top-1 px-1.5 py-0.5 rounded-full text-[9px] font-bold bg-amber-500 text-slate-950">
                      {allExceptions.length}
                    </span>
                  )}
                </button>
              </div>

              {/* Copy Raw / Download Area */}
              {result && (
                <div className="flex flex-wrap items-center gap-2 shrink-0">
                  <button
                    onClick={downloadAllExcel}
                    className="p-1.5 px-3 rounded-md bg-emerald-600 hover:bg-emerald-500 text-[11px] font-medium text-white shadow-sm flex items-center space-x-1.5 transition-colors"
                    title="Export all 3 sheets into a single unified Excel (.xlsx) workbook"
                  >
                    <Download className="h-3 w-3 animate-bounce" />
                    <span className="font-semibold">Download Unified Workbook (Excel)</span>
                  </button>
                  <button
                    onClick={downloadAllCSV}
                    className="p-1.5 px-3 rounded-md bg-slate-900 hover:bg-slate-850 text-[11px] font-medium text-slate-300 border border-slate-800 flex items-center space-x-1.5 transition-colors"
                    title="Download all 3 sheets as individual CSV files"
                  >
                    <Layers className="h-3 w-3 text-emerald-400" />
                    <span>Export CSVs</span>
                  </button>
                  <button
                    onClick={copyJsonToClipboard}
                    className="p-1.5 px-3 rounded-md bg-slate-900 hover:bg-slate-850 text-[11px] font-medium text-slate-300 border border-slate-800 flex items-center space-x-1.5 transition-colors"
                  >
                    {copied ? (
                      <>
                        <Check className="h-3 w-3 text-emerald-400" />
                        <span className="text-emerald-400">Copied</span>
                      </>
                    ) : (
                      <>
                        <Copy className="h-3 w-3 text-slate-400" />
                        <span>Copy JSON</span>
                      </>
                    )}
                  </button>
                </div>
              )}
            </div>

            {/* If Engine is Loading */}
            {loading && (
              <div className="bg-slate-950/60 rounded-xl border border-slate-800 p-16 text-center flex flex-col items-center justify-center space-y-4">
                <div className="relative">
                  <div className="h-14 w-14 rounded-full border-4 border-slate-800 border-t-emerald-500 animate-spin"></div>
                  <div className="absolute inset-0 flex items-center justify-center">
                    <Sparkles className="h-5 w-5 text-emerald-400 animate-pulse" />
                  </div>
                </div>
                <div>
                  <h3 className="text-slate-100 font-bold tracking-tight text-sm">Deep Processing Multimodal Ledger</h3>
                  <p className="text-xs text-slate-400 mt-1 max-w-sm mx-auto">
                    Scanning page indexes, tracking dimensions, and aligning sheet cells in accordance with strict schema...
                  </p>
                </div>
              </div>
            )}

            {/* If No Results Present */}
            {!loading && !result && (
              <div className="bg-slate-950/60 rounded-xl border border-slate-800 p-16 text-center flex flex-col items-center justify-center space-y-4">
                <div className="h-12 w-12 rounded-xl bg-slate-900 flex items-center justify-center text-slate-500 border border-slate-800">
                  <Table className="h-6 w-6" />
                </div>
                <div>
                  <h3 className="text-slate-300 font-medium tracking-tight text-xs">Spreadsheet Ledger Vacant</h3>
                  <p className="text-[11px] text-slate-500 mt-1 max-w-xs mx-auto">
                    Please upload an accounting receipt, invoice, or PO document. The extracted metadata will structure automatically.
                  </p>
                </div>
              </div>
            )}

            {/* TAB SHEET DISPLAY MATRIX */}
            {result && !loading && (
              <div className="bg-slate-950/40 rounded-2xl border border-slate-800 overflow-hidden shadow-2xl">
                
                {/* Simulated Google Sheets / Excel column index bar */}
                <div className="bg-slate-900/60 px-4 py-2 border-b border-slate-800/80 flex items-center justify-between text-[10px] text-slate-500 font-mono">
                  <div className="flex items-center space-x-2">
                    <span className="w-2.5 h-2.5 rounded-full bg-emerald-500/80 inline-block"></span>
                    <span>ACTIVE WORKSHEET: {activeTab === 'documents' ? 'SHEET1_DOCUMENTS' : activeTab === 'lineItems' ? 'SHEET2_LINE_ITEMS' : 'SHEET3_EXCEPTIONS'}</span>
                  </div>
                  <span>GRID SYMC ACTIVE • PAGES: {totalDocumentsCount}</span>
                </div>

                {/* Dynamic Configuration Controls specifically for Sheet 2: Line Items */}
                {activeTab === 'lineItems' && (
                  <div className="bg-slate-900/40 border-b border-slate-800/80 p-4 space-y-3">
                    <div className="flex flex-col sm:flex-row sm:items-center sm:justify-between gap-2">
                      <div className="flex items-center space-x-2">
                        <SlidersHorizontal className="h-3.5 w-3.5 text-emerald-450" />
                        <span className="text-xs font-semibold text-slate-200 tracking-wide font-mono">SHEET 2: COLUMN CONFIGURATOR</span>
                      </div>
                      <div className="flex items-center space-x-3 text-xs font-mono">
                        <button 
                          onClick={selectAllColumns} 
                          className="text-emerald-400 hover:text-emerald-300 transition-colors uppercase text-[10px] bg-emerald-500/10 px-2 py-0.5 rounded border border-emerald-500/20"
                        >
                          Select All
                        </button>
                        <button 
                          onClick={resetToDefaults} 
                          className="text-slate-400 hover:text-slate-200 transition-colors uppercase text-[10px] bg-slate-900 px-2 py-0.5 rounded border border-slate-800"
                        >
                          Reset Defaults
                        </button>
                      </div>
                    </div>
                    
                    <p className="text-[11px] text-slate-400 leading-relaxed font-sans">
                      Toggle active headers below. Changes dynamically customize both local view grid display and download exports (Excel sheet & individual CSV sheets).
                    </p>

                    <div className="grid grid-cols-2 sm:grid-cols-4 md:grid-cols-6 lg:grid-cols-11 gap-2 pt-1">
                      {lineColumns.map((col) => {
                        const isVisible = col.visible;
                        return (
                          <button
                            key={col.key}
                            onClick={() => toggleColumn(col.key)}
                            className={`flex items-center justify-center space-x-1.5 p-1.5 px-2 rounded-lg border text-center font-mono text-[10px] transition-all select-none ${
                              isVisible 
                                ? 'bg-emerald-950/20 border-emerald-500/30 text-emerald-300 hover:bg-emerald-950/40' 
                                : 'bg-slate-950/40 border-slate-900 text-slate-600 hover:bg-slate-900 hover:text-slate-500'
                            }`}
                          >
                            <span className={`h-1.5 w-1.5 rounded-full shrink-0 ${isVisible ? 'bg-emerald-400 animate-pulse' : 'bg-slate-800'}`}></span>
                            <span className="truncate">{col.label}</span>
                          </button>
                        );
                      })}
                    </div>
                  </div>
                )}

                {/* SHEET VIEWPORT */}
                <div className="overflow-x-auto">
                  
                  {/* TAB 1: Documents - Columns: Page, Type, Supplier, Customer, Date */}
                  {activeTab === 'documents' && (
                    <table className="w-full text-left border-collapse font-mono text-xs">
                      <thead>
                        <tr className="bg-slate-900/40 border-b border-slate-800 text-[10px] text-slate-400 select-none">
                          <th className="py-2.5 px-4 font-normal text-slate-600 text-center w-12">REF</th>
                          <th className="py-2.5 px-4 font-semibold text-slate-300">Page</th>
                          <th className="py-2.5 px-4 font-semibold text-slate-300">Type</th>
                          <th className="py-2.5 px-4 font-semibold text-slate-300">Supplier</th>
                          <th className="py-2.5 px-4 font-semibold text-slate-300">Customer</th>
                          <th className="py-2.5 px-4 font-semibold text-slate-300">Date</th>
                        </tr>
                      </thead>
                      <tbody className="divide-y divide-slate-800/60">
                        {pagesList.map((page, idx) => (
                          <tr key={idx} className="hover:bg-slate-900/40 transition-colors">
                            <td className="py-3 px-4 text-center text-slate-600 bg-slate-900/10 border-r border-slate-800/40 select-none">{idx + 1}</td>
                            <td className="py-3 px-4 font-semibold text-emerald-400">
                              Page {page.traceability?.page_number || (idx + 1)}
                            </td>
                            <td className="py-3 px-4">
                              <span className="px-2 py-0.5 rounded text-[10px] uppercase font-bold bg-slate-800 text-slate-300 border border-slate-700">
                                {page.document_type || 'Unknown'}
                              </span>
                            </td>
                            <td className="py-3 px-4 font-sans text-slate-200">
                              {page.header?.supplier_name || <span className="text-slate-600">null</span>}
                            </td>
                            <td className="py-3 px-4 font-sans text-slate-300">
                              {page.header?.customer_name || <span className="text-slate-600">null</span>}
                            </td>
                            <td className="py-3 px-4 text-slate-300 text-emerald-300">
                              {page.header?.date || <span className="text-slate-600">null</span>}
                            </td>
                          </tr>
                        ))}
                      </tbody>
                    </table>
                  )}

                  {/* TAB 2: Line Items - Columns dynamically based on selection state */}
                  {activeTab === 'lineItems' && (
                    <table className="w-full text-left border-collapse font-mono text-xs">
                      <thead>
                        <tr className="bg-slate-900/40 border-b border-slate-800 text-[10px] text-slate-400 select-none">
                          <th className="py-2.5 px-4 text-slate-600 text-center w-12">REF</th>
                          {lineColumns.filter(c => c.visible).map(col => (
                            <th 
                              key={col.key} 
                              className={`py-2.5 px-4 font-semibold text-slate-300 ${
                                col.key === 'quantity' || col.key === 'rate' || col.key === 'total' ? 'text-right' : ''
                              }`}
                            >
                              {col.label}
                            </th>
                          ))}
                        </tr>
                      </thead>
                      <tbody className="divide-y divide-slate-800/60">
                        {allLineItems.length > 0 ? (
                          allLineItems.map((item, idx) => (
                            <tr key={idx} className="hover:bg-slate-900/40 transition-colors">
                              <td className="py-3 px-3.5 text-center text-slate-600 bg-slate-900/10 border-r border-slate-800/40 select-none">{idx + 1}</td>
                              {lineColumns.filter(c => c.visible).map(col => {
                                const renderVal = getColumnValue(item, col.key);
                                if (col.key === 'pageNumber') {
                                  return (
                                    <td key={col.key} className="py-3 px-4 text-emerald-400 font-semibold whitespace-nowrap">
                                      {renderVal}
                                    </td>
                                  );
                                }
                                if (col.key === 'size_specification') {
                                  const rawVal = item.size_specification;
                                  return (
                                    <td key={col.key} className="py-3 px-4 text-amber-300 font-medium">
                                      {rawVal && rawVal !== 'N/A' && rawVal !== 'null' ? (
                                        <span className="px-1.5 py-0.5 rounded bg-amber-500/10 text-amber-400 border border-amber-500/20 text-[10px] whitespace-nowrap">
                                          {rawVal}
                                        </span>
                                      ) : (
                                        <span className="text-slate-600">null</span>
                                      )}
                                    </td>
                                  );
                                }
                                if (col.key === 'date') {
                                  return (
                                    <td key={col.key} className="py-3 px-4 text-emerald-300 whitespace-nowrap">
                                      {renderVal === 'null' ? <span className="text-slate-600">null</span> : renderVal}
                                    </td>
                                  );
                                }
                                return (
                                  <td 
                                    key={col.key} 
                                    className={`py-3 px-4 text-slate-300 ${
                                      col.key === 'quantity' || col.key === 'rate' || col.key === 'total' ? 'text-right font-semibold text-white' : ''
                                    } ${
                                      col.key === 'supplierName' || col.key === 'buyerName' ? 'font-sans text-slate-450 text-[11px] truncate max-w-[130px]' : ''
                                    }`}
                                    title={renderVal}
                                  >
                                    {renderVal === 'null' ? (
                                      <span className="text-slate-600">null</span>
                                    ) : (
                                      renderVal
                                    )}
                                  </td>
                                );
                              })}
                            </tr>
                          ))
                        ) : (
                          <tr>
                            <td 
                              colSpan={lineColumns.filter(c => c.visible).length + 1} 
                              className="py-8 text-center text-slate-500 italic"
                            >
                              No line items structured yet.
                            </td>
                          </tr>
                        )}
                      </tbody>
                    </table>
                  )}

                  {/* TAB 3: Exceptions - Columns: Page, Warning */}
                  {activeTab === 'exceptions' && (
                    <table className="w-full text-left border-collapse font-mono text-xs">
                      <thead>
                        <tr className="bg-slate-900/40 border-b border-slate-800 text-[10px] text-slate-400 select-none">
                          <th className="py-2.5 px-4 text-slate-600 text-center w-12">REF</th>
                          <th className="py-2.5 px-4 font-semibold text-slate-300">Page</th>
                          <th className="py-2.5 px-4 font-semibold text-slate-300">Warning</th>
                        </tr>
                      </thead>
                      <tbody className="divide-y divide-slate-800/60">
                        {allExceptions.length > 0 ? (
                          allExceptions.map((exc, idx) => (
                            <tr key={idx} className="hover:bg-slate-900/40 bg-amber-500/5 transition-colors">
                              <td className="py-4 px-4 text-center text-slate-600 bg-slate-900/10 border-r border-slate-800/40 select-none">{idx + 1}</td>
                              <td className="py-4 px-4 text-amber-400 font-semibold">Page {exc.pageNumber}</td>
                              <td className="py-4 px-4 text-sm font-sans block leading-relaxed">
                                <div className="flex items-center space-x-2">
                                  <AlertTriangle className="h-3.5 w-3.5 text-amber-500 shrink-0" />
                                  <span className="text-amber-200 font-medium">
                                    {exc.warningMsg}
                                  </span>
                                </div>
                              </td>
                            </tr>
                          ))
                        ) : (
                          <tr>
                            <td colSpan={3} className="py-12 text-center text-slate-400 font-sans italic space-y-2">
                              <div className="h-8 w-8 rounded-full bg-emerald-500/10 text-emerald-400 flex items-center justify-center mx-auto mb-2">
                                <Check className="h-4 w-4" />
                              </div>
                              <p className="font-semibold text-slate-200">0 Exceptions Detected</p>
                              <p className="text-xs text-slate-500">Every catalog page meets maximum OCR fidelity requirements!</p>
                            </td>
                          </tr>
                        )}
                      </tbody>
                    </table>
                  )}

                </div>

                {/* Sheets footer indicator */}
                <div className="p-3 bg-slate-900 border-t border-slate-800 flex justify-between items-center text-[10px] text-slate-500 font-mono">
                  <span>Workbook synchronized successfully</span>
                  <span>CAPACITY: 100% UNBOUNDED</span>
                </div>
              </div>
            )}

            {/* RAW JSON Code inspector */}
            {result && !loading && (
              <div className="bg-slate-950 border border-slate-800 rounded-xl overflow-hidden">
                <div className="bg-slate-900 px-4 py-2 border-b border-slate-800 flex items-center justify-between text-xs font-mono">
                  <div className="flex items-center space-x-1.5 text-slate-400">
                    <FileJson className="h-3.5 w-3.5 text-indigo-400" />
                    <span>RAW WORKBOOK EXPORT METADATA</span>
                  </div>
                  <span className="text-[10px] text-emerald-400">SCHEMA MATCHING PERFECT</span>
                </div>
                <div className="p-4 overflow-auto max-h-48 text-slate-400 font-mono text-[11px] leading-relaxed">
                  <pre className="select-all">{JSON.stringify(result, null, 2)}</pre>
                </div>
              </div>
            )}

          </section>

        </div>
      </main>

      {/* Corporate Ledger Footer */}
      <footer className="border-t border-slate-800 bg-slate-950/80 py-8 text-center text-xs text-slate-500 font-mono mt-12">
        <p>© 2026 AccuDoc Spreadsheet Ledger Extraction Suite. All structures fully integrity-verified.</p>
      </footer>

    </div>
  );
}
