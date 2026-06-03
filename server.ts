import express from "express";
import path from "path";
import multer from "multer";
import * as XLSX from "xlsx";
import { GoogleGenAI } from "@google/genai";
import { createServer as createViteServer } from "vite";

// Global handling of unhandled rejections and uncaught exceptions to prevent silent crashes
process.on("unhandledRejection", (reason, promise) => {
  console.error("CRITICAL: Unhandled Rejection at Promise:", promise, "reason:", reason);
});

process.on("uncaughtException", (error) => {
  console.error("CRITICAL: Uncaught Exception:", error);
});

const app = express();
const PORT = 3000;
const upload = multer({ 
  storage: multer.memoryStorage(),
  limits: { fileSize: 10 * 1024 * 1024 } // 10MB limit
});

// Lazy initialization of the Gemini client to support hot reloading of API keys
const getGeminiClient = () => {
  const apiKey = process.env.GEMINI_API_KEY;
  if (!apiKey) {
    return null;
  }
  return new GoogleGenAI({
    apiKey: apiKey,
    httpOptions: {
      headers: {
        'User-Agent': 'aistudio-build',
      }
    }
  });
};

// Helper utility to introduce delays during API retry backoff
const sleep = (ms: number) => new Promise((resolve) => setTimeout(resolve, ms));

// Helper check for transient / 503 errors and network failures
function isTransientError(error: any): boolean {
  const msg = (error?.message || String(error)).toLowerCase();
  return (
    error?.status === 503 ||
    error?.statusCode === 503 ||
    msg.includes("503") ||
    msg.includes("unavailable") ||
    msg.includes("high demand") ||
    msg.includes("temporary") ||
    msg.includes("try again later") ||
    msg.includes("overloaded") ||
    msg.includes("busy")
  );
}

// Retries requests with exponential backoff and cascades to lighter models if needed
async function generateContentWithRetryAndFallback(ai: any, params: any, maxRetries = 3) {
  let attempt = 0;
  let currentModel = params.model || "gemini-3.5-flash";

  while (attempt < maxRetries) {
    try {
      attempt++;
      console.log(`[Gemini API] Request attempt ${attempt}/${maxRetries} using model: ${currentModel}...`);
      
      const response = await ai.models.generateContent({
        ...params,
        model: currentModel,
      });
      return response;
    } catch (error: any) {
      console.error(`[Gemini API] Attempt ${attempt} failed:`, error);
      
      if (isTransientError(error) && attempt < maxRetries) {
        const delay = 1000 * attempt;
        console.warn(`[Gemini API] Transient/503 error detected, waiting ${delay}ms before retrying...`);
        await sleep(delay);
        
        // On subsequent attempts, if the primary gemini-3.5-flash series is fully congested, 
        // fall back to the lighter & highly available gemini-3.1-flash-lite as recommended
        if (attempt === maxRetries - 1 && currentModel === "gemini-3.5-flash") {
          console.warn(`[Gemini API] Cascading fallback to 'gemini-3.1-flash-lite' due to persistent overload under high demand.`);
          currentModel = "gemini-3.1-flash-lite";
        }
        continue;
      }
      
      throw error;
    }
  }
  throw new Error("Failed after maximum retries");
}

// Extracted core parsing logic for documents
async function extractDocumentData(file: Express.Multer.File) {
  const ai = getGeminiClient();
  if (!ai) {
    throw new Error("Gemini API key is not configured. Please supply an API key in the Settings > Secrets panel.");
  }

  const base64Data = file.buffer.toString("base64");

  const prompt = `
Analyze the following accounting document (Invoice, Delivery Note, Purchase Order, or Job Work Note) and extract structured data matching the schema perfectly. If there are multiple pages in the document, process each independently and return an entry for each page in the 'pages' list.

Rules:
1. Analyze every page independently.
2. Identify the document type.
3. Extract all header information (Supplier Name, Customer Name, Date (YYYY-MM-DD), Total Amount, Taxes, etc.).
4. Extract all line items (Description, Quantity, Unit, Size/Dimensions (as size_specification), Rate, Tax, Total).
5. Size/Dimension format (e.g. 4x8x10mm, 1220x2440, 2MTRx6MTRx12MM) must be captured in size_specification.
6. Page level traceability: Include page_number starting from 1 for every extracted record.
7. Detect duplicate pages and set duplicate_warning to true.
8. Detect OCR uncertainty and add warning string to ocr_uncertainty_warning.
9. If a field is missing or not present, return null. Do not infer values.
10. If a page contains multiple line items, extract all line items.
11. Return valid JSON matching the schema below. No conversational text.

Schema:
{
  "pages": [
    {
      "document_type": "string",
      "header": {
        "supplier_name": "string | null",
        "customer_name": "string | null",
        "date": "string | null",
        "total_amount": "string | null",
        "tax_amount": "string | null"
      },
      "line_items": [
        {
          "description": "string | null",
          "quantity": "string | null",
          "unit": "string | null",
          "size_specification": "string | null",
          "rate": "string | null",
          "tax": "string | null",
          "total": "string | null"
        }
      ],
      "traceability": {
        "page_number": 1,
        "duplicate_warning": false,
        "ocr_uncertainty_warning": "string | null"
      }
    }
  ]
}
`;

  // Extract using the robust retry helper
  const response = await generateContentWithRetryAndFallback(ai, {
    model: "gemini-3.5-flash",
    contents: [
      { text: prompt },
      {
        inlineData: {
          mimeType: file.mimetype,
          data: base64Data,
        },
      },
    ],
    config: {
      responseMimeType: "application/json",
    },
  });

  let rawText = response.text || "{}";
  rawText = rawText.trim();
  if (rawText.startsWith("```")) {
    rawText = rawText.replace(/^```(?:json)?\n?/, "").replace(/```$/, "").trim();
  }

  let json = JSON.parse(rawText);
  // Defensive parsing to guarantee normalized pages output
  if (!json.pages) {
    if (json.document_type || json.header || json.line_items) {
      json = { pages: [json] };
    } else if (Array.isArray(json)) {
      json = { pages: json };
    } else {
      json = { pages: [] };
    }
  }
  return json;
}

app.post("/api/extract", upload.single("document"), async (req, res) => {
  try {
    if (!req.file) {
      return res.status(400).json({ error: "No document file was uploaded." });
    }

    const data = await extractDocumentData(req.file);
    res.json(data);
  } catch (error: any) {
    console.error("Extraction error caught gracefully:", error);
    
    const dbgError = error?.message || String(error);
    const isQuotaExceeded = 
      /quota/i.test(dbgError) || 
      /RESOURCE_EXHAUSTED/i.test(dbgError) || 
      /exhausted/i.test(dbgError) ||
      error?.status === 429 ||
      error?.statusCode === 429 ||
      (error?.status && String(error.status).includes('429'));

    if (isQuotaExceeded) {
      return res.status(429).json({
        isQuotaExceeded: true,
        error: "Google Gemini API Quota Limit Exceeded.",
        message: "You have exceeded the current daily limit or rate limit of the Gemini free tier. Google AI Studio limits free API calls to 15 per minute. Please try again in 45-60 seconds, or set your own API key in Settings -> Secrets."
      });
    }

    if (isTransientError(error)) {
      return res.status(503).json({
        error: "The Google Gemini service is currently experiencing extremely high demand. We automatically attempted to retry and fall back to alternative lighter models, but the services remain busy. Please try uploading your document again in a few seconds."
      });
    }

    res.status(500).json({ 
      error: error.message || "Failed to process the accounting document." 
    });
  }
});

app.post("/api/extract-to-excel", upload.single("document"), async (req, res) => {
  try {
    if (!req.file) {
      return res.status(400).json({ error: "No document file was uploaded." });
    }

    console.log(`[Excel Automation API] Starting extraction for uploaded document: ${req.file.originalname} (${req.file.size} bytes)...`);
    const data = await extractDocumentData(req.file);
    const pages = data.pages || [];

    console.log(`[Excel Automation API] Document successfully parsed. Mapping ${pages.length} pages to sheets...`);
    
    // Create Excel Workbook
    const wb = XLSX.utils.book_new();

    // Sheet 1: Documents mapping list (ensure at least 1 row to prevent blank sheet corruption)
    const docSheetData = pages.length > 0 ? pages.map((page: any, idx: number) => ({
      "Page": page.traceability?.page_number || (idx + 1),
      "Type": page.document_type || '',
      "Supplier": page.header?.supplier_name || 'N/A',
      "Customer": page.header?.customer_name || 'N/A',
      "Date": page.header?.date || 'N/A',
      "Total Amount": page.header?.total_amount || 'N/A',
      "Tax Amount": page.header?.tax_amount || 'N/A'
    })) : [{
      "Page": "N/A",
      "Type": "No documents parsed.",
      "Supplier": "N/A",
      "Customer": "N/A",
      "Date": "N/A",
      "Total Amount": "N/A",
      "Tax Amount": "N/A"
    }];
    const docWS = XLSX.utils.json_to_sheet(docSheetData);
    XLSX.utils.book_append_sheet(wb, docWS, "Sheet 1 - Documents");

    // Sheet 2: Line Items mapping list (ensure at least 1 row to prevent blank sheet corruption)
    const lineSheetData = pages.flatMap((page: any) => {
      const pageNum = page.traceability?.page_number || 1;
      const supplierName = page.header?.supplier_name || 'N/A';
      const customerName = page.header?.customer_name || 'N/A';
      const date = page.header?.date || 'N/A';
      
      return (page.line_items || []).map((item: any) => ({
        "Page": pageNum,
        "Supplier": supplierName,
        "Buyer": customerName,
        "Date": date,
        "Item": item.description || 'N/A',
        "Size": item.size_specification || 'N/A',
        "Qty": item.quantity || 'N/A',
        "Unit": item.unit || 'N/A',
        "Rate": item.rate || 'N/A',
        "Tax": item.tax || 'N/A',
        "Total": item.total || 'N/A'
      }));
    });
    const lineWS = XLSX.utils.json_to_sheet(lineSheetData.length > 0 ? lineSheetData : [{
      "Page": "N/A",
      "Supplier": "N/A",
      "Buyer": "N/A",
      "Date": "N/A",
      "Item": "No line items found.",
      "Size": "N/A",
      "Qty": "N/A",
      "Unit": "N/A",
      "Rate": "N/A",
      "Tax": "N/A",
      "Total": "N/A"
    }]);
    XLSX.utils.book_append_sheet(wb, lineWS, "Sheet 2 - Line Items");

    // Sheet 3: Exceptions mapping warnings (ensure at least 1 row to prevent blank sheet corruption)
    const expSheetData = pages.flatMap((page: any) => {
      const warns: any[] = [];
      const trace = page.traceability;
      if (trace) {
        if (trace.duplicate_warning) {
          warns.push({
            "Page": trace.page_number,
            "Warning": "Duplicate page signature flagged in workbook scope."
          });
        }
        if (trace.ocr_uncertainty_warning) {
          warns.push({
            "Page": trace.page_number,
            "Warning": trace.ocr_uncertainty_warning
          });
        }
      }
      return warns;
    });
    const expWS = XLSX.utils.json_to_sheet(expSheetData.length > 0 ? expSheetData : [{
      "Page": "N/A",
      "Warning": "No exceptions flagged in workbook scope."
    }]);
    XLSX.utils.book_append_sheet(wb, expWS, "Sheet 3 - Exceptions");

    // Convert structured sheets to static buffer native to Node
    console.log("[Excel Automation API] Writing Excel binary stream...");
    const excelBuffer = XLSX.write(wb, { type: 'buffer', bookType: 'xlsx' });

    // Support both raw binary streaming and encoded Base64 encapsulated responses
    if (req.query.base64 === "true") {
      console.log("[Excel Automation API] Sending Base64-encoded JSON response...");
      res.setHeader("Content-Type", "application/json");
      return res.json({
        success: true,
        filename: "accounting_workbook.xlsx",
        excelBase64: excelBuffer.toString("base64"),
        pagesCount: pages.length
      });
    }

    console.log("[Excel Automation API] Sending Excel response stream...");
    res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
    res.setHeader("Content-Disposition", `attachment; filename="accounting_workbook.xlsx"`);
    res.send(excelBuffer);

  } catch (error: any) {
    console.error("[Excel Automation API] Error generating spreadsheet:", error);
    
    const dbgError = error?.message || String(error);
    const isQuotaExceeded = 
      /quota/i.test(dbgError) || 
      /RESOURCE_EXHAUSTED/i.test(dbgError) || 
      /exhausted/i.test(dbgError) ||
      error?.status === 429 ||
      error?.statusCode === 429;

    if (isQuotaExceeded) {
       res.status(429).setHeader("Content-Type", "application/json");
       return res.json({
         error: "Google Gemini API Quota Limit Exceeded. Please retry in 45 seconds."
       });
    }

    if (isTransientError(error)) {
      res.status(503).setHeader("Content-Type", "application/json");
      return res.json({
        error: "Google Gemini is temporarily overloaded under extreme request demand. Please try again in 5 seconds."
      });
    }

    res.status(500).setHeader("Content-Type", "application/json");
    res.json({ 
      error: error.message || "Failed to process document and generate Excel workbook." 
    });
  }
});

// Vite middleware
async function startServer() {
  if (process.env.NODE_ENV !== "production") {
    const vite = await createViteServer({
      server: { middlewareMode: true },
      appType: "spa",
    });
    app.use(vite.middlewares);
  } else {
    const distPath = path.join(process.cwd(), 'dist');
    app.use(express.static(distPath));
    app.get('*', (req, res) => {
      res.sendFile(path.join(distPath, 'index.html'));
    });
  }

  app.listen(PORT, "0.0.0.0", () => {
    console.log(`Server running on http://localhost:${PORT}`);
  });
}

// Start with a caught rejection safety boundary
startServer().catch((error) => {
  console.error("CRITICAL: Failed to launch server:", error);
});

