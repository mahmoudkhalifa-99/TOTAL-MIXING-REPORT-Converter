export interface SheetData {
  sheetName: string;
  description?: string;
  isRtl: boolean;
  headers: string[];
  rows: (string | number)[][];
  totalsRow?: (string | number)[];
}

export interface ExtractedDocument {
  documentTitle: string;
  documentSummary?: string;
  detectedLanguage?: string;
  detectedType?: string;
  sheets: SheetData[];
  metadata?: {
    pageCountEstimate?: string;
    rowCountTotal?: number;
    columnCountMax?: number;
  };
  notes?: string[];
}

export interface ConversionOptions {
  extractionMode: 'auto' | 'mixing_reports' | 'tables_only' | 'financial' | 'raw';
  customInstructions: string;
  parseNumbers: boolean;
  autoDetectRTL: boolean;
}

export interface HistoryItem {
  id: string;
  fileName: string;
  fileSize: number;
  timestamp: number;
  documentTitle: string;
  sheetCount: number;
  rowCount: number;
  data: ExtractedDocument;
  pdfDataUrl?: string;
}
