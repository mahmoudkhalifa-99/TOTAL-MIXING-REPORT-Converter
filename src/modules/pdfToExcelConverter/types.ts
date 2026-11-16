export type PageKind = 'text' | 'scanned' | 'mixed';

export type DocumentScanType = 'text' | 'scanned' | 'mixed';

export type ExtractionEngineMode = 'local_js' | 'python_fastapi';

export type CamelotFlavor = 'auto' | 'lattice' | 'stream';

export type OcrLanguageOption = 'ara+eng' | 'ara' | 'eng';

export type ExportSheetStrategy =
  | 'separate_sheets_per_table'
  | 'separate_sheets_per_page'
  | 'single_merged_sheet';

export interface PageInspection {
  pageNumber: number;
  kind: PageKind;
  textItemCount: number;
  charCount: number;
  imageCount: number;
  width: number;
  height: number;
}

export interface PdfDocumentInspection {
  fileName: string;
  fileSize: number;
  totalPages: number;
  documentKind: DocumentScanType;
  textPagesCount: number;
  scannedPagesCount: number;
  mixedPagesCount: number;
  pages: PageInspection[];
}

export interface ExtractedTable {
  id: string;
  name: string;
  pages: number[];
  isRtl: boolean;
  headers: string[];
  rows: string[][];
  engine: 'pdfjs-coords' | 'tesseract-ocr' | 'camelot-lattice' | 'camelot-stream' | 'pdfplumber';
  accuracyScore: number; // 0..100
  emptyCellRatio: number; // 0..1
  warnings: string[];
}

export interface ConversionProgress {
  percent: number;
  stage: string;
}

export interface ConversionResultSummary {
  tables: ExtractedTable[];
  processedPages: number[];
  skippedScannedPages: number[];
  overallAccuracy: number;
  warnings: string[];
  errors: string[];
}

export interface PythonBackendHealth {
  available: boolean;
  camelot: boolean;
  pdfplumber: boolean;
}
