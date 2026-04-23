export type Pattern = "Z" | "N";
export type Orientation = "portrait" | "landscape";
export type PaperSizeKey = "A4" | "A5" | "Letter" | "B5";

export interface Page {
  id: string;
  image: string;
  width: number;
  height: number;
  pdfPageIndex?: number;
}

export interface PagePipelineResult {
  sourceName: string;
  sourceType: "pdf" | "ppt" | "pptx";
  sourcePdfBytes?: Uint8Array;
  pages: Page[];
  diagnostics?: PipelineDiagnostics;
}

export interface PipelineProgress {
  message: string;
  current?: number;
  total?: number;
}

export interface PipelineDiagnostics {
  sourceType: "pdf" | "ppt" | "pptx";
  slidesParsedCount: number;
  slidesRenderedCount: number;
  snapshotsGeneratedCount: number;
  blankSnapshotsDetected: number;
  uniqueCanvasSizes?: string[];
  aspectRatios?: string[];
  pagesNormalizedCount?: number;
  canonicalSize?: { width: number; height: number };
  consistencyStatus?: "pass" | "warning";
  stabilityRuns?: number;
  stabilityMismatches?: number;
  stabilityHashes?: string[];
  bottomClipChecks?: Array<{
    slideIndex: number;
    slideHeight: number;
    contentBBoxBottom: number;
    finalCaptureHeight: number;
    bleedAdded: number;
    contentExceedsFrame: boolean;
  }>;
  overCaptureDetected?: boolean;
  bboxExtensionAmount?: number;
  textOverflowDetected?: boolean;
  exportImageCount?: number;
  rendererUsed?: "primary" | "fallback";
  failureReasons: string[];
}

export interface LayoutSettings {
  rows: number;
  columns: number;
  gapMm: number;
  paperSize: PaperSizeKey;
  orientation: Orientation;
  pattern: Pattern;
}
