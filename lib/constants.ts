import type { PaperSizeKey } from "./types";

const MM_TO_PT = 72 / 25.4;

export const PAPER_SIZE_MM: Record<PaperSizeKey, { width: number; height: number }> = {
  A4: { width: 210, height: 297 },
  A5: { width: 148, height: 210 },
  Letter: { width: 215.9, height: 279.4 },
  B5: { width: 176, height: 250 }
};

export const PAPER_SIZE_OPTIONS: PaperSizeKey[] = ["A4", "A5", "Letter", "B5"];
export const MAX_PDF_FILE_SIZE_MB = 40;
export const MAX_PDF_FILE_SIZE_BYTES = MAX_PDF_FILE_SIZE_MB * 1024 * 1024;

export const ENCOURAGEMENTS = [
  "你已经完成最重要的一步",
  "整理过的知识会更清晰",
  "考试只是你准备好的结果"
];

export function mmToPt(mm: number): number {
  return mm * MM_TO_PT;
}
