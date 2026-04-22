"use client";

import { GlobalWorkerOptions, getDocument } from "pdfjs-dist/legacy/build/pdf.mjs";
import pdfjsPkg from "pdfjs-dist/package.json";
import { MAX_PDF_FILE_SIZE_BYTES, MAX_PDF_FILE_SIZE_MB } from "./constants";
import type { PagePipelineResult } from "./types";

GlobalWorkerOptions.workerSrc = `https://unpkg.com/pdfjs-dist@${pdfjsPkg.version}/build/pdf.worker.min.mjs`;

function isPdf(file: File): boolean {
  return file.type === "application/pdf" || file.name.toLowerCase().endsWith(".pdf");
}

function isLikelyPdf(bytes: Uint8Array): boolean {
  const probeLen = Math.min(bytes.length, 1024);
  const probe = new TextDecoder("ascii", { fatal: false }).decode(bytes.slice(0, probeLen));
  return probe.includes("%PDF-");
}

function explainNonPdf(bytes: Uint8Array): string {
  const probeLen = Math.min(bytes.length, 300);
  const preview = new TextDecoder("utf-8", { fatal: false }).decode(bytes.slice(0, probeLen)).trim();
  if (!preview) return "返回内容为空";
  if (preview.startsWith("<") || preview.toLowerCase().includes("html")) {
    return "返回内容像 HTML 错误页";
  }
  return `返回内容前缀：${preview.slice(0, 120)}`;
}

function ensurePdf(bytes: Uint8Array, scene: string): Uint8Array {
  if (isLikelyPdf(bytes)) return bytes;
  throw new Error(`${scene} 不是有效 PDF（No PDF header found）。${explainNonPdf(bytes)}`);
}

async function convertInputToPdf(file: File): Promise<Uint8Array> {
  if (!isPdf(file)) {
    throw new Error("当前仅支持 PDF 文件上传。");
  }

  if (file.size > MAX_PDF_FILE_SIZE_BYTES) {
    throw new Error(`文件超出限制：最大支持 ${MAX_PDF_FILE_SIZE_MB}MB PDF。`);
  }

  if (isPdf(file)) {
    const raw = new Uint8Array(await file.arrayBuffer());
    return ensurePdf(raw, "上传文件");
  }
  throw new Error("当前仅支持 PDF 文件上传。");
}

export async function buildPagePipeline(file: File): Promise<PagePipelineResult> {
  const rawPdfBytes = await convertInputToPdf(file);
  // 一份用于后续导出持久保存，一份给 pdf.js 预览，避免 worker 转移导致 buffer detached。
  const sourcePdfBytes = Uint8Array.from(rawPdfBytes);
  const previewPdfBytes = Uint8Array.from(rawPdfBytes);
  let pdf;
  try {
    const loadingTask = getDocument({ data: previewPdfBytes });
    pdf = await loadingTask.promise;
  } catch (error) {
    const msg = error instanceof Error ? error.message : "PDF 解析失败";
    throw new Error(`PDF 解析失败：${msg}`);
  }

  const pages = await Promise.all(
    Array.from({ length: pdf.numPages }, async (_, i) => {
      const page = await pdf.getPage(i + 1);
      const viewport = page.getViewport({ scale: 1.2 });
      const canvas = document.createElement("canvas");
      const context = canvas.getContext("2d");
      if (!context) {
        throw new Error("无法创建预览画布");
      }
      canvas.width = viewport.width;
      canvas.height = viewport.height;
      await page.render({ canvasContext: context, viewport }).promise;
      return {
        id: `${file.name}-${i + 1}`,
        image: canvas.toDataURL("image/png"),
        pdfPageIndex: i
      };
    })
  );

  return {
    sourceName: file.name,
    sourcePdfBytes,
    pages
  };
}
