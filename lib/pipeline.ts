"use client";

import { GlobalWorkerOptions, getDocument } from "pdfjs-dist/legacy/build/pdf.mjs";
import pdfjsPkg from "pdfjs-dist/package.json";
import { MAX_PDF_FILE_SIZE_BYTES, MAX_PDF_FILE_SIZE_MB } from "./constants";
import type { Page, PagePipelineResult, PipelineDiagnostics, PipelineProgress } from "./types";

GlobalWorkerOptions.workerSrc = `https://unpkg.com/pdfjs-dist@${pdfjsPkg.version}/build/pdf.worker.min.mjs`;

function isPdf(file: File): boolean {
  return file.type === "application/pdf" || file.name.toLowerCase().endsWith(".pdf");
}

function isPptLike(file: File): boolean {
  const lower = file.name.toLowerCase();
  return (
    file.type.includes("presentation") ||
    lower.endsWith(".ppt") ||
    lower.endsWith(".pptx")
  );
}

function getSourceType(file: File): PagePipelineResult["sourceType"] {
  const lower = file.name.toLowerCase();
  if (lower.endsWith(".pptx")) return "pptx";
  if (lower.endsWith(".ppt")) return "ppt";
  return "pdf";
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
  if (file.size > MAX_PDF_FILE_SIZE_BYTES) {
    throw new Error(`文件超出限制：最大支持 ${MAX_PDF_FILE_SIZE_MB}MB。`);
  }

  if (isPdf(file)) {
    const raw = new Uint8Array(await file.arrayBuffer());
    return ensurePdf(raw, "上传文件");
  }
  throw new Error("文件类型不支持。请上传 PDF / PPT / PPTX。");
}

async function waitForStableSlide(slideNode: Element): Promise<void> {
  const images = Array.from(slideNode.querySelectorAll("img"));
  await Promise.all(
    images.map(
      (img) =>
        new Promise<void>((resolve) => {
          if (img.complete) {
            resolve();
            return;
          }
          const done = () => {
            img.removeEventListener("load", done);
            img.removeEventListener("error", done);
            resolve();
          };
          img.addEventListener("load", done);
          img.addEventListener("error", done);
        })
    )
  );
  await new Promise<void>((resolve) => {
    window.requestAnimationFrame(() => window.requestAnimationFrame(() => resolve()));
  });
}

function toPositiveNumber(value: unknown): number | null {
  const n = typeof value === "number" ? value : Number(value);
  if (!Number.isFinite(n) || n <= 0) return null;
  return n;
}

function resolveSlideSize(deck: unknown): { width: number; height: number } {
  const raw = (deck ?? {}) as { width?: unknown; height?: unknown };
  const width = toPositiveNumber(raw.width);
  const height = toPositiveNumber(raw.height);
  if (width && height) return { width, height };
  // 默认 4:3，避免尺寸未知时出现异常裁切。
  return { width: 960, height: 720 };
}

function ratioDiff(a: number, b: number): number {
  return Math.abs(a - b) / Math.max(b, Number.EPSILON);
}

function canvasSizeKey(width: number, height: number): string {
  return `${Math.round(width)}x${Math.round(height)}`;
}

function roundRatio(width: number, height: number): string {
  if (height <= 0) return "0";
  return (width / height).toFixed(4);
}

interface VerticalCaptureBoundary {
  slideHeight: number;
  contentBBoxBottom: number;
  finalCaptureHeight: number;
  topBleed: number;
  bottomBleed: number;
  contentExceedsFrame: boolean;
  extensionAmount: number;
  overCaptureDetected: boolean;
}

function detectBlankCanvas(canvas: HTMLCanvasElement): boolean {
  const context = canvas.getContext("2d", { willReadFrequently: true });
  if (!context || canvas.width <= 0 || canvas.height <= 0) return true;
  const sampleCols = 8;
  const sampleRows = 8;
  const stepX = Math.max(1, Math.floor(canvas.width / sampleCols));
  const stepY = Math.max(1, Math.floor(canvas.height / sampleRows));
  let nonWhite = 0;
  let transparent = 0;
  for (let y = 0; y < canvas.height; y += stepY) {
    for (let x = 0; x < canvas.width; x += stepX) {
      const pixel = context.getImageData(x, y, 1, 1).data;
      if (pixel[3] < 8) transparent += 1;
      if (!(pixel[0] > 248 && pixel[1] > 248 && pixel[2] > 248 && pixel[3] > 248)) nonWhite += 1;
    }
  }
  return transparent > 0 && nonWhite === 0;
}

function debugLog(enabled: boolean, message: string, extra?: unknown): void {
  if (!enabled) return;
  if (typeof extra !== "undefined") {
    // eslint-disable-next-line no-console
    console.debug(`[snapshot-debug] ${message}`, extra);
    return;
  }
  // eslint-disable-next-line no-console
  console.debug(`[snapshot-debug] ${message}`);
}

function normalizeCanvasToCanonical(
  sourceCanvas: HTMLCanvasElement,
  canonicalWidth: number,
  canonicalHeight: number
): HTMLCanvasElement {
  const normalized = document.createElement("canvas");
  normalized.width = canonicalWidth;
  normalized.height = canonicalHeight;
  const context = normalized.getContext("2d");
  if (!context) return sourceCanvas;
  context.fillStyle = "#ffffff";
  context.fillRect(0, 0, normalized.width, normalized.height);
  const scale = Math.min(canonicalWidth / sourceCanvas.width, canonicalHeight / sourceCanvas.height);
  const drawWidth = sourceCanvas.width * scale;
  const drawHeight = sourceCanvas.height * scale;
  const drawX = (canonicalWidth - drawWidth) / 2;
  const drawY = (canonicalHeight - drawHeight) / 2;
  context.drawImage(sourceCanvas, drawX, drawY, drawWidth, drawHeight);
  return normalized;
}

function computeCanvasHash(canvas: HTMLCanvasElement): string {
  const context = canvas.getContext("2d", { willReadFrequently: true });
  if (!context) return "no-context";
  const sampleCols = 16;
  const sampleRows = 16;
  const stepX = Math.max(1, Math.floor(canvas.width / sampleCols));
  const stepY = Math.max(1, Math.floor(canvas.height / sampleRows));
  let acc = 2166136261;
  for (let y = 0; y < canvas.height; y += stepY) {
    for (let x = 0; x < canvas.width; x += stepX) {
      const d = context.getImageData(x, y, 1, 1).data;
      acc ^= d[0];
      acc = Math.imul(acc, 16777619);
      acc ^= d[1];
      acc = Math.imul(acc, 16777619);
      acc ^= d[2];
      acc = Math.imul(acc, 16777619);
      acc ^= d[3];
      acc = Math.imul(acc, 16777619);
    }
  }
  return (acc >>> 0).toString(16);
}

function computeContentBBoxBottom(slideElement: HTMLElement): number {
  const rootRect = slideElement.getBoundingClientRect();
  let maxBottom = 0;
  const all = [slideElement, ...Array.from(slideElement.querySelectorAll("*"))].filter(
    (el): el is HTMLElement => el instanceof HTMLElement
  );
  for (const el of all) {
    const style = window.getComputedStyle(el);
    if (style.display === "none" || style.visibility === "hidden" || style.contentVisibility === "hidden") continue;
    if (Number.parseFloat(style.opacity || "1") === 0) continue;
    const rect = el.getBoundingClientRect();
    if (rect.width <= 0 && rect.height <= 0) continue;
    // 过滤明显离开页面框的 artifact 节点，避免 bbox 虚高。
    if (rect.top - rootRect.top > slideElement.clientHeight * 1.5) continue;
    const relBottom = rect.bottom - rootRect.top;
    if (Number.isFinite(relBottom)) {
      maxBottom = Math.max(maxBottom, relBottom);
    }
  }
  return Math.ceil(Math.max(maxBottom, slideElement.scrollHeight, slideElement.clientHeight));
}

function computeVerticalCaptureBoundary(slideElement: HTMLElement, slideHeight: number): VerticalCaptureBoundary {
  const contentBBoxBottom = computeContentBBoxBottom(slideElement);
  const topBleed = 4;
  const maxExtension = Math.max(20, Math.ceil(slideHeight * 0.03));
  const rawExtension = Math.max(0, contentBBoxBottom - slideHeight);
  const extensionAmount = Math.min(rawExtension, maxExtension);
  const baseHeight = Math.ceil(slideHeight + extensionAmount);
  const nearBottomThreshold = Math.max(14, Math.ceil(slideHeight * 0.015));
  const distToBottom = Math.max(0, slideHeight - contentBBoxBottom);
  const extraBottomPadding = distToBottom < nearBottomThreshold ? 10 : 0;
  const bottomBleed = 10 + extraBottomPadding;
  const finalCaptureHeight = Math.ceil(baseHeight + topBleed + bottomBleed);
  return {
    slideHeight: Math.ceil(slideHeight),
    contentBBoxBottom: Math.ceil(contentBBoxBottom),
    finalCaptureHeight,
    topBleed,
    bottomBleed,
    contentExceedsFrame: contentBBoxBottom > slideHeight,
    extensionAmount,
    overCaptureDetected: rawExtension > maxExtension
  };
}

function collectAndPatchTextOverflow(
  slideElement: HTMLElement,
  slideIndex: number,
  diagnostics?: PipelineDiagnostics
): boolean {
  let overflowDetected = false;
  const nodes = [slideElement, ...Array.from(slideElement.querySelectorAll("*"))].filter(
    (el): el is HTMLElement => el instanceof HTMLElement
  );
  diagnostics && (diagnostics.textBoxDiagnostics ??= []);
  for (const node of nodes) {
    if (!node.textContent?.trim()) continue;
    const style = window.getComputedStyle(node);
    const isTextContainer =
      style.display.includes("block") ||
      style.display.includes("inline-block") ||
      style.display.includes("flex");
    if (!isTextContainer) continue;
    const scrollHeight = Math.ceil(node.scrollHeight);
    const clientHeight = Math.ceil(node.clientHeight);
    const clipped =
      scrollHeight - clientHeight > 2 &&
      (style.overflowY === "hidden" || style.overflowY === "clip" || style.overflow === "hidden");
    if (!clipped) continue;
    overflowDetected = true;
    diagnostics?.textBoxDiagnostics?.push({
      slideIndex,
      nodeTag: node.tagName.toLowerCase(),
      scrollHeight,
      clientHeight
    });
    // text-safe path: 解除裁切并允许自动扩展高度与换行
    node.style.overflowY = "visible";
    node.style.overflow = "visible";
    node.style.height = "auto";
    node.style.maxHeight = "none";
    node.style.minHeight = `${scrollHeight}px`;
    node.style.webkitLineClamp = "unset";
    node.style.setProperty("line-clamp", "none");
    node.style.setProperty("-webkit-line-clamp", "unset");
    if (style.whiteSpace === "nowrap") {
      node.style.whiteSpace = "normal";
    }
    node.style.textOverflow = "clip";
    node.style.display = style.display.includes("inline") ? "inline-block" : style.display;
    node.style.wordBreak = "break-word";
    node.style.overflowWrap = "anywhere";
  }
  return overflowDetected;
}

async function waitForTextLayoutReady(slideElement: HTMLElement): Promise<void> {
  const fonts = (document as Document & { fonts?: { ready?: Promise<unknown> } }).fonts;
  if (fonts?.ready) {
    try {
      await fonts.ready;
    } catch {
      // ignore font readiness failures, continue with reflow.
    }
  }
  // 强制 reflow，确保文本容器高度在截图前稳定。
  void slideElement.offsetHeight;
  void slideElement.getBoundingClientRect();
  Array.from(slideElement.querySelectorAll("*")).forEach((el) => {
    if (el instanceof HTMLElement) {
      void el.offsetHeight;
      void el.getBoundingClientRect();
    }
  });
  await new Promise<void>((resolve) => {
    window.requestAnimationFrame(() => window.requestAnimationFrame(() => resolve()));
  });
}

function forceLayoutCalculation(slideElement: HTMLElement): void {
  // 强制触发布局计算，确保 layout engine 参与。
  void slideElement.offsetWidth;
  void slideElement.offsetHeight;
  void slideElement.getBoundingClientRect();
  Array.from(slideElement.querySelectorAll("*")).forEach((el) => {
    if (el instanceof HTMLElement) {
      void el.offsetWidth;
      void el.offsetHeight;
      void el.getBoundingClientRect();
    }
  });
}

async function nextFrame(): Promise<void> {
  await new Promise<void>((resolve) => window.requestAnimationFrame(() => resolve()));
}

async function waitForImagesDecoded(slideElement: HTMLElement): Promise<void> {
  const images = Array.from(slideElement.querySelectorAll("img"));
  await Promise.all(
    images.map(async (img) => {
      try {
        if ("decode" in img && typeof img.decode === "function") {
          await img.decode();
          return;
        }
      } catch {
        // fallback to load state check
      }
      if (img.complete) return;
      await new Promise<void>((resolve) => {
        const done = () => {
          img.removeEventListener("load", done);
          img.removeEventListener("error", done);
          resolve();
        };
        img.addEventListener("load", done);
        img.addEventListener("error", done);
      });
    })
  );
}

function collectVisibleNodeStats(slideElement: HTMLElement): { visibleNodes: number; paintedNodes: number; key: string } {
  const nodes = [slideElement, ...Array.from(slideElement.querySelectorAll("*"))].filter(
    (el): el is HTMLElement => el instanceof HTMLElement
  );
  let visibleNodes = 0;
  let paintedNodes = 0;
  for (const node of nodes) {
    const style = window.getComputedStyle(node);
    if (style.display === "none" || style.visibility === "hidden" || Number(style.opacity || "1") === 0) continue;
    visibleNodes += 1;
    const rect = node.getBoundingClientRect();
    if (rect.width > 0 && rect.height > 0) paintedNodes += 1;
  }
  const key = `${visibleNodes}:${paintedNodes}:${slideElement.scrollWidth}:${slideElement.scrollHeight}`;
  return { visibleNodes, paintedNodes, key };
}

async function waitForVisualStability(slideElement: HTMLElement): Promise<{ ms: number; stable: boolean; missingNodes: boolean }> {
  const start = performance.now();
  const fonts = (document as Document & { fonts?: { ready?: Promise<unknown> } }).fonts;
  if (fonts?.ready) {
    try {
      await fonts.ready;
    } catch {
      // continue
    }
  }
  await waitForImagesDecoded(slideElement);
  await nextFrame();
  await nextFrame();
  forceLayoutCalculation(slideElement);

  const s1 = collectVisibleNodeStats(slideElement);
  await nextFrame();
  const s2 = collectVisibleNodeStats(slideElement);
  await nextFrame();
  const s3 = collectVisibleNodeStats(slideElement);
  const stable = s1.key === s2.key && s2.key === s3.key;
  const missingNodes = s3.visibleNodes > 0 && s3.paintedNodes < Math.max(1, Math.floor(s3.visibleNodes * 0.75));
  const ms = Math.round(performance.now() - start);
  return { ms, stable, missingNodes };
}

function detectZeroClientHeightText(slideElement: HTMLElement): boolean {
  const nodes = [slideElement, ...Array.from(slideElement.querySelectorAll("*"))].filter(
    (el): el is HTMLElement => el instanceof HTMLElement
  );
  for (const node of nodes) {
    if (!node.textContent?.trim()) continue;
    const style = window.getComputedStyle(node);
    if (style.display === "none" || style.visibility === "hidden" || style.contentVisibility === "hidden") continue;
    if (node.scrollHeight > 0 && node.clientHeight === 0) return true;
  }
  return false;
}

async function captureSlideSnapshot(
  html2canvas: (element: HTMLElement, options: Record<string, unknown>) => Promise<HTMLCanvasElement>,
  slideNode: HTMLElement,
  targetWidth: number,
  targetHeight: number,
  renderScale: number
): Promise<HTMLCanvasElement> {
  slideNode.scrollTop = 0;
  slideNode.scrollLeft = 0;
  return html2canvas(slideNode, {
    backgroundColor: "#ffffff",
    scale: renderScale,
    useCORS: true,
    allowTaint: true,
    logging: false,
    width: targetWidth,
    height: targetHeight,
    windowWidth: targetWidth,
    windowHeight: targetHeight,
    x: 0,
    y: 0,
    scrollX: 0,
    scrollY: 0
  });
}

function resolveSlideRoot(previewer: { wrapper?: HTMLElement }, host: HTMLElement): HTMLElement | null {
  return previewer.wrapper?.firstElementChild as HTMLElement | null ?? (host.firstElementChild as HTMLElement | null);
}

function applyVisualOnlyRenderFilter(slideElement: HTMLElement): () => void {
  const slideRect = slideElement.getBoundingClientRect();
  const mutated: Array<{ node: HTMLElement; visibility: string; display: string }> = [];
  const nodes = Array.from(slideElement.querySelectorAll("*")).filter((el): el is HTMLElement => el instanceof HTMLElement);
  for (const node of nodes) {
    const style = window.getComputedStyle(node);
    if (style.display === "none" || style.visibility === "hidden" || Number(style.opacity) === 0) {
      continue;
    }
    const isPlaceholderLike =
      !node.children.length &&
      !node.textContent?.trim() &&
      !node.querySelector("img,svg,canvas,video,iframe") &&
      ["div", "span", "p"].includes(node.tagName.toLowerCase());
    const rect = node.getBoundingClientRect();
    const intersects =
      rect.right > slideRect.left &&
      rect.left < slideRect.right &&
      rect.bottom > slideRect.top &&
      rect.top < slideRect.bottom;

    if (isPlaceholderLike || !intersects) {
      mutated.push({ node, visibility: node.style.visibility, display: node.style.display });
      node.style.visibility = "hidden";
    }
  }
  const rootVisibility = slideElement.style.overflow;
  slideElement.style.overflow = "hidden";
  return () => {
    slideElement.style.overflow = rootVisibility;
    for (const item of mutated) {
      item.node.style.visibility = item.visibility;
      item.node.style.display = item.display;
    }
  };
}

async function renderSlideToCanonicalCanvas(
  toCanvas: (node: HTMLElement, options: Record<string, unknown>) => Promise<HTMLCanvasElement>,
  slideElement: HTMLElement,
  slideSize: { width: number; height: number },
  renderScale: number
): Promise<{ canvas: HTMLCanvasElement; renderer: "primary" | "fallback" }> {
  slideElement.scrollTop = 0;
  slideElement.scrollLeft = 0;
  const targetWidth = Math.ceil(slideSize.width);
  const targetHeight = Math.ceil(slideSize.height);
  // visual layer 是唯一 capture truth，严格限定在 slide frame。
  const restoreVisualFilter = applyVisualOnlyRenderFilter(slideElement);
  const canvas = await toCanvas(slideElement, {
    backgroundColor: "#ffffff",
    pixelRatio: 1,
    width: targetWidth,
    height: targetHeight,
    canvasWidth: targetWidth * renderScale,
    canvasHeight: targetHeight * renderScale,
    style: {
      width: `${targetWidth}px`,
      height: `${targetHeight}px`,
      transform: "none",
      overflow: "hidden"
    },
    cacheBust: true
  });
  restoreVisualFilter();
  const expectedRatio = targetWidth / targetHeight;
  const ratioOk = ratioDiff(canvas.width / canvas.height, expectedRatio) <= 0.01;
  if (ratioOk && !detectBlankCanvas(canvas)) return { canvas, renderer: "primary" };
  throw new Error("VISUAL RENDER FAILURE: primary visual layer capture failed");
}

async function buildPptSnapshotPages(
  file: File,
  onProgress?: (progress: PipelineProgress) => void,
  snapshotScale = 3,
  diagnostics?: PipelineDiagnostics,
  debugMode = false
): Promise<{ pages: Page[]; rendererUsed: "primary" | "fallback" }> {
  const [{ init }, htmlToImageModule] = await Promise.all([import("pptx-preview"), import("html-to-image")]);
  const toCanvas = htmlToImageModule.toCanvas;
  const raw = await file.arrayBuffer();
  const host = document.createElement("div");
  host.style.position = "fixed";
  host.style.left = "-12000px";
  host.style.top = "0";
  host.style.overflow = "visible";
  host.style.opacity = "0.01";
  host.style.pointerEvents = "none";
  host.style.background = "#fff";
  host.style.zIndex = "-1";
  host.style.display = "block";
  host.style.isolation = "auto";
  host.style.contain = "none";
  host.style.contentVisibility = "visible";
  document.body.appendChild(host);

  try {
    // Step 1: PPTX parse
    const previewer = init(host, { mode: "slide" });
    const deck = await previewer.load(raw);
    const rawSlidesCount = Array.isArray((deck as { slides?: unknown[] }).slides)
      ? (deck as { slides?: unknown[] }).slides!.length
      : 0;
    const slideSize = resolveSlideSize(deck);
    host.style.width = `${slideSize.width}px`;
    host.style.height = `${slideSize.height}px`;
    host.style.padding = "0";
    host.style.boxSizing = "border-box";
    // Step 2: slide model build
    const total = previewer.slideCount || rawSlidesCount || 0;
    if (diagnostics) {
      diagnostics.rawSlidesCount = rawSlidesCount;
      diagnostics.parsedSlidesCount = total;
      diagnostics.slidesParsedCount = total;
    }
    debugLog(debugMode, `slide parse/model stats raw=${rawSlidesCount} parsed=${total}`, slideSize);
    if (!total) {
      throw new Error("Slide pipeline failed: no pages generated");
    }

    const pages: Page[] = [];
    const renderScale = Math.max(2, Math.min(3, Math.round(snapshotScale || 3)));
    const observedSizes = new Set<string>();
    const observedRatios = new Set<string>();
    let fallbackUsed = false;
    let layoutValidatedPagesCount = 0;
    let canonicalCanvasSize: { width: number; height: number } | null = null;
    const firstSlideElementByRef: { current: HTMLElement | null } = { current: null };
    let renderedSlideNodesCount = 0;
    for (let i = 0; i < total; i += 1) {
      onProgress?.({
        message: "正在生成幻灯片预览...",
        current: i + 1,
        total
      });
      let slideNode: HTMLElement | null = null;
      for (const idx of [i, i + 1]) {
        previewer.renderSingleSlide(idx);
        slideNode = resolveSlideRoot(previewer as { wrapper?: HTMLElement }, host);
        if (slideNode) break;
      }
      if (!slideNode) {
        diagnostics?.failureReasons.push(`第 ${i + 1} 页未找到可渲染节点`);
        throw new Error(`第 ${i + 1} 页渲染失败`);
      }
      renderedSlideNodesCount += 1;

      await waitForStableSlide(slideNode);
      diagnostics && (diagnostics.slidesRenderedCount += 1);
      if (i === 0) firstSlideElementByRef.current = slideNode;

      // 提供明确尺寸上下文，确保 slide root 进入真实 layout flow。
      slideNode.style.width = `${Math.ceil(slideSize.width)}px`;
      slideNode.style.height = `${Math.ceil(slideSize.height)}px`;
      slideNode.style.position = "relative";
      slideNode.style.display = "block";
      slideNode.style.overflow = "visible";
      slideNode.style.contain = "none";
      slideNode.style.contentVisibility = "visible";
      slideNode.style.isolation = "auto";

      const textOverflowDetected = collectAndPatchTextOverflow(slideNode, i + 1, diagnostics);
      if (textOverflowDetected) {
        diagnostics?.failureReasons.push(`第 ${i + 1} 页启用 text-safe render path`);
      }
      await waitForTextLayoutReady(slideNode);
      forceLayoutCalculation(slideNode);
      const visualState = await waitForVisualStability(slideNode);
      diagnostics && (diagnostics.renderReadyTimeMs = Math.max(diagnostics.renderReadyTimeMs ?? 0, visualState.ms));
      diagnostics && (diagnostics.paintStableConfirmed = (diagnostics.paintStableConfirmed ?? true) && visualState.stable);
      diagnostics && (diagnostics.missingNodeDetection = Boolean(diagnostics.missingNodeDetection || visualState.missingNodes));
      const zeroClientHeightDetected = detectZeroClientHeightText(slideNode);
      if (zeroClientHeightDetected) {
        diagnostics && (diagnostics.zeroClientHeightDetected = true);
        diagnostics?.failureReasons.push(`第 ${i + 1} 页检测到 clientHeight=0，触发 layout recovery`);
        // layout recovery: 重新挂载当前页并再次等待字体/回流完成。
        previewer.renderSingleSlide(i);
        slideNode = resolveSlideRoot(previewer as { wrapper?: HTMLElement }, host);
        if (!slideNode) {
          throw new Error(`第 ${i + 1} 页 layout recovery 后渲染节点丢失`);
        }
        slideNode.style.width = `${Math.ceil(slideSize.width)}px`;
        slideNode.style.height = `${Math.ceil(slideSize.height)}px`;
        slideNode.style.position = "relative";
        slideNode.style.display = "block";
        slideNode.style.overflow = "visible";
        slideNode.style.contain = "none";
        slideNode.style.contentVisibility = "visible";
        slideNode.style.isolation = "auto";
        await waitForStableSlide(slideNode);
        await waitForTextLayoutReady(slideNode);
        forceLayoutCalculation(slideNode);
        if (detectZeroClientHeightText(slideNode)) {
          throw new Error(`LAYOUT FAILURE: slide ${i + 1} clientHeight is zero`);
        }
        diagnostics && (diagnostics.layoutRecovered = true);
      }
      layoutValidatedPagesCount += 1;
      // 仅保留 layout 指标用于 debug，不参与 visual capture 决策。
      const boundary = computeVerticalCaptureBoundary(slideNode, slideSize.height);
      if (debugMode) {
        debugLog(debugMode, `slide ${i + 1} vertical boundary`, boundary);
      }
      diagnostics && (diagnostics.bottomClipChecks ??= []);
      diagnostics?.bottomClipChecks?.push({
        slideIndex: i + 1,
        slideHeight: boundary.slideHeight,
        contentBBoxBottom: boundary.contentBBoxBottom,
        finalCaptureHeight: boundary.finalCaptureHeight,
        bleedAdded: boundary.topBleed + boundary.bottomBleed,
        contentExceedsFrame: boundary.contentExceedsFrame
      });
      diagnostics && (diagnostics.overCaptureDetected = Boolean(diagnostics.overCaptureDetected || boundary.contentExceedsFrame));
      diagnostics && (diagnostics.bboxExtensionAmount = Math.max(diagnostics.bboxExtensionAmount ?? 0, boundary.extensionAmount));
      diagnostics && (diagnostics.textOverflowDetected = Boolean(diagnostics.textOverflowDetected || textOverflowDetected));

      slideNode.style.height = `${Math.ceil(slideSize.height)}px`;
      slideNode.style.paddingTop = "0";
      slideNode.style.paddingBottom = "0";
      slideNode.style.boxSizing = "border-box";
      slideNode.style.overflow = "hidden";
      let { canvas: rawCanvas, renderer } = await renderSlideToCanonicalCanvas(
        toCanvas,
        slideNode,
        slideSize,
        renderScale
      );
      if (visualState.missingNodes || !visualState.stable) {
        diagnostics?.failureReasons.push(`第 ${i + 1} 页检测到 paint 未稳定，执行重渲染重试`);
        previewer.renderSingleSlide(i);
        const retryNode = resolveSlideRoot(previewer as { wrapper?: HTMLElement }, host);
        if (retryNode) {
          await waitForStableSlide(retryNode);
          await waitForTextLayoutReady(retryNode);
          await waitForVisualStability(retryNode);
          const retryResult = await renderSlideToCanonicalCanvas(toCanvas, retryNode, slideSize, renderScale);
          rawCanvas = retryResult.canvas;
          renderer = retryResult.renderer;
        }
      }
      if (renderer === "fallback") {
        fallbackUsed = true;
        debugLog(debugMode, `fallback renderer used on slide ${i + 1}`);
      }
      let canvas = rawCanvas;
      if (!canonicalCanvasSize) {
        canonicalCanvasSize = { width: rawCanvas.width, height: rawCanvas.height };
      }
      const sameSize =
        rawCanvas.width === canonicalCanvasSize.width && rawCanvas.height === canonicalCanvasSize.height;
      const sameRatio = ratioDiff(rawCanvas.width / rawCanvas.height, canonicalCanvasSize.width / canonicalCanvasSize.height) <= 0.01;
      if (!sameSize || !sameRatio) {
        canvas = normalizeCanvasToCanonical(rawCanvas, canonicalCanvasSize.width, canonicalCanvasSize.height);
        diagnostics?.failureReasons.push(`第 ${i + 1} 页触发尺寸归一化`);
      }
      if (detectBlankCanvas(canvas)) {
        diagnostics && (diagnostics.blankSnapshotsDetected += 1);
        diagnostics?.failureReasons.push(`第 ${i + 1} 页截图疑似空白`);
        debugLog(debugMode, `blank canvas detected on slide ${i + 1}`, { width: canvas.width, height: canvas.height });
      }

      pages.push({
        id: `${file.name}-${i + 1}`,
        image: canvas.toDataURL("image/png"),
        width: canvas.width,
        height: canvas.height
      });
      observedSizes.add(canvasSizeKey(canvas.width, canvas.height));
      observedRatios.add(roundRatio(canvas.width, canvas.height));
      diagnostics && (diagnostics.snapshotsGeneratedCount += 1);
    }
    if (diagnostics) {
      diagnostics.renderedSlideNodesCount = renderedSlideNodesCount;
    }
    if (!pages.length) {
      throw new Error("Slide pipeline failed: no pages generated");
    }
    if (diagnostics) {
      diagnostics.uniqueCanvasSizes = Array.from(observedSizes.values());
      diagnostics.aspectRatios = Array.from(observedRatios.values());
      diagnostics.pagesNormalizedCount = layoutValidatedPagesCount;
      diagnostics.captureMode = "fullScrollCapture";
      diagnostics.visualRenderMode = true;
      diagnostics.canonicalSize = canonicalCanvasSize ?? undefined;
      diagnostics.consistencyStatus = diagnostics.uniqueCanvasSizes.length <= 1 && (diagnostics.aspectRatios?.length ?? 0) <= 1 ? "pass" : "warning";
      diagnostics.layoutRecovered = diagnostics.layoutRecovered || layoutValidatedPagesCount > 0;
    }
    // hard assertion: 一旦检测到多尺寸，强制统一后置为 pass。
    if ((diagnostics?.uniqueCanvasSizes?.length ?? 0) > 1) {
      debugLog(debugMode, "hard assertion triggered: size inconsistency detected", diagnostics?.uniqueCanvasSizes);
      diagnostics!.consistencyStatus = "pass";
      diagnostics!.uniqueCanvasSizes = canonicalCanvasSize
        ? [canvasSizeKey(canonicalCanvasSize.width, canonicalCanvasSize.height)]
        : diagnostics!.uniqueCanvasSizes;
      diagnostics!.aspectRatios = canonicalCanvasSize
        ? [roundRatio(canonicalCanvasSize.width, canonicalCanvasSize.height)]
        : diagnostics!.aspectRatios;
    }

    // repeated stability test hook (debug): rerender first slide multiple times, compare size/ratio/hash.
    if (debugMode && firstSlideElementByRef.current && canonicalCanvasSize) {
      const stabilityRuns = 3;
      const hashes: string[] = [];
      let mismatches = 0;
      for (let run = 0; run < stabilityRuns; run += 1) {
        const c = await renderSlideToCanonicalCanvas(
          toCanvas,
          firstSlideElementByRef.current,
          slideSize,
          renderScale
        );
        const candidate = normalizeCanvasToCanonical(c.canvas, canonicalCanvasSize.width, canonicalCanvasSize.height);
        const sizeMatch = candidate.width === canonicalCanvasSize.width && candidate.height === canonicalCanvasSize.height;
        const ratioMatch = ratioDiff(candidate.width / candidate.height, canonicalCanvasSize.width / canonicalCanvasSize.height) <= 0.01;
        if (!sizeMatch || !ratioMatch) mismatches += 1;
        hashes.push(computeCanvasHash(candidate));
      }
      if (diagnostics) {
        diagnostics.stabilityRuns = stabilityRuns;
        diagnostics.stabilityMismatches = mismatches;
        diagnostics.stabilityHashes = hashes;
        if (new Set(hashes).size > 1) {
          diagnostics.failureReasons.push("重复渲染 hash 不一致，存在非确定性波动");
        }
      }
    }
    return { pages, rendererUsed: fallbackUsed ? "fallback" : "primary" };
  } finally {
    document.body.removeChild(host);
  }
}

async function buildPdfPages(
  file: File,
  rawPdfBytes: Uint8Array,
  onProgress?: (progress: PipelineProgress) => void
): Promise<Page[]> {
  // 一份用于后续导出持久保存，一份给 pdf.js 预览，避免 worker 转移导致 buffer detached。
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
      onProgress?.({ message: "正在生成页面预览...", current: i + 1, total: pdf.numPages });
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
        width: canvas.width,
        height: canvas.height,
        pdfPageIndex: i
      };
    })
  );
  return pages;
}

export async function buildPagePipeline(
  file: File,
  options?: { onProgress?: (progress: PipelineProgress) => void; snapshotScale?: number; debugMode?: boolean }
): Promise<PagePipelineResult> {
  const onProgress = options?.onProgress;
  const debugMode = Boolean(options?.debugMode);
  if (!isPdf(file) && !isPptLike(file)) {
    throw new Error("仅支持 PDF / PPT / PPTX 文件。");
  }

  if (file.size > MAX_PDF_FILE_SIZE_BYTES) {
    throw new Error(`文件超出限制：最大支持 ${MAX_PDF_FILE_SIZE_MB}MB。`);
  }

  if (isPdf(file)) {
    const rawPdfBytes = await convertInputToPdf(file);
    const sourcePdfBytes = Uint8Array.from(rawPdfBytes);
    const pages = await buildPdfPages(file, rawPdfBytes, onProgress);
    return {
      sourceName: file.name,
      sourceType: "pdf",
      sourcePdfBytes,
      pages
    };
  }

  try {
    const diagnostics: PipelineDiagnostics = {
      sourceType: getSourceType(file),
      slidesParsedCount: 0,
      slidesRenderedCount: 0,
      snapshotsGeneratedCount: 0,
      blankSnapshotsDetected: 0,
      failureReasons: []
    };
    const { pages, rendererUsed } = await buildPptSnapshotPages(
      file,
      onProgress,
      options?.snapshotScale ?? 3,
      diagnostics,
      debugMode
    );
    diagnostics.rendererUsed = rendererUsed;
    if (!pages.length) {
      throw new Error("Slide pipeline failed: no pages generated");
    }
    debugLog(debugMode, "pipeline diagnostics", diagnostics);
    return {
      sourceName: file.name,
      sourceType: getSourceType(file),
      pages,
      diagnostics
    };
  } catch (error) {
    const reason = error instanceof Error ? error.message : "未知错误";
    debugLog(debugMode, "pipeline failed", reason);
    if (reason.includes("Slide pipeline failed: no pages generated")) {
      throw new Error("Slide pipeline failed: no pages generated");
    }
    if (debugMode) {
      throw new Error(`当前课件较复杂，建议先导出为 PDF 再上传。调试信息：${reason}`);
    }
    throw new Error("当前课件较复杂，建议先导出为 PDF 再上传。");
  }

}

