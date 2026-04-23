"use client";

import { useMemo, useState } from "react";
import { ControlPanel } from "@/components/ControlPanel";
import { PreviewPane } from "@/components/PreviewPane";
import { ENCOURAGEMENTS } from "@/lib/constants";
import { exportImposedPdf } from "@/lib/exportPdf";
import { buildPagePipeline } from "@/lib/pipeline";
import type { LayoutSettings, PagePipelineResult } from "@/lib/types";

const initialSettings: LayoutSettings = {
  rows: 2,
  columns: 2,
  gapMm: 4,
  paperSize: "A4",
  orientation: "portrait",
  pattern: "Z"
};

function downloadPdf(blob: Blob, filename: string): void {
  const url = URL.createObjectURL(blob);
  const anchor = document.createElement("a");
  anchor.href = url;
  anchor.download = filename;
  anchor.style.display = "none";
  document.body.appendChild(anchor);
  anchor.click();
  document.body.removeChild(anchor);

  // 延迟释放，避免部分浏览器在点击瞬间提前失效
  window.setTimeout(() => {
    URL.revokeObjectURL(url);
  }, 1500);
}

export default function HomePage() {
  const [settings, setSettings] = useState<LayoutSettings>(initialSettings);
  const [pipeline, setPipeline] = useState<PagePipelineResult | null>(null);
  const [loading, setLoading] = useState(false);
  const [message, setMessage] = useState<string | null>(null);
  const [statusText, setStatusText] = useState<string | null>(null);
  const [retryFile, setRetryFile] = useState<File | null>(null);
  const [canRetry, setCanRetry] = useState(false);
  const [debugMode, setDebugMode] = useState(false);
  const pages = useMemo(() => pipeline?.pages ?? [], [pipeline]);

  const handleFileChange = async (file: File | null) => {
    if (!file) return;
    const lower = file.name.toLowerCase();
    const isPpt = lower.endsWith(".ppt") || lower.endsWith(".pptx");
    setLoading(true);
    setRetryFile(file);
    setCanRetry(false);
    setMessage(null);
    setStatusText(isPpt ? "正在生成幻灯片预览..." : "正在生成页面预览...");
    try {
      const data = await buildPagePipeline(file, {
        snapshotScale: 3,
        debugMode,
        onProgress: (progress) => {
          if (progress.total && progress.current) {
            setStatusText(`${progress.message}（${progress.current}/${progress.total}）`);
            return;
          }
          setStatusText(progress.message);
        }
      });
      setPipeline(data);
      setMessage(`已加载 ${data.pages.length} 页，进入实时拼版模式。`);
      if (debugMode && data.diagnostics) {
        setMessage(
          `已加载 ${data.pages.length} 页。诊断：parsed=${data.diagnostics.slidesParsedCount}, rendered=${data.diagnostics.slidesRenderedCount}, snapshots=${data.diagnostics.snapshotsGeneratedCount}, blank=${data.diagnostics.blankSnapshotsDetected}, normalized=${data.diagnostics.pagesNormalizedCount ?? 0}, renderer=${data.diagnostics.rendererUsed ?? "primary"}`
        );
      }
    } catch (error) {
      const text = error instanceof Error ? error.message : "文件处理失败";
      setMessage(text);
      setPipeline(null);
      setCanRetry(text.includes("建议先导出为 PDF 再上传"));
    } finally {
      setLoading(false);
      setStatusText(null);
    }
  };

  const handleGenerate = async () => {
    if (!pipeline) return;
    setLoading(true);
    try {
      const bytes = await exportImposedPdf(pipeline, settings);
      const safeBytes = Uint8Array.from(bytes);
      const blob = new Blob([safeBytes], { type: "application/pdf" });
      const filename = `${pipeline.sourceName.replace(/\.[^.]+$/, "")}-study-layout.pdf`;
      downloadPdf(blob, filename);
      const random = ENCOURAGEMENTS[Math.floor(Math.random() * ENCOURAGEMENTS.length)];
      if (debugMode && pipeline.diagnostics) {
        setMessage(
          `${random}（export images=${pipeline.pages.length}, parsed=${pipeline.diagnostics.slidesParsedCount}, rendered=${pipeline.diagnostics.slidesRenderedCount}, snapshots=${pipeline.diagnostics.snapshotsGeneratedCount}, blank=${pipeline.diagnostics.blankSnapshotsDetected}）`
        );
      } else {
        setMessage(random);
      }
    } catch (error) {
      const text = error instanceof Error ? error.message : "导出失败";
      setMessage(text);
    } finally {
      setLoading(false);
    }
  };

  return (
    <main className="min-h-screen p-4 md:p-6">
      <header className="mb-4">
        <h1 className="text-2xl font-bold text-slate-900">Study Imposition Engine</h1>
        <p className="text-sm text-slate-600">仅做页面拼版，不修改单页内容。支持 PDF 与 PPT/PPTX 上传。</p>
      </header>

      <div className="flex flex-col gap-4 lg:flex-row">
        <ControlPanel
          settings={settings}
          loading={loading}
          sourceName={pipeline?.sourceName ?? null}
          statusText={statusText}
          debugMode={debugMode}
          onFileChange={handleFileChange}
          onSettingsChange={setSettings}
          onDebugModeChange={setDebugMode}
          onGenerate={handleGenerate}
          onRetry={canRetry && retryFile ? () => handleFileChange(retryFile) : undefined}
        />
        <PreviewPane pages={pages} settings={settings} />
      </div>

      {debugMode && pipeline?.diagnostics ? (
        <section className="mt-4 rounded-lg border border-amber-200 bg-amber-50 p-4 text-sm text-amber-900">
          <div className="mb-2 font-semibold">Consistency Check</div>
          <div className="grid grid-cols-1 gap-1 md:grid-cols-2">
            <div>Status: {pipeline.diagnostics.consistencyStatus === "pass" ? "Pass" : "Warning"}</div>
            <div>Slide Count: {pipeline.diagnostics.slidesParsedCount}</div>
            <div>
              Canonical Size:{" "}
              {pipeline.diagnostics.canonicalSize
                ? `${pipeline.diagnostics.canonicalSize.width} x ${pipeline.diagnostics.canonicalSize.height}`
                : "-"}
            </div>
            <div>Normalized Pages: {pipeline.diagnostics.pagesNormalizedCount ?? 0}</div>
            <div>Detected Sizes: {(pipeline.diagnostics.uniqueCanvasSizes ?? []).join(", ") || "-"}</div>
            <div>Aspect Ratios: {(pipeline.diagnostics.aspectRatios ?? []).join(", ") || "-"}</div>
            <div>Stability Runs: {pipeline.diagnostics.stabilityRuns ?? 0}</div>
            <div>Stability Mismatches: {pipeline.diagnostics.stabilityMismatches ?? 0}</div>
            <div>Over Capture Detected: {pipeline.diagnostics.overCaptureDetected ? "Yes" : "No"}</div>
            <div>BBox Extension: {pipeline.diagnostics.bboxExtensionAmount ?? 0}</div>
            <div>Text Overflow Detected: {pipeline.diagnostics.textOverflowDetected ? "Yes" : "No"}</div>
            <div>
              Bottom Capture:{" "}
              {(pipeline.diagnostics.bottomClipChecks ?? [])
                .slice(0, 3)
                .map(
                  (item) =>
                    `#${item.slideIndex} h=${item.slideHeight} b=${item.contentBBoxBottom} cap=${item.finalCaptureHeight} bleed=${item.bleedAdded} exceed=${item.contentExceedsFrame ? "Y" : "N"}`
                )
                .join(" | ") || "-"}
            </div>
          </div>
        </section>
      ) : null}

      {message ? (
        <div className="mt-4 rounded-lg border border-indigo-200 bg-indigo-50 px-4 py-3 text-sm text-indigo-800">
          {message}
        </div>
      ) : null}
    </main>
  );
}
