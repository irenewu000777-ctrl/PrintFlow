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
  const pages = useMemo(() => pipeline?.pages ?? [], [pipeline]);

  const handleFileChange = async (file: File | null) => {
    if (!file) return;
    setLoading(true);
    setMessage(null);
    try {
      const data = await buildPagePipeline(file);
      setPipeline(data);
      setMessage(`已加载 ${data.pages.length} 页，进入实时拼版模式。`);
    } catch (error) {
      const text = error instanceof Error ? error.message : "文件处理失败";
      setMessage(text);
      setPipeline(null);
    } finally {
      setLoading(false);
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
      setMessage(random);
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
        <p className="text-sm text-slate-600">仅做页面拼版，不修改单页内容。当前版本仅支持 PDF 上传。</p>
      </header>

      <div className="flex flex-col gap-4 lg:flex-row">
        <ControlPanel
          settings={settings}
          loading={loading}
          sourceName={pipeline?.sourceName ?? null}
          onFileChange={handleFileChange}
          onSettingsChange={setSettings}
          onGenerate={handleGenerate}
        />
        <PreviewPane pages={pages} settings={settings} />
      </div>

      {message ? (
        <div className="mt-4 rounded-lg border border-indigo-200 bg-indigo-50 px-4 py-3 text-sm text-indigo-800">
          {message}
        </div>
      ) : null}
    </main>
  );
}
