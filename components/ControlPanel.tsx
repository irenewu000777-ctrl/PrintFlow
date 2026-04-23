"use client";

import type { ChangeEvent } from "react";
import { MAX_PDF_FILE_SIZE_MB, PAPER_SIZE_OPTIONS } from "@/lib/constants";
import type { LayoutSettings } from "@/lib/types";

interface ControlPanelProps {
  settings: LayoutSettings;
  loading: boolean;
  sourceName: string | null;
  statusText?: string | null;
  debugMode?: boolean;
  onFileChange: (file: File | null) => void;
  onSettingsChange: (next: LayoutSettings) => void;
  onDebugModeChange?: (enabled: boolean) => void;
  onGenerate: () => void;
  onRetry?: () => void;
}

export function ControlPanel({
  settings,
  loading,
  sourceName,
  statusText,
  debugMode,
  onFileChange,
  onSettingsChange,
  onDebugModeChange,
  onGenerate,
  onRetry
}: ControlPanelProps) {
  const handleUpload = (e: ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0] ?? null;
    onFileChange(file);
  };

  return (
    <aside className="w-full rounded-xl border border-slate-200 bg-white p-4 shadow-sm lg:w-[360px]">
      <h2 className="mb-4 text-lg font-semibold">排版控制台</h2>
      <div className="space-y-4 text-sm">
        <label className="block">
          <span className="mb-1 block font-medium">上传文件（PDF / PPT / PPTX）</span>
          <input
            type="file"
            accept=".pdf,.ppt,.pptx,application/pdf,application/vnd.ms-powerpoint,application/vnd.openxmlformats-officedocument.presentationml.presentation"
            onChange={handleUpload}
            className="block w-full rounded border border-slate-300 px-2 py-2"
          />
          <p className="mt-1 text-slate-500">最大文件限额：{MAX_PDF_FILE_SIZE_MB}MB</p>
          {sourceName ? <p className="mt-1 text-slate-500">当前文件：{sourceName}</p> : null}
          {statusText ? <p className="mt-1 text-indigo-600">{statusText}</p> : null}
        </label>

        <div className="grid grid-cols-2 gap-3">
          <label className="block">
            <span className="mb-1 block font-medium">纸张尺寸</span>
            <select
              value={settings.paperSize}
              onChange={(e) => onSettingsChange({ ...settings, paperSize: e.target.value as LayoutSettings["paperSize"] })}
              className="w-full rounded border border-slate-300 px-2 py-2"
            >
              {PAPER_SIZE_OPTIONS.map((item) => (
                <option key={item} value={item}>
                  {item}
                </option>
              ))}
            </select>
          </label>
          <label className="block">
            <span className="mb-1 block font-medium">纸张方向</span>
            <select
              value={settings.orientation}
              onChange={(e) =>
                onSettingsChange({ ...settings, orientation: e.target.value as LayoutSettings["orientation"] })
              }
              className="w-full rounded border border-slate-300 px-2 py-2"
            >
              <option value="portrait">Portrait</option>
              <option value="landscape">Landscape</option>
            </select>
          </label>
        </div>

        <div className="grid grid-cols-2 gap-3">
          <label className="block">
            <span className="mb-1 block font-medium">Rows</span>
            <select
              value={settings.rows}
              onChange={(e) => onSettingsChange({ ...settings, rows: Number(e.target.value) })}
              className="w-full rounded border border-slate-300 px-2 py-2"
            >
              {[1, 2, 3, 4].map((n) => (
                <option key={n} value={n}>
                  {n}
                </option>
              ))}
            </select>
          </label>
          <label className="block">
            <span className="mb-1 block font-medium">Columns</span>
            <select
              value={settings.columns}
              onChange={(e) => onSettingsChange({ ...settings, columns: Number(e.target.value) })}
              className="w-full rounded border border-slate-300 px-2 py-2"
            >
              {[1, 2, 3, 4].map((n) => (
                <option key={n} value={n}>
                  {n}
                </option>
              ))}
            </select>
          </label>
        </div>

        <label className="block">
          <span className="mb-1 block font-medium">排列方式</span>
          <select
            value={settings.pattern}
            onChange={(e) => onSettingsChange({ ...settings, pattern: e.target.value as LayoutSettings["pattern"] })}
            className="w-full rounded border border-slate-300 px-2 py-2"
          >
            <option value="Z">Z-pattern（横向优先）</option>
            <option value="N">N-pattern（纵向优先）</option>
          </select>
        </label>

        <label className="block">
          <span className="mb-1 block font-medium">间距（mm）：{settings.gapMm}</span>
          <input
            type="range"
            min={0}
            max={20}
            step={1}
            value={settings.gapMm}
            onChange={(e) => onSettingsChange({ ...settings, gapMm: Number(e.target.value) })}
            className="w-full"
          />
        </label>
        <label className="flex items-center gap-2">
          <input
            type="checkbox"
            checked={Boolean(debugMode)}
            onChange={(e) => onDebugModeChange?.(e.target.checked)}
          />
          <span className="text-slate-600">调试模式（输出诊断信息）</span>
        </label>

        <button
          onClick={onGenerate}
          disabled={loading || !sourceName}
          className="w-full rounded bg-indigo-600 px-4 py-2 font-medium text-white disabled:cursor-not-allowed disabled:bg-slate-300"
        >
          {loading ? "处理中..." : "Generate Study PDF"}
        </button>
        {onRetry ? (
          <button
            onClick={onRetry}
            disabled={loading}
            className="w-full rounded border border-slate-300 px-4 py-2 font-medium text-slate-700 disabled:cursor-not-allowed disabled:opacity-60"
          >
            Retry
          </button>
        ) : null}
      </div>
    </aside>
  );
}
