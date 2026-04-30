import { PDFDocument } from "pdf-lib";
import { buildSheetLayouts, getPaperPt } from "./layout";
import { mmToPt } from "./constants";
import type { LayoutSettings, PagePipelineResult } from "./types";

export async function exportImposedPdf(data: PagePipelineResult, settings: LayoutSettings): Promise<Uint8Array> {
  try {
    const outDoc = await PDFDocument.create();
    const sourceDoc = data.sourcePdfBytes ? await PDFDocument.load(Uint8Array.from(data.sourcePdfBytes)) : null;
    const embeddedImageCache = new Map<string, Awaited<ReturnType<typeof outDoc.embedPng>>>();
    const layouts = buildSheetLayouts(data.pages, settings);
    const paper = getPaperPt(settings);
    const outerMarginPt = mmToPt(8);
    const gapPt = mmToPt(settings.gapMm);
    const cellWidth = (paper.width - outerMarginPt * 2 - gapPt * (settings.columns - 1)) / settings.columns;
    const cellHeight = (paper.height - outerMarginPt * 2 - gapPt * (settings.rows - 1)) / settings.rows;

    for (const sheet of layouts) {
      const page = outDoc.addPage([paper.width, paper.height]);
      for (let slot = 0; slot < sheet.slots.length; slot += 1) {
        const item = sheet.slots[slot];
        if (!item) continue;

        const row = Math.floor(slot / settings.columns);
        const col = slot % settings.columns;

        const x = outerMarginPt + col * (cellWidth + gapPt);
        const yTop = outerMarginPt + row * (cellHeight + gapPt);
        const y = paper.height - yTop - cellHeight;

        const canUseSourcePdf =
          sourceDoc &&
          typeof item.page.pdfPageIndex === "number" &&
          item.page.pdfPageIndex >= 0 &&
          item.page.pdfPageIndex < sourceDoc.getPageCount();

        if (canUseSourcePdf) {
          const sourcePage = sourceDoc.getPage(item.page.pdfPageIndex!);
          const embeddedPage = await outDoc.embedPage(sourcePage);

          // Compare pdfjs visual orientation (rotation-corrected) vs raw MediaBox orientation.
          // If they differ, the page has CTM+/Rotate transforms that embedPage cannot replicate;
          // fall through to the PNG path which already has the correct rendering.
          const pdfjsLandscape = item.page.width > item.page.height;
          const mediaBoxLandscape = embeddedPage.width > embeddedPage.height;

          if (pdfjsLandscape === mediaBoxLandscape) {
            const scale = Math.min(cellWidth / embeddedPage.width, cellHeight / embeddedPage.height);
            const drawW  = embeddedPage.width  * scale;
            const drawH  = embeddedPage.height * scale;
            const drawX  = x + (cellWidth  - drawW) / 2;
            const drawY  = y + (cellHeight - drawH) / 2;
            page.drawPage(embeddedPage, { x: drawX, y: drawY, width: drawW, height: drawH });
            continue;
          }
          // Orientation mismatch: fall through to PNG path
        }

        let embeddedImage = embeddedImageCache.get(item.page.image);
        if (!embeddedImage) {
          const imageBytes = await fetch(item.page.image).then((res) => res.arrayBuffer());
          embeddedImage = await outDoc.embedPng(imageBytes);
          embeddedImageCache.set(item.page.image, embeddedImage);
        }
        const imageWidth = item.page.width || embeddedImage.width;
        const imageHeight = item.page.height || embeddedImage.height;
        const scale = Math.min(cellWidth / imageWidth, cellHeight / imageHeight);
        const drawWidth = imageWidth * scale;
        const drawHeight = imageHeight * scale;
        const drawX = x + (cellWidth - drawWidth) / 2;
        const drawY = y + (cellHeight - drawHeight) / 2;

        page.drawImage(embeddedImage, {
          x: drawX,
          y: drawY,
          width: drawWidth,
          height: drawHeight
        });
      }
    }

    return outDoc.save();
  } catch (error) {
    const msg = error instanceof Error ? error.message : "未知导出错误";
    throw new Error(`导出 PDF 失败：${msg}`);
  }
}
