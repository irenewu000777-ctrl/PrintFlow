import { PDFDocument, degrees } from "pdf-lib";
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
          const rotation = sourcePage.getRotation().angle; // 0 | 90 | 180 | 270
          const embeddedPage = await outDoc.embedPage(sourcePage);

          // embedPage uses raw MediaBox; swap dims for 90°/270° pages
          const isSwapped = rotation === 90 || rotation === 270;
          const intrinsicW = isSwapped ? embeddedPage.height : embeddedPage.width;
          const intrinsicH = isSwapped ? embeddedPage.width  : embeddedPage.height;

          const scale   = Math.min(cellWidth / intrinsicW, cellHeight / intrinsicH);
          const scaledW = embeddedPage.width  * scale;
          const scaledH = embeddedPage.height * scale;
          const visualW = isSwapped ? scaledH : scaledW;
          const visualH = isSwapped ? scaledW : scaledH;

          const drawX = x + (cellWidth  - visualW) / 2;
          const drawY = y + (cellHeight - visualH) / 2;

          // Adjust anchor so visual bottom-left lands on (drawX, drawY) after CCW rotation
          let embedX: number, embedY: number;
          if (rotation === 90)       { embedX = drawX + scaledH; embedY = drawY;           }
          else if (rotation === 180) { embedX = drawX + scaledW; embedY = drawY + scaledH; }
          else if (rotation === 270) { embedX = drawX;           embedY = drawY + scaledW; }
          else                       { embedX = drawX;           embedY = drawY;           }

          page.drawPage(embeddedPage, {
            x: embedX, y: embedY,
            width: scaledW, height: scaledH,
            rotate: degrees(rotation)
          });
          continue;
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
