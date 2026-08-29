import { CommandType, ImageSourceType, ObjectRelativeFromH, ObjectRelativeFromV } from "@univerjs/core";
import type { ICommand } from "@univerjs/core";
import { TextWrappingStyle } from "@univerjs/docs-drawing";
import type { FDocument } from "@univerjs/docs/facade";

// Word's watermark is not a page property — it is a picture sitting in the
// page header, behind the text, rotated up the page. That is why it repeats
// on every page and why it never gets in the way of typing, and it is
// exactly how this builds one: the text is drawn to a canvas, and the
// resulting image goes into the header as a behind-text drawing.

export const SetWatermarkCommandId = "dockaro.command.watermark";

/** Word tilts its text watermarks up the page at this angle. */
const WATERMARK_ANGLE = -45;
const WATERMARK_COLOR = "#d0d0d0";
const WATERMARK_FONT_PX = 130;

export interface ISetWatermarkParams {
  /** null removes the watermark, as Word's "No watermark" does. */
  text: string | null;
}

/** Draws the watermark text to a transparent PNG sized to fit it. */
function renderWatermark(text: string): { source: string; width: number; height: number } | null {
  const canvas = document.createElement("canvas");
  const measuring = canvas.getContext("2d");
  if (!measuring) return null;

  const font = `bold ${WATERMARK_FONT_PX}px Arial, sans-serif`;
  measuring.font = font;
  const width = Math.ceil(measuring.measureText(text).width) + 40;
  const height = Math.ceil(WATERMARK_FONT_PX * 1.3);

  canvas.width = width;
  canvas.height = height;
  const ctx = canvas.getContext("2d");
  if (!ctx) return null;
  ctx.font = font;
  ctx.fillStyle = WATERMARK_COLOR;
  ctx.textAlign = "center";
  ctx.textBaseline = "middle";
  ctx.fillText(text, width / 2, height / 2);

  return { source: canvas.toDataURL("image/png"), width, height };
}

export function createWatermarkCommand(doc: FDocument): ICommand<ISetWatermarkParams> {
  return {
    id: SetWatermarkCommandId,
    type: CommandType.COMMAND,
    handler: async (_accessor, params) => {
      if (!params) return false;

      // Whatever happens next, the old watermark goes. It is recognised by
      // its rotation, which nothing else in a document shares.
      for (const image of doc.getImages()) {
        if (image.getAngle() === WATERMARK_ANGLE) image.remove();
      }
      if (!params.text) return true;

      const rendered = renderWatermark(params.text);
      if (!rendered) return false;

      const style = doc.getDocumentDataModel().getDocumentStyle();
      const pageWidth = style.pageSize?.width ?? 794;
      const pageHeight = style.pageSize?.height ?? 1123;
      // The rotation happens around the image's centre, so centring the
      // upright box centres the tilted text too — but a box tilted 45
      // degrees is wider than it stands, so it has to be scaled against its
      // rotated extent, 0.707 * (width + height), or it hangs off the page.
      const rotatedExtent = Math.SQRT1_2 * (rendered.width + rendered.height);
      const scale = Math.min(1, (pageWidth * 0.92) / rotatedExtent);
      const width = Math.round(rendered.width * scale);
      const height = Math.round(rendered.height * scale);

      const segmentId = doc.ensurePageHeader();
      const inserted = await doc.insertImage({
        source: rendered.source,
        imageSourceType: ImageSourceType.BASE64,
        width,
        height,
        angle: WATERMARK_ANGLE,
        wrappingStyle: TextWrappingStyle.BEHIND_TEXT,
        positionH: { relativeFrom: ObjectRelativeFromH.PAGE, posOffset: Math.round((pageWidth - width) / 2) },
        positionV: { relativeFrom: ObjectRelativeFromV.PAGE, posOffset: Math.round((pageHeight - height) / 2) },
        textRange: { startOffset: 0, endOffset: 0, collapsed: true, segmentId },
      });
      return inserted != null;
    },
  };
}
