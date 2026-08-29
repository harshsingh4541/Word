import { CommandType, DashStyleType } from "@univerjs/core";
import type { ICommand } from "@univerjs/core";

// Word's Borders group works with a "pen": you choose a line style, a
// weight and a colour, and the Borders button then draws with them. The pen
// itself is UI state, not document state, so setting it is an operation with
// nothing to undo.

export interface BorderPen {
  /** Line width in document pixels. */
  width: number;
  color: string;
  dashStyle: DashStyleType;
}

const pen: BorderPen = {
  width: 1,
  color: "#000000",
  dashStyle: DashStyleType.SOLID,
};

export function getBorderPen(): BorderPen {
  return pen;
}

/**
 * Arm the pen and read it back in one step, for the colour and thickness
 * entries that live inside the Borders dropdown: those have to both remember
 * the choice (so the separate Pen color / Line weight buttons stay in sync)
 * and draw with it immediately, but a dropdown entry only gets to dispatch
 * one command. Setting the pen here lets that single dispatch be the draw.
 */
export function armBorderPen(patch: ISetBorderPenParams): BorderPen {
  if (typeof patch.width === "number") pen.width = patch.width;
  if (typeof patch.color === "string") pen.color = patch.color;
  if (typeof patch.dashStyle === "number") pen.dashStyle = patch.dashStyle;
  return pen;
}

export interface ISetBorderPenParams {
  width?: number;
  color?: string;
  dashStyle?: DashStyleType;
}

export const SetBorderPenCommandId = "dockaro.command.border-pen";

export const SetBorderPenCommand: ICommand<ISetBorderPenParams> = {
  id: SetBorderPenCommandId,
  type: CommandType.OPERATION,
  handler: (_accessor, params) => {
    if (!params) return false;
    armBorderPen(params);
    return true;
  },
};

/** Word's Line Weight list, in points. */
export const BORDER_WEIGHTS = [0.5, 1, 1.5, 2.25, 3, 4.5, 6];

/** Points to the document pixels the border width is stored in. */
export function pointsToPixels(points: number): number {
  return (points * 96) / 72;
}
