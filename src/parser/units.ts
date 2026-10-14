/** Unit conversions between OOXML units and CSS pixels. */

export const EMU_PER_INCH = 914_400;
export const EMU_PER_POINT = 12_700;
export const EMU_PER_CM = 360_000;
/* --------------------------------------------------------- WordprocessingML */

/** Twips (1/1440in) to EMU. */
export function twipsToEmu(value: number | null | undefined): number {
  if (value === null || value === undefined || !Number.isFinite(value)) return 0;
  return Math.round(value * EMU_PER_TWIP);
}

/** Twips to CSS pixels at 96dpi (1440 twips === 96px, so this is value / 15). */
export function twipsToPx(value: number | null | undefined): number {
  if (value === null || value === undefined || !Number.isFinite(value)) return 0;
  return (value * EMU_PER_TWIP) / EMU_PER_PX;
}

/**
 * `w:sz` is in half-points. Font sizes in the model are CSS pixels, matching the
 * DrawingML parser, so both formats produce comparable text.
 */
export function halfPtToPx(value: number | null | undefined): number {
  if (value === null || value === undefined || !Number.isFinite(value)) return 0;
  return ptToPx(value / 2);
}

/** OOXML stores slide geometry in EMU; we render at a 96dpi CSS pixel grid. */
export const EMU_PER_PX = EMU_PER_INCH / 96; // 9525
/** WordprocessingML measures page geometry and indents in twips (1/1440in). */
export const EMU_PER_TWIP = 635; // 914400 / 1440

export function emuToPx(emu: number | null | undefined): number {
  if (emu === null || emu === undefined || !Number.isFinite(emu)) return 0;
  return emu / EMU_PER_PX;
}

export function pxToEmu(px: number): number {
  return Math.round(px * EMU_PER_PX);
}

/** Hundredths of a point (the unit used by `a:rPr/@sz`) to CSS pixels. */
export function hundredthsPtToPx(value: number): number {
  return (value / 100) * (96 / 72);
}

/** Points to CSS pixels. */
export function ptToPx(value: number): number {
  return value * (96 / 72);
}

/** OOXML rotation units (1/60000 of a degree) to degrees. */
export function angleUnitsToDegrees(value: number | null | undefined): number {
  if (!value) return 0;
  return value / 60_000;
}

/** DrawingML "ST_Percentage" (100000 === 100%) to a ratio. */
export function percentageToRatio(value: number | null | undefined): number {
  if (value === null || value === undefined || !Number.isFinite(value)) return 0;
  return value / 100_000;
}

/** Clamp helper used all over the drawing-math code. */
export function clamp(value: number, min: number, max: number): number {
  if (Number.isNaN(value)) return min;
  return value < min ? min : value > max ? max : value;
}

export function round(value: number, digits = 2): number {
  const factor = 10 ** digits;
  return Math.round(value * factor) / factor;
}

/** Rounded px value with a stable string form, keeps DOM attributes compact. */
export function px(value: number): number {
  if (!Number.isFinite(value)) return 0;
  return Math.round(value * 100) / 100;
}
