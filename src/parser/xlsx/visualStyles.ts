import type { CSSProperties } from "react";
import type { ThemeData } from "../../model/types";
import { attr, attrBool, attrNum, child, children, type XmlNode } from "../xml";
import { rgbaToHex } from "../drawingml/color";

const THEME_ORDER = ["lt1", "dk1", "lt2", "dk2", "accent1", "accent2", "accent3", "accent4", "accent5", "accent6", "hlink", "folHlink"];
const INDEXED = ["000000", "FFFFFF", "FF0000", "00FF00", "0000FF", "FFFF00", "FF00FF", "00FFFF",
  "000000", "FFFFFF", "FF0000", "00FF00", "0000FF", "FFFF00", "FF00FF", "00FFFF",
  "800000", "008000", "000080", "808000", "800080", "008080", "C0C0C0", "808080",
  "9999FF", "993366", "FFFFCC", "CCFFFF", "660066", "FF8080", "0066CC", "CCCCFF",
  "000080", "FF00FF", "FFFF00", "00FFFF", "800080", "800000", "008080", "0000FF",
  "00CCFF", "CCFFFF", "CCFFCC", "FFFF99", "99CCFF", "FF99CC", "CC99FF", "FFCC99",
  "3366FF", "33CCCC", "99CC00", "FFCC00", "FF9900", "FF6600", "666699", "969696",
  "003366", "339966", "003300", "333300", "993300", "993366", "333399", "333333"];

export function sheetColor(node: XmlNode | null, theme: ThemeData): string | undefined {
  if (!node) return undefined;
  let hex = attr(node, "rgb")?.slice(-6);
  const slot = THEME_ORDER[attrNum(node, "theme") ?? -1];
  if (!hex && slot && theme.colors[slot]) hex = rgbaToHex(theme.colors[slot]!).replace("#", "");
  if (!hex) hex = INDEXED[attrNum(node, "indexed") ?? -1];
  if (!hex || !/^[\da-f]{6}$/i.test(hex)) return undefined;
  const tint = Math.max(-1, Math.min(1, attrNum(node, "tint") ?? 0));
  const channels = [0, 2, 4].map((offset) => {
    const value = parseInt(hex!.slice(offset, offset + 2), 16);
    return Math.round(tint < 0 ? value * (1 + tint) : value * (1 - tint) + 255 * tint);
  });
  return `rgb(${channels.join(", ")})`;
}

export function parseVisualStyles(root: XmlNode | null, theme: ThemeData): CSSProperties[] {
  const fonts = children(child(root, "fonts"), "font");
  const fills = children(child(root, "fills"), "fill");
  const borders = children(child(root, "borders"), "border");
  const bases = children(child(root, "cellStyleXfs"), "xf");
  return children(child(root, "cellXfs"), "xf").map((xf) => {
    const base = bases[attrNum(xf, "xfId") ?? 0];
    const resource = (name: string) => attrNum(xf, name) ?? attrNum(base, name) ?? 0;
    const font = fonts[resource("fontId")] ?? null;
    const pattern = child(fills[resource("fillId")], "patternFill");
    const border = borders[resource("borderId")];
    const alignment = child(xf, "alignment") ?? child(base, "alignment");
    const size = attrNum(child(font, "sz"), "val");
    const name = attr(child(font, "name"), "val") ?? theme.minorFont ?? "Calibri";
    const style: CSSProperties = {
      fontFamily: `"${name.replace(/["\\]/g, "")}", Arial, sans-serif`,
      fontSize: size === null ? undefined : `${Math.min(Math.max(size, 1), 409) * 96 / 72}px`,
      fontWeight: child(font, "b") && attrBool(child(font, "b"), "val") !== false ? 700 : 400,
      fontStyle: child(font, "i") && attrBool(child(font, "i"), "val") !== false ? "italic" : "normal",
      color: sheetColor(child(font, "color"), theme),
      backgroundColor: attr(pattern, "patternType") === "solid" ? sheetColor(child(pattern, "fgColor"), theme) : undefined,
      verticalAlign: attr(alignment, "vertical") === "center" ? "middle" : attr(alignment, "vertical") === "top" ? "top" : "bottom",
      whiteSpace: attrBool(alignment, "wrapText") ? "pre-wrap" : "pre",
    };
    const horizontal = attr(alignment, "horizontal");
    if (horizontal === "left" || horizontal === "center" || horizontal === "right" || horizontal === "justify") style.textAlign = horizontal;
    const decorations = [child(font, "u") && attr(child(font, "u"), "val") !== "none" ? "underline" : "", child(font, "strike") && attrBool(child(font, "strike"), "val") !== false ? "line-through" : ""].filter(Boolean);
    if (decorations.length) style.textDecoration = decorations.join(" ");
    const indent = attrNum(alignment, "indent");
    if (indent) style.paddingLeft = `${Math.min(indent, 250) * 9 + 4}px`;
    for (const [edge, css] of [["left", "borderLeft"], ["right", "borderRight"], ["top", "borderTop"], ["bottom", "borderBottom"]] as const) {
      const node = child(border, edge);
      const kind = attr(node, "style");
      if (!kind || kind === "none") continue;
      const width = kind === "thick" ? 3 : kind.startsWith("medium") || kind === "double" ? 2 : 1;
      const line = kind === "double" ? "double" : kind.toLowerCase().includes("dash") ? "dashed" : kind === "dotted" || kind === "hair" ? "dotted" : "solid";
      style[css] = `${width}px ${line} ${sheetColor(child(node, "color"), theme) ?? "#000"}`;
    }
    return style;
  });
}
