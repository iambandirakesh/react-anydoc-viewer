import type { ThemeData } from "../../model/types";
import type { SheetPageData } from "./model";
import { attr, attrBool, attrNum, child, children, type XmlNode } from "../xml";
import { sheetColor } from "./visualStyles";

/** Apply bounded, cached-value visual rules; formulas are never evaluated. */
export function applyConditionalFormatting(root: XmlNode, data: SheetPageData, theme: ThemeData, warn: (code: string, message: string) => void, stylesRoot?: XmlNode | null): void {
  const ref = (text: string) => {
    const match = /^\$?([A-Z]+)\$?(\d+)$/i.exec(text);
    if (!match) return null;
    let col = 0;
    for (const letter of match[1]!.toUpperCase()) col = col * 26 + letter.charCodeAt(0) - 64;
    return { row: Number(match[2]) - 1, col: col - 1 };
  };
  const values = new Map<string, number>();
  for (const row of children(child(root, "sheetData"), "row")) for (const cell of children(row, "c")) {
    if (attr(cell, "t") && attr(cell, "t") !== "n") continue;
    const text = child(cell, "v")?.text;
    if (text?.trim() && Number.isFinite(Number(text))) values.set((attr(cell, "r") ?? "").toUpperCase(), Number(text));
  }
  const blocked = new WeakSet<object>();
  const applied = new WeakMap<object, Set<string>>();
  const applyStyle = (cell: SheetPageData["rows"][number][number], style: Record<string, unknown>) => {
    const keys = applied.get(cell) ?? new Set<string>();
    const next = { ...cell.style };
    for (const [key, value] of Object.entries(style)) if (value !== undefined && !keys.has(key)) { (next as Record<string, unknown>)[key] = value; keys.add(key); }
    cell.style = next; applied.set(cell, keys);
  };
  let budget = 200;
  const groups = children(root, "conditionalFormatting");
  if (groups.length > 200) warn("xlsx.conditional-limit", "Only the first 200 conditional formatting groups are applied.");
  const rules = groups.slice(0, 200).flatMap((group) => children(group, "cfRule").slice(0, 200).map((rule) => ({ group, rule }))).sort((a, b) => (attrNum(a.rule, "priority") ?? Infinity) - (attrNum(b.rule, "priority") ?? Infinity));
  if (rules.length > 200) warn("xlsx.conditional-limit", "Only the first 200 conditional formatting rules are applied.");
  for (const { group, rule: currentRule } of rules.slice(0, 200)) {
    const ranges = (attr(group, "sqref") ?? "").split(/\s+/).slice(0, 200).map((range) => {
      const [start, end] = range.split(":"); return { start: ref(start ?? ""), end: ref(end ?? start ?? "") };
    });
    const targets: { cell: SheetPageData["rows"][number][number]; value: number }[] = [];
    data.rows.forEach((row, r) => row.forEach((cell, c) => {
      if (!ranges.some(({ start, end }) => start && end && r >= start.row && r <= end.row && c >= start.col && c <= end.col)) return;
      let n = c + 1, label = "";
      while (n) { n--; label = String.fromCharCode(65 + n % 26) + label; n = Math.floor(n / 26); }
      const value = values.get(`${label}${r + 1}`);
      if (value !== undefined) targets.push({ cell, value });
    }));
    const sorted = targets.map(({ value }) => value).sort((a, b) => a - b);
    if (!sorted.length) continue;
    const min = sorted[0]!, max = sorted[sorted.length - 1]!;
    const threshold = (node: XmlNode) => {
      const type = attr(node, "type"), val = attrNum(node, "val") ?? 0;
      if (type === "min") return min;
      if (type === "max") return max;
      if (type === "num") return val;
      if (type === "percent") return min + (max - min) * val / 100;
      if (type === "percentile") {
        const index = Math.max(0, Math.min(100, val)) / 100 * (sorted.length - 1), low = Math.floor(index);
        return sorted[low]! + (sorted[Math.ceil(index)]! - sorted[low]!) * (index - low);
      }
      return NaN;
    };
    for (const rule of [currentRule]) {
      if (--budget < 0) { warn("xlsx.conditional-limit", "Only the first 200 conditional formatting rules are applied."); return; }
      const type = attr(rule, "type");
      if (type === "cellIs") {
        const formulas = children(rule, "formula").map((item) => Number(item.text.trim()));
        const operator = attr(rule, "operator");
        const dxf = children(child(stylesRoot, "dxfs"), "dxf")[attrNum(rule, "dxfId") ?? -1];
        if (!dxf || !formulas.length || formulas.some((value) => !Number.isFinite(value)) || !children(rule, "formula").every((item) => /^[+-]?(?:\d+\.?\d*|\.\d+)(?:[eE][+-]?\d+)?$/.test(item.text.trim()))) { warn("xlsx.conditionalFormatting", "Comparison rules require numeric constant thresholds and a differential style."); continue; }
        const font = child(dxf, "font"), fill = child(child(dxf, "fill"), "patternFill");
        const style = { color: sheetColor(child(font, "color"), theme), backgroundColor: sheetColor(child(fill, "fgColor"), theme), fontWeight: child(font, "b") ? attrBool(child(font, "b"), "val") === false ? 400 : 700 : undefined, fontStyle: child(font, "i") ? attrBool(child(font, "i"), "val") === false ? "normal" : "italic" : undefined };
        const supported = ["lessThan", "lessThanOrEqual", "greaterThan", "greaterThanOrEqual", "equal", "notEqual", "between", "notBetween"];
        if (!supported.includes(operator ?? "") || (["between", "notBetween"].includes(operator ?? "") && formulas.length < 2)) { warn("xlsx.conditionalFormatting", "Unsupported comparison operator."); continue; }
        for (const { cell, value } of targets) {
          if (blocked.has(cell)) continue;
          const a = formulas[0]!, b = formulas[1]!;
          const matches = operator === "lessThan" ? value < a : operator === "lessThanOrEqual" ? value <= a : operator === "greaterThan" ? value > a : operator === "greaterThanOrEqual" ? value >= a : operator === "equal" ? value === a : operator === "notEqual" ? value !== a : operator === "between" ? value >= a && value <= b : value < a || value > b;
          if (matches) { applyStyle(cell, style); if (attrBool(rule, "stopIfTrue")) blocked.add(cell); }
        }
        continue;
      }
      if (type === "iconSet") {
        const visual = child(rule, "iconSet"), kind = attr(visual, "iconSet") ?? "3TrafficLights1";
        const thresholds = children(visual, "cfvo").map(threshold);
        if (!/^[345](Arrows(?:Gray)?|TrafficLights[12]|Flags|Signs|Symbols2?|Rating|Quarters|RedToBlack)$/.test(kind) || thresholds.some((value) => !Number.isFinite(value)) || thresholds.length !== Number(kind[0])) { warn("xlsx.conditionalFormatting", `Unsupported icon set ${kind}.`); continue; }
        const count = thresholds.length;
        for (const { cell, value } of targets) {
          if (blocked.has(cell) || cell.icon) continue;
          let level = 0;
          for (let i = 1; i < count; i++) if (attrBool(children(visual, "cfvo")[i], "gte") === false ? value > thresholds[i]! : value >= thresholds[i]!) level = i;
          if (attrBool(visual, "reverse")) level = count - 1 - level;
          const ratio = level / (count - 1);
          const text = kind.includes("Arrows") ? (count === 3 ? ["↓", "→", "↑"] : count === 4 ? ["↓", "↘", "↗", "↑"] : ["↓", "↘", "→", "↗", "↑"])[level]! : kind.includes("Flags") ? "⚑" : kind.includes("Symbols") ? ["✕", "!", "✓"][Math.round(ratio * 2)]! : "●";
          cell.icon = { text, label: `${kind}: level ${level + 1} of ${count}`, color: kind.includes("Gray") ? "#666" : ratio < .34 ? "#c43d3d" : ratio < .67 ? "#a57400" : "#20834b" };
          cell.hideValue = attrBool(visual, "showValue") === false;
        }
        continue;
      }
      const visual = child(rule, type === "dataBar" ? "dataBar" : "colorScale");
      const stops = children(visual, "cfvo").map(threshold);
      const colors = children(visual, "color").map((color) => sheetColor(color, theme));
      if (!visual || stops.some((stop) => !Number.isFinite(stop)) || ![2, 3].includes(stops.length) || colors.some((color) => !color) || (type !== "dataBar" && type !== "colorScale")) {
        warn("xlsx.conditionalFormatting", `Conditional formatting rule ${type ?? "unknown"} is not supported.`); continue;
      }
      if (type === "dataBar" && colors[0]) {
        const low = Math.min(stops[0]!, 0), high = Math.max(stops[1]!, 0);

        const loLength = attrNum(visual, "minLength") ?? 10, hiLength = attrNum(visual, "maxLength") ?? 90;
        for (const { cell, value } of targets) {
          if (blocked.has(cell) || cell.dataBar) continue;
          if (low < 0) {
            const axis = high === low ? 100 : -low / (high - low) * 100;
            const width = Math.min(value < 0 ? axis : 100 - axis, Math.abs(value) / Math.max(high - low, 1e-10) * Math.max(0, Math.min(100, hiLength)));
            cell.dataBar = { color: value < 0 ? "#d45a5a" : colors[0]!, left: value < 0 ? axis - width : axis, width };
          } else cell.dataBar = { color: colors[0]!, width: Math.max(0, Math.min(100, low === high ? hiLength : loLength + (hiLength - loLength) * Math.max(0, Math.min(1, (value - low) / (high - low))))) };
          cell.hideValue = attrBool(visual, "showValue") === false;
        }
      } else if (colors.length === stops.length) {
        for (const { cell, value } of targets) {
          const segment = stops.length === 3 && value > stops[1]! ? 1 : 0;
          const low = stops[segment]!, high = stops[segment + 1]!;
          const t = high === low ? .5 : Math.max(0, Math.min(1, (value - low) / (high - low)));
          const a = colors[segment]!.match(/\d+/g)!.map(Number), b = colors[segment + 1]!.match(/\d+/g)!.map(Number);
          if (blocked.has(cell)) continue;
          applyStyle(cell, { backgroundColor: `rgb(${a.map((v, i) => Math.round(v + (b[i]! - v) * t)).join(", ")})` });
        }
      }
    }
  }
}
