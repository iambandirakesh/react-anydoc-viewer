import { readText, type Archive } from "../archive";
import { relationshipsFor } from "../relationships";
import { attr, child, children, parseXml, relAttr, type XmlNode } from "../xml";

const node = (local: string, text = "", attrs: Record<string, string> = {}, items: XmlNode[] = []): XmlNode => ({ name: `c:${local}`, local, text, attrs, children: items });

/** Supply chart caches from direct worksheet references when the producer omits them. No formula evaluation. */
export function createChartReferenceResolver(archive: Archive): (chart: XmlNode | null) => void {
  const workbook = parseXml(readText(archive, "xl/workbook.xml") ?? "");
  const rels = relationshipsFor(archive, "xl/workbook.xml");
  const sheets = new Map(children(child(workbook, "sheets"), "sheet").map((sheet) => [attr(sheet, "name")?.toLowerCase(), rels.get(relAttr(sheet, "id") ?? "")]));
  const stringsRoot = parseXml(readText(archive, "xl/sharedStrings.xml") ?? "");
  const textOf = (item: XmlNode): string => item.local === "t" ? item.text : item.children.map(textOf).join("");
  const strings = children(stringsRoot, "si").map(textOf);
  const cached = new Map<string, Map<string, string | number>>();
  const cellsFor = (name: string) => {
    const rel = sheets.get(name.toLowerCase());
    if (!rel || rel.external) return null;
    if (cached.has(rel.target)) return cached.get(rel.target)!;
    const values = new Map<string, string | number>();
    const root = parseXml(readText(archive, rel.target) ?? "");
    let count = 0;
    for (const row of children(child(root, "sheetData"), "row")) for (const cell of children(row, "c")) {
      if (++count > 50000) break;
      const address = attr(cell, "r")?.toUpperCase();
      if (!address) continue;
      const raw = child(cell, "v")?.text ?? "";
      const type = attr(cell, "t");
      if (type === "e") continue;
      const value = type === "s" ? strings[Number(raw)] ?? "" : type === "inlineStr" ? textOf(child(cell, "is") ?? cell) : type === "str" ? raw : raw.trim() && Number.isFinite(Number(raw)) ? Number(raw) : "";
      values.set(address, value);
    }
    cached.set(rel.target, values);
    return values;
  };
  const col = (text: string) => [...text.toUpperCase()].reduce((value, ch) => value * 26 + ch.charCodeAt(0) - 64, 0) - 1;
  const label = (column: number) => { let result = "", value = column + 1; while (value) { value--; result = String.fromCharCode(65 + value % 26) + result; value = Math.floor(value / 26); } return result; };
  return (chart) => {
    let budget = 50000;
    const visit = (item: XmlNode) => {
      if (["numRef", "strRef"].includes(item.local) && !child(item, "numCache") && !child(item, "strCache")) {
        const formula = child(item, "f")?.text.trim() ?? "";
        const match = /^(?:'((?:[^']|'')+)'|([^!\[\]]+))!\$?([A-Z]+)\$?(\d+)(?::\$?([A-Z]+)\$?(\d+))?$/i.exec(formula);
        if (match) {
          const values = cellsFor((match[1] ?? match[2]!).replace(/''/g, "'"));
          const firstCol = col(match[3]!), lastCol = col(match[5] ?? match[3]!);
          const firstRow = Number(match[4]), lastRow = Number(match[6] ?? match[4]);
          const length = (lastCol - firstCol + 1) * (lastRow - firstRow + 1);
          if (values && firstRow > 0 && lastRow >= firstRow && lastCol >= firstCol && length > 0 && length <= budget) {
            budget -= length;
            const items: (string | number)[] = [];
            for (let r = firstRow; r <= lastRow; r++) for (let c = firstCol; c <= lastCol; c++) items.push(values.get(`${label(c)}${r}`) ?? "");
            const numeric = item.local === "numRef" && !items.some((value) => typeof value === "string" && value !== "");
            if (!numeric) { item.local = "strRef"; item.name = "c:strRef"; }
            item.children.push(node(numeric ? "numCache" : "strCache", "", {}, [node("ptCount", "", { val: String(items.length) }), ...items.map((value, index) => node("pt", "", { idx: String(index) }, [node("v", String(value))]))]));
          }
        }
      }
      for (const nested of item.children) visit(nested);
    };
    if (chart) visit(chart);
  };
}
