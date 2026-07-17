const XML_ENTITIES: Record<string, string> = {
  amp: "&",
  apos: "'",
  gt: ">",
  lt: "<",
  quot: '"'
};

export function decodeXml(value: string): string {
  return value.replace(/&(#x[\da-f]+|#\d+|amp|apos|gt|lt|quot);/gi, (match, entity: string) => {
    if (entity.startsWith("#x")) return String.fromCodePoint(Number.parseInt(entity.slice(2), 16));
    if (entity.startsWith("#")) return String.fromCodePoint(Number.parseInt(entity.slice(1), 10));
    return XML_ENTITIES[entity.toLowerCase()] ?? match;
  });
}

export function encodeXml(value: string): string {
  return value
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&apos;");
}

export function visibleText(xml: string): string {
  const pieces: string[] = [];
  const tokenPattern = /<w:t(?:\s[^>]*)?>([\s\S]*?)<\/w:t>|<w:tab\s*\/>|<w:br(?:\s[^>]*)?\s*\/>/g;
  for (const match of xml.matchAll(tokenPattern)) {
    if (match[1] !== undefined) pieces.push(decodeXml(match[1]));
    else if (match[0].startsWith("<w:tab")) pieces.push("\t");
    else pieces.push("\n");
  }
  return pieces.join("");
}

export function xmlAttribute(xml: string, qualifiedName: string): string | null {
  const escaped = qualifiedName.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
  return xml.match(new RegExp(`${escaped}=["']([^"']+)["']`))?.[1] ?? null;
}

export function tagVal(xml: string, tagName: string): string | null {
  const escaped = tagName.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
  const tag = xml.match(new RegExp(`<${escaped}\\b[^>]*\\bw:val=["']([^"']+)["'][^>]*/?>`));
  return tag?.[1] ?? null;
}

export function childXml(xml: string, tagName: string): string | null {
  const escaped = tagName.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
  return xml.match(new RegExp(`<${escaped}\\b[^>]*>[\\s\\S]*?<\\/${escaped}>`))?.[0] ?? null;
}

export function stripVisiblePrefix(paragraphXml: string, count: number): string {
  let remaining = count;
  return paragraphXml.replace(/<w:t(\s[^>]*)?>([\s\S]*?)<\/w:t>|<w:tab\s*\/>/g, (whole, attrs: string | undefined, inner: string | undefined) => {
    if (remaining <= 0) return whole;
    if (inner === undefined) {
      remaining -= 1;
      return remaining >= 0 ? "" : whole;
    }
    const decoded = decodeXml(inner);
    if (decoded.length <= remaining) {
      remaining -= decoded.length;
      return "";
    }
    const next = decoded.slice(remaining);
    remaining = 0;
    const nextAttrs = /xml:space=/.test(attrs ?? "") || /^\s|\s$/.test(next)
      ? attrs ?? ' xml:space="preserve"'
      : attrs ?? "";
    return `<w:t${nextAttrs}>${encodeXml(next)}</w:t>`;
  });
}

export function paragraphProperties(paragraphXml: string): string | null {
  return paragraphXml.match(/<w:pPr\b[^>]*>[\s\S]*?<\/w:pPr>/)?.[0] ?? null;
}

export function setParagraphNumbering(
  paragraphXml: string,
  numId: number,
  level: number,
  styleId: string
): string {
  const numbering = `<w:numPr><w:ilvl w:val="${level}"/><w:numId w:val="${numId}"/></w:numPr>`;
  const properties = paragraphProperties(paragraphXml);
  if (!properties) {
    return paragraphXml.replace(/^(<w:p\b[^>]*>)/, `$1<w:pPr><w:pStyle w:val="${styleId}"/>${numbering}</w:pPr>`);
  }
  let next = properties.replace(/<w:numPr\b[^>]*>[\s\S]*?<\/w:numPr>/g, "");
  if (/<w:pStyle\b/.test(next)) {
    next = next.replace(/<w:pStyle\b[^>]*\/>/, `<w:pStyle w:val="${styleId}"/>`);
  } else {
    next = next.replace(/^<w:pPr\b[^>]*>/, `$&<w:pStyle w:val="${styleId}"/>`);
  }
  next = next.replace(/<\/w:pPr>$/, `${numbering}</w:pPr>`);
  return paragraphXml.replace(properties, next);
}

export function addBookmark(paragraphXml: string, id: number, name: string): string {
  if (paragraphXml.includes(`w:name="${name}"`)) return paragraphXml;
  const start = `<w:bookmarkStart w:id="${id}" w:name="${name}"/>`;
  const end = `<w:bookmarkEnd w:id="${id}"/>`;
  const properties = paragraphProperties(paragraphXml);
  if (properties) {
    return paragraphXml.replace(properties, `${properties}${start}`).replace(/<\/w:p>$/, `${end}</w:p>`);
  }
  return paragraphXml.replace(/^(<w:p\b[^>]*>)/, `$1${start}`).replace(/<\/w:p>$/, `${end}</w:p>`);
}

function textRun(value: string, runProperties: string): string {
  if (!value) return "";
  const space = /^\s|\s$/.test(value) ? ' xml:space="preserve"' : "";
  return `<w:r>${runProperties}<w:t${space}>${encodeXml(value)}</w:t></w:r>`;
}

export function replaceTextWithRefField(
  paragraphXml: string,
  display: string,
  bookmarkName: string
): { xml: string; replaced: boolean } {
  interface TextRunLocation {
    xmlStart: number;
    xmlEnd: number;
    textStart: number;
    textEnd: number;
    text: string;
    runProperties: string;
  }

  const locations: TextRunLocation[] = [];
  let searchable = "";
  let insideField = false;
  for (const match of paragraphXml.matchAll(/<w:r\b[^>]*>[\s\S]*?<\/w:r>/g)) {
    const runXml = match[0];
    const xmlStart = match.index ?? 0;
    const beginsField = /<w:fldChar\b[^>]*\bw:fldCharType=["']begin["']/.test(runXml);
    const endsField = /<w:fldChar\b[^>]*\bw:fldCharType=["']end["']/.test(runXml);
    if (beginsField) insideField = true;
    if (insideField || /<w:instrText\b/.test(runXml)) {
      searchable += "\u0000";
      if (endsField) insideField = false;
      continue;
    }
    if (/<w:(?:tab|br|cr|drawing|object)\b/.test(runXml)) {
      searchable += "\u0000";
      continue;
    }
    const text = visibleText(runXml);
    if (!text) continue;
    const textStart = searchable.length;
    searchable += text;
    locations.push({
      xmlStart,
      xmlEnd: xmlStart + runXml.length,
      textStart,
      textEnd: searchable.length,
      text,
      runProperties: runXml.match(/<w:rPr\b[^>]*>[\s\S]*?<\/w:rPr>/)?.[0] ?? ""
    });
    if (endsField) insideField = false;
  }

  const matchStart = searchable.indexOf(display);
  if (matchStart < 0) return { xml: paragraphXml, replaced: false };
  const matchEnd = matchStart + display.length;
  const startRun = locations.find((location) => location.textStart <= matchStart && location.textEnd > matchStart);
  const endRun = locations.find((location) => location.textStart < matchEnd && location.textEnd >= matchEnd);
  if (!startRun || !endRun) return { xml: paragraphXml, replaced: false };
  const before = startRun.text.slice(0, matchStart - startRun.textStart);
  const after = endRun.text.slice(matchEnd - endRun.textStart);
  const instruction = ` REF ${bookmarkName} \\r \\h `;
  const replacement = [
    textRun(before, startRun.runProperties),
    `<w:r><w:fldChar w:fldCharType="begin" w:dirty="true"/></w:r>`,
    `<w:r><w:instrText xml:space="preserve">${encodeXml(instruction)}</w:instrText></w:r>`,
    `<w:r><w:fldChar w:fldCharType="separate"/></w:r>`,
    textRun(display, startRun.runProperties),
    `<w:r><w:fldChar w:fldCharType="end"/></w:r>`,
    textRun(after, endRun.runProperties)
  ].join("");
  return {
    xml: `${paragraphXml.slice(0, startRun.xmlStart)}${replacement}${paragraphXml.slice(endRun.xmlEnd)}`,
    replaced: true
  };
}

export function maxNumericAttribute(xml: string, qualifiedName: string): number {
  const escaped = qualifiedName.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
  let max = 0;
  for (const match of xml.matchAll(new RegExp(`${escaped}=["'](\\d+)["']`, "g"))) {
    max = Math.max(max, Number(match[1]));
  }
  return max;
}
