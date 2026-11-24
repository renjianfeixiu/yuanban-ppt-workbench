// Prototype: update only identified content, while keeping all other XML intact.
export const P = 'http://schemas.openxmlformats.org/presentationml/2006/main';
export const A = 'http://schemas.openxmlformats.org/drawingml/2006/main';

export function parseXml(xml, sax) {
  const parser = sax.parser(true, { xmlns: true, position: true, trim: false });
  const stack = [];
  let root;
  let nodeCount = 0;
  parser.ondoctype = () => { throw new Error('DTD is not supported'); };
  parser.onopentag = tag => {
    if (stack.length >= 128 || ++nodeCount > 200000) throw new Error('XML 结构超出支持范围。');
    const node = {
      name: tag.name, local: tag.local, uri: tag.uri,
      attrs: Object.fromEntries(Object.entries(tag.attributes).map(([k, v]) => [k, v.value])),
      attributes: Object.values(tag.attributes),
      start: parser.startTagPosition - 1, openEnd: parser.position,
      children: [], text: '', selfClosing: tag.isSelfClosing,
    };
    if (stack.length) stack.at(-1).children.push(node);
    else root = node;
    stack.push(node);
  };
  parser.ontext = text => { if (stack.length) stack.at(-1).text += text; };
  parser.oncdata = parser.ontext;
  parser.onclosetag = () => {
    const node = stack.pop();
    node.end = parser.position;
    node.closeStart = node.selfClosing ? node.openEnd - 2 : xml.lastIndexOf('</', node.end - 1);
  };
  parser.write(xml).close();
  if (!root) throw new Error('XML 缺少根节点。');
  return root;
}

export const child = (node, uri, local) => node?.children.find(x => x.uri === uri && x.local === local);
export function descendants(node, uri, local) {
  const found = [];
  for (const c of node?.children ?? []) {
    if (c.uri === uri && c.local === local) found.push(c);
    found.push(...descendants(c, uri, local));
  }
  return found;
}
export function drawable(root, id, kind = 'sp') {
  const matches = descendants(root, P, kind).filter(n => {
    const nv = child(n, P, kind === 'pic' ? 'nvPicPr' : 'nvSpPr');
    return child(nv, P, 'cNvPr')?.attrs.id === String(id);
  });
  if (matches.length !== 1) throw new Error(`Expected one ${kind} with id ${id}, got ${matches.length}`);
  return matches[0];
}
export const escapeXml = text => String(text).replaceAll('&', '&amp;').replaceAll('<', '&lt;').replaceAll('>', '&gt;').replaceAll('"', '&quot;');

export function textEdits(xml, shape, paragraphs) {
  const sourceParagraphs = child(shape, P, 'txBody').children.filter(n => n.uri === A && n.local === 'p');
  if (sourceParagraphs.length !== paragraphs.length) throw new Error('Paragraph count must stay unchanged');
  const edits = [];
  sourceParagraphs.forEach((para, pi) => {
    const runs = para.children.filter(n => n.uri === A && n.local === 'r');
    const texts = runs.map(run => child(run, A, 't'));
    if (texts.some(t => !t || t.selfClosing)) throw new Error('Unsupported text run');
    const fragments = typeof paragraphs[pi] === 'string' ? [{ text: paragraphs[pi], style: 'any' }] : paragraphs[pi];
    const assigned = Array(runs.length).fill('');
    let cursor = 0;
    for (const fragment of fragments) {
      let target = -1;
      for (let i = cursor; i < runs.length; i++) {
        const emphasis = !!child(child(runs[i], A, 'rPr'), A, 'gradFill');
        if (fragment.style === 'any' || (fragment.style === 'emphasis') === emphasis) { target = i; break; }
      }
      if (target < 0) throw new Error(`The template has no matching style for paragraph ${pi}`);
      assigned[target] = fragment.text;
      cursor = target + 1;
    }
    texts.forEach((t, i) => {
      const replacement = escapeXml(assigned[i]);
      if (xml.slice(t.openEnd, t.closeStart) !== replacement)
        edits.push({ start: t.openEnd, end: t.closeStart, replacement });
    });
  });
  return edits;
}

export function attributeEdit(xml, node, name, value) {
  const open = xml.slice(node.start, node.openEnd);
  const attrPattern = new RegExp(`(^|\\s)${name.replaceAll(':', '\\:')}=("[^"]*"|'[^']*')`);
  const match = attrPattern.exec(open);
  if (match) {
    const quoteStart = match.index + match[1].length + name.length + 1;
    return { start: node.start + quoteStart + 1, end: node.start + quoteStart + match[2].length - 1, replacement: escapeXml(value) };
  }
  const at = node.openEnd - (node.selfClosing ? 2 : 1);
  return { start: at, end: at, replacement: ` ${name}="${escapeXml(value)}"` };
}

export function applyEdits(xml, edits) {
  const sorted = [...edits].sort((a, b) => b.start - a.start);
  let upper = xml.length;
  for (const e of sorted) {
    if (e.end > upper || e.start > e.end) throw new Error('Overlapping patch ranges');
    xml = xml.slice(0, e.start) + e.replacement + xml.slice(e.end);
    upper = e.start;
  }
  return xml;
}

// An intentionally conservative prototype estimate, not WPS font measurement.
export function estimateFit(shape, paragraphs) {
  const transform = child(child(shape, P, 'spPr'), A, 'xfrm');
  const extent = child(transform, A, 'ext');
  const body = child(shape, P, 'txBody');
  const props = child(body, A, 'bodyPr');
  const widthPt = (+extent.attrs.cx - +(props.attrs.lIns ?? 91440) - +(props.attrs.rIns ?? 91440)) / 12700;
  const heightPt = (+extent.attrs.cy - +(props.attrs.tIns ?? 45720) - +(props.attrs.bIns ?? 45720)) / 12700;
  const ps = body.children.filter(x => x.uri === A && x.local === 'p');
  let usedHeightPt = 0;
  let estimatedLines = 0;
  ps.forEach((para, i) => {
    const ppr = child(para, A, 'pPr');
    const fontSize = Math.max(...descendants(para, A, 'rPr').map(x => +(x.attrs.sz ?? 1400) / 100), 14);
    const pct = +(child(child(ppr, A, 'lnSpc'), A, 'spcPct')?.attrs.val ?? 120000) / 100000;
    const available = widthPt - +(ppr?.attrs.marL ?? 0) / 12700;
    const text = typeof paragraphs[i] === 'string' ? paragraphs[i] : paragraphs[i].map(x => x.text).join('');
    const width = [...text].reduce((n, ch) => n + (/[\u2e80-\uffff]/u.test(ch) ? fontSize : fontSize * 0.58), 0);
    const lines = Math.max(1, Math.ceil(width / available));
    estimatedLines += lines;
    usedHeightPt += lines * fontSize * pct;
  });
  return { method: 'approximate-glyph-width', availableHeightPt: +heightPt.toFixed(2), usedHeightPt: +usedHeightPt.toFixed(2), estimatedLines, fits: usedHeightPt <= heightPt * 0.9 };
}

export function lockedSignature(root, textIds, pictureIds) {
  function visit(n, context) {
    let ctx = context;
    if (n.uri === P && ['sp', 'pic'].includes(n.local)) {
      const nv = child(n, P, n.local === 'sp' ? 'nvSpPr' : 'nvPicPr');
      ctx = { kind: n.local, id: child(nv, P, 'cNvPr')?.attrs.id };
    }
    const attrs = { ...n.attrs };
    if (ctx?.kind === 'pic' && pictureIds.has(ctx.id)) {
      if (n.uri === A && n.local === 'blip') {
        for (const attribute of n.attributes) {
          if (attribute.local === 'embed' && attribute.uri === 'http://schemas.openxmlformats.org/officeDocument/2006/relationships') delete attrs[attribute.name];
        }
      }
      if (n.uri === P && n.local === 'cNvPr') delete attrs.descr;
    }
    return {
      name: n.name,
      attrs: Object.fromEntries(Object.entries(attrs).sort(([a], [b]) => a.localeCompare(b))),
      text: n.uri === A && n.local === 't' && ctx?.kind === 'sp' && textIds.has(ctx.id) ? '[CONTENT]' : n.text,
      children: n.children.map(c => visit(c, ctx)),
    };
  }
  return JSON.stringify(visit(root, null));
}
