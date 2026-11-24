import { createReadStream } from 'node:fs';
import { createHash } from 'node:crypto';
import path from 'node:path';
import { P, A, child, descendants, drawable, textEdits, attributeEdit, applyEdits, escapeXml, lockedSignature } from '../demo-prototype/demo-core.mjs';
import { weightedLength } from './model.mjs';
import { openArchive, writeArchive } from './pptx-archive.mjs';
import { patchNotes } from './pptx-notes.mjs';

import { R, CONTENT_TYPES, xmlTree, relationPart, attribute, relation, emptyRelationships, resolvePart, relativePart, qualified, addRelationship, addContentType } from './pptx-xml.mjs';

export const LEGACY_SHA = '9ae31fbda906baee21cb1b0fb5290442458e453b9787a7c4e7a30c6659ddfd52';
const MAIN_TYPE = 'application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml';
const supportedMedia = { '.png': 'image/png', '.jpg': 'image/jpeg', '.jpeg': 'image/jpeg', '.gif': 'image/gif', '.webp': 'image/webp' };
const posix = path.posix;

export async function sha256(file) {
  const hash = createHash('sha256');
  for await (const chunk of createReadStream(file)) hash.update(chunk);
  return hash.digest('hex');
}

const paragraphsOf = shape => child(shape, P, 'txBody')?.children.filter(n => n.uri === A && n.local === 'p') || [];
const runsOf = paragraph => paragraph.children.filter(n => n.uri === A && n.local === 'r').map(n => child(n, A, 't'));
const textOf = paragraph => paragraph.children.map(n => {
  if (n.uri === A && n.local === 'br') return '\n';
  return child(n, A, 't')?.text || '';
}).join('');
const runFont = para => [...descendants(para, A, 'rPr'), ...descendants(para, A, 'defRPr'), ...descendants(para, A, 'endParaRPr')]
  .map(n => Number(n.attrs.sz) / 100).filter(n => Number.isFinite(n) && n > 0);
const editableParagraph = para => (
  !para.children.some(n => n.uri === A && ['br', 'fld', 'extLst'].includes(n.local)) &&
  runsOf(para).every(t => t && t.children.length === 0)
);
const line = value => String(value ?? '').replace(/[\r\n\t]+/g, ' ').replace(/[\x00-\x08\x0b\x0c\x0e-\x1f\ufffe\uffff]/g, '');
function transformOf(shape) {
  return child(child(shape, P, shape.local === 'grpSp' ? 'grpSpPr' : 'spPr'), A, 'xfrm');
}
function geometry(transform) {
  const offset = child(transform, A, 'off')?.attrs;
  const extent = child(transform, A, 'ext')?.attrs;
  if (!offset || !extent) return null;
  const result = [Number(offset.x), Number(offset.y), Number(extent.cx), Number(extent.cy)];
  return result.every(Number.isFinite) && result[2] > 0 && result[3] > 0 ? result : null;
}
function groupSpace(xfrm, parent) {
  const values = geometry(xfrm);
  const origin = child(xfrm, A, 'chOff')?.attrs;
  const extent = child(xfrm, A, 'chExt')?.attrs;
  if (!values || !(Number(extent?.cx) > 0 && Number(extent?.cy) > 0)) return parent;
  const [x, y, w, h] = values;
  const sx = w / Number(extent.cx), sy = h / Number(extent.cy);
  return {
    x: parent.x + (x - Number(origin?.x || 0) * sx) * parent.sx,
    y: parent.y + (y - Number(origin?.y || 0) * sy) * parent.sy,
    sx: parent.sx * sx,
    sy: parent.sy * sy,
  };
}
function boxOf(xfrm, size, group) {
  const box = geometry(xfrm);
  if (!box) return null;
  return [
    (box[0] * group.sx + group.x) / size.width * 100,
    (box[1] * group.sy + group.y) / size.height * 100,
    box[2] * group.sx / size.width * 100,
    box[3] * group.sy / size.height * 100,
  ].map(n => +n.toFixed(3));
}
function textMetrics(shape, paragraphs, xfrm) {
  const box = geometry(xfrm);
  const body = child(child(shape, P, 'txBody'), A, 'bodyPr')?.attrs || {};
  if (!box || (body.vert && body.vert !== 'horz')) return null;
  const width = (box[2] - Number(body.lIns ?? 91440) - Number(body.rIns ?? 91440)) / 12700;
  const height = (box[3] - Number(body.tIns ?? 45720) - Number(body.bIns ?? 45720)) / 12700;
  if (!(width > 0 && height > 0)) return null;
  return {
    width, height,
    wrap: body.wrap !== 'none',
    paragraphs: paragraphs.map(para => {
      const props = child(para, A, 'pPr');
      const spacing = child(props, A, 'lnSpc');
      const fonts = runFont(para);
      const font = fonts.length ? Math.max(...fonts) : 18;
      const percent = Number(child(spacing, A, 'spcPct')?.attrs.val ?? 110000) / 100000;
      const fixed = Number(child(spacing, A, 'spcPts')?.attrs.val) / 100;
      const margin = Number(props?.attrs.marL || 0) + Number(props?.attrs.marR || 0);
      return { font, line: fixed || font * percent, width: Math.max(1, width - margin / 12700) };
    }),
  };
}
function glyphWidth(text) {
  return [...text].reduce((width, ch) => width + (/\s/u.test(ch) ? 0.33 : /[\u0020-\u024f]/u.test(ch) ? 0.58 : 1), 0);
}
export function estimateSlot(slot, values) {
  if (!slot.metrics) {
    const fits = values.every((value, i) => weightedLength(value) <= slot.budgets[i] + 0.001);
    return { fits, percent: null, estimatedLines: null };
  }
  let height = 0;
  let estimatedLines = 0;
  let widthPercent = 0;
  const metrics = slot.metrics;
  for (const [i, value] of values.entries()) {
    const p = metrics.paragraphs[i];
    const ratio = glyphWidth(value) * p.font / p.width;
    const lines = metrics.wrap ? Math.max(1, Math.ceil(ratio)) : 1;
    if (!metrics.wrap) widthPercent = Math.max(widthPercent, ratio * 100);
    estimatedLines += lines;
    height += p.line * lines;
  }
  const percent = Math.round(Math.max(height / metrics.height * 100, widthPercent));
  return { fits: percent <= 100, percent, estimatedLines };
}
function readShapes(root, slideRelationships, part, size, assignMedia, layoutRoot) {
  const slots = [], pictures = [], blocks = [];
  const layoutShapes = descendants(layoutRoot, P, 'sp');
  const tree = child(child(root, P, 'cSld'), P, 'spTree');
  function visit(node, group) {
    for (const shape of node.children) {
      if (shape.uri !== P) continue;
      if (shape.local === 'grpSp') {
        visit(shape, groupSpace(transformOf(shape), group));
        continue;
      }
      if (shape.local !== 'sp' && shape.local !== 'pic') continue;
      const props = child(child(shape, P, shape.local === 'sp' ? 'nvSpPr' : 'nvPicPr'), P, 'cNvPr');
      const id = props?.attrs.id;
      if (!/^\d+$/.test(id || '')) continue;
      const name = props.attrs.name || `位置 ${id}`;
      let xfrm = transformOf(shape);
      if (!xfrm && shape.local === 'sp') {
        const placeholder = child(child(child(shape, P, 'nvSpPr'), P, 'nvPr'), P, 'ph');
        if (placeholder) {
          const match = layoutShapes.find(layout => {
            const candidate = child(child(child(layout, P, 'nvSpPr'), P, 'nvPr'), P, 'ph');
            return candidate && (candidate.attrs.idx ?? '0') === (placeholder.attrs.idx ?? '0') &&
              (candidate.attrs.type ?? 'obj') === (placeholder.attrs.type ?? 'obj');
          });
          xfrm = transformOf(match || shape);
        }
      }
      const box = boxOf(xfrm, size, group);
      if (shape.local === 'pic') {
        const blip = child(child(shape, P, 'blipFill'), A, 'blip');
        const embed = attribute(blip, R, 'embed');
        const link = relation(slideRelationships, a => a.Id === embed?.value && a.Type === R + '/image' && a.TargetMode !== 'External');
        if (link) {
          const target = resolvePart(part, link.attrs.Target);
          const media = assignMedia(target);
          pictures.push({ id, name, media });
          if (box) blocks.push({ kind: 'picture', id, name, media, box });
        }
        continue;
      }
      const paras = paragraphsOf(shape);
      if (!paras.length) continue;
      const paragraphs = paras.map(textOf);
      const metrics = textMetrics(shape, paras, xfrm);
      const nonVisual = child(shape, P, 'nvSpPr');
      const textBox = child(nonVisual, P, 'cNvSpPr')?.attrs.txBox;
      const placeholder = child(child(nonVisual, P, 'nvPr'), P, 'ph');
      const editable = paras.every(editableParagraph) && (paragraphs.some(Boolean) || textBox === '1' || textBox === 'true' || Boolean(placeholder));
      if (editable) {
        const budgets = paragraphs.map(value => Math.max(1, weightedLength(value)));
        slots.push({ id, name, paragraphs, budgets, metrics });
      }
      if (box && (editable || paragraphs.some(Boolean))) {
        const color = descendants(paras[0], A, 'srgbClr')[0]?.attrs.val;
        const font = metrics?.paragraphs[0]?.font || 18;
        blocks.push({ kind: 'text', id, name, box, paragraphs, editable, font: +(font * 12700 / size.width * 100).toFixed(3), color: /^[0-9a-f]{6}$/i.test(color || '') ? '#' + color : null });
      }
    }
  }
  if (tree) visit(tree, { x: 0, y: 0, sx: 1, sy: 1 });
  return { slots, pictures, blocks };
}

export async function readTemplate(file, name, hash) {
  const archive = await openArchive(file);
  try {
    const sourceHash = hash || await sha256(file);
    const rootRelationships = xmlTree(await archive.xml('_rels/.rels'));
    const main = relation(rootRelationships, a => a.Type === R + '/officeDocument' && a.TargetMode !== 'External');
    if (!main) throw new Error('请选择标准的 .pptx 演示文稿。');
    const presentationPart = resolvePart('', main.attrs.Target);
    const contentTypes = xmlTree(await archive.xml('[Content_Types].xml'));
    if (!child(contentTypes, CONTENT_TYPES, 'Override') || !contentTypes.children.some(n => n.attrs.PartName === '/' + presentationPart && n.attrs.ContentType === MAIN_TYPE)) {
      throw new Error('请选择普通 .pptx 文件。加密文件、.ppt 和宏演示文稿请先另存为 .pptx。');
    }
    const presentation = xmlTree(await archive.xml(presentationPart));
    const rels = xmlTree(await archive.xml(relationPart(presentationPart)));
    const list = child(presentation, P, 'sldIdLst')?.children.filter(n => n.uri === P && n.local === 'sldId') || [];
    if (!list.length || list.length > 200) throw new Error('模板须包含 1–200 页可读取的幻灯片。');
    const sizeNode = child(presentation, P, 'sldSz');
    const size = { width: Number(sizeNode?.attrs.cx), height: Number(sizeNode?.attrs.cy) };
    if (!(size.width > 0 && size.height > 0)) throw new Error('无法读取模板的页面尺寸。');
    const media = [], mediaIds = new Map();
    const assignMedia = target => {
      const mime = supportedMedia[posix.extname(target).toLowerCase()];
      if (!mime || !archive.entries.has(target) || archive.entries.get(target).uncompressedSize > 20 * 1024 * 1024) return null;
      if (!mediaIds.has(target)) {
        const id = media.length;
        mediaIds.set(target, id);
        media.push({ part: target, mime });
      }
      return `/api/templates/${sourceHash}/media/${mediaIds.get(target)}`;
    };
    const pages = [];
    let legacySlideXml;
    const layouts = new Map(), referencedParts = new Set();
    for (const slide of list) {
      const link = relation(rels, a => a.Id === attribute(slide, R, 'id')?.value && a.Type === R + '/slide' && a.TargetMode !== 'External');
      if (!link) throw new Error('PPTX 的页面列表与内部引用不匹配。');
      const part = resolvePart(presentationPart, link.attrs.Target);
      if (referencedParts.has(part)) throw new Error('模板的页面列表含有重复部件。');
      referencedParts.add(part);
      const xml = await archive.xml(part);
      if (sourceHash === LEGACY_SHA && pages.length === 1) legacySlideXml = xml;
      const root = xmlTree(xml);
      if (root.uri !== P || root.local !== 'sld') throw new Error('模板含有不支持的页面类型。');
      const relPart = relationPart(part);
      const relations = archive.entries.has(relPart) ? xmlTree(await archive.xml(relPart)) : null;
      const layout = relation(relations, a => a.Type === R + '/slideLayout' && a.TargetMode !== 'External');
      const layoutPart = layout ? resolvePart(part, layout.attrs.Target) : null;
      if (layoutPart && !layouts.has(layoutPart)) layouts.set(layoutPart, xmlTree(await archive.xml(layoutPart)));
      const shapes = readShapes(root, relations, part, size, assignMedia, layouts.get(layoutPart));
      const background = child(child(child(child(root, P, 'cSld'), P, 'bg'), P, 'bgPr'), A, 'solidFill');
      const color = child(background, A, 'srgbClr')?.attrs.val;
      pages.push({
        number: pages.length + 1, part, ...shapes,
        color: /^[0-9a-f]{6}$/i.test(color || '') ? '#' + color : '#f9f8f5',
      });
    }
    return { file, name, sha: sourceHash, presentationPart, size, pages, media, legacySlideXml, adapter: sourceHash === LEGACY_SHA ? 'warehouse' : 'generic' };
  } finally { archive.close(); }
}

function replacementEdits(xml, shape, values) {
  return paragraphsOf(shape).flatMap((para, i) => {
    const value = line(values[i]);
    if (value === textOf(para)) return [];
    if (!editableParagraph(para)) throw new Error('该段落包含动态字段或手动换行，请在 WPS 中直接修改。');
    const texts = runsOf(para);
    if (!texts.length) {
      if (!value) return [];
      const r = qualified(para, 'r'), t = qualified(para, 't');
      const fragment = `<${r}><${t}>${escapeXml(value)}</${t}></${r}>`;
      if (para.selfClosing) return [{ start: para.openEnd - 2, end: para.openEnd, replacement: '>' + fragment + '</' + para.name + '>' }];
      const tail = child(para, A, 'endParaRPr');
      const at = tail?.start ?? para.closeStart;
      return [{ start: at, end: at, replacement: fragment }];
    }
    // Preserve run boundaries and formatting. The last run receives any extra characters.
    const characters = [...value];
    return texts.flatMap((text, j) => {
      const next = escapeXml(characters.splice(0, j === texts.length - 1 ? characters.length : [...text.text].length).join(''));
      if (text.selfClosing) {
        return next ? [{ start: text.openEnd - 2, end: text.openEnd, replacement: '>' + next + '</' + text.name + '>' }] : [];
      }
      return xml.slice(text.openEnd, text.closeStart) === next ? [] : [{ start: text.openEnd, end: text.closeStart, replacement: next }];
    });
  });
}

export async function exportPresentation({ deck, edits, destination }) {
  const originalHash = await sha256(deck.file);
  if (originalHash !== deck.sha) throw new Error('模板副本已在工作台之外修改，请重新上传。');
  const archive = await openArchive(deck.file);
  try {
    const patches = new Map();
    const load = async part => patches.get(part) ?? await archive.xml(part);
    const editPages = [];
    let textsUpdated = 0, picturesUpdated = 0;
    for (const page of deck.pages) {
      const edit = edits[page.number];
      if (!edit || (!edit.texts.length && !edit.pictures.length && !edit.notes)) continue;
      const initial = await load(page.part);
      const root = xmlTree(initial);
      const xmlEdits = [];
      const richIds = new Set();
      for (const { id, paragraphs, rich } of edit.texts) {
        const shape = drawable(root, id);
        if (paragraphs.length !== paragraphsOf(shape).length) throw new Error('模板段落数量已变化。');
        if (rich) {
          richIds.add(id);
          xmlEdits.push(...textEdits(initial, shape, paragraphs));
        } else xmlEdits.push(...replacementEdits(initial, shape, paragraphs));
      }
      const relsPart = relationPart(page.part);
      let rels = archive.entries.has(relsPart) ? await load(relsPart) : emptyRelationships;
      const pictureIds = new Set();
      for (const picture of edit.pictures) {
        const shape = drawable(root, picture.id, 'pic');
        const blip = child(child(shape, P, 'blipFill'), A, 'blip');
        const embed = attribute(blip, R, 'embed');
        if (!embed) throw new Error('此图片使用外链，不能替换。');
        let index = picturesUpdated + 1, rid, mediaPart;
        do {
          rid = 'rIdYuanbanImage' + index;
          mediaPart = posix.join(posix.dirname(deck.presentationPart), 'media', 'yuanban-image-' + index + '.png');
          index++;
        } while (archive.entries.has(mediaPart) || patches.has(mediaPart) || relation(xmlTree(rels), a => a.Id === rid));
        pictureIds.add(picture.id);
        xmlEdits.push(attributeEdit(initial, blip, embed.name, rid));
        rels = addRelationship(rels, rid, 'image', relativePart(page.part, mediaPart));
        patches.set(mediaPart, picture.bytes);
        picturesUpdated++;
      }
      if (xmlEdits.length) {
        const next = applyEdits(initial, xmlEdits);
        const updated = xmlTree(next);
        if (edit.texts.every(t => t.rich) && lockedSignature(root, richIds, pictureIds) !== lockedSignature(updated, richIds, pictureIds)) {
          throw new Error('模板固定属性发生意外变化，已停止导出。');
        }
        for (const item of edit.texts) {
          const expected = item.rich ? item.paragraphs.map(p => typeof p === 'string' ? p : p.map(fragment => fragment.text).join('')) : item.paragraphs.map(line);
          if (JSON.stringify(paragraphsOf(drawable(updated, item.id)).map(textOf)) !== JSON.stringify(expected)) {
            throw new Error('替换后的文字未通过回读检查。');
          }
        }
        patches.set(page.part, next);
      }
      if (edit.pictures.length) patches.set(relsPart, rels);
      if (edit.notes) await patchNotes({ archive, patches, deck, page, notes: edit.notes, replaceBody: deck.adapter === 'warehouse' && page.number === 2 });
      if (xmlEdits.length || edit.notes) editPages.push(page.number);
      textsUpdated += edit.texts.length;
    }
    if (picturesUpdated) {
      const part = '[Content_Types].xml', xml = await load(part);
      const png = xmlTree(xml).children.find(n => n.local === 'Default' && n.attrs.Extension === 'png');
      if (!png) patches.set(part, addContentType(xml, 'Default', { Extension: 'png', ContentType: 'image/png' }));
      else if (png.attrs.ContentType !== 'image/png') throw new Error('模板 PNG 类型声明不兼容。');
    }
    for (const [name, data] of patches) if (typeof data === 'string' && /\.(xml|rels)$/.test(name)) xmlTree(data);
    await writeArchive(archive, patches, destination);
    if (await sha256(deck.file) !== originalHash) throw new Error('导出期间原模板发生变化。');
    const changedParts = [...patches.keys()].filter(part => archive.entries.has(part));
    return {
      sourceSha256: originalHash, outputSha256: await sha256(destination), changedParts,
      newParts: [...patches.keys()].filter(part => !archive.entries.has(part)),
      unchangedParts: archive.entries.size - changedParts.length, editedSlides: editPages, textsUpdated, picturesUpdated,
    };
  } finally { archive.close(); }
}
