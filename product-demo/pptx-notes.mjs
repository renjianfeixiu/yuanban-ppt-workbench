import path from 'node:path';
import { A, P, child, descendants, applyEdits, escapeXml } from '../demo-prototype/demo-core.mjs';
import { R, xmlTree, relationPart, relation, emptyRelationships, resolvePart, relativePart, qualified, appendXml, addRelationship, addContentType } from './pptx-xml.mjs';

const CONTENT = '[Content_Types].xml';
const posix = path.posix;
const prefix = `xmlns:p="${P}" xmlns:a="${A}" xmlns:r="${R}"`;
const group = '<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/>';
const notesType = kind => `application/vnd.openxmlformats-officedocument.presentationml.${kind}+xml`;
const xmlHeader = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';

function availablePart(archive, patches, folder, stem) {
  let part, i = 1;
  do { part = posix.join(folder, stem + i++ + '.xml'); }
  while (archive.entries.has(part) || patches.has(part) || archive.entries.has(relationPart(part)));
  return part;
}
function availableId(xml) {
  const existing = new Set(xmlTree(xml).children.map(n => n.attrs.Id));
  let i = 1;
  while (existing.has('rIdYuanbanNotes' + i)) i++;
  return 'rIdYuanbanNotes' + i;
}
function noteParas(notes) {
  return notes.split('\n').map(text => text
    ? `<a:p><a:r><a:t>${escapeXml(text)}</a:t></a:r></a:p>`
    : '<a:p/>').join('');
}
function notesShape(type, id, contents) {
  return `<p:sp ${prefix}><p:nvSpPr><p:cNvPr id="${id}" name="${type}"/><p:cNvSpPr/><p:nvPr><p:ph type="${type}"/></p:nvPr></p:nvSpPr><p:spPr/><p:txBody><a:bodyPr/><a:lstStyle/>${contents}</p:txBody></p:sp>`;
}

async function ensureNotesMaster(archive, patches, deck, read) {
  const presentation = deck.presentationPart;
  const relPart = relationPart(presentation);
  const rels = await read(relPart);
  const master = relation(xmlTree(rels), a => a.Type === R + '/notesMaster' && a.TargetMode !== 'External');
  if (master) return resolvePart(presentation, master.attrs.Target);

  const target = availablePart(archive, patches, posix.join(posix.dirname(presentation), 'notesMasters'), 'yuanban-notes-master-');
  const id = availableId(rels);
  patches.set(relPart, addRelationship(rels, id, 'notesMaster', relativePart(presentation, target)));
  const firstTheme = xmlTree(await read(CONTENT)).children.find(n => n.attrs.ContentType === 'application/vnd.openxmlformats-officedocument.theme+xml');
  if (firstTheme && archive.entries.has(firstTheme.attrs.PartName?.slice(1))) {
    patches.set(relationPart(target), addRelationship(emptyRelationships, 'rId1', 'theme', relativePart(target, firstTheme.attrs.PartName.slice(1))));
  }
  const colorMap = '<p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/>';
  patches.set(target, `${xmlHeader}<p:notesMaster ${prefix}><p:cSld><p:spTree>${group}${notesShape('sldImg', 2, '<a:p/>')}${notesShape('body', 3, '<a:p/>')}</p:spTree></p:cSld>${colorMap}</p:notesMaster>`);
  patches.set(CONTENT, addContentType(await read(CONTENT), 'Override', { PartName: '/' + target, ContentType: notesType('notesMaster') }));

  const xml = await read(presentation), root = xmlTree(xml);
  const list = child(root, P, 'notesMasterIdLst');
  const fragment = `<${qualified(root, 'notesMasterId')} xmlns:r="${R}" r:id="${id}"/>`;
  if (list) patches.set(presentation, appendXml(xml, list, fragment));
  else {
    const after = child(root, P, 'sldMasterIdLst')?.end ?? root.openEnd;
    const tag = qualified(root, 'notesMasterIdLst');
    patches.set(presentation, applyEdits(xml, [{ start: after, end: after, replacement: `<${tag}>${fragment}</${tag}>` }]));
  }
  return target;
}

export async function patchNotes({ archive, patches, deck, page, notes, replaceBody }) {
  const read = async part => patches.get(part) ?? await archive.xml(part);
  const slideRelsPart = relationPart(page.part);
  const slideRels = patches.has(slideRelsPart) || archive.entries.has(slideRelsPart) ? await read(slideRelsPart) : emptyRelationships;
  const existing = relation(xmlTree(slideRels), a => a.Type === R + '/notesSlide' && a.TargetMode !== 'External');
  const contents = noteParas(notes);
  if (existing) {
    const part = resolvePart(page.part, existing.attrs.Target);
    const xml = await read(part);
    const root = xmlTree(xml);
    const body = descendants(root, P, 'sp').find(shape => child(child(child(shape, P, 'nvSpPr'), P, 'nvPr'), P, 'ph')?.attrs.type === 'body');
    const txBody = child(body, P, 'txBody');
    const paragraphs = contents.replaceAll('<a:p>', `<a:p xmlns:a="${A}">`).replaceAll('<a:p/>', `<a:p xmlns:a="${A}"/>`);
    if (txBody) {
      const previous = txBody.children.filter(n => n.uri === A && n.local === 'p');
      const updated = replaceBody && previous.length
        ? applyEdits(xml, [{ start: previous[0].start, end: previous.at(-1).end, replacement: paragraphs }])
        : appendXml(xml, txBody, paragraphs);
      patches.set(part, updated);
    }
    else {
      const tree = child(child(root, P, 'cSld'), P, 'spTree');
      const id = Math.max(1, ...descendants(tree, P, 'cNvPr').map(n => Number(n.attrs.id) || 0)) + 1;
      patches.set(part, appendXml(xml, tree, notesShape('body', id, contents)));
    }
    return;
  }

  const master = await ensureNotesMaster(archive, patches, deck, read);
  const folder = posix.join(posix.dirname(deck.presentationPart), 'notesSlides');
  const target = availablePart(archive, patches, folder, 'yuanban-notes-');
  const trip = addRelationship(emptyRelationships, 'rId1', 'slide', relativePart(target, page.part));
  patches.set(relationPart(target), addRelationship(trip, 'rId2', 'notesMaster', relativePart(target, master)));
  patches.set(target, `${xmlHeader}<p:notes ${prefix}><p:cSld><p:spTree>${group}${notesShape('sldImg', 2, '<a:p/>')}${notesShape('body', 3, contents)}</p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:notes>`);
  patches.set(slideRelsPart, addRelationship(slideRels, availableId(slideRels), 'notesSlide', relativePart(page.part, target)));
  patches.set(CONTENT, addContentType(await read(CONTENT), 'Override', { PartName: '/' + target, ContentType: notesType('notesSlide') }));
}
