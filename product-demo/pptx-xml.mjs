import path from 'node:path';
import sax from 'sax';
import { parseXml, applyEdits, escapeXml } from '../demo-prototype/demo-core.mjs';

export const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
export const PACKAGE_REL = 'http://schemas.openxmlformats.org/package/2006/relationships';
export const CONTENT_TYPES = 'http://schemas.openxmlformats.org/package/2006/content-types';
const posix = path.posix;

export const xmlTree = xml => parseXml(xml, sax);
export const relationPart = part => posix.join(posix.dirname(part), '_rels', posix.basename(part) + '.rels');
export const attribute = (node, namespace, name) => node?.attributes.find(a => a.uri === namespace && a.local === name);
export const relation = (root, predicate) => root?.children.find(n => n.uri === PACKAGE_REL && n.local === 'Relationship' && predicate(n.attrs));
export const emptyRelationships = '<Relationships xmlns="' + PACKAGE_REL + '"></Relationships>';
export function resolvePart(owner, target) {
  if (!target) throw new Error('PPTX 中的内部引用缺少目标。');
  const decoded = decodeURI(target);
  const full = decoded.startsWith('/') ? posix.normalize(decoded).slice(1) : posix.normalize(posix.join(posix.dirname(owner), decoded));
  if (full === '..' || full.startsWith('../') || full.includes('\\')) throw new Error('PPTX 包含无效的部件路径。');
  return full;
}
export function relativePart(owner, target) {
  return posix.relative(posix.dirname(owner), target);
}
export function qualified(node, local) {
  const colon = node.name.indexOf(':');
  return colon < 0 ? local : node.name.slice(0, colon + 1) + local;
}
export function appendXml(xml, node, fragment) {
  if (node.selfClosing) {
    return applyEdits(xml, [{ start: node.openEnd - 2, end: node.openEnd, replacement: '>' + fragment + '</' + node.name + '>' }]);
  }
  return applyEdits(xml, [{ start: node.closeStart, end: node.closeStart, replacement: fragment }]);
}
export function addRelationship(xml, id, type, target) {
  const root = xmlTree(xml);
  const name = qualified(root, 'Relationship');
  return appendXml(xml, root, `<${name} Id="${id}" Type="${R}/${type}" Target="${escapeXml(target)}"/>`);
}
export function addContentType(xml, kind, attrs) {
  const root = xmlTree(xml);
  const value = Object.entries(attrs).map(([key, v]) => `${key}="${escapeXml(v)}"`).join(' ');
  return appendXml(xml, root, `<${qualified(root, kind)} ${value}/>`);
}
