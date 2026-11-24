// Throwaway prototype: answers whether selected content can change without moving the template.
import fs from 'node:fs/promises';
import path from 'node:path';
import { fileURLToPath, pathToFileURL } from 'node:url';
import { createRequire } from 'node:module';
import { createHash } from 'node:crypto';
import { execFile } from 'node:child_process';
import { promisify } from 'node:util';
import { createInterface } from 'node:readline/promises';
import { P, A, parseXml, child, descendants, drawable, textEdits, attributeEdit, applyEdits, estimateFit, lockedSignature, escapeXml } from './demo-core.mjs';

const DEMO_DIR = path.dirname(fileURLToPath(import.meta.url));
const STUDY_DIR = path.dirname(DEMO_DIR);
const PRESENTATION_PACKAGES = '/Users/wangdajin/.codex/plugins/cache/openai-primary-runtime/presentations';
const installedVersions = (await fs.readdir(PRESENTATION_PACKAGES)).filter(x => /^\d+\./.test(x)).sort((a, b) => a.localeCompare(b, undefined, { numeric: true }));
const SKILL_DIR = process.env.PPT_PRESENTATION_SKILL_DIR ?? path.join(PRESENTATION_PACKAGES, installedVersions.at(-1), 'skills/presentations');
const MODULES = process.env.RUNTIME_NODE_MODULES ?? '/Users/wangdajin/.cache/codex-runtimes/codex-primary-runtime/dependencies/node/node_modules';
const PYTHON = '/Users/wangdajin/.cache/codex-runtimes/codex-primary-runtime/dependencies/python/bin/python3';
const SOURCE = '/Users/wangdajin/我的文稿/大创/防盲区/ppt/平安货运：面向城乡货运场景的商用车智能安全保障系统（马赛克）.pptx';
const SOURCE_SHA = '9ae31fbda906baee21cb1b0fb5290442458e453b9787a7c4e7a30c6659ddfd52';
const req = createRequire(path.join(MODULES, '__prototype__.js'));
const JSZip = req('jszip');
const sax = req('sax');
const sharp = req('sharp');
const run = promisify(execFile);
const sha = bytes => createHash('sha256').update(bytes).digest('hex');
const assert = (condition, message) => { if (!condition) throw new Error(message); };
const state = (stage, evidence) => console.log(JSON.stringify({ stage, ...evidence }));
const writeJson = (target, value) => fs.writeFile(target, JSON.stringify(value, null, 2) + '\n');
process.env.RUNTIME_NODE_MODULES = MODULES;

const option = (name, fallback) => { const at = process.argv.indexOf(name); return at < 0 ? fallback : process.argv[at + 1]; };
const planPath = path.resolve(option('--plan', path.join(DEMO_DIR, 'demo-plan.json')));
const assetsDirectory = path.resolve(option('--assets-dir', path.join(DEMO_DIR, 'assets')));
const renderScope = option('--render-scope', 'all');
assert(['all', 'changed'].includes(renderScope), 'Unsupported render scope.');
const plan = JSON.parse(await fs.readFile(planPath, 'utf8'));
const sourceBytes = await fs.readFile(SOURCE);
assert(sha(sourceBytes) === SOURCE_SHA, 'Source changed: inspect the new baseline before running this prototype.');
const zip = await JSZip.loadAsync(sourceBytes, { checkCRC32: true });
const slidePart = `ppt/slides/slide${plan.slide}.xml`;
const relsPart = `ppt/slides/_rels/slide${plan.slide}.xml.rels`;
const originalXml = await zip.file(slidePart).async('string');
const originalRels = await zip.file(relsPart).async('string');
const root = parseXml(originalXml, sax);
const textIds = new Set(plan.texts.map(x => x.id));
const pictureIds = new Set(plan.pictures.map(x => x.id));
const baseline = new Map();
for (const [name, entry] of Object.entries(zip.files)) if (!entry.dir) baseline.set(name, sha(await entry.async('nodebuffer')));
state('template-read', { sourceSha256: SOURCE_SHA, targetSlide: plan.slide, textSlots: textIds.size, imageSlots: pictureIds.size, lockedOtherSlides: 21 });

// Push an obviously impossible candidate through the same capacity gate used for real copy.
const longText = '需要保留完整的巡检场景、风险说明、技术步骤和验证条件。'.repeat(100);
const negative = estimateFit(drawable(root, '63'), [longText, '这是一段超限测试文字。']);
assert(!negative.fits, 'The oversized-text gate did not reject the negative case.');
state('oversized-copy-rejected', { slotId: '63', characters: longText.length, ...negative });

let interactiveDecision = 'use-prepared-short-copy';
if (process.argv.includes('--interactive')) {
  const cli = createInterface({ input: process.stdin, output: process.stdout });
  console.log('原型状态：超长文字已拒绝；保留原字号和原框。');
  const answer = await cli.question('输入 1 继续生成短文案示例，输入 2 结束并查看超限结果：');
  cli.close();
  if (answer.trim() === '2') { state('stopped-after-negative-check', { negative }); process.exit(0); }
  interactiveDecision = 'user-selected-prepared-short-copy';
}

const fits = plan.texts.filter(x => ['63', '64', '65', '66', '67'].includes(x.id)).map(x => ({ id: x.id, ...estimateFit(drawable(root, x.id), x.paragraphs) }));
assert(fits.every(x => x.fits), 'Prepared body text does not fit the conservative estimate.');
state('short-copy-capacity-pass', { fits, repairMethod: 'pre-authored short alternative; no automatic summarizer implemented' });

const edits = [];
for (const item of plan.texts) edits.push(...textEdits(originalXml, drawable(root, item.id), item.paragraphs));
const addedMedia = [];
const relNodes = [];
const cropEvidence = [];
for (const [i, item] of plan.pictures.entries()) {
  const assetPath = path.join(assetsDirectory, item.file);
  const bytes = await fs.readFile(assetPath);
  const metadata = await sharp(bytes).metadata();
  assert(metadata.format === 'png', `Asset ${item.file} must be PNG.`);
  const pic = drawable(root, item.id, 'pic');
  const blip = child(child(pic, P, 'blipFill'), A, 'blip');
  const nv = child(child(pic, P, 'nvPicPr'), P, 'cNvPr');
  const rid = `rIdDemoWarehouse${i + 1}`;
  const mediaPart = `ppt/media/warehouseDemo${i + 1}.png`;
  assert(!zip.file(mediaPart), 'New media name collides with the source.');
  edits.push(attributeEdit(originalXml, blip, 'r:embed', rid));
  edits.push(attributeEdit(originalXml, nv, 'descr', item.alt));
  relNodes.push(`<Relationship Id="${rid}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/warehouseDemo${i + 1}.png"/>`);
  zip.file(mediaPart, bytes);
  addedMedia.push({ part: mediaPart, asset: item.file, sha256: sha(bytes), width: metadata.width, height: metadata.height });
  cropEvidence.push({ id: item.id, sourceCrop: child(child(pic, P, 'blipFill'), A, 'srcRect')?.attrs ?? null });
}
const patchedXml = applyEdits(originalXml, edits);
const relRoot = parseXml(originalRels, sax);
const patchedRels = applyEdits(originalRels, [{ start: relRoot.closeStart, end: relRoot.closeStart, replacement: relNodes.join('') }]);
const patchedRoot = parseXml(patchedXml, sax);
assert(lockedSignature(root, textIds, pictureIds) === lockedSignature(patchedRoot, textIds, pictureIds), 'Protected template XML changed.');
assert(/<Default\b[^>]*Extension="png"[^>]*ContentType="image\/png"/.test(await zip.file('[Content_Types].xml').async('string')), 'Source has no PNG content type.');
for (const item of plan.texts) {
  const actual = child(drawable(patchedRoot, item.id), P, 'txBody').children.filter(n => n.uri === A && n.local === 'p').map(p => descendants(p, A, 't').map(t => t.text).join(''));
  const expected = item.paragraphs.map(p => typeof p === 'string' ? p : p.map(x => x.text).join(''));
  assert(JSON.stringify(actual) === JSON.stringify(expected), `Text verification failed for slot ${item.id}.`);
}
const patchedRelRoot = parseXml(patchedRels, sax);
const oldRelations = relRoot.children.map(x => x.attrs);
assert(JSON.stringify(oldRelations) === JSON.stringify(patchedRelRoot.children.slice(0, oldRelations.length).map(x => x.attrs)), 'Existing relationships changed.');
const relationIds = patchedRelRoot.children.map(x => x.attrs.Id);
assert(new Set(relationIds).size === relationIds.length, 'Duplicate relationship ID.');
zip.file(slidePart, patchedXml);
zip.file(relsPart, patchedRels);
const allowedChangedParts = new Set([slidePart, relsPart]);
const editedSlideNumbers = new Set([2]);
let extraTextCount = 0;
const notePartsUpdated = [];
const newNoteParts = [];
for (const page of plan.extraPages || []) {
  const number = Number(page.number);
  assert(Number.isInteger(number) && number >= 1 && number <= 22, 'Additional slide number is invalid.');
  const part = `ppt/slides/slide${number}.xml`;
  const xml = await zip.file(part).async('string');
  const tree = parseXml(xml, sax);
  if (page.texts?.length) {
    assert(number !== 2, 'Slide 2 edits belong to the main content plan.');
    const pageEdits = page.texts.flatMap(t => textEdits(xml, drawable(tree, t.id), t.paragraphs));
    const nextXml = applyEdits(xml, pageEdits);
    const nextTree = parseXml(nextXml, sax);
    assert(lockedSignature(tree, new Set(page.texts.map(t => t.id)), new Set()) === lockedSignature(nextTree, new Set(page.texts.map(t => t.id)), new Set()), 'Additional slide protected properties changed.');
    for (const text of page.texts) {
      const actual = child(drawable(nextTree, text.id), P, 'txBody').children.filter(n => n.uri === A && n.local === 'p').map(p => descendants(p, A, 't').map(t => t.text).join(''));
      assert(JSON.stringify(actual) === JSON.stringify(text.paragraphs), 'Additional slide text did not match the plan.');
    }
    zip.file(part, nextXml); allowedChangedParts.add(part); editedSlideNumbers.add(number); extraTextCount += page.texts.length;
  }
  if (page.notes?.trim()) {
    const pageRelsPart=`ppt/slides/_rels/slide${number}.xml.rels`;
    let pageRelsXml=await zip.file(pageRelsPart).async('string');
    const relationships = parseXml(pageRelsXml, sax);
    let relation = relationships.children.find(n => n.attrs.Type?.endsWith('/notesSlide'));
    let notesPart;
    if(!relation){
      notesPart=`ppt/notesSlides/notesDemoSlide${number}.xml`;
      const notesRelsPart=`ppt/notesSlides/_rels/notesDemoSlide${number}.xml.rels`;
      assert(!zip.file(notesPart) && !zip.file(notesRelsPart),'New notes part collides with source.');
      const templatePart=Object.keys(zip.files).find(name=>/^ppt\/notesSlides\/notesSlide\d+\.xml$/.test(name));
      assert(templatePart,'No source notes style is available.');
      const templateNotes=await zip.file(templatePart).async('string');
      const templateRelsPart=path.posix.join(path.posix.dirname(templatePart),'_rels',path.posix.basename(templatePart)+'.rels');
      const templateRels=await zip.file(templateRelsPart).async('string');
      const templateRelsRoot=parseXml(templateRels,sax);
      const owner=templateRelsRoot.children.find(n=>n.attrs.Type?.endsWith('/slide'));
      assert(owner,'Notes template has no slide owner.');
      const newRels=applyEdits(templateRels,[attributeEdit(templateRels,owner,'Target',`../slides/slide${number}.xml`)]);
      zip.file(notesPart,templateNotes);zip.file(notesRelsPart,newRels);newNoteParts.push(notesPart,notesRelsPart);
      const rid=`rIdDemoNotes${number}`;
      assert(!relationships.children.some(n=>n.attrs.Id===rid),'Notes relationship ID collision.');
      pageRelsXml=applyEdits(pageRelsXml,[{start:relationships.closeStart,end:relationships.closeStart,replacement:`<Relationship Id="${rid}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesSlide" Target="../notesSlides/notesDemoSlide${number}.xml"/>`}]);
      zip.file(pageRelsPart,pageRelsXml);allowedChangedParts.add(pageRelsPart);
      const contentTypes=await zip.file('[Content_Types].xml').async('string');
      const contentRoot=parseXml(contentTypes,sax);
      zip.file('[Content_Types].xml',applyEdits(contentTypes,[{start:contentRoot.closeStart,end:contentRoot.closeStart,replacement:`<Override PartName="/${notesPart}" ContentType="application/vnd.openxmlformats-officedocument.presentationml.notesSlide+xml"/>`}]));
      allowedChangedParts.add('[Content_Types].xml');
    }else{
      assert(relation.attrs.TargetMode !== 'External', 'External notes are unsupported.');
      notesPart=path.posix.normalize(path.posix.join(path.posix.dirname(part),relation.attrs.Target));
    }
    assert(notesPart.startsWith('ppt/notesSlides/'), 'Notes target is outside the expected package area.');
    const notesXml = await zip.file(notesPart).async('string');
    const notesRoot = parseXml(notesXml, sax);
    const notesShape = descendants(notesRoot, P, 'sp').find(s => child(child(child(s, P, 'nvSpPr'), P, 'nvPr'), P, 'ph')?.attrs.type === 'body');
    const txBody = child(notesShape, P, 'txBody');
    const paragraphs = txBody?.children.filter(n => n.uri === A && n.local === 'p');
    assert(paragraphs?.length, 'Notes body paragraphs are unavailable.');
    const sourceRunProps = descendants(paragraphs[0], A, 'rPr')[0];
    const propsXml = sourceRunProps ? notesXml.slice(sourceRunProps.start, sourceRunProps.end) : '<a:rPr lang="zh-CN"/>';
    const replacement = page.notes.split('\n').map(line => `<a:p><a:r>${propsXml}<a:t>${escapeXml(line)}</a:t></a:r></a:p>`).join('');
    const notesEdits=[{ start: paragraphs[0].start, end: paragraphs.at(-1).end, replacement }];
    for(const field of descendants(notesRoot,A,'fld')){
      if(field.attrs.type==='slidenum'){
        const text=child(field,A,'t');if(text&&!text.selfClosing)notesEdits.push({start:text.openEnd,end:text.closeStart,replacement:String(number)});
      }
    }
    const nextNotes = applyEdits(notesXml, notesEdits);
    parseXml(nextNotes, sax);
    zip.file(notesPart, nextNotes); allowedChangedParts.add(notesPart); editedSlideNumbers.add(number); notePartsUpdated.push(notesPart);
  }
}

const OUTPUT_DIR = path.resolve(option('--output-dir', path.join(DEMO_DIR, 'output')));
const BUILD_DIR = path.resolve(option('--build-dir', path.join(STUDY_DIR, '.work', 'demo-prototype')));
const projectFilename = (plan.projectName || '智慧仓储').replace(/[\\/:*?"<>|\x00-\x1f]/g, '').slice(0, 40) || '项目';
const demoLabel=(plan.extraPages||[]).some(p=>p.texts?.length || p.notes?.trim())?'逐页修订Demo':'单页替换Demo';
await fs.mkdir(OUTPUT_DIR, { recursive: true });
await fs.mkdir(BUILD_DIR, { recursive: true });
let version = 1;
let finalPath;
for (;;) {
  finalPath = path.join(OUTPUT_DIR, `${projectFilename}_${demoLabel}_v${version}.pptx`);
  const occupied = await Promise.all([finalPath, path.join(BUILD_DIR, `candidate-v${version}.pptx`)].map(p => fs.access(p).then(() => true, () => false)));
  if (!occupied.some(Boolean)) break;
  version++;
}
const candidatePath = path.join(BUILD_DIR, `candidate-v${version}.pptx`);
const candidateBytes = await zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE', compressionOptions: { level: 3 } });
await fs.writeFile(candidatePath, candidateBytes, { flag: 'wx' });
const generatedZip = await JSZip.loadAsync(candidateBytes, { checkCRC32: true });
const changedParts = [];
for (const [name, hash] of baseline) {
  assert(generatedZip.file(name), `Original part missing: ${name}`);
  if (sha(await generatedZip.file(name).async('nodebuffer')) !== hash) changedParts.push(name);
}
assert(changedParts.every(part => allowedChangedParts.has(part)) && changedParts.includes(slidePart), 'Unexpected source parts changed.');
const newParts = Object.entries(generatedZip.files).filter(([name, item]) => !item.dir && !baseline.has(name)).map(([name]) => name);
assert(JSON.stringify(newParts.sort()) === JSON.stringify([...addedMedia.map(x => x.part),...newNoteParts].sort()), 'Unexpected new package parts.');
state('structural-protection-pass', { originalParts: baseline.size, unchangedOriginalParts: baseline.size - changedParts.length, changedParts, addedMedia: newParts, lockedSignatureMatched: true });

const validator = async (script, args) => {
  try {
    const result = await run(PYTHON, [path.join(SKILL_DIR, 'container_tools', script), candidatePath, ...args], { maxBuffer: 8 * 1024 * 1024 });
    return JSON.parse(result.stdout);
  } catch (error) {
    throw new Error(`${script} failed: ${(error.stdout ?? '') + (error.stderr ?? '')}`);
  }
};
const [integrity, geometry] = await Promise.all([
  validator('inspect_presentation_package_integrity.py', ['--fail-on-findings']),
  validator('inspect_presentation_layout_geometry.py', ['--expected-slide-size-emu', '12192000,6858000', '--expected-slide-count', '22']),
]);
await writeJson(path.join(BUILD_DIR, `integrity-v${version}.json`), integrity);
await writeJson(path.join(BUILD_DIR, `geometry-v${version}.json`), geometry);
assert((integrity.findings ?? []).length === 0, 'Package integrity findings.');
assert((geometry.findings ?? []).length === 0, 'Geometry findings.');
state('package-and-geometry-pass', { integrityFindings: (integrity.findings ?? []).length, geometryFindings: (geometry.findings ?? []).length, geometryWarnings: (geometry.warnings ?? []).length });

// Try the standard full-deck finalizer. An inherited chart defect may block the complete-deck certificate.
const { finalizePresentation } = await import(pathToFileURL(path.join(SKILL_DIR, 'container_tools/artifact_tool_utils.mjs')).href);
let fullDeckFinalizer;
try {
  const finalizedPath = path.join(BUILD_DIR, `standard-final-v${version}.pptx`);
  await finalizePresentation({
    workspaceDir: STUDY_DIR,
    candidatePath,
    finalPath: finalizedPath,
    pythonExecutable: PYTHON,
    integrityValidatorPath: path.join(SKILL_DIR, 'container_tools/inspect_presentation_package_integrity.py'),
    layoutValidatorPath: path.join(SKILL_DIR, 'container_tools/inspect_presentation_layout_geometry.py'),
    explicitTotalSlideCount: 22,
    layoutArgs: ['--expected-slide-size-emu', '12192000,6858000'],
    requiredNativeTableOwnerSlides: [],
    requiredNativeChartOwnerSlides: [18],
    requiredEmbeddedWorkbookChartOwnerSlides: [18],
    materializeLiteralChartWorkbooks: false,
    verifyArtifactToolImport: true,
    receiptPath: path.join(BUILD_DIR, `standard-receipt-v${version}.json`),
  });
  fullDeckFinalizer = { status: 'passed', receipt: path.join(BUILD_DIR, `standard-receipt-v${version}.json`) };
} catch (error) {
  fullDeckFinalizer = {
    status: 'blocked', message: error.message, report: error.report ?? error.details ?? null,
    blockClassification: error.message.includes('bounded regular presentation file') && candidateBytes.length > 128 * 1024 * 1024 ? 'validator-file-size-limit' : 'requires-inspection',
    outputFileBytes: candidateBytes.length, chartModuleLimitBytes: 128 * 1024 * 1024,
  };
  await writeJson(path.join(BUILD_DIR, `standard-finalizer-v${version}.json`), fullDeckFinalizer);
}
state('full-deck-finalizer', { status: fullDeckFinalizer.status, message: fullDeckFinalizer.message ?? 'passed' });

const { importRuntimeModule } = await import(pathToFileURL(path.join(SKILL_DIR, 'container_tools/runtime_helpers.mjs')).href);
const { PresentationFile, FileBlob } = await importRuntimeModule('@oai/artifact-tool');
const presentation = await PresentationFile.importPptx(await FileBlob.load(candidatePath));
assert(presentation.slides.items.length === 22, 'First-party importer slide count differs.');
const previewDir = path.join(OUTPUT_DIR, `previews-v${version}`);
await fs.mkdir(previewDir, { recursive: true });
const renderIndices = renderScope === 'changed' ? [...editedSlideNumbers].sort((a, b) => a - b).map(n => n - 1) : presentation.slides.items.map((_, i) => i);
for (const index of renderIndices) {
  const slide = presentation.slides.items[index];
  const png = await presentation.export({ slide, format: 'png', scale: 1 });
  await fs.writeFile(path.join(previewDir, `slide-${index + 1}.png`), Buffer.from(await png.arrayBuffer()));
  state('rendered', { slide: index + 1, total: 22 });
}
const unchangedSlideRenderComparison = [];
for (let number = 1; number <= 22; number++) {
  if (editedSlideNumbers.has(number) || renderScope !== 'all') continue;
  const paths = [path.join(STUDY_DIR, '.work', 'source-render', `slide-${number}.png`), path.join(previewDir, `slide-${number}.png`)];
  const [before, after] = await Promise.all(paths.map(p => sharp(p).removeAlpha().raw().toBuffer({ resolveWithObject: true })));
  unchangedSlideRenderComparison.push({ slide: number, samePixels: JSON.stringify(before.info) === JSON.stringify(after.info) && before.data.equals(after.data) });
}
assert(unchangedSlideRenderComparison.every(x => x.samePixels), 'An unchanged slide rendered differently from the baseline.');
if (renderScope === 'all') state('other-slides-pixel-comparison-pass', { unchangedSlides: 21 });
const previewPath = path.join(previewDir, 'slide-2.png');
const sourcePreview = path.join(STUDY_DIR, '.work', 'source-render', 'slide-2.png');
const comparisonPath = path.join(OUTPUT_DIR, `第二页前后对照_v${version}.png`);
await sharp({ create: { width: 1280, height: 1440, channels: 3, background: '#ffffff' } }).composite([
  { input: sourcePreview, left: 0, top: 0 },
  { input: previewPath, left: 0, top: 720 },
]).png().toFile(comparisonPath);
assert(sha(await fs.readFile(SOURCE)) === SOURCE_SHA, 'Source changed while building.');
await fs.copyFile(candidatePath, finalPath, (await import('node:fs')).constants.COPYFILE_EXCL);
const report = {
  prototype: true, generatedAt: new Date().toISOString(), topic: plan.topic,
  sourcePath: SOURCE, sourceSha256Before: SOURCE_SHA, sourceSha256After: sha(await fs.readFile(SOURCE)),
  outputPath: finalPath, outputSha256: sha(candidateBytes), editedSlides: [...editedSlideNumbers].sort((a, b) => a - b), totalSlides: 22,
  changedOriginalParts: changedParts, unchangedOriginalParts: baseline.size - changedParts.length, originalParts: baseline.size,
  protectedTemplateSignatureMatched: true, newMedia: addedMedia, preservedCrop: cropEvidence,
  textSlotsUpdated: 13 + extraTextCount, imageSlotsUpdated: 5, notePartsUpdated, newNoteParts, planRevision: plan.planRevision,
  oversizedCandidate: { slot: '63', characters: longText.length, rejected: !negative.fits, ...negative },
  repair: { method: option('--plan') ? 'validated-workflow-copy' : interactiveDecision, description: option('--plan') ? 'Current workflow copy is checked against the prepared template adapter and capacity budgets; no arbitrary AI rewrite function is implemented.' : 'Short copy is manually authored in demo-plan.json; no arbitrary AI rewrite function is implemented.' },
  shortCopyFit: fits, integrity, geometry, fullDeckFinalizer,
  unchangedSlideRenderComparison,
  renderScope,
  unmodifiedSlideValidation: renderScope === 'all' ? 'original-part-bytes-and-fresh-rendered-pixels' : 'original-part-bytes; unmodified slides not freshly rendered in this run',
  previews: { changedSlide: previewPath, comparison: comparisonPath, fullDeck: previewDir },
  nativeWpsReview: { status: 'pending', sourceHasMissingFontWarning: true },
  limitations: [
    'Slide 2 is automatically filled. Other pages only change when explicitly edited or annotated; remaining original topic content requires review.',
    'Text fit is a conservative estimate. Native WPS visual inspection is required.',
    'Original font names and formatting are preserved. Missing fonts remain a source-environment issue.',
    'Original slide 18 charts and embedded workbooks are unchanged. A complete-deck data validation may be blocked by inherited source defects.',
    'Slide 2 uses a prepared template adapter and workflow copy rules; general topic generation and automatic aesthetic scoring are not implemented.',
    'Slideshow animation and cross-device fidelity are not verified by this prototype.',
    'Policy citations are written into native notes only for explicitly selected pages. Full-deck topic adaptation is not automatic in this demo.',
    'Artifact Tool preview clips some decorative headings in both source and output. Native WPS review is necessary.',
  ],
};
const reportPath = path.join(OUTPUT_DIR, `验证报告_v${version}.json`);
await writeJson(reportPath, report);
state('demo-exported', { finalPath, reportPath, previewPath, fullDeckFinalizer: fullDeckFinalizer.status, nativeWpsReview: 'pending' });
