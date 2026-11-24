import http from 'node:http';
import fs from 'node:fs/promises';
import { createReadStream, createWriteStream } from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { randomUUID, createHash } from 'node:crypto';
import { execFile, spawn } from 'node:child_process';
import { promisify } from 'node:util';
import { Transform } from 'node:stream';
import { pipeline } from 'node:stream/promises';
import { PNG } from 'pngjs';
import { columns, planToDraft, normalizeDraft, plainText } from './model.mjs';
import { initialWorkflow, normalizeProject, outlineIssues, invalidateOutline, isApproved, canUsePlan, makePlanDraft, planToPptDraft, workflowHash } from './workflow.mjs';
import { searchPolicies, verifiedPolicies, nationalOfficialUrl } from './policies.mjs';
import { readTemplate, exportPresentation, LEGACY_SHA } from './pptx.mjs';
import { openArchive, MAX_TEMPLATE_BYTES } from './pptx-archive.mjs';
import { automaticEdits, buildEdits, planFields } from './content.mjs';

const ROOT = path.dirname(fileURLToPath(import.meta.url));
const REPO = path.dirname(ROOT);
const DATA = path.resolve(process.env.PPT_DEMO_DATA_DIR || path.join(ROOT, '.work'));
const OUTPUT = process.env.PPT_DEMO_DATA_DIR ? path.join(DATA, 'output') : path.join(ROOT, 'output');
const TEMPLATE_DIR = path.join(DATA, 'templates');
const UPLOAD_DIR = path.join(DATA, 'uploads');
const STATE = path.join(DATA, 'workflow.json');
const ASSET_DIR = path.join(REPO, 'demo-prototype', 'assets');
const SAMPLE_FILE = path.join(REPO, 'examples', '样例模板.pptx');
const SAMPLE_NAME = '样例模板.pptx';
const PORT = Number(process.env.PPT_DEMO_PORT ?? 5186);
const OPEN_BROWSER = process.argv.includes('--open-browser') && process.env.PPT_DEMO_NO_BROWSER !== '1';
const exec = promisify(execFile);
const VERSION = JSON.parse(await fs.readFile(path.join(REPO, 'package.json'), 'utf8')).version;
if (!Number.isInteger(PORT) || PORT < 0 || PORT > 65535) throw new Error('PPT_DEMO_PORT 应为 0–65535 的整数。');

const defaultPlan = JSON.parse(await fs.readFile(path.join(REPO, 'demo-prototype', 'demo-plan.json'), 'utf8'));
const defaultDraft = planToDraft(defaultPlan);
const sample = await readTemplate(SAMPLE_FILE, SAMPLE_NAME);
const files = new Map([[sample.sha, sample]]);
let current = sample;
const readJson = async (file, fallback = null) => {
  try { return JSON.parse(await fs.readFile(file, 'utf8')); }
  catch (error) { if (error.code === 'ENOENT') return fallback; throw error; }
};
const listFiles = directory => fs.readdir(directory).catch(error => {
  if (error.code === 'ENOENT') return [];
  throw error;
});
const workflow = { ...initialWorkflow(defaultDraft), ...await readJson(STATE) };
// Older releases never stored the original PPTX. Keep their edits until that exact template is uploaded.
if (workflow.version === 2) {
  workflow.version = 3;
  workflow.legacyEdits = { pageEdits: workflow.pageEdits, pageReviews: workflow.pageReviews, policySelections: workflow.policySelections, lastDraft: workflow.lastDraft };
  workflow.templateConfirmed = false;
  workflow.pageEdits = {};
  workflow.pageReviews = {};
  workflow.policySelections = {};
}
if (workflow.templateKey && workflow.templateKey !== sample.sha) {
  try {
    if (!/^[a-f0-9]{64}$/.test(workflow.templateKey)) throw new Error('模板标识无效。');
    const meta = await readJson(path.join(TEMPLATE_DIR, workflow.templateKey + '.json'));
    current = await readTemplate(path.join(TEMPLATE_DIR, workflow.templateKey + '.pptx'), meta?.name || '本地模板.pptx');
    if (current.sha !== workflow.templateKey) throw new Error('模板文件已变化。');
    files.set(current.sha, current);
  } catch (error) {
    console.warn('未能恢复上次的模板：' + error.message);
    workflow.templateConfirmed = false;
    current = sample;
  }
}
let writeQueue = Promise.resolve();
function saveWorkflow() {
  const payload = JSON.stringify(workflow, null, 2);
  writeQueue = writeQueue.catch(() => {}).then(async () => {
    await fs.mkdir(DATA, { recursive: true });
    const temporary = STATE + '.' + randomUUID() + '.tmp';
    await fs.writeFile(temporary, payload, { flag: 'wx' });
    await fs.rename(temporary, STATE);
  });
  return writeQueue;
}

const assets = new Map(columns.map(c => [c.asset, { key: c.asset, file: path.join(ASSET_DIR, c.asset), name: c.name + '示意图', url: '/assets/' + c.asset, builtin: true }]));
for (const key of await listFiles(UPLOAD_DIR)) {
  if (!/^[a-f0-9-]{36}\.png$/.test(key)) continue;
  const meta = await readJson(path.join(UPLOAD_DIR, key + '.json')).catch(() => null);
  assets.set(key, { key, file: path.join(UPLOAD_DIR, key), name: meta?.name || '已上传图片.png', url: '/uploads/' + key, builtin: false });
}
function publicTemplate() {
  return {
    name: current.name, sha256: current.sha, adapter: current.adapter,
    slideCount: current.pages.length, targetSlide: current.adapter === 'warehouse' ? 2 : null, size: current.size,
  };
}
function publicWorkflow() {
  const { legacyEdits, ...state } = workflow;
  return { ...state, approved: isApproved(workflow), planReady: canUsePlan(workflow), outlineIssues: outlineIssues(workflow.project), automatic: automaticEdits(current, workflow), fields: planFields(workflow) };
}
function bootstrap() {
  return {
    template: publicTemplate(), draft: workflow.lastDraft, columns, workflow: publicWorkflow(),
    pages: current.pages.map(({ part, slots, ...page }) => ({ ...page, slots: slots.map(({ metrics, ...slot }) => ({ ...slot, measured: Boolean(metrics) })) })),
    assets: [...assets.values()].map(({ file, ...asset }) => asset),
    latestJob: [...jobs.values()].filter(job => job.templateKey === current.sha).sort((a, b) => a.createdAt.localeCompare(b.createdAt)).at(-1) || null,
  };
}
function requirePlan() {
  if (!canUsePlan(workflow)) throw new Error('请先人工审核大纲并保存计划书草稿。');
}
function requireTemplate() {
  requirePlan();
  if (!workflow.templateConfirmed) throw new Error('请先确认模板。');
}
function setTemplate(deck) {
  if (workflow.templateKey !== deck.sha) {
    workflow.pageEdits = {};
    workflow.pageImages = {};
    workflow.pageReviews = {};
    workflow.policySelections = {};
    if (deck.sha === LEGACY_SHA && workflow.legacyEdits) {
      Object.assign(workflow, workflow.legacyEdits);
      delete workflow.legacyEdits;
    } else if (deck.adapter === 'warehouse') workflow.lastDraft = planToPptDraft(workflow, defaultDraft);
  }
  current = deck;
  workflow.templateKey = deck.sha;
  workflow.templateConfirmed = true;
}
function jobFingerprint(draft) {
  return workflowHash({
    template: current.sha, plan: workflow.plan?.markdown, approval: workflow.approval?.fingerprint,
    draft: current.adapter === 'warehouse' ? draft : null, texts: workflow.pageEdits,
    images: workflow.pageImages, notes: workflow.pageReviews, policies: workflow.policySelections,
  });
}

const jobs = new Map();
for (const id of await listFiles(OUTPUT)) {
  if (!/^[a-f0-9-]{36}$/.test(id)) continue;
  const state = await readJson(path.join(OUTPUT, id, 'job-state.json')).catch(() => null);
  if (state?.id === id && state.status === 'ready' && state.templateKey &&
      /^[^\\/:*?"<>|\x00-\x1f]+\.pptx$/.test(state.result?.filename || '')) jobs.set(id, state);
}
let activeJob;
let templateUploading = false;
const jobDirectory = id => path.join(OUTPUT, id);
const saveJob = job => fs.writeFile(path.join(jobDirectory(job.id), 'job-state.json'), JSON.stringify(job, null, 2));
function jobStatus(job, progress, message) {
  job.progress = progress;
  job.message = message;
  job.stages.push({ text: message });
}
function safeName(name) {
  const cleaned = String(name).replace(/[\\/:"*?<>|\x00-\x1f]/g, '_').replace(/^[. ]+|[. ]+$/g, '');
  return [...cleaned].slice(0, 100).join('') || '项目';
}
function startJob(checked, draft) {
  const job = {
    id: randomUUID(), templateKey: current.sha, fingerprint: jobFingerprint(draft), status: 'running',
    createdAt: new Date().toISOString(), progress: 3, message: '正在准备 PPTX 副本', stages: [],
    pageReviews: {}, review: { confirmed: false },
  };
  const editPlan = structuredClone(checked.edits), deck = current;
  const filename = safeName(workflow.project.name) + '_演示副本.pptx';
  jobs.set(job.id, job);
  activeJob = job;

  async function generate() {
    const folder = jobDirectory(job.id);
    const candidate = path.join(folder, 'candidate.tmp');
    try {
      await fs.mkdir(folder, { recursive: true });
      const preview = {}, pageImages = {};
      for (const [page, edit] of Object.entries(editPlan)) {
        preview[page] = Object.fromEntries(edit.texts.map(t => [t.id, t.paragraphs.map(plainText)]));
        pageImages[page] = {};
        for (const picture of edit.pictures) {
          const asset = assets.get(picture.key);
          picture.bytes = await fs.readFile(asset.file);
          pageImages[page][picture.id] = asset.url;
        }
      }
      jobStatus(job, 25, '替换选定的文字、图片与备注');
      const proof = await exportPresentation({ deck, edits: editPlan, destination: candidate });
      jobStatus(job, 84, '已回读文字并核对原模板哈希');
      await fs.rename(candidate, path.join(folder, filename));
      const report = { ...proof, template: deck.name, totalSlides: deck.pages.length, checkedAt: new Date().toISOString(), nativeWpsReview: { confirmed: false } };
      await fs.writeFile(path.join(folder, '验证报告.json'), JSON.stringify(report, null, 2));
      job.result = {
        filename, downloadUrl: '/api/jobs/' + job.id + '/download', reportUrl: '/api/jobs/' + job.id + '/report',
        editedSlides: proof.editedSlides, totalSlides: deck.pages.length, preview, pageImages,
        checks: [
          { status: 'pass', title: '模板保护', detail: `原模板哈希一致，${proof.unchangedParts} 个原有部件按原字节写入副本。` },
          { status: 'pass', title: '内容回读', detail: `核对了 ${proof.textsUpdated} 个文字框，替换了 ${proof.picturesUpdated} 张图片。` },
          { status: 'warning', title: '容量估算', detail: '字体和换行以 WPS 实际显示为准。表格、图表及动画沿用原件。' },
          { status: 'pending', title: 'WPS 视觉复核', detail: '下载副本，在 WPS 检查版面与备注，再记录复核结果。' },
        ],
      };
      jobStatus(job, 100, 'PPTX 副本已生成');
      job.status = 'ready';
      await saveJob(job);
    } catch (error) {
      job.status = 'failed';
      job.error = error.message;
      await fs.unlink(candidate).catch(() => {});
    } finally { activeJob = undefined; }
  }
  void generate();
  return job;
}

const MIME = {
  '.html': 'text/html; charset=utf-8', '.css': 'text/css; charset=utf-8', '.js': 'text/javascript; charset=utf-8',
  '.mjs': 'text/javascript; charset=utf-8', '.json': 'application/json; charset=utf-8', '.png': 'image/png',
  '.pptx': 'application/vnd.openxmlformats-officedocument.presentationml.presentation', '.md': 'text/markdown; charset=utf-8',
};
function send(res, status, value) {
  res.writeHead(status, { 'Content-Type': MIME['.json'], 'Cache-Control': 'no-store' });
  res.end(JSON.stringify(value));
}
const disposition = name => `attachment; filename="yuanban${path.extname(name)}"; filename*=UTF-8''${encodeURIComponent(name)}`;
async function serveFile(res, file, filename) {
  const stat = await fs.stat(file);
  const headers = { 'Content-Type': MIME[path.extname(file)] || 'application/octet-stream', 'Content-Length': stat.size, 'Cache-Control': 'no-cache' };
  if (filename) headers['Content-Disposition'] = disposition(filename);
  res.writeHead(200, headers);
  await pipeline(createReadStream(file), res);
}
async function body(req, limit = 1024 * 1024) {
  const chunks = [];
  let size = 0;
  for await (const chunk of req) {
    size += chunk.length;
    if (size > limit) throw new Error('内容超过支持大小。');
    chunks.push(chunk);
  }
  return Buffer.concat(chunks);
}
const jsonBody = async req => JSON.parse((await body(req)).toString('utf8') || '{}');
function nameHeader(req, fallback) {
  return safeName(decodeURIComponent(req.headers['x-file-name'] || fallback));
}

async function uploadTemplate(req) {
  if (Number(req.headers['content-length']) > MAX_TEMPLATE_BYTES) throw new Error('模板不能超过 160 MiB。');
  await fs.mkdir(TEMPLATE_DIR, { recursive: true });
  const temporary = path.join(TEMPLATE_DIR, randomUUID() + '.tmp');
  const hash = createHash('sha256');
  let size = 0;
  const limiter = new Transform({
    transform(chunk, _, callback) {
      size += chunk.length;
      if (size > MAX_TEMPLATE_BYTES) callback(new Error('模板不能超过 160 MiB。'));
      else { hash.update(chunk); callback(null, chunk); }
    },
  });
  try {
    const originalName = decodeURIComponent(req.headers['x-file-name'] || '上传模板.pptx');
    if (!/\.pptx$/i.test(originalName)) throw new Error('请选择 .pptx 文件。');
    const name = safeName(originalName.slice(0, -5)) + '.pptx';
    await pipeline(req, limiter, createWriteStream(temporary, { flags: 'wx' }));
    const key = hash.digest('hex');
    const destination = path.join(TEMPLATE_DIR, key + '.pptx');
    const deck = await readTemplate(temporary, name, key);
    // An upload can take time. Recheck the workbench before changing its template.
    requirePlan();
    if (activeJob) throw new Error('请等当前 PPT 生成完成，再更换模板。');
    if (key === sample.sha) return sample;
    await fs.rename(temporary, destination);
    await fs.writeFile(path.join(TEMPLATE_DIR, key + '.json'), JSON.stringify({ name }));
    deck.file = destination;
    files.set(key, deck);
    return deck;
  } finally { await fs.unlink(temporary).catch(() => {}); }
}
async function uploadPng(req) {
  const bytes = await body(req, 20 * 1024 * 1024);
  if (bytes.length < 33 || bytes.subarray(0, 8).toString('hex') !== '89504e470d0a1a0a' || bytes.toString('ascii', 12, 16) !== 'IHDR') throw new Error('请选择 PNG 图片。');
  const width = bytes.readUInt32BE(16), height = bytes.readUInt32BE(20);
  if (!width || !height || width * height > 24000000) throw new Error('PNG 图片不能超过 2400 万像素。');
  PNG.sync.read(bytes);
  const key = randomUUID() + '.png', name = nameHeader(req, '上传图片.png');
  const file = path.join(UPLOAD_DIR, key);
  await fs.mkdir(UPLOAD_DIR, { recursive: true });
  await fs.writeFile(file, bytes, { flag: 'wx' });
  await fs.writeFile(file + '.json', JSON.stringify({ name }), { flag: 'wx' });
  const url = '/uploads/' + key;
  assets.set(key, { key, file, name, url, builtin: false });
  return { key, name, url, width, height };
}
function savePage(input) {
  const page = current.pages[Number(input.page) - 1];
  if (!page || (page.number === 2 && current.adapter === 'warehouse')) throw new Error('请选择可编辑的页面。');
  if (input.templateKey !== current.sha) throw new Error('当前模板已更换，请重新打开编辑。');
  const edits = {}, automatic = automaticEdits(current, workflow)[page.number] || {};
  for (const item of input.texts || []) {
    const slot = page.slots.find(candidate => candidate.id === item.id);
    if (!slot || !Array.isArray(item.paragraphs) || item.paragraphs.length !== slot.paragraphs.length) throw new Error('文字位置或段数与原模板不符。');
    const values = item.paragraphs.map(text => String(text ?? ''));
    if (values.some(text => text.length > 12000)) throw new Error('单段文字不能超过 12000 个字符。');
    const baseline = automatic[slot.id] || slot.paragraphs;
    if (values.some((text, i) => text !== baseline[i])) edits[slot.id] = values;
  }
  workflow.pageEdits[page.number] = edits;
  workflow.pageReviews[page.number] = { note: String(input.note || '').slice(0, 4000) };
}
async function serveMedia(res, key, index) {
  const deck = files.get(key), media = deck?.media[index];
  if (!media) return send(res, 404, { error: '图片不存在。' });
  const archive = await openArchive(deck.file);
  try {
    const bytes = await archive.bytes(media.part, 20 * 1024 * 1024);
    res.writeHead(200, { 'Content-Type': media.mime, 'Content-Length': bytes.length, 'Cache-Control': 'private, max-age=3600' });
    res.end(bytes);
  } finally { archive.close(); }
}
async function openWps(file) {
  const configured = process.env.PPT_DEMO_WPS_PATH;
  try {
    if (process.platform === 'darwin') await exec('open', ['-a', configured || 'wpsoffice', file]);
    else {
      const program = configured || (process.platform === 'win32' ? 'wpp.exe' : 'wpp');
      await new Promise((resolve, reject) => {
        const child = spawn(program, [file], { detached: true, stdio: 'ignore', windowsHide: true });
        child.once('error', reject);
        child.once('spawn', () => { child.unref(); resolve(); });
      });
    }
  } catch {
    throw new Error('未能找到 WPS 演示。请下载 PPTX，再用本机 WPS 打开。');
  }
}

function planDownload(res) {
  requirePlan();
  const selected = [...new Set(Object.values(workflow.policySelections).flat())].map(id => workflow.policyRecords[id]).filter(Boolean);
  const links = selected.map(p => `- ${p.title}（${p.issuer}，${p.published || '日期待核对'}）\n  ${p.url}\n  适用条件：${p.use}`);
  const data = workflow.plan.markdown + (links.length ? '\n\n## 已选政策资料\n\n' + links.join('\n\n') : '') + '\n';
  res.writeHead(200, { 'Content-Type': MIME['.md'], 'Content-Disposition': disposition(safeName(workflow.project.name) + '_计划书.md') });
  res.end(data);
}

let workbenchUrl;
const server = http.createServer(async (req, res) => {
  try {
    res.setHeader('X-Content-Type-Options', 'nosniff');
    res.setHeader('Cross-Origin-Resource-Policy', 'same-origin');
    const port = server.address().port;
    const hosts = [`127.0.0.1:${port}`, `localhost:${port}`];
    const localOrigin = req.headers.origin ? hosts.some(host => req.headers.origin === 'http://' + host) : true;
    if (!hosts.includes(req.headers.host) || !localOrigin) return send(res, 403, { error: '请从本机打开工作台。' });
    const { pathname, searchParams } = new URL(req.url, workbenchUrl);
    const get = req.method === 'GET', post = req.method === 'POST';
    if (get && pathname === '/api/health') return send(res, 200, { app: 'ppt-template-product-demo', status: 'ready', version: VERSION, url: workbenchUrl });
    if (get && pathname === '/api/bootstrap') return send(res, 200, bootstrap());
    if (get && pathname === '/api/workflow') return send(res, 200, publicWorkflow());

    if (post && pathname === '/api/workflow') {
      const input = await jsonBody(req);
      const project = normalizeProject(input.project);
      const changed = workflowHash(workflow.project) !== workflowHash(project);
      invalidateOutline(workflow, project);
      if (changed) delete workflow.legacyEdits;
    } else if (post && pathname === '/api/workflow/approve') {
      const input = await jsonBody(req), issues = outlineIssues(workflow.project);
      if (input.confirmed !== true || issues.length) return send(res, 422, { error: issues.join(' ') || '请人工确认当前大纲。' });
      workflow.approval = { fingerprint: workflowHash(workflow.project), approvedAt: new Date().toISOString(), note: String(input.note || '').slice(0, 2000) };
    } else if (post && pathname === '/api/workflow/plan') {
      if (!isApproved(workflow)) return send(res, 409, { error: '当前大纲尚未通过人工审核。' });
      const input = await jsonBody(req);
      const today = new Intl.DateTimeFormat('sv-SE').format(new Date());
      const markdown = input.generate === true ? makePlanDraft(workflow, today) : String(input.markdown || '').trim();
      if (!markdown) throw new Error('计划书草稿不能为空。');
      const revision = (workflow.plan?.revision || 0) + 1;
      const planFolder = path.join(OUTPUT, 'plan-drafts');
      await fs.mkdir(planFolder, { recursive: true });
      const draftFile = safeName(workflow.project.name) + '_v' + revision + '_' + randomUUID() + '.md';
      await fs.writeFile(path.join(planFolder, draftFile), markdown + '\n', { flag: 'wx' });
      workflow.plan = { markdown, revision, generatedAt: new Date().toISOString(), file: draftFile };
      if (current.adapter === 'warehouse') workflow.lastDraft = planToPptDraft(workflow, defaultDraft);
      workflow.templateConfirmed = false;
    } else if (get && pathname === '/api/workflow/plan') {
      return planDownload(res);
    } else if (post && pathname === '/api/template/choose') {
      requirePlan();
      const input = await jsonBody(req);
      if (activeJob || templateUploading) throw new Error('请等当前操作完成，再选择模板。');
      setTemplate(input.sample === true ? sample : current);
      await saveWorkflow();
      return send(res, 200, bootstrap());
    } else if (post && pathname === '/api/template') {
      requirePlan();
      if (activeJob || templateUploading) throw new Error('请等当前操作完成，再上传模板。');
      templateUploading = true;
      try {
        setTemplate(await uploadTemplate(req));
        await saveWorkflow();
        return send(res, 200, bootstrap());
      } finally { templateUploading = false; }
    } else if (post && pathname === '/api/page-edit') {
      requireTemplate();
      savePage(await jsonBody(req));
    } else if (post && pathname === '/api/page-picture') {
      requireTemplate();
      const input = await jsonBody(req), page = current.pages[Number(input.page) - 1];
      if (input.templateKey !== current.sha || !page?.pictures.some(p => p.id === input.id) || (input.key && !assets.has(input.key))) throw new Error('图片位置或文件与当前模板不匹配。');
      const images = workflow.pageImages[page.number] ??= {};
      if (input.key) images[input.id] = input.key;
      else delete images[input.id];
    } else if (post && pathname === '/api/images') {
      return send(res, 201, await uploadPng(req));
    } else if (post && pathname === '/api/policies/search') {
      requireTemplate();
      const input = await jsonBody(req), page = current.pages[Number(input.page) - 1];
      if (!page) throw new Error('页面不存在。');
      const context = workflow.project.problem + ' ' + workflow.project.solution + ' ' + page.slots.flatMap(s => s.paragraphs).join(' ').slice(0, 1000);
      const result = await searchPolicies(String(input.query || '仓储 物流 数字化'), context);
      for (const policy of result.results) workflow.policyRecords[policy.id] = policy;
      await saveWorkflow();
      return send(res, 200, result);
    } else if (post && pathname === '/api/policies/adopt') {
      requireTemplate();
      const input = await jsonBody(req), page = current.pages[Number(input.page) - 1];
      const record = workflow.policyRecords[input.id] || verifiedPolicies.find(p => p.id === input.id);
      if (input.templateKey !== current.sha || !page || !record || !nationalOfficialUrl(record.url)) throw new Error('页面或官方政策记录不可用。');
      workflow.policyRecords[record.id] = record;
      const selected = new Set(workflow.policySelections[page.number] || []);
      if (input.remove) selected.delete(record.id);
      else selected.add(record.id);
      workflow.policySelections[page.number] = [...selected];
    } else if (post && (pathname === '/api/check' || pathname === '/api/jobs')) {
      const input = await jsonBody(req);
      if (pathname === '/api/jobs') requireTemplate();
      const draft = current.adapter === 'warehouse' && workflow.templateConfirmed ? normalizeDraft(input) : workflow.lastDraft;
      const checked = buildEdits(current, workflow, assets, draft);
      if (pathname === '/api/check') {
        if (current.adapter === 'warehouse' && workflow.templateConfirmed && workflowHash(draft) !== workflowHash(workflow.lastDraft)) {
          workflow.lastDraft = draft;
          await saveWorkflow();
        }
        return send(res, 200, { issues: checked.issues, capacities: checked.capacities, canGenerate: checked.canGenerate, fingerprint: jobFingerprint(draft) });
      }
      if (!checked.canGenerate) return send(res, 422, { error: '请处理容量或内容问题后再生成。', issues: checked.issues });
      if (activeJob || templateUploading) return send(res, 409, { error: '请等当前操作完成，再生成 PPT。', job: activeJob });
      workflow.lastDraft = draft;
      const job = startJob(checked, draft);
      await saveWorkflow();
      return send(res, 202, job);
    } else {
      const requestedJob = /^\/api\/jobs\/([a-f0-9-]{36})(?:\/(download|report|review|wps))?$/.exec(pathname);
      if (requestedJob) {
        const job = jobs.get(requestedJob[1]), action = requestedJob[2];
        if (!job) return send(res, 404, { error: '生成任务不存在。' });
        if (get && !action) return send(res, 200, job);
        if (job.status !== 'ready') return send(res, 409, { error: 'PPT 尚未生成。' });
        const fresh = workflow.templateConfirmed && job.fingerprint === jobFingerprint(workflow.lastDraft);
        const pptx = path.join(jobDirectory(job.id), job.result.filename);
        const report = path.join(jobDirectory(job.id), '验证报告.json');
        if (get && action === 'download') {
          if (searchParams.get('reviewed') === '1' && (!fresh || !job.review.confirmed)) return send(res, 409, { error: '请复核当前版本的所有修改页。' });
          return await serveFile(res, pptx, job.result.filename);
        }
        if (get && action === 'report') return await serveFile(res, report, '验证报告.json');
        if (post && action === 'wps') {
          await openWps(pptx);
          return send(res, 200, { message: '已向 WPS 发出打开请求。' });
        }
        if (post && action === 'review') {
          const input = await jsonBody(req), page = Number(input.page);
          if (!fresh || !Number.isInteger(page) || page < 1 || page > job.result.totalSlides) return send(res, 409, { error: '内容已更新，请重新生成再复核。' });
          job.pageReviews[page] = { confirmed: input.confirmed === true, note: String(input.note || '').slice(0, 2000), recordedAt: new Date().toISOString() };
          job.review.confirmed = job.result.editedSlides.every(number => job.pageReviews[number]?.confirmed);
          const proof = await readJson(report);
          proof.nativeWpsReview = { ...job.review, pages: job.pageReviews };
          await fs.writeFile(report, JSON.stringify(proof, null, 2));
          await saveJob(job);
          return send(res, 200, job);
        }
      }
      const media = /^\/api\/templates\/([a-f0-9]{64})\/media\/(\d+)$/.exec(pathname);
      if (get && media) return await serveMedia(res, media[1], Number(media[2]));
      const image = /^\/(assets|uploads)\/([^/]+\.png)$/.exec(pathname);
      if (get && image && assets.get(image[2])?.url === pathname) return await serveFile(res, assets.get(image[2]).file);
      if (get && pathname === '/model.mjs') return await serveFile(res, path.join(ROOT, 'model.mjs'));
      if (get && pathname === '/sample.pptx') return await serveFile(res, SAMPLE_FILE, SAMPLE_NAME);
      const publicFiles = { '/': 'index.html', '/prototype/ppt-template': 'index.html', '/app.js': 'app.js', '/styles.css': 'styles.css' };
      if (get && publicFiles[pathname]) return await serveFile(res, path.join(ROOT, 'public', publicFiles[pathname]));
      return send(res, 404, { error: '没有这个页面或文件。' });
    }
    await saveWorkflow();
    return send(res, 200, publicWorkflow());
  } catch (error) {
    if (!res.headersSent && !res.destroyed) send(res, 400, { error: error.message });
    else res.end();
  }
});

async function openBrowser(url) {
  if (!OPEN_BROWSER) return;
  try {
    if (process.platform === 'darwin') await exec('open', [url]);
    else if (process.platform === 'win32') await exec('rundll32.exe', ['url.dll,FileProtocolHandler', url]);
    else await exec('xdg-open', [url]);
  } catch { console.log('请在浏览器打开 ' + url); }
}
server.on('error', async error => {
  if (error.code === 'EADDRINUSE') {
    try {
      const url = `http://127.0.0.1:${PORT}`;
      const response = await fetch(url + '/api/health', { signal: AbortSignal.timeout(1200), redirect: 'error' });
      const health = await response.json();
      if (response.ok && health.app === 'ppt-template-product-demo' && health.version === VERSION) {
        console.log('工作台已在运行 ' + url);
        await openBrowser(url);
        return;
      }
    } catch { /* Another program owns the port. */ }
    console.error(`端口 ${PORT} 已被占用。请关闭原工作台或设置 PPT_DEMO_PORT 后重试。`);
  } else console.error('工作台启动失败 ' + error.message);
  process.exitCode = 1;
});
server.listen(PORT, '127.0.0.1', () => {
  workbenchUrl = `http://127.0.0.1:${server.address().port}`;
  console.log('工作台已就绪：' + workbenchUrl);
  void openBrowser(workbenchUrl);
});
