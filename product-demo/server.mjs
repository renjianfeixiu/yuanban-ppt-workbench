// Local functional demo. Real PPTX generation, one supported template, no cloud service or database.
import http from 'node:http';
import fs from 'node:fs/promises';
import { createReadStream } from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { createRequire } from 'node:module';
import { createHash, randomUUID } from 'node:crypto';
import { spawn, execFile } from 'node:child_process';
import { promisify } from 'node:util';
import { columns, planToDraft, normalizeDraft, draftToPlan, weightedLength } from './model.mjs';
import { P, A, parseXml, child, descendants, drawable, estimateFit, textEdits } from '../demo-prototype/demo-core.mjs';
import { initialWorkflow, normalizeProject, outlineIssues, invalidateOutline, isApproved, canUsePlan, makePlanDraft, planToPptDraft, workflowHash } from './workflow.mjs';
import { searchPolicies, verifiedPolicies, nationalOfficialUrl } from './policies.mjs';

const ROOT = path.dirname(fileURLToPath(import.meta.url));
const STUDY = path.dirname(ROOT);
const ENGINE = path.join(STUDY, 'demo-prototype');
const MODULES = '/Users/wangdajin/.cache/codex-runtimes/codex-primary-runtime/dependencies/node/node_modules';
const require = createRequire(path.join(MODULES, '__product_demo__.js'));
const JSZip = require('jszip');
const sax = require('sax');
const sharp = require('sharp');
const exec = promisify(execFile);
const PORT = Number(process.env.PPT_DEMO_PORT || 5186);
const WORKBENCH_URL = `http://127.0.0.1:${PORT}/prototype/ppt-template`;
const OPEN_BROWSER = process.argv.includes('--open-browser') && process.env.PPT_DEMO_NO_BROWSER !== '1';
if (process.argv.includes('--open-browser')) console.log('正在读取 PPT 模板，服务就绪后会打开工作台，请保留此终端窗口。');
const SOURCE = '/Users/wangdajin/我的文稿/大创/防盲区/ppt/平安货运：面向城乡货运场景的商用车智能安全保障系统（马赛克）.pptx';
const SOURCE_SHA = '9ae31fbda906baee21cb1b0fb5290442458e453b9787a7c4e7a30c6659ddfd52';
const digest = bytes => createHash('sha256').update(bytes).digest('hex');
const templateBytes = await fs.readFile(SOURCE);
if (digest(templateBytes) !== SOURCE_SHA) throw new Error('原模板已变化，需要重新标注后再启动。');
const sourceZip = await JSZip.loadAsync(templateBytes);
const sourceXml = await sourceZip.file('ppt/slides/slide2.xml').async('string');
const sourceRoot = parseXml(sourceXml, sax);
const defaultPlan = JSON.parse(await fs.readFile(path.join(ENGINE, 'demo-plan.json'), 'utf8'));
const defaultDraft = planToDraft(defaultPlan);
const workflowPath = path.join(ROOT, '.work', 'workflow.json');
let workflow;
try { workflow = JSON.parse(await fs.readFile(workflowPath, 'utf8')); } catch { workflow = initialWorkflow(defaultDraft); }
const saveWorkflow = async () => { await fs.mkdir(path.dirname(workflowPath), { recursive: true }); await fs.writeFile(workflowPath, JSON.stringify(workflow, null, 2)); };
const pageModels = [];
for (let number = 1; number <= 22; number++) {
  const xml = await sourceZip.file(`ppt/slides/slide${number}.xml`).async('string');
  const root = parseXml(xml, sax);
  const slots = [];
  for (const shape of descendants(root, P, 'sp')) {
    const props = child(child(shape, P, 'nvSpPr'), P, 'cNvPr');
    const body = child(shape, P, 'txBody');
    if (!props || !body) continue;
    const paragraphs = body.children.filter(n => n.uri === A && n.local === 'p').map(p => descendants(p, A, 't').map(t => t.text).join(''));
    if (!paragraphs.some(x => x.trim())) continue;
    try { textEdits(xml, shape, paragraphs); } catch { continue; }
    slots.push({ id: props.attrs.id, name: props.attrs.name, paragraphs, budgets: paragraphs.map(weightedLength) });
  }
  pageModels.push({ number, slots, originalContent: true });
}
function publicWorkflow() { return { ...workflow, approved: isApproved(workflow), planReady: canUsePlan(workflow), canGenerate: canUsePlan(workflow) && workflow.templateConfirmed, outlineIssues: outlineIssues(workflow.project) }; }
function requirePlan() { if (!canUsePlan(workflow)) throw new Error('请先人工审核大纲并生成当前计划书草稿。'); }
function jobHash(draft) { return workflowHash({ draft, pageEdits: workflow.pageEdits, policies: workflow.policySelections, notes: Object.fromEntries(Object.entries(workflow.pageReviews).map(([k,v]) => [k,v.note || ''])), planRevision: workflow.plan?.revision }); }
function extraPages() {
  return pageModels.filter(page => page.number === 2 || workflow.pageEdits[page.number] || workflow.policySelections[page.number]?.length).map(page => ({
    number: page.number,
    texts: page.number === 2 ? [] : Object.entries(workflow.pageEdits[page.number] || {}).map(([id, paragraphs]) => ({ id, paragraphs })),
    notes: [page.number === 2 ? `${workflow.project.topic}\n依据已审核大纲的计划书草稿第 ${workflow.plan.revision} 版整理。\n事实与阶段：${workflow.project.facts || '待补充'}\n配图为概念示意，不作为现场实测证据。其他原模板页面仍需逐页核对。` : '', workflow.pageReviews[page.number]?.note || '', ...(workflow.policySelections[page.number] || []).map(id => {
      const p = workflow.policyRecords[id]; return p ? `${p.title}\n来源：${p.issuer}；网页发布：${p.published || '待核对'}\n${p.url}\n用途及条件：${p.use}` : '';
    })].filter(Boolean).join('\n\n'),
  }));
}
const assets = new Map(columns.map(c => [c.asset, { path: path.join(ENGINE, 'assets', c.asset), file: c.asset, name: c.name + '示意图', url: `/assets/${c.asset}`, builtin: true }]));
const uploadDir = path.join(ROOT, '.work', 'uploads');
for (const file of await fs.readdir(uploadDir).catch(() => [])) {
  if (!/^[a-f0-9-]{36}\.png$/.test(file)) continue;
  const meta = await fs.readFile(path.join(uploadDir, file + '.json'), 'utf8').then(JSON.parse).catch(() => ({}));
  assets.set(file, { path: path.join(uploadDir, file), file, name: meta.name || '已上传配图.png', url: `/uploads/${file}`, builtin: false });
}
const jobs = new Map();
for (const id of await fs.readdir(path.join(ROOT, 'output')).catch(() => [])) {
  if (!/^[a-f0-9-]{36}$/.test(id)) continue;
  try {
    const output = path.join(ROOT, 'output', id);
    const state = JSON.parse(await fs.readFile(path.join(output, 'job-state.json'), 'utf8'));
    if (state.id !== id || state.status !== 'ready' || !/^验证报告_v\d+\.json$/.test(state.reportFile)) continue;
    const reportPath = path.join(output, state.reportFile);
    const report = JSON.parse(await fs.readFile(reportPath, 'utf8'));
    if (path.dirname(report.outputPath) !== output) continue;
    await fs.access(report.outputPath);
    jobs.set(id, { ...state, output, build: path.join(ROOT, '.work', 'jobs', id), reportPath, report });
  } catch { /* Only complete, locally saved candidates are resumed. */ }
}
const latestJob = () => [...jobs.values()].sort((a, b) => a.createdAt.localeCompare(b.createdAt)).at(-1);
async function saveReadyJob(job) {
  await fs.writeFile(path.join(job.output, 'job-state.json'), JSON.stringify({ ...publicJob(job), reportFile: path.basename(job.reportPath) }, null, 2));
}
const headings = [{ id: '15', field: 'projectName', name: '项目简称', limit: 4 }, { id: '16', field: 'subtitle', name: '副标题', limit: 10 }, { id: '42', field: 'pageTitle', name: '页面标题', limit: 4 }];
const MIME = { '.html': 'text/html; charset=utf-8', '.md': 'text/plain; charset=utf-8', '.css': 'text/css; charset=utf-8', '.js': 'text/javascript; charset=utf-8', '.mjs': 'text/javascript; charset=utf-8', '.png': 'image/png', '.json': 'application/json; charset=utf-8', '.pptx': 'application/vnd.openxmlformats-officedocument.presentationml.presentation' };
let activeJob;
let startingJob = false;

function checkDraft(input) {
  const draft = normalizeDraft(input);
  const issues = [];
  const capacities = [];
  if (!draft.topic) issues.push({ field: 'topic', type: 'error', message: '请填写项目题目。' });
  for (const h of headings) {
    const value = draft[h.field];
    if (!value) issues.push({ field: h.field, type: 'error', message: `请填写${h.name}。` });
    else if ((h.field === 'projectName' ? [...value].length : weightedLength(value)) > h.limit) issues.push({ field: h.field, type: 'error', message: `${h.name}超过原位置的保守预算（${h.limit} 个字），请精简。` });
  }
  for (const c of columns) {
    const section = draft.sections.find(s => s.id === c.id);
    if (!section.heading || weightedLength(section.heading) > c.headingLimit) issues.push({ field: c.id, type: 'error', message: `${c.name}的栏标题需控制在 ${c.headingLimit} 个汉字宽度以内。` });
    section.paragraphs.forEach((p, i) => {
      if (!p.text) issues.push({ field: c.id, paragraph: i, type: 'error', message: `${section.heading || c.name}第 ${i + 1} 段为空。` });
      const absent = p.emphasis.split(/[、,，]/).map(x => x.trim()).filter(x => x && !p.text.includes(x));
      if (absent.length) issues.push({ field: c.id, paragraph: i, type: 'warning', message: `重点词“${absent.join('、')}”不在正文中，可调整或清空。` });
    });
    const fit = estimateFit(drawable(sourceRoot, c.id), section.paragraphs.map(p => p.text));
    capacities.push({ id: c.id, heading: section.heading, ...fit, percent: Math.round(fit.usedHeightPt / fit.availableHeightPt * 100) });
    if (!fit.fits) issues.push({ field: c.id, type: 'error', message: `${section.heading || c.name}预计超出框高，请精简正文；原字号和原框不变。` });
    if (!assets.has(section.assetKey)) issues.push({ field: c.id, type: 'error', message: `${c.name}的图片不可用，请重新选择。` });
  }
  let plan;
  try {
    plan = draftToPlan(draft, Object.fromEntries([...assets].map(([key, asset]) => [key, asset.file])));
    for (const text of plan.texts) textEdits(sourceXml, drawable(sourceRoot, text.id), text.paragraphs);
  } catch (error) { issues.push({ field: 'content', type: 'error', message: error.message }); }
  return { draft, plan, issues, capacities, canGenerate: !issues.some(x => x.type === 'error'), estimated: true };
}

function publicJob(job) {
  return { id: job.id, fingerprint: job.fingerprint, status: job.status, progress: job.progress, stages: job.stages, message: job.message, error: job.error, createdAt: job.createdAt, completedAt: job.completedAt, review: job.review, pageReviews: job.pageReviews || {}, result: job.result, draft: job.draft };
}
function recordStage(job, data) {
  const labels = {
    'template-read': [12, '已读取原模板与第二页位置'],
    'oversized-copy-rejected': [18, '超限拦截规则已检查'],
    'short-copy-capacity-pass': [24, '当前文案通过容量估算'],
    'structural-protection-pass': [58, '固定版式与未编辑页面保持'],
    'package-and-geometry-pass': [72, 'PPT 文件与页面结构通过检查'],
    'full-deck-finalizer': [78, data.status === 'passed' ? '全套检查完成' : '全套图表检查未完成，保留具体原因'],
    rendered: [90, `已生成第 ${data.slide || 2} 页预览`],
    'demo-exported': [98, 'PPT 和检查记录已保存'],
  };
  const match = labels[data.stage];
  if (!match) return;
  job.progress = match[0]; job.message = match[1];
  job.stages.push({ name: data.stage, text: match[1], at: new Date().toISOString() });
}
async function beginJob(checked) {
  const id = randomUUID();
  const build = path.join(ROOT, '.work', 'jobs', id);
  const output = path.join(ROOT, 'output', id);
  const assetDir = path.join(build, 'assets');
  await fs.mkdir(assetDir, { recursive: true });
  for (const section of checked.draft.sections) {
    const asset = assets.get(section.assetKey);
    await fs.copyFile(asset.path, path.join(assetDir, asset.file));
  }
  const planPath = path.join(build, 'plan.json');
  checked.plan.extraPages = extraPages();
  checked.plan.planRevision = workflow.plan?.revision;
  await fs.writeFile(planPath, JSON.stringify(checked.plan, null, 2), { flag: 'wx' });
  workflow.lastDraft = checked.draft; await saveWorkflow();
  const job = { id, fingerprint: jobHash(checked.draft), draft: checked.draft, status: 'running', progress: 3, message: '准备模板副本', stages: [], createdAt: new Date().toISOString(), review: { confirmed: false, note: '' }, pageReviews: {}, output, build };
  jobs.set(id, job); activeJob = job;
  const childProcess = spawn(process.execPath, [path.join(ENGINE, 'run-demo.mjs'), '--plan', planPath, '--output-dir', output, '--build-dir', build, '--assets-dir', assetDir, '--render-scope', 'changed'], { cwd: STUDY, env: { ...process.env, RUNTIME_NODE_MODULES: MODULES }, stdio: ['ignore', 'pipe', 'pipe'] });
  let lines = '', stderr = '';
  childProcess.stdout.on('data', chunk => {
    lines += chunk;
    const pieces = lines.split('\n'); lines = pieces.pop();
    for (const line of pieces) {
      try { recordStage(job, JSON.parse(line)); } catch { /* Non-state library output does not become product copy. */ }
    }
  });
  childProcess.stderr.on('data', chunk => { stderr = (stderr + chunk).slice(-20000); });
  childProcess.on('error', error => { job.status = 'failed'; job.error = error.message; activeJob = undefined; });
  childProcess.on('close', async code => {
    if (activeJob?.id === job.id) activeJob = undefined;
    if (code !== 0) { job.status = 'failed'; job.error = stderr || `生成进程退出：${code}`; return; }
    try {
      const files = await fs.readdir(output);
      const name = files.find(x => /^验证报告_v\d+\.json$/.test(x));
      const report = JSON.parse(await fs.readFile(path.join(output, name), 'utf8'));
      job.report = report; job.reportPath = path.join(output, name);
      job.result = {
        filename: path.basename(report.outputPath), previewUrl: `/api/jobs/${id}/preview`, downloadUrl: `/api/jobs/${id}/download`, reportUrl: `/api/jobs/${id}/report`, pagePreviewUrls: Object.fromEntries(report.editedSlides.map(n => [n, `/api/jobs/${id}/preview?page=${n}`])),
        editedSlides: report.editedSlides, totalSlides: report.totalSlides,
        templateProtected: report.protectedTemplateSignatureMatched && report.sourceSha256Before === report.sourceSha256After,
        unchangedParts: report.unchangedOriginalParts, changedParts: report.changedOriginalParts,
        copyFit: report.shortCopyFit, fullDeckFinalizer: report.fullDeckFinalizer.status,
        checks: [
          { status: 'pass', title: '固定模板保护', detail: `原文件未变，${report.unchangedOriginalParts} 个原部件保持一致，母版、字号、位置与裁切保留；改动页 ${report.editedSlides.join('、')}。` },
          { status: 'pass', title: '文案容量估算', detail: '五栏正文通过原框预算；仍需要 WPS 核对实际字体显示。' },
          { status: 'pass', title: 'PPT 文件结构', detail: '文件完整性与页面几何检查通过，保持 22 页。' },
          { status: 'warning', title: '原模板字体', detail: '原模板曾在 WPS 中提示缺失字体，本次保留字体声明。' },
          { status: report.fullDeckFinalizer.status === 'passed' ? 'pass' : 'warning', title: '全套图表检查', detail: report.fullDeckFinalizer.blockClassification === 'validator-file-size-limit' ? '文件超过图表校验模块的 128 MiB 上限，完整检查未完成。' : report.fullDeckFinalizer.status === 'passed' ? '标准全套检查通过。' : '标准全套检查未完成，详见报告。' },
          { status: 'pending', title: 'WPS 视觉复核', detail: `逐页核对修改页 ${report.editedSlides.join('、')} 的标题、正文换行、图片裁切和引用备注。` },
        ],
      };
      job.status = 'ready'; job.progress = 100; job.message = '生成完成，等待 WPS 视觉复核'; job.completedAt = new Date().toISOString();
      await saveReadyJob(job);
    } catch (error) { job.status = 'failed'; job.error = '无法读取生成结果：' + error.message; }
  });
  return job;
}

const send = (res, status, value) => { res.writeHead(status, { 'Content-Type': 'application/json; charset=utf-8', 'Cache-Control': 'no-store' }); res.end(JSON.stringify(value)); };
async function body(req, limit = 1024 * 1024) {
  const chunks = []; let size = 0;
  for await (const chunk of req) { size += chunk.length; if (size > limit) throw new Error('文件或内容超过本 Demo 的支持大小。'); chunks.push(chunk); }
  return Buffer.concat(chunks);
}
const jsonBody = async req => JSON.parse((await body(req)).toString('utf8'));
async function serveFile(res, filename, download) {
  const stat = await fs.stat(filename);
  const headers = { 'Content-Type': MIME[path.extname(filename)] || 'application/octet-stream', 'Content-Length': stat.size, 'Cache-Control': 'no-cache' };
  if (download) headers['Content-Disposition'] = `attachment; filename="ppt-demo${path.extname(filename)}"; filename*=UTF-8''${encodeURIComponent(path.basename(filename))}`;
  res.writeHead(200, headers); createReadStream(filename).pipe(res);
}

const server = http.createServer(async (req, res) => {
  try {
    if (req.headers.origin && ![`http://127.0.0.1:${PORT}`, `http://localhost:${PORT}`].includes(req.headers.origin)) return send(res, 403, { error: '请在本地 Demo 页面操作。' });
    const url = new URL(req.url, `http://127.0.0.1:${PORT}`);
    const pathname = decodeURIComponent(url.pathname);
    if (req.method === 'GET' && pathname === '/api/health') return send(res, 200, {
      app: 'ppt-template-product-demo', status: 'ready', version: '0.2.1', pid: process.pid, url: WORKBENCH_URL,
    });
    if (req.method === 'GET' && pathname === '/api/bootstrap') return send(res, 200, {
      template: { name: path.basename(SOURCE), slideCount: 22, targetSlide: 2, ratio: '16:9', sizeMb: +(templateBytes.length / 1024 / 1024).toFixed(1), sha256: SOURCE_SHA },
      draft: workflow.lastDraft || defaultDraft, columns, workflow: publicWorkflow(), pages: pageModels, assets: [...assets].map(([key, a]) => ({ key, url: a.url, name: a.name })),
      latestJob: latestJob() ? publicJob(latestJob()) : null,
    });
    if (req.method === 'POST' && pathname === '/api/check') {
      const checked = checkDraft(await jsonBody(req));
      if(canUsePlan(workflow) && workflow.templateConfirmed){workflow.lastDraft=checked.draft;await saveWorkflow();}
      return send(res, 200, { issues: checked.issues, capacities: checked.capacities, canGenerate: checked.canGenerate, fingerprint: jobHash(checked.draft) });
    }
    if (pathname === '/api/workflow' && req.method === 'GET') return send(res, 200, publicWorkflow());
    if (pathname === '/api/workflow' && req.method === 'POST') {
      const input = await jsonBody(req);
      invalidateOutline(workflow, normalizeProject(input.project));
      await saveWorkflow(); return send(res, 200, publicWorkflow());
    }
    if (pathname === '/api/workflow/approve' && req.method === 'POST') {
      const input = await jsonBody(req); const issues = outlineIssues(workflow.project);
      if (issues.length) return send(res, 422, { error: issues.join(' ') });
      if (input.confirmed !== true) return send(res, 422, { error: '需要人工确认大纲后再继续。' });
      workflow.approval = { fingerprint: workflowHash(workflow.project), approvedAt: new Date().toISOString(), note: String(input.note || '').slice(0, 2000) };
      await saveWorkflow(); return send(res, 200, publicWorkflow());
    }
    if (pathname === '/api/workflow/plan' && req.method === 'POST') {
      if (!isApproved(workflow)) return send(res, 409, { error: '当前大纲未通过人工审核。' });
      const input = await jsonBody(req);
      const markdown = input.generate === true ? makePlanDraft(workflow, new Intl.DateTimeFormat('sv-SE', { timeZone: 'Asia/Shanghai' }).format(new Date())) : String(input.markdown || '').trim();
      if (!markdown) throw new Error('计划书草稿不能为空。');
      const revision = (workflow.plan?.revision || 0) + 1;
      const filename = path.join(ROOT, 'output', 'plan-drafts', `计划书草稿_${workflow.project.name}_v${revision}_${Date.now()}.md`);
      await fs.mkdir(path.dirname(filename), { recursive: true }); await fs.writeFile(filename, markdown, { flag: 'wx' });
      workflow.plan = { markdown, revision, path: filename, generatedAt: new Date().toISOString(), outlineFingerprint: workflow.approval.fingerprint };
      workflow.lastDraft = planToPptDraft(workflow, defaultDraft);
      workflow.templateConfirmed = false;
      await saveWorkflow(); return send(res, 200, publicWorkflow());
    }
    if (pathname === '/api/workflow/plan' && req.method === 'GET') {
      requirePlan();
      const selected=[...new Set(Object.values(workflow.policySelections).flat())].map(id=>workflow.policyRecords[id]).filter(Boolean);
      if(!selected.length)return serveFile(res,workflow.plan.path,true);
      const appendix='\n\n## 逐页审核后选取的政策参考\n\n'+selected.map(p=>`- ${p.title}（${p.issuer}，网页发布 ${p.published||'待核对'}）\n  ${p.url}\n  使用边界：${p.use}`).join('\n\n')+'\n';
      const content=workflow.plan.markdown+appendix;
      const file=path.join(path.dirname(workflow.plan.path),`计划书草稿_含政策引用_${workflowHash(content).slice(0,12)}.md`);
      try{await fs.writeFile(file,content,{flag:'wx'});}catch(error){if(error.code!=='EEXIST')throw error;}
      return serveFile(res,file,true);
    }
    if (pathname === '/api/template/choose' && req.method === 'POST') { requirePlan(); workflow.templateConfirmed = true; await saveWorkflow(); return send(res, 200, publicWorkflow()); }
    if (pathname === '/api/page-edit' && req.method === 'POST') {
      requirePlan(); const input = await jsonBody(req); const number = Number(input.page); const page = pageModels.find(p => p.number === number);
      if (!page || number === 2) throw new Error('第二页请在内容编辑区修改，其他页可按原位置修订。');
      const edits = {};
      for (const item of input.texts || []) {
        const slot = page.slots.find(s => s.id === item.id);
        if (!slot || item.paragraphs?.length !== slot.paragraphs.length) throw new Error('文字位置或段数与原模板不符。');
        const paragraphs = item.paragraphs.map(x => String(x));
        paragraphs.forEach((text, i) => { if (weightedLength(text) > slot.budgets[i]) throw new Error(`${slot.name}第 ${i + 1} 段超过原文长度预算，请精简。`); });
        if (JSON.stringify(paragraphs) !== JSON.stringify(slot.paragraphs)) edits[item.id] = paragraphs;
      }
      workflow.pageEdits[number] = edits;
      workflow.pageReviews[number] = { ...workflow.pageReviews[number], note: String(input.note || '').slice(0, 4000), confirmed: false };
      await saveWorkflow(); return send(res, 200, publicWorkflow());
    }
    if (pathname === '/api/policies/search' && req.method === 'POST') {
      const input = await jsonBody(req); const number = Number(input.page || 2); const page = pageModels.find(p => p.number === number);
      if (!page) throw new Error('页面不存在。');
      const results = await searchPolicies(String(input.query || '仓储 物流 数字化'), workflow.project.problem + ' ' + workflow.project.solution + ' ' + page.slots.map(s => s.paragraphs.join(' ')).join(' ').slice(0, 1000));
      results.results.forEach(p => { workflow.policyRecords[p.id] = p; });
      await saveWorkflow(); return send(res, 200, results);
    }
    if (pathname === '/api/policies/adopt' && req.method === 'POST') {
      requirePlan(); const input = await jsonBody(req); const page = Number(input.page); const record = workflow.policyRecords[input.id] || verifiedPolicies.find(p => p.id === input.id);
      if (!pageModels.some(p => p.number === page) || !record || !nationalOfficialUrl(record.url)) throw new Error('该页或官方政策记录不可用。');
      workflow.policyRecords[record.id] = record;
      const ids = new Set(workflow.policySelections[page] || []);
      if (input.remove) ids.delete(record.id); else ids.add(record.id);
      workflow.policySelections[page] = [...ids]; workflow.pageReviews[page] = { ...workflow.pageReviews[page], confirmed: false };
      await saveWorkflow(); return send(res, 200, publicWorkflow());
    }
    if (req.method === 'POST' && pathname === '/api/template') {
      requirePlan();
      const hash = createHash('sha256'); let size = 0;
      for await (const chunk of req) { size += chunk.length; if (size > 160 * 1024 * 1024) throw new Error('当前 Demo 支持最大 160 MiB 的模板文件。'); hash.update(chunk); }
      if (hash.digest('hex') !== SOURCE_SHA) return send(res, 422, { error: '当前 Demo 只支持已经标注的这份平安货运模板。其他模板需要先完成位置标注，文件未被写入。' });
      workflow.templateConfirmed = true; await saveWorkflow();
      return send(res, 200, { matched: true, size, workflow: publicWorkflow(), message: '文件与已标注模板一致，可以继续。' });
    }
    if (req.method === 'POST' && pathname === '/api/images') {
      const bytes = await body(req, 20 * 1024 * 1024);
      const metadata = await sharp(bytes, { limitInputPixels: 24000000 }).metadata();
      if (metadata.format !== 'png') return send(res, 422, { error: '最小 Demo 支持 PNG 配图，请选择 PNG 文件。' });
      const key = randomUUID() + '.png'; const imagePath = path.join(ROOT, '.work', 'uploads', key);
      await fs.mkdir(path.dirname(imagePath), { recursive: true }); await fs.writeFile(imagePath, bytes, { flag: 'wx' });
      const name = decodeURIComponent(req.headers['x-file-name'] || '新配图.png');
      const asset = { path: imagePath, file: key, name, url: `/uploads/${key}`, builtin: false };
      await fs.writeFile(imagePath + '.json', JSON.stringify({ name }), { flag: 'wx' });
      assets.set(key, asset);
      return send(res, 201, { key, url: asset.url, name, width: metadata.width, height: metadata.height });
    }
    if (req.method === 'POST' && pathname === '/api/jobs') {
      requirePlan(); if (!workflow.templateConfirmed) return send(res, 409, { error: '请在计划书草稿之后选择或上传 PPT 模板。' });
      if (startingJob || activeJob?.status === 'running') return send(res, 409, { error: '已有任务在生成，请等当前任务完成。', job: activeJob ? publicJob(activeJob) : undefined });
      const checked = checkDraft(await jsonBody(req));
      if (!checked.canGenerate) return send(res, 422, { error: '请先处理超限或缺少的内容。', issues: checked.issues });
      startingJob = true;
      try { return send(res, 202, publicJob(await beginJob(checked))); }
      finally { startingJob = false; }
    }
    const jobRoute = /^\/api\/jobs\/([a-f0-9-]{36})(?:\/(preview|download|report|review|wps))?$/.exec(pathname);
    if (jobRoute) {
      const job = jobs.get(jobRoute[1]); if (!job) return send(res, 404, { error: '这个任务不在当前服务会话中，请重新生成。' });
      const action = jobRoute[2];
      if (!action && req.method === 'GET') return send(res, 200, publicJob(job));
      if (job.status !== 'ready') return send(res, 409, { error: 'PPT 还没有生成完成。' });
      if (action === 'preview' && req.method === 'GET') {
        const number = Number(url.searchParams.get('page') || 2);
        return serveFile(res, job.report.editedSlides.includes(number) ? path.join(job.report.previews.fullDeck, `slide-${number}.png`) : path.join(STUDY, '.work', 'source-render', `slide-${number}.png`));
      }
      if (action === 'download' && req.method === 'GET') {
        if(url.searchParams.get('reviewed')==='1' && (!canUsePlan(workflow) || job.fingerprint!==jobHash(workflow.lastDraft) || !job.report.editedSlides.every(n=>job.pageReviews?.[n]?.confirmed)))return send(res,409,{error:'当前版本尚未完成所有修改页的 WPS 人工复核，或内容已变化。'});
        return serveFile(res, job.report.outputPath, true);
      }
      if (action === 'report' && req.method === 'GET') return serveFile(res, job.reportPath, true);
      if (action === 'wps' && req.method === 'POST') {
        await exec('/usr/bin/open', ['-a', '/Applications/wpsoffice.app', job.report.outputPath]);
        return send(res, 200, { requested: true, message: `已向 WPS 发出打开请求，请核对修改页 ${job.report.editedSlides.join('、')}。` });
      }
      if (action === 'review' && req.method === 'POST') {
        const input = await jsonBody(req);
        const page=Number(input.page||2);if(!pageModels.some(p=>p.number===page))throw new Error('复核页不存在。');
        job.pageReviews[page]={confirmed:input.confirmed===true,note:typeof input.note==='string'?input.note.slice(0,2000):'',recordedAt:new Date().toISOString()};
        job.review = { confirmed: job.report.editedSlides.every(n=>job.pageReviews[n]?.confirmed), note: job.pageReviews[page].note, recordedAt: new Date().toISOString() };
        job.report.nativeWpsReview = { status: job.review.confirmed ? 'manual-confirmation-recorded' : 'pending', ...job.review, pages:job.pageReviews, explanation: 'Per-page explicit reviewer records; opening WPS alone never marks review complete.' };
        await fs.writeFile(job.reportPath, JSON.stringify(job.report, null, 2));
        await saveReadyJob(job);
        return send(res, 200, publicJob(job));
      }
    }
    const sourcePreview = /^\/source\/slide-(\d+)\.png$/.exec(pathname);
    if (req.method === 'GET' && sourcePreview && +sourcePreview[1] >= 1 && +sourcePreview[1] <= 22) return serveFile(res, path.join(STUDY, '.work', 'source-render', `slide-${+sourcePreview[1]}.png`));
    const imageRoute = /^\/(assets|uploads)\/([^/]+\.png)$/.exec(pathname);
    if (req.method === 'GET' && imageRoute && assets.has(imageRoute[2])) return serveFile(res, assets.get(imageRoute[2]).path);
    if (req.method === 'GET' && pathname === '/model.mjs') return serveFile(res, path.join(ROOT, 'model.mjs'));
    const publicFiles = { '/': 'index.html', '/prototype/ppt-template': 'index.html', '/app.js': 'app.js', '/styles.css': 'styles.css' };
    if (req.method === 'GET' && publicFiles[pathname]) return serveFile(res, path.join(ROOT, 'public', publicFiles[pathname]));
    return send(res, 404, { error: '没有这个页面或文件。' });
  } catch (error) { if (!res.headersSent) send(res, 400, { error: error.message }); else res.end(); }
});
async function openWorkbench() {
  if (!OPEN_BROWSER) return;
  try { await exec('/usr/bin/open', [WORKBENCH_URL]); }
  catch { console.error(`浏览器未能自动打开，请手动打开：${WORKBENCH_URL}`); }
}
server.on('error', async error => {
  if (error.code === 'EADDRINUSE') {
    try {
      const response = await fetch(`http://127.0.0.1:${PORT}/api/health`, { redirect: 'error', signal: AbortSignal.timeout(1500) });
      const health = await response.json();
      if (response.ok && health.app === 'ppt-template-product-demo' && health.status === 'ready') {
        console.log(`工作台已经在运行：${WORKBENCH_URL}\n请保留原来的服务终端窗口。`);
        await openWorkbench();
        return;
      }
    } catch { /* Keep any existing listener untouched. */ }
    console.error(`端口 ${PORT} 已被其他服务占用，Demo 没有启动；请保留错误信息。`);
  } else console.error(`Demo 启动失败：${error.message}`);
  process.exitCode = 1;
});
server.listen(PORT, '127.0.0.1', () => {
  console.log(`工作台已就绪：${WORKBENCH_URL}\n按 Control+C 可以停止服务。`);
  openWorkbench();
});
