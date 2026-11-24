import test from 'node:test';
import assert from 'node:assert/strict';
import { spawn } from 'node:child_process';
import { once } from 'node:events';
import fs from 'node:fs/promises';
import { tmpdir } from 'node:os';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { setTimeout as delay } from 'node:timers/promises';
import { openArchive, writeArchive } from '../pptx-archive.mjs';
import { readTemplate, sha256 } from '../pptx.mjs';
import { xmlTree } from '../pptx-xml.mjs';
import { P, applyEdits, child, drawable, textEdits } from '../../demo-prototype/demo-core.mjs';

const repo = fileURLToPath(new URL('../../', import.meta.url));
const sample = path.join(repo, 'examples', '样例模板.pptx');
const entry = path.join(repo, 'product-demo', 'server.mjs');

async function runWorkbench(directory, t) {
  const child = spawn(process.execPath, [entry], {
    cwd: directory,
    env: { ...process.env, PPT_DEMO_DATA_DIR: path.join(directory, 'data'), PPT_DEMO_PORT: '0' },
    stdio: ['ignore', 'pipe', 'pipe'],
  });
  const stopped = once(child, 'exit');
  let log = '';
  child.stdout.on('data', chunk => { log += chunk; });
  child.stderr.on('data', chunk => { log += chunk; });
  t.after(async () => {
    if (child.exitCode === null && !child.signalCode) child.kill();
    await stopped;
  });
  let base;
  for (let i = 0; i < 180; i++) {
    base = log.match(/http:\/\/127\.0\.0\.1:\d+/)?.[0];
    if (base || child.exitCode !== null) break;
    await delay(75);
  }
  assert.ok(base, log);
  return async (route, input, status = 200, fileName) => {
    const isFile = Buffer.isBuffer(input);
    const options = input === undefined ? {} : {
      method: 'POST',
      headers: isFile
        ? { 'Content-Type': 'application/octet-stream', 'X-File-Name': encodeURIComponent(fileName) }
        : { 'Content-Type': 'application/json' },
      body: isFile ? input : JSON.stringify(input),
    };
    const response = await fetch(base + route, { ...options, signal: AbortSignal.timeout(15000) });
    const isJson = response.headers.get('content-type')?.includes('application/json');
    const value = isJson ? await response.json() : Buffer.from(await response.arrayBuffer());
    assert.equal(response.status, status, isJson ? JSON.stringify(value) : route);
    return value;
  };
}

async function exportJob(request) {
  const started = await request('/api/jobs', {}, 202);
  for (let i = 0; i < 100; i++) {
    const job = await request('/api/jobs/' + started.id);
    assert.notEqual(job.status, 'failed', job.error);
    if (job.status === 'ready') return job;
    await delay(80);
  }
  assert.fail('生成没有完成');
}

async function alternativeTemplate(destination, source, deck) {
  const archive = await openArchive(source);
  try {
    // Simulate a different document: a distinct title and reordered slide references.
    const xml = await archive.xml(deck.presentationPart);
    const list = child(xmlTree(xml), P, 'sldIdLst');
    const [first, second] = list.children;
    const part = deck.pages[1].part;
    const slide = await archive.xml(part);
    const title = deck.pages[1].slots.find(slot => slot.name === '页面标题');
    const changedSlide = applyEdits(slide, textEdits(slide, drawable(xmlTree(slide), title.id), ['团队培训']));
    const reordered = applyEdits(xml, [
      { start: first.start, end: first.end, replacement: xml.slice(second.start, second.end) },
      { start: second.start, end: second.end, replacement: xml.slice(first.start, first.end) },
    ]);
    await writeArchive(archive, new Map([[deck.presentationPart, reordered], [part, changedSlide]]), destination);
  } finally { archive.close(); }
}

test('空白目录启动、样例生成、上传另一份模板并保留未更改部件', { timeout: 45000 }, async t => {
  const directory = await fs.mkdtemp(path.join(tmpdir(), 'yuanban-workbench-'));
  const request = await runWorkbench(directory, t);
  t.after(() => fs.rm(directory, { recursive: true, force: true }));
  assert.match((await request('/')).toString(), /计划书|PPT/);
  assert.equal((await request('/api/health')).version, '0.3.0');
  const starter = await request('/api/bootstrap');
  assert.equal(starter.template.slideCount, 3);
  await request('/api/workflow/plan', { generate: true }, 409);
  await request('/api/workflow/approve', { confirmed: true });
  await request('/api/workflow/plan', { generate: true });
  assert.match((await request('/api/workflow/plan')).toString(), /待补充/);
  const example = await request('/api/template/choose', { sample: true });
  assert.equal(example.workflow.templateConfirmed, true);
  assert.ok(Object.keys(example.workflow.automatic).length === 3);
  const sampleJob = await exportJob(request);
  const sampleOutput = path.join(directory, 'sample-output.pptx');
  await fs.writeFile(sampleOutput, await request(sampleJob.result.downloadUrl));
  const original = await readTemplate(sample, '样例模板');
  const rendered = await readTemplate(sampleOutput, '样例输出');
  assert.equal(rendered.pages.length, 3);
  assert.ok(rendered.pages[0].slots.some(s => s.paragraphs.includes('智慧仓储')));
  assert.equal(await sha256(sample), original.sha);

  const alternate = path.join(directory, '另一份模板.pptx');
  await alternativeTemplate(alternate, sample, original);
  await request('/api/template', Buffer.from('文件内容不正确'), 400, '假的.pptx');
  const uploaded = await request('/api/template', await fs.readFile(alternate), 200, '另一份模板.pptx');
  const templateKey = uploaded.template.sha256;
  assert.notEqual(templateKey, original.sha);
  assert.equal(uploaded.pages[0].number, 1);
  const title = uploaded.pages[0].slots.find(s => s.paragraphs.includes('团队培训'));
  assert.ok(title, '识别 PPTX 内部的真实页面顺序');
  const tooLong = { page: 1, templateKey, texts: [{ id: title.id, paragraphs: ['过长'.repeat(400)] }], note: '记录培训目标' };
  await request('/api/page-edit', tooLong);
  assert.equal((await request('/api/check', {})).canGenerate, false);
  await request('/api/jobs', {}, 422);
  await request('/api/page-edit', { ...tooLong, texts: [{ id: title.id, paragraphs: ['课程安排'] }] });
  await request('/api/policies/adopt', { page: 1, templateKey, id: 'smart-supply-2025' });

  const picture = uploaded.pages[1].pictures[0];
  const asset = await request('/api/images', await fs.readFile(path.join(repo, 'demo-prototype/assets/sensor-flow.png')), 201, '新配图.png');
  await request('/api/page-picture', { page: 2, templateKey, id: picture.id, key: asset.key });
  const job = await exportJob(request);
  assert.equal(job.result.totalSlides, 3);
  assert.equal(job.result.preview[1][title.id][0], '课程安排');
  const output = path.join(directory, 'output.pptx');
  await fs.writeFile(output, await request(job.result.downloadUrl));
  const report = await request(job.result.reportUrl);
  assert.equal(report.sourceSha256, await sha256(alternate));
  assert.equal(report.outputSha256, await sha256(output));
  assert.equal(report.picturesUpdated, 1);
  const source = await openArchive(alternate), final = await openArchive(output);
  try {
    for (const [part] of source.entries) {
      if (!report.changedParts.includes(part)) assert.deepEqual(await final.bytes(part), await source.bytes(part), part);
    }
    const noteParts = [...final.entries.keys()].filter(name => /notesSlides\/[^/]+\.xml$/.test(name));
    const notes = await Promise.all(noteParts.map(name => final.xml(name)));
    assert.ok(notes.some(xml => xml.includes('记录培训目标') && xml.includes('app.www.gov.cn')));
  } finally { source.close(); final.close(); }
  await request(job.result.downloadUrl + '?reviewed=1', undefined, 409);
  for (const page of job.result.editedSlides) await request('/api/jobs/' + job.id + '/review', { page, confirmed: true });
  assert.equal((await request(job.result.downloadUrl + '?reviewed=1')).length > 0, true);
  await request('/api/page-edit', { page: 1, templateKey, texts: [{ id: title.id, paragraphs: ['培训安排'] }] });
  await request(job.result.downloadUrl + '?reviewed=1', undefined, 409);
});
