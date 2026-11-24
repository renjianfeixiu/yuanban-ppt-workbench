import fs from 'node:fs/promises';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { readTemplate, exportPresentation } from '../product-demo/pptx.mjs';
import { automaticEdits } from '../product-demo/content.mjs';
import { draftToPlan, planToDraft, columns } from '../product-demo/model.mjs';
import { initialWorkflow } from '../product-demo/workflow.mjs';

const here = path.dirname(fileURLToPath(import.meta.url));
const repo = path.dirname(here);
function option(name, fallback) {
  const index = process.argv.indexOf(name);
  if (index === -1) return fallback;
  if (!process.argv[index + 1] || process.argv[index + 1].startsWith('--')) throw new Error(name + ' 缺少参数。');
  return process.argv[index + 1];
}

const template = path.resolve(option('--template', path.join(repo, 'examples', '样例模板.pptx')));
const deck = await readTemplate(template, path.basename(template));
const demoPlan = JSON.parse(await fs.readFile(path.join(here, 'demo-plan.json'), 'utf8'));
const draft = planToDraft(demoPlan);
const output = path.resolve(option('--output-dir', path.join(here, 'output')));
const edits = {};
if (deck.adapter === 'warehouse') {
  const planFile = option('--plan');
  const plan = planFile ? JSON.parse(await fs.readFile(planFile, 'utf8'))
    : draftToPlan(draft, Object.fromEntries(columns.map(column => [column.asset, column.asset])));
  const assets = path.resolve(option('--assets-dir', path.join(here, 'assets')));
  edits[2] = {
    texts: plan.texts.map(text => ({ ...text, rich: true })),
    pictures: await Promise.all(plan.pictures.map(async pic => ({ id: pic.id, bytes: await fs.readFile(path.join(assets, path.basename(pic.file))) }))),
    notes: '智慧仓储概念示例。配图为方案示意。',
  };
} else {
  const workflow = initialWorkflow(draft);
  // Command-line demo uses the bundled example, without saving an approval in the workbench.
  workflow.plan = { markdown: '' };
  for (const [number, fields] of Object.entries(automaticEdits(deck, workflow))) {
    edits[number] = { texts: Object.entries(fields).map(([id, paragraphs]) => ({ id, paragraphs })), pictures: [], notes: '' };
  }
}
await fs.mkdir(output, { recursive: true });
const destination = path.join(output, '概念方案_' + Date.now() + '.pptx');
const proof = await exportPresentation({ deck, edits, destination });
console.log(JSON.stringify({ destination, ...proof }, null, 2));
