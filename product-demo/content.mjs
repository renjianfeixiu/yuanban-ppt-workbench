import { drawable, estimateFit, textEdits } from '../demo-prototype/demo-core.mjs';
import { columns, draftToPlan, normalizeDraft, weightedLength } from './model.mjs';
import { estimateSlot } from './pptx.mjs';
import { xmlTree } from './pptx-xml.mjs';

const marker = /\{\{([a-z][a-z0-9.]+)\}\}/g;
const singleLine = text => String(text ?? '').replace(/[\r\n\t]+/g, ' ');
export function planFields(workflow) {
  const { project, plan } = workflow;
  const fields = Object.fromEntries(['topic', 'name', 'audience', 'problem', 'solution', 'facts'].map(key => [key, project[key]]));
  const planTitle = /^#\s+(.+)$/m.exec(plan?.markdown || '');
  if (planTitle) fields.topic = planTitle[1];
  const sections = new Map();
  let active;
  for (const line of (plan?.markdown || '').split('\n')) {
    const heading = /^##\s+(\d+)[.、]\s+(.+)$/.exec(line);
    if (heading) {
      active = { title: heading[2], points: [] };
      sections.set(heading[1], active);
    } else if (/^#/.test(line)) active = undefined;
    else if (active && line.trim()) active.points.push(line.trim().replace(/^[-*]\s+/, ''));
  }
  project.chapters.forEach((chapter, i) => {
    const fromPlan = sections.get(String(i + 1));
    fields[`chapter.${i + 1}.title`] = fromPlan?.title || chapter.title;
    fields[`chapter.${i + 1}.points`] = fromPlan?.points.join(' ') || chapter.points;
  });
  return fields;
}

export function automaticEdits(deck, workflow) {
  if (!workflow.plan || deck.adapter === 'warehouse') return {};
  const fields = planFields(workflow), found = {};
  for (const page of deck.pages) {
    for (const slot of page.slots) {
      const values = slot.paragraphs.map(text => text.replace(marker, (matched, field) => Object.hasOwn(fields, field) ? singleLine(fields[field]) : matched));
      if (values.some((value, i) => value !== slot.paragraphs[i])) (found[page.number] ??= {})[slot.id] = values;
    }
  }
  return found;
}

export function checkLegacyDraft(input, deck, assets) {
  const draft = normalizeDraft(input);
  const issues = [], capacities = [];
  const headingLimits = [
    ['projectName', '项目简称', 4],
    ['subtitle', '副标题', 10],
    ['pageTitle', '页面标题', 4],
  ];
  const root = xmlTree(deck.legacySlideXml);
  for (const [field, name, limit] of headingLimits) {
    const value = draft[field];
    if (!value || (field === 'projectName' ? [...value].length : weightedLength(value)) > limit) {
      issues.push({ field, type: 'error', message: `${name}应为 1–${limit} 个汉字宽度。` });
    }
  }
  for (const column of columns) {
    const section = draft.sections.find(s => s.id === column.id);
    if (!section.heading || weightedLength(section.heading) > column.headingLimit) {
      issues.push({ field: column.id, type: 'error', message: `${column.name}的栏标题应为 1–${column.headingLimit} 个汉字宽度。` });
    }
    if (section.paragraphs.some(p => !p.text)) issues.push({ field: column.id, type: 'error', message: column.name + '包含空白段落。' });
    const fit = estimateFit(drawable(root, column.id), section.paragraphs.map(p => p.text));
    capacities.push({ id: column.id, heading: section.heading, ...fit, percent: Math.round(fit.usedHeightPt / fit.availableHeightPt * 100) });
    if (!fit.fits) issues.push({ field: column.id, type: 'error', message: `${section.heading}正文预计超出原框，请精简。` });
    if (!assets.has(section.assetKey)) issues.push({ field: column.id, type: 'error', message: `${column.name}的配图不可用。` });
  }
  let plan;
  try {
    plan = draftToPlan(draft, Object.fromEntries([...assets].map(([key]) => [key, key])));
    for (const text of plan.texts) textEdits(deck.legacySlideXml, drawable(root, text.id), text.paragraphs);
  } catch (error) { issues.push({ field: 'content', type: 'error', message: error.message }); }
  return { draft, plan, capacities, issues };
}

export function buildEdits(deck, workflow, assets, draft) {
  const automatic = automaticEdits(deck, workflow);
  const legacy = deck.adapter === 'warehouse' && workflow.templateConfirmed ? checkLegacyDraft(draft, deck, assets) : null;
  const capacities = legacy?.capacities || [], issues = legacy?.issues || [];
  const edits = {};
  if (!workflow.templateConfirmed) return { edits, automatic, capacities, issues, canGenerate: false };
  for (const page of deck.pages) {
    const number = page.number;
    const selected = { ...automatic[number], ...workflow.pageEdits[number] };
    const texts = [];
    if (legacy && number === 2) {
      if (legacy.plan) texts.push(...legacy.plan.texts.map(t => ({ ...t, rich: true })));
    } else {
      for (const [id, values] of Object.entries(selected)) {
        const slot = page.slots.find(item => item.id === id);
        if (!slot || !Array.isArray(values) || values.length !== slot.paragraphs.length) {
          issues.push({ page: number, field: id, type: 'error', message: `第 ${number} 页存在与当前模板不匹配的文字位置。` });
          continue;
        }
        const paragraphs = values.map(singleLine);
        if (paragraphs.every((value, i) => value === slot.paragraphs[i])) continue;
        const fit = estimateSlot(slot, paragraphs);
        capacities.push({ page: number, id, heading: slot.name, ...fit });
        if (!fit.fits) issues.push({ page: number, field: id, type: 'error', message: `第 ${number} 页「${slot.name}」超过预计容量，请精简文字。` });
        texts.push({ id, paragraphs });
      }
    }
    const pictures = [];
    if (legacy && number === 2 && legacy.plan) {
      pictures.push(...legacy.plan.pictures.map(item => ({ id: item.id, key: item.file })));
    } else {
      for (const [id, key] of Object.entries(workflow.pageImages[number] || {})) {
        if (!page.pictures.some(item => item.id === id) || !assets.has(key)) {
          issues.push({ page: number, field: id, type: 'error', message: `第 ${number} 页的配图不可用，请重新选择。` });
        } else pictures.push({ id, key });
      }
    }
    const refs = (workflow.policySelections[number] || []).map(id => workflow.policyRecords[id]).filter(Boolean);
    const notes = [
      legacy && number === 2 ? `${workflow.project.topic}\n依据已审核的大纲和计划书草稿整理。配图为概念示意。` : '',
      workflow.pageReviews[number]?.note,
      ...refs.map(p => `${p.title}\n${p.issuer}　${p.published || '日期待核对'}\n${p.url}\n适用条件：${p.use}`),
    ].filter(Boolean).join('\n\n');
    if (texts.length || pictures.length || notes) edits[number] = { texts, pictures, notes };
  }
  return { edits, automatic, capacities, issues, draft: legacy?.draft, canGenerate: !issues.some(i => i.type === 'error') };
}
