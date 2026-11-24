import { createHash } from 'node:crypto';

export const outlineDefaults = [
  ['项目概述', '说明项目定位、服务对象和拟解决的问题。'],
  ['背景与需求', '梳理实际场景、现有困难和需要核对的全国政策依据。'],
  ['解决方案', '说明感知、告警、人工复核和记录流程。'],
  ['实施与验证', '先进行小范围测试，记录误报、响应和适用条件。'],
  ['运行与商业模式', '讨论部署和维护方式，成本及收费依据待补充。'],
  ['团队与资源', '列出岗位分工、需要的设备和合作条件。'],
  ['风险与推进计划', '识别误报、数据处理和现场操作风险，分阶段推进。'],
  ['资料与政策引用', '列出实测记录、预算来源和已确认的官方材料。'],
];
export const workflowHash = value => createHash('sha256').update(JSON.stringify(value)).digest('hex');
export function initialWorkflow(draft) {
  return { version: 3, templateKey: null, pageImages: {}, project: { topic: draft.topic, name: draft.projectName, audience: '中小仓库管理者与巡检人员', problem: '通道阻塞、货架周边风险与告警后的人工复核', solution: '视觉与传感器辅助巡检，异常告警后由人员复核并保留记录', facts: '当前为概念方案，未提供实测结果、试点记录或经营数据。', policyScope: '全国官方政策', chapters: outlineDefaults.map(([title, points], i) => ({ id: String(i + 1), title, points })) }, approval: null, plan: null, templateConfirmed: false, pageEdits: {}, pageReviews: {}, policySelections: {}, policyRecords: {}, lastDraft: draft };
}
export function normalizeProject(input) {
  const project = {};
  for (const key of ['topic', 'name', 'audience', 'problem', 'solution', 'facts']) project[key] = String(input?.[key] ?? '').trim().slice(0, 12000);
  project.policyScope = '全国官方政策';
  project.chapters = (Array.isArray(input?.chapters) ? input.chapters : []).slice(0, 30).map((c, i) => ({ id: String(i + 1), title: String(c?.title || '').trim().slice(0, 100), points: String(c?.points || '').trim().slice(0, 8000) }));
  return project;
}
export function outlineIssues(project) {
  const issues = [];
  if (!project.topic) issues.push('项目题目为空。');
  if (!project.name || [...project.name].length > 60) issues.push('项目简称需为 1–60 个字符。');
  if (/[\\/:*?"<>|\x00-\x1f]/.test(project.name)) issues.push('项目简称不能包含路径符号或控制字符。');
  if (!project.audience || !project.problem || !project.solution) issues.push('请补充服务对象、问题和解决方案。');
  if (project.chapters.length < 3) issues.push('大纲至少需要三个章节。');
  project.chapters.forEach((c, i) => { if (!c.title || !c.points) issues.push(`第 ${i + 1} 章的标题或要点为空。`); });
  return issues;
}
export function invalidateOutline(workflow, project) {
  if (workflowHash(workflow.project) !== workflowHash(project)) {
    workflow.approval = null; workflow.plan = null; workflow.templateConfirmed = false;
    workflow.pageEdits = {}; workflow.pageImages = {}; workflow.pageReviews = {}; workflow.policySelections = {};
  }
  workflow.project = project;
  return workflow;
}
export function isApproved(workflow) { return workflow.approval?.fingerprint === workflowHash(workflow.project); }
export function canUsePlan(workflow) { return isApproved(workflow) && !!workflow.plan?.markdown; }

export function makePlanDraft(workflow, date) {
  if (!isApproved(workflow)) throw new Error('请先人工审核并确认当前大纲。');
  const p = workflow.project;
  const context = [
    `本项目拟面向${p.audience}，围绕${p.problem}开展方案设计。`,
    `拟采用的基本方式为：${p.solution}。这些表述属于计划方案，实际效果需要通过测试记录核对。`,
  ];
  const paragraphs = p.chapters.map((c, i) => {
    const points = c.points.split('\n').map(x => x.replace(/^[-*]\s*/, '').trim()).filter(Boolean);
    const body = points.map(point => point.endsWith('。') ? point : point + '。').join('\n\n');
    let next = '后续需明确本章所需资料、负责人和完成条件，形成可核对的记录。';
    if (/验证|实施|技术|方案/.test(c.title)) next = '实施时应先明确场景、输入条件和人工复核方式，再按同一条件记录测试结果。指标阈值和效果数字为【待补充】，不能把预期写成实测。';
    if (/商业|预算|资金|运行/.test(c.title)) next = '预算、单价、收入、采购或合作情况均应由真实资料支持，当前具体金额及依据为【待补充】。方案可以先讨论成本构成和运行方式。';
    if (/政策|背景|需求/.test(c.title)) next = '相关政策材料应保留官方原文、发布单位、日期和适用条件；在逐页审核中检索后再确认引用，不能据此推断项目已获资助。';
    if (/团队|资源/.test(c.title)) next = '成员信息、资历、分工与设备清单为【待补充】，仅列出需要承担的任务，不虚构奖项、合同或合作单位。';
    return `## ${i + 1}. ${c.title}\n\n${body}\n\n${next}`;
  });
  return `# ${p.topic}\n\n计划书草稿 · ${date}\n\n依据已人工确认的大纲整理。未提供的数据、资历和成果以【待补充】标记，草稿可继续编辑。\n\n## 项目基本设定\n\n${context.join('\n\n')}\n\n已提供的事实与阶段说明：${p.facts || '【待补充】'}\n\n${paragraphs.join('\n\n')}\n\n## 待补资料清单\n\n- 实际调研与测试记录、条件和样本说明。\n- 设备、实施及维护预算的计算依据。\n- 团队身份、能力和分工材料。\n- 合作、采购、知识产权及资质原件（如适用）。\n- 确认适用的全国官方政策及引用位置。\n`;
}
export function planToPptDraft(workflow, draft) {
  if (!canUsePlan(workflow)) throw new Error('请先生成当前大纲的计划书草稿。');
  const p = workflow.project;
  const next = structuredClone(draft);
  next.topic = p.topic; next.projectName = p.name;
  next.subtitle = '项目概念方案';
  const sections = [];
  let current;
  for (const raw of workflow.plan.markdown.split('\n')) {
    const heading = /^##\s+(?:\d+[.、]\s*)?(.+)$/.exec(raw.trim());
    if (heading) { current = { title: heading[1], lines: [] }; sections.push(current); }
    else if (current && raw.trim() && !/^[#-]/.test(raw.trim())) current.lines.push(raw.trim());
  }
  const pick = (chapter, pattern, fallback) => sections.filter(s => chapter.test(s.title)).flatMap(s => s.lines).find(x => pattern.test(x) && [...x].length <= 65) || fallback;
  const values = [
    [`面向${p.audience}。`, `拟解决${p.problem}。`],
    [pick(/方案|功能|技术/, /感知|拟采用|基本方式|告警|功能/, `方案拟采用${p.solution}。`), '实施方式与适用条件需要人工核对。'],
    [pick(/方案|技术|部署/, /本地|模块|部署/, '先明确部署与使用条件。'), '实际效果需要测试记录支持。', '根据现场条件逐步完善方案。'],
    ['当前内容为计划书草稿，不代表实测成果。', pick(/验证|实施|测试/, /阈值|场景|测试|记录/, '拟记录测试条件、样本和实际结果。')],
    [pick(/风险|推进|计划/, /小范围|分阶段|推进|风险/, '先进行小范围测试，再评估应用条件。'), '结合用户反馈和材料核对，分阶段推进。'],
  ];
  next.sections.forEach((s, i) => { s.paragraphs = values[i].map(text => ({ text, emphasis: '' })); });
  return next;
}
