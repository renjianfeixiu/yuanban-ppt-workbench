// Functional prototype: map a five-part outline into the existing page without changing its containers.
export const columns = [
  { id: '63', headingId: '71', pictureId: '1026', name: '应用场景', count: 2, emphasis: [1, 1], headingLimit: 7, asset: 'warehouse-concept.png' },
  { id: '64', headingId: '72', pictureId: '3', name: '核心功能', count: 2, emphasis: [1, 1], headingLimit: 7, asset: 'sensor-flow.png' },
  { id: '65', headingId: '73', pictureId: '4', name: '方案特点', count: 3, emphasis: [1, 0, 2], headingLimit: 6, asset: 'local-processing.png' },
  { id: '66', headingId: '74', pictureId: '7', name: '验证计划', count: 2, emphasis: [1, 1], headingLimit: 6, asset: 'review-dashboard.png' },
  { id: '67', headingId: '75', pictureId: '2', name: '推进安排', count: 2, emphasis: [1, 2], headingLimit: 7, asset: 'pilot-roadmap.png' },
];

const string = value => typeof value === 'string' ? value.replaceAll('\r\n', '\n').trim() : '';
export const plainText = paragraph => typeof paragraph === 'string' ? paragraph : paragraph.map(x => x.text).join('');
export const weightedLength = text => [...text].reduce((n, ch) => n + (/[\u2e80-\uffff]/u.test(ch) ? 1 : 0.65), 0);

export function planToDraft(plan) {
  const slot = id => plan.texts.find(x => x.id === id);
  return {
    topic: plan.topic,
    projectName: plainText(slot('15').paragraphs[0]),
    subtitle: plainText(slot('16').paragraphs[0]),
    pageTitle: plainText(slot('42').paragraphs[0]),
    sections: columns.map(c => ({
      id: c.id, heading: plainText(slot(c.headingId).paragraphs[0]),
      paragraphs: slot(c.id).paragraphs.map(p => ({ text: plainText(p), emphasis: typeof p === 'string' ? '' : p.filter(x => x.style === 'emphasis').map(x => x.text).join('、') })),
      assetKey: c.asset,
      imageDescription: plan.pictures.find(x => x.id === c.pictureId).alt,
    })),
  };
}

export function normalizeDraft(input) {
  if (!input || typeof input !== 'object') throw new Error('缺少项目内容。');
  const draft = {
    topic: string(input.topic), projectName: string(input.projectName), subtitle: string(input.subtitle), pageTitle: string(input.pageTitle),
    sections: columns.map(c => {
      const item = input.sections?.find(x => x.id === c.id);
      if (!item || item.paragraphs?.length !== c.count) throw new Error(`${c.name}需要 ${c.count} 段文字，才能保留原段落结构。`);
      return { id: c.id, heading: string(item.heading), paragraphs: item.paragraphs.map(p => ({ text: string(p.text), emphasis: string(p.emphasis) })), assetKey: string(item.assetKey), imageDescription: string(item.imageDescription) };
    }),
  };
  if (JSON.stringify(draft).length > 80000) throw new Error('这页内容过长，请先精简大纲。');
  return draft;
}

export function richParagraph(paragraph, limit) {
  const text = paragraph.text;
  const words = [...new Set(paragraph.emphasis.split(/[、,，]/).map(x => x.trim()).filter(Boolean))];
  if (words.length > limit) throw new Error(`这一段最多保留 ${limit} 处黄色重点。`);
  const found = words.map(word => ({ start: text.indexOf(word), word })).filter(x => x.start >= 0).sort((a, b) => a.start - b.start || b.word.length - a.word.length);
  const fragments = [];
  let cursor = 0;
  for (const item of found) {
    if (item.start < cursor) throw new Error('重点词存在重叠，请保留一个完整词组。');
    fragments.push({ style: 'normal', text: text.slice(cursor, item.start) });
    fragments.push({ style: 'emphasis', text: item.word });
    cursor = item.start + item.word.length;
  }
  fragments.push({ style: 'normal', text: text.slice(cursor) });
  return fragments;
}

export function draftToPlan(draft, assetFiles) {
  const texts = [
    { id: '42', paragraphs: [draft.pageTitle] }, { id: '15', paragraphs: [draft.projectName] }, { id: '16', paragraphs: [draft.subtitle] },
    ...columns.map(c => ({ id: c.headingId, paragraphs: [draft.sections.find(x => x.id === c.id).heading] })),
    ...columns.map(c => { const s = draft.sections.find(x => x.id === c.id); return { id: c.id, paragraphs: s.paragraphs.map((p, i) => richParagraph(p, c.emphasis[i])) }; }),
  ];
  return { topic: draft.topic, projectName: draft.projectName, slide: 2, texts, pictures: columns.map(c => { const s = draft.sections.find(x => x.id === c.id); return { id: c.pictureId, file: assetFiles[s.assetKey], alt: s.imageDescription || `${s.heading}概念示意` }; }) };
}

export function draftToOutline(draft) {
  return draft.sections.map(s => `## ${s.heading}\n${s.paragraphs.map(p => p.text).join('\n\n')}`).join('\n\n');
}

export function applyOutline(draft, outline) {
  const chunks = [];
  let current;
  for (const line of outline.replaceAll('\r\n', '\n').split('\n')) {
    const heading = /^\s*#{1,3}\s+(.+?)\s*$/.exec(line);
    if (heading) { current = { heading: heading[1], lines: [] }; chunks.push(current); }
    else if (current && line.trim()) current.lines.push(line.trim().replace(/^[-*]\s+/, ''));
    else if (!current && line.trim()) throw new Error('请用 ## 标题 开始每一栏，再逐段填写正文。');
  }
  if (chunks.length !== 5) throw new Error('这页有固定五栏，请提供五个 ## 标题。');
  const next = structuredClone(draft);
  chunks.forEach((chunk, index) => {
    const column = columns[index];
    if (chunk.lines.length !== column.count) throw new Error(`第 ${index + 1} 栏需要 ${column.count} 段，当前有 ${chunk.lines.length} 段。不会自动删除或补造文字。`);
    const target = next.sections[index];
    target.heading = chunk.heading;
    target.paragraphs = chunk.lines.map((text, pi) => ({ text, emphasis: target.paragraphs[pi].emphasis.split(/[、,，]/).filter(word => word.trim() && text.includes(word.trim())).join('、') }));
  });
  return next;
}
