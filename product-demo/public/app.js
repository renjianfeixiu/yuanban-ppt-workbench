import { columns, normalizeDraft, draftToOutline, applyOutline } from '/model.mjs';

const app = document.querySelector('#app');
const escape = value => String(value ?? '').replaceAll('&', '&amp;').replaceAll('<', '&lt;').replaceAll('>', '&gt;').replaceAll('"', '&quot;').replaceAll("'", '&#39;');
const icons = {
  arrow: '<path d="M5 12h14m-5-5 5 5-5 5"/>', back: '<path d="M19 12H5m5-5-5 5 5 5"/>',
  check: '<path d="m5 12 4 4 10-10"/>', lock: '<rect x="5" y="10" width="14" height="11" rx="2"/><path d="M8 10V7a4 4 0 0 1 8 0v3m-4 5v2"/>',
  upload: '<path d="M12 16V3m-5 5 5-5 5 5M4 15v5h16v-5"/>', download: '<path d="M12 3v13m-5-5 5 5 5-5M4 17v4h16v-4"/>',
  edit: '<path d="m16 3 5 5-12 12-6 1 1-6Z"/><path d="m14 5 5 5"/>', alert: '<path d="m12 3 10 18H2Z"/><path d="M12 9v4m0 4h.01"/>',
  eye: '<path d="M2 12s4-7 10-7 10 7 10 7-4 7-10 7S2 12 2 12Z"/><circle cx="12" cy="12" r="3"/>',
  refresh: '<path d="M20 7v5h-5M4 17v-5h5M19 12a7 7 0 0 0-12-5M5 12a7 7 0 0 0 12 5"/>',
  close: '<path d="m6 6 12 12M6 18 18 6"/>', image: '<rect x="3" y="3" width="18" height="18" rx="2"/><circle cx="8" cy="8" r="1.5"/><path d="m3 17 6-6 4 4 3-3 5 5"/>',
};
const icon = name => `<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.6" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true">${icons[name] || icons.check}</svg>`;
const button = (label, action, style = 'secondary', iconName, disabled = false, extra = '') => `<button class="btn ${style}" data-action="${action}" ${disabled ? 'disabled' : ''} ${extra}>${iconName ? icon(iconName) : ''}${label}</button>`;
const state = { step: 1, draft: null, template: null, assets: new Map(), selected: 0, sourceSlide: 2, view: 'generated', checking: false, checks: null, job: null, dirty: false, starting: false, modal: null, modalError: '', longBackup: null, reviewNote: '', uploading: false, workflow: null, outlineDirty: false, planText: '', planDirty: false, generationView: 'edit', pages: [], reviewPage: 2, policyQuery: '仓储 物流 数字化', policyResults: null, policyLoading: false, policySequence: 0, outlineConfirmed: false, outlineNote: '' };
let checkSequence = 0, checkTimer, pollTimer, toastTimer;
let renderedScreen = '';
const busy = () => state.starting || state.job?.status === 'running' || state.uploading;
const stale = () => !!state.job && (state.dirty || state.checks?.fingerprint !== state.job.fingerprint);
const section = () => state.draft.sections[state.selected];
const capacity = id => state.checks?.capacities?.find(x => x.id === id);
const errors = () => state.checks?.issues?.filter(x => x.type === 'error') || [];

async function api(url, options = {}) {
  const response = await fetch(url, options);
  const result = await response.json();
  if (!response.ok) { const error = new Error(result.error || '操作未完成。'); error.data = result; throw error; }
  return result;
}
const post = (url, value) => api(url, { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify(value) });
function toast(message) { const node = document.querySelector('#toast'); node.textContent = message; node.classList.add('visible'); clearTimeout(toastTimer); toastTimer = setTimeout(() => node.classList.remove('visible'), 4200); }

function render() {
  const screenKey = `${state.step}:${state.generationView}:${state.reviewPage}`;
  const screenChanged = renderedScreen !== screenKey;
  const active = document.activeElement;
  const focus = !screenChanged && active?.dataset?.focus ? { key: active.dataset.focus, start: active.selectionStart, end: active.selectionEnd } : null;
  app.innerHTML = `
    <header class="topbar"><div class="brand"><div class="brand-mark" aria-hidden="true"></div><div><strong>原版</strong><span>计划书与 PPT 工作台</span></div></div><div class="top-meta"><span class="local-dot">本地运行</span><span class="divider"></span><span>六步工作流 Demo</span></div></header>
    <div class="app-shell"><aside class="sidebar"><div class="side-label">项目制作流程</div><nav class="steps" aria-label="制作流程">${['设定大纲', '人工审核', '计划书草稿', '上传模板', '自动生成', '逐页审核'].map((name, i) => `<button class="step ${state.step === i + 1 ? 'active' : ''} ${state.step > i + 1 ? 'complete' : ''}" data-action="step" data-step="${i + 1}" aria-current="${state.step === i + 1 ? 'step' : 'false'}"><span class="step-number">${state.step > i + 1 ? icon('check') : String(i + 1).padStart(2, '0')}</span><span>${name}<small>${['明确题目与章节', '确认事实和内容安排', '整理成可编辑文稿', '在草稿后选择设计', '适配现有文字位置', '修改、政策资料与复核'][i]}</small></span></button>`).join('')}</nav><div class="side-foot"><strong><span class="mini-lock">${icon('lock')}</span>大纲先确认，设计后选择</strong><p>全国官方政策优先。<br>第 2 页自动填充。<br>22 页可逐页检查修订。</p></div></aside><main class="main">${[null, outlineScreen, outlineReviewScreen, planScreen, templateScreen, state.generationView === 'edit' ? editorScreen : generateScreen, reviewScreen][state.step]()}</main></div>
    ${modalScreen()}`;
  if (focus) { const target = document.querySelector(`[data-focus="${focus.key}"]`); if (target) { target.focus({ preventScroll: true }); if (typeof focus.start === 'number') try { target.setSelectionRange(focus.start, focus.end); } catch { /* Some controls have no selection. */ } } }
  if(screenChanged)window.scrollTo({top:0,left:0,behavior:'instant'});
  renderedScreen=screenKey;
}

function templateScreen() {
  return `<div class="page-head"><div><div class="eyebrow">04 / PPT TEMPLATE</div><h1>草稿准备好，再选择设计</h1><p>请上传 PPTX 模板，或复用已经提供并标注好的模板。</p></div>${button('选择 PPTX 文件', 'template-upload', 'secondary', 'upload', busy())}</div>
  <div class="template-grid"><section class="panel"><div class="panel-head"><h2>平安货运 · 原模板</h2><span>第 ${state.sourceSlide} / 22 页</span></div><div class="stage-body"><img class="slide-image" src="/source/slide-${state.sourceSlide}.png" alt="原模板第 ${state.sourceSlide} 页" /><div class="caption"><span>原模板预览</span><span class="tag ${state.sourceSlide === 2 ? '' : 'neutral'}">${state.sourceSlide === 2 ? '本次可编辑页' : '保持原内容'}</span></div></div><div class="thumb-strip" aria-label="模板页面">${Array.from({ length: 22 }, (_, i) => `<button class="thumb ${state.sourceSlide === i + 1 ? 'active' : ''}" data-action="source-page" data-page="${i + 1}"><img loading="lazy" src="/source/slide-${i + 1}.png" alt="第 ${i + 1} 页缩略图" /><span>${String(i + 1).padStart(2, '0')}${i === 1 ? ' · 可编辑' : ''}</span></button>`).join('')}</div></section>
  <aside class="panel details"><div class="eyebrow">TEMPLATE READY</div><h2>设计已经定好<br>内容仍然可编辑</h2><p>已读取这份模板，并标注第二页。</p><ul class="rule-list"><li>${icon('check')}<div>背景、母版与版式<small>固定属性保持原样</small></div></li><li>${icon('check')}<div>文字位置与字体格式<small>13 个文字位置可替换</small></div></li><li>${icon('check')}<div>图片框与原有裁切<small>5 张配图可更换</small></div><li>${icon('check')}<div>其他 21 页<small>本次保留原内容</small></div></li></ul><hr /><div class="file-row"><div class="file-icon">P</div><div class="file-info"><strong title="${escape(state.template.name)}">${escape(state.template.name)}</strong><span>22 页 · 16:9 · ${state.template.sizeMb} MiB</span></div></div><p class="note" style="margin-top:18px">最小 Demo 支持这份已标注模板。上传相同文件可核对；其他模板暂需先做位置标注。</p></aside></div>
  <div class="footer-actions"><p>${state.workflow.templateConfirmed ? '模板已确认，继续查看计划书的摘要分配。' : '计划书草稿已准备，请选择文件或使用已经提供的模板。'}</p>${button('使用模板，自动分配内容', 'choose-template', 'primary', 'arrow', busy() || !state.workflow.planReady)}</div>`;
}

function meter(fit) {
  return `<div class="meter ${fit && !fit.fits ? 'over' : ''}"><span style="width:${Math.min(100, fit?.percent || 0)}%"></span></div>`;
}
function editorScreen() {
  const item = section(), column = columns[state.selected], fit = capacity(item.id), asset = state.assets.get(item.assetKey);
  const bounds = [{ l: 1.87, t: 27.2, w: 17.38, h: 39 }, { l: 20.89, t: 30.38, w: 19.04, h: 31.47 }, { l: 40.61, t: 31.56, w: 18.62, h: 39 }, { l: 60.72, t: 31, w: 18.62, h: 35.24 }, { l: 81.24, t: 26.42, w: 16.9, h: 34.9 }][state.selected];
  const columnIssues = state.checks?.issues?.filter(x => x.field === item.id) || [];
  return `<div class="page-head"><div><div class="eyebrow">05 / PPT CONTENT</div><h1>把计划书摘要放进原模板</h1><p>当前适配第二页摘要；其余页在逐页审核中按原文字位置修订。</p></div><div class="actions">${button('调整五栏摘要', 'outline', 'secondary small', 'upload', busy())}${button(state.longBackup ? '恢复测试前文案' : '试一下超长文案', state.longBackup ? 'restore-long' : 'long-demo', 'text small', 'refresh', busy())}</div></div>
  <div class="project-strip panel"><div><span>项目题目</span><strong>${escape(state.draft.topic)}</strong><small>${escape(state.draft.projectName)} · ${escape(state.draft.subtitle)}</small></div>${button('编辑题目', 'project', 'secondary small', 'edit', busy())}</div>
  ${state.job && stale() ? `<div class="notice info" style="margin-bottom:17px">${icon('refresh')}内容已修改，上一版检查记录已失效。再次生成后再进行 WPS 复核。</div>` : ''}
  <div class="editor-grid"><section class="panel"><div class="column-tabs" role="tablist" aria-label="五栏内容">${columns.map((c, i) => `<button class="column-tab ${i === state.selected ? 'active' : ''} ${errors().some(x => x.field === c.id) ? 'has-error' : ''}" data-action="column" data-column="${i}" role="tab" aria-selected="${i === state.selected}">${escape(state.draft.sections[i].heading || c.name)}</button>`).join('')}</div><div class="column-editor"><div class="field"><label for="heading">栏标题<small>最多 ${column.headingLimit} 个汉字宽度</small></label><input id="heading" data-focus="heading" data-section="${item.id}" data-part="heading" value="${escape(item.heading)}" ${busy() ? 'disabled' : ''} /></div>
  ${item.paragraphs.map((p, i) => `<div class="field paragraph-field"><label for="paragraph-${i}">第 ${i + 1} 段<small>${i === 1 && column.id === '65' ? '原模板普通段落' : '沿用原段落格式'}</small></label><textarea id="paragraph-${i}" data-focus="paragraph-${i}" data-section="${item.id}" data-paragraph="${i}" data-part="text" ${busy() ? 'disabled' : ''}>${escape(p.text)}</textarea><div class="field-bottom"><span>内容完整优先，超限时精简措辞</span><span id="count-${i}">${[...p.text].length} 字</span></div>${column.emphasis[i] ? `<div class="emphasis-field"><span>黄色重点</span><input aria-label="第 ${i + 1} 段黄色重点" data-focus="emphasis-${i}" data-section="${item.id}" data-paragraph="${i}" data-part="emphasis" value="${escape(p.emphasis)}" placeholder="最多 ${column.emphasis[i]} 处，用顿号分隔" ${busy() ? 'disabled' : ''} /></div>` : ''}</div>`).join('')}
  <div class="capacity"><div class="capacity-row"><span>${state.checking ? '正在检查容量…' : fit?.fits ? '当前正文预计可容纳' : fit ? '正文预计超出原框' : '等待容量估算'}</span><strong style="color:${fit?.fits === false ? 'var(--red)' : 'var(--green)'}">${fit ? `${fit.percent}%` : '—'}</strong></div>${meter(fit)}<p class="note">按原字号、边距与行距估算占用比例，实际换行在 WPS 中确认。</p></div>${columnIssues.map(issue => `<div class="notice ${issue.type === 'error' ? 'error' : ''}" style="margin-top:12px">${icon('alert')}${escape(issue.message)}</div>`).join('')}</div></section>
  <section class="panel editor-reference"><div class="panel-head"><h2>文字位置参考</h2><span>第 2 页 · ${escape(item.heading)}</span></div><div class="stage-body"><div class="reference-frame"><img class="slide-image" src="/source/slide-2.png" alt="原模板文字位置参考" /><div class="slot-highlight" style="left:${bounds.l}%;top:${bounds.t}%;width:${bounds.w}%;height:${bounds.h}%"></div></div><div class="caption"><span>选中的原文字框已标出</span><span>生成后查看新内容</span></div></div><div class="asset-editor"><img src="${escape(asset?.url || '')}" alt="${escape(item.heading)}配图" /><div><strong>${escape(item.heading)} · 配图</strong><p>${asset?.builtin ? '内置 AI 概念示意图' : escape(asset?.name || '新配图')}</p><div class="actions" style="margin-top:10px">${button('更换 PNG', 'image-upload', 'secondary small', 'image', busy())}${!asset?.builtin ? button('恢复示例图', 'reset-image', 'text small', null, busy()) : ''}</div></div></div><div class="notice">${icon('eye')}图片会沿用原框和原裁切。生成后请检查主体是否完整；概念图不能替代实测材料。</div></section></div>
  <div class="footer-actions"><p>${state.checking ? '正在核对当前内容…' : errors().length ? `还有 ${errors().length} 个问题需要处理` : '五栏内容已准备，可以进入生成检查。'}</p>${button('检查与生成', 'generate-step', 'primary', 'arrow', busy() || state.checking)}</div>`;
}

function capacityList() {
  return `<div class="capacity-list">${columns.map(c => { const fit = capacity(c.id); return `<div class="capacity-item"><div><span>${escape(state.draft.sections.find(x => x.id === c.id).heading)}</span><small style="color:${fit?.fits === false ? 'var(--red)' : ''}">${fit ? `${fit.percent}% · ${fit.estimatedLines} 行` : '待检查'}</small></div>${meter(fit)}</div>`; }).join('')}<p class="note">比例表示原框高度的估算占用量，保留 10% 余量。超过预算会拦截生成。</p></div>`;
}
function generateScreen() {
  const job = state.job;
  const generating = job?.status === 'running' || state.starting;
  const issues = state.checks?.issues || [];
  return `<div class="page-head"><div><div class="eyebrow">05 / VALIDATE & GENERATE</div><h1>先检查，再生成副本</h1><p>计划书摘要、逐页修订与确认的政策引用一起写入新副本。</p></div>${button('返回内容分配', 'editor', 'secondary small', 'back', generating)}</div>
  <div class="audit-summary"><div class="summary-cell"><strong>13</strong><span>可替换文字位置</span><small>字号、位置和样式保留</small></div><div class="summary-cell"><strong>5</strong><span>可更换配图</span><small>原图框与裁切保留</small></div><div class="summary-cell"><strong style="color:${errors().length ? 'var(--red)' : 'var(--green)'}">${state.checking ? '…' : errors().length}</strong><span>当前阻塞问题</span><small>${errors().length ? '处理后才能生成' : '当前文案通过预检查'}</small></div></div>
  <div class="generate-grid"><section class="panel"><div class="panel-head"><h2>当前内容检查</h2><span>${state.checking ? '检查中' : '针对当前文案'}</span></div>${issues.length ? `<div class="issue-list">${issues.map((issue, i) => `<div class="issue ${issue.type}"><p>${escape(issue.message)}</p>${button('去修改', 'fix-issue', 'text small', 'edit', generating, `data-issue="${i}"`)}</div>`).join('')}</div>` : `<div class="empty-state">${icon('check')}<div><strong>可以生成这一页</strong><p>容量估算与文字格式匹配检查已通过。</p></div></div>`}<div class="checks-list"><div class="check-row"><span class="symbol">${icon('lock')}</span><div><strong>固定模板保护</strong><p>生成时逐项核对位置、字体、图片框、母版和其他 21 页。</p></div></div><div class="check-row warning"><span class="symbol">${icon('alert')}</span><div><strong>已有模板问题单独保留</strong><p>字体缺失及全套图表校验限制会显示在结果中，不计为“全部通过”。</p></div></div></div></section>
  <aside><section class="panel"><div class="panel-head"><h2>五栏容量</h2><span>估算值</span></div>${capacityList()}</section><section class="panel" style="margin-top:20px">${generating ? `<div class="progress-block"><div class="progress-title"><strong><span class="busy-indicator"></span> 正在生成</strong><span>${job?.progress || 3}%</span></div><div class="progress-meter"><span style="width:${job?.progress || 3}%"></span></div><p class="note" style="margin-top:13px">${escape(job?.message || '正在建立任务')}</p><ul class="stage-list">${(job?.stages || []).slice(-5).map(s => `<li>${icon('check')}${escape(s.text)}</li>`).join('')}</ul></div>` : `<div class="generation-start"><h3>${job?.status === 'failed' ? '这一轮生成未完成' : '生成可编辑的 PPT'}</h3><p>${job?.status === 'failed' ? '已保留当前文案，请查看原因后重试。' : '约需几十秒。保留完整 22 页，仅替换第二页文字和图片。'}</p>${job?.status === 'failed' ? `<details class="note"><summary>查看错误原因</summary><pre style="white-space:pre-wrap;font-size:10px">${escape(job.error)}</pre></details>` : ''}${button('生成 PPT 副本', 'generate', 'primary', 'arrow', state.checking || !state.checks?.canGenerate || busy())}${job?.status === 'ready' && !stale() ? button('查看生成结果', 'review-step', 'text', 'eye') : ''}</div>`}</section></aside></div>`;
}

function reviewScreen() {
  const job = state.job, result = job?.result;
  if (!result || job.status !== 'ready') return `<div class="page-head"><div><div class="eyebrow">04 / Review & Export</div><h1>查看结果，核对实际显示</h1><p>生成后会在这里展示新页面和检查记录。</p></div></div><section class="panel"><div class="generation-start"><h3>${job?.status === 'running' ? 'PPT 正在生成' : '还没有生成结果'}</h3><p>先在第三步检查当前内容，再生成一个新的 PPT 副本。</p>${button('前往检查与生成', 'generate-step', 'primary', 'arrow')}</div></section>`;
  const isStale = stale(), allConfirmed = job.review?.confirmed && !isStale;
  const confirmed=job.pageReviews?.[state.reviewPage]?.confirmed && !isStale;
  const checks = result.checks.map(item => item.title === 'WPS 视觉复核' && allConfirmed ? { ...item, status: 'pass', detail: '当前版本所有修改页均记录了人工复核确认。' } : item);
  const page = state.reviewPage;
  const previewUrl = `${result.previewUrl}?page=${page}`;
  return `<div class="page-head"><div><div class="eyebrow">06 / PAGE REVIEW</div><h1>${isStale ? '修订已保存，请更新 PPT' : `逐页核对 · 第 ${page} 页`}</h1><p>${escape(state.draft.projectName)} · ${result.editedSlides.includes(page) ? '本页包含已写入的修改或引用备注' : '本页仍是原模板内容，需确认是否适用于新项目'}。</p></div><div class="actions">${button('修改本页文字', 'page-edit', 'secondary', 'edit', busy())}${button(isStale ? '更新 PPT' : '内容分配', isStale ? 'generate-step' : 'editor', 'secondary', isStale ? 'refresh' : 'back')}</div></div>
  <div class="page-review-strip panel" aria-label="逐页选择">${Array.from({length:22},(_,i)=>`<button data-action="review-page" data-page="${i+1}" class="page-chip ${page===i+1?'active':''}"><strong>${String(i+1).padStart(2,'0')}</strong><span>${i===1?'自动摘要':Object.keys(state.workflow.pageEdits[i+1]||{}).length?'有文字修订':state.workflow.policySelections[i+1]?.length?'有引用':'待核对'}</span></button>`).join('')}</div>
  ${isStale ? `<div class="notice error" style="margin-bottom:20px">${icon('alert')}这里是上一版结果。当前文案尚未重新生成，旧的复核确认不能用于当前内容。</div>` : ''}
  <div class="review-grid"><section class="panel"><div class="panel-head"><h2>${isStale ? `上一版 · 第 ${page} 页` : `第 ${page} 页`}</h2><div class="view-tabs" aria-label="预览方式">${[['generated', '当前页'], ['source', '原模板'], ['compare', '对照']].map(([key, text]) => `<button data-action="view" data-view="${key}" class="${state.view === key ? 'active' : ''}">${text}</button>`).join('')}</div></div><div class="stage-body">${state.view === 'compare' ? `<div class="compare-grid"><div><small>原模板</small><img class="slide-image" src="/source/slide-${page}.png" alt="原模板第${page}页" /></div><div><small>当前页</small><img class="slide-image" src="${previewUrl}" alt="当前版本第${page}页" /></div></div>` : `<img class="slide-image" src="${state.view === 'source' ? `/source/slide-${page}.png` : previewUrl}" alt="第${page}页实际预览" />`}<div class="caption"><span>网页预览与 WPS 可能有差异</span><span class="tag success">固定模板保护通过</span></div></div><div class="checks-list">${checks.map(item => `<div class="check-row ${item.status}"><span class="symbol">${icon(item.status === 'pass' ? 'check' : item.status === 'warning' ? 'alert' : 'eye')}</span><div><strong>${item.title}</strong><p>${escape(item.detail)}</p></div></div>`).join('')}</div></section>
  <aside><section class="panel policy-card">${policyScreen()}</section><section class="panel review-card" style="margin-top:18px"><div class="eyebrow">VISUAL REVIEW</div><h2>在 WPS 中做最后核对</h2><p>检查当前版本实际修改过的页面，原模板内容仍需主题核对。</p>${button('在 WPS 中打开', 'open-wps', 'secondary', 'eye', isStale)}<p class="note">打开请求成功不代表复核完成。当前查看页 ${page}。</p><hr /><label class="review-checkbox"><input id="review-confirm" type="checkbox" ${confirmed ? 'checked' : ''} ${isStale ? 'disabled' : ''} />我已在 WPS 核对当前版本第 ${page} 页。</label><textarea id="review-note" data-focus="review-note" placeholder="本页复核备注（可选）" ${isStale ? 'disabled' : ''}>${escape(state.reviewNote)}</textarea>${button('保存本页复核备注', 'save-review', 'text small', null, isStale)}<p class="note">修改页复核进度：${result.editedSlides.filter(n=>job.pageReviews?.[n]?.confirmed).length} / ${result.editedSlides.length}</p><div class="download-row">${button('下载已复核 PPT', 'download-reviewed', 'primary', 'download', !allConfirmed)}${button('先下载候选 PPT', 'download-candidate', 'secondary', 'download', isStale)}${button('下载检查记录', 'download-report', 'text small', 'download', isStale)}</div><p class="job-badge">${escape(result.filename)}</p><hr /><p class="note">未修改页保留原材料；表格、图表和原图片需要另行适配。政策引用写入本页备注，不自动变成项目成果。</p></section></aside></div>`;
}

function modalScreen() {
  if (!state.modal) return '';
  if (state.modal.type === 'page') {
    const page = state.pages.find(p=>p.number===state.modal.page);
    return `<div class="modal-backdrop"><section class="modal" role="dialog" aria-modal="true" aria-labelledby="modal-title"><div class="modal-head"><h2 id="modal-title">修订第 ${page.number} 页文字</h2><button data-action="close-page" aria-label="关闭">${icon('close')}</button></div><div class="modal-body"><p class="note">保持原文字位置和段数。预算以原文长度为参考，最终需在 WPS 核对；图表、表格与原图片不会自动改写。</p>${page.slots.map(slot=>`<div class="field page-slot"><label>${escape(slot.name)}</label>${slot.paragraphs.map((p,i)=>`<textarea aria-label="位置 ${slot.id} 第 ${i+1} 段" data-page-slot="${slot.id}" data-paragraph="${i}" data-focus="page-${slot.id}-${i}">${escape(state.modal.values[slot.id][i])}</textarea><p class="note">第 ${i+1} 段 · 原文预算约 ${Math.floor(slot.budgets[i])} 个汉字宽度</p>`).join('')}</div>`).join('')}<div class="field page-slot"><label for="page-note">本页修订说明（写入备注）</label><textarea id="page-note" data-focus="page-note">${escape(state.modal.note)}</textarea></div>${state.modalError?`<p class="modal-error">${escape(state.modalError)}</p>`:''}</div><div class="modal-foot"><span class="note">保存后需更新 PPT，旧复核会失效</span>${button('保存本页修订','save-page','primary','check')}</div></section></div>`;
  }
  const isOutline = state.modal.type === 'outline';
  return `<div class="modal-backdrop"><section class="modal" role="dialog" aria-modal="true" aria-labelledby="modal-title"><div class="modal-head"><h2 id="modal-title">${isOutline ? '把大纲整理到五栏' : '项目题目与页面标题'}</h2><button data-action="close-modal" aria-label="关闭">${icon('close')}</button></div><div class="modal-body">${isOutline ? `<p class="note">按顺序填写五个 ## 标题，每段独占一行。对应段数为 2、2、3、2、2；不自动删掉内容或补造信息。</p><textarea id="outline-input" data-focus="outline" aria-label="五栏大纲">${escape(state.modal.value)}</textarea>` : `<div class="form-head" style="padding:0"><div class="field wide"><label>完整项目题目</label><input data-global="topic" data-focus="topic" value="${escape(state.draft.topic)}" /></div><div class="field"><label>项目简称<small>4 个汉字宽度</small></label><input data-global="projectName" data-focus="projectName" value="${escape(state.draft.projectName)}" /></div><div class="field"><label>页面标题<small>4 个汉字宽度</small></label><input data-global="pageTitle" data-focus="pageTitle" value="${escape(state.draft.pageTitle)}" /></div><div class="field wide"><label>副标题<small>10 个汉字宽度</small></label><input data-global="subtitle" data-focus="subtitle" value="${escape(state.draft.subtitle)}" /></div></div><p class="note" style="margin-top:18px">标题预算沿用原位置。改题目后，请同时编辑正文和配图，让内容与新题目一致。</p>`}${state.modalError ? `<p class="modal-error">${escape(state.modalError)}</p>` : ''}</div><div class="modal-foot"><span class="note">文字保持可编辑，固定版式保留</span>${button(isOutline ? '整理到原有位置' : '完成', isOutline ? 'apply-outline' : 'close-modal', 'primary', 'check')}</div></section></div>`;
}

function applyWorkflow(workflow, resetDraft=false) {
  state.workflow=workflow;
  if (resetDraft && workflow.lastDraft) state.draft=structuredClone(workflow.lastDraft);
  if (!state.planDirty) state.planText=workflow.plan?.markdown || '';
}
function outlineScreen() {
  const p=state.workflow.project;
  return `<div class="page-head"><div><div class="eyebrow">01 / OUTLINE</div><h1>先把项目讲清楚</h1><p>先确定题目、事实和章节，再选择 PPT 的设计。</p></div>${button('保存大纲，人工审核','save-outline','primary','arrow',busy())}</div><section class="panel outline-project"><div class="panel-head"><h2>项目设定</h2><span>政策范围：全国官方政策</span></div><div class="workflow-form"><div class="field wide"><label for="wf-topic">项目题目</label><input id="wf-topic" data-wf="topic" data-focus="wf-topic" value="${escape(p.topic)}" /></div><div class="field"><label for="wf-name">项目简称<small>1–4 个字</small></label><input id="wf-name" data-wf="name" data-focus="wf-name" value="${escape(p.name)}" /></div><div class="field"><label for="wf-audience">服务对象</label><input id="wf-audience" data-wf="audience" data-focus="wf-audience" value="${escape(p.audience)}" /></div>${[['problem','要解决的问题'],['solution','解决方案'],['facts','已有事实、阶段与资料']].map(([key,label])=>`<div class="field wide"><label for="wf-${key}">${label}</label><textarea id="wf-${key}" data-wf="${key}" data-focus="wf-${key}">${escape(p[key])}</textarea></div>`).join('')}</div></section><section class="panel" style="margin-top:20px"><div class="panel-head"><h2>计划书章节</h2>${button('增加章节','add-chapter','text small',null)}</div><div class="chapter-list">${p.chapters.map((c,i)=>`<article class="chapter"><span class="chapter-number">${String(i+1).padStart(2,'0')}</span><div class="field"><label>章节标题</label><input aria-label="第 ${i+1} 章标题" data-chapter="${i}" data-part="title" data-focus="chapter-title-${i}" value="${escape(c.title)}" /><label style="margin-top:12px">本章要点</label><textarea aria-label="第 ${i+1} 章要点" data-chapter="${i}" data-part="points" data-focus="chapter-points-${i}">${escape(c.points)}</textarea></div>${button('移除','remove-chapter','text small',null,p.chapters.length<=3,`data-chapter-index="${i}"`)}</article>`).join('')}</div></section><div class="footer-actions"><p>修改大纲后，旧审核和草稿会失效。</p>${button('保存大纲，人工审核','save-outline','primary','arrow')}</div>`;
}
function outlineReviewScreen() {
  const w=state.workflow,p=w.project;
  return `<div class="page-head"><div><div class="eyebrow">02 / HUMAN REVIEW</div><h1>确认大纲，再开始写草稿</h1><p>人工核对需要表达的内容、已有事实和待补资料。</p></div>${button('返回改大纲','outline-step','secondary','back')}</div><div class="workflow-two-col"><section class="panel"><div class="plan-reading"><h2>${escape(p.topic)}</h2><p><strong>服务对象：</strong>${escape(p.audience)}</p><p><strong>问题：</strong>${escape(p.problem)}</p><p><strong>方案：</strong>${escape(p.solution)}</p><p><strong>事实与阶段：</strong>${escape(p.facts||'待补充')}</p>${p.chapters.map((c,i)=>`<h3>${i+1}. ${escape(c.title)}</h3><p>${escape(c.points).replaceAll('\n','<br>')}</p>`).join('')}</div></section><aside class="panel review-card"><div class="eyebrow">REVIEW GATE</div><h2>大纲审核</h2><p>没有资料支持的测试、收入、合作与资质不能直接作为成果。</p>${w.outlineIssues.map(issue=>`<div class="notice error" style="margin-bottom:12px">${escape(issue)}</div>`).join('')}<label class="review-checkbox"><input id="outline-confirm" type="checkbox" ${state.outlineConfirmed?'checked':''} />我已核对题目、章节安排和事实边界，确认当前大纲可用于生成草稿。</label><textarea id="outline-note" data-focus="outline-note" placeholder="审核意见（可选）">${escape(state.outlineNote)}</textarea>${button('确认大纲并生成草稿','approve-outline','primary','check',!state.outlineConfirmed || w.outlineIssues.length>0)}${w.approved?'<p class="note">当前版本已记录大纲确认，修改大纲后需重新审核。</p>':''}</aside></div>`;
}
function markdownPreview(text) {
  return text.split('\n').map(line=>line.startsWith('### ')?`<h3>${escape(line.slice(4))}</h3>`:line.startsWith('## ')?`<h2>${escape(line.slice(3))}</h2>`:line.startsWith('# ')?`<h1>${escape(line.slice(2))}</h1>`:line.startsWith('- ')?`<p class="plan-bullet">${escape(line)}</p>`:line.trim()?`<p>${escape(line)}</p>`:'').join('');
}
function planScreen() {
  const w=state.workflow;
  return `<div class="page-head"><div><div class="eyebrow">03 / PLAN DRAFT</div><h1>把大纲整理成计划书草稿</h1><p>草稿可编辑，缺少的数据和材料保留为待补充项。</p></div><div class="actions">${button('生成计划书草稿','make-plan','secondary','refresh',!w.approved)}${w.plan?'<a class="btn secondary" href="/api/workflow/plan" download>下载草稿</a>':''}</div></div>${!w.approved?`<div class="notice error">请先人工审核大纲。${button('前往人工审核','review-outline-step','text small')}</div>`:w.plan?`<div class="plan-draft-grid"><section class="panel"><div class="panel-head"><h2>草稿正文</h2><span>版本 ${w.plan.revision}${state.planDirty?' · 有未保存修改':''}</span></div><div class="plan-edit"><textarea id="plan-text" data-focus="plan-text" aria-label="计划书草稿正文">${escape(state.planText)}</textarea></div></section><section class="panel"><div class="panel-head"><h2>阅读预览</h2><span>待补充项目需核对</span></div><article class="plan-reading">${markdownPreview(state.planText)}</article></section></div><div class="footer-actions"><p>依据审核大纲整理，尚未接入任意题目的大模型扩写。</p><div class="actions">${button('保存草稿','save-plan','secondary','check',!state.planDirty)}${button('草稿就绪，上传模板','template-step','primary','arrow',state.planDirty)}</div></div>`:`<section class="panel"><div class="generation-start"><h3>大纲已确认，可以生成草稿</h3><p>依据项目设定与已审核章节整理内容，不补造实测、金额和资质。</p>${button('生成计划书草稿','make-plan','primary','arrow')}</div></section>`}`;
}
function policyScreen() {
  const result=state.policyResults;
  const adopted=state.workflow.policySelections[state.reviewPage]||[];
  return `<div class="panel-head"><h2>本页政策材料</h2><span>全国官方</span></div><div class="policy-controls"><label for="policy-query">检索关键词</label><input id="policy-query" data-focus="policy-query" value="${escape(state.policyQuery)}" /><div class="actions">${button(state.policyLoading?'查询中…':'自动查询政策','policy-search','primary small','refresh',state.policyLoading)}</div><p class="note">检索只发送关键词，不发送完整计划书。切换页面后按本页内容匹配资料。</p></div>${state.policyLoading?'<div class="policy-status"><span class="busy-indicator"></span> 正在查询官方资料…</div>':result?`<div class="policy-status">${escape(result.onlineMessage)}</div><div class="policy-results">${result.results.map(p=>`<article class="policy-result"><a href="${escape(p.url)}" target="_blank" rel="noopener noreferrer">${escape(p.title)}</a><p class="policy-meta">${escape(p.issuer)} · 网页发布 ${escape(p.published||'待核对')}${p.issued?' · 发文 '+p.issued:''}</p><span class="tag ${p.historical?'danger':'neutral'}">${p.historical?'历史背景，规划期已结束':p.sourceKind==='live-search'?'联网找到，待核对适用性':'已核验官方原文'}</span><p>${escape(p.summary)}</p><p class="policy-condition">${escape(p.use)}</p>${button(adopted.includes(p.id)?'移除本页引用':'引用到本页备注','adopt-policy','secondary small','check',busy(),`data-policy-id="${p.id}" data-remove="${adopted.includes(p.id)}"`)}</article>`).join('')}</div><div class="policy-controls"><a class="btn text small" href="${escape(result.searchUrl)}" target="_blank" rel="noopener noreferrer">在政府网站继续检索</a><p class="note">引用后需要更新 PPT。来源和适用条件将写入本页备注。</p></div>`:'<p class="note" style="padding:20px">可以按本页主题自动查找政策材料，先核对原文，再选择引用。</p>'}`;
}
async function queryPolicy() {
  const sequence=++state.policySequence;
  state.policyLoading=true; render();
  try {
    const result=await post('/api/policies/search',{page:state.reviewPage,query:state.policyQuery});
    if(sequence!==state.policySequence)return;
    state.policyResults=result;state.policyLoading=false;render();
  } catch(error){if(sequence===state.policySequence){state.policyLoading=false;toast(error.message);render();}}
}

function markDirty() { checkSequence++; state.dirty = true; state.checking = true; clearTimeout(checkTimer); checkTimer = setTimeout(runCheck, 380); }
async function runCheck() {
  const sequence = ++checkSequence;
  state.checking = true;
  try {
    const checked = await post('/api/check', state.draft);
    if (sequence !== checkSequence) return;
    state.checks = checked; state.checking = false;
    if (state.job && checked.fingerprint === state.job.fingerprint) state.dirty = false;
    render();
  } catch (error) { if (sequence === checkSequence) { state.checking = false; toast(error.message); render(); } }
}

async function pollJob() {
  clearTimeout(pollTimer);
  if (!state.job || state.job.status !== 'running') return;
  try {
    state.job = await api(`/api/jobs/${state.job.id}`);
    if (state.job.status === 'ready') { state.step = 6; state.view = 'generated'; state.reviewNote = state.job.review?.note || ''; toast('PPT 已生成，可以逐页检查与查询政策。'); queryPolicy(); }
    if (state.job.status === 'failed') toast('生成未完成，已保留文案和错误原因。');
    render();
  } catch (error) { toast(error.message); }
  if (state.job?.status === 'running') pollTimer = setTimeout(pollJob, 1500);
}

app.addEventListener('input', event => {
  const target = event.target;
  if(target.dataset.wf){state.workflow.project[target.dataset.wf]=target.value;state.outlineDirty=true;state.outlineConfirmed=false;state.workflow.approved=false;return;}
  if(target.dataset.chapter!==undefined){state.workflow.project.chapters[+target.dataset.chapter][target.dataset.part]=target.value;state.outlineDirty=true;state.outlineConfirmed=false;state.workflow.approved=false;return;}
  if(target.id==='outline-note'){state.outlineNote=target.value;return;}
  if(target.id==='plan-text'){state.planText=target.value;state.planDirty=true;return;}
  if(target.id==='policy-query'){state.policyQuery=target.value;return;}
  if(target.dataset.pageSlot){state.modal.values[target.dataset.pageSlot][+target.dataset.paragraph]=target.value;return;}
  if(target.id==='page-note'){state.modal.note=target.value;return;}
  if (target.id === 'outline-input') { state.modal.value = target.value; return; }
  if (target.id === 'review-note') { state.reviewNote = target.value; return; }
  if (target.dataset.global) state.draft[target.dataset.global] = target.value;
  else if (target.dataset.section) {
    const item = state.draft.sections.find(x => x.id === target.dataset.section);
    if (target.dataset.paragraph !== undefined) {
      item.paragraphs[+target.dataset.paragraph][target.dataset.part] = target.value;
      if (target.dataset.part === 'text') { const count = document.querySelector(`#count-${target.dataset.paragraph}`); if (count) count.textContent = [...target.value].length + ' 字'; }
    } else item[target.dataset.part] = target.value;
  } else return;
  markDirty();
});

app.addEventListener('change', async event => {
  if(event.target.id==='outline-confirm'){state.outlineConfirmed=event.target.checked;render();return;}
  if (event.target.id !== 'review-confirm') return;
  if (stale()) { event.target.checked = false; toast('内容已变更，请重新生成后再复核。'); return; }
  try { state.job = await post(`/api/jobs/${state.job.id}/review`, { page:state.reviewPage,confirmed: event.target.checked, note: state.reviewNote }); render(); toast(event.target.checked ? '已记录本页复核确认。' : '已取消本页复核确认。'); } catch (error) { toast(error.message); render(); }
});

app.addEventListener('click', async event => {
  const node = event.target.closest('[data-action]'); if (!node || node.disabled) return;
  const action = node.dataset.action;
  try {
    if (action === 'step') {
      const step=+node.dataset.step;
      if(state.outlineDirty && step!==1)throw new Error('请先保存修改后的大纲，再重新审核。');
      if(step>=3 && !state.workflow.approved)throw new Error('请先人工审核大纲。');
      if(step>=4 && (!state.workflow.planReady || state.planDirty))throw new Error('请先生成并保存计划书草稿。');
      if(step>=5 && !state.workflow.templateConfirmed)throw new Error('请先选择或上传 PPT 模板。');
      state.step=step;render();if(step===6 && state.job?.status==='ready')queryPolicy();
    }
    if (action === 'editor') { state.step = 5; state.generationView='edit'; render(); }
    if (action === 'generate-step') { await runCheck(); state.step = 5; state.generationView='checks'; render(); }
    if (action === 'review-step') { state.step = 6; render(); queryPolicy(); }
    if(action==='save-outline'){const w=await post('/api/workflow',{project:state.workflow.project});state.outlineDirty=false;state.planDirty=false;applyWorkflow(w);state.step=2;render();}
    if(action==='outline-step'){state.step=1;render();}
    if(action==='review-outline-step'){state.step=2;render();}
    if(action==='add-chapter'){state.workflow.project.chapters.push({id:String(state.workflow.project.chapters.length+1),title:'新章节',points:''});state.outlineDirty=true;state.outlineConfirmed=false;render();}
    if(action==='remove-chapter'){state.workflow.project.chapters.splice(+node.dataset.chapterIndex,1);state.outlineDirty=true;state.outlineConfirmed=false;render();}
    if(action==='approve-outline'){applyWorkflow(await post('/api/workflow/approve',{confirmed:state.outlineConfirmed,note:state.outlineNote}));applyWorkflow(await post('/api/workflow/plan',{generate:true}),true);state.planDirty=false;state.step=3;state.dirty=true;render();await runCheck();}
    if(action==='make-plan'){state.planDirty=false;applyWorkflow(await post('/api/workflow/plan',{generate:true}),true);state.step=3;state.dirty=true;render();await runCheck();}
    if(action==='save-plan'){applyWorkflow(await post('/api/workflow/plan',{markdown:state.planText}),true);state.planDirty=false;state.dirty=true;render();await runCheck();}
    if(action==='template-step'){if(state.planDirty)throw new Error('请先保存计划书草稿。');state.step=4;render();}
    if(action==='choose-template'){applyWorkflow(await post('/api/template/choose',{}),true);state.step=5;state.generationView='edit';state.dirty=true;render();await runCheck();}
    if(action==='review-page'){state.reviewPage=+node.dataset.page;state.policyResults=null;state.reviewNote=state.job.pageReviews?.[state.reviewPage]?.note||'';state.view='generated';render();queryPolicy();}
    if(action==='policy-search')await queryPolicy();
    if(action==='adopt-policy'){applyWorkflow(await post('/api/policies/adopt',{page:state.reviewPage,id:node.dataset.policyId,remove:node.dataset.remove==='true'}));state.dirty=true;await runCheck();toast(node.dataset.remove==='true'?'已移除本页引用，请更新 PPT。':'政策已加入本页引用备注，请更新 PPT。');}
    if(action==='page-edit'){
      if(state.reviewPage===2){state.step=5;state.generationView='edit';render();}
      else{const page=state.pages.find(p=>p.number===state.reviewPage);state.modal={type:'page',page:page.number,values:Object.fromEntries(page.slots.map(s=>[s.id,structuredClone(state.workflow.pageEdits[page.number]?.[s.id]||s.paragraphs)])),note:state.workflow.pageReviews[page.number]?.note||''};state.modalError='';render();}
    }
    if(action==='close-page'){state.modal=null;state.modalError='';render();}
    if(action==='save-page'){
      try{applyWorkflow(await post('/api/page-edit',{page:state.modal.page,texts:Object.entries(state.modal.values).map(([id,paragraphs])=>({id,paragraphs})),note:state.modal.note}));state.modal=null;state.dirty=true;render();await runCheck();toast('本页修订已保存，更新 PPT 后核对效果。');}
      catch(error){state.modalError=error.message;render();}
    }
    if (action === 'source-page') { state.sourceSlide = +node.dataset.page; render(); }
    if (action === 'column') { state.selected = +node.dataset.column; render(); }
    if (action === 'view') { state.view = node.dataset.view; render(); }
    if (action === 'project') { state.modal = { type: 'project' }; state.modalError = ''; render(); }
    if (action === 'outline') { state.modal = { type: 'outline', value: draftToOutline(state.draft) }; state.modalError = ''; render(); }
    if (action === 'close-modal') { state.modal = null; state.modalError = ''; render(); await runCheck(); }
    if (action === 'apply-outline') {
      try { state.draft = applyOutline(state.draft, state.modal.value); state.modal = null; state.modalError = ''; markDirty(); render(); await runCheck(); toast('五栏大纲已放入对应文字位置。'); }
      catch (error) { state.modalError = error.message; render(); }
    }
    if (action === 'long-demo') { const p = state.draft.sections[0].paragraphs[0]; state.longBackup = structuredClone(p); p.text = '需要保留完整的巡检场景、风险说明、技术步骤和验证条件。'.repeat(100); p.emphasis = ''; state.selected = 0; markDirty(); await runCheck(); toast('已放入超长测试文字，生成会被拦截。'); }
    if (action === 'restore-long') { state.draft.sections[0].paragraphs[0] = state.longBackup; state.longBackup = null; markDirty(); await runCheck(); toast('已恢复测试前文案，其他修改保留。'); }
    if (action === 'template-upload') document.querySelector('#template-file').click();
    if (action === 'image-upload') document.querySelector('#picture-file').click();
    if (action === 'reset-image') { section().assetKey = columns[state.selected].asset; markDirty(); await runCheck(); }
    if (action === 'fix-issue') {
      const issue = state.checks.issues[+node.dataset.issue]; const index = columns.findIndex(c => c.id === issue.field);
      state.step = 5; state.generationView='edit';
      if (index >= 0) state.selected = index; else state.modal = { type: 'project' };
      render(); document.querySelector(index >= 0 ? `#paragraph-${issue.paragraph || 0}` : `[data-global="${issue.field}"]`)?.focus();
    }
    if (action === 'generate') {
      state.starting = true; state.step = 5; state.generationView='checks'; render();
      try { state.job = await post('/api/jobs', normalizeDraft(state.draft)); state.dirty = false; state.starting = false; render(); pollTimer = setTimeout(pollJob, 800); }
      catch (error) { state.starting = false; if (error.data?.job) { state.job = error.data.job; pollTimer = setTimeout(pollJob, 800); } throw error; }
    }
    if (action === 'open-wps') { const result = await post(`/api/jobs/${state.job.id}/wps`, {}); toast(result.message); }
    if (action === 'save-review') { state.job = await post(`/api/jobs/${state.job.id}/review`, { page:state.reviewPage,confirmed: !!state.job.pageReviews?.[state.reviewPage]?.confirmed, note: state.reviewNote }); toast('本页复核备注已保存。'); render(); }
    if (['download-reviewed', 'download-candidate', 'download-report'].includes(action)) {
      if (stale()) throw new Error('当前内容已变化，请重新生成。');
      if (action === 'download-reviewed' && !state.job.review?.confirmed) throw new Error('请先完成 WPS 视觉复核。');
      const a = document.createElement('a'); a.href = action === 'download-report' ? state.job.result.reportUrl : state.job.result.downloadUrl+(action==='download-reviewed'?'?reviewed=1':''); a.download = ''; document.body.append(a); a.click(); a.remove();
    }
  } catch (error) { toast(error.message); render(); }
});

document.querySelector('#template-file').addEventListener('change', async event => {
  const file = event.target.files[0]; if (!file) return;
  state.uploading = true; render(); toast('正在核对模板文件…');
  try { const result = await api('/api/template', { method: 'POST', headers: { 'Content-Type': 'application/octet-stream' }, body: file }); state.sourceSlide = 2; if(result.workflow)applyWorkflow(result.workflow,true);toast(result.message); }
  catch (error) { toast(error.message); }
  finally { state.uploading = false; event.target.value = ''; render(); }
});
document.querySelector('#picture-file').addEventListener('change', async event => {
  const file = event.target.files[0]; if (!file) return;
  const id = section().id; state.uploading = true; render();
  try { const image = await api('/api/images', { method: 'POST', headers: { 'Content-Type': 'application/octet-stream', 'X-File-Name': encodeURIComponent(file.name) }, body: file }); state.assets.set(image.key, image); state.draft.sections.find(s => s.id === id).assetKey = image.key; markDirty(); toast('新配图已选好，生成后检查原裁切效果。'); }
  catch (error) { toast(error.message); }
  finally { state.uploading = false; event.target.value = ''; render(); await runCheck(); }
});

try {
  const bootstrap = await api('/api/bootstrap'); state.draft = bootstrap.draft; state.template = bootstrap.template;state.workflow=bootstrap.workflow;state.pages=bootstrap.pages;state.planText=state.workflow.plan?.markdown||'';
  state.assets = new Map(bootstrap.assets.map(a => [a.key, { ...a, builtin: !a.key.includes('-') || columns.some(c => c.asset === a.key) }]));
  state.job = bootstrap.latestJob; render(); await runCheck();
  if (state.job?.status === 'running') pollTimer = setTimeout(pollJob, 800);
} catch (error) { app.innerHTML = `<div class="loading">无法读取工作台：${escape(error.message)}</div>`; }
