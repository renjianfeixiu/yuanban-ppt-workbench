import { createHash } from 'node:crypto';

export const verifiedPolicies = [
  { id: 'smart-supply-2025', title: '加快数智供应链发展专项行动计划', issuer: '商务部等8部门', published: '2025-05-26', issued: '2025-03-24', url: 'https://app.www.gov.cn/govdata/gov/202505/26/529839/article.html', keywords: ['仓储', '供应链', '物流', '机器视觉', '智能传感', '智能', '方案'], summary: '提出推进数智供应链建设，并涉及机器视觉、智能传感和智能仓储等技术应用。', use: '可作为拟探索智能仓储场景的政策背景；不能证明项目已经落地或获得资金。', historical: false },
  { id: 'sme-digital-2025', title: '中小企业数字化赋能专项行动方案（2025—2027年）', issuer: '工业和信息化部、财政部等4部门', published: '2024-12-14', issued: '2024-12-12', url: 'https://app.www.gov.cn/govdata/gov/202412/14/522696/article.html', keywords: ['中小', '数字化', '转型', '人工智能', '企业', '商业', '运行'], summary: '提出体系化推进中小企业数字化转型与人工智能应用。', use: '需核对企业类型、试点地区和具体实施条件，不能写成自动享有补贴。', historical: false },
  { id: 'logistics-cost-2024', title: '有效降低全社会物流成本行动方案', issuer: '中共中央办公厅、国务院办公厅', published: '2024-11-28', issued: null, url: 'https://www.mee.gov.cn/zcwj/zyygwj/202411/t20241128_1097450.shtml', keywords: ['物流', '仓储', '成本', '管理', '运行', '需求'], summary: '提出从全局和战略高度推动物流成本下降，涉及运输、仓储和管理成本。', use: '用于背景与需求说明；项目降本幅度仍需实测或预算证据。', historical: false },
  { id: 'ai-plus-2025', title: '国务院关于深入实施“人工智能+”行动的意见', issuer: '国务院', published: '2025-08-27', issued: null, url: 'https://www.cac.gov.cn/2025-08/27/c_1758018277755538.htm', keywords: ['人工智能', '智能', '技术', '创新', '解决方案', '风险'], summary: '提出推动人工智能与各行业各领域融合，并强调安全可控。', use: '可解释方案探索方向；模型效果、实施资质及数据条件需要独立核对。', historical: false },
  { id: 'logistics-data-2025', title: '关于开展物流数据开放互联试点工作的通知', issuer: '国家发展改革委等部门', published: '2025-02-15', issued: null, url: 'https://app.www.gov.cn/govdata/gov/202502/15/524612/article.html', keywords: ['物流', '数据', '互联', '记录', '平台', '风险'], summary: '围绕物流数据开放互联开展试点，探索数据交换和应用机制。', use: '试点适用条件应单独核对，不能据此宣称已入选试点。', historical: false },
  { id: 'commerce-logistics-history', title: '商贸物流高质量发展专项行动计划（2021—2025年）', issuer: '商务部等9部门', published: '2021-08-09', issued: '2021-08-06', url: 'https://ltfzs.mofcom.gov.cn/smwltxjs/zcfb/art/2021/art_e0c80e8e4b1944bab00384b73971dc1f.html', keywords: ['物流', '仓储', '商贸', '历史'], summary: '历史规划材料，可用于说明相关发展背景。', use: '规划期已结束，只能标为历史背景；引用当前支持措施需找现行文件。', historical: true },
].map(p => ({ ...p, scope: '全国', verifiedAt: '2026-10-12', sourceKind: 'verified-catalog' }));

const ministries = ['gov.cn', 'ndrc.gov.cn', 'miit.gov.cn', 'mofcom.gov.cn', 'mot.gov.cn', 'mof.gov.cn', 'cac.gov.cn', 'mem.gov.cn', 'samr.gov.cn', 'most.gov.cn', 'mee.gov.cn', 'npc.gov.cn'];
export function nationalOfficialUrl(value) {
  try { const u = new URL(value); return ['http:', 'https:'].includes(u.protocol) && !u.username && !u.password && ministries.some(host => host === 'gov.cn' ? ['www.gov.cn', 'app.www.gov.cn', 'english.www.gov.cn'].includes(u.hostname) : u.hostname === host || u.hostname.endsWith('.' + host)); } catch { return false; }
}
const unescape = value => value.replace(/<!\[CDATA\[([\s\S]*?)\]\]>/g, '$1').replace(/<[^>]+>/g, ' ').replace(/&amp;/g, '&').replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&quot;/g, '"').replace(/&#39;/g, "'").replace(/\s+/g, ' ').trim();
const tag = (xml, name) => unescape(new RegExp(`<${name}[^>]*>([\\s\\S]*?)<\\/${name}>`, 'i').exec(xml)?.[1] || '');
const idFor = url => 'live-' + createHash('sha256').update(url).digest('hex').slice(0, 12);

export async function searchPolicies(query, extra = '') {
  const keywords = String(query).trim().slice(0, 100);
  if (!keywords) throw new Error('请填写政策检索关键词。');
  const context = keywords + ' ' + extra;
  const ranked = verifiedPolicies.map(p => ({ ...p, relevance: p.keywords.reduce((n, k) => n + (context.includes(k) ? 1 : 0), 0) - (p.historical ? 2 : 0) })).filter(p => p.relevance > 0).sort((a, b) => b.relevance - a.relevance);
  const catalog = ranked.length ? ranked : verifiedPolicies.filter(p => !p.historical).slice(0, 3).map(p => ({ ...p, relevance: 0, matchHint: '没有强匹配，需人工判断是否适用' }));
  let onlineStatus = 'unavailable', onlineMessage = '', live = [];
  try {
    const url = 'https://www.bing.com/search?format=rss&q=' + encodeURIComponent(`site:gov.cn ${keywords} 政策`);
    const response = await fetch(url, { signal: AbortSignal.timeout(12000), headers: { 'User-Agent': 'Mozilla/5.0' } });
    if (!response.ok) throw new Error('搜索服务返回 ' + response.status);
    const xml = await response.text();
    if (!/<rss[\s>]/i.test(xml)) throw new Error('搜索服务没有返回可读取的结果。');
    live = [...xml.matchAll(/<item[\s>]([\s\S]*?)<\/item>/gi)].map(m => ({ title: tag(m[1], 'title'), url: tag(m[1], 'link'), summary: tag(m[1], 'description').slice(0, 160) })).filter(p => nationalOfficialUrl(p.url) && p.title).slice(0, 8).map(p => ({ ...p, id: idFor(p.url), issuer: new URL(p.url).hostname, published: null, scope: '全国官方站点', sourceKind: 'live-search', historical: false, use: '检索摘要仅供找资料，发文日期、时效和适用条件请在原文中核对。' }));
    onlineStatus = live.length ? 'available' : 'no-official-result';
    onlineMessage = live.length ? `本次联网返回 ${live.length} 条全国官方站点结果。` : '本次联网检索没有返回可确认的全国官方结果，展示已核验资料库匹配。';
  } catch (error) { onlineMessage = `联网检索未完成（${error.message}），以下为已核验官方资料库的匹配结果。`; }
  const existing = new Set(catalog.map(p => p.url));
  return { query: keywords, searchedAt: new Date().toISOString(), scope: '全国官方政策', onlineStatus, onlineMessage, results: [...catalog, ...live.filter(p => !existing.has(p.url))].slice(0, 10), searchUrl: 'https://sousuo.www.gov.cn/sousuo/search.shtml?searchWord=' + encodeURIComponent(keywords) };
}
