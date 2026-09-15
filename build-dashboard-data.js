const fs = require('fs');
const zlib = require('zlib');
const XLSX = require('./xlsx.full.min.js');

const input = 'brand-research-summary-v9.xlsx';
const output = 'brand-research-summary-v9.json.gz';
const workbook = XLSX.read(fs.readFileSync(input), { type: 'buffer' });
const sheet = workbook.Sheets['品牌汇总'];
if (!sheet) throw new Error('未找到“品牌汇总”子表');
const rows = XLSX.utils.sheet_to_json(sheet, { defval: '' });

const fields = [
  ['袖子类型', ['袖型', '袖长']], ['面料类型', ['材质大类', '面料']], ['袖边类型', ['袖口']],
  ['领型'], ['风格', ['场合风格']], ['品类', ['类目', 'category', '大类']], ['小类'],
  ['是否衬衫'], ['版型'], ['弹力等级'], ['克重等级'], ['闭合方式'], ['商品状态', ['属性状态']],
  ['颜色'], ['品牌', ['brand', '品牌名']], ['标题'], ['竞品ASIN', ['竞品AS', 'ASIN', 'asin']], ['父ASIN'], ['抓取时间'],
  ['链接', ['url', 'link', '产品链接']], ['Buybox价格'], ['评分'], ['父体Rating数'], ['大类排名'],
  ['小类排名'], ['变体数'], ['FBA运费'], ['销售天数'], ['上架日期']
];

const safeNum = value => String(value ?? '').trim() === '92' ? 0 : Number.parseFloat(value) || 0;
const priceBand = value => { const p = safeNum(value); return !p ? '' : p < 5 ? '5以下' : p < 10 ? '5-9.99' : p < 15 ? '10-14.99' : p < 20 ? '15-19.99' : p < 25 ? '20-24.99' : p < 30 ? '25-29.99' : '30以上'; };
const ratingCountBand = value => { const n = safeNum(value); return n < 200 ? '0-199' : n < 500 ? '200-499' : n < 1000 ? '500-999' : n < 3000 ? '1000-2999' : n < 5000 ? '3000-4999' : '5000以上'; };
const ratingBand = value => { const n = safeNum(value); return !n ? '' : n < 4 ? '4分以下' : n >= 4.5 ? '4.5分以上' : n.toFixed(1) + '分'; };
const normalizeFit = value => { const raw = String(value || '').trim(), v = raw.toLowerCase(); if (!raw) return ''; if (['常规', '宽松', '修身', '弹力贴身', '运动版型'].includes(raw)) return raw; if (/straight/.test(v)) return '直筒'; if (/plus|big & tall/.test(v)) return '大码'; if (/tailored|fitted|slim/.test(v)) return '修身'; if (/casual|comfortable|moderate|overall|semi|regular|classic/.test(v)) return '常规'; return ''; };
const normalizeClosure = value => { const raw = String(value || '').trim(), v = raw.toLowerCase(); if (!raw) return ''; if (['套头', '纽扣', '抽绳', '拉链', '松紧', '钩扣', '按扣', '插扣', '系带'].includes(raw)) return raw; if (/button/.test(v)) return '纽扣'; if (/zip/.test(v)) return '拉链'; if (/drawstring/.test(v)) return '抽绳'; if (/lace|bow|strap/.test(v)) return '系带'; if (/clasp|ring/.test(v)) return '钩扣'; if (/open|no closure|no laces/.test(v)) return '无闭合'; return ''; };
const dateRank = value => { const m = String(value || '').match(/(20\d{2})[-\/]?(\d{1,2})[-\/]?(\d{1,2})/); return m ? Number(m[1]) * 10000 + Number(m[2]) * 100 + Number(m[3]) : 0; };

const yearCounts = {};
rows.forEach(row => { const m = String(row['抓取时间'] || '').match(/(20\d{2})/); if (m) yearCounts[m[1]] = (yearCounts[m[1]] || 0) + 1; });
const snapshotYear = Object.entries(yearCounts).sort((a, b) => b[1] - a[1])[0]?.[0] || '';
const snapshotMonth = row => { const m = String(row['月份'] || '').match(/(\d{1,2})月/); return m && snapshotYear ? snapshotYear + String(Number(m[1])).padStart(2, '0') : ''; };

const grouped = new Map();
let sameTimeConflicts = 0;
rows.forEach(row => {
  const month = snapshotMonth(row), parent = String(row['父ASIN'] || row['竞品ASIN'] || '').trim();
  if (!month || !parent) return;
  const key = month + '|' + parent, rank = dateRank(row['抓取时间']), current = grouped.get(key);
  if (!current || rank > current.rank) grouped.set(key, { row, rank, month });
  else if (rank === current.rank && safeNum(row['父体销量' + month]) !== safeNum(current.row['父体销量' + month])) sameTimeConflicts++;
});

const deduped = [...grouped.values()];
const snapshotMonths = [...new Set(deduped.map(item => item.month))].sort();
const latestMonth = snapshotMonths.at(-1);
const sourceMonths = [...new Set(Object.keys(rows[0] || {}).map(key => String(key).match(/^父体销量(20\d{4})$/)?.[1]).filter(Boolean))].sort();
const months = sourceMonths.filter(month => month <= latestMonth);

function mapRow(source, salesMonths, snapshotOnly) {
  const item = {};
  fields.forEach(([key, aliases = []]) => {
    const sourceKey = [key, ...aliases].find(name => Object.prototype.hasOwnProperty.call(source, name));
    const value = sourceKey ? String(source[sourceKey] ?? '').trim() : '';
    if (value) item[key] = value;
  });
  item['商品ASIN'] = item['竞品ASIN'] || '';
  item['竞品ASIN'] = item['父ASIN'] || item['竞品ASIN'] || '';
  item['版型'] = normalizeFit(item['版型']);
  item['闭合方式'] = normalizeClosure(item['闭合方式']);
  item['价格梯度'] = priceBand(item['Buybox价格']);
  item['评分数梯度'] = ratingCountBand(item['父体Rating数']);
  item['评分梯度'] = ratingBand(item['评分']);
  item.sales = {};
  if (snapshotOnly) { item._researchMonth = snapshotMonth(source); item.sales[item._researchMonth] = safeNum(source['父体销量' + item._researchMonth]); }
  else salesMonths.forEach(month => item.sales[month] = safeNum(source['父体销量' + month]));
  return item;
}

const latest = deduped.filter(item => item.month === latestMonth).map(item => mapRow(item.row, months, false));
const trend = deduped.map(item => mapRow(item.row, snapshotMonths, true));
const payload = {
  version: 'summary-v9', sourceName: '月度品牌调研_汇总_归类版.xlsx · 品牌汇总', months, trendMonths: snapshotMonths,
  latest, trend, qa: { sourceRows: rows.length, dedupedRows: deduped.length, latestParents: latest.length, snapshotMonths, sameTimeConflicts }
};
const json = JSON.stringify(payload);
fs.writeFileSync(output, zlib.gzipSync(json, { level: 9 }));
console.log(JSON.stringify({ output, jsonBytes: Buffer.byteLength(json), bytes: fs.statSync(output).size, ...payload.qa }));
