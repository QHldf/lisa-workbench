const { execSync } = require('child_process');
const fs = require('fs');

const SEARCH_DIR = '/home/sandbox/.openclaw/workspace/skills/xiaoyi-web-search';

function search(query, count = 5) {
  console.log(`\n🔍 Searching: ${query}`);
  try {
    const result = execSync(
      `node ./scripts/search.js "${query}" -n ${count}`,
      { cwd: SEARCH_DIR, timeout: 60000, encoding: 'utf-8' }
    );
    return result;
  } catch (e) {
    console.error(`Search failed: ${e.message}`);
    return '';
  }
}

// Products to search, grouped by bank
const products = [
  // === 建设银行 ===
  { bank: '建设银行', product: '建易贷', query: '建设银行 建易贷 产品大纲 申请条件 征信' },

  // === 工商银行 ===
  { bank: '工商银行', product: '融e借', query: '工商银行 融e借 产品大纲 申请条件 征信' },
  { bank: '工商银行', product: 'e抵快贷', query: '工商银行 e抵快贷 产品大纲 申请条件 征信要求' },

  // === 农业银行 ===
  { bank: '农业银行', product: '小微网贷(纳税e贷)', query: '农业银行 纳税e贷 产品大纲 申请条件 征信' },
  { bank: '农业银行', product: '科技贷(原)', query: '农业银行 科技贷 原版 产品大纲 申请条件' },
  { bank: '农业银行', product: '抵押e贷', query: '农业银行 抵押e贷 产品大纲 申请条件 征信要求' },

  // === 中国银行 ===
  { bank: '中国银行', product: '惠如愿·信用贷', query: '中国银行 惠如愿 信用贷 产品大纲 申请条件' },
  { bank: '中国银行', product: '杭岗贷', query: '中国银行 杭岗贷 产品大纲 申请条件' },

  // === 交通银行 ===
  { bank: '交通银行', product: '普惠E贷', query: '交通银行 普惠E贷 产品大纲 申请条件' },
  { bank: '交通银行', product: '惠商贷', query: '交通银行 惠商贷 产品大纲 申请条件' },
  { bank: '交通银行', product: '线上税融通', query: '交通银行 线上税融通 产品大纲 申请条件' },
  { bank: '交通银行', product: '外贸贷/外贸快贷', query: '交通银行 外贸贷 外贸快贷 产品大纲 申请条件' },
  { bank: '交通银行', product: '展业e贷', query: '交通银行 展业e贷 产品大纲 申请条件' },
];

const results = [];
for (const p of products) {
  const output = search(p.query);
  results.push({ ...p, output });
  // Brief pause
  execSync('sleep 1');
}

fs.writeFileSync('search_results_raw.json', JSON.stringify(results, null, 2), 'utf-8');
console.log(`\n✅ Saved ${results.length} search results to search_results_raw.json`);
