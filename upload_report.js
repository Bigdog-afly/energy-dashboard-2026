/**
 * upload_report.js - 周报上传与数据处理脚本
 *
 * 用法:
 *   node upload_report.js                          # 处理默认文件
 *   node upload_report.js "path/to/report.xlsx"    # 指定文件
 *   node upload_report.js --dry-run                # 预览变更不写入
 */

const XLSX = require('xlsx');
const fs = require('fs');
const path = require('path');

const BASE_DIR = __dirname;
const DATA_FILE = path.join(BASE_DIR, 'dashboard_data.json');
const HISTORY_DIR = path.join(BASE_DIR, 'report_history');
const CHANGELOG_FILE = path.join(HISTORY_DIR, 'changes.md');

// ========== 工具函数 ==========
function gv(ws, r, c) {
  const addr = XLSX.utils.encode_cell({ r, c });
  return ws[addr] ? ws[addr].v : null;
}

function getSheetByName(wb, names) {
  if (!Array.isArray(names)) names = [names];
  for (const name of names) {
    if (wb.Sheets[name]) return wb.Sheets[name];
  }
  return null;
}

function safeNum(val, def = 0) {
  if (val == null) return def;
  const n = Number(val);
  return isNaN(n) ? def : n;
}

function todayStr() {
  return new Date().toISOString().slice(0, 10);
}

function timestampStr() {
  const d = new Date();
  return d.toISOString().slice(0, 10) + '_' +
    String(d.getHours()).padStart(2, '0') +
    String(d.getMinutes()).padStart(2, '0') +
    String(d.getSeconds()).padStart(2, '0');
}

// ========== 各工程解析函数 ==========
function parseSection养老院(ws) {
  const rd = XLSX.utils.decode_range(ws['!ref']);
  let total = 0, done = 0, investing = 0, notstart = 0, totalInvest = 0;
  const byCity = {};
  for (let r = rd.s.r + 2; r <= rd.e.r; r++) {
    const idx = gv(ws, r, 0);
    const city = gv(ws, r, 1);
    if (typeof idx !== 'number' || !city) continue;
    total++;
    totalInvest += safeNum(gv(ws, r, 5));
    const tp = gv(ws, r, 7);
    const is = gv(ws, r, 8);
    const key = String(city).substring(0, 4);
    if (!byCity[key]) byCity[key] = { t: 0, d: 0, inv: 0, nst: 0 };
    byCity[key].t++;
    if (is === '是') { done++; byCity[key].d++; }
    else if (tp && String(tp).trim()) { investing++; byCity[key].inv++; }
    else { notstart++; byCity[key].nst++; }
  }
  return { t: total, d: done, inv: investing, nst: notstart, totalInvest: totalInvest.toFixed(0), pct: (total ? (done / total * 100).toFixed(1) : '0.0'), bc: byCity };
}

function parseSection高层小区(ws) {
  const rd = XLSX.utils.decode_range(ws['!ref']);
  let total = 0, done = 0, investing = 0, notstart = 0;
  const byCity = {};
  for (let r = rd.s.r + 2; r <= rd.e.r; r++) {
    const idx = gv(ws, r, 0);
    const city = gv(ws, r, 1);
    if (typeof idx !== 'number' || !city) continue;
    total++;
    const tp = gv(ws, r, 8);
    const is = gv(ws, r, 9);
    const key = String(city).substring(0, 4);
    if (!byCity[key]) byCity[key] = { t: 0, d: 0, inv: 0, nst: 0 };
    byCity[key].t++;
    if (is === '是') { done++; byCity[key].d++; }
    else if (tp && String(tp).trim()) { investing++; byCity[key].inv++; }
    else { notstart++; byCity[key].nst++; }
  }
  return { t: total, d: done, inv: investing, nst: notstart, pct: (total ? (done / total * 100).toFixed(1) : '0.0'), bc: byCity };
}

function parseSection标准化考场(ws) {
  const rd = XLSX.utils.decode_range(ws['!ref']);
  let total = 0, done = 0, investing = 0, notstart = 0;
  const byCity = {};
  for (let r = rd.s.r + 2; r <= rd.e.r; r++) {
    const idx = gv(ws, r, 0);
    const city = gv(ws, r, 1);
    if (typeof idx !== 'number' || !city) continue;
    total++;
    const tp = gv(ws, r, 8);
    const is = gv(ws, r, 9);
    const key = String(city).substring(0, 4);
    if (!byCity[key]) byCity[key] = { t: 0, d: 0, inv: 0, nst: 0 };
    byCity[key].t++;
    if (is === '是') { done++; byCity[key].d++; }
    else if (tp && String(tp).trim()) { investing++; byCity[key].inv++; }
    else { notstart++; byCity[key].nst++; }
  }
  return { t: total, d: done, inv: investing, nst: notstart, pct: (total ? (done / total * 100).toFixed(1) : '0.0'), bc: byCity };
}

function parseSection频繁停电(ws) {
  const rd = XLSX.utils.decode_range(ws['!ref']);
  let total = 0, done = 0, investing = 0, notstart = 0;
  const byCity = {};
  for (let r = rd.s.r + 2; r <= rd.e.r; r++) {
    const city = gv(ws, r, 0);
    const branch = gv(ws, r, 2);
    if (!city || typeof city !== 'string' || !branch) continue;
    total++;
    const is = gv(ws, r, 17);
    const pd = gv(ws, r, 18);
    const key = city.substring(0, 4);
    if (!byCity[key]) byCity[key] = { t: 0, d: 0, inv: 0, nst: 0 };
    byCity[key].t++;
    if (is === '是') { done++; byCity[key].d++; }
    else if (pd && typeof pd === 'number' && pd > 46000) { investing++; byCity[key].inv++; }
    else { notstart++; byCity[key].nst++; }
  }
  return { t: total, d: done, inv: investing, nst: notstart, pct: (total ? (done / total * 100).toFixed(1) : '0.0'), bc: byCity };
}

function parseSection局部绝缘化(ws) {
  const rd = XLSX.utils.decode_range(ws['!ref']);
  let total = 0, done = 0;
  const byCity = {};
  for (let r = rd.s.r + 2; r <= rd.e.r; r++) {
    const br = gv(ws, r, 2);
    if (!br || typeof br !== 'string') continue;
    total++;
    const key = br.substring(0, 4);
    if (!byCity[key]) byCity[key] = { t: 0, d: 0 };
    byCity[key].t++;
    if (gv(ws, r, 20) && String(gv(ws, r, 20)).trim()) { done++; byCity[key].d++; }
  }
  return { t: total, d: done, pct: (total ? (done / total * 100).toFixed(1) : '0.0'), bc: byCity };
}

function parseSection树线矛盾(ws) {
  const rd = XLSX.utils.decode_range(ws['!ref']);
  let total = 0, done = 0;
  const byCity = {};
  for (let r = rd.s.r + 2; r <= rd.e.r; r++) {
    const idx = gv(ws, r, 0);
    if (typeof idx !== 'number') continue;
    total++;
    const city = gv(ws, r, 1);
    const key = String(city).substring(0, 4);
    if (!byCity[key]) byCity[key] = { t: 0, d: 0 };
    byCity[key].t++;
    if (gv(ws, r, 15) === '是') { done++; byCity[key].d++; }
  }
  return { t: total, d: done, pct: (total ? (done / total * 100).toFixed(1) : '0.0'), bc: byCity };
}

function parseSection异常台区(ws) {
  const rd = XLSX.utils.decode_range(ws['!ref']);
  let total = 0, done = 0, notstart = 0;
  const byCity = {};
  for (let r = rd.s.r + 2; r <= rd.e.r; r++) {
    const idx = gv(ws, r, 0);
    const branch = gv(ws, r, 2);
    if (typeof idx !== 'number' || !branch) continue;
    total++;
    const key = branch.substring(0, 4);
    if (!byCity[key]) byCity[key] = { t: 0, d: 0, nst: 0 };
    byCity[key].t++;
    if (gv(ws, r, 22) === '是') { done++; byCity[key].d++; }
    else { notstart++; byCity[key].nst++; }
  }
  return { t: total, d: done, nst: notstart, pct: (total ? (done / total * 100).toFixed(1) : '0.0'), bc: byCity };
}

function parseSection示范区(ws) {
  const rd = XLSX.utils.decode_range(ws['!ref']);
  let total = 0, totalInvest = 0, totalCompletion = 0;
  const byCity = {};
  for (let r = rd.s.r + 2; r <= rd.e.r; r++) {
    const idx = gv(ws, r, 0);
    if (typeof idx !== 'number') continue;
    total++;
    const invest = safeNum(gv(ws, r, 4));
    const comp = safeNum(gv(ws, r, 30));
    totalInvest += invest;
    if (gv(ws, r, 30) != null) totalCompletion += safeNum(gv(ws, r, 30));
    const key = String(gv(ws, r, 1)).substring(0, 4);
    if (!byCity[key]) byCity[key] = { t: 0, c: 0 };
    byCity[key].t++;
    if (gv(ws, r, 30) != null) byCity[key].c += comp;
  }
  return { t: total, totalInvest: totalInvest.toFixed(0), avgComp: (total ? (totalCompletion / total * 100).toFixed(1) : '0.0'), bc: byCity };
}

function parseSection跨产权小区(ws) {
  const rd = XLSX.utils.decode_range(ws['!ref']);
  let total = 0, done = 0, direct = 0, reform = 0, build = 0;
  const byCity = {};
  for (let r = rd.s.r + 3; r <= rd.e.r; r++) {
    const name = gv(ws, r, 3);
    if (!name || String(name).startsWith('附') || String(name).startsWith('错')) continue;
    total++;
    const city = gv(ws, r, 1);
    const key = String(city).substring(0, 4);
    if (!byCity[key]) byCity[key] = { t: 0, d: 0 };
    byCity[key].t++;
    const mode = gv(ws, r, 8);
    if (mode === '直接移交') direct++;
    else if (mode === '先改造后移交') reform++;
    else if (mode === '以建代改') build++;
    if (gv(ws, r, 17) && String(gv(ws, r, 17)).trim()) { done++; byCity[key].d++; }
  }
  return { t: total, d: done, dir: direct, ref: reform, bld: build, pct: (total ? (done / total * 100).toFixed(1) : '0.0'), bc: byCity };
}

function parseSection保电场所(ws) {
  const rd = XLSX.utils.decode_range(ws['!ref']);
  let total = 0;
  const byCity = {};
  for (let r = rd.s.r + 4; r <= rd.e.r; r++) {
    const idx = gv(ws, r, 0);
    if (typeof idx !== 'number') continue;
    total++;
    const city = gv(ws, r, 1);
    if (city) {
      const key = String(city).substring(0, 4);
      byCity[key] = (byCity[key] || 0) + 1;
    }
  }
  return { t: total, bc: byCity };
}

function parseSection农牧区线路(ws) {
  const rd = XLSX.utils.decode_range(ws['!ref']);
  let total = 0, done = 0, totalInvest = 0;
  const byCity = {};
  for (let r = rd.s.r + 1; r <= rd.e.r; r++) {
    const unit = gv(ws, r, 0);
    if (typeof unit !== 'number') continue;
    total++;
    totalInvest += safeNum(gv(ws, r, 5));
    const key = String(gv(ws, r, 1)).substring(0, 4);
    if (!byCity[key]) byCity[key] = { t: 0, d: 0 };
    byCity[key].t++;
    if (gv(ws, r, 11) === '是') { done++; byCity[key].d++; }
  }
  return { t: total, d: done, totalInvest: totalInvest.toFixed(0), pct: (total ? (done / total * 100).toFixed(1) : '0.0'), bc: byCity };
}

// 保电场所解析函数（如果 sheet 不存在则返回空数据）
function parseSection保电场所(ws) {
  if (!ws) return { t: 0, bc: {} };
  const rd = XLSX.utils.decode_range(ws['!ref']);
  let total = 0;
  const byCity = {};
  for (let r = rd.s.r + 4; r <= rd.e.r; r++) {
    const idx = gv(ws, r, 0);
    if (typeof idx !== 'number') continue;
    total++;
    const city = gv(ws, r, 1);
    if (city) {
      const key = String(city).substring(0, 4);
      byCity[key] = (byCity[key] || 0) + 1;
    }
  }
  return { t: total, bc: byCity };
}

// ========== 主处理流程 ==========
function processExcel(filePath) {
  console.log('='.repeat(60));
  console.log('周报数据处理脚本');
  console.log('='.repeat(60));
  console.log('输入文件:', filePath);

  // 读取 Excel
  const wb = XLSX.readFile(filePath);
  console.log('工作表列表:', wb.SheetNames.join(', '));
  console.log('');

  const dashboard = {
    updateTime: todayStr(),
    sections: {}
  };

  // 解析各工程
  const parsers = {
    '养老院': { sheet: '1养老院', fn: parseSection养老院 },
    '高层小区': { sheet: '2高层小区', fn: parseSection高层小区 },
    '标准化考场': { sheet: '3标准化考场', fn: parseSection标准化考场 },
    '频繁停电': { sheet: '4频繁停电', fn: parseSection频繁停电 },
    '局部绝缘化': { sheet: ['5-1局部绝缘化'], fn: parseSection局部绝缘化 },
    '树线矛盾': { sheet: ['5-2树线矛盾'], fn: parseSection树线矛盾 },
    '异常台区': { sheet: '6异常台区', fn: parseSection异常台区 },
    '示范区': { sheet: '7.示范区', fn: parseSection示范区 },
    '跨产权小区': { sheet: '9.跨产权供电小区', fn: parseSection跨产权小区 },
    '保电场所': { sheet: ['10.周期性保电场所再梳理', '10保电场所', '保电场所', '8-1引领区工程进度'], fn: parseSection保电场所 },
    '农牧区线路': { sheet: ['11.1700余条农牧区配电线路清单', '10.1700余条农牧区配电线路清单'], fn: parseSection农牧区线路 }
  };

  for (const [name, cfg] of Object.entries(parsers)) {
    const ws = getSheetByName(wb, cfg.sheet);
    if (!ws) {
      console.log(`  ⚠️  跳过 ${name}: 未找到工作表 "${cfg.sheet}"`);
      dashboard.sections[name] = { t: 0, pct: '0.0' };
      continue;
    }
    try {
      dashboard.sections[name] = cfg.fn(ws);
      const pct = dashboard.sections[name].pct || '0.0';
      console.log(`  ✅ ${name}: ${dashboard.sections[name].t}项, 完成率 ${pct}%`);
    } catch (err) {
      console.log(`  ❌ ${name}: 解析错误 - ${err.message}`);
      dashboard.sections[name] = { t: 0, pct: '0.0', error: err.message };
    }
  }

  // 计算汇总
  let grandTotal = 0, grandDone = 0, grandInvesting = 0, grandNotstart = 0, grandInvest = 0;
  for (const [name, s] of Object.entries(dashboard.sections)) {
    grandTotal += s.t || 0;
    grandDone += s.d || 0;
    grandInvesting += s.inv || 0;
    grandNotstart += s.nst || 0;
    grandInvest += parseFloat(s.totalInvest) || 0;
  }
  dashboard.grand = {
    total: grandTotal,
    done: grandDone,
    investing: grandInvesting,
    notstart: grandNotstart,
    invest: grandInvest.toFixed(0),
    pct: (grandTotal ? (grandDone / grandTotal * 100).toFixed(1) : '0.0')
  };

  return dashboard;
}

// ========== 变更检测 ==========
function detectChanges(newData) {
  let changes = [];
  let oldData = null;

  // 读取旧数据
  if (fs.existsSync(DATA_FILE)) {
    try {
      oldData = JSON.parse(fs.readFileSync(DATA_FILE, 'utf8'));
    } catch (e) { /* ignore */ }
  }

  if (!oldData) {
    return { hasChanges: true, summary: '首次生成数据', details: [] };
  }

  const oldGrand = oldData.grand;
  const newGrand = newData.grand;

  // 总体变化
  if (oldGrand.pct !== newGrand.pct) {
    changes.push(`总体完成率: ${oldGrand.pct}% → ${newGrand.pct}% (${(parseFloat(newGrand.pct) - parseFloat(oldGrand.pct)).toFixed(1)}%)`);
  }

  // 各工程变化
  const sectionNames = Object.keys(newData.sections);
  for (const name of sectionNames) {
    const oldPct = oldData.sections[name]?.pct;
    const newPct = newData.sections[name]?.pct;
    if (oldPct != null && newPct != null && oldPct !== newPct) {
      const diff = (parseFloat(newPct) - parseFloat(oldPct)).toFixed(1);
      const arrow = diff > 0 ? '↑' : diff < 0 ? '↓' : '=';
      changes.push(`${name}: ${oldPct}% → ${newPct}% (${arrow}${diff}%)`);
    }
  }

  return {
    hasChanges: changes.length > 0,
    summary: `发现 ${changes.length} 项变化`,
    details: changes
  };
}

// ========== 历史记录管理 ==========
function saveHistory(newData, changes) {
  if (!fs.existsSync(HISTORY_DIR)) {
    fs.mkdirSync(HISTORY_DIR, { recursive: true });
  }

  // 保存快照
  const snapshotFile = path.join(HISTORY_DIR, timestampStr() + '.json');
  fs.writeFileSync(snapshotFile, JSON.stringify(newData, null, 2));

  // 更新 changelog
  const ts = new Date().toLocaleString('zh-CN');
  let logEntry = `\n## ${ts}\n\n${changes.summary}\n\n`;
  changes.details.forEach(d => { logEntry += `- ${d}\n`; });
  logEntry += '\n';

  if (!fs.existsSync(CHANGELOG_FILE)) {
    fs.writeFileSync(CHANGELOG_FILE, '# 周报变更记录\n\n' + logEntry);
  } else {
    fs.writeFileSync(CHANGELOG_FILE, fs.readFileSync(CHANGELOG_FILE, 'utf8') + logEntry);
  }

  return snapshotFile;
}

// ========== 主入口 ==========
function main() {
  const args = process.argv.slice(2);
  const dryRun = args.includes('--dry-run');
  const filePath = args.find(a => a.endsWith('.xlsx') || a.endsWith('.xls')) || '鄂尔多斯供电公司2026年十项民生工程7.2.xlsx';
  const fullPath = path.resolve(BASE_DIR, filePath);

  if (!fs.existsSync(fullPath)) {
    console.error(`❌ 文件不存在: ${fullPath}`);
    process.exit(1);
  }

  // 处理 Excel
  const newData = processExcel(fullPath);

  // 变更检测
  const changes = detectChanges(newData);

  console.log('\n' + '='.repeat(60));
  console.log('变更检测: ' + changes.summary);
  changes.details.forEach(d => console.log('  ' + d));
  console.log('='.repeat(60));

  if (dryRun) {
    console.log('\n[预览模式] 未写入文件');
    console.log('如需实际更新，请移除 --dry-run 参数');
    return;
  }

  // 保存历史快照
  const snapshotFile = saveHistory(newData, changes);
  console.log('\n历史快照:', snapshotFile);

  // 写入新数据
  fs.writeFileSync(DATA_FILE, JSON.stringify(newData, null, 2));
  console.log('数据文件已更新:', DATA_FILE);

  // 控制台总结
  console.log('\n📊 总体进度: ' + newData.grand.pct + '% (' + newData.grand.done + '/' + newData.grand.total + ' 项已完成)');
  console.log('💰 总投资: ' + Number(newData.grand.invest).toLocaleString() + ' 万元');

  // Top 3 提升
  const improving = changes.details.filter(d => d.includes('↑')).sort((a, b) => {
    const aDiff = parseFloat(a.match(/↑([\d.]+)%/)?.[1] || 0);
    const bDiff = parseFloat(b.match(/↑([\d.]+)%/)?.[1] || 0);
    return bDiff - aDiff;
  }).slice(0, 3);

  const declining = changes.details.filter(d => d.includes('↓')).sort((a, b) => {
    const aDiff = parseFloat(a.match(/↓([\d.]+)%/)?.[1] || 0);
    const bDiff = parseFloat(b.match(/↓([\d.]+)%/)?.[1] || 0);
    return aDiff - bDiff;
  }).slice(0, 3);

  if (improving.length > 0) {
    console.log('\n📈 提升最多的工程:');
    improving.forEach(d => console.log('  +' + d));
  }
  if (declining.length > 0) {
    console.log('\n📉 下降的工程:');
    declining.forEach(d => console.log('  ' + d));
  }
}

main();
