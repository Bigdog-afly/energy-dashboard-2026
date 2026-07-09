const XLSX = require('xlsx');
const wb = XLSX.readFile('鄂尔多斯供电公司2026年十项民生工程7.2.xlsx');

function getVal(ws, row, col) {
  const addr = XLSX.utils.encode_cell({ r: row, c: col });
  return ws[addr] ? ws[addr].v : null;
}

const stats = {};

// ===== 1. 养老院 =====
{
  const ws = wb.Sheets['1养老院'];
  const rd = XLSX.utils.decode_range(ws['!ref']);
  let total = 0, done = 0, investing = 0, notstart = 0;
  let byCity = {};
  let totalInvest = 0;
  for (let r = rd.s.r + 2; r <= rd.e.r; r++) {
    const idx = getVal(ws, r, 0);
    const city = getVal(ws, r, 1);
    const invest = getVal(ws, r, 5);
    const type = getVal(ws, r, 7);
    const isDone = getVal(ws, r, 8);
    if (typeof idx !== 'number' || !city) continue;
    total++;
    totalInvest += invest || 0;
    const key = String(city).substring(0, 4);
    if (!byCity[key]) byCity[key] = { total: 0, done: 0 };
    byCity[key].total++;
    if (isDone === '是') { done++; byCity[key].done++; }
    else if (type && String(type).trim()) investing++;
    else notstart++;
  }
  stats.养老院 = { total, done, investing, notstart, totalInvest: totalInvest.toFixed(0) };
  stats.养老院.byCity = byCity;
}

// ===== 2. 高层小区 =====
{
  const ws = wb.Sheets['2高层小区'];
  const rd = XLSX.utils.decode_range(ws['!ref']);
  let total = 0, done = 0, investing = 0, notstart = 0;
  let byCity = {};
  for (let r = rd.s.r + 2; r <= rd.e.r; r++) {
    const idx = getVal(ws, r, 0);
    const city = getVal(ws, r, 1);
    const type = getVal(ws, r, 8);
    const isDone = getVal(ws, r, 9);
    if (typeof idx !== 'number' || !city) continue;
    total++;
    const key = String(city).substring(0, 4);
    if (!byCity[key]) byCity[key] = { total: 0, done: 0 };
    byCity[key].total++;
    if (isDone === '是') { done++; byCity[key].done++; }
    else if (type && String(type).trim()) investing++;
    else notstart++;
  }
  stats.高层小区 = { total, done, investing, notstart };
  stats.高层小区.byCity = byCity;
}

// ===== 3. 标准化考场 =====
{
  const ws = wb.Sheets['3标准化考场'];
  const rd = XLSX.utils.decode_range(ws['!ref']);
  let total = 0, done = 0, investing = 0, notstart = 0;
  let byCity = {};
  for (let r = rd.s.r + 2; r <= rd.e.r; r++) {
    const idx = getVal(ws, r, 0);
    const city = getVal(ws, r, 1);
    const type = getVal(ws, r, 8);
    const isDone = getVal(ws, r, 9);
    if (typeof idx !== 'number' || !city) continue;
    total++;
    const key = String(city).substring(0, 4);
    if (!byCity[key]) byCity[key] = { total: 0, done: 0 };
    byCity[key].total++;
    if (isDone === '是') { done++; byCity[key].done++; }
    else if (type && String(type).trim()) investing++;
    else notstart++;
  }
  stats.标准化考场 = { total, done, investing, notstart };
  stats.标准化考场.byCity = byCity;
}

// ===== 4. 频繁停电 =====
{
  const ws = wb.Sheets['4频繁停电'];
  const rd = XLSX.utils.decode_range(ws['!ref']);
  let total = 0, done = 0, investing = 0, notstart = 0;
  let byCity = {};
  for (let r = rd.s.r + 2; r <= rd.e.r; r++) {
    const city = getVal(ws, r, 0);
    const isDone = getVal(ws, r, 19);
    const type = getVal(ws, r, 20);
    if (!city || typeof city !== 'string') continue;
    total++;
    const key = city.substring(0, 4);
    if (!byCity[key]) byCity[key] = { total: 0, done: 0 };
    byCity[key].total++;
    if (isDone === '是') { done++; byCity[key].done++; }
    else if (type && String(type).trim()) investing++;
    else notstart++;
  }
  stats.频繁停电 = { total, done, investing, notstart };
  stats.频繁停电.byCity = byCity;
}

// ===== 5-1 局部绝缘化 =====
{
  const ws = wb.Sheets['5-1局部绝缘化'];
  const rd = XLSX.utils.decode_range(ws['!ref']);
  let total = 0, done = 0;
  let byCity = {};
  for (let r = rd.s.r + 2; r <= rd.e.r; r++) {
    const branch = getVal(ws, r, 2);
    const done2 = getVal(ws, r, 20);
    if (!branch || typeof branch !== 'string') continue;
    total++;
    const key = branch.substring(0, 4);
    if (!byCity[key]) byCity[key] = { total: 0, done: 0 };
    byCity[key].total++;
    if (done2 && String(done2).trim()) { done++; byCity[key].done++; }
  }
  stats.局部绝缘化 = { total, done };
  stats.局部绝缘化.byCity = byCity;
}

// ===== 5-2 树线矛盾 =====
{
  const ws = wb.Sheets['5-2树线矛盾'];
  const rd = XLSX.utils.decode_range(ws['!ref']);
  let total = 0, done = 0;
  let byCity = {};
  for (let r = rd.s.r + 2; r <= rd.e.r; r++) {
    const city = getVal(ws, r, 0);
    if (typeof city !== 'string') continue;
    total++;
    const key = city.substring(0, 4);
    if (!byCity[key]) byCity[key] = { total: 0, done: 0 };
    byCity[key].total++;
    const isDone = getVal(ws, r, 15);
    if (isDone === '是') { done++; byCity[key].done++; }
  }
  stats.树线矛盾 = { total, done };
  stats.树线矛盾.byCity = byCity;
}

// ===== 6 异常台区 =====
{
  const ws = wb.Sheets['6异常台区'];
  const rd = XLSX.utils.decode_range(ws['!ref']);
  let total = 0, done = 0, notstart = 0;
  let byCity = {};
  for (let r = rd.s.r + 2; r <= rd.e.r; r++) {
    const idx = getVal(ws, r, 0);
    const branch = getVal(ws, r, 2);
    if (typeof idx !== 'number' || !branch) continue;
    total++;
    const key = branch.substring(0, 4);
    if (!byCity[key]) byCity[key] = { total: 0, done: 0 };
    byCity[key].total++;
    const isDone = getVal(ws, r, 22);
    if (isDone === '是') { done++; byCity[key].done++; }
    else notstart++;
  }
  stats.异常台区 = { total, done, notstart };
  stats.异常台区.byCity = byCity;
}

// ===== 7 示范区 =====
{
  const ws = wb.Sheets['7.示范区'];
  const rd = XLSX.utils.decode_range(ws['!ref']);
  let total = 0, totalInvest = 0, totalCompletionRate = 0;
  let byCity = {};
  for (let r = rd.s.r + 2; r <= rd.e.r; r++) {
    const idx = getVal(ws, r, 0);
    if (typeof idx !== 'number') continue;
    total++;
    const city = getVal(ws, r, 1);
    const invest = getVal(ws, r, 4);
    const completion = getVal(ws, r, 30);
    totalInvest += invest || 0;
    if (completion != null) totalCompletionRate += completion;
    const key = String(city).substring(0, 4);
    if (!byCity[key]) byCity[key] = { total: 0, completion: 0 };
    byCity[key].total++;
    if (completion != null) byCity[key].completion += completion;
  }
  stats.示范区 = {
    total,
    totalInvest: totalInvest.toFixed(0),
    avgCompletion: total ? ((totalCompletionRate / total) * 100).toFixed(1) : 0
  };
  stats.示范区.byCity = byCity;
}

// ===== 9 跨产权小区 =====
{
  const ws = wb.Sheets['9.跨产权供电小区'];
  const rd = XLSX.utils.decode_range(ws['!ref']);
  let total = 0, done = 0, direct = 0, reform = 0, build = 0;
  let byCity = {};
  for (let r = rd.s.r + 3; r <= rd.e.r; r++) {
    const name = getVal(ws, r, 3);
    if (!name || String(name).startsWith('附') || String(name).startsWith('错')) continue;
    total++;
    const city = getVal(ws, r, 1);
    const key = String(city).substring(0, 4);
    if (!byCity[key]) byCity[key] = { total: 0, done: 0 };
    byCity[key].total++;
    const mode = getVal(ws, r, 8);
    if (mode === '直接移交') direct++;
    else if (mode === '先改造后移交') reform++;
    else if (mode === '以建代改') build++;
    var w = getVal(ws, r, 17);
    if (w && String(w).trim()) { done++; byCity[key].done++; }
  }
  stats.跨产权小区 = { total, done, direct, reform, build };
  stats.跨产权小区.byCity = byCity;
}

// ===== 10 周期性保电场所 =====
{
  const ws = wb.Sheets['10.周期性保电场所再梳理'];
  const rd = XLSX.utils.decode_range(ws['!ref']);
  let total = 0, done = 0;
  for (let r = rd.s.r + 4; r <= rd.e.r; r++) {
    const idx = getVal(ws, r, 0);
    if (typeof idx !== 'number') continue;
    total++;
  }
  stats.保电场所 = { total, done };
}

// ===== 11 农牧区线路 =====
{
  const ws = wb.Sheets['11.1700余条农牧区配电线路清单'];
  const rd = XLSX.utils.decode_range(ws['!ref']);
  let total = 0, done = 0, totalInvest = 0;
  for (let r = rd.s.r + 1; r <= rd.e.r; r++) {
    const unit = getVal(ws, r, 0);
    if (typeof unit !== 'number') continue;
    total++;
    const invest = getVal(ws, r, 5);
    totalInvest += invest || 0;
    const isDone = getVal(ws, r, 11);
    if (isDone === '是') done++;
  }
  stats.农牧区线路 = { total, done, totalInvest: totalInvest.toFixed(0) };
}

// ===== Output summary =====
const output = {};
for (const key of Object.keys(stats)) {
  const s = stats[key];
  output[key] = {
    ...s,
    donePct: s.total ? ((s.done / s.total) * 100).toFixed(1) : 0,
    byCity: s.byCity
  };
}
console.log(JSON.stringify(output, null, 2));
