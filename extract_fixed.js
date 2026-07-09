const XLSX = require('xlsx');
const wb = XLSX.readFile('鄂尔多斯供电公司2026年十项民生工程7.2.xlsx');

function getVal(ws, row, col) {
  const addr = XLSX.utils.encode_cell({ r: row, c: col });
  return ws[addr] ? ws[addr].v : null;
}

// ========== 4. 频繁停电 ==========
{
  const ws = wb.Sheets['4频繁停电'];
  const rd = XLSX.utils.decode_range(ws['!ref']);
  let total = 0, done = 0, investing = 0, notstart = 0;
  let byCity = {};
  for (let r = rd.s.r + 2; r <= rd.e.r; r++) {
    const city = getVal(ws, r, 0);
    if (!city || typeof city !== 'string') continue;
    // Check if this is a data row by looking at column C (供电分公司)
    const branch = getVal(ws, r, 2);
    if (!branch) continue;
    total++;
    const key = city.substring(0, 4);
    if (!byCity[key]) byCity[key] = { total: 0, done: 0, investing: 0, notstart: 0 };
    byCity[key].total++;
    // Column T (index 19) = 是否完成治理, Column U (index 20) = 存在问题
    const isDone = getVal(ws, r, 19);
    const type = getVal(ws, r, 20);
    if (isDone === '是') { done++; byCity[key].done++; }
    else if (type && String(type).trim()) { investing++; byCity[key].investing++; }
    else { notstart++; byCity[key].notstart++; }
  }
  console.log('=== 频繁停电 ===');
  console.log('总数:', total, '完成:', done, '进行中:', investing, '未开始:', notstart);
  console.log('各地市:', JSON.stringify(byCity));
}

// ========== 5-2 树线矛盾 ==========
{
  const ws = wb.Sheets['5-2树线矛盾'];
  const rd = XLSX.utils.decode_range(ws['!ref']);
  let total = 0, done = 0;
  let byCity = {};
  for (let r = rd.s.r + 2; r <= rd.e.r; r++) {
    const city = getVal(ws, r, 0);
    if (!city || typeof city !== 'string') continue;
    total++;
    const key = city.substring(0, 4);
    if (!byCity[key]) byCity[key] = { total: 0, done: 0 };
    byCity[key].total++;
    // Column P (index 15) = 是否已完成整改
    const isDone = getVal(ws, r, 15);
    if (isDone === '是') { done++; byCity[key].done++; }
  }
  console.log('\n=== 树线矛盾 ===');
  console.log('总数:', total, '已完成:', done);
  console.log('各地市:', JSON.stringify(byCity));
}

// ========== 6 异常台区 ==========
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
    if (!byCity[key]) byCity[key] = { total: 0, done: 0, notstart: 0 };
    byCity[key].total++;
    // Column W (index 22) = 是否完成治理
    const isDone = getVal(ws, r, 22);
    if (isDone === '是') { done++; byCity[key].done++; }
    else { notstart++; byCity[key].notstart++; }
  }
  console.log('\n=== 异常台区 ===');
  console.log('总数:', total, '已完成:', done, '未完成:', notstart);
}

// ========== 9 跨产权小区 ==========
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
    // Column I (index 8) = 改造接收模式
    const mode = getVal(ws, r, 8);
    if (mode === '直接移交') direct++;
    else if (mode === '先改造后移交') reform++;
    else if (mode === '以建代改') build++;
    // Column R (index 17) = 竣工时间
    var w = getVal(ws, r, 17);
    if (w && String(w).trim()) { done++; byCity[key].done++; }
  }
  console.log('\n=== 跨产权小区 ===');
  console.log('总数:', total, '已完成:', done, '直接移交:', direct, '先改造后移交:', reform, '以建代改:', build);
}

// ========== 10 周期性保电场所 ==========
{
  const ws = wb.Sheets['10.周期性保电场所再梳理'];
  const rd = XLSX.utils.decode_range(ws['!ref']);
  let total = 0;
  let byCity = {};
  for (let r = rd.s.r + 4; r <= rd.e.r; r++) {
    const idx = getVal(ws, r, 0);
    if (typeof idx !== 'number') continue;
    total++;
    const city = getVal(ws, r, 1);
    if (city) {
      const key = String(city).substring(0, 4);
      if (!byCity[key]) byCity[key] = 0;
      byCity[key]++;
    }
  }
  console.log('\n=== 保电场所 ===');
  console.log('总数:', total);
  console.log('各地市:', JSON.stringify(byCity));
}
