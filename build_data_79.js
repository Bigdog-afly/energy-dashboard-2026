const XLSX = require('xlsx');
const fs = require('fs');
const path = require('path');

const wb = XLSX.readFile('十项民生工程进度调度表7.9 报万处.xlsx');
function gv(ws,r,c) { const a=XLSX.utils.encode_cell({r,c}); return ws[a]?ws[a].v:null; }

// ========== 解析函数 ==========
function parseSection养老院(ws) {
  const rd = XLSX.utils.decode_range(ws['!ref']);
  let t=0,d=0,inv=0,nst=0,ti=0;
  const bc = {};
  // 数据从R2开始，列：序号(0) 供电分公司(1) ... 投资金额(4) ... 改造项目类型(6) 是否完成治理(7)
  for(let r=rd.s.r+2; r<=rd.e.r; r++) {
    const idx = gv(ws,r,0);
    if(typeof idx !== 'number') continue;
    t++;
    ti += (gv(ws,r,4)||0);
    const city = gv(ws,r,1);
    const k = String(city).substring(0,4);
    if(!bc[k]) bc[k]={t:0,d:0,inv:0,nst:0};
    bc[k].t++;
    const isDone = gv(ws,r,7);
    const projType = gv(ws,r,6);
    if(isDone==='是') { d++; bc[k].d++; }
    else if(projType && String(projType).trim()) { inv++; bc[k].inv++; }
    else { nst++; bc[k].nst++; }
  }
  // 过滤只保留鄂尔多斯
  const filteredBC = {};
  Object.entries(bc).forEach(([k,v]) => {
    if(k.includes('鄂尔多斯')) filteredBC[k] = v;
  });
  return { t, d, inv, nst, totalInvest: ti.toFixed(0), pct: (t?(d/t*100).toFixed(1):'0.0'), bc: filteredBC };
}

function parseSection高层小区(ws) {
  const rd = XLSX.utils.decode_range(ws['!ref']);
  let t=0,d=0,inv=0,nst=0;
  const bc = {};
  // 数据从R2开始，列：序号(0) 供电分公司(1) ... 是否完成治理(8) 计划完成治理日期(6) 改造项目类型(7)
  for(let r=rd.s.r+2; r<=rd.e.r; r++) {
    const idx = gv(ws,r,0);
    if(typeof idx !== 'number') continue;
    t++;
    const city = gv(ws,r,1);
    const k = String(city).substring(0,4);
    if(!bc[k]) bc[k]={t:0,d:0,inv:0,nst:0};
    bc[k].t++;
    const isDone = gv(ws,r,8);
    const projType = gv(ws,r,7);
    if(isDone==='是') { d++; bc[k].d++; }
    else if(projType && String(projType).trim()) { inv++; bc[k].inv++; }
    else { nst++; bc[k].nst++; }
  }
  const filteredBC = {};
  Object.entries(bc).forEach(([k,v]) => { if(k.includes('鄂尔多斯')) filteredBC[k]=v; });
  return { t, d, inv, nst, pct: (t?(d/t*100).toFixed(1):'0.0'), bc: filteredBC };
}

function parseSection标准化考场(ws) {
  const rd = XLSX.utils.decode_range(ws['!ref']);
  let t=0,d=0,inv=0,nst=0;
  const bc = {};
  // 数据从R2开始，列：序号(0) 归属供电分公司(1) ... 是否完成治理(8) 计划完成治理日期(6) 改造项目类型(7)
  for(let r=rd.s.r+2; r<=rd.e.r; r++) {
    const idx = gv(ws,r,0);
    if(typeof idx !== 'number') continue;
    t++;
    const city = gv(ws,r,1);
    const k = String(city).substring(0,4);
    if(!bc[k]) bc[k]={t:0,d:0,inv:0,nst:0};
    bc[k].t++;
    const isDone = gv(ws,r,8);
    const projType = gv(ws,r,7);
    if(isDone==='是') { d++; bc[k].d++; }
    else if(projType && String(projType).trim()) { inv++; bc[k].inv++; }
    else { nst++; bc[k].nst++; }
  }
  const filteredBC = {};
  Object.entries(bc).forEach(([k,v]) => { if(k.includes('鄂尔多斯')) filteredBC[k]=v; });
  return { t, d, inv, nst, pct: (t?(d/t*100).toFixed(1):'0.0'), bc: filteredBC };
}

function parseSection频繁停电(ws) {
  const rd = XLSX.utils.decode_range(ws['!ref']);
  let t=0,d=0,inv=0,nst=0,ti=0;
  const bc = {};
  // 数据从R2开始，列：线段所在单位(0) ... 是否完成治理(17) 治理完成累计投入资金(20)
  for(let r=rd.s.r+2; r<=rd.e.r; r++) {
    const city = gv(ws,r,0);
    if(!city || typeof city !== 'string') continue;
    t++;
    ti += (gv(ws,r,20)||0);
    const k = city.substring(0,4);
    if(!bc[k]) bc[k]={t:0,d:0,inv:0,nst:0};
    bc[k].t++;
    const isDone = gv(ws,r,17);
    if(isDone==='是') { d++; bc[k].d++; }
    else { nst++; bc[k].nst++; }
  }
  const filteredBC = {};
  Object.entries(bc).forEach(([k,v]) => { if(k.includes('鄂尔多斯')) filteredBC[k]=v; });
  return { t, d, inv, nst, totalInvest: ti.toFixed(0), pct: (t?(d/t*100).toFixed(1):'0.0'), bc: filteredBC };
}

function parseSection局部绝缘化(ws) {
  const rd = XLSX.utils.decode_range(ws['!ref']);
  let t=0,d=0;
  const bc = {};
  // 数据从R4开始，列：序号(0) 旗县区供电分公司(1) ... 已实现绝缘化总长度(5)
  for(let r=rd.s.r+4; r<=rd.e.r; r++) {
    const br = gv(ws,r,1);
    if(!br || typeof br !== 'string') continue;
    t++;
    const len = gv(ws,r,5);
    const k = br.substring(0,4);
    if(!bc[k]) bc[k]={t:0,d:0};
    bc[k].t++;
    if(len && parseFloat(len)>0) { d++; bc[k].d++; }
  }
  const filteredBC = {};
  Object.entries(bc).forEach(([k,v]) => { if(k.includes('鄂尔多斯')) filteredBC[k]=v; });
  return { t, d, pct: (t?(d/t*100).toFixed(1):'0.0'), bc: filteredBC };
}

function parseSection树线矛盾(ws) {
  const rd = XLSX.utils.decode_range(ws['!ref']);
  let t=0,d=0;
  const bc = {};
  // 数据从R2开始，列：序号(0) 线路产权单位(1) ... 是否构成缺陷(9) 是否下达隐患告知书(12) 完成整改时间(13)
  for(let r=rd.s.r+2; r<=rd.e.r; r++) {
    const idx = gv(ws,r,0);
    if(typeof idx !== 'number') continue;
    t++;
    const city = gv(ws,r,1);
    const k = String(city).substring(0,4);
    if(!bc[k]) bc[k]={t:0,d:0};
    bc[k].t++;
    const completedTime = gv(ws,r,13);
    if(completedTime != null && completedTime !== '') { d++; bc[k].d++; }
  }
  const filteredBC = {};
  Object.entries(bc).forEach(([k,v]) => { if(k.includes('鄂尔多斯')) filteredBC[k]=v; });
  return { t, d, pct: (t?(d/t*100).toFixed(1):'0.0'), bc: filteredBC };
}

function parseSection示范区(ws) {
  const rd = XLSX.utils.decode_range(ws['!ref']);
  let t=0,ti=0,tc=0;
  const bc = {};
  // 数据从R2开始，列：序号(0) 盟市供电公司(1) 统计单位名称(2) 建设类型(3) 年度建设投资计划(4) 用户数(5) 线路条数(6) 线路长度(7) ASAI(9)
  for(let r=rd.s.r+2; r<=rd.e.r; r++) {
    const idx = gv(ws,r,0);
    if(typeof idx !== 'number') continue;
    t++;
    const invest = gv(ws,r,4);
    ti += (invest||0);
    const users = gv(ws,r,5);
    const lines = gv(ws,r,6);
    const lengthKm = gv(ws,r,7);
    const asai = gv(ws,r,9);
    const unitName = gv(ws,r,2);
    const k = String(unitName).substring(0,4);
    if(!bc[k]) bc[k]={t:0,users:0,lines:0,lengthKm:0,asai:0};
    bc[k].t++;
    bc[k].users += (users||0);
    bc[k].lines += (lines||0);
    bc[k].lengthKm += (lengthKm||0);
    if(asai!=null) bc[k].asai += asai;
  }
  // 过滤只保留鄂尔多斯
  const filteredBC = {};
  Object.entries(bc).forEach(([k,v]) => {
    if(k.includes('鄂尔多斯')) {
      v.asaiAvg = v.t>0 ? (v.asai/v.t).toFixed(6) : '0';
      filteredBC[k] = v;
    }
  });
  return { t, totalInvest: ti.toFixed(2), avgComp: '0.0', bc: filteredBC };
}

function parseSection异常台区(ws) {
  const rd = XLSX.utils.decode_range(ws['!ref']);
  let t=0,d=0,nst=0;
  const bc = {};
  // 数据从R4开始，列：序号(0) 旗县区供电分公司(1) ... 可研投资(5)
  // 没有明确的完成状态列，假设全部未开始
  for(let r=rd.s.r+4; r<=rd.e.r; r++) {
    const idx = gv(ws,r,0);
    if(typeof idx !== 'number') continue;
    t++;
    const city = gv(ws,r,1);
    const k = String(city).substring(0,4);
    if(!bc[k]) bc[k]={t:0,d:0,nst:0};
    bc[k].t++;
    nst++;
    bc[k].nst++;
  }
  const filteredBC = {};
  Object.entries(bc).forEach(([k,v]) => { if(k.includes('鄂尔多斯')) filteredBC[k]=v; });
  return { t, d, nst, pct: (t?(d/t*100).toFixed(1):'0.0'), bc: filteredBC };
}

function parseSection跨产权小区(ws) {
  const rd = XLSX.utils.decode_range(ws['!ref']);
  let t=0,d=0,dir=0,ref=0,bld=0;
  const bc = {};
  // 数据从R2开始，列：序号(0) 供电分公司(1) ... 改造接收模式(7) 改造移交协议签订时间(12)
  for(let r=rd.s.r+2; r<=rd.e.r; r++) {
    const idx = gv(ws,r,0);
    if(typeof idx !== 'number') continue;
    t++;
    const city = gv(ws,r,1);
    const k = String(city).substring(0,4);
    if(!bc[k]) bc[k]={t:0,d:0};
    bc[k].t++;
    const mode = gv(ws,r,7);
    if(mode==='直接移交') dir++;
    else if(mode==='先改造后移交') ref++;
    else if(mode==='以建代改') bld++;
    const agreementTime = gv(ws,r,12);
    if(agreementTime != null && agreementTime !== '') { d++; bc[k].d++; }
  }
  const filteredBC = {};
  Object.entries(bc).forEach(([k,v]) => { if(k.includes('鄂尔多斯')) filteredBC[k]=v; });
  return { t, d, dir, ref, bld, pct: (t?(d/t*100).toFixed(1):'0.0'), bc: filteredBC };
}

function parseSection农牧区线路(ws) {
  const rd = XLSX.utils.decode_range(ws['!ref']);
  let t=0,d=0,ti=0;
  const bc = {};
  // 数据从R1开始，列：序号(0) 旗县区供电分公司(1) ... 资金预算(4) 涉及线路条数(5) ... 是否完成治理(10)
  for(let r=rd.s.r+1; r<=rd.e.r; r++) {
    const unit = gv(ws,r,0);
    if(typeof unit !== 'number') continue;
    t++;
    ti += (gv(ws,r,4)||0);
    const city = gv(ws,r,1);
    const k = String(city).substring(0,4);
    if(!bc[k]) bc[k]={t:0,d:0};
    bc[k].t++;
    if(gv(ws,r,10)==='是') { d++; bc[k].d++; }
  }
  const filteredBC = {};
  Object.entries(bc).forEach(([k,v]) => { if(k.includes('鄂尔多斯')) filteredBC[k]=v; });
  return { t, d, totalInvest: ti.toFixed(2), pct: (t?(d/t*100).toFixed(1):'0.0'), bc: filteredBC };
}

// ========== 主处理流程 ==========
const D = { updateTime: '2026-07-09', sections: {} };

// 养老院
{ const ws=wb.Sheets['1养老院']; D.sections.养老院 = parseSection养老院(ws); }
// 高层小区
{ const ws=wb.Sheets['2高层小区']; D.sections.高层小区 = parseSection高层小区(ws); }
// 标准化考场
{ const ws=wb.Sheets['3标准化考场']; D.sections.标准化考场 = parseSection标准化考场(ws); }
// 频繁停电
{ const ws=wb.Sheets['4频繁停电']; D.sections.频繁停电 = parseSection频繁停电(ws); }
// 局部绝缘化
{ const ws=wb.Sheets['5-1局部绝缘化']; D.sections.局部绝缘化 = parseSection局部绝缘化(ws); }
// 树线矛盾
{ const ws=wb.Sheets['5-2树线矛盾']; D.sections.树线矛盾 = parseSection树线矛盾(ws); }
// 示范区
{ const ws=wb.Sheets['7.示范区']; D.sections.示范区 = parseSection示范区(ws); }
// 异常台区
{ const ws=wb.Sheets['6异常台区']; D.sections.异常台区 = parseSection异常台区(ws); }
// 跨产权小区
{ const ws=wb.Sheets['9.跨产权供电小区']; D.sections.跨产权小区 = parseSection跨产权小区(ws); }
// 农牧区线路
{ const ws=wb.Sheets['11.1700余条农牧区配电线路清单']; D.sections.农牧区线路 = parseSection农牧区线路(ws); }

// 汇总
let gt=0,gd=0,gi=0,gns=0,gv2=0;
for(const[n,s] of Object.entries(D.sections)) {
  gt += s.t||0; gd += s.d||0; gi += s.inv||0; gns += s.nst||0;
  gv2 += parseFloat(s.totalInvest)||0;
}
D.grand = { total:gt, done:gd, investing:gi, notstart:gns, invest:gv2.toFixed(0), pct:(gt?(gd/gt*100).toFixed(1):'0.0') };

fs.writeFileSync('dashboard_data.json', JSON.stringify(D));
console.log('Done. Grand:', gt, gd, gi, gns, D.grand.pct+'%');
console.log('Sections:', Object.keys(D.sections).map(n => n+':'+D.sections[n].t+'/'+D.sections[n].d).join(', '));
