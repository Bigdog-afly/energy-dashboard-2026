const XLSX = require('xlsx');
const wb = XLSX.readFile('鄂尔多斯供电公司2026年十项民生工程7.2.xlsx');
function gv(ws,r,c) { const a=XLSX.utils.encode_cell({r,c}); return ws[a]?ws[a].v:null; }

const dashboard = {
  updateTime: '2026-07-08',
  sections: {}
};

// ===== 1. 养老院 =====
{
  const ws = wb.Sheets['1养老院'];
  const rd = XLSX.utils.decode_range(ws['!ref']);
  let total=0,done=0,investing=0,notstart=0,totalInvest=0;
  let byCity={};
  for(let r=rd.s.r+2;r<=rd.e.r;r++){
    const idx=gv(ws,r,0); const city=gv(ws,r,1);
    if(typeof idx!=='number'||!city) continue;
    total++;
    const invest=gv(ws,r,5); totalInvest+=invest||0;
    const type=gv(ws,r,7); const isDone=gv(ws,r,8);
    const key=String(city).substring(0,4);
    if(!byCity[key]) byCity[key]={total:0,done:0,investing:0,notstart:0};
    byCity[key].total++;
    if(isDone==='是'){done++;byCity[key].done++;}
    else if(type&&String(type).trim()){investing++;byCity[key].investing++;}
    else{notstart++;byCity[key].notstart++;}
  }
  dashboard.sections.养老院={total,done,investing,notstart,totalInvest:totalInvest.toFixed(0),donePct:(done/total*100).toFixed(1),byCity};
}

// ===== 2. 高层小区 =====
{
  const ws = wb.Sheets['2高层小区'];
  const rd = XLSX.utils.decode_range(ws['!ref']);
  let total=0,done=0,investing=0,notstart=0;
  let byCity={};
  for(let r=rd.s.r+2;r<=rd.e.r;r++){
    const idx=gv(ws,r,0); const city=gv(ws,r,1);
    if(typeof idx!=='number'||!city) continue;
    total++;
    const type=gv(ws,r,8); const isDone=gv(ws,r,9);
    const key=String(city).substring(0,4);
    if(!byCity[key]) byCity[key]={total:0,done:0,investing:0,notstart:0};
    byCity[key].total++;
    if(isDone==='是'){done++;byCity[key].done++;}
    else if(type&&String(type).trim()){investing++;byCity[key].investing++;}
    else{notstart++;byCity[key].notstart++;}
  }
  dashboard.sections.高层小区={total,done,investing,notstart,donePct:(total?done/total*100:0).toFixed(1),byCity};
}

// ===== 3. 标准化考场 =====
{
  const ws = wb.Sheets['3标准化考场'];
  const rd = XLSX.utils.decode_range(ws['!ref']);
  let total=0,done=0,investing=0,notstart=0;
  let byCity={};
  for(let r=rd.s.r+2;r<=rd.e.r;r++){
    const idx=gv(ws,r,0); const city=gv(ws,r,1);
    if(typeof idx!=='number'||!city) continue;
    total++;
    const type=gv(ws,r,8); const isDone=gv(ws,r,9);
    const key=String(city).substring(0,4);
    if(!byCity[key]) byCity[key]={total:0,done:0,investing:0,notstart:0};
    byCity[key].total++;
    if(isDone==='是'){done++;byCity[key].done++;}
    else if(type&&String(type).trim()){investing++;byCity[key].investing++;}
    else{notstart++;byCity[key].notstart++;}
  }
  dashboard.sections.标准化考场={total,done,investing,notstart,donePct:(total?done/total*100:0).toFixed(1),byCity};
}

// ===== 4. 频繁停电 =====
{
  const ws = wb.Sheets['4频繁停电'];
  const rd = XLSX.utils.decode_range(ws['!ref']);
  let total=0,done=0,investing=0,notstart=0;
  let byCity={};
  for(let r=rd.s.r+2;r<=rd.e.r;r++){
    const city=gv(ws,r,0); const branch=gv(ws,r,2);
    if(!city||typeof city!=='string'||!branch) continue;
    total++;
    const isDone=gv(ws,r,17); const planDate=gv(ws,r,18);
    const key=city.substring(0,4);
    if(!byCity[key]) byCity[key]={total:0,done:0,investing:0,notstart:0};
    byCity[key].total++;
    if(isDone==='是'){done++;byCity[key].done++;}
    else if(planDate&&typeof planDate==='number'&&planDate>46000){investing++;byCity[key].investing++;}
    else{notstart++;byCity[key].notstart++;}
  }
  dashboard.sections.频繁停电={total,done,investing,notstart,donePct:(total?done/total*100:0).toFixed(1),byCity};
}

// ===== 5-1 局部绝缘化 =====
{
  const ws = wb.Sheets['5-1局部绝缘化'];
  const rd = XLSX.utils.decode_range(ws['!ref']);
  let total=0,done=0;
  let byCity={};
  for(let r=rd.s.r+2;r<=rd.e.r;r++){
    const branch=gv(ws,r,2); const done2=gv(ws,r,20);
    if(!branch||typeof branch!=='string') continue;
    total++;
    const key=branch.substring(0,4);
    if(!byCity[key]) byCity[key]={total:0,done:0};
    byCity[key].total++;
    if(done2&&String(done2).trim()){done++;byCity[key].done++;}
  }
  dashboard.sections.局部绝缘化={total,done,donePct:(total?done/total*100:0).toFixed(1),byCity};
}

// ===== 5-2 树线矛盾 =====
{
  const ws = wb.Sheets['5-2树线矛盾'];
  const rd = XLSX.utils.decode_range(ws['!ref']);
  let total=0,done=0;
  let byCity={};
  for(let r=rd.s.r+2;r<=rd.e.r;r++){
    const idx=gv(ws,r,0);
    if(typeof idx!=='number') continue;
    total++;
    const city=gv(ws,r,1);
    const key=String(city).substring(0,4);
    if(!byCity[key]) byCity[key]={total:0,done:0};
    byCity[key].total++;
    if(gv(ws,r,15)==='是'){done++;byCity[key].done++;}
  }
  dashboard.sections.树线矛盾={total,done,donePct:(total?done/total*100:0).toFixed(1),byCity};
}

// ===== 6 异常台区 =====
{
  const ws = wb.Sheets['6异常台区'];
  const rd = XLSX.utils.decode_range(ws['!ref']);
  let total=0,done=0,notstart=0;
  let byCity={};
  for(let r=rd.s.r+2;r<=rd.e.r;r++){
    const idx=gv(ws,r,0); const branch=gv(ws,r,2);
    if(typeof idx!=='number'||!branch) continue;
    total++;
    const key=branch.substring(0,4);
    if(!byCity[key]) byCity[key]={total:0,done:0,notstart:0};
    byCity[key].total++;
    if(gv(ws,r,22)==='是'){done++;byCity[key].done++;}
    else{notstart++;byCity[key].notstart++;}
  }
  dashboard.sections.异常台区={total,done,notstart,donePct:(total?done/total*100:0).toFixed(1),byCity};
}

// ===== 7 示范区 =====
{
  const ws = wb.Sheets['7.示范区'];
  const rd = XLSX.utils.decode_range(ws['!ref']);
  let total=0,totalInvest=0,totalCompletion=0;
  let byCity={};
  for(let r=rd.s.r+2;r<=rd.e.r;r++){
    const idx=gv(ws,r,0);
    if(typeof idx!=='number') continue;
    total++;
    const invest=gv(ws,r,4); const comp=gv(ws,r,30);
    totalInvest+=invest||0;
    if(comp!=null) totalCompletion+=comp;
    const key=String(gv(ws,r,1)).substring(0,4);
    if(!byCity[key]) byCity[key]={total:0,completion:0};
    byCity[key].total++;
    if(comp!=null) byCity[key].completion+=comp;
  }
  dashboard.sections.示范区={total,totalInvest:totalInvest.toFixed(0),avgCompletion:(total?((totalCompletion/total)*100).toFixed(1):0),byCity};
}

// ===== 9 跨产权小区 =====
{
  const ws = wb.Sheets['9.跨产权供电小区'];
  const rd = XLSX.utils.decode_range(ws['!ref']);
  let total=0,done=0,direct=0,reform=0,build=0;
  let byCity={};
  for(let r=rd.s.r+3;r<=rd.e.r;r++){
    const name=gv(ws,r,3);
    if(!name||String(name).startsWith('附')||String(name).startsWith('错')) continue;
    total++;
    const city=gv(ws,r,1); const key=String(city).substring(0,4);
    if(!byCity[key]) byCity[key]={total:0,done:0};
    byCity[key].total++;
    const mode=gv(ws,r,8);
    if(mode==='直接移交') direct++;
    else if(mode==='先改造后移交') reform++;
    else if(mode==='以建代改') build++;
    const w=gv(ws,r,17);
    if(w&&String(w).trim()){done++;byCity[key].done++;}
  }
  dashboard.sections.跨产权小区={total,done,direct,reform,build,donePct:(total?done/total*100:0).toFixed(1),byCity};
}

// ===== 10 保电场所 =====
{
  const ws = wb.Sheets['10.周期性保电场所再梳理'];
  const rd = XLSX.utils.decode_range(ws['!ref']);
  let total=0;
  let byCity={};
  for(let r=rd.s.r+4;r<=rd.e.r;r++){
    const idx=gv(ws,r,0);
    if(typeof idx!=='number') continue;
    total++;
    const city=gv(ws,r,1);
    if(city){
      const key=String(city).substring(0,4);
      byCity[key]=(byCity[key]||0)+1;
    }
  }
  dashboard.sections.保电场所={total,byCity};
}

// ===== 11 农牧区线路 =====
{
  const ws = wb.Sheets['11.1700余条农牧区配电线路清单'];
  const rd = XLSX.utils.decode_range(ws['!ref']);
  let total=0,done=0,totalInvest=0;
  let byCity={};
  for(let r=rd.s.r+1;r<=rd.e.r;r++){
    const unit=gv(ws,r,0);
    if(typeof unit!=='number') continue;
    total++;
    const invest=gv(ws,r,5); totalInvest+=invest||0;
    const key=String(gv(ws,r,1)).substring(0,4);
    if(!byCity[key]) byCity[key]={total:0,done:0};
    byCity[key].total++;
    if(gv(ws,r,11)==='是'){done++;byCity[key].done++;}
  }
  dashboard.sections.农牧区线路={total,done,totalInvest:totalInvest.toFixed(0),donePct:(total?done/total*100:0).toFixed(1),byCity};
}

// ===== Grand totals =====
let grandTotal=0,grandDone=0,grandInvesting=0,grandNotstart=0,grandInvest=0;
for(const[name,s] of Object.entries(dashboard.sections)){
  grandTotal+=s.total||0;
  grandDone+=(s.done||0);
  grandInvesting+=(s.investing||0);
  grandNotstart+=(s.notstart||0);
  grandInvest+=(parseFloat(s.totalInvest)||0);
}
dashboard.grandTotal=grandTotal;
dashboard.grandDone=grandDone;
dashboard.grandInvesting=grandInvesting;
dashboard.grandNotstart=grandNotstart;
dashboard.grandInvest=grandInvest.toFixed(0);
dashboard.grandPct=(grandTotal?(grandDone/grandTotal*100):0).toFixed(1);

console.log(JSON.stringify(dashboard, null, 2));
