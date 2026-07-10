const XLSX = require('xlsx');
const wb = XLSX.readFile('鄂尔多斯供电公司2026年十项民生工程7.2.xlsx');
function gv(ws,r,c) { const a=XLSX.utils.encode_cell({r,c}); return ws[a]?ws[a].v:null; }

const D = { updateTime:'2026-07-08', sections:{} };

// 1. 养老院
{
  const ws=wb.Sheets['1养老院'], rd=XLSX.utils.decode_range(ws['!ref']);
  let t=0,d=0,inv=0,nst=0,ti=0;
  let bc={};
  for(let r=rd.s.r+2;r<=rd.e.r;r++){
    const idx=gv(ws,r,0),city=gv(ws,r,1);
    if(typeof idx!=='number'||!city) continue;
    t++; ti+=(gv(ws,r,5)||0);
    const tp=gv(ws,r,7),is=gv(ws,r,8);
    const k=String(city).substr(0,4);
    if(!bc[k])bc[k]={t:0,d:0,inv:0,nst:0};
    bc[k].t++;
    if(is==='是'){d++;bc[k].d++;}
    else if(tp&&String(tp).trim()){inv++;bc[k].inv++;}
    else{nst++;bc[k].nst++;}
  }
  D.sections.养老院={t,d,inv,nst,totalInvest:ti.toFixed(0),pct:(d/t*100).toFixed(1),bc};
}

// 2. 高层小区
{
  const ws=wb.Sheets['2高层小区'], rd=XLSX.utils.decode_range(ws['!ref']);
  let t=0,d=0,inv=0,nst=0;
  let bc={};
  for(let r=rd.s.r+2;r<=rd.e.r;r++){
    const idx=gv(ws,r,0),city=gv(ws,r,1);
    if(typeof idx!=='number'||!city) continue;
    t++;
    const tp=gv(ws,r,8),is=gv(ws,r,9);
    const k=String(city).substr(0,4);
    if(!bc[k])bc[k]={t:0,d:0,inv:0,nst:0};
    bc[k].t++;
    if(is==='是'){d++;bc[k].d++;}
    else if(tp&&String(tp).trim()){inv++;bc[k].inv++;}
    else{nst++;bc[k].nst++;}
  }
  D.sections.高层小区={t,d,inv,nst,pct:(t?d/t*100:0).toFixed(1),bc};
}

// 3. 标准化考场
{
  const ws=wb.Sheets['3标准化考场'], rd=XLSX.utils.decode_range(ws['!ref']);
  let t=0,d=0,inv=0,nst=0;
  let bc={};
  for(let r=rd.s.r+2;r<=rd.e.r;r++){
    const idx=gv(ws,r,0),city=gv(ws,r,1);
    if(typeof idx!=='number'||!city) continue;
    t++;
    const tp=gv(ws,r,8),is=gv(ws,r,9);
    const k=String(city).substr(0,4);
    if(!bc[k])bc[k]={t:0,d:0,inv:0,nst:0};
    bc[k].t++;
    if(is==='是'){d++;bc[k].d++;}
    else if(tp&&String(tp).trim()){inv++;bc[k].inv++;}
    else{nst++;bc[k].nst++;}
  }
  D.sections.标准化考场={t,d,inv,nst,pct:(t?d/t*100:0).toFixed(1),bc};
}

// 4. 频繁停电
{
  const ws=wb.Sheets['4频繁停电'], rd=XLSX.utils.decode_range(ws['!ref']);
  let t=0,d=0,inv=0,nst=0;
  let bc={};
  for(let r=rd.s.r+2;r<=rd.e.r;r++){
    const city=gv(ws,r,0),branch=gv(ws,r,2);
    if(!city||typeof city!=='string'||!branch) continue;
    t++;
    const is=gv(ws,r,17),pd=gv(ws,r,18);
    const k=city.substr(0,4);
    if(!bc[k])bc[k]={t:0,d:0,inv:0,nst:0};
    bc[k].t++;
    if(is==='是'){d++;bc[k].d++;}
    else if(pd&&typeof pd==='number'&&pd>46000){inv++;bc[k].inv++;}
    else{nst++;bc[k].nst++;}
  }
  D.sections.频繁停电={t,d,inv,nst,pct:(t?d/t*100:0).toFixed(1),bc};
}

// 5-1 局部绝缘化
{
  const ws=wb.Sheets['5-1局部绝缘化'], rd=XLSX.utils.decode_range(ws['!ref']);
  let t=0,d=0;
  let bc={};
  for(let r=rd.s.r+2;r<=rd.e.r;r++){
    const br=gv(ws,r,2),dn=gv(ws,r,20);
    if(!br||typeof br!=='string') continue;
    t++;
    const k=br.substr(0,4);
    if(!bc[k])bc[k]={t:0,d:0};
    bc[k].t++;
    if(dn&&String(dn).trim()){d++;bc[k].d++;}
  }
  D.sections.局部绝缘化={t,d,pct:(t?d/t*100:0).toFixed(1),bc};
}

// 5-2 树线矛盾
{
  const ws=wb.Sheets['5-2树线矛盾'], rd=XLSX.utils.decode_range(ws['!ref']);
  let t=0,d=0;
  let bc={};
  for(let r=rd.s.r+2;r<=rd.e.r;r++){
    const idx=gv(ws,r,0);
    if(typeof idx!=='number') continue;
    t++;
    const city=gv(ws,r,1);
    const k=String(city).substr(0,4);
    if(!bc[k])bc[k]={t:0,d:0};
    bc[k].t++;
    if(gv(ws,r,15)==='是'){d++;bc[k].d++;}
  }
  D.sections.树线矛盾={t,d,pct:(t?d/t*100:0).toFixed(1),bc};
}

// 6 异常台区
{
  const ws=wb.Sheets['6异常台区'], rd=XLSX.utils.decode_range(ws['!ref']);
  let t=0,d=0,nst=0;
  let bc={};
  for(let r=rd.s.r+2;r<=rd.e.r;r++){
    const idx=gv(ws,r,0),br=gv(ws,r,2);
    if(typeof idx!=='number'||!br) continue;
    t++;
    const k=br.substr(0,4);
    if(!bc[k])bc[k]={t:0,d:0,nst:0};
    bc[k].t++;
    if(gv(ws,r,22)==='是'){d++;bc[k].d++;}
    else{nst++;bc[k].nst++;}
  }
  D.sections.异常台区={t,d,nst,pct:(t?d/t*100:0).toFixed(1),bc};
}

// 7 示范区
{
  const ws=wb.Sheets['7.示范区'], rd=XLSX.utils.decode_range(ws['!ref']);
  let t=0,ti=0,tc=0;
  let bc={};
  for(let r=rd.s.r+2;r<=rd.e.r;r++){
    const idx=gv(ws,r,0);
    if(typeof idx!=='number') continue;
    t++;
    const inv=gv(ws,r,4),comp=gv(ws,r,30);
    ti+=(inv||0);
    if(comp!=null) tc+=comp;
    const k=String(gv(ws,r,1)).substr(0,4);
    if(!bc[k])bc[k]={t:0,c:0};
    bc[k].t++;
    if(comp!=null) bc[k].c+=comp;
  }
  D.sections.示范区={t,totalInvest:ti.toFixed(0),avgComp:(t?(tc/t*100).toFixed(1):0),bc};
}

// 9 跨产权小区
{
  const ws=wb.Sheets['9.跨产权供电小区'], rd=XLSX.utils.decode_range(ws['!ref']);
  let t=0,d=0,dir=0,ref=0,bld=0;
  let bc={};
  for(let r=rd.s.r+3;r<=rd.e.r;r++){
    const nm=gv(ws,r,3);
    if(!nm||String(nm).startsWith('附')||String(nm).startsWith('错')) continue;
    t++;
    const city=gv(ws,r,1),k=String(city).substr(0,4);
    if(!bc[k])bc[k]={t:0,d:0};
    bc[k].t++;
    const mode=gv(ws,r,8);
    if(mode==='直接移交') dir++;
    else if(mode==='先改造后移交') ref++;
    else if(mode==='以建代改') bld++;
    const w=gv(ws,r,17);
    if(w&&String(w).trim()){d++;bc[k].d++;}
  }
  D.sections.跨产权小区={t,d,dir,ref,bld,pct:(t?d/t*100:0).toFixed(1),bc};
}

// 10 农牧区线路（sheet名: 10.1700余条农牧区配电线路清单）
{
  const ws=wb.Sheets['10.1700余条农牧区配电线路清单'];
  if (!ws) {
    console.warn('Warning: No sheet found for 农牧区线路');
    D.sections.农牧区线路 = {t:0,d:0,totalInvest:'0',pct:'0.0',bc:{}};
  } else {
    const rd = XLSX.utils.decode_range(ws['!ref']);
    let t=0,d=0,ti=0;
    let bc={};
    for(let r=rd.s.r+1;r<=rd.e.r;r++){
      const unit=gv(ws,r,0);
      if(typeof unit!=='number') continue;
      t++;
      ti+=(gv(ws,r,5)||0);
      const k=String(gv(ws,r,1)).substr(0,4);
      if(!bc[k])bc[k]={t:0,d:0};
      bc[k].t++;
      if(gv(ws,r,11)==='是'){d++;bc[k].d++;}
    }
    D.sections.农牧区线路={t,d,totalInvest:ti.toFixed(0),pct:(t?d/t*100:0).toFixed(1),bc};
  }
}

// 汇总
let gt=0,gd=0,gi=0,gns=0,gv2=0;
for(const[n,s] of Object.entries(D.sections)){
  gt+=s.t||0; gd+=(s.d||0); gi+=(s.inv||0); gns+=(s.nst||0); gv2+=(parseFloat(s.totalInvest)||0);
}
D.grand={total:gt,done:gd,investing:gi,notstart:gns,invest:gv2.toFixed(0),pct:(gt?gd/gt*100:0).toFixed(1)};

require('fs').writeFileSync('C:/Users/生产技术部/Desktop/dashboard_data.json', JSON.stringify(D));
console.log('Done. Grand:', gt, gd, gi, gns, D.grand.pct+'%');
