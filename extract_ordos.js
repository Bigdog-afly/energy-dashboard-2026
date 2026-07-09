const XLSX = require('xlsx');
const wb = XLSX.readFile('鄂尔多斯供电公司2026年十项民生工程7.2.xlsx');
function gv(ws,r,c) { const a=XLSX.utils.encode_cell({r,c}); return ws[a]?ws[a].v:null; }
function hasOrdos(ws,r){return (gv(ws,r,1)&&String(gv(ws,r,1)).includes('鄂尔多斯'))||(gv(ws,r,0)&&String(gv(ws,r,0)).includes('鄂尔多斯'));}

const O={};

// 1. 养老院 (city=col1, branch=col2, invest=col5, type=col7, done=col8)
{
  const ws=wb.Sheets['1养老院'],rd=XLSX.utils.decode_range(ws['!ref']);
  let t=0,d=0,inv=0,nst=0,ti=0;
  for(let r=rd.s.r+2;r<=rd.e.r;r++){
    if(!hasOrdos(ws,r)) continue;
    t++; ti+=(gv(ws,r,5)||0);
    const tp=gv(ws,r,7),is=gv(ws,r,8);
    if(is==='是'){d++}else if(tp&&String(tp).trim()){inv++}else{nst++}
  }
  O.养老院={t,tot:d,inv,nst,totalInvest:ti.toFixed(0),pct:(t?(d/t*100).toFixed(1):'0.0')};
}

// 2. 高层小区 (city=col1, branch=col2, type=col8, done=col9)
{
  const ws=wb.Sheets['2高层小区'],rd=XLSX.utils.decode_range(ws['!ref']);
  let t=0,d=0,inv=0,nst=0;
  for(let r=rd.s.r+2;r<=rd.e.r;r++){
    if(!hasOrdos(ws,r)) continue;
    t++;
    const tp=gv(ws,r,8),is=gv(ws,r,9);
    if(is==='是'){d++}else if(tp&&String(tp).trim()){inv++}else{nst++}
  }
  O.高层小区={t,tot:d,inv,nst,pct:(t?(d/t*100).toFixed(1):'0.0')};
}

// 3. 标准化考场 (city=col1, branch=col2, type=col8, done=col9)
{
  const ws=wb.Sheets['3标准化考场'],rd=XLSX.utils.decode_range(ws['!ref']);
  let t=0,d=0,inv=0,nst=0;
  for(let r=rd.s.r+2;r<=rd.e.r;r++){
    if(!hasOrdos(ws,r)) continue;
    t++;
    const tp=gv(ws,r,8),is=gv(ws,r,9);
    if(is==='是'){d++}else if(tp&&String(tp).trim()){inv++}else{nst++}
  }
  O.标准化考场={t,tot:d,inv,nst,pct:(t?(d/t*100).toFixed(1):'0.0')};
}

// 4. 频繁停电 (city=col0, branch=col2, done=col17, planDate=col18)
{
  const ws=wb.Sheets['4频繁停电'],rd=XLSX.utils.decode_range(ws['!ref']);
  let t=0,d=0,inv=0,nst=0;
  for(let r=rd.s.r+2;r<=rd.e.r;r++){
    if(!hasOrdos(ws,r)) continue;
    t++;
    const is=gv(ws,r,17),pd=gv(ws,r,18);
    if(is==='是'){d++}else if(pd&&typeof pd==='number'&&pd>46000){inv++}else{nst++}
  }
  O.频繁停电={t,tot:d,inv,nst,pct:(t?(d/t*100).toFixed(1):'0.0')};
}

// 5-1 局部绝缘化 (branch=col2, done=col20)
{
  const ws=wb.Sheets['5-1局部绝缘化'],rd=XLSX.utils.decode_range(ws['!ref']);
  let t=0,d=0;
  for(let r=rd.s.r+2;r<=rd.e.r;r++){
    const br=gv(ws,r,2); if(!br||typeof br!=='string') continue;
    if(!br.includes('鄂尔多斯')&&!br.includes('东胜')&&!br.includes('铁西')&&!br.includes('伊金霍洛')&&!br.includes('鄂托克')&&!br.includes('杭锦')&&!br.includes('达拉特')&&!br.includes('康巴什')&&!br.includes('准格尔')&&!br.includes('乌审')) continue;
    t++;
    if(gv(ws,r,20)&&String(gv(ws,r,20)).trim()) d++;
  }
  O.局部绝缘化={t,tot:d,pct:(t?(d/t*100).toFixed(1):'0.0')};
}

// 5-2 树线矛盾 (city=col1, done=col15)
{
  const ws=wb.Sheets['5-2树线矛盾'],rd=XLSX.utils.decode_range(ws['!ref']);
  let t=0,d=0;
  for(let r=rd.s.r+2;r<=rd.e.r;r++){
    if(!hasOrdos(ws,r)) continue;
    t++;
    if(gv(ws,r,15)==='是') d++;
  }
  O.树线矛盾={t,tot:d,pct:(t?(d/t*100).toFixed(1):'0.0')};
}

// 6 异常台区 (idx=col0, branch=col2, done=col22)
{
  const ws=wb.Sheets['6异常台区'],rd=XLSX.utils.decode_range(ws['!ref']);
  let t=0,d=0,nst=0;
  for(let r=rd.s.r+2;r<=rd.e.r;r++){
    const idx=gv(ws,r,0),branch=gv(ws,r,2);
    if(typeof idx!=='number'||!branch) continue;
    if(!branch.includes('鄂尔多斯')&&!branch.includes('东胜')&&!branch.includes('铁西')&&!branch.includes('伊金霍洛')&&!branch.includes('鄂托克')&&!branch.includes('康巴什')&&!branch.includes('准格尔')) continue;
    t++;
    if(gv(ws,r,22)==='是') d++; else nst++;
  }
  O.异常台区={t,tot:d,nst,pct:(t?(d/t*100).toFixed(1):'0.0')};
}

// 7 示范区 (city=col1, invest=col4, comp=col30)
{
  const ws=wb.Sheets['7.示范区'],rd=XLSX.utils.decode_range(ws['!ref']);
  let t=0,ti=0,tc=0;
  for(let r=rd.s.r+2;r<=rd.e.r;r++){
    const city=gv(ws,r,1); if(!city||!String(city).includes('鄂尔多斯')) continue;
    t++;
    const inv=gv(ws,r,4),comp=gv(ws,r,30);
    ti+=(inv||0); if(comp!=null) tc+=comp;
  }
  O.示范区={t,totalInvest:ti.toFixed(0),avgComp:(t?(tc/t*100).toFixed(1):'0.0')};
}

// 9 跨产权小区 (city=col1, name=col3, mode=col8, completed=col17)
{
  const ws=wb.Sheets['9.跨产权供电小区'],rd=XLSX.utils.decode_range(ws['!ref']);
  let t=0,d=0,dir=0,ref=0,bld=0;
  for(let r=rd.s.r+3;r<=rd.e.r;r++){
    const city=gv(ws,r,1),name=gv(ws,r,3);
    if(!city||!name) continue;
    if(!city.includes('鄂尔多斯')) continue;
    if(String(name).startsWith('附')||String(name).startsWith('错')) continue;
    t++;
    const mode=gv(ws,r,8);
    if(mode==='直接移交') dir++;
    else if(mode==='先改造后移交') ref++;
    else if(mode==='以建代改') bld++;
    if(gv(ws,r,17)&&String(gv(ws,r,17)).trim()) d++;
  }
  O.跨产权小区={t,tot:d,dir,ref,bld,pct:(t?(d/t*100).toFixed(1):'0.0')};
}

// 10 农牧区线路 (city=col1, invest=col5, done=col11)
{
  const ws=wb.Sheets['10.1700余条农牧区配电线路清单'],rd=XLSX.utils.decode_range(ws['!ref']);
  let t=0,d=0,ti=0;
  for(let r=rd.s.r+1;r<=rd.e.r;r++){
    const unit=gv(ws,r,0); if(typeof unit!=='number') continue;
    const city=gv(ws,r,1); if(!city||!city.includes('鄂尔多斯')) continue;
    t++; ti+=(gv(ws,r,5)||0);
    if(gv(ws,r,11)==='是') d++;
  }
  O.农牧区线路={t,tot:d,totalInvest:ti.toFixed(0),pct:(t?(d/t*100).toFixed(1):'0.0')};
}

// Grand totals
let gt=0,gd=0,gi=0,gns=0,gv2=0;
for(const[n,s] of Object.entries(O)){
  gt+=s.t||0; gd+=(s.tot||0); gi+=(s.inv||0); gns+=(s.nst||0); gv2+=(parseFloat(s.totalInvest)||0);
}
O._grand={total:gt,done:gd,investing:gi,notstart:gns,invest:gv2.toFixed(0),pct:(gt?(gd/gt*100).toFixed(1):'0.0')};

require('fs').writeFileSync('ordos_data.json',JSON.stringify(O));
console.log(JSON.stringify(O,null,2));
