(function(){
const $=s=>document.querySelector(s);
const SAMPLE=window.__SAMPLE__;
const norm=s=>String(s==null?'':s).normalize('NFD').replace(/[̀-ͯ]/g,'').replace(/đ/g,'d').replace(/Đ/g,'D').toLowerCase().replace(/\s+/g,' ').trim();
const nf=new Intl.NumberFormat('vi-VN'); const nf1=new Intl.NumberFormat('vi-VN',{maximumFractionDigits:1,minimumFractionDigits:1});
const fmtN=v=>v==null||isNaN(v)?'–':nf.format(Math.round(v));
const fmt1=v=>v==null||isNaN(v)?'–':nf1.format(v);
const fmtP=v=>v==null||isNaN(v)?'–':Math.round(v*100)+'%';
const fmtVND=v=>{if(v==null||isNaN(v))return '–'; if(Math.abs(v)>=1e9) return nf1.format(v/1e9)+' tỷ'; if(Math.abs(v)>=1e6) return nf1.format(v/1e6)+' tr'; return nf.format(Math.round(v));};
const iso=d=>d.getFullYear()+'-'+String(d.getMonth()+1).padStart(2,'0')+'-'+String(d.getDate()).padStart(2,'0');
const dmy=d=>d?String(d.getDate()).padStart(2,'0')+'/'+String(d.getMonth()+1).padStart(2,'0')+'/'+d.getFullYear():'–';
const MONTHS=['T1','T2','T3','T4','T5','T6','T7','T8','T9','T10','T11','T12'];

// ---------- parsing
function toDate(v){
  if(v==null||v==='') return null;
  if(v instanceof Date) return isNaN(v)?null:new Date(v.getFullYear(),v.getMonth(),v.getDate());
  if(typeof v==='number'){const d=new Date(Math.round((v-25569)*86400000)); return new Date(d.getUTCFullYear(),d.getUTCMonth(),d.getUTCDate());}
  const s=String(v).trim(); let m;
  if((m=s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})/))) return new Date(+m[1],+m[2]-1,+m[3]);
  if((m=s.match(/^(\d{1,2})[\/.-](\d{1,2})[\/.-](\d{4})/))) return new Date(+m[3],+m[2]-1,+m[1]);
  return null;
}
const num=v=>{if(v==null||v==='')return null; if(typeof v==='number')return v; const x=parseFloat(String(v).replace(/[^\d.,-]/g,'').replace(/\./g,'').replace(',','.')); return isNaN(x)?null:x;};
const str=v=>v==null?'':String(v).trim();
function table(aoa,keyHeader){
  // find header row within first 12 rows containing keyHeader
  let hi=-1; for(let i=0;i<Math.min(12,aoa.length);i++){ if((aoa[i]||[]).some(c=>norm(c)===norm(keyHeader))){hi=i;break;} }
  if(hi<0) return null;
  const hdr=aoa[hi].map(norm);
  const col=(...names)=>{for(const n of names){const k=norm(n); let i=hdr.findIndex(h=>h===k); if(i<0) i=hdr.findIndex(h=>h.startsWith(k)); if(i>=0) return i;} return -1;};
  return {rows:aoa.slice(hi+1),col};
}
function buildModel(sheets){
  // sheets: {nv:aoa, kh:aoa, lop:aoa, tr:aoa}
  const errs=[];
  const T={nv:table(sheets.nv||[], 'Mã NV'), kh:table(sheets.kh||[], 'Mã khóa')||table(sheets.kh||[], 'Mã lớp'), lop:table(sheets.lop||[], 'Mã lớp'), tr:table(sheets.tr||[], 'Mã lớp'), kp:sheets.kp?table(sheets.kp,'Nội dung đào tạo'):null};
  const miss=Object.entries({nv:'DS Nhan vien',kh:'DM Khoa hoc',lop:'Lop hoc',tr:'Training Record'}).filter(([k])=>!T[k]).map(([,v])=>v);
  if(miss.length) throw new Error('Không tìm thấy sheet hoặc dòng tiêu đề: '+miss.join(', ')+'. Hãy dùng đúng file mẫu SCC_Training_Record_Dashboard.xlsx.');
  const emp=new Map();
  {const t=T.nv,c={id:t.col('Mã NV'),name:t.col('Họ tên'),khoi:t.col('Khối'),pb:t.col('Phòng ban'),cd:t.col('Chức danh'),cb:t.col('Cấp bậc'),join:t.col('Ngày vào làm'),leave:t.col('Ngày nghỉ việc'),tt:t.col('Tình trạng')};
   for(const r of t.rows){const id=str(r[c.id]); if(!id) continue; emp.set(id,{id,name:str(r[c.name]),khoi:str(r[c.khoi]),pb:str(r[c.pb]),cd:str(r[c.cd]),cb:str(r[c.cb]),join:toDate(r[c.join]),leave:toDate(r[c.leave]),quit:c.tt>=0&&norm(r[c.tt])==='da nghi'});}}
  const crs=new Map();
  {const t=T.kh,c={id:t.col('Mã khóa','Mã lớp'),name:t.col('Tên khóa học'),nhom:t.col('Nhóm khóa học'),ht:t.col('Hình thức'),nb:t.col('Nội bộ'),h:t.col('Thời lượng chuẩn'),test:t.col('Có bài kiểm tra'),pass:t.col('Điểm đạt')};
   for(const r of t.rows){const id=str(r[c.id]); if(!id) continue; crs.set(id,{id,name:str(r[c.name]),nhom:str(r[c.nhom])||'(Chưa phân nhóm)',ht:str(r[c.ht]),nb:str(r[c.nb]),h:num(r[c.h])||0,test:norm(r[c.test])==='co',pass:num(r[c.pass])});}}
  const cls=new Map(); const byId=new Map(); const badCls=[];
  {const t=T.lop,c={id:t.col('Mã lớp'),kh:t.col('Mã khóa'),plan:t.col('Trong / Ngoài','Trong','Kế hoạch'),d0:t.col('Ngày bắt đầu'),d1:t.col('Ngày kết thúc'),h:t.col('Số giờ'),gv:t.col('Giảng viên'),khoi:t.col('Khối tổ chức'),dd:t.col('Địa điểm'),cp:t.col('Chi phí kế hoạch'),ca:t.col('Chi phí thực tế'),st:t.col('Trạng thái lớp'),note:t.col('Ghi chú')};
   if(c.kh<0) c.kh=c.id;
   for(const r of t.rows){const id=str(r[c.id]); if(!id) continue; const k=str(r[c.kh]); const co=crs.get(k);
     const d0=toDate(r[c.d0]);
     if(!co) badCls.push(id+' – mã "'+k+'" chưa có trong DM Khoa hoc'); if(!d0){badCls.push(id+' – thiếu Ngày bắt đầu'); continue;}
     const key=id+'|'+iso(d0);
     if(cls.has(key)){badCls.push(id+' ngày '+dmy(d0)+' – trùng đợt (cùng mã, cùng ngày), chỉ tính dòng đầu'); continue;}
     const o={id,key,kh:k,co,name:co?co.name:k,nhom:co?co.nhom:'(Chưa phân nhóm)',ht:co?co.ht:'',nb:co?co.nb:'',plan:norm(r[c.plan]).startsWith('trong'),d0,d1:toDate(r[c.d1]),h:num(r[c.h])??(co?co.h:0),gv:str(r[c.gv]),khoi:str(r[c.khoi])||'Toàn công ty',dd:str(r[c.dd]),cp:num(r[c.cp])||0,ca:num(r[c.ca]),st:str(r[c.st])||'Kế hoạch',note:str(r[c.note]),reg:0,done:0};
     cls.set(key,o); if(!byId.has(id)) byId.set(id,[]); byId.get(id).push(o);}}
  const recs=[]; const badRec=[]; const seen=new Set();
  {const t=T.tr,c={cl:t.col('Mã lớp'),dt:t.col('Ngày học'),nv:t.col('Mã NV'),st:t.col('Trạng thái tham dự'),hr:t.col('Giờ học thực tế'),pre:t.col('Điểm trước'),post:t.col('Điểm sau'),sat:t.col('Hài lòng'),app:t.col('Áp dụng sau'),cert:t.col('Đã cấp chứng nhận'),kqg:t.col('Kết quả ghi nhận','Kết quả')};
   t.rows.forEach((r,i)=>{const cid=str(r[c.cl]),eid=str(r[c.nv]); if(!cid&&!eid) return; const rowNo=i+1;
     const d=c.dt>=0?toDate(r[c.dt]):null; const dLabel=d?dmy(d):'';
     if(!cid||!eid){badRec.push({row:rowNo,msg:'Thiếu mã lớp hoặc mã NV',cid,d:dLabel,eid});return;}
     let cl=d?cls.get(cid+'|'+iso(d)):null;
     if(!cl&&!d){const l=byId.get(cid); if(l&&l.length===1) cl=l[0];}
     if(!cl){badRec.push({row:rowNo,msg:byId.has(cid)?(d?'Không có đợt học ngày này ở sheet Lop hoc':'Thiếu Ngày học (mã lớp có nhiều đợt)'):'Mã lớp không có trong sheet Lop hoc',cid,d:dLabel,eid});return;}
     const e=emp.get(eid);
     if(!e){badRec.push({row:rowNo,msg:'Mã NV không có trong DS Nhan vien',cid,d:dLabel,eid});}
     const key=cl.key+'|'+eid; if(seen.has(key)){badRec.push({row:rowNo,msg:'Trùng dòng (cùng NV, cùng lớp, cùng ngày) – không tính lần 2',cid,d:dLabel,eid});return;} seen.add(key);
     const st=str(r[c.st])||'Đã đăng ký'; const done=norm(st)==='hoan thanh';
     const hrs=done?(num(r[c.hr])??cl.h):0; const pre=num(r[c.pre]),post=num(r[c.post]);
     let res=''; const kg=c.kqg>=0?norm(r[c.kqg]):'';
     if(done&&(kg==='dat'||kg==='khong dat'||kg==='khong danh gia')){ res=kg==='dat'?'pass':kg==='khong dat'?'fail':'na'; }
     else if(done){ if(!cl.co||!cl.co.test) res='na'; else if(post==null) res='nodata'; else res=post>=(cl.co.pass??0)?'pass':'fail'; }
     cl.reg++; if(done) cl.done++;
     recs.push({cl,e,eid,khoi:e?e.khoi:'(NV không xác định)',st,stn:norm(st),done,hrs,pre,post,res,sat:num(r[c.sat]),app:norm(r[c.app]),year:cl.d0.getFullYear(),month:cl.d0.getMonth()});
   });}
  // Kế hoạch đào tạo năm (sheet 'Ke hoach dao tao'): mỗi dòng = 1 chương trình × 1 đợt dự kiến
  const plan=[];
  if(T.kp){const t=T.kp,c={y:t.col('Năm'),stt:t.col('STT'),sec:t.col('Nhóm KH','Nhóm kế hoạch'),code:t.col('Mã lớp'),name:t.col('Nội dung đào tạo'),dt:t.col('Đối tượng'),nb:t.col('Nội bộ'),dv:t.col('Đơn vị thực hiện','Đơn vị'),pb:t.col('Phòng ban đề xuất'),bud:t.col('Ngân sách dự kiến','Ngân sách','Chi phí dự kiến'),s:t.col('Bắt đầu dự kiến','Bắt đầu'),e:t.col('Kết thúc dự kiến','Kết thúc'),st:t.col('Trạng thái'),dn:t.col('Ngày thực hiện'),note:t.col('Ghi chú')};
   t.rows.forEach((r,i)=>{const name=str(r[c.name]); if(!name) return; const s0=toDate(r[c.s]), e0=toDate(r[c.e]), dn=toDate(r[c.dn]);
     const y=num(r[c.y])||((s0||e0||dn)?(s0||e0||dn).getFullYear():null); if(!y) return; const code=str(r[c.code]); const co=crs.get(code);
     plan.push({row:i+1,y,stt:str(r[c.stt]),sec:str(r[c.sec])||'(Chưa phân nhóm)',code,co,nhom:co?co.nhom:'',name,dt:str(r[c.dt]),nb:str(r[c.nb])||(co?co.nb:''),dv:str(r[c.dv]),pb:str(r[c.pb]),bud:num(r[c.bud])||0,s:s0||e0,e:e0||s0,st:norm(r[c.st]),dn,note:str(r[c.note])});});}
  return {emp,crs,cls,recs,badCls,badRec,plan};
}

// ---------- minimal XLSX reader (zip + XML, no external library)
async function inflate(u8){const ds=new DecompressionStream('deflate-raw'); const s=new Blob([u8]).stream().pipeThrough(ds); return new Uint8Array(await new Response(s).arrayBuffer());}
async function unzip(buf){
  const dv=new DataView(buf.buffer,buf.byteOffset,buf.byteLength); let e=-1;
  for(let i=buf.length-22;i>=Math.max(0,buf.length-70000);i--){ if(dv.getUint32(i,true)===0x06054b50){e=i;break;} }
  if(e<0) throw new Error('File không phải định dạng .xlsx hợp lệ (hãy lưu lại file dưới dạng Excel Workbook .xlsx).');
  const n=dv.getUint16(e+10,true); let p=dv.getUint32(e+16,true); const files={}; const td=new TextDecoder();
  for(let k=0;k<n;k++){ if(dv.getUint32(p,true)!==0x02014b50) break;
    const method=dv.getUint16(p+10,true), csize=dv.getUint32(p+20,true), nl=dv.getUint16(p+28,true), xl=dv.getUint16(p+30,true), cl=dv.getUint16(p+32,true), off=dv.getUint32(p+42,true);
    const name=td.decode(buf.subarray(p+46,p+46+nl)); files[name]={method,csize,off}; p+=46+nl+xl+cl; }
  return {names:Object.keys(files),async text(name){const f=files[name]; if(!f) return null; const ln=dv.getUint16(f.off+26,true), le=dv.getUint16(f.off+28,true); const st=f.off+30+ln+le; const raw=buf.subarray(st,st+f.csize); const data=f.method===8?await inflate(raw):raw; return td.decode(data);}};
}
const xmlDoc=s=>new DOMParser().parseFromString(s,'application/xml');
function colIdx(ref){let n=0; for(const ch of ref){const c=ch.charCodeAt(0); if(c<65||c>90) break; n=n*26+(c-64);} return n-1;}
async function readWorkbook(buf){
  const z=await unzip(buf);
  const wbx=xmlDoc(await z.text('xl/workbook.xml')||''); const rel=xmlDoc(await z.text('xl/_rels/workbook.xml.rels')||'');
  const rmap={}; for(const r of rel.getElementsByTagName('Relationship')){rmap[r.getAttribute('Id')]=r.getAttribute('Target');}
  const sheets=[...wbx.getElementsByTagName('sheet')].map(s=>{const rid=s.getAttribute('r:id')||s.getAttributeNS('http://schemas.openxmlformats.org/officeDocument/2006/relationships','id'); let t=rmap[rid]||''; t=t.startsWith('/')?t.slice(1):'xl/'+t; return {name:s.getAttribute('name'),path:t};});
  let ss=[]; const sst=await z.text('xl/sharedStrings.xml');
  if(sst){ss=[...xmlDoc(sst).getElementsByTagName('si')].map(si=>{let s=''; for(const t of si.getElementsByTagName('t')){ if(t.parentNode&&t.parentNode.nodeName==='rPh') continue; s+=t.textContent;} return s;});}
  async function aoa(path){const x=await z.text(path); if(!x) return null; const d=xmlDoc(x); const out=[];
    for(const row of d.getElementsByTagName('row')){const ri=(+row.getAttribute('r')||out.length+1)-1; const arr=[];
      for(const c of row.getElementsByTagName('c')){const ci=colIdx(c.getAttribute('r')||''); const t=c.getAttribute('t'); const v=c.getElementsByTagName('v')[0]; let val=null;
        if(t==='s') val=v?ss[+v.textContent]:null; else if(t==='inlineStr'){const is=c.getElementsByTagName('t')[0]; val=is?is.textContent:null;} else if(t==='str') val=v?v.textContent:null; else if(t==='b') val=v?v.textContent==='1':null; else if(t==='e') val=null; else if(v){const f=parseFloat(v.textContent); val=isNaN(f)?v.textContent:f;}
        if(ci>=0) arr[ci]=val;}
      out[ri]=arr;}
    for(let i=0;i<out.length;i++) if(!out[i]) out[i]=[]; return out;}
  const pick=async(...keys)=>{const s=sheets.find(s=>keys.some(k=>norm(s.name)===k))||sheets.find(s=>keys.some(k=>norm(s.name).includes(k))); return s?await aoa(s.path):null;};
  return {nv:await pick('ds nhan vien','nhan vien'),kh:await pick('dm khoa hoc','khoa hoc'),lop:await pick('lop hoc'),tr:await pick('training record'),kp:await pick('ke hoach dao tao')};
}

// ---------- state
let M=null; let source={name:'Dữ liệu mẫu',sample:true};
const S={tab:'perf',month:'hours',status:'Tất cả'};
try{const t=localStorage.getItem('scc-td-tab'); if(t) S.tab=t;}catch(e){}

function setModel(model,src){M=model; source=src; document.body.classList.add('loaded'); $('#landing').hidden=true; $('#app').hidden=false; $('#topbar').hidden=false; initFilters(); render();}
function showLanding(){document.body.classList.remove('loaded'); $('#landing').hidden=false; $('#app').hidden=true; $('#topbar').hidden=true;}
function initFilters(){
  const years=[...new Set([...M.cls.values()].map(c=>c.d0&&c.d0.getFullYear()).filter(Boolean))].sort();
  const today=new Date(); const yNow=today.getFullYear();
  const fy=$('#fYear'); const prevY=+fy.value; fy.replaceChildren(...years.map(y=>new Option(y,y)));
  fy.value=years.includes(prevY)?prevY:years.includes(yNow)?yNow:(years[years.length-1]||yNow);
  YEARS=years; const pc=$('#fCmp').value; fillCmp(); if([...$('#fCmp').options].some(o=>o.value===pc)) $('#fCmp').value=pc;
  const khois=[...new Set([...M.emp.values()].map(e=>e.khoi).filter(Boolean))].sort();
  $('#fKhoi').replaceChildren(new Option('Tất cả','*'),...khois.map(k=>new Option(k,k))); fillPb();
  const nhoms=[...new Set([...M.crs.values()].map(c=>c.nhom))];
  $('#fNhom').replaceChildren(new Option('Tất cả','*'),...nhoms.map(k=>new Option(k,k)));
  const d=$('#fDate'); if(!d.value) d.value=iso(today);
  const sts=['Tất cả','Kế hoạch','Đã xác nhận','Hoàn thành','Dời lịch','Hủy'];
  $('#fStatus').replaceChildren(...sts.map(s=>new Option(s,s)));
}

let YEARS=[];
function fillPb(){const K=$('#fKhoi').value, fp=$('#fPb'), v=fp.value;
  const pbs=[...new Set([...M.emp.values()].filter(e=>e.pb&&!e.quit&&!e.leave&&(K==='*'||e.khoi===K)).map(e=>e.pb))].sort((a,b)=>a.localeCompare(b,'vi'));
  fp.replaceChildren(new Option('Tất cả','*'),...pbs.map(p=>new Option(p,p))); fp.value=pbs.includes(v)?v:'*';}
function fillCmp(){const Y=+$('#fYear').value; const fc=$('#fCmp'); const v=fc.value;
  fc.replaceChildren(new Option('Không so sánh',''),...YEARS.filter(y=>y!==Y).sort((a,b)=>b-a).map(y=>new Option(y,y)));
  fc.value='';}
// headcount reference date for a year: the as-of date when it is in that year, else 31/12
const refDate=(Y,T)=>T.getFullYear()===Y?T:new Date(Y,11,31);
// NV 'Đã nghỉ' chưa có Ngày nghỉ việc: không biết nghỉ lúc nào -> không tính vào headcount
const activeAt=(e,d)=>(!e.join||e.join<=d)&&(e.leave?e.leave>d:!e.quit);
function filt(Yo,cut){
  const Y=Yo??+$('#fYear').value, K=$('#fKhoi').value, PB=$('#fPb').value||'*', N=$('#fNhom').value; const T=toDate($('#fDate').value)||new Date();
  const F0=Yo==null?toDate($('#fFrom').value):null; const D0=F0&&F0.getFullYear()===Y?F0:null; const D1=D0?(T.getFullYear()===Y?T:new Date(Y,11,31)):null;
  const inCut=d=>(!cut||d<=cut)&&(!D0||(d>=D0&&d<=D1));
  const recs=M.recs.filter(r=>r.year===Y&&inCut(r.cl.d0)&&(K==='*'||r.khoi===K)&&(PB==='*'||(r.e&&r.e.pb===PB))&&(N==='*'||r.cl.nhom===N));
  const clsAll=[...M.cls.values()];
  const cls=clsAll.filter(c=>c.d0&&c.d0.getFullYear()===Y&&inCut(c.d0)&&(K==='*'||c.khoi===K)&&(N==='*'||c.nhom===N));
  const R=cut||refDate(Y,T); const emps=[...M.emp.values()].filter(e=>activeAt(e,R)&&(K==='*'||e.khoi===K)&&(PB==='*'||e.pb===PB));
  const plan=(M.plan||[]).filter(p=>p.y===Y&&(N==='*'||p.nhom===N)&&(!D0||(p.s&&p.e&&p.s<=D1&&p.e>=D0)||(p.dn&&p.dn>=D0&&p.dn<=D1)));
  return {Y,K,PB,N,T,R,recs,cls,clsAll,emps,plan,cut,D0,D1};
}
// compare year: same period (cùng kỳ) when the viewed year is still running, else full year
function cmpF(){return null; const v=$('#fCmp').value; if(!v) return null; const Y=+$('#fYear').value, Yc=+v; const T=toDate($('#fDate').value)||new Date();
  const cut=T.getFullYear()===Y?new Date(Yc,T.getMonth(),Math.min(T.getDate(),new Date(Yc,T.getMonth()+1,0).getDate())):null;
  const F=filt(Yc,cut); F.full=filt(Yc); F.cut=cut; return F;}
const avg=a=>a.length?a.reduce((x,y)=>x+y,0)/a.length:null;
function perfM(F){
  const done=F.recs.filter(r=>r.done); const hc=F.emps.length;
  const hcSet=new Set(F.emps.map(e=>e.id)); const lset=new Set(done.map(r=>r.eid)); const learners=lset.size; let uniq=0; lset.forEach(id=>{if(hcSet.has(id)) uniq++;}); const hrs=done.reduce((a,r)=>a+r.hrs,0); const hasH=done.some(r=>r.hrs>0);
  const attD=F.recs.filter(r=>r.stn!=='da dang ky').length;
  const pass=done.filter(r=>r.res==='pass').length, fail=done.filter(r=>r.res==='fail').length;
  const sats=done.filter(r=>r.sat>0).map(r=>r.sat);
  const appEval=done.filter(r=>['co','mot phan','khong'].includes(r.app)); const appYes=appEval.filter(r=>r.app==='co').length;
  return {done,hc,uniq,learners,hasH,hrs,hpn:hc&&hasH?hrs/hc:null,cov:hc?uniq/hc:null,attN:done.length,attD,att:attD?done.length/attD:null,pass,tested:pass+fail,pr:(pass+fail)?pass/(pass+fail):null,sat:avg(sats),satN:sats.length,appYes,appN:appEval.length,ap:appEval.length?appYes/appEval.length:null};
}
function budM(F){const hasP=F.plan.length>0; const plan=hasP?F.plan.reduce((a,p)=>a+p.bud,0):F.cls.reduce((a,c)=>a+c.cp,0), act=F.cls.reduce((a,c)=>a+(c.ca||0),0);
  const doneRec=M.recs.filter(r=>F.cls.includes(r.cl)&&r.done); const hrs=doneRec.reduce((a,r)=>a+r.hrs,0);
  const intA=F.cls.filter(c=>norm(c.nb)==='noi bo').reduce((a,c)=>a+(c.ca||0),0);
  return {plan,act,rem:plan-act,u:plan?act/plan:null,perRec:doneRec.length?act/doneRec.length:null,perH:hrs?act/hrs:null,intS:act?intA/act:null,nRec:doneRec.length,nDone:F.cls.filter(c=>c.st==='Hoàn thành').length,hasP};}
// delta badge: kind 'pct' (relative change), 'pp' (percentage points), 'abs' (absolute, 1 decimal); good:+1 higher is better, -1 lower is better, 0 neutral
function dlt(cur,prev,Yc,kind='pct',good=1){
  if(cur==null||prev==null||isNaN(cur)||isNaN(prev)) return {t:'–',e:' '+Yc+': chưa có số',c:''};
  let d,txt; if(kind==='pp'){d=cur-prev; txt=(d>=0?'+':'')+Math.round(d*100)+' điểm %';}
  else if(kind==='abs'){d=cur-prev; txt=(d>=0?'+':'')+nf1.format(d);}
  else {if(!prev) return {t:'–',e:' '+Yc+' = 0',c:''}; d=(cur-prev)/Math.abs(prev); txt=(d>=0?'+':'')+Math.round(d*100)+'%';}
  const ar=Math.abs(d)<1e-9?'■':d>0?'▲':'▼'; const c=good===0||Math.abs(d)<1e-9?'':((d>0)===(good>0)?'up':'down');
  return {t:ar+' '+txt,e:'vs '+Yc,c};
}

const ICO={
 clock:'<circle cx="12" cy="12" r="8"/><path d="M12 8v4l2.5 2"/>',
 users:'<circle cx="9" cy="9" r="3"/><path d="M3.5 18c.8-3 3-4.5 5.5-4.5s4.7 1.5 5.5 4.5"/><path d="M15.5 6.5a3 3 0 010 5M17.5 13.8c1.5.6 2.6 2 3 4.2"/>',
 gauge:'<path d="M4.5 16a7.5 7.5 0 1115 0"/><path d="M12 16l3.5-4"/>',
 shield:'<path d="M12 3.5l7 3v5c0 4.4-3 7.6-7 9-4-1.4-7-4.6-7-9v-5l7-3z"/><path d="M9 12l2 2 4-4"/>',
 check:'<rect x="4" y="4" width="16" height="16" rx="4"/><path d="M8.5 12.2l2.4 2.4 4.6-5"/>',
 star:'<path d="M12 4.5l2.3 4.7 5.2.8-3.8 3.6.9 5.1L12 16.3l-4.6 2.4.9-5.1-3.8-3.6 5.2-.8z"/>',
 trend:'<path d="M4 16l5-5 3.5 3.5L20 7"/><path d="M15 7h5v5"/>',
 spark:'<path d="M12 3.5v3M12 17.5v3M3.5 12h3M17.5 12h3M6 6l2 2M16 16l2 2M6 18l2-2M16 8l2-2"/>',
 user:'<circle cx="12" cy="8.5" r="3.5"/><path d="M5 19.5c1-3.6 3.8-5.5 7-5.5s6 1.9 7 5.5"/>',
 coin:'<ellipse cx="12" cy="7" rx="7" ry="3"/><path d="M5 7v5c0 1.7 3.1 3 7 3s7-1.3 7-3V7"/><path d="M5 12v5c0 1.7 3.1 3 7 3s7-1.3 7-3v-5"/>',
 pie:'<path d="M12 4a8 8 0 108 8h-8z"/><path d="M14.5 2.5a8 8 0 017 7h-7z"/>',
 cal:'<rect x="4" y="5" width="16" height="15" rx="3"/><path d="M8 3v4M16 3v4M4 10h16"/>',
 book:'<path d="M5 5.5A2.5 2.5 0 017.5 3H19v15H7.5A2.5 2.5 0 005 20.5z"/><path d="M5 20.5V5.5"/>',
 warn:'<path d="M12 4l9 15H3z"/><path d="M12 10v4M12 16.8v.2"/>',
 x:'<circle cx="12" cy="12" r="8"/><path d="M9.5 9.5l5 5M14.5 9.5l-5 5"/>',
 home:'<path d="M4 10.5L12 4l8 6.5V19a1 1 0 01-1 1h-4.5v-5.5h-5V20H5a1 1 0 01-1-1z"/>',
 ext:'<circle cx="12" cy="12" r="8"/><path d="M4 12h16M12 4c2.2 2.3 3.3 5 3.3 8s-1.1 5.7-3.3 8c-2.2-2.3-3.3-5-3.3-8s1.1-5.7 3.3-8z"/>'};
function icon(n){const s=document.createElementNS('http://www.w3.org/2000/svg','svg'); s.setAttribute('width','17'); s.setAttribute('height','17'); s.setAttribute('viewBox','0 0 24 24'); s.setAttribute('fill','none'); s.setAttribute('stroke','currentColor'); s.setAttribute('stroke-width','1.9'); s.setAttribute('stroke-linecap','round'); s.setAttribute('stroke-linejoin','round'); s.innerHTML=ICO[n]||ICO.spark; return s;}
// ---------- render helpers
const el=(tag,attrs={},...kids)=>{const n=document.createElement(tag); for(const[k,v] of Object.entries(attrs)){ if(k==='class') n.className=v; else if(k==='text') n.textContent=v; else n.setAttribute(k,v);} for(const k of kids) if(k!=null) n.append(k); return n;};
function kpis(host,items){host.replaceChildren(...items.map(it=>{const k=el('div',{class:'kpi'+(it.wide?' kw':'')});
  const hd=el('div',{class:'hd'}); const ic=el('span',{class:'ico'}); ic.append(icon(it.i)); hd.append(ic,el('div',{class:'l',text:it.l})); k.append(hd);
  if(it.multi){const w=el('div',{class:'kmw'}); it.multi.forEach(x=>{const r=el('div',{class:'km'}); r.append(el('div',{class:'km-l',text:x.l}));
      const v=el('div',{class:'v'}); v.append(document.createTextNode(x.p==null||isNaN(x.p)?'–':fmtP(x.p))); r.append(v);
      if(!it.dual){const bar=el('div',{class:'km-bar'}); const f=el('i'); f.style.width=(x.p==null||isNaN(x.p)?0:Math.max(0,Math.min(100,x.p*100)))+'%'; bar.append(f); r.append(bar);}
      r.append(el('div',{class:'n',text:x.n}));
      if(x.d){const dw=el('div',{class:'dlw'}); dw.append(el('span',{class:'dl '+(x.d.c||''),text:x.d.t})); if(x.d.e) dw.append(el('span',{class:'dle',text:x.d.e.trim()})); r.append(dw);}
      w.append(r);}); k.append(w);
    if(it.dual){k.append(stackBar(it.dual.att,it.dual.pr,'kbar')); const lg=el('div',{class:'slg'}); lg.append(el('span',{class:'k1',text:'Đạt'}),el('span',{class:'k2',text:'Tham dự'})); k.append(lg);}
    return k;}
  const v=el('div',{class:'v'}); const m=String(it.v).match(/^(.*?)(\s*\/\s*5|\s+(tr|tỷ|giờ))?$/); v.append(document.createTextNode(m?m[1]:it.v)); if(it.of!=null) v.append(el('span',{class:'of',text:'/'+it.of})); const U={tr:'triệu đ','tỷ':'tỷ đ','giờ':'giờ'}; const u=it.u||(m&&m[3]?U[m[3]]:(it.vnd&&it.v!=='–'?'đ':null)); if(u) v.append(el('small',{text:u})); k.append(v);
  if(it.st){const s=el('span',{class:'st'}); const d=el('span',{class:'dot'}); d.style.background=`var(--${it.st[0]})`; s.append(d,document.createTextNode(it.st[1])); k.append(s);}
  if(it.n&&!it.bl) k.append(el('div',{class:'n',text:it.n}));
  if(it.bl){const u=el('ul',{class:'kbl'}); it.bl.forEach(([a,v])=>{const li=el('li'); li.append(el('span',{text:a}),el('b',{text:v})); u.append(li);}); k.append(u); if(it.n) k.append(el('div',{class:'n kbn',text:it.n}));}
  if(it.d){const w=el('div',{class:'dlw'}); w.append(el('span',{class:'dl '+(it.d.c||''),text:it.d.t})); if(it.d.e) w.append(el('span',{class:'dle',text:it.d.e.trim()})); k.append(w);}
  if(it.p!=null&&!isNaN(it.p)){const mt=el('div',{class:'meter'}); const b=el('i'); b.style.width=Math.max(0,Math.min(100,it.p*100))+'%'; mt.append(b); k.append(mt);}
  return k;}));}
const status=(v,good,warn)=>v==null||isNaN(v)?null:(v>=good?['good','Đạt mục tiêu ≥ '+fmtP(good)]:v>=warn?['warn','Cần theo dõi']:['crit','Dưới '+fmtP(warn)]);

const tip=$('#tip');
function showTip(ev,lines){tip.replaceChildren(); lines.forEach((l,i)=>{const d=el('div'); if(i===0){d.append(el('b',{text:l[0]}));} else {if(l[2]){const s=el('span',{class:'sw'}); s.style.background=l[2]; s.style.width='10px'; s.style.height='2px'; s.style.marginRight='6px'; s.style.verticalAlign='middle'; d.append(s);} d.append(el('b',{text:l[1]}),document.createTextNode(' '+l[0]));} tip.append(d);}); tip.hidden=false; moveTip(ev);}
function moveTip(ev){const w=tip.offsetWidth,h=tip.offsetHeight; let x=ev.clientX+14,y=ev.clientY+14; if(x+w>innerWidth-8)x=ev.clientX-w-14; if(y+h>innerHeight-8)y=ev.clientY-h-14; tip.style.left=x+'px'; tip.style.top=y+'px';}
const hideTip=()=>{tip.hidden=true};
const NS='http://www.w3.org/2000/svg';
const sv=(tag,a={})=>{const n=document.createElementNS(NS,tag); for(const[k,v] of Object.entries(a)) n.setAttribute(k,v); return n;};
function niceMax(v){if(v<=0) return 1; const p=Math.pow(10,Math.floor(Math.log10(v))); const f=v/p; const st=[1,1.2,1.6,2,2.4,3,4,5,6,8,10]; return st.find(x=>f<=x+1e-9)*p;}
// vertical bars; series: [{name,color,values}] grouped
function vbars(host,cats,series,fmt,{h=240,tickFmt=fmt,highlight=-1,intTicks=false,label=fmt,max:fixMax=null,ticks=4,small=false}={}){
  host.replaceChildren(); const W=Math.max(320,host.clientWidth||600); const pad={l:48,r:8,t:22,b:26};
  let max=niceMax(Math.max(0,...series.flatMap(s=>s.values))); if(intTicks) max=Math.max(ticks,Math.ceil(Math.max(0,...series.flatMap(s=>s.values))/ticks)*ticks); if(fixMax) max=fixMax;
  const svg=sv('svg',{viewBox:`0 0 ${W} ${h}`,width:'100%',height:h,role:'img'});
  const iw=W-pad.l-pad.r, ih=h-pad.t-pad.b; const y=v=>pad.t+ih-(v/max)*ih;
  for(let i=0;i<=ticks;i++){const v=max*i/ticks; svg.append(sv('line',{x1:pad.l,x2:W-pad.r,y1:y(v),y2:y(v),stroke:'var(--grid)','stroke-width':1}));
    const t=sv('text',{x:pad.l-6,y:y(v)+4,'text-anchor':'end','font-size':11,fill:'var(--muted)'}); t.textContent=tickFmt(v); svg.append(t);}
  const bw=iw/cats.length; const n=series.length; const gap=4; const barW=Math.min(30,(bw*(n>1?0.84:0.7)-(n-1)*gap)/n);
  cats.forEach((c,i)=>{const cx=pad.l+bw*i+bw/2; const gx=cx-(n*barW+(n-1)*gap)/2;
    const t=sv('text',{x:cx,y:h-8,'text-anchor':'middle','font-size':small?10:11,fill:'var(--muted)'}); t.textContent=c; svg.append(t);
    series.forEach((s,j)=>{const v=s.values[i]||0; const x=gx+j*(barW+gap); const yy=y(v); const hh=Math.max(0,pad.t+ih-yy);
      if(hh>0){const r=Math.min(4,barW/2,hh); svg.append(sv('path',{d:`M${x},${pad.t+ih}V${yy+r}a${r},${r} 0 0 1 ${r},${-r}H${x+barW-r}a${r},${r} 0 0 1 ${r},${r}V${pad.t+ih}Z`,fill:s.colors?s.colors[i]:s.color}));
        if(label){const lt=sv('text',{x:x+barW/2,y:yy-5,'text-anchor':'middle','font-size':small?8.5:n>1?9.5:11,'font-weight':600,fill:'var(--ink-2)'}); lt.textContent=label(v); svg.append(lt);}}});
    const hit=sv('rect',{x:pad.l+bw*i,y:pad.t,width:bw,height:ih,fill:'transparent',tabindex:0});
    const lines=[[c]].concat(series.map(s=>[s.name,(s.tips?s.tips[i]:fmt(s.values[i]||0)),s.colors?s.colors[i]:s.color]));
    hit.addEventListener('pointermove',e=>showTip(e,lines)); hit.addEventListener('pointerleave',hideTip);
    hit.addEventListener('focus',e=>{const b=hit.getBoundingClientRect(); showTip({clientX:b.left+b.width/2,clientY:b.top+20},lines);}); hit.addEventListener('blur',hideTip);
    svg.append(hit);});
  svg.append(sv('line',{x1:pad.l,x2:W-pad.r,y1:pad.t+ih,y2:pad.t+ih,stroke:'var(--axis)','stroke-width':1}));
  host.append(svg);
}
// horizontal bars with labels; items [{label,value,sub}] value fraction or number
function hbars(host,items,fmt,{max=null,color='var(--s1)',heads=null}={}){
  host.replaceChildren(); if(!items.length){host.append(el('div',{class:'empty',text:'Chưa có dữ liệu'}));return;}
  const mx=max??niceMax(Math.max(...items.map(i=>i.value)));
  const box=el('div',{class:'hbx'}); const bh=items.length<=6?24:16; const wide=items.some(i=>i.cnt);
  const cols=wide?'minmax(90px,26%) 1fr 104px 46px 128px':'minmax(90px,30%) 1fr auto';
  const cell=(t,css)=>{const d=el('div',{text:t}); d.style.cssText='font-size:12.5px;font-variant-numeric:tabular-nums;text-align:right;white-space:nowrap;'+(css||''); return d;};
  if(wide&&heads){const h=el('div',{class:'hbh'}); h.style.cssText=`display:grid;grid-template-columns:${cols};gap:10px;align-items:end;padding-bottom:6px;border-bottom:1px solid var(--grid)`;
    heads.forEach((t,i)=>h.append(cell(t,`font-size:11px;font-weight:700;letter-spacing:.03em;color:var(--muted);text-transform:uppercase;white-space:normal;line-height:1.3;text-align:${i<2?'left':'right'}`))); box.append(h);}
  items.forEach(it=>{const row=el('div'); row.style.cssText=`display:grid;grid-template-columns:${cols};gap:10px;align-items:center`;
    const lab=el('div',{text:it.label}); lab.style.cssText='font-size:12.5px;color:var(--ink-2);overflow:hidden;text-overflow:ellipsis;white-space:nowrap'; lab.title=it.label;
    const track=el('div'); track.style.cssText=`height:${bh}px;position:relative;background:var(--surface-2);border-radius:6px`;
    const bar=el('div'); bar.style.cssText=`height:${bh}px;border-radius:6px;background:${color};width:${Math.max(0,Math.min(100,it.value/mx*100))}%`;
    track.append(bar); row.append(lab,track);
    if(wide){row.append(cell(it.cnt||'','color:var(--ink)'),cell(fmt(it.value),'font-weight:700;color:var(--ink)'),cell(it.txt||'','color:var(--ink-2)'));}
    else{const v=el('div',{text:it.txt?'':fmt(it.value)}); if(it.txt){const b=el('b',{text:fmt(it.value)}); b.style.cssText='display:inline-block;min-width:38px;font-weight:700'; const x=el('span',{text:it.txt}); x.style.cssText='color:var(--ink-2);display:inline-block;min-width:92px;margin-left:8px'; v.append(b,x);} v.style.cssText='font-size:12.5px;font-variant-numeric:tabular-nums;color:var(--ink);min-width:44px;text-align:right;white-space:nowrap'; row.append(v);}
    const tl=[[it.label],['',fmt(it.value),color]].concat(it.sub?[[it.sub,'',null]]:[]);
    row.addEventListener('pointermove',e=>showTip(e,tl)); row.addEventListener('pointerleave',hideTip);
    box.append(row);});
  host.append(box);
}
// bảng phòng ban gọn: STT · Khối · Phòng ban · thanh (đã học/tổng · %) · Số khóa · Giờ học
const PBS={i:2,d:-1};
const ROM=['I','II','III','IV','V','VI','VII','VIII','IX','X'];
const KORD=['Nhà máy','Nội nghiệp','Thương mại'];
const kNm=k=>String(k||'').replace(/^Khối\s+/,'');
function pbTable(host,items,hasH){host.replaceChildren(); if(!items.length){host.append(el('div',{class:'empty',text:'Chưa có dữ liệu'})); return;}
  const t=el('div',{class:'pbt'}); const hd=el('div',{class:'pbr pbh'}); const body=el('div');
  const K=[null,it=>it.label,it=>it.value,it=>it.nk,it=>hasH?it.hh:it.n];
  const heads=['STT','Phòng ban','NV đang làm đã học / tổng · tỷ lệ','Số khóa học',hasH?'Giờ học (gồm NV đã nghỉ)':'Lượt học (gồm NV đã nghỉ)'];
  const hcs=heads.map((x,i)=>{const c=el('div',{class:(i===0||i>=3?'r':'')+(i?' srt':''),text:x}); if(i){c.tabIndex=0; c.title='Bấm để sắp xếp trong từng khối'; const go=()=>{if(PBS.i===i) PBS.d=-PBS.d; else {PBS.i=i; PBS.d=i>=2?-1:1;} draw();}; c.addEventListener('click',go); c.addEventListener('keydown',ev=>{if(ev.key==='Enter'||ev.key===' '){ev.preventDefault(); go();}});} hd.append(c); return c;});
  const cmp=(x,y)=>typeof x==='number'?x-y:String(x).localeCompare(String(y),'vi',{numeric:true});
  const ko=k=>{const i=KORD.indexOf(kNm(k)); return i<0?99:i;};
  const groups=[...new Set(items.map(it=>it.khoi||''))].sort((x,y)=>ko(x)-ko(y)||kNm(x).localeCompare(kNm(y),'vi'));
  const bar=(v,u,h2)=>{const tr=el('div',{class:'pbbar'}); const f=el('i'); f.style.width=Math.max(0,Math.min(100,v*100))+'%'; tr.append(f);
    const lb=el('span',{class:v<0.35?'out':''}); lb.append(el('b',{text:fmtP(v)}),document.createTextNode(` · ${fmtN(u)}/${fmtN(h2)} NV`)); tr.append(lb); return tr;};
  const sumRow=(label,arr,cls)=>{const u=arr.reduce((x,it)=>x+it.u,0), h2=arr.reduce((x,it)=>x+it.h,0), hh=arr.reduce((x,it)=>x+it.hh,0), n=arr.reduce((x,it)=>x+it.n,0);
    const ks=new Set(arr.flatMap(it=>[...(it.ks||[])])); const r=el('div',{class:'pbr '+cls});
    const lab=el('div',{class:'nm',text:label}); lab.style.gridColumn='1 / span 2'; r.append(lab,bar(h2?u/h2:0,u,h2),el('div',{class:'r',text:fmtN(ks.size)}),el('div',{class:'r',text:hasH?fmtN(hh)+' giờ':fmtN(n)+' lượt'}));
    r.title=`${label}: ${fmtN(u)}/${fmtN(h2)} NV đang làm đã học · ${fmtN(arr.length)} phòng ban`; return r;};
  function draw(){hcs.forEach((c,i)=>c.dataset.sort=i&&i===PBS.i?(PBS.d>0?'asc':'desc'):'');
    const out=[]; let n=0;
    groups.forEach((g,gi)=>{const arr=items.filter(it=>(it.khoi||'')===g).sort((a,b2)=>cmp(K[PBS.i](a),K[PBS.i](b2))*PBS.d||(b2.value-a.value)||(b2.hh-a.hh)||a.label.localeCompare(b2.label,'vi'));
      out.push(sumRow(`${ROM[gi]||gi+1}. ${g?(/^Khối/i.test(g)?g:'Khối '+g):'Chưa xác định khối'}`,arr,'pbsec'));
      arr.forEach(it=>{const r=el('div',{class:'pbr'}); r.append(el('div',{class:'r mu',text:String(++n)}));
        const nm=el('div',{class:'nm',text:it.label}); nm.title=it.label; r.append(nm,bar(it.value,it.u,it.h));
        r.append(el('div',{class:'r',text:fmtN(it.nk)}),el('div',{class:'r',text:hasH?fmtN(it.hh)+' giờ':fmtN(it.n)+' lượt'}));
        r.title=it.sub||''; out.push(r);});});
    body.replaceChildren(...out);}
  t.append(hd,body); host.append(t); draw();}
function legend(host,items){host.replaceChildren(...items.map(([n,c])=>{const s=el('span'); const w=el('span',{class:'sw'}); w.style.background=c; s.append(w,document.createTextNode(n)); return s;}));}
const SST=new Map();
function tbl(host,cols,rows,empty='Chưa có dữ liệu',opt={}){
  host.replaceChildren(); if(host.nextElementSibling&&host.nextElementSibling.classList.contains('fnote')) host.nextElementSibling.remove(); if(!rows.length){host.append(el('div',{class:'empty',text:empty}));return;}
  const t=el('table'),th=el('thead'),tr=el('tr'); cols.forEach(c=>tr.append(el('th',{class:c.num?'num':'',text:c.h}))); th.append(tr); t.append(th);
  const tb=el('tbody'); rows.forEach(r=>{const x=el('tr'); x._r=r; cols.forEach(c=>{const v=c.v(r); const td=el('td',{class:c.num?'num':''}); if(v instanceof Node) td.append(v); else td.textContent=v; if(c.fv) td.dataset.fv=c.fv(r); x.append(td);}); tb.append(x);}); t.append(tb); host.append(t);
  const key=host.id||host.closest('[id]')?.id||'t';
  let renum=()=>{};
  if(opt.stt){const h0=el('th',{class:'num stt',text:'STT'}); tr.prepend(h0); [...tb.rows].forEach(x=>x.prepend(el('td',{class:'num stt'})));
    renum=()=>{let n=0; [...tb.rows].forEach(x=>{if(!x.hidden) x.cells[0].textContent=++n;});}; cols=[{h:'STT',nosort:1,num:1,v:()=>''},...cols];}
  if(!opt.sort){if(rows.length>1) addFilters(t,key); return;}
  // dòng tổng (subtotal) ngay dưới tiêu đề – tính theo các dòng đang hiển thị
  let upd=()=>{};
  if(opt.total){const st=el('tr',{class:'subtot'}); const tds=cols.map(c=>{const d=el('td',{class:c.num?'num':''}); st.append(d); return d;}); th.append(st);
    upd=()=>{renum(); const vis=[...tb.rows].filter(x=>!x.hidden).map(x=>x._r); cols.forEach((c,i)=>{const v=c.tot?c.tot(vis):''; tds[i].replaceChildren(); if(v instanceof Node) tds[i].append(v); else tds[i].textContent=v;});};
    requestAnimationFrame(()=>{const hh=tr.offsetHeight; tds.forEach(d=>d.style.top=hh+'px');});}
  // sắp xếp: bấm tiêu đề cột (bấm lần nữa để đảo chiều)
  const ss=SST.get(key)||{i:-1,d:1}; SST.set(key,ss); const ths=[...tr.cells];
  const kf=c=>c.sv||(r=>{const v=c.v(r); return v instanceof Node?v.textContent:v;});
  const cmp=(x,y)=>{const nx=x==null||x==='–'||x==='', ny=y==null||y==='–'||y==='';
    if(nx||ny) return nx&&ny?0:nx?1:-1; if(typeof x==='number'&&typeof y==='number') return x-y; if(x instanceof Date&&y instanceof Date) return x-y; return String(x).localeCompare(String(y),'vi',{numeric:true});};
  const doSort=()=>{ths.forEach((h2,i)=>h2.dataset.sort=i===ss.i?(ss.d>0?'asc':'desc'):''); if(ss.i<0||!cols[ss.i]) return; const k=kf(cols[ss.i]);
    [...tb.rows].map(x=>({x,v:k(x._r)})).sort((a,b)=>{const nb=(a.v==null||a.v==='–'||a.v==='')-(b.v==null||b.v==='–'||b.v===''); return nb||cmp(a.v,b.v)*ss.d;}).forEach(o=>tb.append(o.x)); renum();};
  ths.forEach((h2,i)=>{const c=cols[i]; if(c.filter||c.nosort) return; h2.classList.add('sortable'); h2.tabIndex=0; h2.title='Bấm để sắp xếp';
    const go=()=>{if(ss.i===i) ss.d=-ss.d; else {ss.i=i; ss.d=c.num||c.desc?-1:1;} doSort();};
    h2.addEventListener('click',e=>{if(e.target.closest('.fbtn')) return; go();}); h2.addEventListener('keydown',e=>{if(e.key==='Enter'||e.key===' '){e.preventDefault(); go();}});});
  doSort();
  const fcols=cols.map((c,i)=>c.filter?i:-1).filter(i=>i>=0);
  const upd2=()=>{renum(); upd();};
  if(fcols.length&&rows.length>1) addFilters(t,key,cols.map((c,i)=>i).filter(i=>!fcols.includes(i)),upd2); else upd2();
}
// ---- lọc theo cột (kiểu AutoFilter Excel): nút phễu ở tiêu đề, trạng thái giữ qua các lần vẽ lại
const FST=new Map(); let FPOP=null;
const FUN='<svg width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.4" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><path d="M3 5h18l-7 8.5V19l-4 2v-7.5z"/></svg>';
const cv=td=>((td&&td.dataset.fv!=null?td.dataset.fv:td?td.textContent:'')||'').replace(/\s+/g,' ').trim()||'(Trống)';
function closeFilter(){if(FPOP){FPOP.remove(); FPOP=null;}}
document.addEventListener('pointerdown',e=>{if(FPOP&&!FPOP.contains(e.target)&&!e.target.closest('.fbtn')) closeFilter();});
document.addEventListener('keydown',e=>{if(e.key==='Escape') closeFilter();});
addEventListener('scroll',e=>{if(FPOP&&!FPOP.contains(e.target)&&Date.now()-FPOP._t>400) closeFilter();},true);
function addFilters(t,key,skip=[],after=null){
  if(!t||!t.tHead||!t.tBodies[0]) return; const st=FST.get(key)||new Map(); FST.set(key,st);
  const ths=[...t.tHead.rows[0].cells]; const body=[...t.tBodies[0].rows]; const data=body.filter(r=>!r.classList.contains('sec'));
  const pass=(tr,ex)=>{for(const [i,set] of st){if(i!==ex&&!set.has(cv(tr.cells[i]))) return false;} return true;};
  const host=t.closest('.tbl')||t.parentNode;
  const apply=()=>{let n=0; data.forEach(tr=>{const ok=pass(tr,-1); tr.hidden=!ok; if(ok) n++;});
    let sec=null,any=false; const fin=()=>{if(sec) sec.hidden=!any;}; body.forEach(tr=>{if(tr.classList.contains('sec')){fin(); sec=tr; any=false;} else if(!tr.hidden) any=true;}); fin();
    ths.forEach((th,i)=>{const b=th.querySelector('.fbtn'); if(b) b.classList.toggle('on',st.has(i));});
    let nt=host.nextElementSibling&&host.nextElementSibling.classList.contains('fnote')?host.nextElementSibling:null;
    if(st.size){if(!nt){nt=el('div',{class:'fnote nopdf'}); host.after(nt);} nt.replaceChildren(document.createTextNode(`Đang lọc: hiển thị ${fmtN(n)} / ${fmtN(data.length)} dòng · `)); const a=el('a',{href:'#',text:'Bỏ tất cả bộ lọc'}); a.addEventListener('click',e=>{e.preventDefault(); st.clear(); apply();}); nt.append(a);}
    else if(nt) nt.remove(); if(after) after();};
  ths.forEach((th,i)=>{if(skip.includes(i)) return; th.querySelector('.fbtn')?.remove();
    const b=el('button',{class:'fbtn nopdf',type:'button',title:'Lọc cột này','aria-label':'Lọc cột '+th.textContent}); b.innerHTML=FUN;
    b.addEventListener('click',e=>{e.stopPropagation(); if(FPOP&&FPOP._b===b){closeFilter(); return;} openFilter(b,th.textContent,i);}); th.append(b);});
  function openFilter(b,title,i){closeFilter();
    const cnt=new Map(); data.filter(tr=>pass(tr,i)).forEach(tr=>{const v=cv(tr.cells[i]); cnt.set(v,(cnt.get(v)||0)+1);});
    (st.get(i)||new Set()).forEach(v=>{if(!cnt.has(v)) cnt.set(v,0);});
    const vals=[...cnt.keys()].sort((a,b)=>a.localeCompare(b,'vi',{numeric:true}));
    const cur=st.get(i); const sel=new Set(cur?[...cur]:vals);
    const p=el('div',{class:'fpop',role:'dialog','aria-label':'Lọc '+title}); p._b=b; p._t=Date.now();
    p.append(el('div',{class:'fh',text:'Lọc: '+title}));
    const q=el('input',{type:'search',placeholder:'Tìm giá trị…'}); p.append(q);
    const list=el('div',{class:'fl'}); p.append(list);
    const commit=()=>{if(sel.size===vals.length&&vals.every(v=>sel.has(v))) st.delete(i); else st.set(i,new Set(sel)); apply();};
    const allL=el('label',{class:'all'}); const allC=el('input',{type:'checkbox'}); allL.append(allC,document.createTextNode('(Chọn tất cả)'));
    const boxes=[];
    const syncAll=()=>{const vis=boxes.filter(x=>!x.l.hidden); allC.checked=vis.every(x=>x.c.checked); allC.indeterminate=!allC.checked&&vis.some(x=>x.c.checked);};
    vals.forEach(v=>{const l=el('label'); const c=el('input',{type:'checkbox'}); c.checked=sel.has(v); l.append(c,el('span',{text:v}),el('span',{class:'fc',text:fmtN(cnt.get(v))}));
      c.addEventListener('change',()=>{c.checked?sel.add(v):sel.delete(v); syncAll(); commit();}); boxes.push({v,l,c}); list.append(l);});
    allC.addEventListener('change',()=>{boxes.filter(x=>!x.l.hidden).forEach(x=>{x.c.checked=allC.checked; allC.checked?sel.add(x.v):sel.delete(x.v);}); syncAll(); commit();});
    list.prepend(allL); syncAll();
    q.addEventListener('input',()=>{const k=norm(q.value); boxes.forEach(x=>x.l.hidden=!!k&&!norm(x.v).includes(k)); syncAll();});
    const ff=el('div',{class:'ff'}); const clr=el('button',{class:'btn',type:'button',text:'Xóa lọc cột'}); const ok=el('button',{class:'btn primary',type:'button',text:'Xong'});
    clr.addEventListener('click',()=>{st.delete(i); apply(); closeFilter();}); ok.addEventListener('click',closeFilter); ff.append(clr,ok); p.append(ff);
    document.body.append(p); FPOP=p; const r=b.getBoundingClientRect(); const w=p.offsetWidth, hh=p.offsetHeight;
    p.style.left=Math.max(12,Math.min(innerWidth-w-12,r.left-10))+'px'; p.style.top=(r.bottom+6+hh>innerHeight-8?Math.max(8,r.top-hh-6):r.bottom+6)+'px'; q.focus({preventScroll:true});}
  // bỏ giá trị lọc không còn trong dữ liệu mới
  for(const [i,set] of [...st]){if(i>=ths.length) st.delete(i);}
  apply();
}
// 2 thanh tiến độ trong 1 ô: tỷ lệ tham dự + tỷ lệ đạt
function stackBar(att,pr,cls){const t=el('div',{class:'sbar '+(cls||'')}); const a=el('i',{class:'sa'}); const p=el('i',{class:'sp'});
  a.style.width=(att==null?0:Math.min(100,att*100))+'%'; p.style.width=(att==null||pr==null?0:Math.min(100,att*pr*100))+'%'; t.append(a,p); return t;}
function rateCell(done,reg,pass,tested){const att=reg?done/reg:null, pr=tested?pass/tested:null; const w=el('div',{class:'rcell'});
  const tx=el('div',{class:'rtx'}); tx.append(el('span',{class:'k1',text:'Đạt '}),el('b',{text:pr==null?'–':fmtP(pr)}),document.createTextNode(' / '),el('span',{class:'k2',text:'Tham dự '}),el('b',{text:att==null?'–':fmtP(att)}));
  w.append(tx,stackBar(att,pr)); w.title=`Tỷ lệ đạt ${pr==null?'–':fmtP(pr)} (${fmtN(pass)}/${fmtN(tested)} lượt có kết quả) · Tỷ lệ tham dự ${att==null?'–':fmtP(att)} (${fmtN(done)}/${fmtN(reg)} lượt)`; return w;}
function rateBars(done,reg,pass,tested){const w=el('div',{class:'rbx'});
  [['Tham dự',reg?done/reg:null,`${fmtN(done)}/${fmtN(reg)} lượt tham dự`],['Đạt',tested?pass/tested:null,tested?`${fmtN(pass)}/${fmtN(tested)} lượt đạt`:'chưa có kết quả đánh giá']].forEach(([l,p,t])=>{
    const r=el('div',{class:'rbr'}); r.title=t; r.append(el('span',{class:'l',text:l})); const tr=el('span',{class:'t'}); const i=el('i'); i.style.width=(p==null?0:Math.min(100,p*100))+'%'; if(l==='Đạt') i.classList.add('ok'); tr.append(i); r.append(tr,el('b',{text:p==null?'–':fmtP(p)})); w.append(r);});
  return w;}
function ratioBar(done,reg,mx){const w=el('div',{class:'rb'}); const t=el('div',{class:'tr'}); t.style.width=Math.max(4,reg/mx*100)+'%'; const i=el('i'); i.style.width=(reg?done/reg*100:0)+'%'; t.append(i);
  const tw=el('div',{class:'rbw'}); tw.append(t); const tx=el('span',{class:'tx'}); tx.append(el('b',{text:fmtN(done)}),document.createTextNode('/'+fmtN(reg))); w.append(tw,tx);
  w.title=`${fmtN(done)} hoàn thành / ${fmtN(reg)} đăng ký`+(reg?` (${fmtP(done/reg)})`:''); return w;}
const STAR='<path d="M8 1.2l2 4.3 4.6.6-3.4 3.2.9 4.6L8 11.6l-4.1 2.3.9-4.6L1.4 6.1 6 5.5z"/>';
function stars(v){const w=el('span',{class:'stars'}); if(v==null){w.append(el('span',{class:'sv',text:'–'})); return w;}
  const mk=c=>{let s=''; for(let i=0;i<5;i++) s+=`<svg width="16" height="16" viewBox="0 0 16 16" fill="${c}" aria-hidden="true">${STAR}</svg>`; return s;};
  const sx=el('span',{class:'sx'}); sx.innerHTML=mk('#e3e7ef'); const fg=el('span',{class:'fg'}); fg.innerHTML=mk('#f2a33a'); fg.style.width=(v/5*100)+'%'; sx.append(fg);
  w.append(sx,el('span',{class:'sv',text:fmt1(v)})); w.title=fmt1(v)+' / 5'; w.setAttribute('aria-label',fmt1(v)+' trên 5 sao'); return w;}
const STC={'Hoàn thành':'good','Đã xác nhận':'s2','Kế hoạch':'muted','Dời lịch':'warn','Hủy':'crit','Đang thực hiện':'warn','Trễ hạn':'crit','Khi phát sinh':'axis'};
function pill(s){const p=el('span',{class:'pill'}); const d=el('span',{class:'dot'}); d.style.background=`var(--${STC[s]||'muted'})`; p.append(d,document.createTextNode(s)); return p;}

// ---------- render
function render(){
  if(M) renderNV();
  if(!M) return; ['perf','budget','sched','years','data'].forEach(p=>$('#p-'+p).hidden=p!==S.tab); const F=filt();
  $('#srcChip').textContent=source.sample?'Dữ liệu mẫu':source.name; $('#srcChip').className='chip '+(source.sample?'sample':'ok');
  $('#rowsChip').textContent=`${fmtN(M.emp.size)} NV · ${fmtN(M.cls.size)} lớp · ${fmtN(M.recs.length)} lượt ghi nhận`;
  $('#ffKhoi').classList.toggle('on',F.K!=='*'); $('#ffFrom').classList.toggle('on',!!F.D0); $('#ffPb').classList.toggle('on',F.PB!=='*'); $('#ffNhom').classList.toggle('on',F.N!=='*'); $('#ffYear').classList.add('on'); $('#ffCmp').classList.toggle('on',!!$('#fCmp').value);
  {const C=cmpF(), n=$('#cmpNote'); n.hidden=!C||S.tab==='years'||S.tab==='data'; if(C) n.textContent=C.cut?`Đang so sánh cùng kỳ: 01/01–${dmy(F.T).slice(0,5)} năm ${F.Y} với cùng giai đoạn năm ${C.Y}. Biểu đồ theo tháng hiện đủ 12 tháng của năm ${C.Y} để tham khảo.`:`Đang so sánh cả năm ${F.Y} với cả năm ${C.Y}.`;}
  renderAlerts(); renderPerf(F); renderBud(F); renderSch(F); renderYears(F); renderData();
  document.querySelectorAll('.tab').forEach(t=>t.setAttribute('aria-selected',String(t.dataset.p===S.tab)));
  ['perf','budget','sched','years','data'].forEach(p=>$('#p-'+p).hidden=p!==S.tab);
}
function renderAlerts(){
  const host=$('#alerts'); host.replaceChildren(); const n=M.badRec.length+M.badCls.length; if(!n) return;
  const a=el('div',{class:'alert'}); a.append(el('b',{text:`File có ${n} dòng cần kiểm tra. `}),document.createTextNode('Các dòng lỗi không được tính hoặc tính thiếu thông tin. Xem chi tiết ở tab "Dữ liệu & hướng dẫn".'));
  host.append(a);
}
function renderPerf(F){
  const P=perfM(F); const C=cmpF(); const Q=C?perfM(C):null; const Yc=C&&C.Y; const D=(k,kind,g)=>Q?dlt(P[k],Q[k],(C.cut?'cùng kỳ ':'')+Yc,kind,g):null;
  const done=P.done;
  kpis($('#kPerf'),[
    {i:'users',l:'Số NV được đào tạo',v:fmtN(P.uniq),of:fmtN(P.hc),p:P.cov,n:`độ phủ ${fmtP(P.cov)} NV đang làm`+(P.learners>P.uniq?` · ${fmtN(P.learners-P.uniq)} người đã nghỉ`:''),d:D('cov','pp',1)},
    {i:'star',u:P.satN?'/ 5':'',l:'Mức độ hài lòng',v:P.satN?fmt1(P.sat):'–',n:`${fmtN(P.satN)} lượt đánh giá`,d:D('sat','abs',1)},
    {i:'clock',u:P.hasH?'giờ':'',l:'Tổng số giờ đào tạo',v:P.hasH?fmtN(P.hrs):'–',n:P.hasH?`${fmtN(P.attN)} lượt tham dự`:`${fmtN(P.attN)} lượt tham dự · chưa có số giờ (điền Thời lượng chuẩn ở DM Khoa hoc)`,d:D('hrs','pct',1)},
    {i:'gauge',u:P.hasH?`giờ / ${fmtN(P.hc)} NV`:`/ ${fmtN(P.hc)} NV`,l:'Số giờ đào tạo / NV',v:fmt1(P.hpn),n:P.hasH?`${fmtN(P.hrs)} giờ ÷ ${fmtN(P.hc)} NV đang làm`:'chưa có số giờ',d:D('hpn','pct',1)},
    {i:'check',l:'Tỷ lệ đạt / Tỷ lệ tham dự',wide:1,dual:{att:P.att,pr:P.pr},multi:[
      {l:'Tỷ lệ đạt',p:P.pr,n:`${fmtN(P.pass)} / ${fmtN(P.tested)} lượt có kết quả đánh giá đạt`,d:D('pr','pp',1)},
      {l:'Tỷ lệ tham dự',p:P.att,n:`${fmtN(P.attN)} / ${fmtN(P.attD)} lượt được mời đã tham dự`,d:D('att','pp',1)}]},
    (()=>{const t=P.hrs||0, f=(m)=>{const h=P.done.filter(r=>norm(r.cl.nb)===m).reduce((a,r)=>a+r.hrs,0); const k=new Set(P.done.filter(r=>norm(r.cl.nb)===m).map(r=>r.cl.kh)).size; return {h,k};}; const a=f('noi bo'); return {i:'home',u:P.hasH?'giờ':'',l:'Số giờ đào tạo nội bộ',v:P.hasH?fmtN(a.h):'–',p:t?a.h/t:null,n:`${fmtP(t?a.h/t:null)} tổng giờ · ${fmtN(a.k)} khóa do SCC tự đào tạo`,d:Q?dlt(a.h,(()=>{const x=Q.done.filter(r=>norm(r.cl.nb)==='noi bo').reduce((s,r)=>s+r.hrs,0); return x;})(),(C.cut?'cùng kỳ ':'')+Yc,'pct',1):null};})(),
    (()=>{const t=P.hrs||0; const dn=P.done.filter(r=>norm(r.cl.nb)==='ben ngoai'); const h=dn.reduce((a,r)=>a+r.hrs,0); const k=new Set(dn.map(r=>r.cl.kh)).size; return {i:'ext',u:P.hasH?'giờ':'',l:'Số giờ đào tạo bên ngoài',v:P.hasH?fmtN(h):'–',p:t?h/t:null,n:`${fmtP(t?h/t:null)} tổng giờ · ${fmtN(k)} khóa thuê / cử đi học bên ngoài`,d:Q?dlt(h,Q.done.filter(r=>norm(r.cl.nb)==='ben ngoai').reduce((s,r)=>s+r.hrs,0),(C.cut?'cùng kỳ ':'')+Yc,'pct',0):null};})(),
  ]);
  // month
  const mm=(S.month==='hours'&&!P.hasH)?'lượt':S.month; const mv=(F2,dn)=>MONTHS.map((_,m)=>{const d=dn.filter(r=>r.month===m); return mm==='hours'?d.reduce((a,r)=>a+r.hrs,0):mm==='nv'?new Set(d.filter(r=>r.e&&r.e.join&&r.e.join.getFullYear()===F2.Y).map(r=>r.eid)).size:d.length;});
  const nm={hours:'Giờ đào tạo',lượt:'Lượt học',nv:'NV mới được đào tạo'}[mm];
  $('#mTitle').textContent=nm+' theo tháng'+(S.month==='hours'&&!P.hasH?' (chưa có số giờ)':'');
  const ser=Q?[{name:'Năm '+Yc,color:'var(--navy)',values:mv(C.full,perfM(C.full).done)},{name:'Năm '+F.Y,color:'var(--accent)',values:mv(F,done)}]:[{name:nm,color:'var(--accent)',values:mv(F,done)}];
  legend($('#lgMonth'),Q?ser.map(s=>[s.name,s.color]):[]);
  vbars($('#cMonth'),MONTHS,ser,fmtN,{tickFmt:fmtN,label:v=>fmtN(v)});
  // khoi (not filtered by khoi)
  const allActive=[...M.emp.values()].filter(e=>e.khoi&&activeAt(e,F.R)); const khois=[...new Set(allActive.map(e=>e.khoi))].sort();
  const yrDone=M.recs.filter(r=>r.year===F.Y&&r.done&&(F.N==='*'||r.cl.nhom===F.N));
  const kHasH=yrDone.some(r=>r.hrs>0);
  hbars($('#cKhoi'),khois.map(k=>{const ak=new Set(allActive.filter(e=>e.khoi===k).map(e=>e.id)); const h=ak.size; const dk=yrDone.filter(r=>r.khoi===k); const u=new Set(dk.filter(r=>ak.has(r.eid)).map(r=>r.eid)).size; const hh=dk.reduce((a,r)=>a+r.hrs,0);
    return {label:k,value:h?u/h:0,hh,n:dk.length,cnt:`${fmtN(u)} / ${fmtN(h)}`,txt:kHasH?`${fmtN(hh)} giờ`:`${fmtN(dk.length)} lượt`,sub:`${fmtN(u)}/${fmtN(h)} NV đang làm đã học · ${kHasH?fmtN(hh)+' giờ · BQ '+fmt1(h?hh/h:0)+' giờ/NV':fmtN(dk.length)+' lượt học'}`};}).sort((a,b)=>b.value-a.value||b.hh-a.hh||b.n-a.n),fmtP,{max:1,heads:['Khối','',  'NV đang làm đã học / tổng','Tỷ lệ',kHasH?'Giờ học (gồm NV đã nghỉ)':'Lượt học (gồm NV đã nghỉ)']});
  // phòng ban (theo bộ lọc Khối, không lọc Phòng ban)
  {const ap=allActive.filter(e=>e.pb&&(F.K==='*'||e.khoi===F.K)); const pbs=[...new Set(ap.map(e=>e.pb))];
   const items=pbs.map(p=>{const ak=new Set(ap.filter(e=>e.pb===p).map(e=>e.id)); const h=ak.size; const dk=yrDone.filter(r=>r.e&&r.e.pb===p); const u=new Set(dk.filter(r=>ak.has(r.eid)).map(r=>r.eid)).size; const hh=dk.reduce((a,r)=>a+r.hrs,0);
     const kc=new Map(); ap.filter(e=>e.pb===p).forEach(e=>kc.set(e.khoi,(kc.get(e.khoi)||0)+1)); const khoi=[...kc.entries()].sort((x,y)=>y[1]-x[1])[0][0];
     const ks=new Set(dk.map(r=>r.cl.kh)); return {label:p,khoi,u,h,ks,nk:ks.size,value:h?u/h:0,hh,n:dk.length,cnt:`${fmtN(u)} / ${fmtN(h)}`,txt:kHasH?`${fmtN(hh)} giờ`:`${fmtN(dk.length)} lượt`,sub:`${fmtN(u)}/${fmtN(h)} NV đang làm đã học · ${kHasH?fmtN(hh)+' giờ':fmtN(dk.length)+' lượt học'}`};}).sort((a,b)=>b.value-a.value||b.n-a.n);
   pbTable($('#cPb'),items,kHasH);
   {const ph=$('#pbHint'); if(ph) ph.textContent=`% NV đang làm đã học ≥ 1 lớp · kèm ${kHasH?'tổng giờ':'số lượt học'} · xếp giảm dần · ${items.length} phòng ban${F.K==='*'?'':' thuộc '+F.K}`;}}
  // nhom
  const nh=[...new Set([...M.crs.values()].map(c=>c.nhom))];
  const byH=P.hasH; const nhv=nh.map(n=>({n,v:done.filter(r=>r.cl.nhom===n).reduce((a,r)=>a+(byH?r.hrs:1),0)})).sort((a,b)=>b.v-a.v);
  $('#nTitle').textContent=(byH?'Giờ đào tạo':'Lượt học hoàn thành')+' theo nhóm khóa học';
  const drawNhom=h=>vbars($('#cNhom'),nhv.map(x=>x.n),[{name:byH?'Giờ đào tạo':'Lượt học',color:'var(--accent)',values:nhv.map(x=>x.v)}],fmtN,{h,label:fmtN});
  // course table
  const byC=new Map(); F.recs.forEach(r=>{const k=r.cl.kh; if(!byC.has(k)) byC.set(k,{name:r.cl.name,nhom:r.cl.nhom,nb:r.cl.nb,emp:new Set(),cls:new Set(),reg:0,done:0,hrs:0,p:0,f:0,sat:[]}); const o=byC.get(k); o.cls.add(r.cl.key); o.reg++; if(r.done){o.done++; o.emp.add(r.eid); o.hrs+=r.hrs; if(r.res==='pass')o.p++; if(r.res==='fail')o.f++; if(r.sat>0)o.sat.push(r.sat);}});
  const nbo=x=>{const n=norm(x.nb); return n==='noi bo'?0:n==='ben ngoai'?2:1;}; const cr=[...byC.values()].sort((a,b)=>nbo(a)-nbo(b)||b.done-a.done||b.reg-a.reg); const mxReg=Math.max(1,...cr.map(r=>r.reg));
  const SUM=(R,f)=>R.reduce((x,r)=>x+(f(r)||0),0);
  $('#cTitle')&&($('#cTitle').textContent=`Theo khóa học – năm ${F.Y}`);
  tbl($('#tCourse'),[
    {h:'Khóa học',v:r=>r.name,sv:r=>r.name,tot:R=>`Tổng · ${fmtN(R.length)} khóa`},
    {h:'Phân loại',v:r=>r.nb||'–',sv:r=>{const n=norm(r.nb); return n==='noi bo'?0:n==='ben ngoai'?1:2;}},
    {h:'Nhóm',v:r=>r.nhom,sv:r=>r.nhom},
    {h:'Số lần tổ chức',num:1,v:r=>fmtN(r.cls.size),sv:r=>r.cls.size,tot:R=>fmtN(SUM(R,r=>r.cls.size))},
    {h:'Số NV học',num:1,v:r=>fmtN(r.emp.size),sv:r=>r.emp.size,tot:R=>{const u=new Set(R.flatMap(r=>[...r.emp])); const act=[...u].filter(id=>{const e=M.emp.get(id); return e&&activeAt(e,F.R);}).length; const d=el('span'); d.append(document.createTextNode(fmtN(u.size))); d.title=`${fmtN(act)} NV đang làm + ${fmtN(u.size-act)} NV đã nghỉ`; d.append(el('small',{class:'sub2',text:`${fmtN(act)} đang làm`})); return d;}},
    {h:'Lượt tham dự',num:1,v:r=>fmtN(r.done),sv:r=>r.done,tot:R=>fmtN(SUM(R,r=>r.done))},
    {h:'Giờ',num:1,v:r=>r.hrs?fmtN(r.hrs):'–',sv:r=>r.hrs||null,tot:R=>fmtN(SUM(R,r=>r.hrs))},
    {h:'Tỷ lệ đạt / Tỷ lệ tham dự',v:r=>rateCell(r.done,r.reg,r.p,r.p+r.f),sv:r=>(r.p+r.f?r.p/(r.p+r.f):0)+(r.reg?r.done/r.reg:0)/1000,desc:1,tot:R=>rateCell(SUM(R,r=>r.done),SUM(R,r=>r.reg),SUM(R,r=>r.p),SUM(R,r=>r.p+r.f))},
    {h:'Hài lòng',v:r=>stars(r.sat.length?avg(r.sat):null),sv:r=>r.sat.length?avg(r.sat):null,desc:1,tot:R=>{const a=R.flatMap(r=>r.sat); return stars(a.length?avg(a):null);}}],cr,'Chưa có dữ liệu',{sort:1,total:1,stt:1});
  renderInsights(F,P,C,Q);
  drawNhom(Math.max(300,Math.round($('#insights').offsetHeight)));  // cao bằng ô Tóm tắt nhanh bên cạnh
}

// ---------- insights: short, factual sentences computed from the current filter (no invented targets)
const vt=v=>fmtVND(v).replace(/ tr$/,' triệu đ').replace(/ tỷ$/,' tỷ đ').replace(/(\d)$/,'$1 đ');
function rich(t){const f=document.createDocumentFragment(); t.split(/(\*\*[^*]+\*\*)/).forEach(p=>{if(!p) return; if(p.startsWith('**')) f.append(el('b',{text:p.slice(2,-2)})); else f.append(document.createTextNode(p));}); return f;}
// Phân tích nhanh: mỗi điểm = nhận định + ý nghĩa + gợi ý hành động
function insRender(host,out){host.replaceChildren(...out.map(o=>{const d=el('div',{class:'it'}); const ic=el('span',{class:'ic '+(o.kind||'')}); ic.append(icon(o.ic)); const p=el('div');
  p.append(el('div',{class:'ih',text:o.title})); if(o.body){const x=el('div',{class:'ib'}); x.append(rich(o.body)); p.append(x);}  d.append(ic,p); return d;}));}
const evalNote=()=>{const ys=YEARS.slice().sort(); const r=ys.length>1?`Giai đoạn ${ys[0]}–${ys[ys.length-1]}`:`Năm ${ys[0]}`; return `${r} chưa thực hiện đánh giá sau khóa học; nội dung này sẽ được bổ sung trong thời gian tới.`;};
function renderInsights(F,P,C,Q){
  const out=[]; const add=(ic,kind,title,body)=>out.push({ic,kind,title,body});
  const done=P.done; const actIds=new Set(F.emps.map(e=>e.id)); const host=$('#insights');
  if(!done.length){add('cal','info','Chưa có dữ liệu trong kỳ','Chưa có lượt học hoàn thành nào theo bộ lọc đang chọn.'); insRender(host,out); return;}
  // 1. độ phủ & chiều sâu
  const pe=new Map(); done.forEach(r=>{if(!actIds.has(r.eid)) return; let o=pe.get(r.eid); if(!o){o={k:new Set()}; pe.set(r.eid,o);} o.k.add(r.cl.kh);});
  const n1=[...pe.values()].filter(o=>o.k.size===1).length, n3=[...pe.values()].filter(o=>o.k.size>=3).length;
  if(P.hc&&pe.size) add('users',P.cov>=0.9?'good':'info',`Độ phủ ${fmtP(P.cov)} – nền tảng tốt để đi sâu`,
    `${fmtN(P.uniq)}/${fmtN(P.hc)} NV đang làm đã tham gia ít nhất 1 khóa, bình quân ${fmt1(P.hpn)} giờ/NV. ${fmtN(n1)} NV (${fmtP(n1/pe.size)}) học 1 khóa, ${fmtN(n3)} NV (${fmtP(n3/pe.size)}) học từ 3 khóa trở lên – độ phủ hiện đến nhiều từ các khóa bắt buộc và hội nhập, chiều sâu chuyên môn sẽ tăng dần khi các khóa theo nhu cầu được triển khai.`);
  // 2. chương trình diện rộng
  if(P.hrs){const ch=new Map(); done.forEach(r=>ch.set(r.cl.name,(ch.get(r.cl.name)||0)+r.hrs)); const cs=[...ch.entries()].sort((a,b)=>b[1]-a[1]);
    if(cs.length>2){const t2=cs[0][1]+cs[1][1]; if(t2/P.hrs>=0.4) add('trend','info',`2 chương trình diện rộng đóng góp ${fmtP(t2/P.hrs)} số giờ`,
      `**${cs[0][0]}** (${fmtN(cs[0][1])} giờ) và **${cs[1][0]}** (${fmtN(cs[1][1])} giờ) có số học viên lớn nên kéo tổng giờ lên đáng kể; ${fmtN(cs.length-2)} khóa còn lại quy mô nhỏ hơn, tập trung vào nhu cầu chuyên môn của từng phòng.`);}}
  // 3. nội bộ
  const B=budM(F);
  if(P.hrs){const hI=done.filter(r=>norm(r.cl.nb)==='noi bo').reduce((a,r)=>a+r.hrs,0), hE=P.hrs-hI; const cI=F.cls.filter(c=>norm(c.nb)==='noi bo').reduce((a,c)=>a+(c.ca||0),0), cE=B.act-cI;
    if(hI&&hE&&cE>0) add('home','good',`Đào tạo nội bộ đóng góp ${fmtP(hI/P.hrs)} số giờ`,
      `Các khóa nội bộ gần như không phát sinh chi phí, trong khi giờ học bên ngoài bình quân ${vt(cE/hE)}/giờ – nguồn lực nội bộ đang được tận dụng hiệu quả.`);}
  // 4. phòng ban
  {const ap=F.emps.filter(e=>e.pb); const lset=new Set(done.map(r=>r.eid)); const m=new Map(); ap.forEach(e=>{const o=m.get(e.pb)||{h:0,u:0}; o.h++; if(lset.has(e.id)) o.u++; m.set(e.pb,o);});
    const all=[...m.entries()].filter(([,o])=>o.h>=5); const low=all.filter(([,o])=>o.u/o.h<0.9).map(([k,o])=>({k,c:o.u/o.h,o})).sort((a,b)=>a.c-b.c);
    if(low.length) add('gauge','info',`${all.length-low.length}/${all.length} phòng ban đạt độ phủ từ 90% trở lên`,`Các phòng còn lại: `+low.slice(0,4).map(x=>`**${x.k}** ${fmtP(x.c)} (${fmtN(x.o.u)}/${fmtN(x.o.h)})`).join(' · ')+'.');}
  // 5. NV chưa tham gia
  if(P.hc){const lset=new Set(done.map(r=>r.eid)); const no=F.emps.filter(e=>!lset.has(e.id));
    if(no.length) add('user','info',`${fmtN(no.length)} NV chưa tham gia khóa nào`,(()=>{const g=new Map(); no.forEach(e=>{const k=e.pb||e.khoi||'Chưa rõ'; g.set(k,(g.get(k)||0)+1);}); const gs=[...g.entries()].sort((a,b)=>b[1]-a[1]);
      const R=F.R||new Date(), nw=no.filter(e=>e.join&&(R-e.join)<=90*864e5).length;
      let t=`Chiếm ${fmtP(no.length/P.hc)} NV đang làm, tập trung ở `+gs.slice(0,3).map(([k,v])=>`**${k}** (${fmtN(v)} NV)`).join(', ')+(gs.length>3?` và ${fmtN(gs.length-3)} phòng khác`:'')+'.';
      if(nw) t+=` Trong đó ${fmtN(nw)} NV mới vào làm trong 3 tháng gần đây.`; return t;})());}
  // 6. tiến độ kế hoạch
  const Sm=schM(F);
  if(Sm.byPlan){const late=Sm.its.filter(x=>x.k==='late'); const lb=late.filter(x=>/bat buoc/.test(norm(x.p.sec)));
    add('cal','info',`Đã triển khai ${fmtN(Sm.num)}/${fmtN(Sm.den)} khóa kế hoạch (${fmtP(Sm.rate)})`,
      `${fmtN(Sm.nRun)} khóa đang thực hiện`+(late.length?`, ${fmtN(late.length)} khóa đang chờ xếp lịch`+(lb.length?` (gồm ${fmtN(lb.length)} khóa bắt buộc: ${lb.map(x=>x.p.name.replace(/\s*\(Đợt T\d+\)$/,'')).slice(0,3).join('; ')})`:''):'')+'.');
  // 7. ngân sách
    if(B.plan){const u=B.act/B.plan; const lateBud=late.reduce((a,x)=>a+(x.p.bud||0),0);
      add('coin','info',`Đã sử dụng ${fmtP(u)} ngân sách`,lateBud?`Phần lớn ngân sách còn lại (${vt(lateBud)}) thuộc các khóa đang chờ xếp lịch; nếu triển khai đủ, chi phí cả năm khoảng ${vt(B.act+lateBud)} (${fmtP((B.act+lateBud)/B.plan)} ngân sách) – vẫn nằm trong ngân sách.`:`${vt(B.act)} / ${vt(B.plan)}.`);}}
  // 8. người đã nghỉ
  if(P.hrs){const lv=done.filter(r=>!actIds.has(r.eid)); const hL=lv.reduce((a,r)=>a+r.hrs,0); const nL=new Set(lv.map(r=>r.eid)).size;
    if(nL) add('ext','info',`${fmtP(1-hL/P.hrs)} giờ đào tạo thuộc về NV đang làm`,`${fmtN(nL)} người đã học trong kỳ rồi nghỉ việc (${fmtN(hL)} giờ, ${fmtP(hL/P.hrs)} tổng giờ).`);}
  // 9. đánh giá
  if(!P.satN) add('star','info','Chưa có dữ liệu đánh giá sau khóa',evalNote());
  else add('star',P.sat>=4?'good':'info',`Hài lòng bình quân ${fmt1(P.sat)}/5`,`${fmtN(P.satN)} lượt đánh giá.`);
  insRender(host,out);
}
function renderNV(){
  const host=$('#nvRes'); if(!M) return; const q=norm($('#qNV').value);
  if(q.length<2){host.replaceChildren(el('p',{class:'hint',text:'Nhập ít nhất 2 ký tự để tìm.'})); return;}
  const emps=[...M.emp.values()].filter(e=>norm(e.id).includes(q)||norm(e.name).includes(q)).slice(0,30);
  if(!emps.length){host.replaceChildren(el('p',{class:'hint',text:'Không tìm thấy nhân viên.'})); return;}
  const ids=new Set(emps.map(e=>e.id)); const recs=M.recs.filter(r=>ids.has(r.eid)).sort((a,b)=>b.cl.d0-a.cl.d0||a.eid.localeCompare(b.eid));
  const RES={pass:'Đạt',fail:'Không đạt',na:'–',nodata:'Chưa có điểm','':'–'};
  const sum=el('p',{class:'hint'}); sum.textContent=`${emps.length} nhân viên khớp${emps.length===30?' (hiện tối đa 30)':''} · ${fmtN(recs.length)} lượt học`;
  const d=el('div',{class:'tbl'});
  tbl(d,[{h:'Mã NV',v:r=>r.eid},{h:'Họ tên',v:r=>r.e?r.e.name:''},{h:'Phòng ban',v:r=>r.e?r.e.pb:''},{h:'Ngày học',v:r=>dmy(r.cl.d0)},{h:'Khóa học',v:r=>r.cl.name},{h:'Nhóm',v:r=>r.cl.nhom},{h:'Tham dự',v:r=>r.st},{h:'Giờ',num:1,v:r=>r.hrs?fmt1(r.hrs):'–'},{h:'Kết quả',v:r=>RES[r.res]??'–'}],recs,'Nhân viên chưa có lượt học nào trong file');
  const noRec=emps.filter(e=>!recs.some(r=>r.eid===e.id));
  host.replaceChildren(sum,d); if(noRec.length) host.append(el('p',{class:'hint',text:'Chưa có lượt học: '+noRec.map(e=>e.id+' '+e.name).join(', ')}));
}
function renderCal(F){
  const host=$('#calSch'); host.classList.add('cal'); host.replaceChildren(); $('#calTitle').textContent=`Lịch đào tạo năm ${F.Y}: kế hoạch và thực hiện`;
  const by=new Map(); F.cls.forEach(c=>{if(!by.has(c.kh)) by.set(c.kh,{kh:c.kh,name:c.name,nhom:c.nhom,plan:false,ss:[]}); const o=by.get(c.kh); o.ss.push(c); if(c.plan) o.plan=true;});
  const rows=[...by.values()].sort((a,b)=>(b.plan-a.plan)||(Math.min(...a.ss.map(c=>c.d0))-Math.min(...b.ss.map(c=>c.d0))));
  if(!rows.length){host.append(el('div',{class:'empty',text:'Chưa có lớp nào trong năm này'})); return;}
  const curM=F.T.getFullYear()===F.Y?F.T.getMonth():-1;
  const state=c=>c.st==='Hoàn thành'?'done':c.st==='Hủy'?'cx':c.st==='Dời lịch'?'dl':(c.d0<=F.T?'late':'up');
  const LBL={done:'Đã tổ chức',up:'Dự kiến',late:'Quá hạn, chưa cập nhật',dl:'Dời lịch',cx:'Hủy'};
  const t=el('table'), th=el('thead'), hr=el('tr'); hr.append(el('th',{text:'Khóa học'}),el('th',{text:'Nhóm'})); MONTHS.forEach((m,i)=>hr.append(el('th',{class:'m'+(i===curM?' now':''),text:m}))); hr.append(el('th',{text:'Thực hiện'})); th.append(hr); t.append(th);
  const tb=el('tbody'); let sec=null;
  rows.forEach(r=>{const s=r.plan?'Trong kế hoạch':'Ngoài kế hoạch'; if(s!==sec){sec=s; const tr=el('tr',{class:'sec'}); const td=el('td',{text:s}); td.colSpan=15; tr.append(td); tb.append(tr);}
    const tr=el('tr'); const nm=el('td',{class:'nm'}); nm.append(document.createTextNode(r.name)); nm.append(el('small',{text:r.kh})); tr.append(nm,el('td',{text:r.nhom}));
    MONTHS.forEach((_,m)=>{const td=el('td',{class:'m'+(m===curM?' now':'')}); r.ss.filter(c=>c.d0.getMonth()===m).forEach(c=>{const k=state(c); const i=el('i',{class:'mk '+k}); i.title=`${dmy(c.d0)} · ${LBL[k]}${c.st!=='Kế hoạch'&&k!=='late'&&k!=='up'?'':' ('+c.st+')'}${c.reg?' · '+c.done+'/'+c.reg+' học viên':''}${c.note?' · '+c.note:''}`; td.append(i);}); tr.append(td);});
    const done=r.ss.filter(c=>c.st==='Hoàn thành').length, late=r.ss.filter(c=>state(c)==='late').length;
    tr.append(el('td',{text:`${done}/${r.ss.length} lớp`+(late?` · ${late} quá hạn`:'')})); tb.append(tr);});
  t.append(tb); host.append(t);
}
// nhóm kế hoạch của 1 lớp: theo Mã lớp trong sheet Kế hoạch → nếu Ngoài kế hoạch = Phát sinh → còn lại theo nhóm khóa học
const LO=s=>String(s||'').replace(/^[IVXLC]+\.\s*/,'').trim();
function planRowOf(F,c){const rs=F.plan.filter(p=>p.code&&p.code===c.kh); if(!rs.length) return null; if(rs.length===1) return rs[0];
  return rs.find(p=>p.dn&&+p.dn===+c.d0)||rs.find(p=>p.s&&c.d0>=p.s&&c.d0<=new Date(+p.e+864e5))||rs[0];}
function typeFn(F){return c=>{if(!c.plan) return 'Phát sinh'; const p=planRowOf(F,c); return p?LO(p.sec):(c.nhom==='Bắt buộc'?'Bắt buộc':'Theo nhu cầu');};}
function renderBud(F){
  const B=budM(F); const typeOf=typeFn(F); const C=cmpF(); const Q=C?budM(C):null; const Yc=C&&C.Y; const D=(k,kind,g)=>Q?dlt(B[k],Q[k],(C.cut?'cùng kỳ ':'')+Yc,kind,g):null;
  const vu=v=>{const t=fmtVND(v); const m=t.match(/^(.*) (tr|tỷ)$/); return m?[m[1],m[2]==='tr'?'triệu đ':'tỷ đ']:[t,'đ'];};
  // ---- thẻ ngân sách tổng: kế hoạch / đã dùng / còn lại + thanh tiến độ
  {const h=$('#bHero'); h.replaceChildren();
   const hd=el('div',{class:'bh-hd'}); const ic=el('span',{class:'ico'}); ic.append(icon('coin')); const tt=el('div'); tt.append(el('h3',{text:`Ngân sách đào tạo năm ${F.Y}`})); hd.append(ic,tt); h.append(hd);
   const u=B.plan?B.act/B.plan:null; const over=B.act>B.plan&&B.plan>0;
   const nums=el('div',{class:'bh-nums'});
   const nb=(lab,v,dot,pc)=>{const d=el('div'); const l=el('span',{class:'lb'}); if(dot){const i=el('i'); i.style.background=dot; l.append(i);} l.append(document.createTextNode(lab)); const [a,b]=vu(v); const x=el('div',{class:'vv'}); x.append(el('b',{text:a}),el('small',{text:b})); if(pc!=null) x.append(el('em',{text:fmtP(pc)})); d.append(l,x); return d;};
   nums.append(nb('Ngân sách kế hoạch',B.plan,null,null),nb('Đã sử dụng',B.act,over?'var(--crit)':'var(--accent)',u),nb(over?'Vượt ngân sách':'Còn lại',Math.abs(B.rem),over?'var(--crit)':'#d5dae5',B.plan?Math.abs(B.rem)/B.plan:null));
   h.append(nums);
   const bar=el('div',{class:'bh-bar'}); const fill=el('i'); fill.style.width=Math.min(100,(u||0)*100)+'%'; if(over) fill.style.background='var(--crit)'; bar.append(fill);
   const lab=el('span',{class:'bh-pc',text:u==null?'–':fmtP(u)+' đã dùng'}); bar.append(lab);
   const y0=new Date(F.Y,0,1), y1=new Date(F.Y+1,0,1); const tp=F.T<y0?0:F.T>=y1?1:(F.T-y0)/(y1-y0);
   if(tp>0&&tp<1){const nw=el('span',{class:'bh-now'}); nw.style.left=(tp*100)+'%'; nw.append(el('span',{text:`${dmy(F.T).slice(0,5)} · Tiến độ năm ${fmtP(tp)}`})); bar.append(nw);}
   h.append(bar);
   const sc=el('div',{class:'bh-scale'}); ['0%','25%','50%','75%','100%'].forEach(t=>sc.append(el('span',{text:t}))); h.append(sc);
   const note=el('div',{class:'bh-note'});
   note.textContent=Q?'':B.plan?`Đã dùng ${fmtP(u)} ngân sách (${vt(B.act)} / ${vt(B.plan)})`+(tp>0&&tp<1?`, trong khi tiến độ thời gian năm ${F.Y} là ${fmtP(tp)} (tính đến ${dmy(F.T).slice(0,5)}).`:'.')+(B.nDone?` ${fmtN(B.nDone)} lớp đã tổ chức trong năm.`:''):'Chưa có ngân sách kế hoạch cho năm này.';
   if(Q){const d=D('act','pct',0); const w=el('span',{class:'dl '+(d.c||''),text:d.t}); note.append(document.createTextNode(' Chi phí thực tế '),w,document.createTextNode(' '+d.e.trim()+'.'));}
   if(Q) h.append(note);
   // phân bổ theo nhóm kế hoạch
   if(B.hasP){const secs=[...new Set(F.plan.map(p=>p.sec))].sort((a,b)=>a.localeCompare(b,'vi',{numeric:true})); const c2s=new Map(); F.plan.forEach(p=>{if(p.code&&!c2s.has(p.code)) c2s.set(p.code,p.sec);});
     const ps=secs.find(x=>/phat sinh/.test(norm(x))); const secOf=c=>{const t=typeOf(c); return secs.find(x=>LO(x)===t)||t;};
     const act=new Map(); F.cls.forEach(c=>{if(c.ca){const k=secOf(c); act.set(k,(act.get(k)||0)+c.ca);}});
     const all=[...secs,...[...act.keys()].filter(k=>!secs.includes(k))];
     const box=el('div',{class:'bh-secs'});
     all.forEach(k=>{const pl=F.plan.filter(p=>p.sec===k).reduce((a,p)=>a+p.bud,0), ac=act.get(k)||0; const r=el('div',{class:'bs'});
       const n=el('div',{class:'bs-n'}); n.append(document.createTextNode(k==='Phát sinh'?'IV. Phát sinh':k)); if(k==='Phát sinh') n.append(el('small',{text:' (ngoài kế hoạch, chưa có ngân sách)'}));
       const tr=el('div',{class:'bs-t'}); const i=el('i'); i.style.width=(pl?Math.min(100,ac/pl*100):ac?100:0)+'%'; if(ac>pl) i.style.background='var(--crit)'; tr.append(i);
       const v=el('div',{class:'bs-v'}); v.append(el('b',{text:fmtVND(ac)}),document.createTextNode(' / '+fmtVND(pl))); r.append(n,tr,v); box.append(r);});
     const hh=el('div',{class:'bs bs-h'}); hh.append(el('div',{text:'Theo nhóm kế hoạch'}),el('div'),el('div',{text:'Thực tế / Kế hoạch'})); box.prepend(hh); h.append(box);}
  }
  kpis($('#kBud'),[
    {i:'user',vnd:1,l:'Chi phí / lượt học',v:fmtVND(B.perRec),n:`thực tế / ${fmtN(B.nRec)} lượt tham dự`,d:D('perRec','pct',-1)},
    {i:'clock',vnd:1,l:'Chi phí / giờ đào tạo',v:fmtVND(B.perH),n:'thực tế / giờ đào tạo',d:D('perH','pct',-1)},
  ]);
  const mon=(F2,f)=>MONTHS.map((_,m)=>F2.cls.filter(c=>c.d0.getMonth()===m).reduce((a,c)=>a+f(c),0));
  const monP=F2=>F2.plan.length?MONTHS.map((_,m)=>F2.plan.filter(p=>p.s&&p.s.getMonth()===m).reduce((a,p)=>a+p.bud,0)):mon(F2,c=>c.cp);
  const short=v=>v>=1e9?nf1.format(v/1e9)+' tỷ':v>=1e5?nf1.format(v/1e6):fmtN(v);
  const ser=Q?[{name:'Thực tế '+Yc,color:'var(--navy)',values:mon(C.full,c=>c.ca||0)},{name:'Thực tế '+F.Y,color:'var(--accent)',values:mon(F,c=>c.ca||0)}]
             :[{name:'Kế hoạch',color:'var(--navy)',values:monP(F)},{name:'Thực tế',color:'var(--accent)',values:mon(F,c=>c.ca||0)}];
  const naB=F.plan.filter(p=>!p.s).reduce((a,p)=>a+p.bud,0);
  $('#bTitle').textContent=Q?`Chi phí thực tế theo tháng: ${F.Y} và ${Yc}`:'Bảng theo dõi chi phí theo tháng';
  legend($('#lgBud'),ser.map(s=>[s.name,s.color]));
  vbars($('#cBud'),MONTHS,ser,fmtVND,{label:short});
  const nh=[...new Set([...M.crs.values()].map(c=>c.nhom))];
  const plN=n=>B.hasP?F.plan.filter(p=>p.nhom===n).reduce((a,p)=>a+p.bud,0):F.cls.filter(c=>c.nhom===n).reduce((a,c)=>a+c.cp,0);
  hbars($('#cBudNhom'),nh.map(n=>({label:n,value:F.cls.filter(c=>c.nhom===n).reduce((a,c)=>a+(c.ca||0),0),sub:'Kế hoạch: '+fmtVND(plN(n))})).sort((a,b)=>b.value-a.value),fmtVND,{color:'var(--accent)'});
  // bảng: chương trình kế hoạch chưa hoàn thành + các lớp; Kế hoạch trước → Hoàn thành → Hủy; ngày mới → cũ
  const rows=F.plan.map(p=>({p,k:pst(p,F.T)})).filter(x=>x.k!=='done').map(({p,k})=>({src:LO(p.sec),id:p.code,name:p.name,d:p.s,dl:p.s?(dmy(p.s).slice(0,5)+' – '+dmy(p.e)):'Khi phát sinh',st:PST[k][0],nb:p.nb,cp:p.bud,ca:null,done:null,g:k==='cx'?2:0}))
    .concat(F.cls.map(c=>({src:typeOf(c),id:c.id,name:c.name,d:c.d0,dl:dmy(c.d0),st:c.st,nb:c.nb,cp:c.cp,ca:c.ca,done:c.done,g:c.st==='Hoàn thành'?1:c.st==='Hủy'?2:0})))
    .sort((a,b)=>a.g-b.g||((b.d?+b.d:-1)-(a.d?+a.d:-1)));
  const SB=(R,f)=>R.reduce((x,r)=>x+(f(r)||0),0); const vn=v=>v==null?'–':fmtN(v);
  if($('#tBud')) tbl($('#tBud'),[
    {h:'Mã lớp',v:r=>r.id||'–',sv:r=>r.id||null,tot:R=>`Tổng · ${fmtN(R.length)} dòng`},
    {h:'Khóa học / chương trình',v:r=>r.name,sv:r=>r.name},
    {h:'Phân loại KH',v:r=>r.src,sv:r=>r.src},
    {h:'Ngày',v:r=>r.dl,sv:r=>r.d||null,desc:1},
    {h:'Trạng thái',v:r=>pill(r.st),sv:r=>r.st},
    {h:'Nội bộ / ngoài',v:r=>r.nb,sv:r=>r.nb},
    {h:'Kế hoạch',num:1,v:r=>r.cp?fmtN(r.cp):'–',sv:r=>r.cp||null,tot:R=>fmtN(SB(R,r=>r.cp))},
    {h:'Thực tế',num:1,v:r=>vn(r.ca),sv:r=>r.ca,tot:R=>fmtN(SB(R,r=>r.ca))},
    {h:'Chênh lệch',num:1,v:r=>r.ca==null?'–':fmtN(r.ca-r.cp),sv:r=>r.ca==null?null:r.ca-r.cp,tot:R=>fmtN(SB(R,r=>r.ca)-SB(R,r=>r.cp))},
    {h:'Lượt HT',num:1,v:r=>vn(r.done),sv:r=>r.done,tot:R=>fmtN(SB(R,r=>r.done))}],rows,'Chưa có dữ liệu',{sort:1,total:1});
}
// trạng thái 1 dòng kế hoạch tại ngày T: cột Trạng thái nhập tay trước, để trống thì tự tính theo ngày
const PST={done:['Hoàn thành','var(--good)'],run:['Đang thực hiện','var(--warn)'],plan:['Kế hoạch','#b8bfcc'],late:['Trễ hạn','var(--crit)'],cx:['Hủy','#6b7385'],na:['Khi phát sinh','#d5dae5']};
function pst(p,T){if(p.st==='hoan thanh') return 'done'; if(p.st==='huy') return 'cx'; if(p.st==='dang thuc hien') return 'run'; if(!p.s) return 'na'; if(p.e<T) return 'late'; if(p.s<=T) return 'run'; return 'plan';}
function schM(F){const T=F.cut||F.T;
  if(F.plan.length){const its=F.plan.map(p=>({p,k:pst(p,T)})); const cnt=k=>its.filter(x=>x.k===k).length;
    const due=its.filter(x=>x.k!=='na'), dueDone=due.filter(x=>x.k==='done');  // lớp KH cả năm (có lịch) và lớp KH đã diễn ra
    const adhoc=F.cls.filter(c=>!c.plan&&c.st==='Hoàn thành'); const nAd=new Set(adhoc.map(c=>c.kh)).size; const num=dueDone.length+nAd;  // khóa ad-hoc (ngoài KH) đã diễn ra cũng tính vào tử số
    // tử số = khóa trong KH đã hoàn thành + khóa Phát sinh (ngoài KH) đã diễn ra; mẫu số = mọi khóa trong sheet Kế hoạch
    const typeOf=typeFn(F); const dn=F.cls.filter(c=>c.st==='Hoàn thành'&&c.d0<=T);
    const ORD={'Bắt buộc':0,'Theo nhu cầu':1,'Dự phòng':2,'Phát sinh':3};
    const grp=new Map(); [...new Set(F.plan.map(p=>LO(p.sec))),'Phát sinh'].sort((a,b)=>(ORD[a]??9)-(ORD[b]??9)).forEach(k=>grp.set(k,{plan:0,done:0,late:0,run:0,rem:0,cls:0}));
    const G=k=>{if(!grp.has(k)) grp.set(k,{plan:0,done:0,late:0,run:0,rem:0,cls:0}); return grp.get(k);};
    its.forEach(x=>{const g=G(LO(x.p.sec)); g.plan++; if(x.k==='done') g.done++; else g.rem++; if(x.k==='late') g.late++; if(x.k==='run') g.run++;});
    const psC=new Set(dn.filter(c=>!c.plan).map(c=>c.kh)); G('Phát sinh').done=psC.size;
    dn.forEach(c=>G(typeOf(c)).cls++);
    const allNum=dueDone.length+psC.size, den=F.plan.length;
    return {byPlan:true,unit:'lớp',its,due,dueDone,adhoc,nAd,num:allNum,den,grp,nPlan:its.filter(x=>x.k!=='na').length,nProg:new Set(F.plan.map(p=>p.stt||p.name)).size,rate:den?allNum/den:null,nRun:cnt('run'),nLate:cnt('late'),nUp:cnt('plan'),nNa:cnt('na'),nCx:cnt('cx'),nDone:F.cls.filter(c=>c.st==='Hoàn thành').length,nOut:F.cls.filter(c=>c.st==='Hoàn thành'&&!c.plan).length};}
  const planned=F.cls.filter(c=>c.plan); const due=planned.filter(c=>c.d0<=F.T); const dueDone=due.filter(c=>c.st==='Hoàn thành');
  return {unit:'lớp',planned,due,dueDone,nPlan:planned.length,rate:due.length?dueDone.length/due.length:null,nDone:F.cls.filter(c=>c.st==='Hoàn thành').length,nOut:F.cls.filter(c=>c.st==='Hoàn thành'&&!c.plan).length,nCx:F.cls.filter(c=>c.st==='Hủy'||c.st==='Dời lịch').length};}
function renderSch(F){
  const Sm=schM(F); const {due,dueDone}=Sm; const C=cmpF(); const Q=C?schM(C):null; const Qf=C?schM(C.full):null; const Yc=C&&C.Y; const Dl=(k,kind,g)=>Q?dlt(Sm[k],Q[k],(C.cut?'cùng kỳ ':'')+Yc,kind,g):null;
  const upS=F.clsAll.filter(c=>c.d0&&c.d0>=F.T&&['Kế hoạch','Đã xác nhận','Dời lịch'].includes(c.st)&&(F.K==='*'||c.khoi===F.K)&&(F.N==='*'||c.nhom===F.N)).sort((a,b)=>a.d0-b.d0);
  if(Sm.byPlan){
    const T0=F.T, inM=x=>x.p.s&&x.p.e&&x.p.s.getMonth()===T0.getMonth()&&x.p.e.getMonth()===T0.getMonth()&&x.p.s.getFullYear()===T0.getFullYear();
    const runs=Sm.its.filter(x=>x.k==='run'), rM=runs.filter(inM).length;
    const dn=F.cls.filter(c=>c.st==='Hoàn thành');
    kpis($('#kSch'),[
      {i:'cal',l:'Tỷ lệ hoàn thành kế hoạch',v:fmtN(Sm.num),of:fmtN(Sm.den),u:`khóa (${fmtP(Sm.rate)})`,p:Sm.rate,bl:[...Sm.grp].map(([k,g])=>[k,k==='Phát sinh'?`${fmtN(g.done)} khóa`:`${fmtN(g.done)}/${fmtN(g.plan)} khóa`])},
      {i:'book',u:'khóa',l:'Số khóa theo kế hoạch',v:fmtN(Sm.den),bl:[...Sm.grp].filter(([k])=>k!=='Phát sinh').map(([k,g])=>[k,`${fmtN(g.plan)} khóa`])},
      {i:'clock',u:'khóa',l:'Số khóa đang thực hiện',v:fmtN(Sm.nRun),bl:[[`Trong tháng ${T0.getMonth()+1}`,`${fmtN(rM)} khóa`],['Định kỳ cả năm',`${fmtN(runs.length-rM)} khóa`]],n:runs.length-rM?'Định kỳ: '+[...new Set(runs.filter(x=>!inM(x)).map(x=>x.p.name.replace(/\s*–\s*Khối.*$/,'').replace(/^Hướng Dẫn\s+/,'').replace(/\s+Cho\s+/,' ').replace(/Nhân Viên/g,'NV')))].join('; '):''},
      {i:'warn',u:'khóa',l:'Số khóa trễ hạn',v:fmtN(Sm.nLate),bl:[...Sm.grp].filter(([k])=>k!=='Phát sinh').map(([k,g])=>[k,`${fmtN(g.late)} khóa`])},
      {i:'cal',u:'khóa',l:'Số khóa còn lại theo kế hoạch',v:fmtN(Sm.den-Sm.dueDone.length),bl:[...Sm.grp].filter(([k])=>k!=='Phát sinh').map(([k,g])=>[k,`${fmtN(g.rem)} khóa`])},
      {i:'check',u:'lớp',l:'Số lớp đã tổ chức',v:fmtN(Sm.nDone),bl:[...Sm.grp].map(([k,g])=>[k,`${fmtN(g.cls)} lớp`])},
    ]);
    renderGantt(F,Sm);
    // theo tháng (tháng của Kết thúc dự kiến)
    const D=MONTHS.map((_,m)=>due.filter(x=>(x.p.s||x.p.dn).getMonth()===m).length), H=MONTHS.map((_,m)=>dueDone.filter(x=>(x.p.s||x.p.dn).getMonth()===m).length+new Set(Sm.adhoc.filter(c=>c.d0.getMonth()===m).map(c=>c.kh)).size);
    {const t=$('#sTitle'); t.replaceChildren(document.createTextNode('Số Khóa Học Theo Tháng ')); const bd=el('span',{class:'tbadge'}); bd.append(el('b',{text:`${fmtN(Sm.num)}/${fmtN(Sm.den)}`}),document.createTextNode(` khóa đã diễn ra / kế hoạch · ${fmtP(Sm.rate)}`)); t.append(bd);}
    legend($('#lgSch'),[['Khóa kế hoạch','var(--navy)'],['Đã diễn ra','var(--good)']]);
    $('#uTitle').textContent='Lịch Đào Tạo';
    const ord={'Trễ hạn':0,'Đang thực hiện':1,'Kế hoạch':2};
    const up=Sm.its.filter(x=>x.k==='run'||x.k==='plan'||x.k==='late').map(x=>({d:x.p.s,t:dmy(x.p.s).slice(0,5)+' – '+dmy(x.p.e).slice(0,5),name:x.p.name,who:x.p.dv,st:PST[x.k][0]}))
      .concat(upS.map(c=>({d:c.d0,t:dmy(c.d0),name:c.name+' (lớp)',who:c.gv,st:c.st}))).sort((a,b)=>(ord[a.st]??3)-(ord[b.st]??3)||a.d-b.d);
    tbl($('#tUp'),[{h:'Thời gian',v:r=>r.t,sv:r=>r.d},{h:'Nội dung',v:r=>r.name,sv:r=>r.name},{h:'Đơn vị thực hiện',v:r=>r.who||'–',sv:r=>r.who||null},{h:'Trạng thái',v:r=>pill(r.st),sv:r=>ord[r.st]??3}],up.slice(0,20),'Không có lớp sắp tới',{sort:1});
    const ch=$('#tUp').closest('.card').offsetHeight; const hh=ch?Math.max(240,ch-$('#lgSch').offsetHeight-$('#sTitle').closest('.ch').offsetHeight-60):240;
    vbars($('#cSch'),MONTHS,[{name:'Khóa kế hoạch',color:'var(--navy)',values:D},{name:'Đã diễn ra',color:'var(--good)',values:H}],fmtN,{intTicks:true,label:fmtN,h:hh});
  } else {
    const over=due.filter(c=>!['Hoàn thành','Hủy'].includes(c.st));
    kpis($('#kSch'),[
      {i:'cal',l:'Tỷ lệ hoàn thành kế hoạch',v:fmtN(dueDone.length),of:fmtN(due.length),u:'lớp',p:Sm.rate,n:`${fmtP(Sm.rate)} lớp trong KH đã đến hạn · HR Action Plan 3.3`,d:Dl('rate','pp',1)},
      {i:'book',u:'lớp',l:'Số lớp theo kế hoạch',v:fmtN(Sm.nPlan),n:'cột Kế hoạch = Trong kế hoạch',d:Qf?dlt(Sm.nPlan,Qf.nPlan,'cả năm '+Yc,'pct',0):null},
      {i:'check',u:'lớp',l:'Số lớp đã tổ chức',v:fmtN(Sm.nDone),n:`gồm ${fmtN(Sm.nOut)} lớp ngoài kế hoạch`,d:Dl('nDone','pct',1)},
      {i:'x',u:'lớp',l:'Số lớp hủy / dời lịch',v:fmtN(Sm.nCx),n:'ghi lý do ở cột Ghi chú',d:Dl('nCx','pct',-1)},
      {i:'warn',u:'lớp',l:'Đến hạn chưa cập nhật',v:fmtN(over.length),st:over.length?['warn','Cần cập nhật trạng thái']:['good','Đã cập nhật đủ'],n:'ngày bắt đầu đã qua, chưa Hoàn thành/Hủy'},
      {i:'cal',u:'lớp',l:'Số lớp sắp tới',v:fmtN(upS.length),n:'từ ngày tính'},
    ]);
    $('#calLg').innerHTML='<span><i class="mk done"></i>Đã tổ chức</span><span><i class="mk up"></i>Dự kiến (chưa đến hạn)</span><span><i class="mk late"></i>Quá hạn, chưa cập nhật</span><span><i class="mk dl"></i>Dời lịch</span><span><i class="mk cx"></i>Hủy</span>';
    $('#calSch').classList.remove('gantt'); renderCal(F);
    const D=MONTHS.map((_,m)=>due.filter(c=>c.d0.getMonth()===m).length), H=MONTHS.map((_,m)=>dueDone.filter(c=>c.d0.getMonth()===m).length);
    $('#sTitle').textContent='Lớp trong kế hoạch: đến hạn và hoàn thành';
    legend($('#lgSch'),[['Đến hạn','var(--navy)'],['Hoàn thành','var(--accent)']]);
    vbars($('#cSch'),MONTHS,[{name:'Đến hạn',color:'var(--navy)',values:D},{name:'Hoàn thành',color:'var(--accent)',values:H}],fmtN,{intTicks:true,label:fmtN});
    $('#uTitle').textContent='Lớp sắp tới';
    tbl($('#tUp'),[{h:'Ngày',v:c=>dmy(c.d0)},{h:'Mã lớp',v:c=>c.id},{h:'Khóa học',v:c=>c.name},{h:'Khối',v:c=>c.khoi},{h:'Trạng thái',v:c=>pill(c.st)},{h:'Đăng ký',num:1,v:c=>fmtN(c.reg),sv:c=>c.reg}],upS.slice(0,15),'Không có lớp sắp tới',{sort:1});
  }
  const st=$('#fStatus').value||'Tất cả';
  tbl($('#tCls'),[{h:'Ngày',v:c=>dmy(c.d0),sv:c=>c.d0,desc:1},{h:'Mã lớp',v:c=>c.id},{h:'Khóa học',v:c=>c.name},{h:'Nhóm',v:c=>c.nhom},{h:'KH',v:c=>c.plan?'Trong':'Ngoài'},{h:'Khối',v:c=>c.khoi||null},{h:'Giảng viên',v:c=>c.gv||null},{h:'Trạng thái',v:c=>pill(c.st),sv:c=>c.st},{h:'Đăng ký',num:1,v:c=>fmtN(c.reg),sv:c=>c.reg},{h:'Hoàn thành',num:1,v:c=>fmtN(c.done),sv:c=>c.done},{h:'Ghi chú',v:c=>c.note||null}],
    F.cls.filter(c=>st==='Tất cả'||c.st===st).sort((a,b)=>b.d0-a.d0),'Chưa có dữ liệu',{sort:1});
}
// sắp xếp Gantt: bấm tiêu đề cột, các dòng được xếp lại BÊN TRONG từng nhóm I–IV
const GSS={i:-1,d:1};
function ganttSort(t){const tb=t.tBodies[0]; const ths=[...t.tHead.rows[0].cells];
  const val=(tr,i)=>{const td=tr.cells[i]; let x=(td.dataset.fv!=null?td.dataset.fv:td.textContent).trim(); if(!x||x==='–') return null;
    if(i===0) return tr._g; if(i===3||i===4){const m=x.match(/^(\d{2})\/(\d{2})/); return m?+m[2]*100+ +m[1]:null;}
    if(i===6) return td.dataset.sv===''||td.dataset.sv==null?null:+td.dataset.sv; if(i===5) return {'Trễ hạn':0,'Đang thực hiện':1,'Kế hoạch':2,'Hoàn thành':3,'Khi phát sinh':4,'Hủy':5}[x]??6; return x;};
  const groups=[]; [...tb.rows].forEach(r=>{if(r.classList.contains('sec')) groups.push({h:r,rows:[]}); else if(groups.length) groups[groups.length-1].rows.push(r); else groups.push({h:null,rows:[r]});});
  let gi=0; groups.forEach(g=>g.rows.forEach((r,j)=>{r._o=j; r._g=gi++;}));
  const draw=()=>{ths.forEach((h2,i)=>h2.dataset.sort=i===GSS.i?(GSS.d>0?'asc':'desc'):'');
    groups.forEach(g=>{const rs=g.rows.slice().sort((a,b)=>{if(GSS.i<0) return a._o-b._o; const x=val(a,GSS.i),y=val(b,GSS.i); if(x==null||y==null) return x==null&&y==null?a._o-b._o:x==null?1:-1;
      return (typeof x==='number'?x-y:String(x).localeCompare(String(y),'vi',{numeric:true}))*GSS.d||a._o-b._o;}); let after=g.h; rs.forEach(r=>{if(after) after.after(r); else tb.prepend(r); after=r;});});
    let n=0; [...tb.rows].forEach(r=>{if(!r.classList.contains('sec')) r.cells[0].textContent=++n;});};
  ths.forEach((h2,i)=>{if(i>6) return; h2.classList.add('sortable'); h2.tabIndex=0; h2.title='Bấm để sắp xếp (trong từng nhóm)';
    const go=()=>{if(GSS.i===i) GSS.d=-GSS.d; else {GSS.i=i; GSS.d=1;} draw();}; h2.addEventListener('click',go); h2.addEventListener('keydown',e=>{if(e.key==='Enter'||e.key===' '){e.preventDefault(); go();}});});
  draw();}
// ---------- Gantt: mỗi dòng kế hoạch = 1 thanh trên trục 12 tháng
function renderGantt(F,Sm){
  const host=$('#calSch'); host.classList.remove('cal'); host.classList.add('gantt'); host.replaceChildren();
  $('#calTitle').textContent=`Bảng theo dõi tiến độ Đào Tạo ${F.Y}`;
  $('#calLg').innerHTML=['done','run','plan','late'].map(k=>`<span><i class="gsw" style="background:${PST[k][1]}"></i>${PST[k][0]}</span>`).join('')+'<span><i class="gdm"></i>Ngày thực hiện (lớp đã tổ chức)</span><span><i class="gnw"></i>Ngày tính ('+dmy(F.T).slice(0,5)+')</span>';
  const y0=new Date(F.Y,0,1), y1=new Date(F.Y+1,0,1); const pos=d=>Math.max(0,Math.min(100,(d-y0)/(y1-y0)*100)); const DAY=864e5;
  const showNow=F.T>=y0&&F.T<y1, nowP=pos(F.T); const curM=showNow?F.T.getMonth():-1;
  const t=el('table'); const th=el('thead'), hr=el('tr');
  [['STT','c-stt'],['Nội dung đào tạo','c-nm'],['Đơn vị thực hiện','c-dv'],['Bắt đầu','c-d'],['Kết thúc','c-d'],['Trạng thái','c-st'],['Chi phí','c-cp']].forEach(([x,c])=>{const th=el('th',{class:c,text:x}); if(c==='c-cp') th.append(el('small',{class:'thu',text:'(Triệu Đ)'})); hr.append(th);});
  const tlh=el('th',{class:'c-tl'}); const mh=el('div',{class:'gmh'}); MONTHS.forEach((m,i)=>mh.append(el('span',{class:i===curM?'now':'',text:m}))); tlh.append(mh); hr.append(tlh); th.append(hr); t.append(th);
  const tb=el('tbody');
  const tl=(s,e,k,extra)=>{const d=el('div',{class:'gtl'});
    if(showNow){const n=el('i',{class:'gnow'}); n.style.left=nowP+'%'; d.append(n);}
    if(!s){if(extra&&extra.length) extra.forEach(x=>{const m=el('i',{class:x.cls}); m.style.left=pos(x.d)+'%'; m.title=x.t; d.append(m);}); else d.append(el('span',{class:'gna',text:'Khi có phát sinh'})); return d;}
    const b=el('div',{class:'gbar '+k}); const l=pos(s), r=pos(new Date(+e+DAY)); b.style.left=l+'%'; b.style.width=Math.max(2,r-l)+'%'; b.style.background=PST[k][1]; d.append(b);
    (extra||[]).forEach(x=>{const m=el('i',{class:x.cls}); m.style.left=pos(x.d)+'%'; m.title=x.t; d.append(m);}); return d;};
  const tr1=v=>v==null?'–':nf1.format(v/1e6);
  const cost=(act,bud)=>{const w=el('div',{class:'gcp'}); const t=el('div',{class:'gcpt'}); t.append(el('b',{text:act?tr1(act):'–'}),document.createTextNode(' / '+(bud?tr1(bud):'–'))); w.append(t);
    if(bud||act){const k=el('span',{class:'gcpb'}); const f=el('i'); f.style.width=(bud?Math.min(100,(act||0)/bud*100):100)+'%'; if(!bud||(act||0)>bud) f.classList.add('over'); k.append(f); w.append(k);}
    w.title=`Thực tế ${act?fmtN(act)+' đ':'–'} / ngân sách ${bud?fmtN(bud)+' đ':'–'}`; return w;};
  const row=(cells,tlNode,title)=>{const tr=el('tr'); cells.forEach(([v,c,fv,sv])=>{const td=el('td',{class:c}); if(fv!=null) td.dataset.fv=fv; if(sv!==undefined) td.dataset.sv=sv==null?'':sv; if(v instanceof Node) td.append(v); else if(c==='c-dv'||c==='c-nt'){const x=el('div',{class:'clamp',text:v}); td.title=v; td.append(x);} else td.textContent=v; tr.append(td);}); const td=el('td',{class:'c-tl'}); td.append(tlNode); if(title) td.title=title; tr.append(td); tb.append(tr);};
  const secs=[...new Set(Sm.its.map(x=>x.p.sec))].sort((a,b)=>a.localeCompare(b,'vi',{numeric:true}));
  const sk=x=>[parseFloat(x.p.stt)||1e9,x.p.s?+x.p.s:Infinity]; const typeOfG=typeFn(F);
  const secHead=(name,bars)=>{const w=el('div',{class:'gsh'}); w.append(el('b',{text:name}));
    bars.filter(Boolean).forEach(([a,b2,txt,money])=>{const x=el('span',{class:'gsb'+(money?' m':'')}); x.append(el('small',{text:money?'Chi phí':'Tiến độ'})); const tk=el('span',{class:'gsbt'}); const f=el('i'); f.style.width=(b2?Math.min(100,a/b2*100):a?100:0)+'%'; if(money&&b2&&a>b2) f.style.background='var(--crit)'; tk.append(f); x.append(tk,el('em',{text:txt})); w.append(x);}); return w;};
  secs.forEach(sec=>{const its=Sm.its.filter(x=>x.p.sec===sec).sort((a,b)=>{const A=sk(a),B=sk(b); return A[0]-B[0]||A[1]-B[1];});
    const bud=its.reduce((a,x)=>a+x.p.bud,0); const done=its.filter(x=>x.k==='done').length;
    const spent=F.cls.filter(c=>typeOfG(c)===LO(sec)).reduce((a,c)=>a+(c.ca||0),0);
    const tr=el('tr',{class:'sec'}); const td=el('td'); td.colSpan=8; const hd=done; td.append(secHead(sec,[[hd,its.length,`${fmtN(hd)}/${fmtN(its.length)} khóa đã diễn ra`],bud||spent?[spent,bud,`${vt(spent)} / ${vt(bud)}`+'',1]:null])); tr.append(td); tb.append(tr);
    its.forEach(({p,k})=>{
      const nm=el('div'); nm.append(el('span',{text:p.name})); 
      const ex=[]; if(p.dn) ex.push({d:p.dn,cls:'gdm',t:'Thực hiện '+dmy(p.dn)});
      if(p.code&&p.s) F.clsAll.filter(c=>c.kh===p.code&&c.st==='Hoàn thành'&&c.d0>=p.s&&c.d0<=p.e&&(!p.dn||+c.d0!==+p.dn)).forEach(c=>ex.push({d:c.d0,cls:'gdm',t:`Lớp ${c.id} ${dmy(c.d0)}${c.done?' · '+c.done+' học viên':''}`}));
      row([[p.stt||'',''],[nm,'c-nm',p.name],[p.dv||'–','c-dv'],[p.s?dmy(p.s).slice(0,5):'–','c-d'],[p.e?dmy(p.e).slice(0,5):'–','c-d'],[pill(PST[k][0]),'c-st'],(()=>{const a=F.clsAll.filter(c=>c.kh===p.code&&planRowOf(F,c)===p).reduce((x,c)=>x+(c.ca||0),0); return [cost(a,p.bud),'c-cp',null,a||null];})()],
        tl(p.s||(k==='done'?p.dn:null),p.e||(k==='done'?p.dn:null),k,ex),`${p.name}\n${p.s?dmy(p.s)+' – '+dmy(p.e):'Khi có phát sinh'} · ${PST[k][0]}${p.dn?' · thực hiện '+dmy(p.dn):''}${p.bud?' · ngân sách '+fmtN(p.bud)+' đ':''}`);});
  });
  // lớp ngoài kế hoạch (phát sinh) đã/sắp tổ chức trong năm
  const out=F.cls.filter(c=>!c.plan).sort((a,b)=>a.d0-b.d0);
  if(out.length){const oc=out.reduce((a,c)=>a+(c.ca||0),0), nk=new Set(out.map(c=>c.kh)).size, nd=new Set(out.filter(c=>c.st==='Hoàn thành').map(c=>c.kh)).size;
    {const tr=el('tr',{class:'sec'}); const td=el('td'); td.colSpan=8; const sh=secHead('IV. Phát sinh (ngoài kế hoạch)',[[nd,nk,`${fmtN(nd)} khóa đã diễn ra`]]); sh.append(el('span',{class:'gsx',text:`Chi phí ${vt(oc)} (chưa có trong ngân sách kế hoạch)`})); td.append(sh); tr.append(td); tb.append(tr);}
    out.forEach(c=>{const k=c.st==='Hoàn thành'?'done':c.st==='Hủy'?'cx':c.d0<F.T?'late':'plan'; const e=c.d1&&c.d1>=c.d0?c.d1:c.d0;
      const nm=el('div'); nm.append(el('span',{text:c.name})); 
      row([['–',''],[nm,'c-nm',c.name],[c.gv||'–','c-dv'],[dmy(c.d0).slice(0,5),'c-d'],[dmy(e).slice(0,5),'c-d'],[pill(k==='late'?'Trễ hạn':c.st),'c-st'],[cost(c.ca,null),'c-cp',null,c.ca||null]],tl(c.d0,e,k),`${c.name} · ${dmy(c.d0)} · ${c.st}`);});}
  t.append(tb); host.append(t); if(host.nextElementSibling&&host.nextElementSibling.classList.contains('fnote')) host.nextElementSibling.remove(); ganttSort(t);
}

// Chi phí đào tạo các năm: mỗi năm 1 thanh tiến độ thực tế / ngân sách kế hoạch
function yProg(rows,sel,T){const box=el('div',{class:'ypg'});
  rows.slice().sort((a,b)=>b.y-a.y).forEach(r=>{const B=r.b; const u=B.plan?B.act/B.plan:null; const over=B.plan&&B.act>B.plan;
    const row=el('div',{class:'ypr'+(r.y===sel.y?' cur':'')}); row.append(el('div',{class:'yl',text:String(r.y)+(r.y===T.getFullYear()?'*':'')}));
    const tr=el('div',{class:'yt'}); const f=el('i'); f.style.width=(u==null?0:Math.min(100,u*100))+'%'; if(over) f.style.background='var(--crit)'; tr.append(f);
    tr.append(el('span',{class:'yp'+((u||0)<0.12?' out':''),text:u==null?(B.act?'chưa có ngân sách KH':'chưa có số liệu'):fmtP(u)}));
    const y0=new Date(r.y,0,1),y1=new Date(r.y+1,0,1); if(T>y0&&T<y1){const n=el('b',{class:'ynow'}); n.style.left=((T-y0)/(y1-y0)*100)+'%'; n.title='Ngày tính '+dmy(T)+' · tiến độ năm '+fmtP((T-y0)/(y1-y0)); tr.append(n);}
    const v=el('div',{class:'yv'}); v.append(el('b',{text:fmtVND(B.act)}),document.createTextNode(B.plan?' / '+fmtVND(B.plan):''));
    row.append(tr,v); row.title=`${r.y}: chi phí ${fmtN(B.act)} đ`+(B.plan?` / ngân sách ${fmtN(B.plan)} đ (${fmtP(u)})`:' · chưa có ngân sách kế hoạch'); box.append(row);});
  return box;}
function renderYears(F){
  const ys=YEARS.slice().sort((a,b)=>a-b); const T=F.T; const curY=T.getFullYear();
  const rows=ys.map(y=>{const Fy=filt(y); const p=perfM(Fy), b=budM(Fy), sc=schM(Fy); return {y,p,b,sc,n:sc.nDone,Fy};});
  const sel=rows.find(r=>r.y===F.Y)||rows[rows.length-1];
  const prevY=ys.filter(y=>y<sel.y).pop();
  let prev=null, lab='';
  if(prevY!=null){const cut=sel.y===curY?new Date(prevY,T.getMonth(),Math.min(T.getDate(),new Date(prevY,T.getMonth()+1,0).getDate())):null; const Fp=filt(prevY,cut); prev={p:perfM(Fp),b:budM(Fp),sc:schM(Fp)}; lab=(cut?'cùng kỳ ':'')+prevY;}
  const M_=[
    {i:'users',t:'Số NV được đào tạo',g:r=>r.p.cov,f:fmtP,big:r=>[fmtN(r.p.uniq),'/'+fmtN(r.p.hc)+' NV'],tip:r=>`${fmtP(r.p.cov)} · ${fmtN(r.p.uniq)}/${fmtN(r.p.hc)} NV`,d:['cov','pp',1,'p']},
    {i:'check',t:'Số lượt tham dự',g:r=>r.p.attN,f:fmtN,big:r=>[fmtN(r.p.attN),'lượt'],tip:r=>`${fmtN(r.p.attN)} lượt · ${fmtN(r.p.learners)} người`,d:['attN','pct',1,'p']},
    {i:'clock',ok:r=>r.p.hasH,t:'Tổng số giờ đào tạo',g:r=>r.p.hrs,f:fmtN,big:r=>[fmtN(r.p.hrs),'giờ'],d:['hrs','pct',1,'p']},
    {i:'gauge',ok:r=>r.p.hasH,t:'Số giờ đào tạo / NV',g:r=>r.p.hpn,f:fmt1,big:r=>[fmt1(r.p.hpn),`giờ / ${fmtN(r.p.hc)} NV`],tip:r=>`${fmt1(r.p.hpn)} giờ/NV · ${fmtN(r.p.hrs)} giờ ÷ ${fmtN(r.p.hc)} NV`,d:['hpn','pct',1,'p']},
    {i:'check',t:'Số lớp đã tổ chức',g:r=>r.n,f:fmtN,big:r=>[fmtN(r.n),'lớp'],int:1,d:['nDone','pct',1,'sc']},
    {i:'star',ok:r=>r.p.satN>0,t:'Mức độ hài lòng',g:r=>r.p.satN?r.p.sat:null,f:fmt1,big:r=>[r.p.satN?fmt1(r.p.sat):'–','/ 5'],max5:1,d:['sat','abs',1,'p']},
    {i:'coin',prog:1,ok:r=>r.b.act>0||r.b.plan>0,t:'Tổng chi phí đào tạo',g:r=>r.b.act,f:fmtVND,lf:v=>nf1.format(v/1e6),big:r=>{const t=fmtVND(r.b.act).split(' '); return [t[0],t[1]==='tỷ'?'tỷ đ':t[1]==='tr'?'triệu đ':'đ'];},d:['act','pct',0,'b']},
  ];
  const host=$('#yCharts'); host.replaceChildren(); host.classList.toggle('wide',rows.length>8);
  const miss=M_.filter(m=>m.ok&&!rows.some(m.ok)).map(m=>m.t);
  M_.filter(m=>!m.ok||rows.some(m.ok)).forEach(m=>{const c=el('div',{class:'yc'}); const hd=el('div',{class:'hd'}); const ic=el('span',{class:'ico'}); ic.append(icon(m.i)); hd.append(ic,el('h3',{text:m.t})); c.append(hd);
    const bg=m.big(sel); const cv=el('div',{class:'cv'}); cv.append(document.createTextNode(bg[0])); cv.append(el('small',{text:bg[1]})); c.append(cv);
    const sub=el('div',{class:'sub'}); sub.textContent=`Năm ${sel.y}`+(m.tip&&m.g(sel)!=null&&m.f===fmtP?' · '+fmtP(m.g(sel)):''); c.append(sub);
    if(prev){const [k,kind,g,src]=m.d; const dd=dlt(sel[src][k],prev[src][k],lab,kind,g); const w=el('div',{class:'dlw'}); w.style.marginTop='6px'; w.append(el('span',{class:'dl '+(dd.c||''),text:dd.t})); if(dd.e) w.append(el('span',{class:'dle',text:dd.e})); const box=el('div',{class:'kpi'}); box.style.cssText='padding:0;box-shadow:none;background:none'; box.append(w); c.append(box);}
    if(m.prog){c.classList.add('yprog'); host.append(c); c.append(yProg(rows,sel,T)); return;}
    const ch=el('div',{class:'chart'}); c.append(ch); host.append(c);
    const vals=rows.map(r=>{const v=m.g(r); return v==null||isNaN(v)?0:v;});
    const many=rows.length>12; vbars(ch,rows.map(r=>(many?"'"+String(r.y).slice(2):String(r.y))+(r.y===curY?'*':'')),[{name:m.t,values:vals,colors:rows.map(r=>r.y===sel.y?'var(--accent)':'var(--navy)'),tips:rows.map(r=>m.tip?m.tip(r):m.f(m.g(r)))}],m.f,{h:190,label:v=>m.lf?m.lf(v):m.f(v),tickFmt:m.max5?fmtN:m.f===fmtVND?(v=>fmtN(v/1e6)):m.f,small:many,intTicks:!!m.int,max:m.max5?5:m.f===fmtP?1:null,ticks:m.max5?5:4});
  });
  // Phân tích nhanh (so năm trước, cùng kỳ nếu năm chưa hết)
  {const out=[]; const add=(ic,kind,title,body)=>out.push({ic,kind,title,body}); const s0=sel, P1=s0.p, B1=s0.b; const pv=prev; const L=lab;
   const pc=(a,b)=>b?(a-b)/b:null; const sg=x=>x==null?'':(x>=0?'+':'−')+fmtP(Math.abs(x));
   if(P1.hc) add('users',pv&&P1.cov>=pv.p.cov?'good':'info',`Độ phủ ${fmtP(P1.cov)}`+(pv?` (${P1.cov>=pv.p.cov?'+':'−'}${nf1.format(Math.abs(P1.cov-pv.p.cov)*100)} điểm % so với ${L})`:''),
     `${fmtN(P1.uniq)}/${fmtN(P1.hc)} NV đang làm đã tham gia đào tạo`+(pv?`; ${L}: ${fmtP(pv.p.cov)}.`:'.'));
   if(pv&&P1.hasH&&pv.p.hasH){const g=dn=>{const m=new Map(); dn.forEach(r=>m.set(r.cl.nhom,(m.get(r.cl.nhom)||0)+r.hrs)); return m;}; const a=g(P1.done), b0=g(pv.p.done);
     const ks=[...new Set([...a.keys(),...b0.keys()])].map(k=>({k,d:(a.get(k)||0)-(b0.get(k)||0)})).sort((x,y)=>y.d-x.d); const dH=P1.hrs-pv.p.hrs;
     add('clock',dH>=0?'good':'info',`Giờ học ${sg(pc(P1.hrs,pv.p.hrs))}, chủ yếu từ nhóm ${ks[0].k}`,
       `${fmtN(P1.hrs)} giờ so với ${fmtN(pv.p.hrs)} giờ (${L}); nhóm **${ks[0].k}** tăng ${fmtN(ks[0].d)} giờ`+(dH>0?` (${fmtP(ks[0].d/dH)} mức tăng)`:'')+(ks[ks.length-1].d<0?`, nhóm ${ks[ks.length-1].k} thấp hơn ${fmtN(-ks[ks.length-1].d)} giờ.`:'.'));}
   if(pv&&B1.act&&pv.b.act&&P1.hrs&&pv.p.hrs){const c1=B1.act/P1.hrs, c0=pv.b.act/pv.p.hrs;
     add('coin',c1<=c0?'good':'info',`Chi phí / giờ học ${c1<=c0?'giảm':'tăng'} ${fmtP(Math.abs(pc(c1,c0)))}`,`${vt(c1)}/giờ so với ${vt(c0)}/giờ (${L}): tổng chi phí ${sg(pc(B1.act,pv.b.act))} trong khi giờ học ${sg(pc(P1.hrs,pv.p.hrs))}, nhờ tỷ trọng đào tạo nội bộ cao hơn.`);}
   if(P1.hasH&&P1.hrs){const sh=p=>p.hrs?p.done.filter(r=>norm(r.cl.nb)==='noi bo').reduce((x,r)=>x+r.hrs,0)/p.hrs:null; const a=sh(P1), b0=pv&&pv.p.hasH?sh(pv.p):null;
     add('home','info',`Nội bộ chiếm ${fmtP(a)} số giờ`+(b0!=null?` (${L}: ${fmtP(b0)})`:''),b0!=null?`Cơ cấu đào tạo ${a>b0?'nghiêng nhiều hơn về nội bộ':'nghiêng nhiều hơn về bên ngoài'} so với năm trước.`:'');}
   if(P1.satN) add('star',pv&&pv.p.satN&&P1.sat>=pv.p.sat?'good':'info',`Hài lòng ${fmt1(P1.sat)}/5`,`${fmtN(P1.satN)} lượt đánh giá`+(pv&&pv.p.satN?`; ${L}: ${fmt1(pv.p.sat)}/5.`:'.'));
   else add('star','info','Chưa có dữ liệu đánh giá sau khóa',evalNote());
   insRender($('#yQuick'),out);
   $('#yqSub').textContent=pv?`${sel.y} so với ${L}`:`Năm ${sel.y}`;}
}
function renderData(){
  const host=$('#dq'); host.replaceChildren();
  const s=el('p',{class:'hint'}); s.textContent=`Đã đọc: ${fmtN(M.emp.size)} nhân viên · ${fmtN(M.crs.size)} khóa học · ${fmtN(M.cls.size)} lớp · ${fmtN(M.recs.length)} dòng Training Record hợp lệ`+(M.plan.length?` · ${fmtN(M.plan.length)} dòng Kế hoạch đào tạo.`:'. Chưa có sheet Ke hoach dao tao (ngân sách & tiến độ kế hoạch tính theo sheet Lop hoc).'); host.append(s);
  if(!M.badCls.length&&!M.badRec.length){host.append(el('p',{text:'Không phát hiện lỗi dữ liệu.'}));}
  if(M.badCls.length){host.append(el('h3',{text:`Sheet Lop hoc: ${M.badCls.length} lỗi`})); const u=el('ul'); M.badCls.slice(0,50).forEach(x=>u.append(el('li',{text:x}))); host.append(u);}
  if(M.badRec.length){host.append(el('h3',{text:`Sheet Training Record: ${M.badRec.length} dòng lỗi`})); const d=el('div',{class:'tbl'}); tbl(d,[{h:'Dòng dữ liệu thứ',num:1,v:r=>r.row},{h:'Mã lớp',v:r=>r.cid},{h:'Ngày học',v:r=>r.d||''},{h:'Mã NV',v:r=>r.eid},{h:'Lỗi',v:r=>r.msg}],M.badRec.slice(0,500)); host.append(d);}
  const fx=t=>el('span',{class:'fx',text:t});
  tbl($('#tDef'),[{h:'Chỉ số',v:r=>r[0]},{h:'Công thức',v:r=>fx(r[1])},{h:'Ý nghĩa',v:r=>r[2]}],[
    ['Độ phủ đào tạo','NV đang làm có ≥ 1 lượt tham dự ÷ Tổng NV đang làm × 100%','Bao nhiêu phần trăm nhân sự hiện có đã được đào tạo ít nhất 1 lần trong năm. NV đang làm = đã vào làm và chưa nghỉ tại ngày tính (năm cũ: 31/12).'],
    ['Số lượt học hoàn thành','Σ dòng Training Record có Trạng thái tham dự = Hoàn thành','Quy mô hoạt động đào tạo; 1 người học 3 lớp = 3 lượt.'],
    ['Tổng số giờ đào tạo','Σ Giờ học của các lượt tham dự (Giờ học thực tế → nếu trống: Số giờ lớp → nếu trống: Thời lượng chuẩn)','Tổng thời gian công ty đầu tư cho đào tạo.'],
    ['Số giờ đào tạo / NV','Tổng giờ đào tạo ÷ Tổng NV đang làm','Bình quân mỗi nhân viên được đào tạo bao nhiêu giờ; so được giữa các năm, Khối, Phòng ban dù quy mô khác nhau.'],
    ['Tỷ lệ tham dự','Hoàn thành ÷ (Hoàn thành + Vắng + Không hoàn thành) × 100%','Người được mời có thực sự đi học và học hết lớp không.'],
    ['Tỷ lệ đạt','Lượt Đạt ÷ (Lượt Đạt + Lượt Không đạt) × 100%  (Đạt = cột Kết quả là Đạt, hoặc Điểm sau ≥ Điểm đạt của khóa)','Chất lượng tiếp thu sau đào tạo; chỉ tính lượt có kết quả đánh giá.'],
    ['Mức độ hài lòng','Σ Điểm hài lòng ÷ Số lượt đánh giá  (thang 1–5)','Học viên đánh giá lớp học tốt đến mức nào.'],
    ['Tỷ lệ hoàn thành kế hoạch','(Khóa trong kế hoạch đã diễn ra + Khóa ad-hoc / ngoài kế hoạch đã diễn ra) ÷ Tổng số khóa theo kế hoạch cả năm × 100%. Khóa KH: mỗi dòng có lịch ở sheet Ke hoach dao tao (Trạng thái = Hoàn thành là đã diễn ra); khóa ad-hoc: lớp Ngoài kế hoạch, Hoàn thành ở sheet Lop hoc. Chương trình phát sinh chưa có lịch không tính vào mẫu số','Phòng có triển khai đúng kế hoạch đào tạo đã duyệt không; mục tiêu 100% (HR Action Plan 3.3).'],
    ['Trạng thái kế hoạch (Gantt)','Cột Trạng thái nhập tay (Hoàn thành / Đang thực hiện / Hủy) → nếu trống: Kết thúc dự kiến < ngày tính = Trễ hạn · Bắt đầu ≤ ngày tính ≤ Kết thúc = Đang thực hiện · chưa tới Bắt đầu = Kế hoạch','Xanh lá = đã hoàn thành · Vàng = đang thực hiện · Xám = kế hoạch · Đỏ = trễ hạn.'],
    ['Sử dụng ngân sách','Σ Chi phí thực tế (sheet Lop hoc) ÷ Σ Ngân sách dự kiến (sheet Ke hoach dao tao) × 100%','Đã tiêu bao nhiêu so với ngân sách đào tạo được duyệt; vạch đứt cho biết tiến độ thời gian của năm (tính đến ngày tính) để so sánh.'],
    ['Chi phí / lượt học','Chi phí thực tế ÷ Lượt học hoàn thành','Bình quân 1 lượt học tốn bao nhiêu tiền.'],
    ['Tỷ trọng chi phí nội bộ','Chi phí lớp Nội bộ ÷ Tổng chi phí thực tế × 100%','Mức tự chủ đào tạo bằng nguồn lực nội bộ so với thuê ngoài.'],
    ['Tăng / giảm so với năm khác','Chỉ số tỷ lệ: A − B (điểm %) · Chỉ số số lượng: (A − B) ÷ B × 100%','A = năm đang xem, B = năm so sánh. Năm đang xem chưa hết năm thì B lấy cùng kỳ (cùng khoảng 01/01 → ngày tính).'],
  ]);
}


// ---------- PDF export (không cần thư viện ngoài): chụp từng tab qua SVG foreignObject → JPEG → ghép PDF A4 ngang
const nextFrame=()=>new Promise(r=>requestAnimationFrame(()=>setTimeout(r,60)));
const latin=s=>{const u=new Uint8Array(s.length); for(let i=0;i<s.length;i++) u[i]=s.charCodeAt(i)&255; return u;};
function buildPdf(pages){ // pages: [{jpeg:Uint8Array,w,h,dw,dh}] trên khổ 842×595 pt
  const parts=[]; let off=0; const offs=[]; const push=u=>{parts.push(u); off+=u.length;};
  const obj=(id,chunks)=>{offs[id]=off; push(latin(id+' 0 obj\n')); chunks.forEach(push); push(latin('\nendobj\n'));};
  push(latin('%PDF-1.4\n%\xE2\xE3\xCF\xD3\n'));
  const n=pages.length; obj(1,[latin('<< /Type /Catalog /Pages 2 0 R >>')]);
  obj(2,[latin(`<< /Type /Pages /Kids [${pages.map((_,i)=>(3+i*3)+' 0 R').join(' ')}] /Count ${n} >>`)]);
  pages.forEach((p,i)=>{const pid=3+i*3; if(p.full){p.dw=842; p.dh=842*p.h/p.w;} const PH=p.full?p.dh:p.dh+40, OX=p.full?0:20;
    obj(pid,[latin(`<< /Type /Page /Parent 2 0 R /MediaBox [0 0 842 ${PH.toFixed(2)}] /Resources << /XObject << /Im${i} ${pid+2} 0 R >> >> /Contents ${pid+1} 0 R >>`)]);
    const cs=`0.933 0.945 0.965 rg 0 0 842 ${PH.toFixed(2)} re f q ${p.dw.toFixed(2)} 0 0 ${p.dh.toFixed(2)} ${OX} ${OX} cm /Im${i} Do Q`;
    obj(pid+1,[latin(`<< /Length ${cs.length} >>\nstream\n${cs}\nendstream`)]);
    obj(pid+2,[latin(`<< /Type /XObject /Subtype /Image /Width ${p.w} /Height ${p.h} /ColorSpace /DeviceRGB /BitsPerComponent 8 /Filter /DCTDecode /Length ${p.jpeg.length} >>\nstream\n`),p.jpeg,latin('\nendstream')]);});
  const xo=off; let x=`xref\n0 ${3+n*3}\n0000000000 65535 f \n`; for(let id=1;id<3+n*3;id++) x+=String(offs[id]).padStart(10,'0')+' 00000 n \n';
  push(latin(x+`trailer\n<< /Size ${3+n*3} /Root 1 0 R >>\nstartxref\n${xo}\n%%EOF`));
  return new Blob(parts,{type:'application/pdf'});
}
function jpegBytes(cv){const b=atob(cv.toDataURL('image/jpeg',0.9).split(',')[1]); const u=new Uint8Array(b.length); for(let i=0;i<b.length;i++) u[i]=b.charCodeAt(i); return u;}
async function rasterize(node,width,breaksOut,bsel,fixedH){
  const css=[...document.querySelectorAll('style')].map(s=>s.textContent).join('\n');
  const wrap=document.createElement('div'); wrap.className='pdfx'; wrap.style.cssText=`position:absolute;left:-100000px;top:0;width:${width}px;background:var(--page);padding:18px;box-sizing:border-box;font-family:'Be Vietnam Pro',system-ui,sans-serif;color:var(--ink)`;
  if(fixedH){wrap.classList.add('pg'); wrap.style.height=fixedH+'px'; wrap.style.overflow='hidden';}
  wrap.append(node); document.body.append(wrap); await nextFrame();
  const H=fixedH||Math.ceil(wrap.scrollHeight), top0=wrap.getBoundingClientRect().top;
  wrap.querySelectorAll(bsel||'.kpis,.card,.yc,.grid2,.ygrid,.pdfhead,.cmpnote').forEach(e=>{const r=e.getBoundingClientRect(); breaksOut.push(Math.round(r.top-top0)-8);});
  wrap.style.position='static'; wrap.style.left='0';
  const html=new XMLSerializer().serializeToString(wrap); wrap.remove();
  const svg=`<svg xmlns="http://www.w3.org/2000/svg" width="${width}" height="${H}"><foreignObject x="0" y="0" width="100%" height="100%"><div xmlns="http://www.w3.org/1999/xhtml"><style>${css.replace(/&/g,'&amp;').replace(/</g,'&lt;')}</style>${html}</div></foreignObject></svg>`;
  const img=new Image(); img.decoding='sync';
  await new Promise((ok,bad)=>{img.onload=ok; img.onerror=()=>bad(new Error('render')); setTimeout(()=>bad(new Error('timeout')),30000); img.src='data:image/svg+xml;charset=utf-8,'+encodeURIComponent(svg);});
  const sc=2, cv=document.createElement('canvas'); cv.width=width*sc; cv.height=H*sc; const g=cv.getContext('2d'); g.fillStyle='#eef1f6'; g.fillRect(0,0,cv.width,cv.height); g.scale(sc,sc); g.drawImage(img,0,0);
  cv.toDataURL('image/jpeg',0.5); // ném lỗi ngay nếu trình duyệt chặn (canvas bị "tainted")
  if(fixedH) return {cv,H,sc};
  // cắt phần nền trống thừa ở cuối (font trong ảnh có thể khác font trên trang → nội dung ngắn hơn)
  const W=cv.width; let last=cv.height-1; const step=32*sc;
  outer: for(let y=cv.height;y>0;y-=step){const y0=Math.max(0,y-step), d=g.getImageData(0,y0,W,y-y0).data;
    for(let i=d.length-4;i>=0;i-=4){if(Math.abs(d[i]-238)>6||Math.abs(d[i+1]-241)>6||Math.abs(d[i+2]-246)>6){last=y0+Math.floor(i/4/W); break outer;}}}
  const Ht=Math.min(cv.height,last+1+16*sc);
  if(Ht<cv.height-4){const c2=document.createElement('canvas'); c2.width=W; c2.height=Ht; c2.getContext('2d').drawImage(cv,0,0); return {cv:c2,H:Ht/sc,sc};}
  return {cv,H,sc};
}
function filterSummary(F){const sel=id=>{const e=$(id); return e.options[e.selectedIndex]?e.options[e.selectedIndex].text:'';};
  return `Năm ${F.Y}${$('#fCmp').value?' · so với '+$('#fCmp').value:''} · Khối: ${sel('#fKhoi')} · Phòng ban: ${sel('#fPb')} · Nhóm: ${sel('#fNhom')} · ${F.D0?'Kỳ: '+dmy(F.D0)+' – '+dmy(F.D1):'Đến ngày: '+dmy(F.T)}`;}
async function saveFile(blob,name){
  try{ if(window.claude&&typeof window.claude.use==='function'){ const d=await Promise.race([window.claude.use('downloads'),new Promise(r=>setTimeout(()=>r(null),4000))]);
      if(d){ try{await d.save({filename:name,data:blob}); return true;}catch(e){ if(e&&e.code==='declined') return false; if(e&&!['unavailable','not_granted','capability_disabled','capability_removed'].includes(e.code)) throw e; } } } }catch(e){}
  const a=document.createElement('a'); a.href=URL.createObjectURL(blob); a.download=name; document.body.append(a); a.click(); setTimeout(()=>{URL.revokeObjectURL(a.href); a.remove();},3000); return true;
}

// ---------- dùng chung cho PDF / PPT: bỏ khung đang thu gọn (đóng) và dàn lại bố cục
const EXPORT_TABS=[['perf','Hiệu quả đào tạo'],['budget','Ngân sách'],['sched','Kế hoạch & lịch'],['years','So sánh các năm']];
function tabHasContent(t){const p=$('#p-'+t); return !!p&&[...p.querySelectorAll('.kpis,.card,.yc')].some(c=>!c.classList.contains('collapsed')&&!c.classList.contains('nopdf')&&!c.closest('.nopdf'));}
function prepPanel(t){const src=$('#p-'+t); if(!src) return null; const p=src.cloneNode(true);
  p.removeAttribute('hidden'); p.removeAttribute('id'); p.querySelectorAll('[id]').forEach(e=>e.removeAttribute('id'));
  p.querySelectorAll('.nopdf,.card.collapsed').forEach(e=>e.remove()); p.querySelectorAll('.colbtn').forEach(e=>e.remove());
  // lưới 2 cột chỉ còn 1 khung → khung đó trải hết chiều ngang; lưới trống → bỏ
  p.querySelectorAll('.grid2,.ygrid').forEach(g=>{const k=[...g.children].filter(x=>!x.hidden); if(!k.length) g.remove(); else if(g.classList.contains('grid2')){if(k.length===1) g.style.gridTemplateColumns='minmax(0,1fr)'; else if(k.length%2) k[k.length-1].style.gridColumn='1 / -1';}});
  if(!p.querySelector('.kpis,.card,.yc')) return null; return p;}
// dựng các trang báo cáo khổ 16:9 (dùng chung cho PDF và PPT – hai file giống hệt nhau)
// mỗi trang lặp lại thanh tiêu đề (phần đang trình bày + bộ lọc); khung dài được ngắt theo dòng, ưu tiên ngắt ở đầu nhóm
const PGW=1360, PGH=765;
function pgHead(name,F0){const hd=el('div',{class:'pdfhead'}); const lg=document.querySelector('.brandlogo');
  const L=el('div',{class:'ph-l'}); if(lg){const im=el('img',{class:'ph-logo',src:lg.getAttribute('src'),alt:'SCC'}); L.append(im,el('i',{class:'ph-sep'}));}
  const tt=el('div'); tt.append(el('div',{class:'ph-t',text:'Training Dashboard'}),el('div',{class:'ph-s',text:'Phòng Nhân sự · L&OD · Saigon Cosmetics Corporation'})); L.append(tt);
  const Rr=el('div',{class:'ph-r'}); Rr.append(el('div',{class:'ph-tab',text:name}));
  const ch=el('div',{class:'ph-chips'}); const sel=id=>{const e=$(id); return e.options[e.selectedIndex]?e.options[e.selectedIndex].text:'';};
  [['Năm',String(F0.Y)+($('#fCmp').value?' so với '+$('#fCmp').value:'')],['Khối',sel('#fKhoi')],['Phòng ban',sel('#fPb')],['Nhóm',sel('#fNhom')],F0.D0?['Kỳ',dmy(F0.D0)+' – '+dmy(F0.D1)]:['Đến ngày',dmy(F0.T)]].forEach(([k,v])=>{const c=el('span'); c.append(el('em',{text:k}),document.createTextNode(' '+v)); ch.append(c);});
  Rr.append(ch); hd.append(L,Rr); return hd;}
function pgUnits(c){const tb=c.querySelector('table tbody'); if(tb) return {box:tb,isSec:x=>x.classList.contains('sec')};
  const pb=c.querySelector('.pbt'); if(pb&&pb.children[1]) return {box:pb.children[1],isSec:x=>x.classList.contains('pbsec')};
  const ins=c.querySelector('.ins'); if(ins) return {box:ins,isSec:()=>false}; return null;}
const pgItems=U=>[...U.box.children].filter(x=>!x.hidden);
function pgCont(x){const t=x.querySelector('b,.nm')||x.querySelector('td')||x; t.append(document.createTextNode(' (tiếp)'));}
function pgTitleCont(c){const h=c.querySelector('h3'); if(h&&!/\(tiếp\)$/.test(h.textContent)) h.append(document.createTextNode(' (tiếp)'));}
async function buildPages(lbl){
  const F0=filt(); const TABS=EXPORT_TABS.filter(([t])=>tabHasContent(t)); const out=[];
  const host=document.createElement('div'); host.className='pdfx pg'; host.style.cssText=`position:absolute;left:-100000px;top:0;width:${PGW}px;height:${PGH}px;padding:18px;box-sizing:border-box;background:var(--page);font-family:'Be Vietnam Pro',system-ui,sans-serif;color:var(--ink)`;
  document.body.append(host);
  try{
  for(let ti=0;ti<TABS.length;ti++){const [t,name]=TABS[ti]; lbl.textContent=`Đang dàn trang ${ti+1}/${TABS.length}…`;
    S.tab=t; render(); await nextFrame();
    const p=prepPanel(t); if(!p) continue;
    const q=[...p.children].filter(x=>!x.hidden&&x.offsetParent!==undefined);
    const cn=$('#cmpNote'); if(!cn.hidden&&t!=='years'){const c=cn.cloneNode(true); c.className='cmpnote'; c.removeAttribute('id'); c.style.cssText='font-size:12.5px;color:var(--ink-2);padding:8px 12px;background:#fff4e8;border:1px solid #f6c38f;border-radius:10px'; q.unshift(c);}
    let cur=null;
    const newPage=()=>{const node=el('div',{class:'pgnode'}); node.append(pgHead(name,F0)); const body=el('div',{class:'pgbody'}); const ft=el('div',{class:'pdffoot',text:' '}); node.append(body,ft); host.replaceChildren(node); cur={node,body,ft,name}; return cur;};
    const fits=()=>cur.body.scrollHeight<=cur.body.clientHeight+1;
    const flush=()=>{if(cur&&cur.body.children.length){host.replaceChildren(); out.push(cur);} cur=null;};
    // tách 1 khung có bảng/danh sách: lấy phần vừa với chỗ trống còn lại, phần còn lại thành khung "(tiếp)"
    const splitCard=(b,minRows)=>{const U0=pgUnits(b); if(!U0) return null;
      const c=b.cloneNode(true); const U=pgUnits(c); const all=pgItems(U); [...U.box.children].forEach(x=>x.remove()); cur.body.append(c);
      let n=0,dn=0,lastSec=null;
      for(let i=0;i<all.length;i++){const x=all[i];
        if(U.isSec(x)){const grp=[x]; for(let j=i+1;j<all.length&&!U.isSec(all[j]);j++) grp.push(all[j]); grp.forEach(g=>U.box.append(g)); const ok=fits(); grp.forEach(g=>g.remove()); if(!ok&&dn>0) break;}
        U.box.append(x); if(!fits()){x.remove(); break;} n++; if(U.isSec(x)) lastSec=x; else dn++;}
      c.remove(); if(dn<minRows||n>=all.length) return null;
      while(n>0&&U.isSec(all[n-1])){n--; all[n].remove(); lastSec=null; for(let k=n-1;k>=0;k--) if(U.isSec(all[k])){lastSec=all[k]; break;}}
      const rest=b.cloneNode(true); const R=pgUnits(rest); const rall=pgItems(R); [...R.box.children].filter(x=>x.hidden).forEach(x=>x.remove()); rall.slice(0,n).forEach(x=>x.remove());
      const first=pgItems(R)[0]; if(lastSec&&first&&!R.isSec(first)){const s2=lastSec.cloneNode(true); pgCont(s2); R.box.prepend(s2);}
      pgTitleCont(rest); return [c,rest];};
    const splitGrid=b=>{const kids=[...b.children].filter(x=>!x.hidden); if(kids.length<=1) return kids.length?[kids[0]]:[];
      const per=b.classList.contains('ygrid')?3:2; if(kids.length<=per) return kids;
      const rows=[]; for(let i=0;i<kids.length;i+=per){const g=b.cloneNode(false); g.style.cssText=b.style.cssText; kids.slice(i,i+per).forEach(k=>{k.style.gridColumn=''; g.append(k);}); if(g.children.length===1&&per===2) g.style.gridTemplateColumns='minmax(0,1fr)'; rows.push(g);} return rows;};
    newPage(); let guard=0;
    while(q.length&&guard++<400){const b=q.shift();
      cur.body.append(b); if(fits()) continue; b.remove();
      const empty=!cur.body.children.length;
      if(!empty){const room=cur.body.clientHeight-cur.body.scrollHeight;
        if(room>220&&pgUnits(b)&&!b.matches('.grid2,.ygrid')){const sp=splitCard(b,3); if(sp){cur.body.append(sp[0]); flush(); newPage(); q.unshift(sp[1]); continue;}}
        flush(); newPage(); q.unshift(b); continue;}
      if(b.matches('.grid2,.ygrid')){const r=splitGrid(b); if(r.length&&!(r.length===1&&r[0]===b)){q.unshift(...r); continue;}}
      if(pgUnits(b)){const sp=splitCard(b,1); if(sp){cur.body.append(sp[0]); flush(); newPage(); q.unshift(sp[1]); continue;}}
      // không tách được: thu nhỏ cho vừa 1 trang
      cur.body.append(b); const h=b.getBoundingClientRect().height, mx=cur.body.clientHeight; if(h>mx) b.style.zoom=String(Math.max(0.45,(mx-2)/h)); flush(); newPage();}
    flush();}
  // đánh số trang rồi chụp
  for(let i=0;i<out.length;i++){const pg=out[i]; lbl.textContent=`Đang xuất trang ${i+1}/${out.length}…`;
    pg.ft.textContent=`SCC Training Dashboard · ${pg.name} · Trang ${i+1}/${out.length} · Xuất lúc ${new Date().toLocaleString('vi-VN')}`;
    const {cv}=await rasterize(pg.node,PGW,[],null,PGH);
    out[i]={jpeg:jpegBytes(cv),w:cv.width,h:cv.height,full:1};}
  } finally{host.remove();}
  return {pages:out,F0};
}
async function exportPdf(){
  if(!M) return; const btn=$('#pdfBtn'), lbl=$('#pdfLbl'); btn.disabled=true; const prevTab=S.tab;
  try{
    const {pages,F0}=await buildPages(lbl); if(!pages.length) return;
    lbl.textContent='Đang lưu…';
    const blob=buildPdf(pages); const name=`SCC_Training_Dashboard_${F0.Y}_${iso(new Date())}.pdf`;
    await saveFile(blob,name);
  }catch(err){ console.error(err);
    $('#alerts').prepend(el('div',{class:'alert',text:'Trình duyệt này không cho xuất PDF trực tiếp. Hãy dùng Chrome hoặc Edge, hoặc nhấn Ctrl+P → "Lưu dưới dạng PDF".'})); }
  finally{ S.tab=prevTab; render(); btn.disabled=false; lbl.textContent='Tải PDF'; }
}

// ---------- xuất PowerPoint (tự tạo file .pptx, không cần thư viện ngoài): giống hệt file PDF, mỗi trang = 1 slide
const CRCT=(()=>{const t=new Uint32Array(256); for(let n=0;n<256;n++){let c=n; for(let k=0;k<8;k++) c=c&1?0xEDB88320^(c>>>1):c>>>1; t[n]=c>>>0;} return t;})();
function crc32(u){let c=0xFFFFFFFF; for(let i=0;i<u.length;i++) c=CRCT[(c^u[i])&255]^(c>>>8); return (c^0xFFFFFFFF)>>>0;}
function zipStore(files){const enc=new TextEncoder(); const parts=[], cen=[]; let off=0;
  const w16=(d,o,v)=>{d[o]=v&255; d[o+1]=(v>>>8)&255;}, w32=(d,o,v)=>{d[o]=v&255; d[o+1]=(v>>>8)&255; d[o+2]=(v>>>16)&255; d[o+3]=(v>>>24)&255;};
  files.forEach(f=>{const nm=enc.encode(f.name); const data=typeof f.data==='string'?enc.encode(f.data):f.data; const crc=crc32(data);
    const h=new Uint8Array(30+nm.length); w32(h,0,0x04034b50); w16(h,4,20); w16(h,6,0x0800); w16(h,8,0); w16(h,10,0); w16(h,12,0x21); w32(h,14,crc); w32(h,18,data.length); w32(h,22,data.length); w16(h,26,nm.length); w16(h,28,0); h.set(nm,30);
    const c=new Uint8Array(46+nm.length); w32(c,0,0x02014b50); w16(c,4,20); w16(c,6,20); w16(c,8,0x0800); w16(c,10,0); w16(c,12,0); w16(c,14,0x21); w32(c,16,crc); w32(c,20,data.length); w32(c,24,data.length); w16(c,28,nm.length); w32(c,42,off); c.set(nm,46);
    parts.push(h,data); cen.push(c); off+=h.length+data.length;});
  const cs=cen.reduce((x,c)=>x+c.length,0); const end=new Uint8Array(22); w32(end,0,0x06054b50); w16(end,8,files.length); w16(end,10,files.length); w32(end,12,cs); w32(end,16,off);
  return new Blob([...parts,...cen,end],{type:'application/vnd.openxmlformats-officedocument.presentationml.presentation'});}
function b64bytes(d){const s=atob(d.split(',')[1]); const u=new Uint8Array(s.length); for(let i=0;i<s.length;i++) u[i]=s.charCodeAt(i); return u;}
function makePptx(SWin,SHin){
  const E=v=>Math.round(v*914400), X=s=>String(s).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;');
  const NS='xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"';
  const HD='<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'; const GRP='<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>';
  const slides=[], media=[]; const xf=(x,y,w,h)=>`<a:xfrm><a:off x="${E(x)}" y="${E(y)}"/><a:ext cx="${E(w)}" cy="${E(h)}"/></a:xfrm>`;
  const P={addSlide(bg){const s={bg:bg||'FFFFFF',sh:[],rel:[],id:2};
      s.rect=(x,y,w,h,col)=>{s.sh.push(`<p:sp><p:nvSpPr><p:cNvPr id="${s.id++}" name="Shape ${s.id}"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr>${xf(x,y,w,h)}<a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:solidFill><a:srgbClr val="${col}"/></a:solidFill><a:ln><a:noFill/></a:ln></p:spPr></p:sp>`); return s;};
      s.text=(t,o)=>{const sz=Math.round((o.size||14)*100); s.sh.push(`<p:sp><p:nvSpPr><p:cNvPr id="${s.id++}" name="Text ${s.id}"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr><p:spPr>${xf(o.x,o.y,o.w,o.h)}<a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/></p:spPr><p:txBody><a:bodyPr wrap="square" lIns="0" tIns="0" rIns="0" bIns="0" anchor="${o.anchor||'t'}"><a:normAutofit/></a:bodyPr><a:lstStyle/><a:p><a:pPr algn="${o.align||'l'}"/><a:r><a:rPr lang="vi-VN" sz="${sz}" b="${o.bold?1:0}"${o.spc?` spc="${o.spc}"`:''} dirty="0"><a:solidFill><a:srgbClr val="${o.color||'141B2D'}"/></a:solidFill><a:latin typeface="Arial"/><a:cs typeface="Arial"/></a:rPr><a:t>${X(t)}</a:t></a:r></a:p></p:txBody></p:sp>`); return s;};
      s.image=(bytes,ext,x,y,w,h)=>{const n=media.length+1; media.push({name:`ppt/media/image${n}.${ext}`,data:bytes}); const rid='rId'+(s.rel.length+2); s.rel.push(`<Relationship Id="${rid}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image${n}.${ext}"/>`);
        s.sh.push(`<p:pic><p:nvPicPr><p:cNvPr id="${s.id++}" name="Picture ${s.id}"/><p:cNvPicPr><a:picLocks noChangeAspect="1"/></p:cNvPicPr><p:nvPr/></p:nvPicPr><p:blipFill><a:blip r:embed="${rid}"/><a:stretch><a:fillRect/></a:stretch></p:blipFill><p:spPr>${xf(x,y,w,h)}<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:pic>`); return s;};
      slides.push(s); return s;},
    blob(title){const F=[]; const n=slides.length;
      F.push({name:'[Content_Types].xml',data:HD+'<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Default Extension="jpeg" ContentType="image/jpeg"/><Default Extension="png" ContentType="image/png"/><Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/><Override PartName="/ppt/slideMasters/slideMaster1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml"/><Override PartName="/ppt/slideLayouts/slideLayout1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"/><Override PartName="/ppt/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>'+slides.map((_,i)=>`<Override PartName="/ppt/slides/slide${i+1}.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>`).join('')+'<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/><Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/></Types>'});
      F.push({name:'_rels/.rels',data:HD+'<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/><Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/></Relationships>'});
      const now=new Date().toISOString().replace(/\.\d+Z$/,'Z');
      F.push({name:'docProps/core.xml',data:HD+`<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"><dc:title>${X(title)}</dc:title><dc:creator>Phòng Nhân sự · L&amp;OD</dc:creator><dcterms:created xsi:type="dcterms:W3CDTF">${now}</dcterms:created><dcterms:modified xsi:type="dcterms:W3CDTF">${now}</dcterms:modified></cp:coreProperties>`});
      F.push({name:'docProps/app.xml',data:HD+`<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"><Application>Microsoft Office PowerPoint</Application><Slides>${n}</Slides><Company>Saigon Cosmetics Corporation</Company></Properties>`});
      F.push({name:'ppt/presentation.xml',data:HD+`<p:presentation ${NS} saveSubsetFonts="1"><p:sldMasterIdLst><p:sldMasterId id="2147483648" r:id="rId1"/></p:sldMasterIdLst><p:sldIdLst>${slides.map((_,i)=>`<p:sldId id="${256+i}" r:id="rId${i+3}"/>`).join('')}</p:sldIdLst><p:sldSz cx="${E(SWin||13.333)}" cy="${E(SHin||7.5)}"/><p:notesSz cx="6858000" cy="9144000"/></p:presentation>`});
      F.push({name:'ppt/_rels/presentation.xml.rels',data:HD+'<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="slideMasters/slideMaster1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/>'+slides.map((_,i)=>`<Relationship Id="rId${i+3}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide${i+1}.xml"/>`).join('')+'</Relationships>'});
      const sc=(nm,c)=>`<a:${nm}><a:srgbClr val="${c}"/></a:${nm}>`; const sf='<a:solidFill><a:schemeClr val="phClr"/></a:solidFill>';
      F.push({name:'ppt/theme/theme1.xml',data:HD+`<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="SCC"><a:themeElements><a:clrScheme name="SCC"><a:dk1><a:srgbClr val="141B2D"/></a:dk1><a:lt1><a:srgbClr val="FFFFFF"/></a:lt1>${sc('dk2','14213D')}${sc('lt2','EEF1F6')}${sc('accent1','F28C28')}${sc('accent2','2F6FDE')}${sc('accent3','1AA673')}${sc('accent4','F2B01E')}${sc('accent5','D64545')}${sc('accent6','4A5470')}${sc('hlink','2F6FDE')}${sc('folHlink','7D869B')}</a:clrScheme><a:fontScheme name="SCC"><a:majorFont><a:latin typeface="Arial"/><a:ea typeface=""/><a:cs typeface=""/></a:majorFont><a:minorFont><a:latin typeface="Arial"/><a:ea typeface=""/><a:cs typeface=""/></a:minorFont></a:fontScheme><a:fmtScheme name="SCC"><a:fillStyleLst>${sf}${sf}${sf}</a:fillStyleLst><a:lnStyleLst><a:ln w="6350">${sf}</a:ln><a:ln w="12700">${sf}</a:ln><a:ln w="19050">${sf}</a:ln></a:lnStyleLst><a:effectStyleLst><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle></a:effectStyleLst><a:bgFillStyleLst>${sf}${sf}${sf}</a:bgFillStyleLst></a:fmtScheme></a:themeElements><a:objectDefaults/><a:extraClrSchemeLst/></a:theme>`});
      F.push({name:'ppt/slideMasters/slideMaster1.xml',data:HD+`<p:sldMaster ${NS}><p:cSld><p:bg><p:bgRef idx="1001"><a:schemeClr val="bg1"/></p:bgRef></p:bg><p:spTree>${GRP}</p:spTree></p:cSld><p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/><p:sldLayoutIdLst><p:sldLayoutId id="2147483649" r:id="rId1"/></p:sldLayoutIdLst><p:txStyles><p:titleStyle/><p:bodyStyle/><p:otherStyle/></p:txStyles></p:sldMaster>`});
      F.push({name:'ppt/slideMasters/_rels/slideMaster1.xml.rels',data:HD+'<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="../theme/theme1.xml"/></Relationships>'});
      F.push({name:'ppt/slideLayouts/slideLayout1.xml',data:HD+`<p:sldLayout ${NS} type="blank" preserve="1"><p:cSld name="Blank"><p:spTree>${GRP}</p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sldLayout>`});
      F.push({name:'ppt/slideLayouts/_rels/slideLayout1.xml.rels',data:HD+'<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="../slideMasters/slideMaster1.xml"/></Relationships>'});
      slides.forEach((s,i)=>{F.push({name:`ppt/slides/slide${i+1}.xml`,data:HD+`<p:sld ${NS}><p:cSld><p:bg><p:bgPr><a:solidFill><a:srgbClr val="${s.bg}"/></a:solidFill><a:effectLst/></p:bgPr></p:bg><p:spTree>${GRP}${s.sh.join('')}</p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sld>`});
        F.push({name:`ppt/slides/_rels/slide${i+1}.xml.rels`,data:HD+'<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>'+s.rel.join('')+'</Relationships>'});});
      return zipStore(F.concat(media));}};
  return P;}
async function exportPpt(){
  if(!M) return; const btn=$('#pptBtn'), lbl=$('#pptLbl'); btn.disabled=true; const prevTab=S.tab;
  try{
    const {pages,F0}=await buildPages(lbl); if(!pages.length) return;
    lbl.textContent='Đang lưu…';
    // mỗi trang PDF = 1 slide 16:9, ảnh phủ kín slide – giống hệt file PDF
    const pp=makePptx(13.333,7.5);
    pages.forEach(p=>{const s=pp.addSlide('EEF1F6'); s.image(p.jpeg,'jpeg',0,0,13.333,7.5);});
    await saveFile(pp.blob('SCC Training Dashboard'),`SCC_Training_Dashboard_${F0.Y}_${iso(new Date())}.pptx`);
  }catch(err){ console.error(err);
    $('#alerts').prepend(el('div',{class:'alert',text:'Chưa xuất được file PowerPoint. Hãy dùng Chrome hoặc Edge và thử lại.'})); }
  finally{ S.tab=prevTab; render(); btn.disabled=false; lbl.textContent='Tải PPT'; }
}
// ---------- nút thu gọn / mở rộng cho từng khung
const CHEV='<svg width="16" height="16" viewBox="0 0 24 24" fill="none" aria-hidden="true"><path d="M6 9l6 6 6-6" stroke="currentColor" stroke-width="2.2" stroke-linecap="round" stroke-linejoin="round"/></svg>';
document.querySelectorAll('#app .card').forEach(card=>{
  let ch=card.querySelector(':scope > .ch');
  if(!ch){const h3=card.querySelector(':scope > h3'); if(!h3) return; ch=document.createElement('div'); ch.className='ch'; h3.before(ch); ch.append(h3);}
  const b=document.createElement('button'); b.type='button'; b.className='colbtn'; b.title='Thu gọn / mở rộng'; b.setAttribute('aria-expanded','true'); b.innerHTML=CHEV;
  b.addEventListener('click',()=>{const c=card.classList.toggle('collapsed'); b.setAttribute('aria-expanded',String(!c)); if(!c) render();});
  ch.append(b);
});
// ---------- events
['#fPb','#fNhom','#fDate','#fStatus','#fCmp'].forEach(s=>$(s).addEventListener('change',render));
$('#fFrom').addEventListener('change',()=>{const d=toDate($('#fFrom').value); if(d&&YEARS.includes(d.getFullYear())&&+$('#fYear').value!==d.getFullYear()){$('#fYear').value=d.getFullYear(); fillCmp();}
  const t=toDate($('#fDate').value); if(d&&t&&t<d){const e=new Date(d.getFullYear(),d.getMonth()+1,0); $('#fDate').value=iso(e);} render();});
$('#fKhoi').addEventListener('change',()=>{fillPb(); render();});
$('#qNV').addEventListener('input',()=>renderNV());
$('#pdfBtn').addEventListener('click',exportPdf);
$('#pptBtn').addEventListener('click',exportPpt);
$('#fYear').addEventListener('change',()=>{$('#fFrom').value=''; fillCmp(); render();});
document.querySelectorAll('.tab').forEach(t=>t.addEventListener('click',()=>{S.tab=t.dataset.p; try{localStorage.setItem('scc-td-tab',S.tab)}catch(e){} render();}));
$('#mSeg').addEventListener('click',e=>{const b=e.target.closest('button'); if(!b) return; S.month=b.dataset.m; $('#mSeg').querySelectorAll('button').forEach(x=>x.setAttribute('aria-pressed',String(x===b))); render();});
// merge several workbooks (e.g. one file per year): later files' columns are re-mapped onto the first file's headers by name
const KEYH={nv:'Mã NV',kh:'Mã lớp',lop:'Mã lớp',tr:'Mã lớp',kp:'Nội dung đào tạo'};
function hdrRow(aoa,key){for(let i=0;i<Math.min(12,aoa.length);i++){if((aoa[i]||[]).some(c=>norm(c)===norm(key))) return i;} return -1;}
function mergeBooks(books){
  if(books.length===1) return books[0];
  const out={};
  for(const k of ['nv','kh','lop','tr','kp']){
    let base=null,h1=null;
    for(const b of books){const a=b[k]; if(!a) continue; const hi=hdrRow(a,KEYH[k]); if(hi<0) continue;
      if(!base){base=a.slice(); h1=a[hi].map(norm); continue;}
      const h2=a[hi].map(norm); const idx=h1.map(h=>h?h2.indexOf(h):-1);
      for(const r of a.slice(hi+1)) base.push(idx.map(j=>j<0?null:r[j]));}
    out[k]=base;}
  return out;
}
let handles=null, lastFiles=null;
// lưu "tay cầm" file (File System Access) vào IndexedDB để Làm mới / mở lại trang đọc đúng file cũ
const HDB={open(){return new Promise((res,rej)=>{try{const r=indexedDB.open('scc-training-dash',1); r.onupgradeneeded=()=>r.result.createObjectStore('h'); r.onsuccess=()=>res(r.result); r.onerror=()=>rej(r.error);}catch(e){rej(e);}});},
  async set(v){try{const db=await this.open(); db.transaction('h','readwrite').objectStore('h').put(v,'last');}catch(e){}},
  async get(){try{const db=await this.open(); return await new Promise(r=>{const q=db.transaction('h').objectStore('h').get('last'); q.onsuccess=()=>r(q.result||null); q.onerror=()=>r(null);});}catch(e){return null;}}};
function keepHandles(hs){handles=hs; HDB.set(hs); const rb=$('#reopenBtn'); if(rb) rb.hidden=true;}
async function canRead(hs,ask){for(const h of hs){if(!h.queryPermission) continue; let p=await h.queryPermission({mode:'read'}); if(p!=='granted'&&ask) p=await h.requestPermission({mode:'read'}); if(p!=='granted') return false;} return true;}
async function loadHandles(hs,refresh){handles=hs; return loadFiles(await Promise.all(hs.map(h=>h.getFile())),refresh);}
const fi=$('#fileIn');
async function pick(){
  // File System Access API keeps a handle so "Làm mới" can re-read without choosing again (Chrome/Edge, when allowed)
  if(window.showOpenFilePicker){
    try{const hs=await window.showOpenFilePicker({multiple:true,types:[{description:'Excel',accept:{'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet':['.xlsx','.xlsm']}}]});
      keepHandles(hs); loadFiles(await Promise.all(hs.map(h=>h.getFile()))); return;}
    catch(e){if(e&&e.name==='AbortError') return;}
  }
  handles=null; fi.click();
}
$('#pickBtn3').addEventListener('click',pick); $('#landDrop').addEventListener('click',e=>{if(!e.target.closest('button')) pick();}); $('#landDrop').addEventListener('keydown',e=>{if(e.key==='Enter'||e.key===' '){e.preventDefault(); pick();}});
$('#refreshBtn').addEventListener('click',async()=>{
  const btn=$('#refreshBtn'); if(btn.classList.contains('spin')) return; btn.classList.add('spin');
  try{let hs=handles; if(!hs&&!source.sample) hs=await HDB.get();
    if(hs&&hs.length){try{if(await canRead(hs,true)){await loadHandles(hs,true); return;}}catch(e){}}
    const host=$('#alerts'); host.replaceChildren(el('div',{class:'alert info',text:source.sample?'Chọn file SCC Training Record.xlsx để xem dữ liệu thật.':'Trình duyệt chưa cho phép tự đọc lại file. Hãy chọn file SCC Training Record.xlsx một lần – từ lần sau chỉ cần bấm Làm mới.'}));
    await pick();}
  finally{btn.classList.remove('spin');}
});
// mở lại file lần trước: tự tải nếu đã được cấp quyền, nếu chưa thì hiện nút trên trang đầu
(async()=>{const hs=await HDB.get(); if(!hs||!hs.length) return;
  try{if(await canRead(hs,false)){await loadHandles(hs,false); return;}}catch(e){}
  const rb=$('#reopenBtn'); rb.textContent='Mở lại: '+hs.map(h=>h.name).join(', '); rb.hidden=false;
  rb.addEventListener('click',async e=>{e.stopPropagation(); try{if(await canRead(hs,true)){rb.hidden=true; await loadHandles(hs,false); return;}}catch(err){} pick();});})();
$('#tryBtn').addEventListener('click',()=>{handles=null; setModel(buildModel(SAMPLE),{name:'Dữ liệu mẫu',sample:true});});
fi.addEventListener('change',()=>{if(fi.files.length) loadFiles([...fi.files]); fi.value='';});
const drop=document.body;
['dragenter','dragover'].forEach(ev=>document.addEventListener(ev,e=>{e.preventDefault(); drop.classList.add('over');}));
['dragleave','drop'].forEach(ev=>document.addEventListener(ev,e=>{e.preventDefault(); if(ev==='drop'||!e.relatedTarget) drop.classList.remove('over');}));
document.addEventListener('drop',e=>{const dt=e.dataTransfer; const hp=dt?[...dt.items].filter(i=>i.kind==='file'&&i.getAsFileSystemHandle).map(i=>i.getAsFileSystemHandle()):[];
  const fs=dt&&[...dt.files].filter(f=>/\.xls[xm]$/i.test(f.name)); if(fs&&fs.length){handles=null; loadFiles(fs);
    Promise.all(hp).then(hs=>{hs=hs.filter(h=>h&&h.kind==='file'&&/\.xls[xm]$/i.test(h.name)); if(hs.length) keepHandles(hs);}).catch(()=>{});}});
$('#sampleBtn').addEventListener('click',()=>{M=null; handles=null; showLanding(); $('#alerts').replaceChildren(); $('#alertsLand').replaceChildren();});
async function loadFiles(files,refresh=false){
  const host=$('#landing').hidden?$('#alerts'):$('#alertsLand'); $('#alertsLand').replaceChildren();
  const names=files.map(f=>f.name).join(', ');
  host.replaceChildren(el('div',{class:'alert info',text:'Đang đọc '+names+'…'}));
  document.querySelectorAll('section.panel').forEach(p=>p.style.opacity='.5');
  try{const books=[]; for(const f of files) books.push(await readWorkbook(new Uint8Array(await f.arrayBuffer())));
    const m=buildModel(mergeBooks(books)); if(!refresh){$('#fDate').value=''; $('#fFrom').value='';}
    lastFiles=files; setModel(m,{name:files.length>1?files.length+' file':files[0].name,sample:false});
    if(refresh){const t=new Date(); $('#alerts').prepend(el('div',{class:'alert info',text:`Đã làm mới lúc ${String(t.getHours()).padStart(2,'0')}:${String(t.getMinutes()).padStart(2,'0')} từ ${names}.`}));}}
  catch(err){if(!M) showLanding(); host.replaceChildren(el('div',{class:'alert',text:'Không đọc được file: '+(err&&err.message||err)}));}
  document.querySelectorAll('section.panel').forEach(p=>p.style.opacity='');
}
let rt; addEventListener('resize',()=>{clearTimeout(rt); rt=setTimeout(render,150);});
showLanding();
})();
