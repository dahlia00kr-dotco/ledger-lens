/* ============================================================
   상태 / 저장소
   ============================================================ */
const $=s=>document.querySelector(s);
const esc=s=>String(s??'').replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
const pad=n=>String(n).padStart(2,'0');
const fmt=n=>(n==null||isNaN(n))?'':Math.round(n).toLocaleString('ko-KR');
const pct=n=>(n==null||!isFinite(n))?'':(n>0?'+':'')+n.toFixed(1)+'%';
const median=a=>{if(!a.length)return 0;const s=[...a].sort((x,y)=>x-y),m=s.length>>1;return s.length%2?s[m]:(s[m-1]+s[m])/2};
const nowIso=()=>new Date().toISOString().slice(0,16).replace('T',' ');
const DBKEY='jv-review-db-v1';
const defaultSettings=()=>({minMonths:3,minRatio:0.5,monthlyRatio:0.8,tol:10,mode:'va'});
const defaultDB=()=>({version:1,patterns:{},once:{},issues:{},fixed:{},settings:defaultSettings(),log:[]});
function normDB(d){const b=defaultDB();d=Object.assign(b,d||{});d.settings=Object.assign(defaultSettings(),d.settings||{});return d}
function loadDB(){try{const s=localStorage.getItem(DBKEY);if(s)return normDB(JSON.parse(s))}catch(e){}return defaultDB()}
// 포트폴리오 체험판: 샘플 데이터를 자동으로 불러오고, 내 파일 업로드·패턴DB 파일 저장은 숨깁니다.
// 실무용으로 쓰려면 false로 바꾸세요.
const DEMO=true;
let DB=DEMO?defaultDB():loadDB();
let ROWS=[],VOUCH=new Map(),FILES=[],SIGS=new Set();
let fileHandle=null,dirty=false,saveTimer=null;
const S={kind:'all',month:null,tab:'summary',metric:'n',onlyFlag:false,onlyNoVendor:false,q:'',prFilter:'todo',expanded:new Set(),selected:new Set(),dbq:'',ckFilter:'all'};
let C={};

function saveDB(){
  DB.updatedAt=nowIso();
  if(DEMO){renderStatus();return}
  try{localStorage.setItem(DBKEY,JSON.stringify(DB))}catch(e){}
  dirty=true;
  if(fileHandle){clearTimeout(saveTimer);saveTimer=setTimeout(()=>writeHandle(true),800)}
  renderStatus();
}
async function writeHandle(silent){
  try{
    if(fileHandle.queryPermission&&(await fileHandle.queryPermission({mode:'readwrite'}))!=='granted'){
      if(silent)return false;
      if((await fileHandle.requestPermission({mode:'readwrite'}))!=='granted')return false;
    }
    const w=await fileHandle.createWritable();await w.write(JSON.stringify(DB,null,1));await w.close();
    dirty=false;renderStatus();if(!silent)toast('저장됨 · '+fileHandle.name);return true;
  }catch(e){if(!silent)toast('파일 저장 실패: '+e.message);return false}
}
async function saveDbFile(){
  if(fileHandle&&await writeHandle(false))return;
  if(window.showSaveFilePicker){
    try{fileHandle=await showSaveFilePicker({suggestedName:'전표검토_패턴DB.json',types:[{description:'JSON',accept:{'application/json':['.json']}}]});
      await writeHandle(false);return}catch(e){if(e.name==='AbortError')return;fileHandle=null}
  }
  download(JSON.stringify(DB,null,1),'전표검토_패턴DB_'+new Date().toISOString().slice(0,10)+'.json','application/json');
  dirty=false;renderStatus();
}
async function openDbFile(){
  if(window.showOpenFilePicker){
    try{const [h]=await showOpenFilePicker({types:[{description:'JSON',accept:{'application/json':['.json']}}]});
      const txt=await (await h.getFile()).text();applyDbText(txt,h.name);fileHandle=h;dirty=false;renderStatus();return}
    catch(e){if(e.name==='AbortError')return}
  }
  $('#fDb').click();
}
function applyDbText(txt,name){
  try{const d=JSON.parse(txt);if(!d.patterns)throw new Error('패턴DB 형식이 아닙니다');DB=normDB(d);
    try{localStorage.setItem(DBKEY,JSON.stringify(DB))}catch(e){}
    toast(`패턴DB 불러옴 · ${name} (정상패턴 ${Object.keys(DB.patterns).length}개)`);refresh();
  }catch(e){alert('패턴DB를 읽지 못했습니다: '+e.message)}
}
function download(content,name,type){
  const a=document.createElement('a');a.href=URL.createObjectURL(new Blob([content],{type}));a.download=name;a.click();
  setTimeout(()=>URL.revokeObjectURL(a.href),2000);
}
function toast(m){const t=$('#toast');t.textContent=m;t.classList.add('show');clearTimeout(t._t);t._t=setTimeout(()=>t.classList.remove('show'),2600)}

/* ============================================================
   엑셀 파싱 (더존 iCUBE 전표출력 형식 + 유사 형식)
   ============================================================ */
function dstr(v){
  if(v==null||v==='')return null;
  if(typeof v==='number'){const p=XLSX.SSF.parse_date_code(v);return p?`${p.y}-${pad(p.m)}-${pad(p.d)}`:null}
  const m=String(v).match(/(\d{4})\D?(\d{1,2})\D?(\d{1,2})/);return m?`${m[1]}-${pad(m[2])}-${pad(m[3])}`:null;
}
function num(v){if(v==null||v==='')return 0;if(typeof v==='number')return v;const n=parseFloat(String(v).replace(/[,\s원]/g,''));return isNaN(n)?0:n}
const str=v=>v==null?'':String(v).trim();
function tplOf(s){
  return String(s||'').replace(/\[[^\]]*\]/g,'').replace(/\([^)]*\)/g,'').replace(/\d+/g,'#').replace(/\s+/g,' ').trim();
}
function parseWorkbook(wb,fname){
  const out=[];
  for(const sn of wb.SheetNames){
    const aoa=XLSX.utils.sheet_to_json(wb.Sheets[sn],{header:1,raw:true,defval:null});
    let hi=-1;
    for(let i=0;i<Math.min(30,aoa.length);i++){const r=(aoa[i]||[]).map(str);if(r.includes('계정과목명')||r.includes('계정명')||r.includes('계정과목')){hi=i;break}}
    if(hi<0)continue;
    const h=aoa[hi].map(str);
    const f=(...n)=>{for(const x of n){const i=h.indexOf(x);if(i>=0)return i}return -1};
    const iDate=f('기표일자','전표일자','회계일자','일자','결의일'),iNo=f('기표번호','전표번호','결의No','번호'),iSeq=f('순번');
    const iAN=f('계정과목명','계정명','계정과목');let iAC=f('계정코드');if(iAC<0&&h[iAN-1]==='코드')iAC=iAN-1;
    const iVN=f('거래처명','거래처');let iVC=f('거래처코드');if(iVC<0&&iVN>0&&h[iVN-1]==='코드')iVC=iVN-1;
    const iMemo=f('적요','적요명'),iDr=f('차변','차변금액'),iCr=f('대변','대변금액');
    if(iDate<0||iAN<0||iDr<0||iCr<0)continue;
    for(let i=hi+1;i<aoa.length;i++){
      const r=aoa[i];if(!r)continue;
      const date=dstr(r[iDate]);if(!date)continue;
      const aName=str(r[iAN]);if(!aName)continue;
      const dr=num(r[iDr]),cr=num(r[iCr]);if(!dr&&!cr)continue;
      const aCode=iAC>=0?str(r[iAC]):'';const vName=iVN>=0?str(r[iVN]):'';const vCode=iVC>=0?str(r[iVC]):'';
      const memo=iMemo>=0?str(r[iMemo]):'';const vno=iNo>=0?str(r[iNo]):'';
      out.push({date,ym:date.slice(0,7),vno,seq:iSeq>=0?str(r[iSeq]):'',vid:date+'#'+vno,
        aKey:aCode||aName,aName,vKey:vCode?'C'+vCode:(vName?'N'+vName:'-'),vName:vName||'(거래처 없음)',noVendor:!vCode&&!vName,
        memo,tpl:tplOf(memo),dr,cr,side:dr?'차':'대',amt:dr+cr,file:fname});
    }
  }
  return out;
}
async function loadJournals(files){
  let added=0,dup=0;
  for(const file of files){
    const buf=await file.arrayBuffer();
    let wb;try{wb=XLSX.read(buf,{type:'array'})}catch(e){alert(file.name+' 읽기 실패: '+e.message);continue}
    const rows=parseWorkbook(wb,file.name);
    if(!rows.length){alert(file.name+': 전표 형식(계정과목명/차변/대변/기표일자 열)을 찾지 못했습니다.');continue}
    for(const r of rows){
      const sig=[r.date,r.vno,r.seq,r.aKey,r.vKey,r.dr,r.cr,r.memo].join('|');
      if(SIGS.has(sig)){dup++;continue}SIGS.add(sig);ROWS.push(r);added++;
    }
    FILES.push({name:file.name,n:rows.length});
  }
  ROWS.sort((a,b)=>a.date<b.date?-1:a.date>b.date?1:(+a.vno||0)-(+b.vno||0)||(+a.seq||0)-(+b.seq||0));
  VOUCH=new Map();for(const r of ROWS){let v=VOUCH.get(r.vid);if(!v)VOUCH.set(r.vid,v=[]);v.push(r)}
  S.month=null;
  toast(`${added.toLocaleString()}라인 추가`+(dup?` · 중복라인 ${dup.toLocaleString()}건 제외`:''));
  refresh();
}

const SAMPLES=[['sample/journal_2025_01-11.xlsx','샘플_전표_2025년_1-11월.xlsx'],['sample/journal_2025_12.xlsx','샘플_전표_2025년_12월.xlsx']];
async function loadSample(){
  if(ROWS.length&&!confirm('불러온 전표를 지우고 샘플 데이터로 바꿀까요?'))return;
  try{
    const files=await Promise.all(SAMPLES.map(async([u,n])=>{const r=await fetch(u);if(!r.ok)throw new Error(u);return new File([await r.blob()],n)}));
    ROWS=[];VOUCH=new Map();FILES=[];SIGS=new Set();S.selected.clear();S.expanded.clear();
    await loadJournals(files);
  }catch(e){alert('샘플 데이터를 불러오지 못했습니다. (로컬 파일로 열었다면 웹서버/배포 주소에서 실행해 주세요)\n'+e.message)}
}

/* ============================================================
   분석
   ============================================================ */
const pkeyOf=r=>r.vKey+'|'+r.aKey+'|'+r.side;
function gkeyOf(r,mode){return mode+'|'+r.vKey+'|'+r.aKey+'|'+r.side+(mode==='vat'?'|'+r.tpl:'')}

function compute(){
  C={};
  C.months=[...new Set(ROWS.map(r=>r.ym))].sort();
  if(!S.month||!C.months.includes(S.month))S.month=C.months[C.months.length-1]||null;
  const M=S.month;
  C.hist=C.months.filter(m=>m<M);
  C.cur=ROWS.filter(r=>r.ym===M);
  C.prevMonth=C.hist[C.hist.length-1]||null;

  // 거래처별 과거 계정 사용 이력
  C.vendorHist=new Map();
  for(const r of ROWS){if(r.ym>=M)continue;
    let v=C.vendorHist.get(r.vKey);if(!v)C.vendorHist.set(r.vKey,v=new Map());
    const k=r.aKey+'|'+r.side;const e=v.get(k)||{aName:r.aName,side:r.side,n:0};e.n++;v.set(k,e)}
  const dbVendors=new Set(Object.values(DB.patterns).map(p=>p.vKey));

  // 1) 신규 계정패턴
  const pg=new Map();
  for(const r of C.cur){const k=pkeyOf(r);let g=pg.get(k);if(!g)pg.set(k,g={pkey:k,vKey:r.vKey,vName:r.vName,aKey:r.aKey,aName:r.aName,side:r.side,noVendor:r.noVendor,rows:[],amt:0});g.rows.push(r);g.amt+=r.amt}
  C.newPatterns=[];
  for(const g of pg.values()){
    if(DB.patterns[g.pkey])continue;
    const dk=M+'|'+g.pkey;
    g.status=DB.once[dk]?'once':DB.issues[dk]?'issue':'todo';
    g.note=(DB.once[dk]||DB.issues[dk]||{}).note||'';
    g.newVendor=!g.noVendor&&!C.vendorHist.has(g.vKey)&&!dbVendors.has(g.vKey);
    g.seenInHist=!!(C.vendorHist.get(g.vKey)?.get(g.aKey+'|'+g.side));
    C.newPatterns.push(g);
  }
  for(const g of C.newPatterns)g.kind=g.noVendor?'nov':g.newVendor?'newv':'exist';
  const kr={exist:0,nov:1,newv:2};
  C.newPatterns.sort((a,b)=>(a.status==='todo'?0:1)-(b.status==='todo'?0:1)||kr[a.kind]-kr[b.kind]||a.vName.localeCompare(b.vName,'ko'));
  C.todoCount=C.newPatterns.filter(g=>g.status==='todo').length;
  C.newVendorCount=new Set(C.newPatterns.filter(g=>g.newVendor).map(g=>g.vKey)).size;
  C.unseededHist=C.hist.length>0&&Object.keys(DB.patterns).length===0;

  // 2) 고정거래
  C.rec=computeRecurring(DB.settings.mode);
  C.flagCount={miss:0,dup:0,chg:0};
  for(const g of C.rec){if(g.muted)continue;if(g.fMiss)C.flagCount.miss++;if(g.fDup)C.flagCount.dup++;if(g.fChg)C.flagCount.chg++}
  C.check=C.rec.filter(g=>!g.muted&&(g.pinned||g.freq>=DB.settings.monthlyRatio));

  // 3) 중복전표 의심 (같은 달, 라인 구성이 완전히 동일한 전표)
  const sig=new Map();
  for(const vid of new Set(C.cur.map(r=>r.vid))){
    const lines=VOUCH.get(vid);
    const s=lines.map(r=>[r.aKey,r.vKey,r.dr,r.cr].join(':')).sort().join(';');
    if(!sig.has(s))sig.set(s,[]);sig.get(s).push(vid);
  }
  C.dupVouchers=[...sig.values()].filter(a=>a.length>1);
}

function computeRecurring(mode){
  const map=new Map();
  for(const r of ROWS){
    const k=gkeyOf(r,mode);let g=map.get(k);
    if(!g)map.set(k,g={k,vKey:r.vKey,vName:r.vName,aKey:r.aKey,aName:r.aName,side:r.side,noVendor:r.noVendor,tpl:mode==='vat'?r.tpl:'',m:{}});
    const c=g.m[r.ym]||(g.m[r.ym]={n:0,amt:0,rows:[]});c.n++;c.amt+=r.amt;c.rows.push(r);
  }
  const H=C.hist,M=S.month,st=DB.settings,tol=st.tol/100;
  const minM=Math.max(st.minMonths,Math.ceil(H.length*st.minRatio));
  const out=[];
  for(const g of map.values()){
    const fx=DB.fixed[g.k];g.pinned=fx==='pin';g.muted=fx==='mute';
    const hm=H.filter(m=>g.m[m]);g.present=hm.length;g.freq=H.length?hm.length/H.length:0;
    if(!(g.pinned||(H.length&&g.present>=minM)))continue;
    const ns=hm.map(m=>g.m[m].n),as=hm.map(m=>g.m[m].amt);
    g.medN=median(ns);g.maxN=Math.max(0,...ns);g.medAmt=median(as);
    g.first=hm[0]||Object.keys(g.m).sort()[0];
    // 전체 월 기준 통계(그리드 색칠용)
    const am=C.months.filter(m=>g.m[m]);
    g.allMedN=median(am.map(m=>g.m[m].n));g.allMedAmt=median(am.map(m=>g.m[m].amt));g.allFirst=am[0];g.allLast=am[am.length-1];
    g.allFreq=am.length/C.months.length;
    const cur=g.m[M];g.cur=cur;g.prevAmt=C.prevMonth&&g.m[C.prevMonth]?g.m[C.prevMonth].amt:null;
    const monthly=g.pinned||g.freq>=st.monthlyRatio;
    g.fMiss=!cur&&monthly;
    g.fGap=!cur&&!monthly;
    g.fDup=!!cur&&(cur.n>Math.max(g.maxN,1)||(g.medN<=2&&cur.n>g.medN&&hasSameAmt(cur.rows)));
    g.chgPct=cur&&g.medAmt?(cur.amt-g.medAmt)/Math.abs(g.medAmt)*100:null;
    g.prevPct=cur&&g.prevAmt?(cur.amt-g.prevAmt)/Math.abs(g.prevAmt)*100:null;
    g.fChg=!!cur&&g.medAmt!==0&&Math.abs(cur.amt-g.medAmt)>Math.abs(g.medAmt)*tol;
    g.score=(g.fMiss?4:0)+(g.fDup?3:0)+(g.fChg?1:0);
    out.push(g);
  }
  out.sort((a,b)=>b.score-a.score||a.noVendor-b.noVendor||a.vName.localeCompare(b.vName,'ko')||a.aName.localeCompare(b.aName,'ko'));
  return out;
}
function hasSameAmt(rows){const s=new Set();for(const r of rows){if(s.has(r.amt))return true;s.add(r.amt)}return false}
function cellClass(g,m,metric){
  const c=g.m[m];
  if(!c){const exp=g.allFreq>=DB.settings.monthlyRatio||g.pinned;return (m>=g.allFirst&&(m<=g.allLast||exp)&&exp)?'miss':'na'}
  if(metric==='n'){const med=g.allMedN;return (c.n>=med+1&&c.n>=med*1.5)?'dup':''}
  const med=g.allMedAmt;return med&&Math.abs(c.amt-med)>Math.abs(med)*DB.settings.tol/100?'chg':'';
}

/* ============================================================
   의사결정 (정상패턴 / 이번만 허용 / 이슈)
   ============================================================ */
function patternMeta(g){return{vKey:g.vKey,vName:g.vName,aKey:g.aKey,aName:g.aName,side:g.side}}
function decide(pkeys,action,note){
  const M=S.month;
  for(const pk of pkeys){
    const g=C.newPatterns.find(x=>x.pkey===pk);if(!g)continue;
    const dk=M+'|'+pk;delete DB.once[dk];delete DB.issues[dk];
    if(action==='approve')DB.patterns[pk]={...patternMeta(g),source:'review',month:M,at:nowIso(),note:note||''};
    else if(action==='once')DB.once[dk]={...patternMeta(g),month:M,at:nowIso(),note:note||'',amt:g.amt,n:g.rows.length};
    else if(action==='issue')DB.issues[dk]={...patternMeta(g),month:M,at:nowIso(),note:note||'',amt:g.amt,n:g.rows.length};
    DB.log.push({at:nowIso(),month:M,action,pkey:pk,label:`${g.vName} / ${g.aName}(${g.side})`,note:note||''});
  }
  S.selected.clear();saveDB();refresh();
  const lbl={approve:'정상패턴으로 등록',once:'이번만 허용',issue:'이슈로 표시',undo:'결정 취소'}[action];
  toast(`${pkeys.length}건 ${lbl}`);
}
function seedFromHistory(){
  const H=new Set(C.hist);let add=0;
  for(const r of ROWS){if(!H.has(r.ym))continue;const pk=pkeyOf(r);
    const p=DB.patterns[pk];
    if(!p){DB.patterns[pk]={vKey:r.vKey,vName:r.vName,aKey:r.aKey,aName:r.aName,side:r.side,source:'history',month:r.ym,at:nowIso(),note:'',n:1};add++}
    else if(p.source==='history'){p.n=(p.n||0)+1}
  }
  DB.log.push({at:nowIso(),month:S.month,action:'seed',label:`${C.hist[0]}~${C.hist[C.hist.length-1]} 이력 수집`,note:`${add}개 추가`});
  saveDB();refresh();toast(`과거 이력에서 정상패턴 ${add.toLocaleString()}개 수집`);
}
function setFixed(k,val){if(DB.fixed[k]===val)delete DB.fixed[k];else DB.fixed[k]=val;saveDB();refresh()}

/* ============================================================
   렌더링
   ============================================================ */
const TABS=[['summary','요약'],['patterns','신규 계정패턴'],['grid','고정거래 월별'],['check','고정전표 체크'],['db','패턴DB'],['settings','설정']];
function refresh(){compute();renderAll()}
function renderAll(){renderMonthSel();renderStatus();renderTabs();renderView()}
function renderMonthSel(){
  const sel=$('#selMonth');sel.innerHTML=(C.months||[]).slice().reverse().map(m=>`<option ${m===S.month?'selected':''}>${m}</option>`).join('')||'<option>-</option>';
  sel.disabled=!ROWS.length;
}
function renderStatus(){
  if(DEMO){$('#status').innerHTML=ROWS.length?`<b>체험용 샘플 데이터</b> · 가상의 제조회사 2025년 전표 · ${ROWS.length.toLocaleString()}라인 · 정상패턴 <b>${Object.keys(DB.patterns).length.toLocaleString()}</b>개 · 새로고침하면 처음 상태로 돌아갑니다`:'';return}
  const np=Object.keys(DB.patterns).length;
  $('#status').innerHTML=ROWS.length
    ?`불러온 파일 <b>${FILES.map(f=>esc(f.name)).join(', ')}</b> · ${ROWS.length.toLocaleString()}라인 · ${C.months[0]} ~ ${C.months[C.months.length-1]} · 정상패턴 <b>${np.toLocaleString()}</b>개 `+
      (fileHandle?`· DB파일 <b>${esc(fileHandle.name)}</b> `:'')+(dirty?'· <span class="dirty">DB 변경사항 미저장</span>':'')
    :`정상패턴 ${np.toLocaleString()}개 저장됨`+(dirty?' · <span class="dirty">DB 변경사항 미저장</span>':'');
}
function renderTabs(){
  const cnt={patterns:C.todoCount,grid:(C.flagCount?.miss||0)+(C.flagCount?.dup||0),check:C.check?C.check.filter(g=>g.fMiss||g.fChg||g.fDup).length:0};
  $('#tabs').innerHTML=TABS.map(([k,l])=>`<button class="chip ${S.tab===k?'active':''}" data-tab="${k}">${l}${cnt[k]!=null&&ROWS.length?` <span class="cnt ${cnt[k]?'':'zero'}">${cnt[k]}</span>`:''}</button>`).join('');
}
function renderView(){
  const v=$('#view');
  const fid=document.activeElement?.id,pos=document.activeElement?.selectionStart;
  if(!ROWS.length&&S.tab!=='settings'&&S.tab!=='db'){v.innerHTML=emptyView();return}
  v.innerHTML=({summary:vSummary,patterns:vPatterns,grid:vGrid,check:vCheck,db:vDb,settings:vSettings}[S.tab])();
  if(fid){const el=document.getElementById(fid);if(el){el.focus();try{el.setSelectionRange(pos,pos)}catch(e){}}}
}
function emptyView(){
  if(DEMO)return `<div class="card empty"><h2>샘플 데이터를 불러오는 중…</h2><p>가상의 제조회사 2025년 전표를 준비하고 있습니다.</p></div>`;
  return `<div class="card empty"><h2>전표 엑셀을 불러오세요</h2>
  <p>더존 iCUBE 전표출력 엑셀(기표일자·계정과목명·거래처명·적요·차변·대변)을 그대로 넣으면 됩니다. 여러 파일(예: 작년 전체 + 이번 달)을 함께 불러올 수 있고, 데이터는 이 브라우저 안에서만 처리됩니다.</p>
  <div class="toolbar" style="justify-content:center"><button class="btn primary" data-act="sample">샘플 데이터로 체험하기</button>
  <label class="btn">내 전표 엑셀 불러오기<input type="file" multiple accept=".xlsx,.xls,.csv" hidden data-journal></label></div>
  <p class="muted" style="font-size:13px;margin:12px auto 0">샘플은 가상의 제조회사 2025년 전표입니다. 12월에 누락·중복·계정 오류·금액 변동을 일부러 심어 두었습니다.
   · 샘플 엑셀 받기: ${SAMPLES.map(([u,n])=>`<a class="lnk" href="${u}" download="${n}">${n.replace('샘플_전표_','').replace('.xlsx','')}</a>`).join(' / ')}</p>
  <div class="steps">
   <div><b>1. 이력 불러오기</b><span>과거 전표를 불러오면 거래처별 계정 사용 패턴과 고정거래를 자동으로 찾습니다.</span></div>
   <div><b>2. 정상패턴 수집</b><span>검토월 이전 이력을 정상패턴으로 한 번에 등록합니다. (최초 1회)</span></div>
   <div><b>3. 월 결산 검토</b><span>새 패턴은 정상등록 / 이번만 허용 / 이슈로 판단합니다. 고정거래 누락·중복·금액변동을 확인합니다.</span></div>
   <div><b>4. 패턴DB 저장</b><span>판단 결과가 쌓인 패턴DB(JSON)를 저장해 두면 다음 달 검토가 더 정확해집니다.</span></div>
  </div></div>`;
}
function seedBanner(){
  if(!C.hist.length)return `<div class="banner"><div><h3>검토월 이전 이력이 없습니다</h3><p>정상패턴과 고정거래를 판단하려면 검토월보다 이전 달의 전표가 필요합니다. 검토월을 바꾸거나 과거 전표를 함께 불러오세요.</p></div></div>`;
  if(C.unseededHist)return `<div class="banner"><div><h3>정상패턴이 아직 없습니다</h3><p>${C.hist[0]} ~ ${C.hist[C.hist.length-1]} 이력에 나온 거래처·계정 조합을 정상패턴으로 한 번에 등록할 수 있습니다.</p></div><button class="btn primary" data-act="seed">과거 이력으로 정상패턴 수집</button></div>`;
  return '';
}

/* ---- 요약 ---- */
function demoIntro(){
  return `<div class="notice"><span class="tag t-new">SAMPLE</span>
   <span>실무에서 쓰는 ERP 전표출력 엑셀 양식을 그대로 재현하고, 회사명·거래처·금액은 모두 <b>가상의 더미데이터</b>로 채웠습니다.</span>
   <span class="muted">샘플 엑셀: ${SAMPLES.map(([u,n])=>`<a class="lnk" href="${u}" download="${n}">${n.replace('샘플_전표_','').replace('.xlsx','')}</a>`).join(' · ')}</span></div>`;
}
function vSummary(){
  const M=S.month,cur=C.cur,vcount=new Set(cur.map(r=>r.vid)).size,dr=cur.reduce((s,r)=>s+r.dr,0);
  const k=(l,v,unit,alert,tab)=>`<div class="kpi ${alert?'alert':''} ${tab?'link':''}" ${tab?`data-tab="${tab}"`:''}><div class="l">${l}</div><div class="v">${v}<small>${unit||''}</small></div></div>`;
  const mrows=C.months.map(m=>{const rs=ROWS.filter(r=>r.ym===m);return{m,v:new Set(rs.map(r=>r.vid)).size,l:rs.length,dr:rs.reduce((s,r)=>s+r.dr,0)}});
  const maxDr=Math.max(...mrows.map(x=>x.dr),1);
  return (DEMO?demoIntro():'')+seedBanner()+`
  <div class="kpis">
   ${k(M+' 전표',vcount.toLocaleString(),'건')}
   ${k('라인 수',cur.length.toLocaleString(),'')}
   ${k('차변 합계',fmt(dr/1e6),'백만원')}
   ${k('신규 계정패턴 미검토',C.todoCount,'건',C.todoCount>0,'patterns')}
   ${k('신규 거래처',C.newVendorCount,'곳',false,'patterns')}
   ${k('고정거래 누락',C.flagCount.miss,'건',C.flagCount.miss>0,'check')}
   ${k('건수 증가(중복 의심)',C.flagCount.dup,'건',C.flagCount.dup>0,'grid')}
   ${k('금액 변동 ±'+DB.settings.tol+'%',C.flagCount.chg,'건',false,'check')}
   ${k('동일 전표 중복 의심',C.dupVouchers.length,'쌍',C.dupVouchers.length>0)}
  </div>
  <div class="cols2">
   <div class="card"><h2>월별 기표 현황</h2><p class="sub">전표 건수가 평소보다 크게 적거나 많은 달은 기표 누락·중복을 의심해 볼 수 있습니다.</p>
    <div class="tw"><table><thead><tr><th>월</th><th class="n">전표</th><th class="n">라인</th><th class="n">차변 합계</th><th style="width:30%"></th></tr></thead><tbody>
    ${mrows.map(x=>`<tr ${x.m===M?'class="sel"':''}><td class="nowrap">${x.m}</td><td class="n">${x.v.toLocaleString()}</td><td class="n">${x.l.toLocaleString()}</td><td class="n">${fmt(x.dr)}</td><td><div style="height:8px;border-radius:9999px;background:var(--secondary)"><div style="height:8px;border-radius:9999px;background:var(--ink);width:${(x.dr/maxDr*100).toFixed(1)}%"></div></div></td></tr>`).join('')}
    </tbody></table></div></div>
   <div class="card"><h2>동일 전표 중복 의심</h2><p class="sub">${M}에 계정·거래처·금액 구성이 완전히 같은 전표가 2건 이상인 경우입니다.</p>
    ${C.dupVouchers.length?`<div class="tw"><table><thead><tr><th>전표</th><th>거래처</th><th>적요</th><th class="n">금액</th></tr></thead><tbody>
    ${C.dupVouchers.map(arr=>{const l=VOUCH.get(arr[0]);const amt=l.reduce((s,r)=>s+r.dr,0);const v=l.find(r=>!r.noVendor)||l[0];
      return `<tr><td class="nowrap">${arr.map(id=>`<button class="lnk" data-voucher="${esc(id)}">${esc(id.replace('#',' #'))}</button>`).join('<br>')}</td><td>${esc(v.vName)}</td><td class="memo">${esc(l[0].memo)}</td><td class="n">${fmt(amt)}</td></tr>`}).join('')}
    </tbody></table></div>`:'<p class="muted">해당 없음</p>'}
   </div>
  </div>`;
}

/* ---- 신규 계정패턴 ---- */
function vPatterns(){
  const f=S.prFilter,q=S.q.trim().toLowerCase();
  let list=C.newPatterns.filter(g=>f==='all'||(f==='todo'?g.status==='todo':g.status!=='todo'));
  if(S.kind!=='all')list=list.filter(g=>g.kind===S.kind);
  if(q)list=list.filter(g=>(g.vName+' '+g.aName+' '+g.rows.map(r=>r.memo).join(' ')).toLowerCase().includes(q));
  const kc=k=>C.newPatterns.filter(g=>g.kind===k&&(S.prFilter==='all'||(S.prFilter==='todo'?g.status==='todo':g.status!=='todo'))).length;
  const nDone=C.newPatterns.length-C.todoCount;
  const sideTxt=s=>s==='차'?'차변':'대변';
  const rowsHtml=list.map(g=>{
    const known=[...(C.vendorHist.get(g.vKey)?.values()||[])].sort((a,b)=>b.n-a.n);
    const dbKnown=Object.values(DB.patterns).filter(p=>p.vKey===g.vKey&&!C.vendorHist.get(g.vKey)?.has(p.aKey+'|'+p.side));
    const exp=S.expanded.has(g.pkey),sel=S.selected.has(g.pkey);
    const memos=[...new Set(g.rows.map(r=>r.memo))];
    const st={todo:'',once:'<span class="tag t-gray">이번만 허용</span>',issue:'<span class="tag t-miss">이슈</span>'}[g.status];
    return `<tr class="${sel?'sel':''}">
     <td class="c"><input type="checkbox" data-sel="${esc(g.pkey)}" ${sel?'checked':''} ${g.status!=='todo'?'disabled':''}></td>
     <td><b style="color:var(--ink)">${esc(g.vName)}</b><br>${g.newVendor?'<span class="tag t-new">신규 거래처</span>':''}${g.noVendor?'<span class="tag t-gray">거래처 없음</span>':''}${g.seenInHist?'<span class="tag t-gray">이력엔 있으나 미등록</span>':''}</td>
     <td class="nowrap"><b style="color:var(--ink)">${esc(g.aName)}</b> <span class="muted">${sideTxt(g.side)}</span></td>
     <td class="n">${g.rows.length}</td><td class="n">${fmt(g.amt)}</td>
     <td>${known.length?known.slice(0,8).map(e=>`<span class="acc">${esc(e.aName)}·${e.side} ${e.n}</span>`).join('')+(known.length>8?`<span class="muted">외 ${known.length-8}</span>`:''):''}${dbKnown.slice(0,5).map(p=>`<span class="acc">${esc(p.aName)}·${p.side}</span>`).join('')}${!known.length&&!dbKnown.length?'<span class="muted">이력 없음</span>':''}</td>
     <td class="memo">${esc(memos[0])}${memos.length>1?` <span class="muted">외 ${memos.length-1}</span>`:''}</td>
     <td class="nowrap">${st}${g.note?`<div class="muted" style="font-size:12px;max-width:180px;white-space:normal">${esc(g.note)}</div>`:''}
       ${g.status==='todo'
         ?`<button class="btn sm" data-dec="approve" data-pk="${esc(g.pkey)}">정상등록</button> <button class="btn sm" data-dec="once" data-pk="${esc(g.pkey)}">이번만</button> <button class="btn sm ghost" data-dec="issue" data-pk="${esc(g.pkey)}">이슈</button>`
         :`<button class="btn sm ghost" data-dec="undo" data-pk="${esc(g.pkey)}">취소</button> ${g.status==='once'?`<button class="btn sm" data-dec="approve" data-pk="${esc(g.pkey)}">정상등록</button>`:''}`}
       <button class="btn sm ghost" data-exp="${esc(g.pkey)}">${exp?'접기':'전표 보기'}</button></td>
    </tr>${exp?`<tr class="sub"><td colspan="8">${linesTable(g.rows,true)}</td></tr>`:''}`;
  }).join('');
  return seedBanner()+`<div class="card">
   <h2>신규 계정패턴 검토 · ${S.month}</h2>
   <p class="sub">정상패턴 DB에 없는 <b>거래처 × 계정 × 차/대</b> 조합입니다. 오른쪽 "기존 계정"과 비교해 정상이면 등록하고, 이번 달만 예외면 이번만 허용, 잘못된 기표면 이슈로 남기세요.</p>
   <div class="toolbar">
    <div class="seg">${[['todo','미검토 '+C.todoCount],['done','이번 달 결정 '+nDone],['all','전체']].map(([k,l])=>`<button class="${S.prFilter===k?'on':''}" data-prf="${k}">${l}</button>`).join('')}</div>
    <div class="seg">${[['all','전체'],['exist','기존 거래처·새 계정 '+kc('exist')],['nov','거래처 없음 '+kc('nov')],['newv','신규 거래처 '+kc('newv')]].map(([k,l])=>`<button class="${S.kind===k?'on':''}" data-kind="${k}">${l}</button>`).join('')}</div>
    <input type="search" id="q" placeholder="거래처·계정·적요 검색" value="${esc(S.q)}">
    <span class="sp"></span>
    ${S.prFilter!=='done'&&list.some(g=>g.status==='todo')?`<button class="btn sm ghost" data-act="selall">미검토 전체 선택</button>`:''}
   </div>
   ${list.length?`<div class="tw"><table><thead><tr><th></th><th>거래처</th><th>계정</th><th class="n">라인</th><th class="n">금액</th><th>이 거래처 기존 계정 (차/대 · 건수)</th><th>적요</th><th>판단</th></tr></thead><tbody>${rowsHtml}</tbody></table></div>`
     :`<p class="muted" style="padding:24px 0;text-align:center">${S.prFilter==='todo'?'미검토 신규 패턴이 없습니다 ✓':'해당 항목이 없습니다'}</p>`}
   ${S.selected.size?`<div class="bulk">선택 ${S.selected.size}건 <button class="btn primary" data-bulk="approve">정상패턴 등록</button><button class="btn" data-bulk="once">이번만 허용</button><button class="btn ghost" style="color:#fff" data-bulk="clear">선택 해제</button></div>`:''}
  </div>`;
}
function linesTable(rows,showVoucher){
  return `<table><thead><tr><th>일자</th><th>전표</th><th>계정</th><th>거래처</th><th>적요</th><th class="n">차변</th><th class="n">대변</th></tr></thead><tbody>
  ${rows.map(r=>`<tr><td class="nowrap">${r.date}</td><td class="nowrap">${showVoucher?`<button class="lnk" data-voucher="${esc(r.vid)}">#${esc(r.vno)}</button>`:'#'+esc(r.vno)}</td><td class="nowrap">${esc(r.aName)}</td><td>${esc(r.vName)}</td><td class="memo">${esc(r.memo)}</td><td class="n">${r.dr?fmt(r.dr):''}</td><td class="n">${r.cr?fmt(r.cr):''}</td></tr>`).join('')}
  </tbody></table>`;
}

/* ---- 고정거래 월별 ---- */
function filteredRec(){
  const q=S.q.trim().toLowerCase();
  return C.rec.filter(g=>{
    if(g.muted&&!S.showMuted)return false;
    if(S.onlyNoVendor&&!g.noVendor)return false;
    if(S.onlyFlag&&!(g.fMiss||g.fDup||(S.metric!=='n'&&g.fChg)))return false;
    if(q&&!(g.vName+' '+g.aName+' '+g.tpl).toLowerCase().includes(q))return false;
    return true;
  });
}
function vGrid(){
  const mode=DB.settings.mode,metric=S.metric,list=filteredRec(),M=S.month;
  const unit=metric==='k'?1000:1;
  const val=c=>!c?'':metric==='n'?c.n:fmt(c.amt/unit);
  const flagTags=g=>(g.fMiss?'<span class="tag t-miss">누락</span>':'')+(g.fDup?'<span class="tag t-dup">건수↑</span>':'')+(g.fChg?`<span class="tag t-chg">${pct(g.chgPct)}</span>`:'')+(g.fGap?'<span class="tag t-gray">미발생</span>':'')+(!g.fMiss&&!g.fDup&&!g.fChg&&!g.fGap?'<span class="tag t-ok">정상</span>':'');
  const shown=list.slice(0,600);
  return seedBanner()+`<div class="card">
   <h2>고정거래 월별 ${metric==='n'?'건수':'금액'}</h2>
   <p class="sub">검토월 이전 ${C.hist.length}개월 중 ${Math.max(DB.settings.minMonths,Math.ceil(C.hist.length*DB.settings.minRatio))}개월 이상 발생한 반복거래입니다. 셀을 누르면 해당 월 라인이 열립니다.</p>
   <div class="toolbar">
    <div class="seg">${[['n','건수'],['a','금액(원)'],['k','금액(천원)']].map(([k,l])=>`<button class="${metric===k?'on':''}" data-metric="${k}">${l}</button>`).join('')}</div>
    <div class="seg">${[['va','거래처+계정'],['vat','거래처+계정+적요']].map(([k,l])=>`<button class="${mode===k?'on':''}" data-mode="${k}">${l}</button>`).join('')}</div>
    <input type="search" id="q" placeholder="거래처·계정·적요 검색" value="${esc(S.q)}">
    <label class="ck"><input type="checkbox" data-opt="onlyFlag" ${S.onlyFlag?'checked':''}>이상 항목만</label>
    <label class="ck"><input type="checkbox" data-opt="onlyNoVendor" ${S.onlyNoVendor?'checked':''}>거래처 없음(결산성)만</label>
    <label class="ck"><input type="checkbox" data-opt="showMuted" ${S.showMuted?'checked':''}>제외 항목 표시</label>
   </div>
   <div class="legend"><span class="tag t-miss">누락</span>매월 발생하던 거래가 없음 ${metric==='n'?'<span class="tag t-dup">건수↑</span>평소 건수의 1.5배 이상':`<span class="tag t-chg">변동</span>중앙값 대비 ±${DB.settings.tol}% 초과`} <span>· 굵은 테두리 열 = 검토월 · ${list.length.toLocaleString()}개 항목${list.length>600?' (상위 600개 표시)':''}</span></div>
   <div class="tw grid"><table><thead><tr>
    <th class="stick">거래처</th><th>계정</th>${mode==='vat'?'<th>적요 유형</th>':''}
    ${C.months.map(m=>`<th class="n ${m===M?'rm':''}">${m.slice(2).replace('-','.')}</th>`).join('')}
    <th class="c">발생</th><th>${M} 판정</th><th></th></tr></thead><tbody>
   ${shown.map(g=>`<tr>
     <td class="stick"><b style="color:var(--ink)">${esc(g.vName)}</b>${g.pinned?' <span class="tag t-pin">고정</span>':''}${g.muted?' <span class="tag t-gray">제외</span>':''}</td>
     <td class="nowrap">${esc(g.aName)} <span class="muted">${g.side}</span></td>${mode==='vat'?`<td class="memo">${esc(g.tpl)}</td>`:''}
     ${C.months.map(m=>`<td class="cell ${cellClass(g,m,metric==='n'?'n':'a')} ${m===M?'rm':''}" data-cell="${esc(g.k)}" data-m="${m}">${val(g.m[m])}</td>`).join('')}
     <td class="c muted">${Object.keys(g.m).length}/${C.months.length}</td>
     <td class="nowrap">${flagTags(g)}</td>
     <td class="nowrap"><button class="btn sm ghost" data-fix="pin" data-k="${esc(g.k)}" title="매월 반드시 있어야 하는 항목으로 지정">${g.pinned?'고정 해제':'고정'}</button><button class="btn sm ghost" data-fix="mute" data-k="${esc(g.k)}" title="검토 대상에서 제외">${g.muted?'제외 해제':'제외'}</button></td>
   </tr>`).join('')}
   </tbody></table></div></div>`;
}

/* ---- 고정전표 체크 ---- */
function vCheck(){
  const M=S.month,f=S.ckFilter,q=S.q.trim().toLowerCase();
  let list=C.check;
  if(f==='issue')list=list.filter(g=>g.fMiss||g.fDup||g.fChg);
  if(f==='close')list=list.filter(g=>g.noVendor);
  if(q)list=list.filter(g=>(g.vName+' '+g.aName+' '+g.tpl).toLowerCase().includes(q));
  const issues=C.check.filter(g=>g.fMiss||g.fDup||g.fChg).length;
  const stat=g=>g.fMiss?'<span class="tag t-miss">누락</span>':(g.fDup?'<span class="tag t-dup">중복 의심</span>':'')+(g.fChg?'<span class="tag t-chg">금액 변동</span>':'')+(!g.fDup&&!g.fChg?'<span class="tag t-ok">정상</span>':'');
  return seedBanner()+`<div class="card">
   <h2>고정전표 체크리스트 · ${M}</h2>
   <p class="sub">이력의 ${Math.round(DB.settings.monthlyRatio*100)}% 이상 달에 발생한 항목과 "고정"으로 지정한 항목입니다. 감가상각·급여·충당금·미수수익처럼 매월 기표해야 하는 결산전표의 누락과 금액 변동을 확인하세요.</p>
   <div class="toolbar">
    <div class="seg">${[['all','전체 '+C.check.length],['issue','이상 '+issues],['close','거래처 없음(결산성)']].map(([k,l])=>`<button class="${f===k?'on':''}" data-ckf="${k}">${l}</button>`).join('')}</div>
    <div class="seg">${[['va','거래처+계정'],['vat','거래처+계정+적요']].map(([k,l])=>`<button class="${DB.settings.mode===k?'on':''}" data-mode="${k}">${l}</button>`).join('')}</div>
    <input type="search" id="q" placeholder="거래처·계정·적요 검색" value="${esc(S.q)}">
   </div>
   <div class="tw"><table><thead><tr><th>판정</th><th>거래처</th><th>계정</th>${DB.settings.mode==='vat'?'<th>적요 유형</th>':''}<th class="n">이력 건수(중앙)</th><th class="n">${M} 건수</th><th class="n">전월 금액</th><th class="n">${M} 금액</th><th class="n">전월比</th><th class="n">이력 중앙값</th><th class="n">중앙값比</th><th></th></tr></thead><tbody>
   ${list.map(g=>`<tr>
    <td class="nowrap">${stat(g)}</td>
    <td><b style="color:var(--ink)">${esc(g.vName)}</b>${g.pinned?' <span class="tag t-pin">고정</span>':''}</td>
    <td class="nowrap">${esc(g.aName)} <span class="muted">${g.side}</span></td>${DB.settings.mode==='vat'?`<td class="memo">${esc(g.tpl)}</td>`:''}
    <td class="n">${g.medN}</td><td class="n">${g.cur?g.cur.n:0}</td>
    <td class="n">${fmt(g.prevAmt)}</td><td class="n"><b style="color:var(--ink)">${g.cur?fmt(g.cur.amt):'-'}</b></td>
    <td class="n">${pct(g.prevPct)}</td><td class="n">${fmt(g.medAmt)}</td><td class="n" style="${g.fChg?'color:var(--info);font-weight:700':''}">${pct(g.chgPct)}</td>
    <td class="nowrap"><button class="btn sm ghost" data-cell="${esc(g.k)}" data-m="${M}">라인</button><button class="btn sm ghost" data-fix="mute" data-k="${esc(g.k)}">제외</button></td>
   </tr>`).join('')||`<tr><td colspan="12" class="muted c" style="padding:24px">해당 항목이 없습니다</td></tr>`}
   </tbody></table></div></div>`;
}

/* ---- 패턴DB ---- */
function vDb(){
  const q=S.dbq.trim().toLowerCase();
  let pats=Object.entries(DB.patterns);
  if(q)pats=pats.filter(([k,p])=>(p.vName+' '+p.aName+' '+(p.note||'')).toLowerCase().includes(q));
  pats.sort((a,b)=>a[1].vName.localeCompare(b[1].vName,'ko')||a[1].aName.localeCompare(b[1].aName,'ko'));
  const nHist=Object.values(DB.patterns).filter(p=>p.source==='history').length;
  const dec=[...Object.entries(DB.once).map(([k,v])=>['once',k,v]),...Object.entries(DB.issues).map(([k,v])=>['issue',k,v])].sort((a,b)=>b[2].month.localeCompare(a[2].month));
  const fixed=Object.entries(DB.fixed);
  return `<div class="kpis">
    <div class="kpi"><div class="l">정상패턴 전체</div><div class="v">${Object.keys(DB.patterns).length.toLocaleString()}</div></div>
    <div class="kpi"><div class="l">이력 수집</div><div class="v">${nHist.toLocaleString()}</div></div>
    <div class="kpi"><div class="l">검토 후 등록</div><div class="v">${(Object.keys(DB.patterns).length-nHist).toLocaleString()}</div></div>
    <div class="kpi"><div class="l">이번만 허용 / 이슈</div><div class="v">${Object.keys(DB.once).length} / ${Object.keys(DB.issues).length}</div></div>
   </div>
   <div class="card"><h2>정상패턴</h2><p class="sub">거래처 × 계정 × 차/대 조합. 삭제하면 다음 검토 때 다시 신규 패턴으로 잡힙니다.</p>
    <div class="toolbar"><input type="search" id="dbq" placeholder="거래처·계정 검색" value="${esc(S.dbq)}"><span class="muted">${pats.length.toLocaleString()}개${pats.length>500?' (500개만 표시 · 검색으로 좁히세요)':''}</span></div>
    <div class="tw"><table><thead><tr><th>거래처</th><th>계정</th><th>차/대</th><th>출처</th><th>기준월</th><th>등록일시</th><th>메모</th><th></th></tr></thead><tbody>
    ${pats.slice(0,500).map(([k,p])=>`<tr><td>${esc(p.vName)}</td><td>${esc(p.aName)}</td><td>${p.side}</td><td>${p.source==='history'?'<span class="tag t-gray">이력</span>':'<span class="tag t-ok">검토 승인</span>'}</td><td class="nowrap">${esc(p.month||'')}</td><td class="nowrap muted">${esc(p.at||'')}</td><td class="memo">${esc(p.note||'')}</td><td><button class="btn sm ghost" data-delpat="${esc(k)}">삭제</button></td></tr>`).join('')||'<tr><td colspan="8" class="muted c" style="padding:24px">정상패턴이 없습니다</td></tr>'}
    </tbody></table></div></div>
   <div class="cols2">
    <div class="card"><h2>이번만 허용 · 이슈 기록</h2><p class="sub">월별 판단 이력입니다.</p>
     <div class="tw"><table><thead><tr><th>월</th><th>구분</th><th>거래처 / 계정</th><th class="n">금액</th><th>메모</th></tr></thead><tbody>
     ${dec.map(([t,k,v])=>`<tr><td class="nowrap">${v.month}</td><td>${t==='once'?'<span class="tag t-gray">이번만</span>':'<span class="tag t-miss">이슈</span>'}</td><td>${esc(v.vName)} / ${esc(v.aName)}(${v.side})</td><td class="n">${fmt(v.amt)}</td><td class="memo">${esc(v.note)}</td></tr>`).join('')||'<tr><td colspan="5" class="muted c" style="padding:16px">기록 없음</td></tr>'}
     </tbody></table></div></div>
    <div class="card"><h2>고정 / 제외 지정 항목</h2><p class="sub">고정거래 탭에서 지정한 항목입니다.</p>
     <div class="tw"><table><thead><tr><th>구분</th><th>항목</th><th></th></tr></thead><tbody>
     ${fixed.map(([k,v])=>{const p=k.split('|');return `<tr><td>${v==='pin'?'<span class="tag t-pin">고정</span>':'<span class="tag t-gray">제외</span>'}</td><td>${esc(p[1].slice(1)||'(거래처 없음)')} / ${esc(p[2])} ${esc(p[3])}${p[4]?` · <span class="muted">${esc(p[4])}</span>`:''}</td><td><button class="btn sm ghost" data-unfix="${esc(k)}">해제</button></td></tr>`}).join('')||'<tr><td colspan="3" class="muted c" style="padding:16px">없음</td></tr>'}
     </tbody></table></div></div>
   </div>`;
}

/* ---- 설정 ---- */
function vSettings(){
  const s=DB.settings;
  return `<div class="card"><h2>판정 기준</h2><p class="sub">값을 바꾸면 바로 다시 계산됩니다. 설정은 패턴DB에 함께 저장됩니다.</p>
   <div class="form">
    <div><label>고정거래 최소 발생 월수</label><input type="number" min="1" data-set="minMonths" value="${s.minMonths}"><div class="hint">이력에서 이 개월 수 이상 나오면 반복거래로 봅니다.</div></div>
    <div><label>고정거래 최소 발생 비율</label><input type="number" min="0" max="1" step="0.05" data-set="minRatio" value="${s.minRatio}"><div class="hint">이력 개월 수 × 비율 (예: 11개월 × 0.5 = 6개월). 위 월수와 비교해 큰 쪽을 씁니다.</div></div>
    <div><label>매월 발생(누락 판정) 비율</label><input type="number" min="0" max="1" step="0.05" data-set="monthlyRatio" value="${s.monthlyRatio}"><div class="hint">이 비율 이상 달에 있던 항목이 검토월에 없으면 "누락"으로 표시합니다.</div></div>
    <div><label>금액 변동 허용폭 (%)</label><input type="number" min="0" step="1" data-set="tol" value="${s.tol}"><div class="hint">이력 중앙값 대비 이 비율을 넘으면 "금액 변동"으로 표시합니다.</div></div>
   </div></div>
   <div class="card"><h2>정상패턴 관리</h2>
    <p class="sub">검토월(${esc(S.month||'-')}) 이전의 모든 이력을 정상패턴으로 추가합니다. 이미 있는 패턴은 그대로 둡니다. 과거 데이터를 추가로 불러왔을 때도 다시 실행할 수 있습니다.</p>
    <div class="toolbar"><button class="btn primary" data-act="seed" ${C.hist&&C.hist.length?'':'disabled'}>과거 이력으로 정상패턴 수집</button>
    <button class="btn" data-act="resetdb">패턴DB 초기화</button></div>
    <p class="muted" style="font-size:12px;margin:8px 0 0">${DEMO?'체험판에서는 판단 결과가 저장되지 않으며, 새로고침하면 처음 상태로 돌아갑니다.':''}</p>
    <p class="muted real-only" style="font-size:12px;margin:8px 0 0">패턴DB는 이 브라우저에 자동 저장되지만, 브라우저 데이터가 지워지면 사라질 수 있습니다. "패턴DB 저장"으로 JSON 파일을 Dropbox 등에 보관하세요. Chrome/Edge에서는 한 번 저장 위치를 고르면 이후 변경 시 같은 파일에 자동 저장됩니다.</p>
   </div>`;
}

/* ---- 모달 ---- */
function openModal(html){$('#modal').innerHTML=html;$('#scrim').classList.add('open')}
function closeModal(){$('#scrim').classList.remove('open')}
function showVoucher(vid){
  const l=VOUCH.get(vid)||[];const dr=l.reduce((s,r)=>s+r.dr,0),cr=l.reduce((s,r)=>s+r.cr,0);
  openModal(`<div class="hd"><div><h3>전표 ${esc(vid.replace('#',' · #'))}</h3><p>${l.length}라인 · 차변 ${fmt(dr)} · 대변 ${fmt(cr)}</p></div><button class="x" data-close>✕</button></div>
   <div class="tw">${linesTable(l,false)}</div>`);
}
function showCell(k,m){
  const g=C.rec.find(x=>x.k===k);if(!g)return;const c=g.m[m];
  openModal(`<div class="hd"><div><h3>${esc(g.vName)} · ${esc(g.aName)} (${g.side})</h3><p>${m} · ${c?`${c.n}건 · ${fmt(c.amt)}원`:'발생 없음'} · 이력 중앙값 ${g.medN}건 / ${fmt(g.medAmt)}원${g.tpl?' · 적요유형 '+esc(g.tpl):''}</p></div><button class="x" data-close>✕</button></div>
   ${c?`<div class="tw">${linesTable(c.rows,true)}</div>`:'<p class="muted">이 달에는 해당 거래가 없습니다. 기표 누락인지, 거래가 종료됐는지 확인하세요. 종료된 거래라면 "제외"로 지정하면 다음부터 표시되지 않습니다.</p>'}`);
}
function askNote(title,cb){
  openModal(`<div class="hd"><div><h3>${esc(title)}</h3><p>판단 근거를 남겨두면 다음 검토 때 참고할 수 있습니다. (선택)</p></div><button class="x" data-close>✕</button></div>
   <textarea id="noteText" placeholder="예: 신규 계약에 따른 정상 계정 / 계정 오류로 수정 요청"></textarea>
   <div class="toolbar" style="margin-top:16px;justify-content:flex-end"><button class="btn" data-close>취소</button><button class="btn primary" id="noteOk">확인</button></div>`);
  setTimeout(()=>$('#noteText').focus(),30);
  $('#noteOk').onclick=()=>{const v=$('#noteText').value.trim();closeModal();cb(v)};
}

/* ---- 엑셀 내보내기 ---- */
function exportXlsx(){
  if(!ROWS.length){toast('불러온 전표가 없습니다');return}
  const M=S.month,wb=XLSX.utils.book_new();
  const stName={todo:'미검토',once:'이번만 허용',issue:'이슈'};
  const add=(name,data)=>XLSX.utils.book_append_sheet(wb,XLSX.utils.json_to_sheet(data),name);
  add('신규계정패턴',C.newPatterns.map(g=>({판단:stName[g.status],메모:g.note,거래처:g.vName,신규거래처:g.newVendor?'Y':'',계정:g.aName,차대:g.side,라인수:g.rows.length,금액:g.amt,
    기존계정:[...(C.vendorHist.get(g.vKey)?.values()||[])].map(e=>`${e.aName}(${e.side})`).join(', '),적요:[...new Set(g.rows.map(r=>r.memo))].join(' / '),전표:[...new Set(g.rows.map(r=>r.vid))].join(', ')})));
  add('고정전표체크',C.check.map(g=>({판정:g.fMiss?'누락':[g.fDup?'중복의심':'',g.fChg?'금액변동':''].filter(Boolean).join(',')||'정상',거래처:g.vName,계정:g.aName,차대:g.side,적요유형:g.tpl,
    이력건수중앙:g.medN,검토월건수:g.cur?g.cur.n:0,전월금액:g.prevAmt,검토월금액:g.cur?g.cur.amt:null,전월비:g.prevPct==null?null:+g.prevPct.toFixed(1),이력중앙값:g.medAmt,중앙값비:g.chgPct==null?null:+g.chgPct.toFixed(1)})));
  const grid=metric=>C.rec.filter(g=>!g.muted).map(g=>{const o={거래처:g.vName,계정:g.aName,차대:g.side};if(g.tpl)o.적요유형=g.tpl;for(const m of C.months){const c=g.m[m];o[m]=c?(metric==='n'?c.n:c.amt):null}
    o.판정=g.fMiss?'누락':[g.fDup?'건수증가':'',g.fChg?'금액변동':''].filter(Boolean).join(',')||'';return o});
  add('월별건수',grid('n'));add('월별금액',grid('a'));
  add('중복전표의심',C.dupVouchers.flatMap((arr,i)=>arr.flatMap(id=>VOUCH.get(id).map(r=>({그룹:i+1,전표:id,계정:r.aName,거래처:r.vName,적요:r.memo,차변:r.dr,대변:r.cr})))));
  XLSX.writeFile(wb,`전표검토결과_${M}.xlsx`);
}

/* ============================================================
   이벤트
   ============================================================ */
$('#fJournal').addEventListener('change',e=>{if(e.target.files.length)loadJournals([...e.target.files]);e.target.value=''});
$('#fDb').addEventListener('change',async e=>{const f=e.target.files[0];if(f)applyDbText(await f.text(),f.name);e.target.value=''});
$('#bLoadDb').onclick=openDbFile;
$('#bSaveDb').onclick=saveDbFile;
$('#bExport').onclick=exportXlsx;
$('#selMonth').onchange=e=>{S.month=e.target.value;S.selected.clear();S.expanded.clear();refresh()};
$('#scrim').addEventListener('click',e=>{if(e.target.id==='scrim'||e.target.closest('[data-close]'))closeModal()});
document.addEventListener('keydown',e=>{if(e.key==='Escape')closeModal()});
document.addEventListener('change',e=>{
  const t=e.target;
  if(t.matches('[data-journal]')&&t.files.length){loadJournals([...t.files])}
  else if(t.dataset.sel){t.checked?S.selected.add(t.dataset.sel):S.selected.delete(t.dataset.sel);renderView()}
  else if(t.dataset.opt){S[t.dataset.opt]=t.checked;renderView()}
  else if(t.dataset.set){const v=parseFloat(t.value);if(!isNaN(v)){DB.settings[t.dataset.set]=v;saveDB();compute();renderTabs();renderStatus()}}
});
let qTimer;
document.addEventListener('input',e=>{
  const t=e.target;
  if(t.id==='q'){S.q=t.value;clearTimeout(qTimer);qTimer=setTimeout(renderView,150)}
  if(t.id==='dbq'){S.dbq=t.value;clearTimeout(qTimer);qTimer=setTimeout(renderView,150)}
});
document.addEventListener('click',e=>{
  const t=e.target.closest('button,[data-tab],[data-cell]');if(!t||t.closest('#scrim')&&!t.dataset.voucher&&!t.dataset.close)return;
  const d=t.dataset;
  if(d.tab){S.tab=d.tab;S.q='';renderTabs();renderView();window.scrollTo(0,0)}
  else if(d.voucher)showVoucher(d.voucher);
  else if(d.cell)showCell(d.cell,d.m);
  else if(d.metric){S.metric=d.metric;renderView()}
  else if(d.mode){DB.settings.mode=d.mode;saveDB();refresh()}
  else if(d.prf){S.prFilter=d.prf;S.selected.clear();renderView()}
  else if(d.kind){S.kind=d.kind;S.selected.clear();renderView()}
  else if(d.ckf){S.ckFilter=d.ckf;renderView()}
  else if(d.exp){S.expanded.has(d.exp)?S.expanded.delete(d.exp):S.expanded.add(d.exp);renderView()}
  else if(d.dec){
    const pk=d.pk;
    if(d.dec==='undo'){const dk=S.month+'|'+pk;delete DB.once[dk];delete DB.issues[dk];DB.log.push({at:nowIso(),month:S.month,action:'undo',pkey:pk});saveDB();refresh();toast('결정 취소')}
    else if(d.dec==='issue')askNote('이슈로 표시',n=>decide([pk],'issue',n));
    else decide([pk],d.dec,'');
  }
  else if(d.bulk){
    if(d.bulk==='clear'){S.selected.clear();renderView()}
    else decide([...S.selected],d.bulk,'');
  }
  else if(d.act==='sample')loadSample();
  else if(d.act==='selall'){C.newPatterns.filter(g=>g.status==='todo'&&(S.kind==='all'||g.kind===S.kind)).forEach(g=>S.selected.add(g.pkey));renderView()}
  else if(d.act==='seed'){if(confirm(`${C.hist[0]} ~ ${C.hist[C.hist.length-1]} 전표의 모든 거래처·계정 조합을 정상패턴으로 등록합니다.\n(과거 이력에 오류 기표가 있었다면 그것도 정상으로 등록되니, 이후 패턴DB 탭에서 정리할 수 있습니다.)`))seedFromHistory()}
  else if(d.act==='resetdb'){if(confirm('정상패턴·판단기록·고정/제외 지정을 모두 지웁니다. 계속할까요?\n(먼저 "패턴DB 저장"으로 백업하는 것을 권장합니다)')){const st=DB.settings;DB=defaultDB();DB.settings=st;saveDB();refresh()}}
  else if(d.fix)setFixed(d.k,d.fix);
  else if(d.unfix){delete DB.fixed[d.unfix];saveDB();refresh()}
  else if(d.delpat){delete DB.patterns[d.delpat];saveDB();refresh()}
});
if(!DEMO)window.addEventListener('beforeunload',e=>{if(dirty&&!fileHandle){e.preventDefault();e.returnValue=''}});
if(DEMO){
  document.body.classList.add('demo');refresh();
  loadSample().then(()=>{if(ROWS.length&&!Object.keys(DB.patterns).length)seedFromHistory()});
}else refresh();
