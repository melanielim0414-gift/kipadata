"use strict";

const state={files:[],activeFile:0,activeSheet:"",page:1,pageSize:30,search:"",chart:"line",rows:[],columns:[],types:{}};
const $=s=>document.querySelector(s), $$=s=>[...document.querySelectorAll(s)];
const colors=["#6ee7d2","#53a7ff","#f7c86b","#bb8cff","#ff7b7b","#8fd36d"];

$("#folderButton").onclick=$("#emptyFolderButton").onclick=chooseFolder;
$("#folderInput").onchange=e=>loadFiles([...e.target.files]);
$("#themeButton").onclick=()=>document.documentElement.classList.toggle("light");
$("#sheetSelect").onchange=e=>selectSheet(e.target.value);
$("#xSelect").onchange=renderAll;
$("#metricSelect").onchange=renderAll;
$("#aggregateSelect").onchange=renderAll;
$("#searchInput").oninput=e=>{state.search=e.target.value.trim().toLowerCase();state.page=1;renderTable()};
$("#prevPage").onclick=()=>{state.page--;renderTable()};
$("#nextPage").onclick=()=>{state.page++;renderTable()};
$("#exportButton").onclick=exportSummary;
$("#chartTabs").onclick=e=>{if(e.target.dataset.chart){state.chart=e.target.dataset.chart;$$('.chart-tabs button').forEach(b=>b.classList.toggle('active',b===e.target));renderCharts()}};
window.addEventListener("resize",debounce(renderCharts,120));

async function chooseFolder(){
  if(window.showDirectoryPicker){
    try{const dir=await showDirectoryPicker({mode:"read"});const files=[];for await(const entry of dir.values())if(entry.kind==="file"&&/\.(xlsx|csv|tsv)$/i.test(entry.name))files.push(await entry.getFile());await loadFiles(files);return}catch(e){if(e.name!=="AbortError")showError("폴더를 읽지 못했습니다.")}
  }
  $("#folderInput").click();
}

async function loadFiles(files){
  const selected=files.filter(f=>/\.(xlsx|csv|tsv)$/i.test(f.name));
  if(!selected.length)return showError("지원되는 엑셀 또는 CSV 파일이 없습니다.");
  setBusy(true);const parsed=[];
  for(const file of selected){
    try{parsed.push(await parseFile(file))}catch(error){console.error(error);parsed.push({name:file.name,size:file.size,error:"파일을 해석하지 못했습니다."})}
  }
  state.files=parsed;state.activeFile=Math.max(0,parsed.findIndex(f=>!f.error));renderFileList();
  if(parsed.some(f=>!f.error))selectFile(state.activeFile);else showError("선택한 파일을 해석하지 못했습니다.");
  setBusy(false);
}

async function parseFile(file){
  const ext=file.name.split('.').pop().toLowerCase();
  if(ext==="csv"||ext==="tsv"){const text=await readTextSmart(file);return{name:file.name,size:file.size,sheets:{"데이터":parseDelimited(text,ext==="tsv"?'\t':detectDelimiter(text))}}}
  const buffer=await file.arrayBuffer();return{name:file.name,size:file.size,sheets:await parseXlsx(buffer)};
}

async function readTextSmart(file){
  const buf=await file.arrayBuffer();let text=new TextDecoder("utf-8",{fatal:false}).decode(buf);
  if((text.match(/�/g)||[]).length>2)try{text=new TextDecoder("euc-kr").decode(buf)}catch{}return text.replace(/^\uFEFF/,"");
}
function detectDelimiter(text){const first=text.split(/\r?\n/,1)[0];return[',',';','\t','|'].sort((a,b)=>first.split(b).length-first.split(a).length)[0]}
function parseDelimited(text,delimiter){
  const rows=[];let row=[],cell="",quoted=false;
  for(let i=0;i<text.length;i++){const c=text[i],n=text[i+1];if(c==='"'&&quoted&&n==='"'){cell+='"';i++}else if(c==='"')quoted=!quoted;else if(c===delimiter&&!quoted){row.push(cell);cell=""}else if((c==='\n'||(c==='\r'&&n==='\n'))&&!quoted){row.push(cell);rows.push(row);row=[];cell="";if(c==='\r')i++}else cell+=c}
  if(cell||row.length){row.push(cell);rows.push(row)}return normalizeMatrix(rows);
}

async function parseXlsx(buffer){
  const zip=await JSZip.loadAsync(buffer),parser=new DOMParser(),shared=[],styles={dateStyles:new Set()};
  if(zip.file("xl/sharedStrings.xml")){const doc=parser.parseFromString(await zip.file("xl/sharedStrings.xml").async("text"),"text/xml");doc.querySelectorAll("si").forEach(si=>shared.push([...si.querySelectorAll("t")].map(t=>t.textContent).join("")))}
  if(zip.file("xl/styles.xml")){const doc=parser.parseFromString(await zip.file("xl/styles.xml").async("text"),"text/xml"),custom={};doc.querySelectorAll("numFmt").forEach(n=>custom[n.getAttribute("numFmtId")]=n.getAttribute("formatCode")||"");doc.querySelectorAll("cellXfs xf").forEach((x,i)=>{const id=+x.getAttribute("numFmtId");if([14,15,16,17,18,19,20,21,22,45,46,47].includes(id)||/[ymdhis]/i.test(custom[id]||""))styles.dateStyles.add(i)})}
  const wb=parser.parseFromString(await zip.file("xl/workbook.xml").async("text"),"text/xml"),rels=parser.parseFromString(await zip.file("xl/_rels/workbook.xml.rels").async("text"),"text/xml"),relMap={};
  rels.querySelectorAll("Relationship").forEach(r=>relMap[r.getAttribute("Id")]=r.getAttribute("Target"));const sheets={};
  for(const s of wb.querySelectorAll("sheet")){const name=s.getAttribute("name"),rid=s.getAttribute("r:id"),target=relMap[rid]||"",path=target.startsWith('/')?target.slice(1):`xl/${target.replace(/^\.\//,'')}`;if(!zip.file(path))continue;const doc=parser.parseFromString(await zip.file(path).async("text"),"text/xml"),matrix=[];
    doc.querySelectorAll("sheetData row").forEach(row=>row.querySelectorAll("c").forEach(c=>{const ref=c.getAttribute("r")||"A1",col=columnIndex(ref),r=(parseInt(ref.match(/\d+/)?.[0]||"1")-1),type=c.getAttribute("t"),style=+(c.getAttribute("s")||0),v=c.querySelector("v")?.textContent??"",inline=c.querySelector("is t")?.textContent;matrix[r]??=[];let value=inline??v;if(type==="s")value=shared[+v]??"";else if(type==="b")value=v==="1";else if(type!=="str"&&type!=="inlineStr"&&v!==""){value=Number(v);if(styles.dateStyles.has(style)&&Number.isFinite(value))value=excelDate(value)}matrix[r][col]=value}));sheets[name]=normalizeMatrix(matrix)}return sheets;
}
function columnIndex(ref){let n=0;for(const c of ref.match(/[A-Z]+/i)?.[0]||"A")n=n*26+c.toUpperCase().charCodeAt(0)-64;return n-1}
function excelDate(serial){const d=new Date(Date.UTC(1899,11,30)+serial*86400000);return serial%1?d.toISOString().replace('T',' ').slice(0,19):d.toISOString().slice(0,10)}
function normalizeMatrix(matrix){
  const clean=matrix.filter(r=>r?.some(v=>v!==undefined&&v!==null&&String(v).trim()!==""));if(!clean.length)return{columns:[],rows:[]};const width=Math.max(...clean.map(r=>r.length));let headerIndex=clean.findIndex(r=>r.filter(v=>v!==undefined&&v!==null&&String(v).trim()!=="").length>=Math.min(2,width));if(headerIndex<0)headerIndex=0;const raw=clean[headerIndex],seen={};const columns=Array.from({length:width},(_,i)=>{let h=String(raw[i]??`열 ${i+1}`).trim()||`열 ${i+1}`;seen[h]=(seen[h]||0)+1;return seen[h]>1?`${h} (${seen[h]})`:h});const rows=clean.slice(headerIndex+1).map(r=>Object.fromEntries(columns.map((c,i)=>[c,coerce(r[i])]))).filter(r=>Object.values(r).some(v=>v!==""&&v!==null));return{columns,rows};
}
function coerce(v){if(v===undefined||v===null)return"";if(typeof v!=="string")return v;const t=v.trim();if(t==="")return"";if(/^[-+]?\d[\d,]*(\.\d+)?%$/.test(t))return parseFloat(t.replace(/,/g,''))/100;if(/^[-+]?\d[\d,]*(\.\d+)?$/.test(t))return Number(t.replace(/,/g,''));return t}

function renderFileList(){
  $("#fileCount").textContent=state.files.length;$("#fileList").innerHTML=state.files.map((f,i)=>`<button class="file-item ${i===state.activeFile?'active':''}" data-index="${i}" ${f.error?'title="'+f.error+'"':''}><span class="file-icon">${f.name.split('.').pop().toUpperCase()}</span><span><div class="file-name">${esc(f.name)}</div><div class="file-meta">${f.error?'읽기 실패':Object.keys(f.sheets).length+'개 시트 · '+formatBytes(f.size)}</div></span></button>`).join('');$$('.file-item').forEach(b=>b.onclick=()=>selectFile(+b.dataset.index));
}
function selectFile(index){const f=state.files[index];if(!f||f.error)return;state.activeFile=index;renderFileList();$("#pageTitle").textContent=f.name.replace(/\.(xlsx|csv|tsv)$/i,'');$("#sheetSelect").innerHTML=Object.keys(f.sheets).map(s=>`<option>${esc(s)}</option>`).join('');selectSheet(Object.keys(f.sheets)[0]);$("#emptyState").hidden=true;$("#dashboard").hidden=false;$("#exportButton").disabled=false}
function selectSheet(name){const f=state.files[state.activeFile],sheet=f.sheets[name];state.activeSheet=name;state.rows=sheet.rows;state.columns=sheet.columns;state.types=inferTypes(sheet);state.page=1;populateSelectors();renderAll()}
function inferTypes(sheet){const out={};sheet.columns.forEach(c=>{const values=sheet.rows.map(r=>r[c]).filter(v=>v!==""&&v!=null).slice(0,300),nums=values.filter(v=>typeof v==="number"&&Number.isFinite(v)).length,dates=values.filter(v=>isDateLike(v)).length;out[c]=nums/Math.max(values.length,1)>.75?"number":dates/Math.max(values.length,1)>.75?"date":"category"});return out}
function isDateLike(v){return typeof v==="string"&&(/^\d{4}[-/.]\d{1,2}([-/ .]\d{1,2})?/.test(v)||/^\d{1,2}[-/.]\d{1,2}[-/.]\d{2,4}/.test(v))&&!Number.isNaN(Date.parse(v))}
function populateSelectors(){const opts=state.columns.map(c=>`<option value="${escAttr(c)}">${esc(c)}</option>`).join(''),numbers=state.columns.filter(c=>state.types[c]==="number"),dims=state.columns.filter(c=>state.types[c]!=="number");$("#xSelect").innerHTML=opts;$("#metricSelect").innerHTML=(numbers.length?numbers:state.columns).map(c=>`<option value="${escAttr(c)}">${esc(c)}</option>`).join('');$("#xSelect").value=(dims.find(c=>state.types[c]==="date")||dims[0]||state.columns[0]||"");$("#metricSelect").value=numbers[0]||state.columns[0]||"";const xt=state.types[$("#xSelect").value];$("#aggregateSelect").value=xt==="date"?(state.rows.length>500?"month":"none"):"none"}

function renderAll(){if(!state.rows.length){renderEmptySheet();return}renderKpis();renderCharts();renderInsights();renderQuality();renderTable()}
function renderEmptySheet(){$("#kpiGrid").innerHTML='<div class="kpi-card">표 형식의 데이터가 없습니다.</div>';$("#insights").innerHTML='';$("#qualityContent").innerHTML='';clearCanvas($("#trendCanvas"));clearCanvas($("#distributionCanvas"));renderTable()}
function seriesData(){
  const x=$("#xSelect").value,metric=$("#metricSelect").value,mode=$("#aggregateSelect").value;let numeric=state.columns.filter(c=>state.types[c]==="number");if(!numeric.includes(metric))numeric=[metric];if(mode==="none")return{labels:state.rows.map((r,i)=>display(r[x])||String(i+1)),series:numeric.slice(0,4).map(c=>({name:c,values:state.rows.map(r=>num(r[c]))})),raw:true};
  const groups=new Map();state.rows.forEach(r=>{let key=r[x];if(mode==="day"||mode==="month"||mode==="year"){const d=new Date(key);if(Number.isNaN(+d))return;key=mode==="day"?d.toISOString().slice(0,10):mode==="month"?d.toISOString().slice(0,7):String(d.getUTCFullYear())}else key=String(key||"(빈 값)");if(!groups.has(key))groups.set(key,{});numeric.slice(0,4).forEach(c=>{const v=num(r[c]);if(v!=null)(groups.get(key)[c]??=[]).push(v)})});const labels=[...groups.keys()];return{labels,series:numeric.slice(0,4).map(c=>({name:c,values:labels.map(k=>average(groups.get(k)[c]||[]))})),raw:false};
}
function renderKpis(){const metric=$("#metricSelect").value,vals=state.rows.map(r=>num(r[metric])).filter(v=>v!=null),first=vals[0],last=vals.at(-1),change=first&&last!=null?(last-first)/Math.abs(first):null,missing=state.rows.length-vals.length;const cards=[['행 수',formatInt(state.rows.length),`${state.columns.length}개 열`],['평균',formatNum(average(vals)),metric],['범위',`${formatNum(Math.min(...vals))} – ${formatNum(Math.max(...vals))}`,'최솟값 – 최댓값'],['첫 값 대비',change==null?'—':formatPct(change),missing?`${missing}개 빈 값`:'결측값 없음']];$("#kpiGrid").innerHTML=cards.map((c,i)=>`<article class="kpi-card"><div class="kpi-label">${esc(c[0])}</div><div class="kpi-value ${i===3&&change!=null?(change>=0?'positive':'negative'):''}">${esc(c[1])}</div><div class="kpi-sub">${esc(c[2])}</div></article>`).join('')}
function renderCharts(){if(!state.rows.length)return;const data=seriesData();$("#trendTitle").textContent=`${$("#metricSelect").value} 추세`;drawTrend($("#trendCanvas"),data,state.chart);drawDistribution($("#distributionCanvas"),state.rows.map(r=>num(r[$("#metricSelect").value])).filter(v=>v!=null));$("#legend").innerHTML=data.series.map((s,i)=>`<span style="--legend:${colors[i]}">${esc(s.name)}</span>`).join('')}

function canvasContext(canvas,height){const rect=canvas.getBoundingClientRect(),dpr=Math.min(devicePixelRatio||1,2);canvas.width=Math.max(1,rect.width*dpr);canvas.height=Math.max(1,height*dpr);canvas.style.height=height+'px';const ctx=canvas.getContext('2d');ctx.scale(dpr,dpr);return{ctx,w:rect.width,h:height}}
function clearCanvas(c){const x=c.getContext('2d');x.clearRect(0,0,c.width,c.height)}
function drawTrend(canvas,data,type){
  const {ctx,w,h}=canvasContext(canvas,340),pad={l:55,r:16,t:20,b:42},pw=w-pad.l-pad.r,ph=h-pad.t-pad.b,all=data.series.flatMap(s=>s.values).filter(v=>v!=null);if(!all.length)return;let min=Math.min(...all),max=Math.max(...all);if(min===max){min-=1;max+=1}ctx.clearRect(0,0,w,h);ctx.font='11px system-ui';ctx.strokeStyle=css('--line');ctx.fillStyle=css('--muted');ctx.lineWidth=1;
  for(let i=0;i<5;i++){const y=pad.t+ph*i/4,val=max-(max-min)*i/4;ctx.beginPath();ctx.moveTo(pad.l,y);ctx.lineTo(w-pad.r,y);ctx.stroke();ctx.fillText(compact(val),4,y+4)}
  const count=data.labels.length,step=pw/Math.max(count-1,1),stride=Math.max(1,Math.ceil(count/6));data.labels.forEach((l,i)=>{if(i%stride&&i!==count-1)return;const x=pad.l+step*i;ctx.fillText(truncate(String(l),12),Math.min(x,w-70),h-14)});
  data.series.forEach((s,si)=>{ctx.strokeStyle=colors[si];ctx.fillStyle=colors[si];ctx.lineWidth=2;if(type==='bar'){const bw=Math.max(1,pw/Math.max(count,1)/data.series.length*.72);s.values.forEach((v,i)=>{if(v==null)return;const x=pad.l+pw*i/Math.max(count,1)+si*bw,y=pad.t+(max-v)/(max-min)*ph;ctx.fillRect(x,y,bw,Math.max(1,pad.t+ph-y))})}else{ctx.beginPath();let started=false;s.values.forEach((v,i)=>{if(v==null)return;const x=pad.l+step*i,y=pad.t+(max-v)/(max-min)*ph;started?ctx.lineTo(x,y):ctx.moveTo(x,y);started=true});ctx.stroke()}});
  canvas.onmousemove=e=>{const rect=canvas.getBoundingClientRect(),i=Math.max(0,Math.min(count-1,Math.round((e.clientX-rect.left-pad.l)/Math.max(step,1)))),tip=$("#chartTooltip");tip.style.left=`${pad.l+step*i}px`;tip.style.top=`${Math.max(50,e.clientY-rect.top)}px`;tip.innerHTML=`<strong>${esc(data.labels[i])}</strong><br>${data.series.map(s=>`${esc(shortName(s.name))}: ${formatNum(s.values[i])}`).join('<br>')}`;tip.style.opacity=1};canvas.onmouseleave=()=>$("#chartTooltip").style.opacity=0;
}
function drawDistribution(canvas,vals){const {ctx,w,h}=canvasContext(canvas,240);ctx.clearRect(0,0,w,h);if(vals.length<2)return;const bins=Math.min(12,Math.max(5,Math.round(Math.sqrt(vals.length)))),min=Math.min(...vals),max=Math.max(...vals),step=(max-min||1)/bins,counts=Array(bins).fill(0);vals.forEach(v=>counts[Math.min(bins-1,Math.floor((v-min)/step))]++);const top=Math.max(...counts),pad=28,bw=(w-pad*2)/bins;counts.forEach((c,i)=>{const bh=(h-55)*c/top;ctx.fillStyle=i===counts.indexOf(top)?css('--accent'):css('--accent-2');ctx.globalAlpha=.82;ctx.fillRect(pad+i*bw+2,h-28-bh,Math.max(2,bw-4),bh)});ctx.globalAlpha=1;ctx.fillStyle=css('--muted');ctx.font='11px system-ui';ctx.fillText(compact(min),pad,h-9);ctx.fillText(compact(max),w-pad-ctx.measureText(compact(max)).width,h-9)}

function renderInsights(){const metric=$("#metricSelect").value,vals=state.rows.map(r=>num(r[metric])).filter(v=>v!=null),ins=[];if(vals.length){const first=vals[0],last=vals.at(-1),chg=first?((last-first)/Math.abs(first)):0;ins.push(['전체 변화',`첫 값 ${formatNum(first)}에서 마지막 값 ${formatNum(last)}로 ${chg>=0?'증가':'감소'}했습니다 (${formatPct(chg)}).`]);const mean=average(vals),sd=std(vals),out=vals.filter(v=>Math.abs(v-mean)>2*sd).length;ins.push(['변동성',`평균 ${formatNum(mean)}, 표준편차 ${formatNum(sd)}입니다. 평균에서 2표준편차 이상 벗어난 값은 ${formatInt(out)}개입니다.`]);const nums=state.columns.filter(c=>state.types[c]==="number"&&c!==metric),corrs=nums.map(c=>[c,correlation(state.rows.map(r=>num(r[metric])),state.rows.map(r=>num(r[c])))]).filter(x=>Number.isFinite(x[1])).sort((a,b)=>Math.abs(b[1])-Math.abs(a[1]));if(corrs[0])ins.push(['가장 강한 관계',`${shortName(corrs[0][0])} 열과의 상관계수는 ${corrs[0][1].toFixed(2)}입니다. ${Math.abs(corrs[0][1])>.7?'강한 선형 관계가 관찰됩니다.':'뚜렷한 선형 관계는 제한적입니다.'}`])}else ins.push(['분석 값 확인','숫자형 열을 선택하면 변화율과 분포를 계산합니다.']);$("#insights").innerHTML=ins.map(x=>`<div class="insight"><strong>${x[0]}</strong><p>${esc(x[1])}</p></div>`).join('')}
function renderQuality(){let missing=0,dupes=0,total=state.rows.length*state.columns.length;state.rows.forEach(r=>state.columns.forEach(c=>{if(r[c]===""||r[c]==null)missing++}));const keys=new Set;state.rows.forEach(r=>{const k=JSON.stringify(state.columns.map(c=>r[c]));if(keys.has(k))dupes++;keys.add(k)});const complete=total?1-missing/total:0;$("#qualityContent").innerHTML=`<div class="quality-grid"><div class="quality-item"><strong>${formatPct(complete)}</strong><span>값 완성도</span></div><div class="quality-item"><strong>${formatInt(missing)}</strong><span>빈 셀</span></div><div class="quality-item"><strong>${formatInt(dupes)}</strong><span>중복 행</span></div><div class="quality-item"><strong>${state.columns.length}</strong><span>분석 열</span></div></div><div class="quality-bar"><i style="width:${complete*100}%"></i></div>`}
function renderTable(){const rows=state.rows.filter(r=>!state.search||state.columns.some(c=>String(r[c]).toLowerCase().includes(state.search))),pages=Math.max(1,Math.ceil(rows.length/state.pageSize));state.page=Math.min(state.page,pages);const start=(state.page-1)*state.pageSize,slice=rows.slice(start,start+state.pageSize);$("#dataTable thead").innerHTML=`<tr>${state.columns.map(c=>`<th title="${escAttr(c)}">${esc(c)}</th>`).join('')}</tr>`;$("#dataTable tbody").innerHTML=slice.map(r=>`<tr>${state.columns.map(c=>`<td title="${escAttr(display(r[c]))}">${esc(display(r[c]))}</td>`).join('')}</tr>`).join('');$("#rowStatus").textContent=`${formatInt(rows.length)}개 행`;$("#pageInfo").textContent=`${state.page} / ${pages}`;$("#prevPage").disabled=state.page<=1;$("#nextPage").disabled=state.page>=pages}
function exportSummary(){const metric=$("#metricSelect").value,vals=state.rows.map(r=>num(r[metric])).filter(v=>v!=null),lines=[["항목","값"],["파일",state.files[state.activeFile].name],["시트",state.activeSheet],["분석 열",metric],["행 수",state.rows.length],["유효 숫자",vals.length],["평균",average(vals)],["최솟값",Math.min(...vals)],["최댓값",Math.max(...vals)],["표준편차",std(vals)]].map(r=>r.map(csvCell).join(','));const blob=new Blob(['\uFEFF'+lines.join('\r\n')],{type:'text/csv;charset=utf-8'}),a=document.createElement('a');a.href=URL.createObjectURL(blob);a.download=`${state.activeSheet}_분석요약.csv`;a.click();URL.revokeObjectURL(a.href)}

function showError(msg){alert(msg)}function setBusy(b){$("#folderButton").disabled=b;$("#emptyFolderButton").disabled=b;$("#folderButton").textContent=b?'읽는 중…':'＋ 폴더 선택'}
function num(v){return typeof v==='number'&&Number.isFinite(v)?v:null}function average(a){return a.length?a.reduce((x,y)=>x+y,0)/a.length:null}function std(a){const m=average(a);return a.length?Math.sqrt(a.reduce((s,v)=>s+(v-m)**2,0)/a.length):0}
function correlation(a,b){const pairs=a.map((x,i)=>[x,b[i]]).filter(p=>p[0]!=null&&p[1]!=null);if(pairs.length<3)return NaN;const ax=average(pairs.map(p=>p[0])),bx=average(pairs.map(p=>p[1])),n=pairs.reduce((s,p)=>s+(p[0]-ax)*(p[1]-bx),0),d=Math.sqrt(pairs.reduce((s,p)=>s+(p[0]-ax)**2,0)*pairs.reduce((s,p)=>s+(p[1]-bx)**2,0));return d?n/d:NaN}
function display(v){if(v==null||v==='')return'';if(typeof v==='number')return formatNum(v);return String(v)}function formatNum(v){if(v==null||!Number.isFinite(v))return'—';return new Intl.NumberFormat('ko-KR',{maximumFractionDigits:2}).format(v)}function formatInt(v){return new Intl.NumberFormat('ko-KR').format(v)}function formatPct(v){return Number.isFinite(v)?new Intl.NumberFormat('ko-KR',{style:'percent',maximumFractionDigits:1,signDisplay:'exceptZero'}).format(v):'—'}function compact(v){return new Intl.NumberFormat('ko-KR',{notation:'compact',maximumFractionDigits:1}).format(v)}
function formatBytes(n){return n<1024?`${n} B`:n<1048576?`${(n/1024).toFixed(1)} KB`:`${(n/1048576).toFixed(1)} MB`}function shortName(s){return String(s).replace(/\([^)]*\)/g,'').replace(/(매출|연매출)$/,'').trim()}function truncate(s,n){return s.length>n?s.slice(0,n-1)+'…':s}function css(v){return getComputedStyle(document.documentElement).getPropertyValue(v).trim()}function csvCell(v){const s=String(v??'');return /[",\n]/.test(s)?`"${s.replace(/"/g,'""')}"`:s}function esc(s){return String(s??'').replace(/[&<>"']/g,m=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[m]))}function escAttr(s){return esc(s)}function debounce(fn,ms){let t;return()=>{clearTimeout(t);t=setTimeout(fn,ms)}}
