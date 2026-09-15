"use strict";

const $ = (selector) => document.querySelector(selector);
const els = {
  fileInput: $("#fileInput"), dropzone: $("#dropzone"), dropStatus: $("#dropStatus"),
  sheetSelect: $("#sheetSelect"), windowSelect: $("#windowSelect"), rangeSelect: $("#rangeSelect"),
  exportBtn: $("#exportBtn"), tableSearch: $("#tableSearch"), toast: $("#toast")
};

const state = { fileName: "", sheets: [], rows: [], columns: {}, analysis: null };
const palette = { raw: "#fb7185", smooth: "#5eead4", truth: "#38bdf8", mint: "#5eead4", cyan: "#38bdf8", amber: "#fbbf24", purple: "#a78bfa", grid: "#20324a", text: "#91a4ba" };

function showToast(message) {
  els.toast.textContent = message;
  els.toast.classList.add("show");
  clearTimeout(showToast.timer);
  showToast.timer = setTimeout(() => els.toast.classList.remove("show"), 2800);
}

function setLoading(active, label = "분석 중") {
  els.dropStatus.textContent = active ? label : "완료";
  els.dropStatus.style.color = active ? "#fbbf24" : "#5eead4";
}

function colIndex(reference) {
  const letters = (reference.match(/[A-Z]+/) || ["A"])[0];
  let value = 0;
  for (const char of letters) value = value * 26 + char.charCodeAt(0) - 64;
  return value - 1;
}

async function unzip(arrayBuffer) {
  const view = new DataView(arrayBuffer);
  let eocd = -1;
  for (let i = view.byteLength - 22; i >= Math.max(0, view.byteLength - 65558); i--) {
    if (view.getUint32(i, true) === 0x06054b50) { eocd = i; break; }
  }
  if (eocd < 0) throw new Error("유효한 XLSX 압축 구조를 찾지 못했습니다.");
  const count = view.getUint16(eocd + 10, true);
  let offset = view.getUint32(eocd + 16, true);
  const decoder = new TextDecoder("utf-8");
  const files = new Map();
  for (let i = 0; i < count; i++) {
    if (view.getUint32(offset, true) !== 0x02014b50) throw new Error("XLSX 항목 정보를 읽을 수 없습니다.");
    const method = view.getUint16(offset + 10, true);
    const compressedSize = view.getUint32(offset + 20, true);
    const fileNameLength = view.getUint16(offset + 28, true);
    const extraLength = view.getUint16(offset + 30, true);
    const commentLength = view.getUint16(offset + 32, true);
    const localOffset = view.getUint32(offset + 42, true);
    const name = decoder.decode(new Uint8Array(arrayBuffer, offset + 46, fileNameLength));
    const localNameLength = view.getUint16(localOffset + 26, true);
    const localExtraLength = view.getUint16(localOffset + 28, true);
    const dataStart = localOffset + 30 + localNameLength + localExtraLength;
    const compressed = new Uint8Array(arrayBuffer.slice(dataStart, dataStart + compressedSize));
    let content;
    if (method === 0) content = compressed;
    else if (method === 8) {
      if (!("DecompressionStream" in window)) throw new Error("이 브라우저는 로컬 XLSX 압축 해제를 지원하지 않습니다.");
      const stream = new Blob([compressed]).stream().pipeThrough(new DecompressionStream("deflate-raw"));
      content = new Uint8Array(await new Response(stream).arrayBuffer());
    } else throw new Error(`지원하지 않는 XLSX 압축 방식입니다 (${method}).`);
    files.set(name.replace(/^\//, ""), content);
    offset += 46 + fileNameLength + extraLength + commentLength;
  }
  return files;
}

function xml(files, path) {
  const bytes = files.get(path);
  if (!bytes) return null;
  return new DOMParser().parseFromString(new TextDecoder("utf-8").decode(bytes), "application/xml");
}

function parseSharedStrings(doc) {
  if (!doc) return [];
  return [...doc.getElementsByTagName("si")].map(si => [...si.getElementsByTagName("t")].map(t => t.textContent).join(""));
}

function parseSheet(doc, sharedStrings) {
  const matrix = [];
  for (const rowNode of doc.getElementsByTagName("row")) {
    const row = [];
    for (const cell of rowNode.getElementsByTagName("c")) {
      const index = colIndex(cell.getAttribute("r") || "A1");
      const type = cell.getAttribute("t");
      const valueNode = cell.getElementsByTagName("v")[0];
      const inlineNode = cell.getElementsByTagName("is")[0];
      let value = "";
      if (type === "inlineStr" && inlineNode) value = [...inlineNode.getElementsByTagName("t")].map(n => n.textContent).join("");
      else if (valueNode) {
        const raw = valueNode.textContent;
        if (type === "s") value = sharedStrings[Number(raw)] ?? "";
        else if (type === "b") value = raw === "1";
        else if (type === "str" || type === "e") value = raw;
        else value = raw === "" ? "" : Number(raw);
      }
      row[index] = value;
    }
    matrix.push(row);
  }
  const headerRowIndex = matrix.findIndex(row => row.filter(v => v !== "" && v != null).length >= 2);
  if (headerRowIndex < 0) return { headers: [], rows: [] };
  const headers = matrix[headerRowIndex].map((value, index) => String(value || `열 ${index + 1}`).trim());
  const rows = matrix.slice(headerRowIndex + 1).filter(row => row.some(v => v !== "" && v != null)).map(row => {
    const record = {};
    headers.forEach((header, i) => { record[header] = row[i] ?? ""; });
    return record;
  });
  return { headers, rows };
}

async function parseXlsx(buffer) {
  const files = await unzip(buffer);
  const workbook = xml(files, "xl/workbook.xml");
  const rels = xml(files, "xl/_rels/workbook.xml.rels");
  if (!workbook || !rels) throw new Error("통합문서 정보를 읽지 못했습니다.");
  const relations = {};
  for (const rel of rels.getElementsByTagName("Relationship")) relations[rel.getAttribute("Id")] = rel.getAttribute("Target");
  const shared = parseSharedStrings(xml(files, "xl/sharedStrings.xml"));
  const sheets = [];
  for (const sheet of workbook.getElementsByTagName("sheet")) {
    const id = sheet.getAttribute("r:id") || sheet.getAttributeNS("http://schemas.openxmlformats.org/officeDocument/2006/relationships", "id");
    const target = (relations[id] || "").replace(/^\//, "").replace(/^xl\//, "");
    const path = target.startsWith("worksheets/") ? `xl/${target}` : `xl/${target}`;
    const sheetDoc = xml(files, path);
    if (sheetDoc) sheets.push({ name: sheet.getAttribute("name") || `Sheet ${sheets.length + 1}`, ...parseSheet(sheetDoc, shared) });
  }
  return sheets;
}

function findColumn(headers, keywords, excludes = []) {
  return headers.find(header => {
    const h = header.toLowerCase();
    return keywords.some(k => h.includes(k)) && !excludes.some(k => h.includes(k));
  });
}

function excelDate(value) {
  if (typeof value !== "number" || value < 20000 || value > 80000) return value;
  return new Date(Math.round((value - 25569) * 86400 * 1000));
}

function formatTime(value) {
  const date = excelDate(value);
  if (date instanceof Date && !Number.isNaN(date.valueOf())) return new Intl.DateTimeFormat("ko-KR", { month: "2-digit", day: "2-digit", hour: "2-digit", minute: "2-digit", hour12: false, timeZone: "UTC" }).format(date);
  return String(value ?? "—");
}

const finite = value => Number.isFinite(Number(value));
const num = value => finite(value) ? Number(value) : NaN;
const mean = values => { const valid = values.filter(Number.isFinite); return valid.length ? valid.reduce((a,b)=>a+b,0)/valid.length : NaN; };
const std = values => { const m = mean(values); return Math.sqrt(mean(values.map(v => (v - m) ** 2))); };
const mae = (a, b) => mean(a.map((v, i) => Number.isFinite(v) && Number.isFinite(b[i]) ? Math.abs(v - b[i]) : NaN));

function movingAverage(values, windowSize) {
  const output = []; let sum = 0; const queue = [];
  values.forEach(value => {
    queue.push(value); if (Number.isFinite(value)) sum += value;
    if (queue.length > windowSize) { const removed = queue.shift(); if (Number.isFinite(removed)) sum -= removed; }
    const count = queue.filter(Number.isFinite).length;
    output.push(count ? sum / count : NaN);
  });
  return output;
}

function inferColumns(sheet) {
  const headers = sheet.headers;
  const numeric = headers.filter(header => sheet.rows.filter(r => finite(r[header])).length >= Math.max(3, sheet.rows.length * .5));
  const raw = findColumn(headers, ["원시", "raw", "측정값", "온도센서"], ["참값", "비교"]);
  return {
    time: findColumn(headers, ["측정시각", "timestamp", "datetime", "일시", "시간"]),
    raw: raw || numeric[0], truth: findColumn(headers, ["참값", "비교용", "reference", "target"]),
    vibration: findColumn(headers, ["진동", "vibration", "rms"]), pressure: findColumn(headers, ["압력", "pressure"]),
    event: findColumn(headers, ["이벤트", "event", "상태", "status"]), numeric
  };
}

function computeAnalysis(rows, columns, windowSize) {
  const raw = rows.map(r => num(r[columns.raw]));
  const truth = columns.truth ? rows.map(r => num(r[columns.truth])) : [];
  const smooth = movingAverage(raw, windowSize);
  const rawMae = truth.length ? mae(raw, truth) : NaN;
  const smoothMae = truth.length ? mae(smooth, truth) : NaN;
  const windows = [5, 15, 30, 60].map(w => ({ label: `${w}분`, value: truth.length ? mae(movingAverage(raw, w), truth) : std(movingAverage(raw, w).map((v,i)=>v-raw[i])) }));
  const events = {};
  rows.forEach(r => { const label = String(r[columns.event] || "미분류").trim(); events[label] = (events[label] || 0) + 1; });
  const nonNormal = columns.event ? rows.filter(r => !/정상|normal|ok/i.test(String(r[columns.event]))).length : 0;
  const errorRows = rows.map((row, index) => ({ row, index, error: truth.length && finite(raw[index]) && finite(truth[index]) ? Math.abs(raw[index] - truth[index]) : Math.abs(raw[index] - smooth[index]) })).sort((a,b)=>b.error-a.error);
  return { raw, truth, smooth, rawMae, smoothMae, windows, events, nonNormal, errorRows };
}

function fmt(value, digits = 3) { return Number.isFinite(value) ? value.toLocaleString("ko-KR", { maximumFractionDigits: digits, minimumFractionDigits: digits }) : "—"; }

function updateDashboard() {
  const rows = state.rows; const c = state.columns; const windowSize = Number(els.windowSelect.value);
  state.analysis = computeAnalysis(rows, c, windowSize);
  const a = state.analysis;
  $("#kpiRows").textContent = rows.length.toLocaleString("ko-KR");
  $("#kpiRawMae").textContent = Number.isFinite(a.rawMae) ? `${fmt(a.rawMae)}°C` : "비교열 없음";
  $("#kpiSmoothMae").textContent = Number.isFinite(a.smoothMae) ? `${fmt(a.smoothMae)}°C` : fmt(std(a.smooth.map((v,i)=>v-a.raw[i])));
  $("#kpiWindow").textContent = `${windowSize}분 이동평균`;
  $("#kpiEvents").textContent = c.event ? a.nonNormal.toLocaleString("ko-KR") : "—";
  $("#kpiEventRate").textContent = c.event ? `전체의 ${fmt(a.nonNormal / Math.max(1, rows.length) * 100, 1)}%` : "이벤트 열 없음";
  const firstTime = c.time && rows[0] ? formatTime(rows[0][c.time]) : "";
  const lastTime = c.time && rows.at(-1) ? formatTime(rows.at(-1)[c.time]) : "";
  $("#kpiPeriod").textContent = firstTime && lastTime ? `${firstTime} – ${lastTime}` : `${state.sheets.length}개 시트`;
  const improvement = Number.isFinite(a.rawMae) && Number.isFinite(a.smoothMae) ? (1 - a.smoothMae / a.rawMae) * 100 : NaN;
  $("#improvementBadge").textContent = Number.isFinite(improvement) ? `오차 ${improvement >= 0 ? "개선" : "증가"} ${fmt(Math.abs(improvement), 1)}%` : "스무딩 변동 확인";
  renderInsights(improvement); renderEventBars(); renderRecords(); drawAll();
  $("#footerStatus").textContent = `${state.fileName} · ${rows.length.toLocaleString("ko-KR")}행 · ${state.sheets.length}개 시트`;
}

function renderInsights(improvement) {
  const a = state.analysis; const c = state.columns; const insights = [];
  if (Number.isFinite(improvement)) insights.push({ title: `${els.windowSelect.value}분 창에서 오차 ${improvement >= 0 ? "감소" : "증가"}`, body: `MAE가 ${fmt(a.rawMae)}°C에서 ${fmt(a.smoothMae)}°C로 ${improvement >= 0 ? "낮아졌습니다" : "높아졌습니다"}.` });
  const best = a.windows.filter(x => Number.isFinite(x.value)).sort((x,y)=>x.value-y.value)[0];
  if (best) insights.push({ title: `비교 창 중 ${best.label} 창이 가장 정확`, body: `참값 대비 MAE ${fmt(best.value)}°C로 네 가지 이동평균 설정 중 가장 낮습니다.` });
  if (c.event) {
    const notable = Object.entries(a.events).filter(([name]) => !/정상|normal|ok/i.test(name)).sort((x,y)=>y[1]-x[1]);
    if (notable[0]) insights.push({ title: `${notable[0][0]} 이벤트가 가장 많음`, body: `비정상 이벤트 ${a.nonNormal.toLocaleString("ko-KR")}건 중 ${notable[0][1].toLocaleString("ko-KR")}건입니다.` });
  }
  if (c.vibration && c.raw) {
    const x = state.rows.map(r=>num(r[c.raw])), y = state.rows.map(r=>num(r[c.vibration]));
    const corr = correlation(x,y);
    if (Number.isFinite(corr)) insights.push({ title: `온도–진동 상관계수 ${fmt(corr, 2)}`, body: `${Math.abs(corr) >= .7 ? "두 센서가 강하게 함께 움직입니다." : Math.abs(corr) >= .4 ? "두 센서가 중간 수준으로 함께 움직입니다." : "두 센서의 동조는 제한적입니다."}` });
  }
  $("#insightList").innerHTML = insights.slice(0,4).map((item,i)=>`<div class="insight"><span class="insight-index">${String(i+1).padStart(2,"0")}</span><div><strong>${escapeHtml(item.title)}</strong><p>${escapeHtml(item.body)}</p></div></div>`).join("") || '<p class="empty-copy">분석 가능한 수치 열이 부족합니다.</p>';
}

function correlation(a,b) {
  const pairs = a.map((x,i)=>[x,b[i]]).filter(([x,y])=>Number.isFinite(x)&&Number.isFinite(y));
  if (pairs.length < 2) return NaN;
  const ax=mean(pairs.map(p=>p[0])), ay=mean(pairs.map(p=>p[1]));
  const numerator=pairs.reduce((s,[x,y])=>s+(x-ax)*(y-ay),0);
  const denominator=Math.sqrt(pairs.reduce((s,[x])=>s+(x-ax)**2,0)*pairs.reduce((s,[,y])=>s+(y-ay)**2,0));
  return denominator ? numerator/denominator : NaN;
}

function renderEventBars() {
  const entries = Object.entries(state.analysis.events).sort((a,b)=>b[1]-a[1]);
  $("#eventCountLabel").textContent = `${entries.length}개 유형`;
  if (!entries.length) { $("#eventBars").innerHTML = '<p class="empty-copy">표시할 이벤트가 없습니다.</p>'; return; }
  const max = Math.max(...entries.map(([,v])=>v));
  $("#eventBars").innerHTML = entries.slice(0,7).map(([name,value])=>`<div class="event-row"><span title="${escapeHtml(name)}">${escapeHtml(shorten(name,9))}</span><div class="bar-track"><div class="bar-fill" style="width:${Math.max(2,value/max*100)}%"></div></div><strong>${value}</strong></div>`).join("");
}

function renderRecords() {
  const c=state.columns, query=els.tableSearch.value.trim().toLowerCase();
  const records=state.analysis.errorRows.filter(item=>!query || String(item.row[c.event]||"").toLowerCase().includes(query)).slice(0,10);
  $("#recordsBody").innerHTML=records.map(({row,error})=>{
    const event=String(row[c.event]||"미분류");
    return `<tr><td>${escapeHtml(formatTime(row[c.time]))}</td><td>${fmt(num(row[c.raw]))}</td><td>${c.truth?fmt(num(row[c.truth])):"—"}</td><td class="error-value">${fmt(error)}</td><td>${c.vibration?fmt(num(row[c.vibration])):"—"}</td><td>${c.pressure?fmt(num(row[c.pressure])):"—"}</td><td><span class="event-pill ${/정상|normal|ok/i.test(event)?"normal":""}">${escapeHtml(event)}</span></td></tr>`;
  }).join("") || '<tr><td colspan="7" class="empty-cell">검색 결과가 없습니다.</td></tr>';
}

function setupCanvas(canvas) {
  const rect=canvas.getBoundingClientRect(), ratio=window.devicePixelRatio||1;
  canvas.width=Math.max(1,Math.round(rect.width*ratio)); canvas.height=Math.max(1,Math.round(rect.height*ratio));
  const ctx=canvas.getContext("2d"); ctx.scale(ratio,ratio); return {ctx,w:rect.width,h:rect.height};
}

function drawLineChart(canvas, series) {
  const {ctx,w,h}=setupCanvas(canvas); const pad={l:44,r:14,t:12,b:28};
  const all=series.flatMap(s=>s.values).filter(Number.isFinite); if(!all.length) return drawEmpty(ctx,w,h);
  let min=Math.min(...all), max=Math.max(...all); const extra=(max-min||1)*.08; min-=extra; max+=extra;
  ctx.font='11px system-ui'; ctx.fillStyle=palette.text; ctx.strokeStyle=palette.grid; ctx.lineWidth=1;
  for(let i=0;i<5;i++){const y=pad.t+(h-pad.t-pad.b)*i/4;ctx.beginPath();ctx.moveTo(pad.l,y);ctx.lineTo(w-pad.r,y);ctx.stroke();const label=max-(max-min)*i/4;ctx.fillText(label.toFixed(1),4,y+4)}
  series.forEach(s=>{ctx.strokeStyle=s.color;ctx.lineWidth=s.width||1.5;ctx.globalAlpha=s.alpha||1;ctx.beginPath();s.values.forEach((v,i)=>{if(!Number.isFinite(v))return;const x=pad.l+(w-pad.l-pad.r)*i/Math.max(1,s.values.length-1);const y=pad.t+(max-v)/(max-min)*(h-pad.t-pad.b);if(i===0)ctx.moveTo(x,y);else ctx.lineTo(x,y)});ctx.stroke()});ctx.globalAlpha=1;
  ctx.fillStyle=palette.text;ctx.fillText("시작",pad.l,h-7);ctx.fillText("최근",w-pad.r-24,h-7);
}

function drawBarChart(canvas, entries) {
  const {ctx,w,h}=setupCanvas(canvas), pad={l:42,r:12,t:16,b:32}; const max=Math.max(...entries.map(e=>e.value).filter(Number.isFinite),1);
  ctx.font='11px system-ui';ctx.fillStyle=palette.text;ctx.strokeStyle=palette.grid;
  for(let i=0;i<4;i++){const y=pad.t+(h-pad.t-pad.b)*i/3;ctx.beginPath();ctx.moveTo(pad.l,y);ctx.lineTo(w-pad.r,y);ctx.stroke();ctx.fillText((max*(1-i/3)).toFixed(2),4,y+4)}
  const gap=10, bw=(w-pad.l-pad.r-gap*(entries.length-1))/entries.length;
  entries.forEach((e,i)=>{const bh=(h-pad.t-pad.b)*(e.value/max),x=pad.l+i*(bw+gap),y=h-pad.b-bh;const gradient=ctx.createLinearGradient(0,y,0,h-pad.b);gradient.addColorStop(0,i===0?palette.mint:palette.cyan);gradient.addColorStop(1,"#1a4b64");ctx.fillStyle=gradient;roundRect(ctx,x,y,bw,bh,5);ctx.fill();ctx.fillStyle=palette.text;ctx.textAlign="center";ctx.fillText(e.label,x+bw/2,h-10)});ctx.textAlign="left";
}

function roundRect(ctx,x,y,w,h,r){const rr=Math.min(r,w/2,h/2);ctx.beginPath();ctx.moveTo(x+rr,y);ctx.arcTo(x+w,y,x+w,y+h,rr);ctx.arcTo(x+w,y+h,x,y+h,rr);ctx.arcTo(x,y+h,x,y,rr);ctx.arcTo(x,y,x+w,y,rr);ctx.closePath()}
function drawEmpty(ctx,w,h){ctx.fillStyle=palette.text;ctx.font='12px system-ui';ctx.textAlign='center';ctx.fillText('표시할 수치 데이터가 없습니다.',w/2,h/2);ctx.textAlign='left'}

function drawAll() {
  if(!state.analysis)return; const requested=Number(els.rangeSelect.value), count=Math.min(requested,state.rows.length), start=Math.max(0,state.rows.length-count), a=state.analysis, c=state.columns;
  drawLineChart($("#temperatureChart"),[{values:a.raw.slice(start),color:palette.raw,alpha:.5,width:1},{values:a.smooth.slice(start),color:palette.smooth,width:2.5},{values:a.truth.slice(start),color:palette.truth,width:1.6,alpha:.9}]);
  const multi=[]; [[c.raw,palette.mint],[c.vibration,palette.amber],[c.pressure,palette.purple]].forEach(([column,color])=>{if(!column)return;const values=state.rows.slice(start).map(r=>num(r[column])),m=mean(values),s=std(values)||1;multi.push({values:values.map(v=>(v-m)/s),color,width:1.5,alpha:.9})});
  drawLineChart($("#multiChart"),multi); drawBarChart($("#errorChart"),a.windows);
}

function activateSheet(index) {
  const sheet=state.sheets[index]; state.rows=sheet.rows; state.columns=inferColumns(sheet);
  if(!state.columns.raw) { showToast("수치 열이 있는 시트를 선택해 주세요."); return; }
  updateDashboard();
}

async function loadFile(fileOrBuffer, name) {
  try {
    setLoading(true,"읽는 중"); const buffer=fileOrBuffer instanceof ArrayBuffer?fileOrBuffer:await fileOrBuffer.arrayBuffer();
    const sheets=await parseXlsx(buffer); if(!sheets.length)throw new Error("읽을 수 있는 시트가 없습니다.");
    state.fileName=name;state.sheets=sheets;
    els.sheetSelect.innerHTML=sheets.map((s,i)=>`<option value="${i}">${escapeHtml(s.name)} (${s.rows.length.toLocaleString("ko-KR")}행)</option>`).join("");
    els.sheetSelect.disabled=false;els.exportBtn.disabled=false;
    let best=0,bestScore=-1;sheets.forEach((sheet,i)=>{const c=inferColumns(sheet),score=sheet.rows.length+(c.raw?10000:0)+(c.truth?1000:0);if(score>bestScore){best=i;bestScore=score}});els.sheetSelect.value=String(best);
    $("#fileSummary").textContent=`${name} · ${sheets.length}개 시트를 로컬에서 분석했습니다.`;activateSheet(best);setLoading(false);showToast("분석이 완료되었습니다.");
  } catch(error) { console.error(error);els.dropStatus.textContent="오류";els.dropStatus.style.color=palette.coral;showToast(error.message||"파일을 읽지 못했습니다."); }
}

function exportCsv() {
  if(!state.rows.length)return;const c=state.columns,a=state.analysis,headers=[...Object.keys(state.rows[0]),`${els.windowSelect.value}분 이동평균`,`절대오차`];
  const quote=v=>`"${String(v??"").replaceAll('"','""')}"`;
  const lines=[headers.map(quote).join(","),...state.rows.map((row,i)=>[...Object.keys(state.rows[0]).map(h=>row[h]),Number.isFinite(a.smooth[i])?a.smooth[i].toFixed(4):"",Number.isFinite(a.errorRows.find(x=>x.index===i)?.error)?a.errorRows.find(x=>x.index===i).error.toFixed(4):""].map(quote).join(","))];
  const blob=new Blob(["\ufeff"+lines.join("\r\n")],{type:"text/csv;charset=utf-8"}),url=URL.createObjectURL(blob),link=document.createElement("a");link.href=url;link.download=`${state.fileName.replace(/\.xlsx$/i,"")}_분석.csv`;link.click();URL.revokeObjectURL(url);showToast("분석 CSV를 저장했습니다.");
}

function escapeHtml(value){return String(value??"").replace(/[&<>'"]/g,ch=>({"&":"&amp;","<":"&lt;",">":"&gt;","'":"&#39;",'"':"&quot;"}[ch]))}
function shorten(value,length){const text=String(value);return text.length>length?`${text.slice(0,length-1)}…`:text}

els.fileInput.addEventListener("change",e=>e.target.files[0]&&loadFile(e.target.files[0],e.target.files[0].name));
els.dropzone.addEventListener("click",()=>els.fileInput.click());els.dropzone.addEventListener("keydown",e=>{if(e.key==="Enter"||e.key===" "){e.preventDefault();els.fileInput.click()}});
["dragenter","dragover"].forEach(type=>els.dropzone.addEventListener(type,e=>{e.preventDefault();els.dropzone.classList.add("dragging")}));["dragleave","drop"].forEach(type=>els.dropzone.addEventListener(type,e=>{e.preventDefault();els.dropzone.classList.remove("dragging")}));
els.dropzone.addEventListener("drop",e=>{const file=[...e.dataTransfer.files].find(f=>/\.xlsx$/i.test(f.name));file?loadFile(file,file.name):showToast(".xlsx 파일을 선택해 주세요.")});
els.sheetSelect.addEventListener("change",()=>activateSheet(Number(els.sheetSelect.value)));els.windowSelect.addEventListener("change",updateDashboard);els.rangeSelect.addEventListener("change",drawAll);els.tableSearch.addEventListener("input",renderRecords);els.exportBtn.addEventListener("click",exportCsv);window.addEventListener("resize",()=>state.analysis&&drawAll());

if (document.modelContext?.registerTool) {
  const signal=new AbortController();
  Promise.resolve(document.modelContext.registerTool({name:"configure_sensor_analysis",title:"센서 분석 설정",description:"현재 대시보드의 이동평균 창과 표시 범위를 변경합니다.",inputSchema:{type:"object",properties:{smoothingWindow:{type:"integer",enum:[5,15,30,60]},rangeMinutes:{type:"integer",enum:[360,720,1440]}},additionalProperties:false},annotations:{readOnlyHint:false,untrustedContentHint:false},execute(input){
    const allowedWindows=[5,15,30,60],allowedRanges=[360,720,1440];
    if(!input||typeof input!=="object"||Array.isArray(input))throw new Error("설정은 객체 형식이어야 합니다.");
    if("smoothingWindow" in input&&!allowedWindows.includes(input.smoothingWindow))throw new Error("스무딩 창은 5, 15, 30, 60분 중 하나여야 합니다.");
    if("rangeMinutes" in input&&!allowedRanges.includes(input.rangeMinutes))throw new Error("표시 범위는 360, 720, 1440분 중 하나여야 합니다.");
    if(input.smoothingWindow){els.windowSelect.value=String(input.smoothingWindow);updateDashboard()}
    if(input.rangeMinutes){els.rangeSelect.value=String(input.rangeMinutes);drawAll()}
    return{window:Number(els.windowSelect.value),range:Number(els.rangeSelect.value),rows:state.rows.length}
  }},{signal:signal.signal})).catch(()=>{});
}

async function loadBundledWorkbook(){if(location.protocol==="file:")return;try{const response=await fetch("센서 스무딩.xlsx",{cache:"no-store"});if(response.ok)await loadFile(await response.arrayBuffer(),"센서 스무딩.xlsx")}catch{/* 파일 선택 방식으로 계속 사용 */}}
loadBundledWorkbook();
