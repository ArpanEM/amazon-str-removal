/* Amazon STR & Removal Control Tower
   Static frontend for GitHub Pages + Supabase. All privileged access is enforced by RLS.
   Configure config.js before use. CSV parser is intentionally dependency-free. */
(() => {
  const CONTROL_TOWER_BUILD = "2026-11-12-horizontal-scroll-controls";
  console.info("Amazon STR & Removal Control Tower build:", CONTROL_TOWER_BUILD);
  const cfg = window.APP_CONFIG || {};
  const $ = (s, root=document) => root.querySelector(s);
  const $$ = (s, root=document) => [...root.querySelectorAll(s)];
  const fmt = n => Number(n || 0).toLocaleString("en-IN");
  const money = (n,c="INR") => new Intl.NumberFormat("en-IN",{style:"currency",currency:c,maximumFractionDigits:2}).format(Number(n||0));
  const today = () => new Date().toISOString().slice(0,10);
  let sb=null, session=null, profile=null, cache={orders:[],shipments:[],receipts:[],fees:[],claims:[],sap:[],imports:[],mappings:[]}, trend=null;
  const say=(id,msg,ok=false)=>{const el=$("#"+id);if(el){el.textContent=msg;el.style.color=ok?"#16845b":"#c2413b";}};
  const esc=s=>String(s??"").replace(/[&<>"']/g,c=>({"&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;","'":"&#39;"}[c]));
  function parseCSV(text){
    const rows=[];let row=[],v="",q=false;
    for(let i=0;i<text.length;i++){const c=text[i];if(c==='"'){if(q&&text[i+1]==='"'){v+='"';i++;}else q=!q;}else if(c===","&&!q){row.push(v);v="";}else if((c==="\n"||c==="\r")&&!q){if(c==="\r"&&text[i+1]==="\n")i++;row.push(v);if(row.some(x=>x!==""))rows.push(row);row=[];v="";}else v+=c;}
    row.push(v);if(row.some(x=>x!==""))rows.push(row);if(!rows.length)return [];
    const rawHeaders=rows.shift().map(h=>h.trim().replace(/^\uFEFF/,""));
    const seen={};const headers=rawHeaders.map(h=>{seen[h]=(seen[h]||0)+1;return seen[h]===1?h:`${h}__${seen[h]}`;});
    return rows.map(r=>Object.fromEntries(headers.map((h,i)=>[h,(r[i]??"").trim()])));
  }
  async function readReportFile(file){
    const ext=(file.name.split('.').pop()||'').toLowerCase();
    if(ext==='csv') return parseCSV(await file.text());
    if(!['xlsx','xls'].includes(ext)) throw new Error('Unsupported file. Please upload .xlsx, .xls or .csv.');
    if(!window.XLSX) throw new Error('Excel reader did not load. Refresh the page and try again.');
    const book=XLSX.read(await file.arrayBuffer(),{type:'array',cellDates:true});
    const sheetName=book.SheetNames.find(n=>!book.Workbook?.Sheets?.find(s=>s.name===n)?.Hidden) || book.SheetNames[0];
    if(!sheetName) throw new Error('No worksheet found in this Excel file.');
    const sheet=book.Sheets[sheetName];
    const rows=XLSX.utils.sheet_to_json(sheet,{defval:'',raw:false});
    if(!rows.length) throw new Error(`The first worksheet (${sheetName}) is empty or has no header row.`);
    return rows.map(r=>Object.fromEntries(Object.entries(r).map(([k,v])=>[String(k).trim(),String(v??'').trim()])));
  }
  const pick=(o,...keys)=>{for(const k of keys){const found=Object.keys(o||{}).find(x=>x.toLowerCase()===k.toLowerCase());if(found)return o[found];}return "";};
  const num=v=>{const n=Number(String(v??"").replace(/,/g,""));return Number.isFinite(n)?n:0;};
  // Amazon India exports dates as DD/MM/YY HH:mm (for example 25/09/26 05:30).
  // The STR database columns are DATE, so normalize to YYYY-MM-DD and discard time.
  // Also accepts ISO dates, DD/MM/YYYY, DD-MM-YY and Excel-parsed date strings.
  const dateISO=v=>{
    if(v===null||v===undefined||v==="")return null;
    if(v instanceof Date&&!Number.isNaN(v.getTime()))return `${v.getFullYear()}-${String(v.getMonth()+1).padStart(2,"0")}-${String(v.getDate()).padStart(2,"0")}`;
    const x=String(v).trim();if(!x)return null;
    let m=x.match(/^(\d{4})-(\d{1,2})-(\d{1,2})(?:[T\s].*)?$/);
    if(m){const y=Number(m[1]),mo=Number(m[2]),d=Number(m[3]);if(y>=1900&&mo>=1&&mo<=12&&d>=1&&d<=31)return `${String(y).padStart(4,"0")}-${String(mo).padStart(2,"0")}-${String(d).padStart(2,"0")}`;return null;}
    m=x.match(/^(\d{1,2})[./-](\d{1,2})[./-](\d{2}|\d{4})(?:[ T].*)?$/);
    if(m){const d=Number(m[1]),mo=Number(m[2]);let y=Number(m[3]);if(m[3].length===2)y+=2000;if(y>=1900&&y<=2100&&mo>=1&&mo<=12&&d>=1&&d<=new Date(y,mo,0).getDate())return `${String(y).padStart(4,"0")}-${String(mo).padStart(2,"0")}-${String(d).padStart(2,"0")}`;return null;}
    return null;
  };
  const norm=s=>String(s??"").trim().toLowerCase();
  async function audit(action,entity,details={}){if(!sb||!session)return;await sb.from("audit_logs").insert({user_id:session.user.id,action,entity,details});}
  async function init(){
    if(!cfg.SUPABASE_URL||!cfg.SUPABASE_ANON_KEY||cfg.SUPABASE_URL.includes("YOUR-PROJECT")){$("#loginMsg").textContent="Configure config.js with your Supabase URL and anon key, then reload.";return;}
    sb=window.supabase.createClient(cfg.SUPABASE_URL,cfg.SUPABASE_ANON_KEY);
    const {data:{session:s}}=await sb.auth.getSession();if(s)await enter(s);
    sb.auth.onAuthStateChange(async(_e,s)=>{if(s&&!session)await enter(s);else if(!s){session=null;$("#appView").classList.add("hidden");$("#loginView").classList.remove("hidden");}});
  }
  async function enter(s){
    session=s;
    let {data:p}=await sb.from("profiles").select("*").eq("id",s.user.id).maybeSingle();
    if(!p){const name=s.user.user_metadata?.full_name||s.user.email.split("@")[0];const ins=await sb.from("profiles").insert({id:s.user.id,email:s.user.email,full_name:name,role:"employee"}).select().single();p=ins.data;}
    profile=p;$("#userName").textContent=p?.full_name||s.user.email;$("#userRole").textContent=p?.role||"employee";$("#userInitial").textContent=(p?.full_name||"U").slice(0,1).toUpperCase();
    $$(".admin-only").forEach(el=>el.classList.toggle("hidden",!["admin","manager"].includes(p?.role)));applyPageAccess(p);
    $("#loginView").classList.add("hidden");$("#appView").classList.remove("hidden");await loadAll();installHorizontalScrollControls();
  }
  async function fetchAllRows(table){
    const orderCol={mail_settings:"updated_at",fc_mappings:"updated_at"}[table]||"created_at";
    const out=[],size=1000;
    for(let from=0;;from+=size){
      const {data,error}=await sb.from(table).select("*").order(orderCol,{ascending:false}).order("id",{ascending:false}).range(from,from+size-1);
      if(error)return {error};
      out.push(...(data||[]));
      if(!data||data.length<size)break;
    }
    return {data:out};
  }
  async function loadAll(){
    const tables=["removal_orders","shipment_lines","warehouse_receipts","amazon_removal_fees","amazon_claims","sap_postings","import_batches","str_records","product_mappings","po_receipts","amazon_invoices","invoice_zip_imports","mail_settings","approval_batches","approval_items","fc_mappings","sap_stock","scat_exports","scat_export_lines","receipt_evidence"];
    const names=["orders","shipments","receipts","fees","claims","sap","imports","str","mappings","poReceipts","invoices","zipImports","mailSettings","approvalBatches","approvalItems","fcMappings","sapStock","scatExports","scatLines","evidence"];
    for(let i=0;i<tables.length;i++){const {data,error}=await fetchAllRows(tables[i]);if(!error)cache[names[i]]=data||[];}
    renderAll();
  }
  function fyDate(d){if(!d)return null;const x=new Date(d+"T00:00:00");return x;}
  function filtered(arr,dateKey){const m=$("#monthFilter").value,fy=$("#fyFilter").value;return arr.filter(x=>{const d=fyDate(x[dateKey]);if(!d)return m==="all"&&fy==="all";const mm=String(d.getMonth()+1).padStart(2,"0"),month=d.getFullYear()+"-"+mm;const fiscal=d.getMonth()>=3?`${d.getFullYear()}-${String(d.getFullYear()+1).slice(-2)}`:`${d.getFullYear()-1}-${String(d.getFullYear()).slice(-2)}`;return(m==="all"||m===month)&&(fy==="all"||fy===fiscal);});}
  function receiptTotals(shipmentId){return cache.receipts.filter(r=>r.shipment_line_id===shipmentId).reduce((a,r)=>{for(const k of ["good_qty","damaged_qty","near_expiry_qty","expired_qty","short_qty","not_received_qty","rework_qty"])a[k]+=num(r[k]);return a;},{good_qty:0,damaged_qty:0,near_expiry_qty:0,expired_qty:0,short_qty:0,not_received_qty:0,rework_qty:0});}
  function populateFilters(){
    const fySel=$("#fyFilter"),mSel=$("#monthFilter");if(!fySel||!mSel)return;
    const dates=[...(cache.str||[]).map(x=>x.invoice_date),...(cache.orders||[]).map(x=>x.request_date),...(cache.shipments||[]).map(x=>x.shipment_date),...(cache.receipts||[]).map(x=>x.receipt_date),...(cache.poReceipts||[]).map(x=>x.receipt_date),...(cache.fees||[]).map(x=>x.fee_date)].filter(Boolean).map(s=>String(s).slice(0,10));
    const fyOf=ym=>{const y=+ym.slice(0,4),m=+ym.slice(5,7);return m>=4?`${y}-${String(y+1).slice(-2)}`:`${y-1}-${String(y).slice(-2)}`;};
    const months=[...new Set(dates.map(s=>s.slice(0,7)))].sort();
    const fys=[...new Set(months.map(fyOf))].sort();
    const fyVal=fySel.value,mVal=mSel.value;
    fySel.innerHTML='<option value="all">All years</option>'+fys.map(f=>`<option value="${f}">FY ${f}</option>`).join("");
    fySel.value=fys.includes(fyVal)?fyVal:"all";
    const monthsForFy=fySel.value==="all"?months:months.filter(ym=>fyOf(ym)===fySel.value);
    mSel.innerHTML='<option value="all">All months</option>'+monthsForFy.map(ym=>{const [y,m]=ym.split("-");return `<option value="${ym}">${new Date(+y,+m-1,1).toLocaleString("en-GB",{month:"long",year:"numeric"})}</option>`;}).join("");
    mSel.value=monthsForFy.includes(mVal)?mVal:"all";
  }
  function renderAll(){
    populateFilters();renderRecon();renderDashCharts();renderReceiptEntries();renderProducts();renderApprovals();renderFcMap();renderScat();freezeAll();
    const orders=filtered(cache.orders,"request_date"),ships=filtered(cache.shipments,"shipment_date"),fees=filtered(cache.fees,"fee_date");
    const shipped=ships.reduce((a,x)=>a+num(x.shipped_qty),0),rt={good_qty:0,damaged_qty:0,near_expiry_qty:0,expired_qty:0,short_qty:0,not_received_qty:0};
    ships.forEach(s=>{const t=receiptTotals(s.id);Object.keys(rt).forEach(k=>rt[k]+=t[k]);});
    const physical=rt.good_qty+rt.damaged_qty+rt.near_expiry_qty+rt.expired_qty,final=Object.values(rt).reduce((a,b)=>a+b,0),pending=Math.max(0,shipped-final);
    const feeCurrencies=[...new Set(fees.map(x=>x.currency).filter(Boolean))],feeTotal=fees.reduce((a,x)=>a+num(x.fee_amount),0),claims=cache.claims.filter(c=>!["closed","resolved"].includes(norm(c.status)));
    $("#kpiOrders").textContent=fmt(new Set(orders.map(o=>o.order_id)).size);$("#kpiShipped").textContent=fmt(shipped);$("#kpiReceived").textContent=fmt(physical);$("#kpiPending").textContent=fmt(pending);$("#kpiFees").textContent=feeCurrencies.length>1?"Multiple currencies":money(feeTotal,feeCurrencies[0]||"INR");$("#kpiClaims").textContent=fmt(claims.length);const strF=filtered(cache.str||[],"invoice_date");$("#kpiInvoices").textContent=fmt(new Set(strF.map(x=>norm(x.invoice_no)).filter(Boolean)).size);$("#kpiPOs").textContent=fmt(new Set(strF.map(x=>norm(x.po_number)).filter(Boolean)).size);
    $("#statusBreakdown").innerHTML=[["Good",rt.good_qty],["Damaged",rt.damaged_qty],["Near expiry",rt.near_expiry_qty],["Expired",rt.expired_qty],["Short",rt.short_qty],["Not received",rt.not_received_qty],["Pending",pending]].map(([n,v])=>`<div class="status-row"><span>${n}</span><div class="bar"><i style="width:${shipped?Math.min(100,v/shipped*100):0}%"></i></div><b>${fmt(v)}</b></div>`).join("");
    $("#summaryTable").innerHTML=[["Removal orders",fmt(new Set(orders.map(o=>o.order_id)).size)],["Shipped quantity",fmt(shipped)],["Physical received",fmt(physical)],["Short quantity",fmt(rt.short_qty)],["Not received quantity",fmt(rt.not_received_qty)],["Pending final allocation",fmt(pending)],["Amazon removal fees",feeCurrencies.length>1?"Multiple currencies":money(feeTotal,feeCurrencies[0]||"INR")]].map(r=>`<tr><td>${r[0]}</td><td><b>${r[1]}</b></td></tr>`).join("");
    drawTrend();renderOrders();renderShipments();renderReceipts();renderFees();renderClaims();renderSAP();renderImports();renderEmployees();renderAudit();
  }
  function drawTrend(){
    const labels=["Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec","Jan","Feb","Mar"],year=2026;
    const ship=labels.map((_,i)=>cache.shipments.filter(s=>{const d=fyDate(s.shipment_date);return d&&d.getFullYear()===(i<9?year:year+1)&&d.getMonth()===(i+3)%12;}).reduce((a,s)=>a+num(s.shipped_qty),0));
    const rec=labels.map((_,i)=>cache.receipts.filter(r=>{const d=fyDate(r.receipt_date);return d&&d.getFullYear()===(i<9?year:year+1)&&d.getMonth()===(i+3)%12;}).reduce((a,r)=>a+["good_qty","damaged_qty","near_expiry_qty","expired_qty"].reduce((z,k)=>z+num(r[k]),0),0));
    if(trend)trend.destroy();const ctx=$("#trendChart");if(!ctx)return;trend=new Chart(ctx,{type:"bar",data:{labels,datasets:[{label:"Shipped",data:ship,backgroundColor:"#1769aa",borderRadius:5},{label:"Physically received",data:rec,backgroundColor:"#32a47a",borderRadius:5}]},options:{responsive:true,plugins:{legend:{position:"bottom"}},scales:{y:{beginAtZero:true,grid:{color:"#edf1f5"}},x:{grid:{display:false}}}}});
  }
  function strFor(orderId,sku){return cache.str.find(r=>(!orderId||norm(r.order_id)===norm(orderId))&&(!sku||norm(r.sku)===norm(sku)))||cache.str.find(r=>sku&&norm(r.sku)===norm(sku))||{};}
  function receiptSummary(s){const t=receiptTotals(s.id),physical=t.good_qty+t.damaged_qty+t.near_expiry_qty+t.expired_qty+t.rework_qty,allocated=physical+t.short_qty+t.not_received_qty;return {...t,physical,pending:Math.max(0,num(s.shipped_qty)-allocated)};}
  function renderOrders(){
    const q=norm($("#orderSearch")?.value);
    const rows=filtered(cache.orders,"request_date").filter(x=>{const st=strFor(x.order_id,x.sku);return [x.order_id,x.sku,x.fnsku,x.sku,st.transaction_id,st.po_number,st.delivery_no,st.billing_document,st.migo_no].some(v=>norm(v).includes(q));}).slice(0,5000);
    $("#ordersTable").innerHTML=rows.map(x=>{const st=strFor(x.order_id,x.sku),m=masterFor(x.sku,x.fnsku,null);const totals=cache.shipments.filter(z=>z.order_id===x.order_id&&((x.fnsku&&z.fnsku===x.fnsku)||(x.sku&&z.sku===x.sku))).reduce((a,z)=>{const t=receiptSummary(z);a.physical+=t.physical;a.short+=t.short_qty;a.notReceived+=t.not_received_qty;a.pending+=t.pending;return a;},{physical:0,short:0,notReceived:0,pending:0});return `<tr><td>${esc(x.order_id)}</td><td>${esc(x.sku)}</td><td>${esc(x.fnsku)}</td><td>${esc(m.material_code||"Not mapped")}</td><td>${esc(m.title||"—")}</td><td>${fmt(x.requested_qty)}</td><td>${fmt(x.cancelled_qty)}</td><td>${fmt(x.shipped_qty)}</td><td>${fmt(totals.physical)}</td><td>${fmt(totals.pending+totals.short+totals.notReceived)}</td><td>${esc(st.po_number||"Pending")}</td><td>${esc(st.delivery_no||"—")}</td><td>${esc(st.billing_document||"—")}</td><td>${esc(st.migo_no||"—")}</td><td><span class="tag">${esc(x.status||"—")}</span></td></tr>`;}).join("")||emptyRow(15);enhanceTable("ordersTable");
  }
  function renderShipments(){
    const q=norm($("#shipmentSearch")?.value);
    const rows=filtered(cache.shipments,"shipment_date").filter(x=>{const y=enrichShipment(x),st=strFor(x.order_id,y.sku);return [x.order_id,x.tracking_number,x.fnsku,x.id,y.sku,y.material_code,masterFor(y.sku,x.fnsku,null).material_code,masterFor(y.sku,x.fnsku,null).title,st.po_number,st.delivery_no,st.billing_document,st.migo_no].some(v=>norm(v).includes(q));}).slice(0,5000);
    $("#shipmentsTable").innerHTML=rows.map(x=>{const y=enrichShipment(x),st=strFor(x.order_id,y.sku),t=receiptSummary(x),m=masterFor(y.sku,x.fnsku,null);return `<tr><td>${esc(x.id.slice(0,8))}</td><td>${esc(x.order_id)}</td><td>${esc(x.tracking_number)} <button class="secondary" data-track="${esc(x.tracking_number)}" data-carrier="${esc(x.raw_data?.carrier||"")}">Track</button></td><td>${esc(m.material_code||"Not mapped")}</td><td>${esc(m.title||"—")}</td><td>${fmt(x.shipped_qty)}</td><td>${fmt(t.physical)}</td><td>${fmt(t.pending)}</td><td>${fmt(t.short_qty)}</td><td>${fmt(t.not_received_qty)}</td><td>${esc(st.po_number||"Pending")}</td><td>${esc(st.delivery_no||"—")}</td><td>${esc(st.billing_document||"—")}</td><td>${esc(st.migo_no||"—")}</td><td>${esc(x.amazon_status||"—")}</td><td>${esc(x.shipment_date||"—")}</td></tr>`;}).join("")||emptyRow(16);enhanceTable("shipmentsTable");
  }
  // Header sorting and per-column filters are available on each main tracker.
  const tableState={};
  function enhanceTable(tbodyId){
    const el=document.getElementById(tbodyId),table=el?.tagName==="TABLE"?el:el?.closest("table"),body=table?.tBodies[0];if(!table||!body)return;const head=table.tHead?.rows[0];if(!head)return;
    let state=tableState[tbodyId]||(tableState[tbodyId]={filters:{},sortCol:-1,asc:true});
    let filterRow=table.querySelector("thead tr.column-filters");
    if(!filterRow){filterRow=document.createElement("tr");filterRow.className="column-filters";[...head.cells].forEach((th,i)=>{const cell=document.createElement("th"),input=document.createElement("input");input.className="column-filter";input.placeholder="Filter…";input.setAttribute("aria-label",`Filter ${th.textContent}`);input.value=state.filters[i]||"";input.onclick=e=>e.stopPropagation();input.oninput=()=>{state.filters[i]=input.value;applyTableFilters(tbodyId);};cell.append(input);filterRow.append(cell);th.title="Click to sort";th.style.cursor="pointer";th.onclick=()=>{if(state.sortCol===i)state.asc=!state.asc;else{state.sortCol=i;state.asc=true;}sortTableRows(tbodyId);};});table.tHead.append(filterRow);}
    if(state.sortCol>=0)sortTableRows(tbodyId);else applyTableFilters(tbodyId);
  }
  function bodyOf(id){const el=document.getElementById(id);return el?.tagName==="TABLE"?el.tBodies[0]:el;}
  function applyTableFilters(id){const body=bodyOf(id),state=tableState[id];if(!body||!state)return;[...body.rows].forEach(row=>{const ok=Object.entries(state.filters).every(([i,v])=>!v||norm(row.cells[Number(i)]?.textContent).includes(norm(v)));row.hidden=!ok;});}
  function sortTableRows(id){const body=bodyOf(id),state=tableState[id];if(!body||!state||state.sortCol<0)return;const rows=[...body.rows];rows.sort((a,b)=>{const x=a.cells[state.sortCol]?.textContent.trim()||"",y=b.cells[state.sortCol]?.textContent.trim()||"",nx=Number(x.replace(/,/g,"")),ny=Number(y.replace(/,/g,""));const cmp=x!==""&&y!==""&&Number.isFinite(nx)&&Number.isFinite(ny)?nx-ny:x.localeCompare(y,undefined,{numeric:true,sensitivity:"base"});return state.asc?cmp:-cmp;});rows.forEach(r=>body.append(r));applyTableFilters(id);}
  // Add always-visible left/right controls to wide report tables so users can reach columns outside the viewport.
  function installHorizontalScrollControls(){
    $$(".table-wrap").forEach(wrap=>{
      if(wrap.querySelector(":scope > .table-scroll-controls")||!wrap.querySelector("table"))return;
      const controls=document.createElement("div");controls.className="table-scroll-controls";
      controls.innerHTML='<span class="scroll-hint">Scroll to view all columns</span><div class="scroll-buttons"><button type="button" class="secondary scroll-left" aria-label="Scroll report left" title="Scroll left">← Left</button><button type="button" class="secondary scroll-right" aria-label="Scroll report right" title="Scroll right">Right →</button></div>';
      controls.querySelector(".scroll-left").addEventListener("click",()=>wrap.scrollBy({left:-Math.max(280,Math.round(wrap.clientWidth*.75)),behavior:"smooth"}));
      controls.querySelector(".scroll-right").addEventListener("click",()=>wrap.scrollBy({left:Math.max(280,Math.round(wrap.clientWidth*.75)),behavior:"smooth"}));
      wrap.prepend(controls);
      const update=()=>{const max=wrap.scrollWidth-wrap.clientWidth;controls.classList.toggle("is-scrollable",max>8);controls.querySelector(".scroll-left").disabled=wrap.scrollLeft<=2;controls.querySelector(".scroll-right").disabled=wrap.scrollLeft>=max-2;};
      wrap.addEventListener("scroll",update,{passive:true});window.addEventListener("resize",update);requestAnimationFrame(update);
    });
  }
  function emptyRow(n){return `<tr><td colspan="${n}" class="muted">No records found.</td></tr>`;}
  function enrichShipment(s){
    const order=cache.orders.find(o=>o.order_id===s.order_id && ((!s.fnsku||o.fnsku===s.fnsku)||(!s.sku||o.sku===s.sku)));
    const maps=cache.mappings.filter(m=>(s.fnsku&&norm(m.fnsku)===norm(s.fnsku))||(!s.fnsku&&s.sku&&norm(m.sku)===norm(s.sku)));
    const unique=[...new Map(maps.map(m=>[`${norm(m.sku)}|${norm(m.material_code)}`,m])).values()];
    let sku=s.sku||order?.sku||"", material=s.material_code||sku||"";
    let mappingStatus=s.mapping_status||"unmapped";
    if(unique.length===1){sku=sku||unique[0].sku||"";material=material||sku||unique[0].material_code||"";mappingStatus="matched";}
    else if(unique.length>1){mappingStatus="conflict";}
    else if(sku||material||order?.sku){sku=sku||order?.sku||"";material=material||sku||order?.material_code||"";mappingStatus=material||sku?"order-detail":"partial";}
    return {...s,sku,material_code:material,mapping_status:mappingStatus};
  }
  function renderReceiptSelector(){
    const type=$("#receiptLookupType")?.value||"order",sel=$("#receiptLookupValue");if(!sel)return;const previous=sel.value;
    const vals=[...new Set(cache.shipments.map(s=>type==="order"?s.order_id:s.tracking_number).filter(Boolean))].sort((a,b)=>a.localeCompare(b,undefined,{numeric:true}));
    const list=$("#receiptLookupList");if(list)list.innerHTML=vals.map(v=>`<option value="${esc(v)}"></option>`).join("");sel.placeholder=type==="order"?"Type an Amazon Order ID to search":"Type a Tracking Number to search";if(!vals.includes(previous))sel.value="";renderReceiptLineEditor();
  }
  function renderReceiptLineEditor(){
    const type=$("#receiptLookupType")?.value||"order",value=$("#receiptLookupValue")?.value,box=$("#receiptLineEditor");if(!box)return;
    if(!value){box.innerHTML='<div class="notice">Select an Order ID or Tracking Number to load all associated material lines and shipped quantities.</div>';return;}
    if(!cache.shipments.some(s=>(type==="order"?s.order_id:s.tracking_number)===value)){box.innerHTML='<div class="notice">No match yet. Keep typing, then pick the Order ID or Tracking Number from the suggestions.</div>';return;}
    const lines=cache.shipments.filter(s=>(type==="order"?s.order_id:s.tracking_number)===value);
    box.innerHTML=lines.map(raw=>{const s=enrichShipment(raw),t=receiptTotals(s.id),used=Object.values(t).reduce((a,b)=>a+num(b),0),remaining=Math.max(0,num(s.shipped_qty)-used),mm=masterFor(s.sku,s.fnsku,null);return `<div class="receipt-line" data-shipment="${esc(s.id)}"><div class="receipt-line-head"><b>${esc(mm.title||'Description not available')}</b><span>SAP material ${esc(mm.material_code||'Not mapped')} · Tracking: ${esc(s.tracking_number||'—')}</span><strong>Amazon shipped: ${fmt(s.shipped_qty)} | Previously received: ${fmt(physOf(s.id))} | Short / Not received: ${fmt(snrOf(s.id))} | Remaining: ${fmt(remaining)}</strong><small>Amazon status: ${esc(s.amazon_status||'—')} (does not confirm physical receipt)</small></div><div class="receipt-qty-grid">${[['good_qty','Good'],['damaged_qty','Damaged'],['near_expiry_qty','Near Expiry'],['expired_qty','Expired'],['rework_qty','Rework'],['short_qty','Short'],['not_received_qty','Not Received']].map(([k,label])=>`<label>${label}<input type="number" min="0" step="1" value="0" data-qty="${k}" inputmode="numeric"></label>`).join('')}</div><div class="batch-summary-row"><span data-batch-summary>No batches added</span> <button type="button" class="secondary" data-open-batches>Add batches</button></div><div class="hidden" data-evidence-row><label>Evidence: photo or video (required for damaged, short or not received)<input type="file" data-evidence accept="image/*,video/*" multiple></label></div></div>`;}).join('')||'<div class="notice">No shipment lines found.</div>';
  }
  function renderReceipts(){
    renderReceiptSelector();
    $("#receiptTable").innerHTML=cache.shipments.map(raw=>{const s=enrichShipment(raw),t=receiptTotals(s.id),physical=t.good_qty+t.damaged_qty+t.near_expiry_qty+t.expired_qty+t.rework_qty,alloc=physical+t.short_qty+t.not_received_qty,pending=Math.max(0,num(s.shipped_qty)-alloc),m=masterFor(s.sku,s.fnsku,null);return `<tr><td>${esc(s.order_id)}</td><td>${esc(s.sku||"—")}</td><td>${esc(s.fnsku||"—")}</td><td>${esc(m.material_code||"Not mapped")}</td><td>${esc(s.tracking_number)}</td><td>${fmt(s.shipped_qty)}</td><td>${fmt(t.good_qty)}</td><td>${fmt(t.damaged_qty)}</td><td>${fmt(t.near_expiry_qty)}</td><td>${fmt(t.expired_qty)}</td><td>${fmt(physical)}</td><td>${fmt(t.short_qty)}</td><td>${fmt(t.not_received_qty)}</td><td><b>${fmt(pending)}</b></td><td><span class="tag">${esc(s.mapping_status)}</span></td></tr>`;}).join("")||emptyRow(15);enhanceTable("receiptTable");
  }
  function renderFees(){const fees=cache.fees;const currencies=[...new Set(fees.map(x=>x.currency).filter(Boolean))], total=fees.reduce((a,x)=>a+num(x.fee_amount),0),qty=fees.reduce((a,f)=>{const o=cache.orders.find(x=>x.order_id===f.order_id&&((f.fnsku&&x.fnsku===f.fnsku)||(f.sku&&x.sku===f.sku)));return a+num(o?.shipped_qty);},0);$("#feeTotal").textContent=currencies.length<=1?money(total,currencies[0]||"INR"):"Multiple currencies";$("#feeCount").textContent=fmt(fees.length);$("#feePerUnit").textContent=currencies.length===1&&qty?money(total/qty,currencies[0]):"—";$("#feesTable").innerHTML=fees.slice(0,500).map(f=>{const m=masterFor(f.sku,f.fnsku,null);return `<tr><td>${esc(f.order_id)}</td><td>${esc(m.material_code||"Not mapped")}</td><td>${esc(m.title||"—")}</td><td>${money(f.fee_amount,f.currency||"INR")}</td><td>${esc(f.currency||"—")}</td><td>${esc(f.fee_date||"—")}</td><td>${esc(f.source_file||"Amazon report")}</td></tr>`;}).join("")||emptyRow(7);}
  function renderClaims(){$("#claimsTable").innerHTML=cache.claims.map(c=>`<tr><td>${esc(c.order_id)}</td><td>${esc(c.claim_ref||"—")}</td><td>${esc(c.claim_type)}</td><td>${fmt(c.quantity)}</td><td>${esc(c.status)}</td><td>${c.amount==null?"—":money(c.amount)}</td></tr>`).join("")||emptyRow(6);}
  function poReceiptTotals(po,sku){const k=`${norm(po)}|${norm(sku)}`,t={good_qty:0,damaged_qty:0,near_expiry_qty:0,expired_qty:0,short_qty:0,not_received_qty:0};(cache.poReceipts||[]).forEach(p=>{if(`${norm(p.po_number)}|${norm(p.sku)}`===k)Object.keys(t).forEach(f=>t[f]+=num(p[f]));});return t;}
  function poLineSummary(po){const map=new Map();(cache.str||[]).filter(r=>String(r.po_number)===String(po)).forEach(r=>{const k=norm(r.sku);const g=map.get(k)||{sku:r.sku||"",material:masterFor(r.sku,null,r.asin).material_code||r.material_code||"",title:masterFor(r.sku,null,r.asin).title||"",sent:0};g.sent+=unitsOf(r);map.set(k,g);});return [...map.values()];}
  function poSearchOptions(){const list=$("#poLookupList");if(!list)return;const pos=[...new Set((cache.str||[]).filter(r=>r.po_number).map(r=>String(r.po_number)))].sort();list.innerHTML=pos.map(p=>`<option value="${esc(p)}"></option>`).join("");}
  function renderPOLineEditor(){
    const po=($("#poLookupValue")?.value||"").trim(),box=$("#poLineEditor");if(!box)return;
    if(!po){box.innerHTML='<div class="notice">Type a PO number to load its SKU lines and sent quantities.</div>';return;}
    const lines=poLineSummary(po);
    if(!lines.length){box.innerHTML='<div class="notice">No match yet. Keep typing, then pick the PO number from the suggestions.</div>';return;}
    box.innerHTML=lines.map(g=>{const t=poReceiptTotals(po,g.sku),used=Object.values(t).reduce((a,b)=>a+num(b),0);return `<div class="receipt-line" data-sku="${esc(g.sku)}"><div class="receipt-line-head"><b>${esc(g.title||'Description not available')}</b><span>SAP material ${esc(g.material||'Not mapped')}</span><strong>Sent: ${fmt(g.sent)} | Previously allocated: ${fmt(used)} | Remaining: ${fmt(Math.max(0,g.sent-used))}</strong></div><div class="receipt-qty-grid">${[['good_qty','Good'],['damaged_qty','Damaged'],['near_expiry_qty','Near Expiry'],['expired_qty','Expired'],['short_qty','Short'],['not_received_qty','Not Received']].map(([k,label])=>`<label>${label}<input type="number" min="0" step="1" value="0" data-qty="${k}" inputmode="numeric"></label>`).join('')}</div></div>`;}).join('');
  }
  async function handlePOReceipt(e){
    e.preventDefault();const po=($("#poLookupValue").value||"").trim();const lines=[...document.querySelectorAll("#poLineEditor .receipt-line")];
    if(!po||!lines.length){say("poReceiptMsg","Type a PO number and load its lines first.");return;}
    const receipt_date=e.target.elements.receipt_date.value,remarks=e.target.elements.remarks.value;let saved=0,errors=[];
    for(const line of lines){const sku=line.dataset.sku;const vals={};line.querySelectorAll("[data-qty]").forEach(i=>vals[i.dataset.qty]=num(i.value));const total=Object.values(vals).reduce((a,b)=>a+b,0);if(!total)continue;
      const sent=poLineSummary(po).find(g=>norm(g.sku)===norm(sku))?.sent||0;const prev=Object.values(poReceiptTotals(po,sku)).reduce((a,b)=>a+num(b),0);
      if(total+prev>sent){errors.push(`${sku}: allocation exceeds sent quantity ${sent}`);continue;}
      const {error}=await sb.from("po_receipts").insert({po_number:po,sku,receipt_date,remarks,...vals});
      if(error)errors.push(`${sku}: ${friendlyDbError(error)}`);else{saved++;await audit("po_receipt_saved","po_receipt",{po_number:po,sku,...vals});}}
    if(saved){say("poReceiptMsg",`${saved} line(s) saved.${errors.length?' Issues: '+errors.join(' | '):''}`,!errors.length);await loadAll();}else say("poReceiptMsg",errors.join(" | ")||"Enter at least one quantity against a PO line.");
  }
  function renderSAP(){
    const records=cache.str||[],groups=new Map(),dayMs=86400000,todayMs=Date.now();
    records.forEach(r=>{const key=r.po_number?`${norm(r.po_number)}|${norm(r.sku)}`:`NOPO|${r.id}`;if(!groups.has(key))groups.set(key,{key,po:r.po_number?String(r.po_number):"",sku:r.sku||"",material:masterFor(r.sku,null,r.asin).material_code||r.material_code||"",title:masterFor(r.sku,null,r.asin).title||"",delivery:new Set(),billing:new Set(),migo:new Set(),poDate:r.po_date||r.invoice_date||"",sent:0});const g=groups.get(key);if(r.delivery_no)g.delivery.add(String(r.delivery_no));if(r.billing_document)g.billing.add(String(r.billing_document));if(r.migo_no)g.migo.add(String(r.migo_no));const pd=r.po_date||r.invoice_date;if(pd&&(!g.poDate||pd<g.poDate))g.poDate=pd;g.sent+=unitsOf(r);});
    const rows=[...groups.values()].map(g=>{const t=g.po?poReceiptTotals(g.po,g.sku):{good_qty:0,damaged_qty:0,near_expiry_qty:0,expired_qty:0,short_qty:0,not_received_qty:0};const phys=t.good_qty+t.damaged_qty+t.near_expiry_qty+t.expired_qty+t.rework_qty,alloc=phys+t.short_qty+t.not_received_qty,pending=Math.max(0,g.sent-alloc),days=pending>0&&g.poDate?Math.max(0,Math.floor((todayMs-new Date(g.poDate).getTime())/dayMs)):0,status=!g.po?"Awaiting PO":(g.sent>0&&pending===0?"Complete":alloc===0?"Pending":"Partial");return {...g,phys,short:alloc-phys,pending,days,status};}).sort((a,b)=>b.days-a.days||a.po.localeCompare(b.po));
    const open=rows.filter(r=>r.status!=="Complete");
    const kpi=document.querySelector("#sapSummary");if(kpi)kpi.innerHTML=`<div class="kpi-grid compact"><article class="kpi"><span>PO lines</span><b>${fmt(rows.length)}</b></article><article class="kpi"><span>Open (not fully received)</span><b>${fmt(open.length)}</b></article><article class="kpi"><span>Pending units</span><b>${fmt(open.reduce((a,r)=>a+r.pending,0))}</b></article><article class="kpi"><span>Oldest pending (days)</span><b>${fmt(Math.max(0,...open.map(r=>r.days)))}</b></article></div>`;
    $("#sapTable").innerHTML=rows.map(r=>`<tr><td>${esc(r.po||"Awaiting PO")}</td><td>${esc([...r.delivery].join(", ")||"—")}</td><td>${esc([...r.billing].join(", ")||"—")}</td><td>${esc([...r.migo].join(", ")||"—")}</td><td>${esc(r.material||"Not mapped")}</td><td>${esc(r.title||"—")}</td><td>${esc(r.poDate||"—")}</td><td>${fmt(r.sent)}</td><td>${fmt(r.phys)}</td><td>${fmt(r.short)}</td><td><b>${fmt(r.pending)}</b></td><td>${r.pending>0?fmt(r.days):"—"}</td><td><span class="tag ${tagFor(r.status)}">${esc(r.status)}</span></td></tr>`).join("")||emptyRow(13);
    enhanceTable("sapTable");poSearchOptions();
    $("#sapExceptions").innerHTML=records.map(r=>({r,m:masterFor(r.sku,null,r.asin)})).filter(({r,m})=>!r.po_number||!r.delivery_no||!r.billing_document||!r.migo_no||!m.material_code).slice(0,1000).map(({r,m})=>`<tr><td>Amazon STR</td><td>${esc(r.transaction_id)}</td><td>${esc(m.material_code||"Not mapped")}</td><td>${esc(!m.material_code?"Missing SAP material (product not in the master)":!r.po_number?"PO pending":!r.delivery_no?"Delivery pending":!r.billing_document?"Billing document pending":"MIGO pending")}</td><td>${esc(r.irn_status||"—")}</td><td>Review</td></tr>`).join("")||emptyRow(6);
  }
  function renderImports(){$("#importsTable").innerHTML=cache.imports.map(x=>`<tr><td>${esc(x.file_name)}</td><td>${esc(x.report_type)}</td><td>${fmt(x.rows_read)}</td><td>${fmt(x.rows_accepted)}</td><td>${fmt(x.rejected_count||0)}</td><td>${esc((x.error_details||[]).slice(0,2).join(" | "))}</td><td>${esc(x.imported_by_email||"—")}</td><td>${esc(x.created_at?.slice(0,10)||"—")}</td></tr>`).join("")||emptyRow(8);}
  async function renderEmployees(){if(!["admin","manager"].includes(profile?.role))return;const {data}=await sb.from("profiles").select("id,email,full_name,role,allowed_pages").order("created_at");allProfiles=data||[];renderAccessPanel();$("#employeesTable").innerHTML=(data||[]).map(p=>`<tr><td>${esc(p.full_name)}</td><td>${esc(p.email)}</td><td>${esc(p.role)}</td><td>${p.id===session.user.id?"Current user":profile?.role!=="admin"?"View only":`<select data-role-user="${p.id}"><option ${p.role==="employee"?"selected":""}>employee</option><option ${p.role==="manager"?"selected":""}>manager</option><option ${p.role==="warehouse"?"selected":""}>warehouse</option><option ${p.role==="admin"?"selected":""}>admin</option></select>`}</td></tr>`).join("")||emptyRow(4);$$("[data-role-user]").forEach(el=>el.addEventListener("change",async()=>{const {error}=await sb.from("profiles").update({role:el.value}).eq("id",el.dataset.roleUser);if(error)alert(friendlyDbError(error));else{await audit("role_updated","profile",{user_id:el.dataset.roleUser,role:el.value});await renderEmployees();}}));}
  async function renderAudit(){if(!["admin","manager"].includes(profile?.role))return;const {data}=await sb.from("audit_logs").select("*,profiles(full_name,email)").order("created_at",{ascending:false}).limit(300);$("#auditTable").innerHTML=(data||[]).map(a=>`<tr><td>${esc(a.created_at?.replace("T"," ").slice(0,19))}</td><td>${esc(a.profiles?.full_name||a.profiles?.email||"—")}</td><td>${esc(a.action)}</td><td>${esc(a.entity)}</td><td>${esc(JSON.stringify(a.details||{}))}</td></tr>`).join("")||emptyRow(5);}
  function nav(page){$$(".page").forEach(p=>p.classList.add("hidden"));$("#"+page+"Page").classList.remove("hidden");$$(".nav-item").forEach(n=>n.classList.toggle("active",n.dataset.page===page));$("#pageTitle").textContent=({dashboard:"Dashboard",orders:"Removal Orders",shipments:"Shipment Tracker",receipts:"Warehouse Receipts",fees:"Amazon Removal Fees",sap:"PO & SAP Tracker",claims:"Amazon Claims",reconciliation:"Reconciliation",reports:"Reports & Imports",employees:"Employee Management",audit:"Audit Logs"})[page]||"Dashboard";}
  async function addImport(file,type,read,accepted){const {error}=await sb.from("import_batches").insert({file_name:file.name,report_type:type,rows_read:read,rows_accepted:accepted,rejected_count:importProblemCount,error_details:importProblems,imported_by:session.user.id,imported_by_email:session.user.email});if(error)throw error;await audit("report_imported","import_batch",{file:file.name,type,rows_read:read,rows_accepted:accepted,rejected_count:importProblemCount});}
  const importProblems=[];let importProblemCount=0;
  function masterFor(sku,fnsku,asin){const match=uniqueMap(cache.mappings,[{field:'fnsku',label:'FNSKU',value:fnsku},{field:'sku',label:'SKU',value:sku},{field:'asin',label:'ASIN',value:asin}].filter(k=>k.value));const mr=match.row||{};return {material_code:mr.material_code||'',title:mr.sap_description||mr.title||'',sapDesc:mr.sap_description||'',pack:num(mr.raw_data?.['Pack Of'])||0,mrp:num(mr.new_mrp)||num(mr.raw_data?.['MRP'])||0,cogs:num(mr.cogs)||num(mr.raw_data?.['COGS'])||0,conflict:match.conflict};}
  const packOfRow=r=>num(r.raw_data?.['Pack Of'])||num(masterFor(r.sku,null,r.asin).pack)||1;
  const unitsOf=r=>(r.sent_units!=null?num(r.sent_units):num(r.quantity)*packOfRow(r));
  const canEdit=()=>["admin","manager"].includes(profile?.role);
  function writeOpts(opts){return canEdit()?opts:{...opts,ignoreDuplicates:true};}
  function friendlyDbError(err){const m=String(err?.message||err||""),c=String(err?.code||"");if(c==="42501"||/row-level security|permission denied/i.test(m))return "You do not have permission for this action. Ask a manager or admin to do it.";if(c==="23505"||/duplicate key/i.test(m))return "This record already exists and was skipped.";return m;}
  function problem(file,row,reason){importProblemCount++;if(importProblems.length<100)importProblems.push(`${file.name} — row ${row}: ${reason}`);}
  function uniqueMap(rows,keys){for(const key of keys){const v=norm(key.value);if(!v)continue;const found=rows.filter(m=>norm(m[key.field])===v);if(found.length){const signatures=[...new Set(found.map(m=>`${norm(m.sku)}|${norm(m.material_code)}`))];return {row:signatures.length===1?found[0]:null,conflict:signatures.length>1,matchedBy:key.label};}}return {row:null,conflict:false,matchedBy:null};}
  async function importOrders(rows,file){let accepted=0,feeAccepted=0,skipped=0;for(let i=0;i<rows.length;i++){const r=rows[i];const x={order_id:pick(r,"order-id","order id","removal-order-id"),sku:pick(r,"sku"),fnsku:pick(r,"fnsku"),material_code:pick(r,"material-code","material code","material"),requested_qty:num(pick(r,"requested-quantity","requested qty")),cancelled_qty:num(pick(r,"cancelled-quantity","cancelled qty")),shipped_qty:num(pick(r,"shipped-quantity","shipped qty")),in_process_qty:num(pick(r,"in-process-quantity","in-process qty")),status:pick(r,"status","order-status"),request_date:dateISO(pick(r,"request-date","created-date")),raw_data:r};if(!x.order_id){problem(file,i+2,'Missing Order ID');continue;}const {data:ordRows,error}=await sb.from("removal_orders").upsert(x,writeOpts({onConflict:"order_id,sku,fnsku"})).select("id");if(error)problem(file,i+2,friendlyDbError(error));else if(ordRows?.length||canEdit())accepted++;else skipped++;
      const rawFee=pick(r,"removal-fee","removal fee"),currency=pick(r,"currency");if(String(rawFee).trim()!==""&&currency){const feeRow={order_id:x.order_id,sku:x.sku||"",fnsku:x.fnsku||"",fee_amount:num(rawFee),currency,fee_date:x.request_date,source_file:file.name,raw_data:r};const {data:feeRows,error:feeError}=await sb.from("amazon_removal_fees").upsert(feeRow,writeOpts({onConflict:"order_id,sku,fnsku"})).select("id");if(feeError)problem(file,i+2,`Fee: ${friendlyDbError(feeError)}`);else if(feeRows?.length||canEdit())feeAccepted++;}}
    await addImport(file,"removal_order_detail",rows.length,accepted);return {accepted,feeAccepted,skipped};}
  async function importShipments(rows,file){let accepted=0,skipped=0;const seqMap=new Map();for(let i=0;i<rows.length;i++){const r=rows[i],order_id=pick(r,"order-id","order id"),tracking_number=pick(r,"tracking-number","tracking number"),fnsku=pick(r,"fnsku"),sku=pick(r,"sku"),asin=pick(r,"asin");if(!order_id||!tracking_number){problem(file,i+2,'Missing Order ID or Tracking Number');continue;}const order=cache.orders.find(o=>o.order_id===order_id&&((fnsku&&o.fnsku===fnsku)||(sku&&o.sku===sku)));const match=uniqueMap(cache.mappings,[{field:'fnsku',label:'FNSKU',value:fnsku},{field:'sku',label:'SKU',value:sku},{field:'asin',label:'ASIN',value:asin}].filter(x=>x.value));const mapped=match.row;const x={order_id,tracking_number,fnsku:fnsku||'',sku:sku||order?.sku||mapped?.sku||'',material_code:order?.material_code||mapped?.material_code||null,mapping_status:match.conflict?'conflict':(sku||order?.sku||mapped)?'matched':'unmapped',shipped_qty:num(pick(r,"shipped-quantity","quantity-shipped","quantity")),amazon_status:pick(r,"status","shipment-status"),shipment_date:dateISO(pick(r,"shipment-date","shipped-date")),disposition:pick(r,"disposition"),raw_data:r};const sk=[x.order_id,x.tracking_number,x.sku,x.fnsku,x.shipped_qty].join("|");const seq=(seqMap.get(sk)||0)+1;seqMap.set(sk,seq);x.line_seq=seq;const {data:shipRows,error}=await sb.from("shipment_lines").upsert(x,writeOpts({onConflict:"order_id,tracking_number,sku,fnsku,shipped_qty,line_seq"})).select("id");if(error)problem(file,i+2,friendlyDbError(error));else if(shipRows?.length||canEdit())accepted++;else skipped++;}await addImport(file,"removal_shipment_detail",rows.length,accepted);return {accepted,skipped};}
  async function importMappings(rows,file){
    let accepted=0,updated=0,skipped=0;
    const cleanTitle=s=>String(s??"").replace(/\s*-?\s*INACTIVE\s*$/i,"").trim();
    const seen=new Set();
    const fields=["asin","sku","fnsku","material_code","title","sap_description","cogs","case_size","channel","combo_check","new_mrp","status"];
    for(let i=0;i<rows.length;i++){
      const r=rows[i];
      const rawTitle=pick(r,"title","product name","product title");
      const x={asin:pick(r,"asin","asin.1")||null,sku:pick(r,"sku")||null,fnsku:pick(r,"fnsku")||null,material_code:pick(r,"material code","material-code","material_code")||null,title:cleanTitle(rawTitle)||null,sap_description:String(pick(r,"sap description","sap material description","material description","sap desc","description")||"").trim()||null,cogs:(()=>{const v=pick(r,"cogs","cogs per unit","cogs/unit","cost");return v===""||v===null||v===undefined?null:num(v);})(),case_size:num(pick(r,"case size","case_size"))||null,channel:pick(r,"channel")||null,combo_check:pick(r,"combo check","combo_check")||null,new_mrp:pick(r,"new mrp","new_mrp","mrp")?num(pick(r,"new mrp","new_mrp","mrp")):null,mapping_version:file.name,raw_data:r};
      if(!x.fnsku&&!x.sku&&!x.asin){problem(file,i+2,'No ASIN, SKU or FNSKU identifier');continue;}
      const explicit=pick(r,"status");
      x.status=explicit?explicit:(/inactive/i.test(String(rawTitle??""))?"Inactive":(!x.fnsku?"Pending FNSKU":"Active"));
      const key=x.fnsku?`F|${norm(x.fnsku)}`:`A|${norm(x.asin)}`;
      if(seen.has(key)){skipped++;continue;}
      seen.add(key);
      const list=cache.mappings||[];
      let existing=x.fnsku?list.find(m=>norm(m.fnsku)===norm(x.fnsku)):null;
      if(!existing&&norm(x.asin))existing=list.find(m=>!norm(m.fnsku)&&norm(m.asin)===norm(x.asin));
      if(existing){
        const changed=fields.some(f=>String(x[f]??"")!==String(existing[f]??""))||String(x.mapping_version)!==String(existing.mapping_version??"");
        if(!changed){skipped++;continue;}
        if(!canEdit()){skipped++;continue;}
        const {error}=await sb.from("product_mappings").update(x).eq("id",existing.id);
        if(error)problem(file,i+2,friendlyDbError(error));else{updated++;Object.assign(existing,x);}
        continue;
      }
      const {data,error}=await sb.from("product_mappings").insert(x).select("id");
      if(error)problem(file,i+2,friendlyDbError(error));
      else{accepted++;list.push({...x,id:data?.[0]?.id});cache.mappings=list;}
    }
    await addImport(file,"product_mapping_master",rows.length,accepted+updated);
    return {accepted,updated,skipped};
  }
  async function importSTR(rows,file,label="str_fc_removal_only"){
    const eligible=rows.filter(r=>norm(pick(r,"Transaction Type","transaction-type","transaction type"))==="fc_removal");let accepted=0,skipped=0,updated=0;
    const strKeyOf=o=>[norm(o.transaction_id),norm(o.invoice_no),norm(o.asin),num(o.quantity)].join('|');
    const strIndex=new Map((cache.str||[]).map(e=>[strKeyOf(e),e]));
    for(let i=0;i<eligible.length;i++){const r=eligible[i],transaction_id=pick(r,"Transaction Id","transaction-id","transaction id"),sourceSku=pick(r,"Sku","SKU"),asin=pick(r,"Asin","ASIN"),fnsku=pick(r,"FNSKU");if(!transaction_id){problem(file,i+2,'FC_REMOVAL row has no Transaction ID');continue;}
      const match=uniqueMap(cache.mappings,[{field:'sku',label:'SKU',value:sourceSku},{field:'fnsku',label:'FNSKU',value:fnsku},{field:'asin',label:'ASIN',value:asin}].filter(x=>x.value));const mapped=match.row;const sku=sourceSku||mapped?.sku||null;const material_code=pick(r,"Material Code","material-code","material code")||sourceSku||mapped?.sku||mapped?.material_code||null;const invoice=pick(r,"Invoice Number","invoice-number","invoice no"),qty=pick(r,"Quantity","transaction-quantity");
      const source_line_key=[transaction_id,invoice,sourceSku,asin,fnsku,qty,i].map(v=>encodeURIComponent(String(v??''))).join('|');
      const x={source_line_key,transaction_id,invoice_no:invoice,material_code,transaction_date:dateISO(pick(r,"Invoice Date","invoice date","Date","transaction-date")),invoice_date:dateISO(pick(r,"Invoice Date","invoice date")),invoice_value:pick(r,"Invoice Value","invoice value")!==""?num(pick(r,"Invoice Value","invoice value")):null,taxable_value:pick(r,"Taxable Value","taxable value")!==""?num(pick(r,"Taxable Value","taxable value")):null,quantity:num(qty),sent_units:num(qty)*(num(pick(r,'Pack Of','pack of'))||num(match.row?.raw_data?.['Pack Of'])||1),hsn:pick(r,"Hsn Code","hsn code","hsn"),gst_rate:num(pick(r,"Igst Rate","igst rate","gst-rate"))||null,transaction_type:"FC_REMOVAL",ship_from_fc:pick(r,"Ship From Fc","ship from fc"),ship_from_storage_location:pick(r,"SAP Storage location"),ship_from_plant_code:pick(r,"SAP plant code'"),sku,asin,order_id:pick(r,"Order Id","order id"),receiver_gstin:pick(r,"Gstin Of Receiver","gstin of receiver"),actual_qty:num(pick(r,"ACTUAL QTY","actual qty")),invoice_quantity:num(pick(r,"Invoice Quantity","invoice quantity")),irn_status:pick(r,"Irn Filing Status","irn filing status"),po_number:pick(r,"Po Number","po number","SAP PO NO","SAP PO No","PO NO"),delivery_no:pick(r,"Delivery No","delivery no","DELIVERY NO"),billing_document:pick(r,"Billing Document","billing document","INVOICE NO","INVOICE No"),migo_no:pick(r,"MIGO No","migo no","MIGO NO"),po_date:dateISO(pick(r,"created date","Created Date","PO Created Date","po created date")),sap_status:pick(r,"Status","status"),mapping_status:match.conflict?'conflict':(sourceSku||mapped)?'matched':'unmapped',raw_data:r};
      const existing=strIndex.get(strKeyOf(x));
      if(existing){const patch={};["po_number","delivery_no","billing_document","migo_no","sap_status","irn_status","po_date"].forEach(k=>{if(x[k]!==undefined&&x[k]!==null&&String(x[k]).trim()!==""&&norm(x[k])!==norm(existing[k]))patch[k]=String(x[k]);});["sku","material_code"].forEach(k=>{if(x[k]&&String(x[k]).trim()!==""&&!norm(existing[k]))patch[k]=String(x[k]);});if(num(x.invoice_quantity)>0&&num(x.invoice_quantity)!==num(existing.invoice_quantity))patch.invoice_quantity=num(x.invoice_quantity);if(num(x.sent_units)>0&&num(x.sent_units)!==num(existing.sent_units))patch.sent_units=num(x.sent_units);
        if(!Object.keys(patch).length)skipped++;else if(!canEdit())skipped++;else{const {error:uErr}=await sb.from("str_records").update(patch).eq("id",existing.id);if(uErr)problem(file,i+2,friendlyDbError(uErr));else{updated++;Object.assign(existing,patch);}}}
      else{const {data:strRows,error}=await sb.from("str_records").upsert(x,writeOpts({onConflict:"source_line_key"})).select("id");if(error)problem(file,i+2,friendlyDbError(error));else if(strRows?.length){accepted++;strIndex.set(strKeyOf(x),{...x,id:strRows[0].id});}else skipped++;}
    }
    await addImport(file,label,rows.length,accepted);return {eligible:eligible.length,accepted,skipped,updated};
  }
  async function importFees(rows,file,cols={}){let accepted=0;for(const r of rows){const order=pick(r,cols.order||"order-id","order-id","removal-order-id","order id"),fee=pick(r,cols.fee||"fee-amount","fee-amount","amount","removal-fee"),currency=pick(r,cols.currency||"currency")||"INR",date=pick(r,cols.date||"posted-date","posted-date","fee-date")||null;if(!order||fee==="")continue;const {data:ord}=await sb.from("removal_orders").select("order_id").eq("order_id",order).limit(1);if(!ord?.length)continue;const {data:fr,error}=await sb.from("amazon_removal_fees").upsert({order_id:order,sku:"",fnsku:"",fee_amount:num(fee),currency,fee_date:date,source_file:file.name,raw_data:r},writeOpts({onConflict:"order_id,sku,fnsku"})).select("id");if(!error&&(fr?.length||canEdit()))accepted++;}await addImport(file,"amazon_removal_fees_only",rows.length,accepted);return accepted;}
  async function handleAmazonImport(e){e.preventDefault();const fd=new FormData(e.target),file=fd.get("file"),type=fd.get("type");try{importProblems.length=0;importProblemCount=0;const rows=await readReportFile(file);let result={};if(type==="orders")result=await importOrders(rows,file);if(type==="shipments")result=await importShipments(rows,file);if(type==="str")result=await importSTR(rows,file);if(type==="str_master")result=await importSTR(rows,file,"str_master_sap");if(type==="fee")result={accepted:await importFees(rows,file)};if(type==="mapping")result=await importMappings(rows,file);if(type==="stock")result=await importStock(rows,file);const rejected=importProblemCount;const detail=importProblems.slice(0,8).join('\n');say("importMsg",`Processed ${rows.length} rows from ${file.name}. ${result?.eligible!=null?`FC_REMOVAL rows: ${result.eligible}. `:""}Accepted: ${result?.accepted??0}.${result?.skipped?` Already present and skipped: ${result.skipped}.`:""}${result?.updated?` Updated existing rows: ${result.updated}.`:""} Rejected / needs review: ${rejected}.${result?.feeAccepted!=null?` Removal fee lines captured: ${result.feeAccepted}.`:""}${detail?'\n'+detail:''}${rejected>8?'\nMore row errors are available in the browser console.':''}`,rejected===0);if(rejected>8)console.warn('Import row errors',importProblems);await loadAll();}catch(err){say("importMsg",err.message||String(err));}}
  async function handleFeeUpload(e){e.preventDefault();const fd=new FormData(e.target),file=fd.get("file");try{const rows=await readReportFile(file);const accepted=await importFees(rows,file,{order:fd.get("order_col"),fee:fd.get("fee_col"),currency:fd.get("currency_col"),date:fd.get("date_col")});say("feeMsg",`Imported ${accepted} fee rows linked to known Amazon removal orders.`,true);await loadAll();}catch(err){say("feeMsg",err.message);}}

  const BATCH_CONDS=[["good","Good"],["damaged","Damaged"],["rework","Rework"]];
  let batchCtx=null;
  function batchRowHtml(r){return `<tr><td><select data-bcond>${BATCH_CONDS.map(([k,l])=>`<option value="${k}" ${r.cond===k?"selected":""}>${l}</option>`).join("")}</select></td><td><input data-bbatch value="${esc(r.batch||"")}" placeholder="e.g. VU0001"></td><td><input data-bqty type="number" min="0" step="1" value="${num(r.qty)}"></td><td><input data-bmfg type="date" value="${esc(r.mfg||"")}"></td><td><input data-bexp type="date" value="${esc(r.exp||"")}"></td><td><button type="button" class="secondary" data-bdel>Remove</button></td></tr>`;}
  function refreshBatchSummary(line){const rows=JSON.parse(line.dataset.batches||"[]");const el=line.querySelector("[data-batch-summary]");if(el)el.textContent=rows.length?`${rows.length} batch row(s): `+rows.map(r=>`${r.batch||"no batch"} ${r.cond} ${r.qty}`).join(", "):"No batches added";}
  function openBatchModal(line){
    if(!line)return;
    const s=cache.shipments.find(x=>x.id===line.dataset.shipment);if(!s)return;
    batchCtx={line,s};
    let rows=[];try{rows=JSON.parse(line.dataset.batches||"[]");}catch(_){rows=[];}
    if(!rows.length){BATCH_CONDS.forEach(([k])=>{const q=num(line.querySelector(`[data-qty="${k}_qty"]`)?.value);if(q>0)rows.push({cond:k,batch:"",qty:q,mfg:"",exp:""});});}
    if(!rows.length)rows.push({cond:"good",batch:"",qty:0,mfg:"",exp:""});
    {const mx=masterFor(s.sku,s.fnsku,null);$("#batchModalTitle").textContent=`${s.order_id} · ${s.tracking_number} · ${mx.material_code||"Not mapped"} ${mx.title||""} · shipped ${s.shipped_qty}`;}
    $("#batchTable tbody").innerHTML=rows.map(batchRowHtml).join("");
    say("batchMsg","");
    $("#batchModal").classList.remove("hidden");
  }
  function saveBatchModal(){
    if(!batchCtx)return;const {line}=batchCtx;
    const rows=[...document.querySelectorAll("#batchTable tbody tr")].map(tr=>({cond:tr.querySelector("[data-bcond]").value,batch:tr.querySelector("[data-bbatch]").value.trim(),qty:num(tr.querySelector("[data-bqty]").value),mfg:tr.querySelector("[data-bmfg]").value,exp:tr.querySelector("[data-bexp]").value})).filter(r=>r.qty>0);
    const sums={good:0,damaged:0,rework:0};rows.forEach(r=>sums[r.cond]+=r.qty);
    BATCH_CONDS.forEach(([k])=>{const inp=line.querySelector(`[data-qty="${k}_qty"]`);if(inp)inp.value=sums[k];});
    line.dataset.batches=JSON.stringify(rows);
    refreshBatchSummary(line);
    $("#batchModal").classList.add("hidden");batchCtx=null;
  }
  async function handleReceipt(e){
    e.preventDefault();const lines=[...document.querySelectorAll("#receiptLineEditor .receipt-line")];if(!lines.length){say("receiptMsg","Select an Order ID or Tracking Number first.");return;}
    const receipt_date=e.target.elements.receipt_date.value,remarks=e.target.elements.remarks.value;let saved=0,errors=[];
    for(const line of lines){
      const s=cache.shipments.find(x=>x.id===line.dataset.shipment);if(!s)continue;
      const vals={};line.querySelectorAll("[data-qty]").forEach(i=>vals[i.dataset.qty]=num(i.value));
      const total=Object.values(vals).reduce((acc,b)=>acc+b,0);if(!total)continue;
      const needEv=num(vals.damaged_qty)+num(vals.short_qty)+num(vals.not_received_qty)>0,evFiles=line.querySelector("[data-evidence]")?.files||[];
      if(needEv){if(!evFiles.length){errors.push(`${s.order_id} / ${s.tracking_number}: attach a photo or video for the damaged, short or not received quantity`);continue;}const ee=checkEvidenceFiles(evFiles);if(ee.length){errors.push(`${s.order_id} / ${s.tracking_number}: ${ee.join(", ")}`);continue;}}
      let rows=[];try{rows=JSON.parse(line.dataset.batches||"[]");}catch(_){rows=[];}
      let entries=[];
      if(rows.length){
        const sumBy={good:0,damaged:0,rework:0};rows.forEach(r=>sumBy[r.cond]+=num(r.qty));
        if(BATCH_CONDS.some(([k])=>num(vals[`${k}_qty`])!==sumBy[k])){errors.push(`${s.order_id} / ${s.tracking_number}: batch quantities do not match the entered quantities. Open Add batches and save again.`);continue;}
        rows.filter(r=>r.qty>0).forEach(r=>{
          const c=r.cond==="good"?classifyGood(num(r.qty),receipt_date,r.exp||null):{good:0,near:0,expired:0};
          entries.push({good_qty:c.good,near_expiry_qty:c.near,expired_qty:c.expired,damaged_qty:r.cond==="damaged"?num(r.qty):0,rework_qty:r.cond==="rework"?num(r.qty):0,short_qty:0,not_received_qty:0,batch_no:r.batch||null,mfg_date:r.mfg||null,exp_date:r.exp||null});
        });
        const other={short_qty:vals.short_qty||0,not_received_qty:vals.not_received_qty||0,near_expiry_qty:vals.near_expiry_qty||0,expired_qty:vals.expired_qty||0};
        if(Object.values(other).some(v=>v>0))entries.push({...other,good_qty:0,damaged_qty:0,rework_qty:0,batch_no:null,mfg_date:null,exp_date:null});
      }else{
        entries.push({...vals,batch_no:null,mfg_date:null,exp_date:null});
      }
      const newTotal=entries.reduce((acc,en)=>acc+Object.values(en).filter(v=>typeof v==="number").reduce((x,y)=>x+y,0),0);
      const prev=Object.values(receiptTotals(s.id)).reduce((acc,b)=>acc+num(b),0);
      if(newTotal+prev>num(s.shipped_qty)){errors.push(`${s.order_id} / ${s.tracking_number}: allocation ${newTotal+prev} exceeds shipped quantity ${s.shipped_qty}`);continue;}
      let lineOk=true;
      for(const en of entries){const {error}=await sb.from("warehouse_receipts").insert({shipment_line_id:s.id,...en,receipt_date,remarks});if(error){lineOk=false;errors.push(`${s.order_id} / ${s.tracking_number}: ${friendlyDbError(error)}`);break;}saved++;}
      if(lineOk){await audit("receipt_saved","warehouse_receipt",{shipment_id:s.id,rows:entries.length,total:newTotal});if(needEv){const up=await uploadEvidence(s.id,evFiles);if(up.errs.length)errors.push(`${s.order_id} / ${s.tracking_number}: receipt saved but evidence failed (${up.errs.join(", ")}). Add it from the Saved receipts list.`);}}
    }
    if(saved){say("receiptMsg",`${saved} receipt row(s) saved.${errors.length?' Issues: '+errors.join(' | '):''}`,!errors.length);await loadAll();}
    else say("receiptMsg",errors.join(" | ")||"Enter at least one quantity against a shipment line.");
  }
  async function handleClaim(e){e.preventDefault();const fd=new FormData(e.target),x=Object.fromEntries(fd.entries());x.quantity=num(x.quantity);x.amount=x.amount===""?null:num(x.amount);const {error}=await sb.from("amazon_claims").insert(x);if(error)alert(friendlyDbError(error));else{e.target.reset();await loadAll();}}
  function exportTable(tableId,fileName){const table=$("#"+tableId);if(!table)return;const rows=[...table.querySelectorAll("tr")].filter(tr=>!tr.hidden&&!tr.classList.contains("column-filters")).map(tr=>[...tr.children].map(td=>(td.innerText||"").trim()));if(window.XLSX){const wb=XLSX.utils.book_new(),ws=XLSX.utils.aoa_to_sheet(rows);XLSX.utils.book_append_sheet(wb,ws,"Report");XLSX.writeFile(wb,fileName.replace(/\.csv$/i,".xlsx"));}else{const csv=rows.map(r=>r.map(v=>`"${String(v).replace(/"/g,'""')}"`).join(",")).join("\r\n");const blob=new Blob(["\uFEFF"+csv],{type:"text/csv;charset=utf-8;"}),url=URL.createObjectURL(blob),a=document.createElement("a");a.href=url;a.download=fileName;a.click();URL.revokeObjectURL(url);}}
  function bind(){
    $("#loginBtn").onclick=async()=>{const {error}=await sb.auth.signInWithPassword({email:$("#loginEmail").value,password:$("#loginPassword").value});say("loginMsg",error?.message||"Signing in…",!error);};
    $("#logoutBtn").onclick=()=>sb.auth.signOut();
    $$("#nav [data-page]").forEach(b=>b.onclick=()=>nav(b.dataset.page));$$("[data-goto]").forEach(b=>b.onclick=()=>nav(b.dataset.goto));
    $("#amazonImportForm").onsubmit=handleAmazonImport;$("#receiptForm").onsubmit=handleReceipt;$("#invoiceImportForm").onsubmit=async e=>{e.preventDefault();const files=e.target.elements.files.files;if(!files.length)return;try{await importInvoiceFiles(files);}catch(err){say("invoiceImportMsg",friendlyDbError(err));}};$("#reconStatus").onchange=renderRecon;document.addEventListener("click",e=>{const b=e.target.closest&&e.target.closest("[data-track]");if(b)trackShipment(b.dataset.track,b.dataset.carrier);});$("#receiptEditForm").onsubmit=saveReceiptEdit;bindBatchModal();bindApprovals();bindScat();bindEvidence();$("#fcMapForm").onsubmit=saveFcMap;$("#fcMapTable").onclick=e=>{const d=e.target.closest&&e.target.closest("[data-fc-del]");if(d)deleteFcMap(d.dataset.fcDel);};$("#warehouseImportForm").onsubmit=async e=>{e.preventDefault();const f=e.target.elements.file.files[0];if(!f)return;try{await importWarehouseExcel(f,e.target.elements.receipt_date.value);}catch(err){say("warehouseImportMsg",err.message||String(err));}};$("#warehouseImportForm").elements.receipt_date.value=today();$("#accessUser").onchange=refreshAccessChecks;$("#accessForm").onsubmit=saveAccess;$("#accessReset").onclick=resetAccess;$("#receiptEditCancel").onclick=()=>$("#receiptEditPanel").classList.add("hidden");$("#productForm").onsubmit=saveProduct;$("#productCancel").onclick=()=>$("#productFormPanel").classList.add("hidden");$("#productAddBtn").onclick=()=>openProductForm("");$("#productSearch").oninput=renderProducts;$("#productsTable").onclick=e=>{const ed=e.target.closest("[data-edit-product]");if(ed){openProductForm(ed.dataset.editProduct);return;}const dl=e.target.closest("[data-delete-product]");if(dl)deleteProduct(dl.dataset.deleteProduct);};$("#receiptEntriesTable").onclick=e=>{const b=e.target.closest("[data-edit-receipt]");if(b)openReceiptEdit(b.dataset.editReceipt);};document.addEventListener("click",e=>{if(e.target.closest&&e.target.closest(".nav-item"))setTimeout(freezeAll,80);});window.addEventListener("resize",()=>freezeAll());$("#reconSearch").oninput=renderRecon;$("#exportRecon").onclick=()=>exportTable("reconTable","reconciliation.xlsx");$("#poReceiptForm").onsubmit=handlePOReceipt;$("#poLookupValue").oninput=renderPOLineEditor;$("#poReceiptForm").elements.receipt_date.value=today();$("#claimForm").onsubmit=handleClaim;
    $("#orderSearch").oninput=renderOrders;$("#shipmentSearch").oninput=renderShipments;$("#exportOrders").onclick=()=>exportTable("ordersTable","amazon_removal_orders.xlsx");$("#exportShipments").onclick=()=>exportTable("shipmentsTable","amazon_shipment_tracker.xlsx");$("#exportReceipts").onclick=()=>exportTable("receiptTable","warehouse_receipt_tracker.xlsx");$("#exportSAP").onclick=()=>exportTable("sapTable","sap_po_tracker.xlsx");$("#monthFilter").onchange=renderAll;$("#fyFilter").onchange=renderAll;
    
    $("#receiptForm").elements.receipt_date.value=today();$("#receiptLookupType").addEventListener("change",renderReceiptSelector);$("#receiptLookupValue").addEventListener("input",renderReceiptLineEditor);
  }
  // ---------- Invoices (bulk ZIP) and reconciliation ----------
  const libLoads={};
  function loadScriptOnce(url){if(!libLoads[url])libLoads[url]=new Promise((res,rej)=>{const s=document.createElement("script");s.src=url;s.onload=res;s.onerror=()=>rej(new Error("Could not load "+url));document.head.appendChild(s);});return libLoads[url];}
  async function ensureJSZip(){await loadScriptOnce("https://cdnjs.cloudflare.com/ajax/libs/jszip/3.10.1/jszip.min.js");return window.JSZip;}
  async function ensurePdfJs(){await loadScriptOnce("https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.min.js");window.pdfjsLib.GlobalWorkerOptions.workerSrc="https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js";return window.pdfjsLib;}
  async function sha256Hex(buf){const d=await crypto.subtle.digest("SHA-256",buf);return [...new Uint8Array(d)].map(b=>b.toString(16).padStart(2,"0")).join("");}
  async function pdfText(bytes,pdfjs){const doc=await pdfjs.getDocument({data:bytes}).promise;let out="";for(let p=1;p<=doc.numPages;p++){const page=await doc.getPage(p);const c=await page.getTextContent();out+=c.items.map(i=>i.str).join(" ")+"\n";}return out;}
  function invoiceDateISO(s){const m=/(\d{2})\.(\d{2})\.(\d{4})/.exec(s||"");return m?`${m[3]}-${m[2]}-${m[1]}`:null;}
  // RA rule: text between the first and second hyphen (ZAXQT-<key>-<location>-<suffix>).
  // Internal hyphens inside the key are kept; the last two parts (location, suffix) are dropped.
  function raKeyOf(ra){const p=String(ra||"").split("-");if(p.length<4||p[0].toUpperCase()!=="ZAXQT")return {key:null,ok:false};return {key:p.slice(1,p.length-2).join("-"),ok:true};}
  function parseInvoiceText(t){
    const g=re=>{const m=re.exec(t);return m?m[1].trim():null;};
    const invoice_no=g(/Invoice Number\s*:?\s*([A-Z0-9]+-T-\d+)/i)||g(/\b([A-Z0-9]{2,6}-T-\d+)\b/);
    const shipment_id=g(/Shipment ID\s*:?\s*(\d{6,})/i);
    const ra_number=g(/RA No\.?\s*:?\s*(\S+)/i);
    const rk=raKeyOf(ra_number);
    return {invoice_no,invoice_date:invoiceDateISO(g(/Invoice Date\s*:?\s*(\d{2}\.\d{2}\.\d{4})/i)),shipment_id,ra_number,ra_key:rk.key,ra_ok:rk.ok,order_number:g(/Order Number\s*:?\s*(\S+)/i),asin:g(/\b(B0[A-Z0-9]{8})\b/)};
  }
  async function importInvoiceFiles(files){
    const msg=$("#invoiceImportMsg"),totals=$("#invoiceImportTotals");
    const T={read:0,accepted:0,duplicates:0,failed:0,unmatched:0,zipSkipped:0};const problems=[];
    say("invoiceImportMsg","Reading invoices...",true);totals.innerHTML="";
    const pdfjs=await ensurePdfJs();
    const seenZip=new Set((cache.zipImports||[]).map(z=>z.file_hash));
    const knownInv=new Set((cache.invoices||[]).map(i=>norm(i.invoice_no)));
    const knownOrders=new Set((cache.orders||[]).map(o=>norm(o.order_id)));
    for(const file of [...files]){
      const isZip=/\.zip$/i.test(file.name);
      const buf=await file.arrayBuffer();
      const fileHash=await sha256Hex(buf);
      if(seenZip.has(fileHash)){T.zipSkipped++;problems.push(`${file.name}: already uploaded, skipped`);continue;}
      const pdfs=[];
      if(isZip){const JSZip=await ensureJSZip();const z=await JSZip.loadAsync(buf);for(const e of Object.values(z.files)){if(!e.dir&&/\.pdf$/i.test(e.name))pdfs.push({name:e.name.split("/").pop(),bytes:await e.async("uint8array")});}}
      else pdfs.push({name:file.name,bytes:new Uint8Array(buf)});
      let zAcc=0,zDup=0,zFail=0,zUnm=0;
      for(const p of pdfs){
        T.read++;
        try{
          const f=parseInvoiceText(await pdfText(p.bytes,pdfjs));
          if(!f.invoice_no){T.failed++;zFail++;problems.push(`${p.name}: invoice number not found`);continue;}
          if(knownInv.has(norm(f.invoice_no))){T.duplicates++;zDup++;continue;}
          const matched=!!f.ra_key&&knownOrders.has(norm(f.ra_key));
          if(!matched){T.unmatched++;zUnm++;}
          const row={invoice_no:f.invoice_no,invoice_date:f.invoice_date,shipment_id:f.shipment_id,ra_number:f.ra_number,ra_key:f.ra_key,order_number:f.order_number,asin:f.asin,source_zip:isZip?file.name:null,source_pdf:p.name,file_hash:fileHash,match_status:!f.ra_ok?"ra_format_review":matched?"matched_order":"unmatched"};
          const {data,error}=await sb.from("amazon_invoices").upsert(row,{onConflict:"invoice_no",ignoreDuplicates:true}).select("id");
          if(error){T.failed++;zFail++;problems.push(`${p.name}: ${friendlyDbError(error)}`);continue;}
          if(!data?.length){T.duplicates++;zDup++;continue;}
          T.accepted++;zAcc++;knownInv.add(norm(f.invoice_no));
        }catch(err){T.failed++;zFail++;problems.push(`${p.name}: ${err.message||err}`);}
      }
      const {error:zErr}=await sb.from("invoice_zip_imports").insert({file_hash:fileHash,file_name:file.name,pdf_count:pdfs.length,accepted:zAcc,duplicates:zDup,failed:zFail,unmatched:zUnm});
      if(zErr)problems.push(`${file.name}: import log not saved (${friendlyDbError(zErr)})`);
      seenZip.add(fileHash);
    }
    await loadAll();
    say("invoiceImportMsg",`Done. Read ${T.read}, accepted ${T.accepted}, duplicates ${T.duplicates}, failed ${T.failed}, unmatched ${T.unmatched}.`,T.failed===0);
    totals.innerHTML=`<div class="kpi-grid compact"><article class="kpi"><span>Invoices read</span><b>${fmt(T.read)}</b></article><article class="kpi"><span>Accepted (new)</span><b>${fmt(T.accepted)}</b></article><article class="kpi"><span>Updated</span><b>0</b><small>Invoices are not changed after first import</small></article><article class="kpi"><span>Duplicates skipped</span><b>${fmt(T.duplicates)}</b></article><article class="kpi"><span>Failed</span><b>${fmt(T.failed)}</b></article><article class="kpi"><span>Unmatched</span><b>${fmt(T.unmatched)}</b><small>RA key has no removal order yet</small></article><article class="kpi"><span>ZIP files already uploaded</span><b>${fmt(T.zipSkipped)}</b></article></div>${problems.length?`<p class="muted">${problems.map(esc).join("<br>")}</p>`:""}`;
  }
  function reconRows(){
    const orders=cache.orders||[],ships=cache.shipments||[],recs=cache.receipts||[],fees=cache.fees||[],invs=cache.invoices||[],strs=cache.str||[],maps=cache.mappings||[];
    const shipBy=new Map();ships.forEach(s=>{const k=`${norm(s.order_id)}|${norm(s.fnsku)}`;shipBy.set(k,[...(shipBy.get(k)||[]),s]);});
    const recBy=new Map();recs.forEach(r=>recBy.set(r.shipment_line_id,[...(recBy.get(r.shipment_line_id)||[]),r]));
    const feeBy=new Map();fees.forEach(f=>{const k=`${norm(f.order_id)}|${norm(f.fnsku)}`;feeBy.set(k,[...(feeBy.get(k)||[]),f]);});
    const asinsFor=fn=>new Set(maps.filter(m=>norm(m.fnsku)===norm(fn)&&m.asin).map(m=>norm(m.asin)));
    return orders.map(o=>{
      const key=`${norm(o.order_id)}|${norm(o.fnsku)}`;
      const lines=shipBy.get(key)||[];
      const asins=asinsFor(o.fnsku);
      const myInv=invs.filter(i=>norm(i.ra_key)===norm(o.order_id)&&(!asins.size||!i.asin||asins.has(norm(i.asin))));
      const shipIds=[...new Set(myInv.map(i=>String(i.shipment_id||"")).filter(Boolean))];
      const myStr=strs.filter(r=>shipIds.includes(String(r.transaction_id))&&(!asins.size||asins.has(norm(r.asin))));
      const strUnits=myStr.reduce((a,r)=>a+unitsOf(r),0);
      const po=myStr.find(r=>r.po_number)||myStr[0]||{};
      const shipped=num(o.shipped_qty);
      const shipTotal=lines.reduce((a,s)=>a+num(s.shipped_qty),0);
      const returned=lines.filter(s=>/return/i.test(s.amazon_status||"")).reduce((a,s)=>a+num(s.shipped_qty),0);
      const expected=Math.max(0,shipped-returned);
      let phys=0,short=0,notrec=0;
      lines.forEach(s=>(recBy.get(s.id)||[]).forEach(r=>{phys+=num(r.good_qty)+num(r.damaged_qty)+num(r.near_expiry_qty)+num(r.expired_qty)+num(r.rework_qty);short+=num(r.short_qty);notrec+=num(r.not_received_qty);}));
      const pending=Math.max(0,expected-phys-short-notrec);
      const fl=feeBy.get(key)||[];
      const fee=fl.length?fl.reduce((a,f)=>a+num(f.fee_amount),0):null;
      const cur=[...new Set(fl.map(f=>f.currency).filter(Boolean))].join(", ");
      const tracking=[...new Set(lines.map(s=>s.tracking_number).filter(Boolean))].join(", ");
      let status;
      if(!shipped)status="Not shipped";
      else if(!myInv.length)status="No invoice";
      else if(!myStr.length)status="STR not found (FC_REMOVAL)";
      else if(strUnits!==shipped)status="Variance: STR vs shipped";
      else if(shipTotal!==shipped)status="Variance: shipment vs order";
      else status="Matched";
      const bmap=new Map();lines.forEach(s=>(recBy.get(s.id)||[]).forEach(r=>{const k=r.batch_no||"No batch";const q=num(r.good_qty)+num(r.damaged_qty)+num(r.near_expiry_qty)+num(r.expired_qty)+num(r.rework_qty);bmap.set(k,(bmap.get(k)||0)+q);}));const batchTxt=[...bmap.entries()].map(([k,q])=>`${k}: ${q}`).join("; ");return {o,invs:myInv,shipIds,strUnits,po,shipTotal,returned,expected,phys,short,notrec,pending,fee,cur,tracking,status,batchTxt};
    });
  }
  function tagFor(s){const t=String(s||"");if(/^(matched|complete)$/i.test(t))return "ok";if(/variance|partial|pending|not shipped/i.test(t))return "warn";if(/no invoice|not found|exception|awaiting/i.test(t))return "bad";return "neutral";}
  function renderDashCharts(){
    const box=$("#dashCharts");if(!box)return;
    const now=new Date(),months=[];for(let k=5;k>=0;k--){const d=new Date(now.getFullYear(),now.getMonth()-k,1);months.push(`${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,"0")}`);}
    const inv=new Map(months.map(m=>[m,new Set()])),ship=new Map(months.map(m=>[m,0]));
    (cache.str||[]).forEach(r=>{const m=String(r.invoice_date||"").slice(0,7);if(inv.has(m)&&r.invoice_no)inv.get(m).add(norm(r.invoice_no));});
    (cache.shipments||[]).forEach(s=>{const m=String(s.shipment_date||"").slice(0,7);if(ship.has(m))ship.set(m,ship.get(m)+num(s.shipped_qty));});
    const iv=months.map(m=>inv.get(m).size),sv=months.map(m=>ship.get(m));
    const mi=Math.max(1,...iv),ms=Math.max(1,...sv);
    const lbl=m=>{const [y,mo]=m.split("-");return new Date(+y,+mo-1,1).toLocaleString("en-GB",{month:"short"})+" "+y.slice(2);};
    const cols=months.map((m,i)=>`<div class="bar-col"><div class="bar-val">${fmt(iv[i])} · ${fmt(sv[i])}</div><div class="pair"><div class="bar-inv" style="height:${Math.round(iv[i]/mi*100)}%" title="Invoices: ${fmt(iv[i])}"></div><div class="bar-ship" style="height:${Math.round(sv[i]/ms*100)}%" title="Shipped units: ${fmt(sv[i])}"></div></div><div class="bar-label">${lbl(m)}</div></div>`).join("");
    const rows=reconRows(),total=rows.length,ok=rows.filter(r=>r.status==="Matched").length,exc=total-ok,deg=total?Math.round(ok/total*360):0;
    const donut=total?`<div class="donut-wrap"><div class="donut" style="background:conic-gradient(#16a34a 0 ${deg}deg,#f59e0b ${deg}deg 360deg)"><div class="donut-center">${Math.round(ok/total*100)}%</div></div><div class="donut-legend"><div><i style="background:#16a34a"></i>Matched: ${fmt(ok)}</div><div><i style="background:#f59e0b"></i>Exceptions: ${fmt(exc)}</div><div class="muted">Total lines: ${fmt(total)}</div></div></div>`:`<p class="muted">No reconciliation lines yet. Upload removal orders, shipments and invoices.</p>`;
    box.innerHTML=`<div class="panel chart-card"><h3>Monthly invoices and shipped units</h3><p class="muted">Last six months. Invoices counted by invoice date; shipped units from shipment date.</p><div class="legend"><span><i style="background:#2a7bc4"></i>Amazon invoices</span><span><i style="background:#14b8a6"></i>Shipped units</span></div><div class="bars">${cols}</div></div><div class="panel chart-card"><h3>Reconciliation status</h3><p class="muted">Share of order and product lines that are fully matched.</p>${donut}</div>`;
  }
  function freezeFirstColumns(table,n=3){
    const rows=[...table.rows];if(!rows.length)return;
    let left=0;const first=rows[0].cells;
    for(let c=0;c<Math.min(n,first.length);c++){
      const w=first[c].getBoundingClientRect().width;
      if(!w)return;
      rows.forEach(r=>{const cell=r.cells[c];if(cell){cell.classList.add("frozen");cell.style.left=left+"px";}});
      left+=w;
    }
  }
  function freezeAll(){
    document.querySelectorAll(".table-wrap table").forEach(t=>{t.querySelectorAll(".frozen").forEach(c=>{c.classList.remove("frozen");c.style.left="";});freezeFirstColumns(t,3);});
  }
  function receiptAllocated(r){return num(r.good_qty)+num(r.damaged_qty)+num(r.near_expiry_qty)+num(r.expired_qty)+num(r.rework_qty)+num(r.short_qty)+num(r.not_received_qty);}
  function canEditReceipt(r){return canEdit()||(profile?.role==="warehouse"&&r.created_by===session.user.id);}
  function renderReceiptEntries(){
    const tbl=$("#receiptEntriesTable");if(!tbl)return;
    const shipById=new Map((cache.shipments||[]).map(s=>[s.id,s]));
    const rows=[...(cache.receipts||[])].sort((x,y)=>String(y.created_at||"").localeCompare(String(x.created_at||"")));
    tbl.querySelector("tbody").innerHTML=rows.map(r=>{const s=shipById.get(r.shipment_line_id)||{},mm=masterFor(s.sku,s.fnsku,null);const btn=canEditReceipt(r)?`<button class="secondary" data-edit-receipt="${esc(r.id)}">Edit</button>`:`<span class="muted">—</span>`;return `<tr><td>${esc(r.receipt_date||"—")}</td><td>${esc(s.order_id||"—")}</td><td>${esc(s.tracking_number||"—")}</td><td>${esc(mm.material_code||"Not mapped")}</td><td>${esc(mm.title||"—")}</td><td>${esc(r.batch_no||"—")}</td><td>${esc(r.exp_date||"—")}</td><td>${fmt(r.good_qty)}</td><td>${fmt(r.damaged_qty)}</td><td>${fmt(r.near_expiry_qty)}</td><td>${fmt(r.expired_qty)}</td><td>${fmt(r.short_qty)}</td><td>${fmt(r.not_received_qty)}</td><td>${fmt(r.rework_qty)}</td><td>${esc(r.remarks||"—")}</td><td>${(()=>{const need=evidenceNeeded(r),cnt=evidenceFor(r.shipment_line_id).length;return need||cnt?`<button class="secondary" data-evidence-open="${esc(r.shipment_line_id)}">${cnt?cnt+" file(s)":"Add"}</button>${need&&!cnt?' <span class="tag bad">Missing</span>':""}`:"—";})()}</td><td>${r.created_by===session.user.id?"You":"Team"}</td><td>${btn}</td></tr>`;}).join("")||emptyRow(18);
    enhanceTable("receiptEntriesTable");freezeAll();
  }
  function openReceiptEdit(id){
    const r=(cache.receipts||[]).find(x=>x.id===id);if(!r)return;
    const f=$("#receiptEditForm");
    f.elements.id.value=r.id;
    ["good_qty","damaged_qty","near_expiry_qty","expired_qty","short_qty","not_received_qty"].forEach(k=>f.elements[k].value=num(r[k]));
    f.elements.receipt_date.value=r.receipt_date||today();f.elements.remarks.value=r.remarks||"";f.elements.batch_no.value=r.batch_no||"";f.elements.rework_qty.value=num(r.rework_qty);f.elements.mfg_date.value=r.mfg_date||"";f.elements.exp_date.value=r.exp_date||"";
    $("#receiptEditPanel").classList.remove("hidden");say("receiptEditMsg","");
    $("#receiptEditPanel").scrollIntoView({behavior:"smooth",block:"center"});
  }
  async function saveReceiptEdit(e){
    e.preventDefault();
    const f=e.target,id=f.elements.id.value,r=(cache.receipts||[]).find(x=>x.id===id);
    if(!r){say("receiptEditMsg","Receipt not found. Reload the page and try again.");return;}
    const keys=["good_qty","damaged_qty","near_expiry_qty","expired_qty","short_qty","not_received_qty","rework_qty"];
    const vals={};keys.forEach(k=>vals[k]=num(f.elements[k].value));
    if(Object.values(vals).some(v=>v<0)){say("receiptEditMsg","Quantities cannot be negative.");return;}
    if(vals.damaged_qty+vals.short_qty+vals.not_received_qty>0&&!evidenceFor(r.shipment_line_id).length){say("receiptEditMsg","Add photo or video evidence first (Evidence button in the list), then save.");return;}
    const ship=(cache.shipments||[]).find(s=>s.id===r.shipment_line_id);const shipped=num(ship?.shipped_qty);
    const others=(cache.receipts||[]).filter(x=>x.shipment_line_id===r.shipment_line_id&&x.id!==id).reduce((acc,x)=>acc+receiptAllocated(x),0);
    const total=Object.values(vals).reduce((acc,b)=>acc+b,0);
    if(others+total>shipped){say("receiptEditMsg",`Allocation exceeds shipped quantity. Shipped ${fmt(shipped)}, already allocated to other entries ${fmt(others)}.`);return;}
    const {error}=await sb.from("warehouse_receipts").update({...vals,receipt_date:f.elements.receipt_date.value,remarks:f.elements.remarks.value||null,batch_no:f.elements.batch_no.value.trim()||null,mfg_date:f.elements.mfg_date.value||null,exp_date:f.elements.exp_date.value||null,manual_edit:true}).eq("id",id);
    if(error){say("receiptEditMsg",friendlyDbError(error));return;}
    await audit("receipt_edited","warehouse_receipt",{id,...vals});
    say("receiptEditMsg","Saved.",true);
    await loadAll();
    $("#receiptEditPanel").classList.add("hidden");
  }
  function productCanEdit(){return canEdit();}
  function renderProducts(){
    const tbl=$("#productsTable");if(!tbl)return;
    const q=norm($("#productSearch")?.value||"");
    const list=(cache.mappings||[]).filter(m=>!q||norm([m.asin,m.sku,m.fnsku,m.material_code,m.sap_description,m.title,m.status].join(" ")).includes(q)).sort((x,y)=>String(x.title||"").localeCompare(String(y.title||"")));
    const act=m=>productCanEdit()?`<button class="secondary" data-edit-product="${esc(m.id)}">Edit</button> <button class="secondary" data-delete-product="${esc(m.id)}">Remove</button>`:`<span class="muted">View only</span>`;
    tbl.querySelector("tbody").innerHTML=list.map(m=>`<tr><td>${esc(m.asin||"—")}</td><td>${esc(m.sku||"—")}</td><td>${esc(m.fnsku||"—")}</td><td>${esc(m.material_code||"—")}</td><td>${esc(m.sap_description||"—")}</td><td>${esc(m.title||"—")}</td><td>${fmt(num(m.raw_data?.["Pack Of"])||1)}</td><td>${fmt(num(m.new_mrp)||num(m.raw_data?.["MRP"])||0)}</td><td>${m.cogs==null?"—":fmt(m.cogs)}</td><td><span class="tag ${m.status==="Active"?"ok":m.status==="Inactive"?"neutral":"warn"}">${esc(m.status||"Active")}</span></td><td>${act(m)}</td></tr>`).join("")||emptyRow(11);
    const addBtn=$("#productAddBtn");if(addBtn)addBtn.classList.toggle("hidden",!productCanEdit());
    enhanceTable("productsTable");freezeAll();
  }
  function openProductForm(id){
    const f=$("#productForm"),panel=$("#productFormPanel");if(!f||!panel)return;
    const m=id?(cache.mappings||[]).find(x=>x.id===id):null;
    f.reset();
    f.elements.id.value=m?.id||"";
    f.elements.asin.value=m?.asin||"";f.elements.sku.value=m?.sku||"";f.elements.fnsku.value=m?.fnsku||"";
    f.elements.material_code.value=m?.material_code||"";f.elements.title.value=m?.title||"";f.elements.sap_description.value=m?.sap_description||"";f.elements.cogs.value=m?.cogs==null?"":m.cogs;
    f.elements.pack_of.value=m?(num(m.raw_data?.["Pack Of"])||1):1;
    f.elements.status.value=m?.status||"Active";
    $("#productFormTitle").textContent=m?"Edit product / SKU":"Add product / SKU";
    panel.classList.remove("hidden");say("productMsg","");
    panel.scrollIntoView({behavior:"smooth",block:"center"});
  }
  async function saveProduct(e){
    e.preventDefault();
    if(!productCanEdit()){say("productMsg","Only a manager or admin can change products.");return;}
    const f=e.target,id=f.elements.id.value;
    const v={asin:f.elements.asin.value.trim()||null,sku:f.elements.sku.value.trim()||null,fnsku:f.elements.fnsku.value.trim()||null,material_code:f.elements.material_code.value.trim()||null,title:f.elements.title.value.trim()||null,sap_description:f.elements.sap_description.value.trim()||null,cogs:f.elements.cogs.value===""?null:num(f.elements.cogs.value),status:f.elements.status.value||"Active"};
    const pack=num(f.elements.pack_of.value)||1;
    if(!v.asin&&!v.sku&&!v.fnsku){say("productMsg","Enter at least an ASIN, SKU or FNSKU.");return;}
    const list=cache.mappings||[];
    const clash=list.find(m=>m.id!==id&&(v.fnsku?norm(m.fnsku)===norm(v.fnsku):(!norm(m.fnsku)&&v.asin&&norm(m.asin)===norm(v.asin))));
    if(clash){say("productMsg",`Duplicate: ${v.fnsku?"this FNSKU":"this ASIN (without FNSKU)"} already exists as "${clash.title||clash.sku||clash.asin}". Edit that row instead.`);return;}
    const prev=id?(list.find(m=>m.id===id)?.raw_data||{}):{};
    const raw={...prev,"Pack Of":pack,"Product Title":v.title,"SKU":v.sku,"FNSKU":v.fnsku,"ASIN":v.asin,"Material Code":v.material_code,"SAP Description":v.sap_description,"COGS":v.cogs,"Status":v.status};
    const row={...v,raw_data:raw};
    if(id){
      const {error}=await sb.from("product_mappings").update(row).eq("id",id);
      if(error){say("productMsg",friendlyDbError(error));return;}
      await audit("product_edited","product_mapping",{id,...v});
    }else{
      const {error}=await sb.from("product_mappings").insert(row);
      if(error){say("productMsg",friendlyDbError(error));return;}
      await audit("product_added","product_mapping",v);
    }
    say("productMsg","Saved.",true);
    await loadAll();
    $("#productFormPanel").classList.add("hidden");
  }
  async function deleteProduct(id){
    if(!productCanEdit())return;
    const m=(cache.mappings||[]).find(x=>x.id===id);if(!m)return;
    if(!confirm(`Remove "${m.title||m.sku||m.asin}" from the product master?\n\nTip: set Status to Inactive instead if old removals still use this product.`))return;
    const {error}=await sb.from("product_mappings").delete().eq("id",id);
    if(error){say("productMsg",friendlyDbError(error));return;}
    await audit("product_removed","product_mapping",{id,asin:m.asin,sku:m.sku,fnsku:m.fnsku});
    await loadAll();
  }
  function physOf(id){const t=receiptTotals(id);return t.good_qty+t.damaged_qty+t.near_expiry_qty+t.expired_qty+t.rework_qty;}
  function snrOf(id){const t=receiptTotals(id);return t.short_qty+t.not_received_qty;}
  const ALL_PAGES=["dashboard","orders","shipments","receipts","fees","sap","reconciliation","claims","approvals","scat","reports","employees","audit"];
  const PAGE_LABELS={dashboard:"Dashboard",orders:"Removal Orders",shipments:"Shipment Tracker",receipts:"Warehouse Receipts",fees:"Removal Fees",sap:"PO & SAP Tracker",reconciliation:"Reconciliation",claims:"Amazon Claims",approvals:"Approvals",scat:"SAP SCAT",reports:"Reports & Imports",employees:"Employees",audit:"Audit Logs"};
  const ROLE_PAGES={warehouse:["receipts"]};
  let allProfiles=[];
  function pagesFor(p){if(!p||p.role==="admin")return ALL_PAGES;if(Array.isArray(p.allowed_pages)&&p.allowed_pages.length)return p.allowed_pages;return ROLE_PAGES[p.role]||ALL_PAGES;}
  function applyPageAccess(p){const allowed=pagesFor(p);document.querySelectorAll(".nav-item").forEach(b=>b.classList.toggle("hidden",!allowed.includes(b.dataset.page)));const cur=document.querySelector(".nav-item.active")?.dataset.page;if(!allowed.includes(cur))nav(allowed[0]);}
  function renderAccessPanel(){
    const panel=$("#accessPanel");if(!panel)return;
    panel.classList.toggle("hidden",profile?.role!=="admin");
    const sel=$("#accessUser");if(!sel)return;
    const prev=sel.value;
    sel.innerHTML=allProfiles.map(u=>`<option value="${esc(u.id)}">${esc(u.full_name||u.email)} (${esc(u.role)})</option>`).join("");
    if(allProfiles.some(u=>u.id===prev))sel.value=prev;
    refreshAccessChecks();
  }
  function refreshAccessChecks(){
    const u=allProfiles.find(x=>x.id===$("#accessUser").value);const box=$("#accessPages");if(!box)return;
    if(!u){box.innerHTML="";return;}
    const on=new Set(pagesFor(u));
    box.innerHTML=ALL_PAGES.map(k=>`<label style="display:flex;gap:8px;align-items:center"><input type="checkbox" value="${k}" ${on.has(k)?"checked":""} ${u.role==="admin"?"disabled":""}> ${PAGE_LABELS[k]}</label>`).join("");
    say("accessMsg",u.role==="admin"?"Admin always has access to all pages.":(Array.isArray(u.allowed_pages)&&u.allowed_pages.length?"Custom access set for this user.":"Using role default."),true);
  }
  async function saveAccess(e){
    e.preventDefault();
    const u=allProfiles.find(x=>x.id===$("#accessUser").value);if(!u||profile?.role!=="admin")return;
    const list=[...document.querySelectorAll("#accessPages input:checked")].map(i=>i.value);
    if(!list.length){say("accessMsg","Select at least one page, or use Reset to role default.");return;}
    const {error}=await sb.from("profiles").update({allowed_pages:list}).eq("id",u.id);
    if(error){say("accessMsg",friendlyDbError(error));return;}
    await audit("access_changed","profile",{user:u.id,pages:list});
    say("accessMsg","Saved.",true);
    await renderEmployees();
  }
  async function resetAccess(){
    const u=allProfiles.find(x=>x.id===$("#accessUser").value);if(!u||profile?.role!=="admin")return;
    const {error}=await sb.from("profiles").update({allowed_pages:null}).eq("id",u.id);
    if(error){say("accessMsg",friendlyDbError(error));return;}
    await audit("access_reset","profile",{user:u.id});
    say("accessMsg","Reset to role default.",true);
    await renderEmployees();
  }
  async function ensureXLSX(){await loadScriptOnce("https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js");return window.XLSX;}
  const hkey=k=>String(k??"").trim().toLowerCase();
  function sheetRecords(ws,X){
    const raw=X.utils.sheet_to_json(ws,{header:1,defval:null,raw:true});
    const hi=raw.findIndex(r=>r.some(c=>hkey(c)==="tracking no"));if(hi<0)return [];
    const head=raw[hi].map(hkey),out=[];
    raw.slice(hi+1).forEach(r=>{if(r.some(c=>hkey(c)==="tracking no"))return;const o={};head.forEach((h,i)=>{if(h)o[h]=r[i]??null;});if(o["tracking no"]!==null&&String(o["tracking no"]).trim()!=="")out.push(o);});
    return out;
  }
  const pickCell=(o,...keys)=>{for(const k of keys){const v=o[hkey(k)];if(v!==null&&v!==undefined&&String(v).trim()!=="")return v;}return null;};
  function dateISOFromCell(v){
    if(v instanceof Date&&!isNaN(v))return `${v.getFullYear()}-${String(v.getMonth()+1).padStart(2,"0")}-${String(v.getDate()).padStart(2,"0")}`;
    const s=String(v??"").trim();const m=/^(\d{4})-(\d{2})-(\d{2})/.exec(s);if(m)return m[0];
    const d=/^(\d{1,2})[\/.-](\d{1,2})[\/.-](\d{4})$/.exec(s);if(d)return `${d[3]}-${d[2].padStart(2,"0")}-${d[1].padStart(2,"0")}`;
    return null;
  }
  function skuMatch(s,sk){const mat=norm(masterFor(s.sku,s.fnsku,null).material_code),ss=norm(s.sku);return mat===sk||ss===sk||(!!ss&&sk.includes(ss))||(!!sk&&ss.includes(sk));}
  function classifyGood(goodQty,rd,exp){
    if(!exp)return{good:goodQty,near:0,expired:0};
    const days=Math.floor((new Date(exp)-new Date(rd))/86400000);
    if(days<0)return{good:0,near:0,expired:goodQty};
    if(days<=90)return{good:0,near:goodQty,expired:0};
    return{good:goodQty,near:0,expired:0};
  }
  async function importWarehouseExcel(file,fallbackDate){
    say("warehouseImportMsg","Reading file...",true);
    const X=await ensureXLSX();
    const wb=X.read(await file.arrayBuffer(),{type:"array",cellDates:true});
    const phyWs=wb.Sheets["PHY DATA"];
    if(!phyWs){say("warehouseImportMsg","Sheet 'PHY DATA' not found in this file.");return;}
    const phy=sheetRecords(phyWs,X);
    const trk=o=>String(pickCell(o,"tracking no")??"").replace(/\.0$/,"").trim();
    const T={read:0,inserted:0,updated:0,unchanged:0,keptManual:0,rejected:0};
    const problems=[],rows=[];
    phy.forEach((o,i)=>{
      T.read++;
      const tracking=trk(o),sku=String(pickCell(o,"sku code")??"").replace(/\.0$/,"").trim();
      const batch=String(pickCell(o,"batch")??"").trim(),box=String(pickCell(o,"box no")??"").trim();
      const good=num(pickCell(o,"good")),rework=num(pickCell(o,"rework")),damage=num(pickCell(o,"damage")),ph=num(pickCell(o,"ph qty"));
      if(!tracking||!sku||!ph){T.rejected++;problems.push(`Row ${i+2}: tracking, SKU or Ph Qty missing`);return;}
      if(good+rework+damage!==ph){T.rejected++;problems.push(`${tracking} / ${sku}: Ph Qty ${ph} is not Good + Rework + Damage`);return;}
      rows.push({i,tracking,sku,batch,box,good,rework,damage,ph,mfg:dateISOFromCell(pickCell(o,"mfg")),exp:dateISOFromCell(pickCell(o,"exp")),rd:dateISOFromCell(pickCell(o,"receipt date","received date"))||fallbackDate});
    });
    const groups=new Map();
    rows.forEach(r=>{const k=norm(r.tracking)+"|"+norm(r.sku);if(!groups.has(k))groups.set(k,[]);groups.get(k).push(r);});
    const ops=[];
    for(const [gk,list] of groups){
      const [tr,sk]=gk.split("|");
      const lines=(cache.shipments||[]).filter(s=>norm(s.tracking_number)===tr&&skuMatch(s,sk)).sort((x,y)=>num(x.line_seq)-num(y.line_seq)||String(x.id).localeCompare(String(y.id)));
      if(!lines.length){list.forEach(r=>{T.rejected++;problems.push(`${r.tracking} / ${r.sku}: no shipment line matches this tracking and SKU`);});continue;}
      const lineIds=new Set(lines.map(l=>l.id));
      const recs=(cache.receipts||[]).filter(x=>lineIds.has(x.shipment_line_id));
      const replaceable=recs.filter(x=>x.source_file===file.name&&!x.manual_edit);
      const fixed=recs.filter(x=>!(x.source_file===file.name&&!x.manual_edit));
      T.keptManual+=fixed.filter(x=>x.source_file===file.name&&x.manual_edit).length;
      const cap=new Map(lines.map(l=>[l.id,num(l.shipped_qty)]));
      fixed.forEach(x=>cap.set(x.shipment_line_id,cap.get(x.shipment_line_id)-receiptAllocated(x)));
      const need=list.reduce((acc,r)=>acc+r.ph,0);
      const free=[...cap.values()].reduce((acc,b)=>acc+Math.max(0,b),0);
      if(need>free){list.forEach(r=>{T.rejected++;problems.push(`${r.tracking} / ${r.sku}: file needs ${need}, only ${free} free on this shipment (manual or other entries use the rest)`);});continue;}
      const capLeft=new Map(lines.map(l=>[l.id,Math.max(0,cap.get(l.id))]));
      let li=0;const segs=[];
      for(const r of list){
        const parts=new Map();
        for(const [kind,qty] of [["good",r.good],["rework",r.rework],["damage",r.damage]]){
          let left=qty;
          while(left>0){
            while(li<lines.length&&capLeft.get(lines[li].id)<=0)li++;
            const lid=lines[li].id,take=Math.min(left,capLeft.get(lid));
            capLeft.set(lid,capLeft.get(lid)-take);left-=take;
            const p=parts.get(lid)||{good:0,rework:0,damage:0};p[kind]+=take;parts.set(lid,p);
          }
        }
        let seq=0;
        for(const [lid,p] of parts){
          seq++;
          const c=classifyGood(p.good,r.rd,r.exp);
          segs.push({key:[norm(r.tracking),norm(r.sku),norm(r.batch),norm(r.box),r.i,seq].join("|"),lineId:lid,
            payload:{good_qty:c.good,damaged_qty:p.damage,near_expiry_qty:c.near,expired_qty:c.expired,rework_qty:p.rework,short_qty:0,not_received_qty:0,receipt_date:r.rd,remarks:`Imported: ${file.name}`,batch_no:r.batch||null,mfg_date:r.mfg,exp_date:r.exp,source_file:file.name,import_key:null,shipment_line_id:lid}});
        }
      }
      const used=new Set();
      for(const s of segs){
        s.payload.import_key=s.key;
        const row=replaceable.find(x=>!used.has(x.id)&&x.import_key===s.key)||replaceable.find(x=>!used.has(x.id)&&x.shipment_line_id===s.lineId&&norm(x.batch_no)===norm(s.payload.batch_no));
        if(row){
          used.add(row.id);
          const p=s.payload;
          const same=["good_qty","damaged_qty","near_expiry_qty","expired_qty","rework_qty"].every(k=>num(row[k])===num(p[k]))&&String(row.receipt_date||"")===String(p.receipt_date||"")&&String(row.exp_date||"")===String(p.exp_date||"")&&String(row.mfg_date||"")===String(p.mfg_date||"")&&String(row.batch_no||"")===String(p.batch_no||"");
          if(same){T.unchanged++;continue;}
          const {shipment_line_id,...upd}=p;
          ops.push({type:"update",id:row.id,payload:upd,delta:receiptAllocated(p)-receiptAllocated(row)});
        }else{
          ops.push({type:"insert",payload:s.payload,delta:receiptAllocated(s.payload)});
        }
      }
      replaceable.forEach(x=>{if(!used.has(x.id)&&receiptAllocated(x)>0)ops.push({type:"update",id:x.id,payload:{good_qty:0,damaged_qty:0,near_expiry_qty:0,expired_qty:0,rework_qty:0,import_key:null},delta:-receiptAllocated(x)});});
    }
    ops.sort((x,y)=>(x.type==="insert")-(y.type==="insert")||x.delta-y.delta);
    for(const op of ops){
      if(op.type==="update"){
        const {error}=await sb.from("warehouse_receipts").update(op.payload).eq("id",op.id);
        if(error){T.rejected++;problems.push(friendlyDbError(error));}else T.updated++;
      }else{
        const {error}=await sb.from("warehouse_receipts").insert(op.payload);
        if(error){T.rejected++;problems.push(friendlyDbError(error));}else T.inserted++;
      }
    }
    await loadAll();
    say("warehouseImportMsg",`Done. Read ${T.read}, inserted ${T.inserted}, updated ${T.updated}, unchanged ${T.unchanged}, kept manual edits ${T.keptManual}, rejected ${T.rejected}.`,T.rejected===0);
    const box=$("#warehouseImportTotals");if(box)box.innerHTML=problems.length?`<p class="muted">${problems.map(esc).join("<br>")}</p>`:"";
  }
  function bindBatchModal(){
    $("#batchAddRow").onclick=()=>$("#batchTable tbody").insertAdjacentHTML("beforeend",batchRowHtml({cond:"good",batch:"",qty:0,mfg:"",exp:""}));
    $("#batchTable").onclick=e=>{const d=e.target.closest&&e.target.closest("[data-bdel]");if(d)d.closest("tr").remove();};
    $("#batchSave").onclick=saveBatchModal;
    $("#batchCancel").onclick=()=>{$("#batchModal").classList.add("hidden");batchCtx=null;};
    $("#receiptLineEditor").addEventListener("change",e=>{const inp=e.target.closest&&e.target.closest("[data-qty]");if(!inp||!["good_qty","damaged_qty","rework_qty"].includes(inp.dataset.qty))return;const line=inp.closest(".receipt-line");if(line&&num(inp.value)>0&&!JSON.parse(line.dataset.batches||"[]").length)openBatchModal(line);});
    document.addEventListener("click",e=>{const b=e.target.closest&&e.target.closest("[data-open-batches]");if(b)openBatchModal(b.closest(".receipt-line"));});
  }
  // ---------- Approvals: damaged / expired / near expiry, prepared in Outlook ----------
  const APPR_CASES=["Damaged","Expired","Near Expiry"];
  // APPR-PURE-START
  const apprMailList=s=>String(s||"").split(/[;,\s]+/).filter(Boolean).join(",");
  const apprEnc=s=>encodeURIComponent(s).replace(/%40/g,"@").replace(/%2C/gi,",");
  function apprOutlookUrl(to,cc,subject,mode){
    const body=encodeURIComponent("Paste the copied table here (Ctrl+V).");
    if(mode==="web")return "https://outlook.office.com/mail/deeplink/compose?to="+apprEnc(to)+(cc?"&cc="+apprEnc(cc):"")+"&subject="+encodeURIComponent(subject)+"&body="+body;
    return "mailto:"+apprEnc(to)+"?subject="+encodeURIComponent(subject)+(cc?"&cc="+apprEnc(cc):"")+"&body="+body;
  }
  function apprTotals(rows){return {qty:rows.reduce((x,r)=>x+r.qty,0),mrp:rows.reduce((x,r)=>x+r.mrpValue,0),cogs:rows.reduce((x,r)=>x+(r.cogsValue||0),0),noCogs:rows.filter(r=>r.cogsValue==null).length};}
  function apprTableHtml(rows,fmt,esc){
    const th='style="border:1px solid #cbd5e1;padding:6px;background:#f1f5f9;text-align:left;font-size:12px"',td='style="border:1px solid #cbd5e1;padding:6px;font-size:12px"',tr='style="border:1px solid #cbd5e1;padding:6px;font-size:12px;text-align:right"';
    const head=["#","Case","Order ID","SAP Material","SAP Description","Batch","Exp date","Qty","MRP","MRP value","COGS","COGS value"].map(x=>`<th ${th}>${x}</th>`).join("");
    const body=rows.map((r,i)=>`<tr><td ${td}>${i+1}</td><td ${td}>${esc(r.case_type)}</td><td ${td}>${esc(r.order_id)}</td><td ${td}>${esc(r.material||"-")}</td><td ${td}>${esc(r.desc||"-")}</td><td ${td}>${esc(r.batch_no||"-")}</td><td ${td}>${esc(r.exp_date||"-")}</td><td ${tr}>${fmt(r.qty)}</td><td ${tr}>${fmt(r.mrp)}</td><td ${tr}>${fmt(r.mrpValue)}</td><td ${tr}>${r.cogs>0?fmt(r.cogs):"-"}</td><td ${tr}>${r.cogsValue==null?"-":fmt(r.cogsValue)}</td></tr>`).join("");
    const t=apprTotals(rows);
    const foot=`<tr><td colspan="7" ${td}><b>Total</b></td><td ${tr}><b>${fmt(t.qty)}</b></td><td ${td}></td><td ${tr}><b>${fmt(t.mrp)}</b></td><td ${td}></td><td ${tr}><b>${fmt(t.cogs)}</b></td></tr>`;
    return `<table style="border-collapse:collapse;width:100%"><thead><tr>${head}</tr></thead><tbody>${body}${foot}</tbody></table><p style="font-size:12px;color:#64748b">MRP value = quantity x MRP. COGS value = quantity x COGS.${t.noCogs?` COGS is not available for ${t.noCogs} line(s), so they are not in the COGS total.`:""}</p>`;
  }
  function apprPlain(rows){
    const t=apprTotals(rows);
    return ["Case\tOrder ID\tSAP Material\tSAP Description\tBatch\tExp date\tQty\tMRP\tMRP value\tCOGS\tCOGS value"].concat(rows.map(r=>[r.case_type,r.order_id,r.material,r.desc,r.batch_no||"",r.exp_date||"",r.qty,r.mrp,r.mrpValue,r.cogs>0?r.cogs:"",r.cogsValue==null?"":r.cogsValue].join("\t"))).concat(["Total\t\t\t\t\t\t"+t.qty+"\t\t"+t.mrp+"\t\t"+t.cogs]).join("\n");
  }
  // APPR-PURE-END
  let apprList=[],apprLast=null;
  function apprSettings(){return (cache.mailSettings||[])[0]||{};}
  function apprCandidates(){
    const from=$("#apprFrom")?.value||"",to=$("#apprTo")?.value||"",cs=$("#apprCase")?.value||"all";
    const shipById=new Map((cache.shipments||[]).map(s=>[s.id,s]));
    const lockedBatch=new Set((cache.approvalBatches||[]).filter(b=>!["Rejected","Cancelled"].includes(b.status)).map(b=>b.id));
    const locked=new Set((cache.approvalItems||[]).filter(i=>lockedBatch.has(i.batch_id)).map(i=>i.receipt_id+"|"+i.case_type));
    const out=[];
    (cache.receipts||[]).forEach(r=>{
      const d=String(r.receipt_date||"").slice(0,10);
      if(from&&d<from)return;if(to&&d>to)return;
      const s=shipById.get(r.shipment_line_id)||{};
      const m=masterFor(s.sku,s.fnsku,null);
      [["Damaged",num(r.damaged_qty)],["Expired",num(r.expired_qty)],["Near Expiry",num(r.near_expiry_qty)]].forEach(([c,q])=>{
        if(q<=0||(cs!=="all"&&cs!==c)||locked.has(r.id+"|"+c))return;
        out.push({receipt_id:r.id,case_type:c,order_id:s.order_id||"",tracking_number:s.tracking_number||"",sku:s.sku||"",material:m.material_code||"",desc:m.title||"",batch_no:r.batch_no||"",exp_date:r.exp_date||"",received:d,qty:q,mrp:m.mrp||0,mrpValue:q*(m.mrp||0),cogs:m.cogs||0,cogsValue:m.cogs>0?q*m.cogs:null});
      });
    });
    return out;
  }
  function apprRowsOf(batchId){
    return (cache.approvalItems||[]).filter(i=>i.batch_id===batchId).map(i=>({case_type:i.case_type,order_id:i.order_id||"",material:i.material_code||"",desc:i.material_desc||"",batch_no:i.batch_no||"",exp_date:i.exp_date||"",qty:num(i.qty),mrp:num(i.mrp),mrpValue:num(i.value),cogs:num(i.cogs),cogsValue:i.cogs_value==null?null:num(i.cogs_value)}));
  }
  const apprAge=b=>Math.floor((Date.now()-new Date(b.followup_sent_at||b.sent_at))/86400000);
  function renderApprovals(){
    const page=$("#approvalsPage");if(!page)return;
    const manager=canEdit();
    $("#apprLocked")?.classList.toggle("hidden",manager);
    $("#apprBody")?.classList.toggle("hidden",!manager);
    if(!manager)return;
    const st=apprSettings(),f=$("#apprSettingsForm");
    if(f&&!f.dataset.filled){
      if(st.id){f.elements.addressee.value=st.sender_name||"";f.elements.to_list.value=st.to_list||"";f.elements.cc_list.value=st.cc_list||"";}
      try{f.elements.open_with.value=localStorage.getItem("apprOpenWith")||"desktop";}catch(_){}
      f.dataset.filled="1";
    }
    const batches=[...(cache.approvalBatches||[])].sort((x,y)=>String(y.sent_at).localeCompare(String(x.sent_at)));
    const last=batches.find(b=>b.status!=="Cancelled");
    const days=last?Math.floor((Date.now()-new Date(last.sent_at))/86400000):null;
    const waiting=batches.filter(b=>b.status==="Sent").length,prepared=batches.filter(b=>b.status==="Prepared").length,due=batches.filter(b=>b.status==="Sent"&&apprAge(b)>=2).length;
    $("#apprDue").textContent=days===null?"No approval mail yet.":days>=14?`Fortnightly approval is due. The last mail was ${days} days ago.`:`Next fortnightly mail due in ${14-days} day(s).`;
    $("#apprPending").textContent=[waiting?`${waiting} mail(s) waiting for approval`:"",due?`${due} follow-up(s) due`:"",prepared?`${prepared} prepared, not marked as sent`:""].filter(Boolean).join(" · ")||"Nothing waiting.";
    if(!$("#apprFrom").value&&last)$("#apprFrom").value=String(last.sent_at).slice(0,10);
    apprList=apprCandidates();
    const per=c=>apprList.filter(r=>r.case_type===c),T=apprTotals(apprList);
    $("#apprSummary").innerHTML=`<div class="kpi-grid compact">${APPR_CASES.map(c=>{const t=apprTotals(per(c));return `<article class="kpi"><span>${c}</span><b>${fmt(t.qty)} units</b><small>MRP ${fmt(t.mrp)} · COGS ${fmt(t.cogs)}</small></article>`;}).join("")}<article class="kpi"><span>Total MRP value</span><b>${fmt(T.mrp)}</b><small>COGS ${fmt(T.cogs)} · ${fmt(apprList.length)} line(s) not yet mailed</small></article></div>`;
    $("#apprItems tbody").innerHTML=apprList.map((r,i)=>`<tr data-idx="${i}"><td><input type="checkbox" checked></td><td>${esc(r.case_type)}</td><td>${esc(r.order_id||"—")}</td><td>${esc(r.tracking_number||"—")}</td><td>${esc(r.material||"Not mapped")}</td><td>${esc(r.desc||"—")}</td><td>${esc(r.batch_no||"—")}</td><td>${esc(r.exp_date||"—")}</td><td>${esc(r.received||"—")}</td><td>${fmt(r.qty)}</td><td>${fmt(r.mrp)}</td><td>${fmt(r.mrpValue)}</td><td>${r.cogs>0?fmt(r.cogs):"—"}</td><td>${r.cogsValue==null?"—":fmt(r.cogsValue)}</td></tr>`).join("")||emptyRow(14);
    enhanceTable("apprItems");freezeAll();
    $("#apprHistory tbody").innerHTML=batches.map(b=>{
      const tag=b.status==="Approved"?"ok":(b.status==="Rejected"||b.status==="Cancelled")?"bad":"warn";
      const fu=b.followup_count?`<br><small>Follow-up sent ${fmt(b.followup_count)} time(s), last ${esc(String(b.followup_sent_at||"").slice(0,10))}</small>`:"";
      const dueTag=(b.status==="Sent"&&apprAge(b)>=2)?' <span class="tag bad">Follow-up due</span>':"";
      const actions=b.status==="Prepared"?`<button class="secondary" data-appr-sent="${esc(b.id)}">Mark as sent</button> <button class="secondary" data-appr-reopen="${esc(b.id)}">Open mail again</button> <button class="secondary" data-appr-cancel="${esc(b.id)}">Cancel</button>`:b.status==="Sent"?`<button class="secondary" data-decide="${esc(b.id)}">Record decision</button> <button class="secondary" data-appr-fu="${esc(b.id)}">Prepare follow-up</button> <button class="secondary" data-appr-fusent="${esc(b.id)}">Follow-up sent</button>`:"—";
      return `<tr><td>${esc(String(b.sent_at||"").slice(0,10))}</td><td>${esc(b.subject||"")}</td><td>${esc(b.to_list||"")}</td><td>${fmt(b.item_count)}</td><td>${fmt(b.total_value)}</td><td>${b.total_cogs_value==null?"—":fmt(b.total_cogs_value)}</td><td><span class="tag ${tag}">${esc(b.status)}</span>${dueTag}${fu}</td><td>${esc(b.approved_by_name||"—")} ${esc(b.approval_date||"")}</td><td>${b.screenshot_path?`<button class="secondary" data-shot="${esc(b.id)}">View screenshot</button>`:"—"}</td><td>${actions}</td></tr>`;
    }).join("")||emptyRow(10);
    enhanceTable("apprHistory");
  }
  async function apprCopyHtml(html,plain){
    try{if(window.ClipboardItem&&navigator.clipboard?.write){await navigator.clipboard.write([new ClipboardItem({"text/html":new Blob([html],{type:"text/html"}),"text/plain":new Blob([plain],{type:"text/plain"})})]);return true;}}catch(_){}
    try{const d=document.createElement("div");d.contentEditable="true";d.style.cssText="position:fixed;left:-9999px;top:0";d.innerHTML=html;document.body.appendChild(d);const rg=document.createRange();rg.selectNodeContents(d);const sel=window.getSelection();sel.removeAllRanges();sel.addRange(rg);const ok=document.execCommand("copy");sel.removeAllRanges();d.remove();return ok;}catch(_){return false;}
  }
  function apprOpenOutlook(m){
    const mode=$("#apprSettingsForm").elements.open_with.value;
    try{localStorage.setItem("apprOpenWith",mode);}catch(_){}
    const url=apprOutlookUrl(m.to,m.cc,m.subject,mode);
    if(mode==="web")window.open(url,"_blank","noopener");else window.location.href=url;
  }
  async function apprDeliver(m){
    apprLast=m;
    const ok=await apprCopyHtml(m.html,m.plain);
    $("#apprPreview").innerHTML=m.html;$("#apprPrepared").classList.remove("hidden");
    apprOpenOutlook(m);
    say("apprMsg",ok?"The table is copied. In Outlook click in the message, press Ctrl+V, check To and CC, and send. Then click Mark as sent in the list below.":"The table could not be copied automatically. Click Copy table below, then paste it in Outlook.",ok);
  }
  const apprIntro=(who,html)=>`<div style="font-family:Arial,sans-serif;color:#0f172a"><p>Dear ${esc(who||"team")},</p>${html}</div>`;
  async function apprPrepare(e){
    e.preventDefault();
    const f=e.target,to=apprMailList(f.elements.to_list.value),cc=apprMailList(f.elements.cc_list.value),who=f.elements.addressee.value.trim();
    if(!to){say("apprMsg","Enter at least one To address.");return;}
    const rows=[...document.querySelectorAll("#apprItems tbody tr[data-idx]")].filter(tr=>tr.querySelector("input[type=checkbox]")?.checked).map(tr=>apprList[Number(tr.dataset.idx)]).filter(Boolean);
    if(!rows.length){say("apprMsg","Select at least one line.");return;}
    const T=apprTotals(rows);
    const subject=`Approval required: damaged / expired / near expiry stock, ${rows.length} line(s), MRP value ${fmt(T.mrp)}, COGS value ${fmt(T.cogs)} (${today()})`;
    const html=apprIntro(who,`<p>Please review the stock below (damaged, expired and near expiry) and reply to this mail with <b>APPROVED</b> or <b>REJECTED</b>, your name and the date.</p>${apprTableHtml(rows,fmt,esc)}`);
    const copied=await apprCopyHtml(html,apprPlain(rows));
    const {data:bd,error:be}=await sb.from("approval_batches").insert({subject,to_list:to,cc_list:cc||null,sender_name:who||null,item_count:rows.length,total_qty:T.qty,total_value:T.mrp,total_cogs_value:T.cogs,status:"Prepared"}).select("id");
    if(be||!bd?.length){say("apprMsg","The mail record was not saved: "+friendlyDbError(be||"no id returned"));return;}
    const batchId=bd[0].id;
    const {error:ie}=await sb.from("approval_items").insert(rows.map(r=>({batch_id:batchId,receipt_id:r.receipt_id,case_type:r.case_type,order_id:r.order_id||null,tracking_number:r.tracking_number||null,sku:r.sku||null,material_code:r.material||null,material_desc:r.desc||null,batch_no:r.batch_no||null,exp_date:r.exp_date||null,qty:r.qty,mrp:r.mrp,value:r.mrpValue,cogs:r.cogs>0?r.cogs:null,cogs_value:r.cogsValue})));
    if(ie){say("apprMsg","The mail lines were not saved: "+friendlyDbError(ie));return;}
    await sb.from("mail_settings").upsert({id:"default",sender_name:who||null,to_list:to,cc_list:cc||null,updated_at:new Date().toISOString()});
    await audit("approval_mail_prepared","approval_batch",{id:batchId,items:rows.length,mrp_value:T.mrp,cogs_value:T.cogs});
    await apprDeliver({to,cc,subject,html,plain:apprPlain(rows),copiedEarly:copied});
    await loadAll();
  }
  async function apprSetStatus(id,patch,auditAction,msg){
    const {error}=await sb.from("approval_batches").update(patch).eq("id",id);
    if(error){say("apprMsg",friendlyDbError(error));return;}
    await audit(auditAction,"approval_batch",{id,...patch});
    say("apprMsg",msg,true);await loadAll();
  }
  const apprMarkSent=id=>apprSetStatus(id,{status:"Sent",sent_at:new Date().toISOString()},"approval_mail_sent","Marked as sent.");
  function apprCancel(id){if(confirm("Cancel this prepared mail? Its lines can then be mailed again."))apprSetStatus(id,{status:"Cancelled"},"approval_mail_cancelled","Cancelled. The lines are free again.");}
  function apprReopen(id){
    const b=(cache.approvalBatches||[]).find(x=>x.id===id);if(!b)return;
    const rows=apprRowsOf(id);
    apprDeliver({to:b.to_list,cc:b.cc_list||"",subject:b.subject,html:apprIntro(b.sender_name,`<p>Please review the stock below (damaged, expired and near expiry) and reply to this mail with <b>APPROVED</b> or <b>REJECTED</b>, your name and the date.</p>${apprTableHtml(rows,fmt,esc)}`),plain:apprPlain(rows)});
  }
  function apprFollowUp(id){
    const b=(cache.approvalBatches||[]).find(x=>x.id===id);if(!b)return;
    const rows=apprRowsOf(id),n=(b.followup_count||0)+1;
    apprDeliver({to:b.to_list,cc:b.cc_list||"",subject:`Follow-up ${n}: ${b.subject}`,html:apprIntro(b.sender_name,`<p>This is follow-up ${n} on the approval request sent on <b>${esc(String(b.sent_at||"").slice(0,10))}</b>. We have not yet received your decision on the stock below. Please reply with <b>APPROVED</b> or <b>REJECTED</b>, your name and the date.</p>${apprTableHtml(rows,fmt,esc)}`),plain:apprPlain(rows)});
  }
  function apprFollowUpSent(id){
    const b=(cache.approvalBatches||[]).find(x=>x.id===id);if(!b)return;
    apprSetStatus(id,{followup_count:(b.followup_count||0)+1,followup_sent_at:new Date().toISOString()},"approval_followup_sent","Follow-up recorded.");
  }
  function apprOpenDecision(id){const p=$("#apprDecisionPanel");if(!p)return;const f=$("#apprDecisionForm");f.reset();f.elements.batch_id.value=id;f.elements.approval_date.value=today();p.classList.remove("hidden");say("apprDecisionMsg","");p.scrollIntoView({behavior:"smooth",block:"center"});}
  async function saveApprovalDecision(e){
    e.preventDefault();
    const f=e.target,id=f.elements.batch_id.value,dec=f.elements.decision.value,file=f.elements.screenshot.files[0];
    const cur=(cache.approvalBatches||[]).find(b=>b.id===id);
    if(!cur){say("apprDecisionMsg","Mail record not found. Reload and try again.");return;}
    let path=cur.screenshot_path||null;
    if(dec==="Approved"&&!file&&!path){say("apprDecisionMsg","An approval needs the screenshot of the approval mail.");return;}
    if(file){const p=`${id}/${Date.now()}_${file.name.replace(/[^\w.\-]/g,"_")}`;const {error:ue}=await sb.storage.from("approval-evidence").upload(p,file);if(ue){say("apprDecisionMsg",friendlyDbError(ue));return;}path=p;}
    const {error}=await sb.from("approval_batches").update({status:dec,approved_by_name:f.elements.approver_name.value.trim()||null,approval_date:f.elements.approval_date.value||null,decision_remarks:f.elements.remarks.value.trim()||null,screenshot_path:path}).eq("id",id);
    if(error){say("apprDecisionMsg",friendlyDbError(error));return;}
    await audit("approval_decision","approval_batch",{id,decision:dec});
    say("apprDecisionMsg","Decision saved.",true);
    $("#apprDecisionPanel").classList.add("hidden");
    await loadAll();
  }
  async function viewApprovalShot(id){
    const b=(cache.approvalBatches||[]).find(x=>x.id===id);if(!b?.screenshot_path)return;
    const {data,error}=await sb.storage.from("approval-evidence").createSignedUrl(b.screenshot_path,600);
    if(error){alert(friendlyDbError(error));return;}
    window.open(data.signedUrl,"_blank","noopener");
  }
  function bindApprovals(){
    $("#apprSettingsForm").onsubmit=apprPrepare;
    $("#apprFrom").onchange=renderApprovals;$("#apprTo").onchange=renderApprovals;$("#apprCase").onchange=renderApprovals;
    $("#apprDecisionForm").onsubmit=saveApprovalDecision;
    $("#apprDecisionCancel").onclick=()=>$("#apprDecisionPanel").classList.add("hidden");
    $("#apprHistory").onclick=e=>{
      const t=e.target.closest&&e.target.closest("button");if(!t)return;const d=t.dataset;
      if(d.decide)apprOpenDecision(d.decide);else if(d.shot)viewApprovalShot(d.shot);else if(d.apprSent)apprMarkSent(d.apprSent);else if(d.apprReopen)apprReopen(d.apprReopen);else if(d.apprCancel)apprCancel(d.apprCancel);else if(d.apprFu)apprFollowUp(d.apprFu);else if(d.apprFusent)apprFollowUpSent(d.apprFusent);
    };
    const cp=async(text,label)=>{try{await navigator.clipboard.writeText(text);say("apprMsg",label+" copied.",true);}catch(_){say("apprMsg","Could not copy "+label+".");}};
    $("#apprCopyTable").onclick=async()=>{if(!apprLast)return;const ok=await apprCopyHtml(apprLast.html,apprLast.plain);say("apprMsg",ok?"Table copied.":"Could not copy the table. Select the preview below and copy it.",ok);};
    $("#apprCopyTo").onclick=()=>apprLast&&cp(apprLast.to,"To");
    $("#apprCopyCc").onclick=()=>apprLast&&cp(apprLast.cc,"CC");
    $("#apprCopySubject").onclick=()=>apprLast&&cp(apprLast.subject,"Subject");
  }
  // ---------- FC mapping: Amazon FC to SAP virtual plant and storage location ----------
  const SAP_SHIP_TO={plant:"DP50",storage:"ED01"};
  function renderFcMap(){
    const tbl=$("#fcMapTable");if(!tbl)return;
    const manager=canEdit();
    $("#fcMapForm")?.classList.toggle("hidden",!manager);
    const rows=[...(cache.fcMappings||[])].sort((x,y)=>String(x.amazon_fc).localeCompare(String(y.amazon_fc)));
    tbl.querySelector("tbody").innerHTML=rows.map(r=>`<tr><td>${esc(r.amazon_fc)}</td><td>${esc(r.sap_virtual_plant||"—")}</td><td>${esc(r.sap_storage_location||"—")}</td><td>${esc(SAP_SHIP_TO.plant)}</td><td>${esc(SAP_SHIP_TO.storage)}</td><td>${esc(r.notes||"—")}</td><td>${manager?`<button class="secondary" data-fc-del="${esc(r.id)}">Remove</button>`:"—"}</td></tr>`).join("")||emptyRow(7);
    enhanceTable("fcMapTable");freezeAll();
  }
  async function saveFcMap(e){
    e.preventDefault();
    if(!canEdit()){say("fcMapMsg","Only a manager or admin can change the FC mapping.");return;}
    const f=e.target;const row={amazon_fc:f.elements.amazon_fc.value.trim().toUpperCase(),sap_virtual_plant:f.elements.sap_virtual_plant.value.trim()||null,sap_storage_location:f.elements.sap_storage_location.value.trim()||null,notes:f.elements.notes.value.trim()||null};
    if(!row.amazon_fc){say("fcMapMsg","Enter the Amazon FC code.");return;}
    const {error}=await sb.from("fc_mappings").upsert(row,{onConflict:"amazon_fc"});
    if(error){say("fcMapMsg",friendlyDbError(error));return;}
    await audit("fc_mapping_saved","fc_mapping",row);
    f.reset();say("fcMapMsg","Saved.",true);await loadAll();
  }
  async function deleteFcMap(id){
    if(!canEdit()||!confirm("Remove this FC mapping?"))return;
    const {error}=await sb.from("fc_mappings").delete().eq("id",id);
    if(error){say("fcMapMsg",friendlyDbError(error));return;}
    await audit("fc_mapping_removed","fc_mapping",{id});await loadAll();
  }
  // ---------- SAP SCAT: PO upload files built from STR ----------
  // SCAT-PURE-START
  const cmpStr=(x,y)=>x<y?-1:x>y?1:0;
  function scatStockIndex(rows){
    const bySl=new Map(),byPlant=new Map(),byMat=new Map(),types=new Set();
    const add=(m,k,t,q)=>{let x=m.get(k);if(!x){x=new Map();m.set(k,x);}x.set(t,(x.get(t)||0)+q);};
    (rows||[]).forEach(s=>{
      const mat=String(s.material||"").trim().replace(/^0+/,""),sl=String(s.sloc||"").trim().toUpperCase(),pl=String(s.plant||"").trim().toUpperCase(),t=String(s.valuation_type||"").trim().toUpperCase(),q=Number(s.actual_stock)||0;
      if(!mat||!t)return;
      types.add(t);add(bySl,mat+"|"+sl,t,q);if(pl)add(byPlant,mat+"|"+pl,t,q);add(byMat,mat,t,q);
    });
    return {bySl,byPlant,byMat,types:[...types].sort()};
  }
  function scatPick(m){
    const arr=[...m.entries()].sort((x,y)=>y[1]-x[1]||cmpStr(x[0],y[0]));
    const detail=arr.map(([t,q])=>t+" "+q).join(" / ");
    if(arr.length===1)return {type:arr[0][0],kind:"single",options:[arr[0][0]],detail};
    const tied=arr.filter(x=>x[1]===arr[0][1]).map(x=>x[0]);
    if(tied.length>1)return {type:"",kind:"tie",options:tied,detail};
    return {type:arr[0][0],kind:"max",options:arr.map(x=>x[0]),detail};
  }
  function scatValuation(idx,mat,plant,sl,manual){
    if(manual)return {type:manual,kind:"manual",level:"manual",options:[manual],detail:"chosen by user"};
    let level="sl",m=idx.bySl.get(mat+"|"+sl);
    if(!m){level="plant";m=idx.byPlant.get(mat+"|"+plant);}
    if(!m){level="any";m=idx.byMat.get(mat);}
    if(!m)return {type:"",kind:"none",level:"none",options:idx.types,detail:"material not in the stock report"};
    return {...scatPick(m),level};
  }
  function scatBuildLine(r,c){
    const issues=[];
    const fc=String(r.ship_from_fc||"").trim().toUpperCase();
    const fm=fc?c.fcMap.get(fc):null;
    if(!fc)issues.push("Ship-from FC is missing in the STR");
    else if(!fm||!fm.plant||!fm.sl)issues.push("FC "+fc+" is not in the FC mapping");
    const mm=c.master(r)||{};
    let material=String(mm.material_code||"").trim();
    if(!material&&/^\d{6,}$/.test(String(r.material_code||"").trim()))material=String(r.material_code).trim();
    if(!material)issues.push("ASIN is not in the product master (no material code)");
    else if(!/^\d+$/.test(material))issues.push("Product master has a material code that is not a single number ("+material+")");
    const qty=c.units(r);
    if(!(qty>0))issues.push("Quantity is zero");
    const tv=(r.taxable_value===null||r.taxable_value===undefined||r.taxable_value==="")?null:Number(r.taxable_value);
    if(tv===null||isNaN(tv))issues.push("Taxable value is missing");
    const value=(tv!==null&&!isNaN(tv)&&qty>0)?Math.round((tv/qty)*100)/100:null;
    const template=/ACCEPT|APPROV/.test(String(r.irn_status||"").toUpperCase())?"approved":"rejected";
    const irn=c.irn(r);
    if(template==="approved"&&!irn)issues.push("IRN status is approved but the IRN number is missing");
    if(!String(r.invoice_no||"").trim())issues.push("Invoice number is missing");
    const okMat=material&&/^\d+$/.test(material)&&fm&&fm.plant&&fm.sl;
    const vt=okMat?scatValuation(c.idx,material.replace(/^0+/,""),fm.plant.toUpperCase(),fm.sl.toUpperCase(),c.manual.get(r.id)||""):{type:"",kind:"none",level:"none",options:c.idx.types,detail:""};
    if(okMat&&!vt.type)issues.push(vt.kind==="tie"?"Valuation type is tied on stock ("+vt.detail+"). Choose one.":"Material is not in the stock report. Choose a valuation type.");
    const status=issues.length?"blocked":((vt.level==="plant"||vt.level==="any")?"review":"ok");
    return {id:r.id,invoice:r.invoice_no||"",invDate:r.invoice_date||"",fc,plant:fm?fm.plant:"",sl:fm?fm.sl:"",material,title:mm.title||"",qty,taxable:tv,value,irnStatus:r.irn_status||"",template,irn,vt:vt.type,vtKind:vt.kind,vtLevel:vt.level,vtOptions:vt.options,vtDetail:vt.detail,issues,status};
  }
  function scatPlan(lines){
    const out={approved:[],rejected:[]};
    ["approved","rejected"].forEach(t=>{
      const rows=lines.filter(l=>l.template===t).sort((x,y)=>cmpStr(x.sl,y.sl)||cmpStr(x.invoice,y.invoice)||cmpStr(x.material,y.material)||cmpStr(x.id,y.id));
      let po=0,item=0,last=null;
      rows.forEach(l=>{if(l.sl!==last){po++;item=0;last=l.sl;}item+=10;out[t].push({...l,po,item});});
    });
    return out;
  }
  function scatSheets(plan,shipTo){
    const A=[["10 20 30","Shipping To","Shipping From",null,null,null,null,null,null,"Shipping to SL",null,"Z103",null,null,"Shipping from",null,null],
      ["PO_ITEM","PLANT","VENDOR","MATERIAL","QUANTITY","PO_UNIT","TAX_CODE","VAL_TYPE","ItM Cat","Reciving STGE_LOC","BATCH","Condition Type","Value","Unit","Issu STGE_LOC","T.P Inv No","T.P  IRN"]];
    plan.approved.forEach(l=>A.push([l.item,shipTo.plant,l.plant,Number(l.material),l.qty,"PC",null,l.vt,null,shipTo.storage,null,"Z103",l.value,"INR",l.sl,l.invoice,l.irn]));
    const R=[["PO_NUMBER","Ship From Plant","Ship TO Plant","PO_ITEM","MATERIAL","QUANTITY","PO_UNIT","VAL_TYPE","Ship to SL","Item Cat","Ship from SL","Invoice no"]];
    plan.rejected.forEach(l=>R.push([l.po,l.plant,shipTo.plant,l.item,Number(l.material),l.qty,"PC",l.vt,shipTo.storage,null,l.sl,l.invoice]));
    return {approved:A,rejected:R};
  }
  // SCAT-PURE-END
  const scatManual=new Map(),scatChecked=new Map();let scatCurrent=[];
  function irnNumberOf(r){const raw=r.raw_data||{};for(const k of Object.keys(raw)){if(String(k).toLowerCase().replace(/[^a-z]/g,"")==="irnnumber"){const v=String(raw[k]??"").trim();if(v)return v;}}return "";}
  function scatContext(){
    const fcMap=new Map((cache.fcMappings||[]).map(m=>[String(m.amazon_fc||"").trim().toUpperCase(),{plant:String(m.sap_virtual_plant||"").trim(),sl:String(m.sap_storage_location||"").trim()}]));
    return {fcMap,idx:scatStockIndex(cache.sapStock),manual:scatManual,master:r=>masterFor(r.sku,null,r.asin),units:unitsOf,irn:irnNumberOf};
  }
  function scatExportedMap(){
    const ex=new Map((cache.scatExports||[]).map(e=>[e.id,e])),m=new Map();
    (cache.scatLines||[]).forEach(l=>{const e=ex.get(l.export_id),d=String(e?.created_at||l.created_at||""),cur=m.get(l.str_record_id);if(!cur||d>cur.date)m.set(l.str_record_id,{date:d,file:e?.file_name||""});});
    return m;
  }
  function scatIsChecked(l){return scatChecked.has(l.id)?scatChecked.get(l.id):(l.status==="ok"&&!l.exported);}
  function scatNote(l){
    if(l.status==="blocked")return l.issues.join("; ");
    const bits=[];
    if(l.vtKind==="max")bits.push("Most stock: "+l.vtDetail);
    if(l.vtLevel==="plant")bits.push("No stock at "+l.sl+"; suggested from plant "+l.plant+" ("+l.vtDetail+"). Please confirm.");
    if(l.vtLevel==="any")bits.push("No stock at plant "+l.plant+"; suggested from other plants ("+l.vtDetail+"). Please confirm.");
    if(l.vtKind==="manual")bits.push("Valuation type chosen by you");
    if(l.exported)bits.push("Exported "+String(l.exported.date).slice(0,10));
    return bits.join("; ")||"—";
  }
  function renderScat(){
    const page=$("#scatPage");if(!page)return;
    const manager=canEdit();
    $("#scatLocked").classList.toggle("hidden",manager);$("#scatBody").classList.toggle("hidden",!manager);
    if(!manager)return;
    const stock=cache.sapStock||[];
    const rd=stock.reduce((acc,s)=>s.report_date>acc?s.report_date:acc,"");
    const age=rd?Math.floor((Date.now()-new Date(rd+"T00:00:00"))/86400000):null;
    $("#scatStrip").innerHTML=stock.length?`Stock report: ${esc(stock[0].source_file||"")} · ${fmt(stock.length)} rows · uploaded ${esc(rd)}${age>2?` <span class="tag bad">${age} days old, upload a fresh report</span>`:""} &nbsp;|&nbsp; FC mappings: ${fmt((cache.fcMappings||[]).length)} &nbsp;|&nbsp; Ship-to: ${SAP_SHIP_TO.plant} / ${SAP_SHIP_TO.storage}`:`<span class="tag bad">No SAP stock report uploaded.</span> Upload it in Reports &amp; Imports (SAP stock report). &nbsp;|&nbsp; FC mappings: ${fmt((cache.fcMappings||[]).length)}`;
    const from=$("#scatFrom").value,to=$("#scatTo").value,tp=$("#scatTemplate").value,show=$("#scatShow").value;
    const ctx=scatContext(),ex=scatExportedMap();
    let lines=(cache.str||[]).map(r=>({...scatBuildLine(r,ctx),exported:ex.get(r.id)||null}));
    lines=lines.filter(l=>(!from||l.invDate>=from)&&(!to||(l.invDate&&l.invDate<=to))&&(tp==="all"||l.template===tp)&&(show==="all"||!l.exported));
    lines.sort((x,y)=>cmpStr(y.invDate,x.invDate)||cmpStr(x.invoice,y.invoice));
    scatCurrent=lines;
    const cnt=s=>lines.filter(l=>l.status===s).length;
    $("#scatSummary").innerHTML=`<div class="kpi-grid compact"><article class="kpi"><span>Ready</span><b>${fmt(cnt("ok"))}</b><small>All fields filled</small></article><article class="kpi"><span>Review</span><b>${fmt(cnt("review"))}</b><small>Valuation type suggested, please confirm</small></article><article class="kpi"><span>Blocked</span><b>${fmt(cnt("blocked"))}</b><small>Fix the reason shown, then they can be exported</small></article></div>`;
    $("#scatTable tbody").innerHTML=lines.map(l=>{
      const opts=[...new Set([l.vt,...l.vtOptions].filter(Boolean))];
      const sel=`<select data-scat-vt="${esc(l.id)}">${l.vt?"":'<option value="">Choose…</option>'}${opts.map(o=>`<option ${o===l.vt?"selected":""}>${esc(o)}</option>`).join("")}</select>`;
      const tag=l.status==="ok"?"ok":l.status==="review"?"warn":"bad";
      return `<tr><td><input type="checkbox" data-scat-check="${esc(l.id)}" ${l.status==="blocked"?"disabled":""} ${scatIsChecked(l)&&l.status!=="blocked"?"checked":""}></td><td>${esc(l.invoice||"—")}</td><td>${esc(l.invDate||"—")}</td><td>${esc(l.fc||"—")}</td><td>${esc(l.plant||"—")}</td><td>${esc(l.sl||"—")}</td><td>${esc(l.material||"—")}</td><td>${esc(l.title||"—")}</td><td>${fmt(l.qty)}</td><td>${l.taxable==null?"—":fmt(l.taxable)}</td><td>${l.value==null?"—":l.value.toFixed(2)}</td><td>${sel}</td><td>${esc(l.irnStatus||"Blank")}</td><td>${l.template==="approved"?"Govt Approved":"Govt Rejected / Blank"}</td><td><span class="tag ${tag}">${l.status==="ok"?"Ready":l.status==="review"?"Review":"Blocked"}</span></td><td style="white-space:normal;min-width:260px">${esc(scatNote(l))}</td></tr>`;
    }).join("")||emptyRow(16);
    enhanceTable("scatTable");freezeAll();updateScatSel();
    const exps=[...(cache.scatExports||[])].sort((x,y)=>String(y.created_at).localeCompare(String(x.created_at)));
    $("#scatHistory tbody").innerHTML=exps.map(e=>`<tr><td>${esc(String(e.created_at||"").slice(0,16).replace("T"," "))}</td><td>${e.template==="approved"?"Govt Approved":"Govt Rejected / Blank"}</td><td>${esc(e.file_name)}</td><td>${fmt(e.line_count)}</td><td>${fmt(e.total_qty)}</td></tr>`).join("")||emptyRow(5);
  }
  function updateScatSel(){const n=scatCurrent.filter(l=>l.status!=="blocked"&&scatIsChecked(l)).length;const el=$("#scatSelCount");if(el)el.textContent=`${n} line(s) selected`;}
  async function scatGenerate(){
    if(!canEdit()){say("scatMsg","Only a manager or admin can generate SCAT files.");return;}
    const lines=scatCurrent.filter(l=>l.status!=="blocked"&&scatIsChecked(l));
    if(!lines.length){say("scatMsg","Select at least one line that is not blocked.");return;}
    const plan=scatPlan(lines),sheets=scatSheets(plan,SAP_SHIP_TO);
    const d=new Date(),stamp=`${d.getFullYear()}${String(d.getMonth()+1).padStart(2,"0")}${String(d.getDate()).padStart(2,"0")}_${String(d.getHours()).padStart(2,"0")}${String(d.getMinutes()).padStart(2,"0")}`;
    const files=[];
    if(plan.approved.length)files.push({t:"approved",name:`SCAT_Govt_Approved_${stamp}.xlsx`,aoa:sheets.approved,rows:plan.approved});
    if(plan.rejected.length)files.push({t:"rejected",name:`SCAT_Govt_Rejected_Blank_${stamp}.xlsx`,aoa:sheets.rejected,rows:plan.rejected});
    say("scatMsg","Saving the export record...",true);
    for(const f of files){
      const {data:eh,error:ee}=await sb.from("scat_exports").insert({template:f.t,file_name:f.name,line_count:f.rows.length,total_qty:f.rows.reduce((acc,l)=>acc+l.qty,0)}).select("id");
      if(ee||!eh?.length){say("scatMsg","Export record was not saved, so no file was created: "+friendlyDbError(ee||"no id returned"));return;}
      f.exportId=eh[0].id;
      const {error:le}=await sb.from("scat_export_lines").insert(f.rows.map(l=>({export_id:f.exportId,str_record_id:l.id,invoice_no:l.invoice,material:l.material,qty:l.qty,value:l.value,val_type:l.vt,val_source:l.vtLevel==="sl"?l.vtKind:l.vtLevel,ship_from_fc:l.fc,ship_from_plant:l.plant,ship_from_sl:l.sl,po_group:l.po,po_item:l.item})));
      if(le){say("scatMsg","Export lines were not saved, so no file was created: "+friendlyDbError(le));return;}
    }
    const X=window.XLSX||await ensureXLSX();
    files.forEach((f,i)=>setTimeout(()=>{const wb=X.utils.book_new();X.utils.book_append_sheet(wb,X.utils.aoa_to_sheet(f.aoa),"Sheet1");X.writeFile(wb,f.name);},i*600));
    await audit("scat_generated","scat_export",{files:files.map(f=>f.name),lines:lines.length});
    say("scatMsg",`${files.length} file(s) created with ${lines.length} line(s). If only one downloaded, allow multiple downloads in your browser.`,true);
    lines.forEach(l=>scatChecked.delete(l.id));
    await loadAll();
  }
  function bindScat(){
    ["scatFrom","scatTo","scatTemplate","scatShow"].forEach(id=>{$("#"+id).onchange=renderScat;});
    $("#scatGenerate").onclick=scatGenerate;
    $("#scatSelectReady").onclick=()=>{scatCurrent.forEach(l=>{if(l.status!=="blocked")scatChecked.set(l.id,true);});renderScat();};
    $("#scatClear").onclick=()=>{scatCurrent.forEach(l=>scatChecked.set(l.id,false));renderScat();};
    $("#scatTable").addEventListener("change",e=>{
      const c=e.target.closest&&e.target.closest("[data-scat-check]");
      if(c){scatChecked.set(c.dataset.scatCheck,c.checked);updateScatSel();return;}
      const s=e.target.closest&&e.target.closest("[data-scat-vt]");
      if(s){if(s.value)scatManual.set(s.dataset.scatVt,s.value);else scatManual.delete(s.dataset.scatVt);renderScat();}
    });
  }
  // ---------- SAP stock report upload (valuation type source) ----------
  async function importStock(rows,file){
    if(!canEdit())throw new Error("Only a manager or admin can upload the SAP stock report.");
    const first=rows[0]||{},keys=Object.keys(first).map(k=>k.trim().toLowerCase());
    for(const need of ["material code","plant","valuation type","st. loc","actual stock"])if(!keys.includes(need))throw new Error('This does not look like the SAP stock report. Column "'+need+'" was not found.');
    const g=(r,k)=>{const v=r[k];return v===undefined||v===null?"":String(v).trim();};
    const uploadId=crypto.randomUUID(),reportDate=today(),recs=[];let skipped=0;
    rows.forEach(r0=>{
      const r={};Object.keys(r0).forEach(k=>r[k.trim().toLowerCase()]=r0[k]);
      const material=g(r,"material code").replace(/^0+/,""),sloc=g(r,"st. loc").toUpperCase(),vt=g(r,"valuation type").toUpperCase();
      if(!material||!sloc||!vt){skipped++;return;}
      recs.push({material,plant:g(r,"plant").toUpperCase()||null,sloc,valuation_type:vt,batch:g(r,"batch number")||null,actual_stock:num(g(r,"actual stock")),report_date:reportDate,source_file:file.name,upload_id:uploadId});
    });
    if(!recs.length)throw new Error("No usable rows were found in the stock report.");
    for(let i=0;i<recs.length;i+=500){
      const {error}=await sb.from("sap_stock").insert(recs.slice(i,i+500));
      if(error){await sb.from("sap_stock").delete().eq("upload_id",uploadId);throw new Error("Stock report was not saved: "+friendlyDbError(error));}
    }
    const {error:de}=await sb.from("sap_stock").delete().neq("upload_id",uploadId);
    if(de)throw new Error("The new stock report was saved, but the old one could not be removed: "+friendlyDbError(de));
    await audit("stock_report_uploaded","sap_stock",{file:file.name,rows:recs.length,skipped});
    await addImport(file,"sap_stock_report",rows.length,recs.length);
    return {accepted:recs.length,skipped};
  }
  // ---------- Evidence (photos and videos) for damaged, short and not received stock ----------
  const EVIDENCE_MAX_MB={image:10,video:30};
  let evidenceCtx=null;
  function evidenceFor(shipmentId){return (cache.evidence||[]).filter(e=>e.shipment_line_id===shipmentId);}
  function evidenceNeeded(r){return num(r.damaged_qty)+num(r.short_qty)+num(r.not_received_qty)>0;}
  function checkEvidenceFiles(files){
    const errs=[];
    [...files].forEach(f=>{const kind=String(f.type).startsWith("image/")?"image":String(f.type).startsWith("video/")?"video":null;if(!kind){errs.push(f.name+": only photos and videos are allowed");return;}if(f.size>EVIDENCE_MAX_MB[kind]*1048576)errs.push(f.name+": larger than "+EVIDENCE_MAX_MB[kind]+" MB");});
    return errs;
  }
  async function uploadEvidence(shipmentId,files){
    const errs=[];let ok=0;
    for(const f of files){
      const p=`${shipmentId}/${Date.now()}_${f.name.replace(/[^\w.\-]/g,"_")}`;
      const {error:ue}=await sb.storage.from("receipt-evidence").upload(p,f,{contentType:f.type});
      if(ue){errs.push(f.name+": "+friendlyDbError(ue));continue;}
      const {error:ie}=await sb.from("receipt_evidence").insert({shipment_line_id:shipmentId,file_path:p,file_name:f.name,mime_type:f.type,size_bytes:f.size});
      if(ie){errs.push(f.name+": "+friendlyDbError(ie));continue;}
      ok++;
    }
    if(ok)await audit("evidence_uploaded","receipt_evidence",{shipment_line_id:shipmentId,files:ok});
    return {ok,errs};
  }
  function renderEvidenceList(){
    const list=evidenceFor(evidenceCtx);
    $("#evidenceList").innerHTML=list.length?list.map(e=>`<div style="display:flex;gap:8px;align-items:center;margin:6px 0"><span style="flex:1">${esc(e.file_name)} <small class="muted">${esc(String(e.created_at||"").slice(0,10))}</small></span><button type="button" class="secondary" data-evidence-view="${esc(e.id)}">View</button></div>`).join(""):'<p class="muted">No evidence uploaded yet.</p>';
  }
  function openEvidence(shipmentId){
    evidenceCtx=shipmentId;
    const s=(cache.shipments||[]).find(x=>x.id===shipmentId)||{};
    {const mx=masterFor(s.sku,s.fnsku,null);$("#evidenceTitle").textContent=`${s.order_id||""} · ${s.tracking_number||""} · ${mx.material_code||"Not mapped"} ${mx.title||""}`;}
    $("#evidenceFile").value="";say("evidenceMsg","");renderEvidenceList();
    $("#evidenceModal").classList.remove("hidden");
  }
  async function viewEvidence(id){
    const e=(cache.evidence||[]).find(x=>x.id===id);if(!e)return;
    const {data,error}=await sb.storage.from("receipt-evidence").createSignedUrl(e.file_path,600);
    if(error){alert(friendlyDbError(error));return;}
    window.open(data.signedUrl,"_blank","noopener");
  }
  function bindEvidence(){
    $("#receiptLineEditor").addEventListener("input",e=>{
      const q=e.target.closest&&e.target.closest("[data-qty]");if(!q)return;
      const line=q.closest(".receipt-line");if(!line)return;
      const v=k=>num(line.querySelector(`[data-qty="${k}"]`)?.value);
      line.querySelector("[data-evidence-row]")?.classList.toggle("hidden",!(v("damaged_qty")+v("short_qty")+v("not_received_qty")>0));
    });
    $("#evidenceClose").onclick=()=>$("#evidenceModal").classList.add("hidden");
    $("#evidenceUpload").onclick=async()=>{
      const files=$("#evidenceFile").files;
      if(!files.length){say("evidenceMsg","Choose at least one photo or video.");return;}
      const errs=checkEvidenceFiles(files);if(errs.length){say("evidenceMsg",errs.join("\n"));return;}
      say("evidenceMsg","Uploading...",true);
      const r=await uploadEvidence(evidenceCtx,files);
      say("evidenceMsg",r.errs.length?`${r.ok} uploaded. Problems: ${r.errs.join(" | ")}`:`${r.ok} file(s) uploaded.`,!r.errs.length);
      await loadAll();renderEvidenceList();$("#evidenceFile").value="";
    };
    $("#evidenceList").onclick=e=>{const b=e.target.closest&&e.target.closest("[data-evidence-view]");if(b)viewEvidence(b.dataset.evidenceView);};
    $("#receiptEntriesTable").addEventListener("click",e=>{const b=e.target.closest&&e.target.closest("[data-evidence-open]");if(b)openEvidence(b.dataset.evidenceOpen);});
  }
  const COURIER_TRACK_PAGES={BLUEDART:{base:"https://www.bluedart.com/home",direct:false},DLV:{base:"https://www.delhivery.com/track-v2/package/",direct:true}};
  function courierPageFor(carrier){const c=String(carrier||"").toUpperCase();if(c.includes("BLUEDART")||c.includes("BLUE_DART"))return COURIER_TRACK_PAGES.BLUEDART;if(c.includes("DLV"))return COURIER_TRACK_PAGES.DLV;return null;}
  async function trackShipment(tracking,carrier){
    try{await navigator.clipboard.writeText(String(tracking));}catch(_){}
    const page=courierPageFor(carrier);
    if(!page){alert(`Tracking ${tracking} copied.\n\nThe courier tracking page for "${carrier||"unknown carrier"}" is not set up yet. Open the courier site and paste the number.`);return;}
    if(page.direct){window.open(page.base+encodeURIComponent(String(tracking)),"_blank","noopener");return;}
    window.open(page.base,"_blank","noopener");
    alert(`Tracking number ${tracking} copied.\n\nOn the Blue Dart page, paste it (Ctrl+V) and press Track.`);
  }
  function renderRecon(){
    const tbl=$("#reconTable");if(!tbl)return;
    const rows=reconRows();
    const st=$("#reconStatus")?.value||"all",q=norm($("#reconSearch")?.value||"");
    const shown=rows.filter(r=>(st==="all"||(st==="matched"?r.status==="Matched":r.status!=="Matched"))&&(!q||norm([r.o.order_id,masterFor(r.o.sku,r.o.fnsku,null).material_code,masterFor(r.o.sku,r.o.fnsku,null).title,r.tracking,r.invs.map(i=>i.invoice_no).join(" "),r.shipIds.join(" ")].join(" ")).includes(q)));
    const sum=f=>shown.reduce((a,r)=>a+f(r),0);
    const sumEl=$("#reconSummary");
    if(sumEl)sumEl.innerHTML=`<div class="kpi-grid compact"><article class="kpi"><span>Order × product lines</span><b>${fmt(shown.length)}</b></article><article class="kpi"><span>Matched</span><b>${fmt(shown.filter(r=>r.status==="Matched").length)}</b></article><article class="kpi"><span>Exceptions</span><b>${fmt(shown.filter(r=>r.status!=="Matched").length)}</b></article><article class="kpi"><span>Order shipped units</span><b>${fmt(sum(r=>num(r.o.shipped_qty)))}</b></article><article class="kpi"><span>Returned to Amazon</span><b>${fmt(sum(r=>r.returned))}</b></article><article class="kpi"><span>Pending at warehouse</span><b>${fmt(sum(r=>r.pending))}</b></article></div>`;
    tbl.querySelector("tbody").innerHTML=shown.map(r=>{const m=masterFor(r.o.sku,r.o.fnsku,null),po=r.po;return `<tr><td>${esc(r.o.order_id)}</td><td>${esc(m.material_code||"Not mapped")}</td><td>${esc(m.title||"—")}</td><td>${esc(r.invs.map(i=>i.invoice_no).join(", ")||"—")}</td><td>${esc(r.invs.map(i=>i.invoice_date).filter(Boolean).join(", ")||"—")}</td><td>${esc(r.shipIds.join(", ")||"—")}</td><td>${esc([...new Set(r.invs.map(i=>i.ra_key).filter(Boolean))].join(", ")||"—")}</td><td>${fmt(r.strUnits)}</td><td>${esc(po.po_number||"Awaiting PO")}</td><td>${esc(po.delivery_no||"—")}</td><td>${esc(po.billing_document||"—")}</td><td>${esc(po.migo_no||"—")}</td><td>${fmt(r.o.requested_qty)}</td><td>${fmt(r.o.shipped_qty)}</td><td>${fmt(r.shipTotal)}</td><td>${fmt(r.returned)}</td><td>${fmt(r.expected)}</td><td>${fmt(r.phys)}</td><td>${fmt(r.short+r.notrec)}</td><td><b>${fmt(r.pending)}</b></td><td>${esc(r.tracking||"—")}</td><td>${r.fee==null?"—":esc(r.fee.toFixed(2))}</td><td>${esc(r.cur||"—")}</td><td>${esc(r.batchTxt||"—")}</td><td><span class="tag ${tagFor(r.status)}">${esc(r.status)}</span></td></tr>`;}).join("")||emptyRow(25);
    enhanceTable("reconTable");freezeAll();
  }
  document.addEventListener("DOMContentLoaded",()=>{bind();init();});
})();