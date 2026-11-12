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
    $$(".admin-only").forEach(el=>el.classList.toggle("hidden",!["admin","manager"].includes(p?.role)));
    $("#loginView").classList.add("hidden");$("#appView").classList.remove("hidden");await loadAll();installHorizontalScrollControls();
  }
  async function loadAll(){
    const tables=["removal_orders","shipment_lines","warehouse_receipts","amazon_removal_fees","amazon_claims","sap_postings","import_batches","str_records","product_mappings"];
    const names=["orders","shipments","receipts","fees","claims","sap","imports","str","mappings"];
    for(let i=0;i<tables.length;i++){const {data,error}=await sb.from(tables[i]).select("*").order("created_at",{ascending:false}).limit(5000);if(!error)cache[names[i]]=data||[];}
    renderAll();
  }
  function fyDate(d){if(!d)return null;const x=new Date(d+"T00:00:00");return x;}
  function filtered(arr,dateKey){const m=$("#monthFilter").value,fy=$("#fyFilter").value;return arr.filter(x=>{const d=fyDate(x[dateKey]);if(!d)return m==="all"&&fy==="all";const mm=String(d.getMonth()+1).padStart(2,"0"),month=d.getFullYear()+"-"+mm;const fiscal=d.getMonth()>=3?`${d.getFullYear()}-${String(d.getFullYear()+1).slice(-2)}`:`${d.getFullYear()-1}-${String(d.getFullYear()).slice(-2)}`;return(m==="all"||m===month)&&(fy==="all"||fy===fiscal);});}
  function receiptTotals(shipmentId){return cache.receipts.filter(r=>r.shipment_line_id===shipmentId).reduce((a,r)=>{for(const k of ["good_qty","damaged_qty","near_expiry_qty","expired_qty","short_qty","not_received_qty"])a[k]+=num(r[k]);return a;},{good_qty:0,damaged_qty:0,near_expiry_qty:0,expired_qty:0,short_qty:0,not_received_qty:0});}
  function renderAll(){
    const orders=filtered(cache.orders,"request_date"),ships=filtered(cache.shipments,"shipment_date"),fees=filtered(cache.fees,"fee_date");
    const shipped=ships.reduce((a,x)=>a+num(x.shipped_qty),0),rt={good_qty:0,damaged_qty:0,near_expiry_qty:0,expired_qty:0,short_qty:0,not_received_qty:0};
    ships.forEach(s=>{const t=receiptTotals(s.id);Object.keys(rt).forEach(k=>rt[k]+=t[k]);});
    const physical=rt.good_qty+rt.damaged_qty+rt.near_expiry_qty+rt.expired_qty,final=Object.values(rt).reduce((a,b)=>a+b,0),pending=Math.max(0,shipped-final);
    const feeCurrencies=[...new Set(fees.map(x=>x.currency).filter(Boolean))],feeTotal=fees.reduce((a,x)=>a+num(x.fee_amount),0),claims=cache.claims.filter(c=>!["closed","resolved"].includes(norm(c.status)));
    $("#kpiOrders").textContent=fmt(new Set(orders.map(o=>o.order_id)).size);$("#kpiShipped").textContent=fmt(shipped);$("#kpiReceived").textContent=fmt(physical);$("#kpiPending").textContent=fmt(pending);$("#kpiFees").textContent=feeCurrencies.length>1?"Multiple currencies":money(feeTotal,feeCurrencies[0]||"INR");$("#kpiClaims").textContent=fmt(claims.length);
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
  function receiptSummary(s){const t=receiptTotals(s.id),physical=t.good_qty+t.damaged_qty+t.near_expiry_qty+t.expired_qty,allocated=physical+t.short_qty+t.not_received_qty;return {...t,physical,pending:Math.max(0,num(s.shipped_qty)-allocated)};}
  function renderOrders(){
    const q=norm($("#orderSearch")?.value);
    const rows=filtered(cache.orders,"request_date").filter(x=>{const st=strFor(x.order_id,x.sku);return [x.order_id,x.sku,x.fnsku,x.sku,st.transaction_id,st.po_number,st.delivery_no,st.billing_document,st.migo_no].some(v=>norm(v).includes(q));}).slice(0,5000);
    $("#ordersTable").innerHTML=rows.map(x=>{const st=strFor(x.order_id,x.sku);const totals=cache.shipments.filter(z=>z.order_id===x.order_id&&((x.fnsku&&z.fnsku===x.fnsku)||(x.sku&&z.sku===x.sku))).reduce((a,z)=>{const t=receiptSummary(z);a.physical+=t.physical;a.short+=t.short_qty;a.notReceived+=t.not_received_qty;a.pending+=t.pending;return a;},{physical:0,short:0,notReceived:0,pending:0});return `<tr><td>${esc(x.order_id)}</td><td>${esc(x.sku)}</td><td>${esc(x.fnsku)}</td><td>${esc(x.sku||"—")}</td><td>—</td><td>${fmt(x.requested_qty)}</td><td>${fmt(x.cancelled_qty)}</td><td>${fmt(x.shipped_qty)}</td><td>${fmt(totals.physical)}</td><td>${fmt(totals.pending+totals.short+totals.notReceived)}</td><td>${esc(st.po_number||"Pending")}</td><td>${esc(st.delivery_no||"—")}</td><td>${esc(st.billing_document||"—")}</td><td>${esc(st.migo_no||"—")}</td><td><span class="tag">${esc(x.status||"—")}</span></td></tr>`;}).join("")||emptyRow(15);enhanceTable("ordersTable");
  }
  function renderShipments(){
    const q=norm($("#shipmentSearch")?.value);
    const rows=filtered(cache.shipments,"shipment_date").filter(x=>{const y=enrichShipment(x),st=strFor(x.order_id,y.sku);return [x.order_id,x.tracking_number,x.fnsku,x.id,y.sku,y.material_code,st.po_number,st.delivery_no,st.billing_document,st.migo_no].some(v=>norm(v).includes(q));}).slice(0,5000);
    $("#shipmentsTable").innerHTML=rows.map(x=>{const y=enrichShipment(x),st=strFor(x.order_id,y.sku),t=receiptSummary(x);return `<tr><td>${esc(x.id.slice(0,8))}</td><td>${esc(x.order_id)}</td><td>${esc(x.tracking_number)}</td><td>${esc(y.sku||"—")}</td><td>${esc(x.fnsku||"—")}</td><td>${esc(y.sku||"—")}</td><td>—</td><td>${fmt(x.shipped_qty)}</td><td>${fmt(t.physical)}</td><td>${fmt(t.pending)}</td><td>${fmt(t.short_qty)}</td><td>${fmt(t.not_received_qty)}</td><td>${esc(st.po_number||"Pending")}</td><td>${esc(st.delivery_no||"—")}</td><td>${esc(st.billing_document||"—")}</td><td>${esc(st.migo_no||"—")}</td><td>${esc(x.amazon_status||"—")}</td><td>${esc(x.shipment_date||"—")}</td></tr>`;}).join("")||emptyRow(18);enhanceTable("shipmentsTable");
  }
  // Header sorting and per-column filters are available on each main tracker.
  const tableState={};
  function enhanceTable(tbodyId){
    const body=$("#"+tbodyId),table=body?.closest("table");if(!table)return;const head=table.tHead?.rows[0];if(!head)return;
    let state=tableState[tbodyId]||(tableState[tbodyId]={filters:{},sortCol:-1,asc:true});
    let filterRow=table.querySelector("thead tr.column-filters");
    if(!filterRow){filterRow=document.createElement("tr");filterRow.className="column-filters";[...head.cells].forEach((th,i)=>{const cell=document.createElement("th"),input=document.createElement("input");input.className="column-filter";input.placeholder="Filter…";input.setAttribute("aria-label",`Filter ${th.textContent}`);input.value=state.filters[i]||"";input.onclick=e=>e.stopPropagation();input.oninput=()=>{state.filters[i]=input.value;applyTableFilters(tbodyId);};cell.append(input);filterRow.append(cell);th.title="Click to sort";th.style.cursor="pointer";th.onclick=()=>{if(state.sortCol===i)state.asc=!state.asc;else{state.sortCol=i;state.asc=true;}sortTableRows(tbodyId);};});table.tHead.append(filterRow);}
    if(state.sortCol>=0)sortTableRows(tbodyId);else applyTableFilters(tbodyId);
  }
  function applyTableFilters(id){const body=$("#"+id),state=tableState[id];if(!body||!state)return;[...body.rows].forEach(row=>{const ok=Object.entries(state.filters).every(([i,v])=>!v||norm(row.cells[Number(i)]?.textContent).includes(norm(v)));row.hidden=!ok;});}
  function sortTableRows(id){const body=$("#"+id),state=tableState[id];if(!body||!state||state.sortCol<0)return;const rows=[...body.rows];rows.sort((a,b)=>{const x=a.cells[state.sortCol]?.textContent.trim()||"",y=b.cells[state.sortCol]?.textContent.trim()||"",nx=Number(x.replace(/,/g,"")),ny=Number(y.replace(/,/g,""));const cmp=x!==""&&y!==""&&Number.isFinite(nx)&&Number.isFinite(ny)?nx-ny:x.localeCompare(y,undefined,{numeric:true,sensitivity:"base"});return state.asc?cmp:-cmp;});rows.forEach(r=>body.append(r));applyTableFilters(id);}
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
    sel.innerHTML='<option value="">Select '+(type==="order"?'Amazon Order ID':'Tracking Number')+'</option>'+vals.map(v=>`<option value="${esc(v)}">${esc(v)}</option>`).join("");if(vals.includes(previous))sel.value=previous;renderReceiptLineEditor();
  }
  function renderReceiptLineEditor(){
    const type=$("#receiptLookupType")?.value||"order",value=$("#receiptLookupValue")?.value,box=$("#receiptLineEditor");if(!box)return;
    if(!value){box.innerHTML='<div class="notice">Select an Order ID or Tracking Number to load all associated material lines and shipped quantities.</div>';return;}
    const lines=cache.shipments.filter(s=>(type==="order"?s.order_id:s.tracking_number)===value);
    box.innerHTML=lines.map(raw=>{const s=enrichShipment(raw),t=receiptTotals(s.id),used=Object.values(t).reduce((a,b)=>a+num(b),0),remaining=Math.max(0,num(s.shipped_qty)-used);return `<div class="receipt-line" data-shipment="${esc(s.id)}"><div class="receipt-line-head"><b>${esc(s.sku||'SKU unavailable')} · Material Code ${esc(s.sku||'—')}</b><span>Tracking: ${esc(s.tracking_number||'—')} · FNSKU: ${esc(s.fnsku||'—')}</span><strong>Amazon shipped: ${fmt(s.shipped_qty)} | Previously allocated: ${fmt(used)} | Remaining: ${fmt(remaining)}</strong><small>Amazon status: ${esc(s.amazon_status||'—')} (does not confirm physical receipt)</small></div><div class="receipt-qty-grid">${[['good_qty','Good'],['damaged_qty','Damaged'],['near_expiry_qty','Near Expiry'],['expired_qty','Expired'],['short_qty','Short'],['not_received_qty','Not Received']].map(([k,label])=>`<label>${label}<input type="number" min="0" step="1" value="0" data-qty="${k}" inputmode="numeric"></label>`).join('')}</div></div>`;}).join('')||'<div class="notice">No shipment lines found.</div>';
  }
  function renderReceipts(){
    renderReceiptSelector();
    $("#receiptTable").innerHTML=cache.shipments.map(raw=>{const s=enrichShipment(raw),t=receiptTotals(s.id),physical=t.good_qty+t.damaged_qty+t.near_expiry_qty+t.expired_qty,alloc=physical+t.short_qty+t.not_received_qty,pending=Math.max(0,num(s.shipped_qty)-alloc);return `<tr><td>${esc(s.order_id)}</td><td>${esc(s.sku||"—")}</td><td>${esc(s.fnsku||"—")}</td><td>${esc(s.sku||"—")}</td><td>${esc(s.tracking_number)}</td><td>${fmt(s.shipped_qty)}</td><td>${fmt(t.good_qty)}</td><td>${fmt(t.damaged_qty)}</td><td>${fmt(t.near_expiry_qty)}</td><td>${fmt(t.expired_qty)}</td><td>${fmt(physical)}</td><td>${fmt(t.short_qty)}</td><td>${fmt(t.not_received_qty)}</td><td><b>${fmt(pending)}</b></td><td><span class="tag">${esc(s.mapping_status)}</span></td></tr>`;}).join("")||emptyRow(15);enhanceTable("receiptTable");
  }
  function renderFees(){const fees=cache.fees;const currencies=[...new Set(fees.map(x=>x.currency).filter(Boolean))], total=fees.reduce((a,x)=>a+num(x.fee_amount),0),qty=fees.reduce((a,f)=>{const o=cache.orders.find(x=>x.order_id===f.order_id&&((f.fnsku&&x.fnsku===f.fnsku)||(f.sku&&x.sku===f.sku)));return a+num(o?.shipped_qty);},0);$("#feeTotal").textContent=currencies.length<=1?money(total,currencies[0]||"INR"):"Multiple currencies";$("#feeCount").textContent=fmt(fees.length);$("#feePerUnit").textContent=currencies.length===1&&qty?money(total/qty,currencies[0]):"—";$("#feesTable").innerHTML=fees.slice(0,500).map(f=>`<tr><td>${esc(f.order_id)}</td><td>${esc(f.sku||"—")}</td><td>${esc(f.fnsku||"—")}</td><td>${money(f.fee_amount,f.currency||"INR")}</td><td>${esc(f.currency||"—")}</td><td>${esc(f.fee_date||"—")}</td><td>${esc(f.source_file||"Amazon report")}</td></tr>`).join("")||emptyRow(7);}
  function renderClaims(){$("#claimsTable").innerHTML=cache.claims.map(c=>`<tr><td>${esc(c.order_id)}</td><td>${esc(c.claim_ref||"—")}</td><td>${esc(c.claim_type)}</td><td>${fmt(c.quantity)}</td><td>${esc(c.status)}</td><td>${c.amount==null?"—":money(c.amount)}</td></tr>`).join("")||emptyRow(6);}
  function renderSAP(){
    const records=cache.str||[],groups=new Map();
    records.forEach(r=>{const ref=String(r.transaction_id||r.invoice_no||r.id||"NO-REFERENCE");const key=`${ref}|${norm(r.invoice_no)}`;if(!groups.has(key))groups.set(key,{key,ref,ship:r.ship_from_fc||"—",rows:[],skus:new Set(),qty:0,po:new Set(),delivery:new Set(),billing:new Set(),migo:new Set(),statuses:new Set()});const g=groups.get(key);g.rows.push(r);if(r.sku)g.skus.add(r.sku);g.qty+=num(r.quantity);if(r.po_number)g.po.add(r.po_number);if(r.delivery_no)g.delivery.add(r.delivery_no);if(r.billing_document)g.billing.add(r.billing_document);if(r.migo_no)g.migo.add(r.migo_no);if(r.sap_status)g.statuses.add(r.sap_status);});
    const list=[...groups.values()].sort((a,b)=>a.ref.localeCompare(b.ref));
    const kpi=document.querySelector("#sapSummary");if(kpi)kpi.innerHTML=`<div class="kpi-grid compact"><article class="kpi"><span>STR lines</span><b>${fmt(records.length)}</b></article><article class="kpi"><span>Individual reference groups</span><b>${fmt(list.length)}</b></article><article class="kpi"><span>PO pending</span><b>${fmt(list.filter(g=>!g.po.size).length)}</b></article><article class="kpi"><span>SAP documents pending</span><b>${fmt(list.filter(g=>!g.delivery.size||!g.billing.size||!g.migo.size).length)}</b></article></div>`;
    $("#sapTable").innerHTML=list.map(g=>`<tr><td>${esc(g.ref)}</td><td>${esc(g.ship)}</td><td>${esc(g.rows[0].irn_status||"—")}</td><td>${fmt(g.rows.length)}</td><td>${esc([...g.skus].join(", ")||"—")}</td><td>${fmt(g.qty)}</td><td>${esc([...g.po].join(", ")||"Pending")}</td><td>${esc([...g.delivery].join(", ")||"—")}</td><td>${esc([...g.billing].join(", ")||"—")}</td><td>${esc([...g.migo].join(", ")||"—")}</td></tr>`).join("")||emptyRow(10);
    $("#sapExceptions").innerHTML=records.filter(r=>!r.po_number||!r.delivery_no||!r.billing_document||!r.migo_no||!r.sku).slice(0,1000).map(r=>`<tr><td>Amazon STR</td><td>${esc(r.transaction_id)}</td><td>${esc(r.sku||"—")}</td><td>${esc(!r.sku?"Missing SKU / SAP Material Code":!r.po_number?"PO pending":!r.delivery_no?"Delivery pending":!r.billing_document?"Billing document pending":"MIGO pending")}</td><td>${esc(r.irn_status||"—")}</td><td>Review</td></tr>`).join("")||emptyRow(6);
  }
  function renderImports(){$("#importsTable").innerHTML=cache.imports.map(x=>`<tr><td>${esc(x.file_name)}</td><td>${esc(x.report_type)}</td><td>${fmt(x.rows_read)}</td><td>${fmt(x.rows_accepted)}</td><td>${fmt(x.rejected_count||0)}</td><td>${esc((x.error_details||[]).slice(0,2).join(" | "))}</td><td>${esc(x.imported_by_email||"—")}</td><td>${esc(x.created_at?.slice(0,10)||"—")}</td></tr>`).join("")||emptyRow(8);}
  async function renderEmployees(){if(!["admin","manager"].includes(profile?.role))return;const {data}=await sb.from("profiles").select("id,email,full_name,role").order("created_at");$("#employeesTable").innerHTML=(data||[]).map(p=>`<tr><td>${esc(p.full_name)}</td><td>${esc(p.email)}</td><td>${esc(p.role)}</td><td>${p.id===session.user.id?"Current user":`<select data-role-user="${p.id}"><option ${p.role==="employee"?"selected":""}>employee</option><option ${p.role==="manager"?"selected":""}>manager</option><option ${p.role==="admin"?"selected":""}>admin</option></select>`}</td></tr>`).join("")||emptyRow(4);$$("[data-role-user]").forEach(el=>el.addEventListener("change",async()=>{const {error}=await sb.from("profiles").update({role:el.value}).eq("id",el.dataset.roleUser);if(error)alert(error.message);else{await audit("role_updated","profile",{user_id:el.dataset.roleUser,role:el.value});await renderEmployees();}}));}
  async function renderAudit(){if(!["admin","manager"].includes(profile?.role))return;const {data}=await sb.from("audit_logs").select("*,profiles(full_name,email)").order("created_at",{ascending:false}).limit(300);$("#auditTable").innerHTML=(data||[]).map(a=>`<tr><td>${esc(a.created_at?.replace("T"," ").slice(0,19))}</td><td>${esc(a.profiles?.full_name||a.profiles?.email||"—")}</td><td>${esc(a.action)}</td><td>${esc(a.entity)}</td><td>${esc(JSON.stringify(a.details||{}))}</td></tr>`).join("")||emptyRow(5);}
  function nav(page){$$(".page").forEach(p=>p.classList.add("hidden"));$("#"+page+"Page").classList.remove("hidden");$$(".nav-item").forEach(n=>n.classList.toggle("active",n.dataset.page===page));$("#pageTitle").textContent=({dashboard:"Dashboard",orders:"Removal Orders",shipments:"Shipment Tracker",receipts:"Warehouse Receipts",fees:"Amazon Removal Fees",sap:"PO & SAP Tracker",claims:"Amazon Claims",reports:"Reports & Imports",employees:"Employee Management",audit:"Audit Logs"})[page]||"Dashboard";}
  async function addImport(file,type,read,accepted){const {error}=await sb.from("import_batches").insert({file_name:file.name,report_type:type,rows_read:read,rows_accepted:accepted,rejected_count:importProblemCount,error_details:importProblems,imported_by:session.user.id,imported_by_email:session.user.email});if(error)throw error;await audit("report_imported","import_batch",{file:file.name,type,rows_read:read,rows_accepted:accepted,rejected_count:importProblemCount});}
  const importProblems=[];let importProblemCount=0;
  function problem(file,row,reason){importProblemCount++;if(importProblems.length<100)importProblems.push(`${file.name} — row ${row}: ${reason}`);}
  function uniqueMap(rows,keys){for(const key of keys){const v=norm(key.value);if(!v)continue;const found=rows.filter(m=>norm(m[key.field])===v);if(found.length){const signatures=[...new Set(found.map(m=>`${norm(m.sku)}|${norm(m.material_code)}`))];return {row:signatures.length===1?found[0]:null,conflict:signatures.length>1,matchedBy:key.label};}}return {row:null,conflict:false,matchedBy:null};}
  async function importOrders(rows,file){let accepted=0,feeAccepted=0;for(let i=0;i<rows.length;i++){const r=rows[i];const x={order_id:pick(r,"order-id","order id","removal-order-id"),sku:pick(r,"sku"),fnsku:pick(r,"fnsku"),material_code:pick(r,"material-code","material code","material"),requested_qty:num(pick(r,"requested-quantity","requested qty")),cancelled_qty:num(pick(r,"cancelled-quantity","cancelled qty")),shipped_qty:num(pick(r,"shipped-quantity","shipped qty")),in_process_qty:num(pick(r,"in-process-quantity","in-process qty")),status:pick(r,"status","order-status"),request_date:dateISO(pick(r,"request-date","created-date")),raw_data:r};if(!x.order_id){problem(file,i+2,'Missing Order ID');continue;}const {error}=await sb.from("removal_orders").upsert(x,{onConflict:"order_id,sku,fnsku"});if(error)problem(file,i+2,error.message);else accepted++;
      const rawFee=pick(r,"removal-fee","removal fee"),currency=pick(r,"currency");if(String(rawFee).trim()!==""&&currency){const feeRow={order_id:x.order_id,sku:x.sku||"",fnsku:x.fnsku||"",fee_amount:num(rawFee),currency,fee_date:x.request_date,source_file:file.name,raw_data:r};const {error:feeError}=await sb.from("amazon_removal_fees").upsert(feeRow,{onConflict:"order_id,sku,fnsku,source_file"});if(!feeError)feeAccepted++;else problem(file,i+2,`Fee: ${feeError.message}`);}}
    await addImport(file,"removal_order_detail",rows.length,accepted);return {accepted,feeAccepted};}
  async function importShipments(rows,file){let accepted=0;for(let i=0;i<rows.length;i++){const r=rows[i],order_id=pick(r,"order-id","order id"),tracking_number=pick(r,"tracking-number","tracking number"),fnsku=pick(r,"fnsku"),sku=pick(r,"sku"),asin=pick(r,"asin");if(!order_id||!tracking_number){problem(file,i+2,'Missing Order ID or Tracking Number');continue;}const order=cache.orders.find(o=>o.order_id===order_id&&((fnsku&&o.fnsku===fnsku)||(sku&&o.sku===sku)));const match=uniqueMap(cache.mappings,[{field:'fnsku',label:'FNSKU',value:fnsku},{field:'sku',label:'SKU',value:sku},{field:'asin',label:'ASIN',value:asin}].filter(x=>x.value));const mapped=match.row;const x={order_id,tracking_number,fnsku,sku:sku||order?.sku||mapped?.sku||null,material_code:order?.material_code||mapped?.material_code||null,mapping_status:match.conflict?'conflict':(sku||order?.sku||mapped)?'matched':'unmapped',shipped_qty:num(pick(r,"shipped-quantity","quantity-shipped","quantity")),amazon_status:pick(r,"status","shipment-status"),shipment_date:dateISO(pick(r,"shipment-date","shipped-date")),disposition:pick(r,"disposition"),raw_data:r};const {error}=await sb.from("shipment_lines").insert(x);if(!error)accepted++;else problem(file,i+2,error.message);}await addImport(file,"removal_shipment_detail",rows.length,accepted);return {accepted};}
  async function importMappings(rows,file){let accepted=0;for(let i=0;i<rows.length;i++){const r=rows[i],x={asin:pick(r,"asin","asin.1"),sku:pick(r,"sku"),fnsku:pick(r,"fnsku"),material_code:pick(r,"material code","material-code","material_code"),title:pick(r,"title","product name"),case_size:num(pick(r,"case size","case_size"))||null,channel:pick(r,"channel"),combo_check:pick(r,"combo check","combo_check"),new_mrp:pick(r,"new mrp","new_mrp")?num(pick(r,"new mrp","new_mrp")):null,mapping_version:file.name,raw_data:r};if(!x.fnsku&&!x.sku&&!x.asin){problem(file,i+2,'No ASIN, SKU or FNSKU identifier');continue;}const {error}=await sb.from("product_mappings").insert(x);if(!error)accepted++;else problem(file,i+2,error.message);}await addImport(file,"product_mapping_master",rows.length,accepted);return {accepted};}
  async function importSTR(rows,file){
    const eligible=rows.filter(r=>norm(pick(r,"Transaction Type","transaction-type","transaction type"))==="fc_removal");let accepted=0;
    for(let i=0;i<eligible.length;i++){const r=eligible[i],transaction_id=pick(r,"Transaction Id","transaction-id","transaction id"),sourceSku=pick(r,"Sku","SKU"),asin=pick(r,"Asin","ASIN"),fnsku=pick(r,"FNSKU");if(!transaction_id){problem(file,i+2,'FC_REMOVAL row has no Transaction ID');continue;}
      const match=uniqueMap(cache.mappings,[{field:'sku',label:'SKU',value:sourceSku},{field:'fnsku',label:'FNSKU',value:fnsku},{field:'asin',label:'ASIN',value:asin}].filter(x=>x.value));const mapped=match.row;const sku=sourceSku||mapped?.sku||null;const material_code=pick(r,"Material Code","material-code","material code")||sourceSku||mapped?.sku||mapped?.material_code||null;const invoice=pick(r,"Invoice Number","invoice-number","invoice no"),qty=pick(r,"Quantity","transaction-quantity");
      const source_line_key=[transaction_id,invoice,sourceSku,asin,fnsku,qty,i].map(v=>encodeURIComponent(String(v??''))).join('|');
      const x={source_line_key,transaction_id,invoice_no:invoice,material_code,transaction_date:dateISO(pick(r,"Invoice Date","invoice date","Date","transaction-date")),invoice_date:dateISO(pick(r,"Invoice Date","invoice date")),invoice_value:pick(r,"Invoice Value","invoice value")!==""?num(pick(r,"Invoice Value","invoice value")):null,taxable_value:pick(r,"Taxable Value","taxable value")!==""?num(pick(r,"Taxable Value","taxable value")):null,quantity:num(qty),hsn:pick(r,"Hsn Code","hsn code","hsn"),gst_rate:num(pick(r,"Igst Rate","igst rate","gst-rate"))||null,transaction_type:"FC_REMOVAL",ship_from_fc:pick(r,"Ship From Fc","ship from fc"),ship_from_storage_location:pick(r,"SAP Storage location"),ship_from_plant_code:pick(r,"SAP plant code'"),sku,asin,order_id:pick(r,"Order Id","order id"),receiver_gstin:pick(r,"Gstin Of Receiver","gstin of receiver"),actual_qty:num(pick(r,"ACTUAL QTY","actual qty")),invoice_quantity:num(pick(r,"Invoice Quantity","invoice quantity")),irn_status:pick(r,"Irn Filing Status","irn filing status"),po_number:pick(r,"Po Number","po number"),delivery_no:pick(r,"Delivery No","delivery no"),billing_document:pick(r,"Billing Document","billing document"),migo_no:pick(r,"MIGO No","migo no"),sap_status:pick(r,"Status","status"),mapping_status:match.conflict?'conflict':(sourceSku||mapped)?'matched':'unmapped',raw_data:r};
      const {error}=await sb.from("str_records").upsert(x,{onConflict:"source_line_key"});if(!error)accepted++;else problem(file,i+2,error.message);
    }
    await addImport(file,"str_fc_removal_only",rows.length,accepted);return {eligible:eligible.length,accepted};
  }
  async function importFees(rows,file,cols={}){let accepted=0;for(const r of rows){const order=pick(r,cols.order||"order-id","order-id","removal-order-id","order id"),fee=pick(r,cols.fee||"fee-amount","fee-amount","amount","removal-fee"),currency=pick(r,cols.currency||"currency")||"INR",date=pick(r,cols.date||"posted-date","posted-date","fee-date")||null;if(!order||fee==="")continue;const {data:ord}=await sb.from("removal_orders").select("order_id").eq("order_id",order).limit(1);if(!ord?.length)continue;const {error}=await sb.from("amazon_removal_fees").insert({order_id:order,fee_amount:num(fee),currency,fee_date:date,source_file:file.name,raw_data:r});if(!error)accepted++;}await addImport(file,"amazon_removal_fees_only",rows.length,accepted);return accepted;}
  async function handleAmazonImport(e){e.preventDefault();const fd=new FormData(e.target),file=fd.get("file"),type=fd.get("type");try{importProblems.length=0;importProblemCount=0;const rows=await readReportFile(file);let result={};if(type==="orders")result=await importOrders(rows,file);if(type==="shipments")result=await importShipments(rows,file);if(type==="str")result=await importSTR(rows,file);if(type==="fee")result={accepted:await importFees(rows,file)};if(type==="mapping")result=await importMappings(rows,file);const rejected=importProblemCount;const detail=importProblems.slice(0,8).join('\n');say("importMsg",`Processed ${rows.length} rows from ${file.name}. ${result?.eligible!=null?`FC_REMOVAL rows: ${result.eligible}. `:""}Accepted: ${result?.accepted??0}. Rejected / needs review: ${rejected}.${result?.feeAccepted!=null?` Removal fee lines captured: ${result.feeAccepted}.`:""}${detail?'\n'+detail:''}${rejected>8?'\nMore row errors are available in the browser console.':''}`,rejected===0);if(rejected>8)console.warn('Import row errors',importProblems);await loadAll();}catch(err){say("importMsg",err.message||String(err));}}
  async function handleFeeUpload(e){e.preventDefault();const fd=new FormData(e.target),file=fd.get("file");try{const rows=await readReportFile(file);const accepted=await importFees(rows,file,{order:fd.get("order_col"),fee:fd.get("fee_col"),currency:fd.get("currency_col"),date:fd.get("date_col")});say("feeMsg",`Imported ${accepted} fee rows linked to known Amazon removal orders.`,true);await loadAll();}catch(err){say("feeMsg",err.message);}}

  async function handleReceipt(e){
    e.preventDefault();const lines=[...document.querySelectorAll("#receiptLineEditor .receipt-line")];if(!lines.length){say("receiptMsg","Select an Order ID or Tracking Number first.");return;}
    const receipt_date=e.target.elements.receipt_date.value,remarks=e.target.elements.remarks.value;let saved=0,errors=[];
    for(const line of lines){const s=cache.shipments.find(x=>x.id===line.dataset.shipment);if(!s)continue;const vals={};line.querySelectorAll("[data-qty]").forEach(i=>vals[i.dataset.qty]=num(i.value));const total=Object.values(vals).reduce((a,b)=>a+b,0);if(!total)continue;const already=receiptTotals(s.id),prev=Object.values(already).reduce((a,b)=>a+num(b),0);if(total+prev>num(s.shipped_qty)){errors.push(`${s.order_id} / ${s.tracking_number}: allocation exceeds shipped quantity ${s.shipped_qty}`);continue;}const {error}=await sb.from("warehouse_receipts").insert({shipment_line_id:s.id,...vals,receipt_date,remarks});if(error)errors.push(`${s.order_id} / ${s.tracking_number}: ${error.message}`);else{saved++;await audit("receipt_saved","warehouse_receipt",{shipment_id:s.id,...vals});}}
    if(saved){say("receiptMsg",`${saved} shipment line(s) saved.${errors.length?' Issues: '+errors.join(' | '):''}`,!errors.length);await loadAll();}else say("receiptMsg",errors.join(" | ")||"Enter at least one quantity against a shipment line.");
  }
  async function handleClaim(e){e.preventDefault();const fd=new FormData(e.target),x=Object.fromEntries(fd.entries());x.quantity=num(x.quantity);x.amount=x.amount===""?null:num(x.amount);const {error}=await sb.from("amazon_claims").insert(x);if(error)alert(error.message);else{e.target.reset();await loadAll();}}
  function exportTable(tableId,fileName){const table=$("#"+tableId);if(!table)return;const rows=[...table.querySelectorAll("tr")].filter(tr=>!tr.hidden&&!tr.classList.contains("column-filters")).map(tr=>[...tr.children].map(td=>(td.innerText||"").trim()));if(window.XLSX){const wb=XLSX.utils.book_new(),ws=XLSX.utils.aoa_to_sheet(rows);XLSX.utils.book_append_sheet(wb,ws,"Report");XLSX.writeFile(wb,fileName.replace(/\.csv$/i,".xlsx"));}else{const csv=rows.map(r=>r.map(v=>`"${String(v).replace(/"/g,'""')}"`).join(",")).join("\r\n");const blob=new Blob(["\uFEFF"+csv],{type:"text/csv;charset=utf-8;"}),url=URL.createObjectURL(blob),a=document.createElement("a");a.href=url;a.download=fileName;a.click();URL.revokeObjectURL(url);}}
  function bind(){
    $("#loginBtn").onclick=async()=>{const {error}=await sb.auth.signInWithPassword({email:$("#loginEmail").value,password:$("#loginPassword").value});say("loginMsg",error?.message||"Signing in…",!error);};
    $("#logoutBtn").onclick=()=>sb.auth.signOut();
    $$("#nav [data-page]").forEach(b=>b.onclick=()=>nav(b.dataset.page));$$("[data-goto]").forEach(b=>b.onclick=()=>nav(b.dataset.goto));
    $("#amazonImportForm").onsubmit=handleAmazonImport;$("#receiptForm").onsubmit=handleReceipt;$("#claimForm").onsubmit=handleClaim;
    $("#orderSearch").oninput=renderOrders;$("#shipmentSearch").oninput=renderShipments;$("#exportOrders").onclick=()=>exportTable("ordersTable","amazon_removal_orders.xlsx");$("#exportShipments").onclick=()=>exportTable("shipmentsTable","amazon_shipment_tracker.xlsx");$("#exportReceipts").onclick=()=>exportTable("receiptTable","warehouse_receipt_tracker.xlsx");$("#exportSAP").onclick=()=>exportTable("sapTable","sap_po_tracker.xlsx");$("#monthFilter").onchange=renderAll;$("#fyFilter").onchange=renderAll;
    const month=$("#monthFilter");for(let y=2025;y<=2028;y++)for(let m=1;m<=12;m++){const v=`${y}-${String(m).padStart(2,"0")}`,o=document.createElement("option");o.value=v;o.textContent=new Date(y,m-1,1).toLocaleString("en-IN",{month:"short",year:"numeric"});month.append(o);}
    $("#receiptForm").elements.receipt_date.value=today();$("#receiptLookupType").addEventListener("change",renderReceiptSelector);$("#receiptLookupValue").addEventListener("change",renderReceiptLineEditor);
  }
  document.addEventListener("DOMContentLoaded",()=>{bind();init();});
})();