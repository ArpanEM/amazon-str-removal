/* Amazon STR & Removal Control Tower
   Static frontend for GitHub Pages + Supabase. All privileged access is enforced by RLS.
   Configure config.js before use. CSV parser is intentionally dependency-free. */
(() => {
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
  const pick=(o,...keys)=>{for(const k of keys){const found=Object.keys(o||{}).find(x=>x.toLowerCase()===k.toLowerCase());if(found)return o[found];}return "";};
  const num=v=>{const n=Number(String(v??"").replace(/,/g,""));return Number.isFinite(n)?n:0;};
  const dateISO=v=>{const x=String(v??"").trim();if(!x)return null;if(/^\d{4}-\d{2}-\d{2}/.test(x))return x.slice(0,10);const m=x.match(/^(\d{1,2})[./-](\d{1,2})[./-](\d{4})$/);if(m)return `${m[3]}-${m[2].padStart(2,"0")}-${m[1].padStart(2,"0")}`;return x;};
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
    $("#loginView").classList.add("hidden");$("#appView").classList.remove("hidden");await loadAll();
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
  function renderOrders(){const q=norm($("#orderSearch")?.value);$("#ordersTable").innerHTML=cache.orders.filter(x=>[x.order_id,x.sku,x.fnsku].some(v=>norm(v).includes(q))).slice(0,500).map(x=>`<tr><td>${esc(x.order_id)}</td><td>${esc(x.sku)}</td><td>${esc(x.fnsku)}</td><td>${fmt(x.requested_qty)}</td><td>${fmt(x.cancelled_qty)}</td><td>${fmt(x.shipped_qty)}</td><td>${fmt(x.in_process_qty)}</td><td><span class="tag">${esc(x.status||"—")}</span></td></tr>`).join("")||emptyRow(8);}
  function renderShipments(){
    const q=norm($("#shipmentSearch")?.value);
    const rows=cache.shipments.filter(x=>[x.order_id,x.tracking_number,x.fnsku,x.id].some(v=>norm(v).includes(q))).slice(0,500);
    $("#shipmentsTable").innerHTML=rows.map(x=>{
      const y=enrichShipment(x);
      return "<tr><td>"+esc(x.id.slice(0,8))+"</td><td>"+esc(x.order_id)+"</td><td>"+esc(x.tracking_number)+"</td><td>"+esc(y.sku||"—")+"</td><td>"+esc(x.fnsku||"—")+"</td><td>"+esc(y.material_code||"—")+"</td><td>"+fmt(x.shipped_qty)+"</td><td>"+esc(x.amazon_status||"—")+"</td><td>"+esc(x.shipment_date||"—")+"</td></tr>";
    }).join("")||emptyRow(9);
  }
  function emptyRow(n){return `<tr><td colspan="${n}" class="muted">No records found.</td></tr>`;}
  function enrichShipment(s){
    const order=cache.orders.find(o=>o.order_id===s.order_id && ((!s.fnsku||o.fnsku===s.fnsku)||(!s.sku||o.sku===s.sku)));
    const maps=cache.mappings.filter(m=>(s.fnsku&&norm(m.fnsku)===norm(s.fnsku))||(!s.fnsku&&s.sku&&norm(m.sku)===norm(s.sku)));
    const unique=[...new Map(maps.map(m=>[`${norm(m.sku)}|${norm(m.material_code)}`,m])).values()];
    let sku=s.sku||order?.sku||"", material=s.material_code||"";
    let mappingStatus=s.mapping_status||"unmapped";
    if(unique.length===1){sku=sku||unique[0].sku||"";material=material||unique[0].material_code||"";mappingStatus="matched";}
    else if(unique.length>1){mappingStatus="conflict";}
    else if(sku||material||order?.sku){sku=sku||order?.sku||"";material=material||order?.material_code||"";mappingStatus=material||sku?"order-detail":"partial";}
    return {...s,sku,material_code:material,mapping_status:mappingStatus};
  }
  function renderReceiptSelector(){
    const sel=$("#receiptShipmentSelect");if(!sel)return;
    const current=sel.value;
    sel.innerHTML='<option value="">Select shipment line</option>'+cache.shipments.map(raw=>{const s=enrichShipment(raw);return `<option value="${esc(s.id)}">${esc(s.order_id)} | ${esc(s.sku||"SKU missing")} | ${esc(s.fnsku||"FNSKU missing")} | ${esc(s.tracking_number)} | Qty ${fmt(s.shipped_qty)}</option>`;}).join("");
    if(current&&cache.shipments.some(s=>s.id===current))sel.value=current;
    showReceiptShipment();
  }
  function showReceiptShipment(){
    const id=$("#receiptShipmentSelect")?.value,s=cache.shipments.find(x=>x.id===id),el=$("#receiptShipmentInfo");if(!el)return;
    if(!s){el.textContent="Select a shipment line to view its product and shipped quantity.";return;}
    const x=enrichShipment(s);
    el.innerHTML=`<b>Order:</b> ${esc(x.order_id)} &nbsp; <b>SKU:</b> ${esc(x.sku||"Not mapped")} &nbsp; <b>FNSKU:</b> ${esc(x.fnsku||"—")} &nbsp; <b>Material:</b> ${esc(x.material_code||"Not mapped")}<br><b>Tracking:</b> ${esc(x.tracking_number)} &nbsp; <b>Shipped:</b> ${fmt(x.shipped_qty)} &nbsp; <b>Mapping:</b> ${esc(x.mapping_status)}`;
  }
  function renderReceipts(){
    renderReceiptSelector();
    $("#receiptTable").innerHTML=cache.shipments.slice(0,1000).map(raw=>{const s=enrichShipment(raw),t=receiptTotals(s.id),physical=t.good_qty+t.damaged_qty+t.near_expiry_qty+t.expired_qty,alloc=physical+t.short_qty+t.not_received_qty,pending=Math.max(0,num(s.shipped_qty)-alloc);return `<tr><td>${esc(s.order_id)}</td><td>${esc(s.sku||"—")}</td><td>${esc(s.fnsku||"—")}</td><td>${esc(s.material_code||"—")}</td><td>${esc(s.tracking_number)}</td><td>${fmt(s.shipped_qty)}</td><td>${fmt(t.good_qty)}</td><td>${fmt(t.damaged_qty)}</td><td>${fmt(t.near_expiry_qty)}</td><td>${fmt(t.expired_qty)}</td><td>${fmt(physical)}</td><td>${fmt(t.short_qty)}</td><td>${fmt(t.not_received_qty)}</td><td><b>${fmt(pending)}</b></td><td><span class="tag">${esc(s.mapping_status)}</span></td></tr>`;}).join("")||emptyRow(15);
  }
  function renderFees(){const fees=cache.fees;const currencies=[...new Set(fees.map(x=>x.currency).filter(Boolean))], total=fees.reduce((a,x)=>a+num(x.fee_amount),0),qty=fees.reduce((a,f)=>{const o=cache.orders.find(x=>x.order_id===f.order_id&&((f.fnsku&&x.fnsku===f.fnsku)||(f.sku&&x.sku===f.sku)));return a+num(o?.shipped_qty);},0);$("#feeTotal").textContent=currencies.length<=1?money(total,currencies[0]||"INR"):"Multiple currencies";$("#feeCount").textContent=fmt(fees.length);$("#feePerUnit").textContent=currencies.length===1&&qty?money(total/qty,currencies[0]):"—";$("#feesTable").innerHTML=fees.slice(0,500).map(f=>`<tr><td>${esc(f.order_id)}</td><td>${esc(f.sku||"—")}</td><td>${esc(f.fnsku||"—")}</td><td>${money(f.fee_amount,f.currency||"INR")}</td><td>${esc(f.currency||"—")}</td><td>${esc(f.fee_date||"—")}</td><td>${esc(f.source_file||"Amazon report")}</td></tr>`).join("")||emptyRow(7);}
  function renderClaims(){$("#claimsTable").innerHTML=cache.claims.map(c=>`<tr><td>${esc(c.order_id)}</td><td>${esc(c.claim_ref||"—")}</td><td>${esc(c.claim_type)}</td><td>${fmt(c.quantity)}</td><td>${esc(c.status)}</td><td>${c.amount==null?"—":money(c.amount)}</td></tr>`).join("")||emptyRow(6);}
  function renderSAP(){
    const records=cache.str||[],groups=new Map();
    records.forEach(r=>{const status=norm(r.irn_status),accepted=status.includes("govt_accepted")||status.includes("govt accepted"),rejected=status.includes("govt_rejected")||status.includes("govt rejected"),blank=!status,ship=(r.ship_from_fc||"Unspecified Ship From").trim();const consolidatable=rejected||blank;const kind=accepted?"GOVT_ACCEPTED":(consolidatable?(r.ship_from_fc?"REJECTED + BLANK":"LOCATION MISSING") : "IRN STATUS REVIEW");const key=accepted?`ACCEPTED|${r.transaction_id}`:(consolidatable&&r.ship_from_fc?`CONSOLIDATED|${ship}`:`REVIEW|${r.transaction_id}`);if(!groups.has(key))groups.set(key,{key,ship,kind,rows:[],skus:new Set(),qty:0,po:new Set(),delivery:new Set(),billing:new Set(),migo:new Set(),statuses:new Set()});const g=groups.get(key);g.rows.push(r);if(r.sku)g.skus.add(r.sku);g.qty+=num(r.quantity);if(r.po_number)g.po.add(r.po_number);if(r.delivery_no)g.delivery.add(r.delivery_no);if(r.billing_document)g.billing.add(r.billing_document);if(r.migo_no)g.migo.add(r.migo_no);if(r.sap_status)g.statuses.add(r.sap_status);});
    const list=[...groups.values()].sort((a,b)=>a.ship.localeCompare(b.ship)||a.kind.localeCompare(b.kind));
    const acceptedCount=list.filter(g=>g.kind==="GOVT_ACCEPTED").length,consolidatedCount=list.filter(g=>g.kind==="REJECTED + BLANK").length;
    const kpi=document.querySelector("#sapSummary");if(kpi)kpi.innerHTML=`<div class="kpi-grid compact"><article class="kpi"><span>STR lines</span><b>${fmt(records.length)}</b></article><article class="kpi"><span>Individual PO groups</span><b>${fmt(acceptedCount)}</b></article><article class="kpi"><span>Consolidated PO groups</span><b>${fmt(consolidatedCount)}</b></article><article class="kpi"><span>PO pending</span><b>${fmt(list.filter(g=>!g.po.size).length)}</b></article></div>`;
    $("#sapTable").innerHTML=list.map(g=>`<tr><td>${esc(g.kind==="GOVT_ACCEPTED"?g.rows[0].transaction_id:"Consolidated – "+g.ship)}</td><td>${esc(g.ship)}</td><td>${esc(g.kind)}</td><td>${fmt(g.rows.length)}</td><td>${fmt(g.skus.size)}</td><td>${fmt(g.qty)}</td><td>${esc([...g.po].join(", ")||"Pending")}</td><td>${esc([...g.delivery].join(", ")||"—")}</td><td>${esc([...g.billing].join(", ")||"—")}</td><td>${esc([...g.migo].join(", ")||"—")}</td></tr>`).join("")||emptyRow(10);
    $("#sapExceptions").innerHTML=records.filter(r=>!r.ship_from_fc||(!["", "govt_accepted", "govt_rejected", "govt accepted", "govt rejected"].includes(norm(r.irn_status)))||(!r.po_number&&!r.delivery_no&&!r.billing_document&&!r.migo_no)).slice(0,500).map(r=>`<tr><td>Amazon STR</td><td>${esc(r.transaction_id)}</td><td>${esc(r.sku||r.material_code||"—")}</td><td>${esc(!r.ship_from_fc?"Missing Ship From FC":!r.irn_status?"Blank IRN status":"SAP documents pending")}</td><td>${esc(r.irn_status||"Blank")}</td><td>Review</td></tr>`).join("")||emptyRow(6);
  }
  function renderImports(){$("#importsTable").innerHTML=cache.imports.map(x=>`<tr><td>${esc(x.file_name)}</td><td>${esc(x.report_type)}</td><td>${fmt(x.rows_read)}</td><td>${fmt(x.rows_accepted)}</td><td>${esc(x.imported_by_email||"—")}</td><td>${esc(x.created_at?.slice(0,10)||"—")}</td></tr>`).join("")||emptyRow(6);}
  async function renderEmployees(){if(!["admin","manager"].includes(profile?.role))return;const {data}=await sb.from("profiles").select("id,email,full_name,role").order("created_at");$("#employeesTable").innerHTML=(data||[]).map(p=>`<tr><td>${esc(p.full_name)}</td><td>${esc(p.email)}</td><td>${esc(p.role)}</td><td>${p.id===session.user.id?"Current user":`<select data-role-user="${p.id}"><option ${p.role==="employee"?"selected":""}>employee</option><option ${p.role==="manager"?"selected":""}>manager</option><option ${p.role==="admin"?"selected":""}>admin</option></select>`}</td></tr>`).join("")||emptyRow(4);$$("[data-role-user]").forEach(el=>el.addEventListener("change",async()=>{const {error}=await sb.from("profiles").update({role:el.value}).eq("id",el.dataset.roleUser);if(error)alert(error.message);else{await audit("role_updated","profile",{user_id:el.dataset.roleUser,role:el.value});await renderEmployees();}}));}
  async function renderAudit(){if(!["admin","manager"].includes(profile?.role))return;const {data}=await sb.from("audit_logs").select("*,profiles(full_name,email)").order("created_at",{ascending:false}).limit(300);$("#auditTable").innerHTML=(data||[]).map(a=>`<tr><td>${esc(a.created_at?.replace("T"," ").slice(0,19))}</td><td>${esc(a.profiles?.full_name||a.profiles?.email||"—")}</td><td>${esc(a.action)}</td><td>${esc(a.entity)}</td><td>${esc(JSON.stringify(a.details||{}))}</td></tr>`).join("")||emptyRow(5);}
  function nav(page){$$(".page").forEach(p=>p.classList.add("hidden"));$("#"+page+"Page").classList.remove("hidden");$$(".nav-item").forEach(n=>n.classList.toggle("active",n.dataset.page===page));$("#pageTitle").textContent=({dashboard:"Dashboard",orders:"Removal Orders",shipments:"Shipment Tracker",receipts:"Warehouse Receipts",fees:"Amazon Removal Fees",sap:"PO & SAP Tracker",claims:"Amazon Claims",reports:"Reports & Imports",employees:"Employee Management",audit:"Audit Logs"})[page]||"Dashboard";}
  async function addImport(file,type,read,accepted){const {error}=await sb.from("import_batches").insert({file_name:file.name,report_type:type,rows_read:read,rows_accepted:accepted,imported_by:session.user.id,imported_by_email:session.user.email});if(error)throw error;await audit("report_imported","import_batch",{file:file.name,type,rows_read:read,rows_accepted:accepted});}
  async function importOrders(rows,file){let accepted=0,feeAccepted=0;for(const r of rows){const x={order_id:pick(r,"order-id","order id","removal-order-id"),sku:pick(r,"sku"),fnsku:pick(r,"fnsku"),material_code:pick(r,"material-code","material code","material"),requested_qty:num(pick(r,"requested-quantity","requested qty")),cancelled_qty:num(pick(r,"cancelled-quantity","cancelled qty")),shipped_qty:num(pick(r,"shipped-quantity","shipped qty")),in_process_qty:num(pick(r,"in-process-quantity","in-process qty")),status:pick(r,"status","order-status"),request_date:(pick(r,"request-date","created-date")||"").slice(0,10)||null,raw_data:r};if(!x.order_id)continue;const {error}=await sb.from("removal_orders").upsert(x,{onConflict:"order_id,sku,fnsku"});if(!error)accepted++;
      // Amazon's Removal Order Detail report is the source of truth for removal fees.
      const rawFee=pick(r,"removal-fee","removal fee");const currency=pick(r,"currency");
      if(String(rawFee).trim()!==""&&currency){const feeRow={order_id:x.order_id,sku:x.sku||"",fnsku:x.fnsku||"",fee_amount:num(rawFee),currency,fee_date:x.request_date,source_file:file.name,raw_data:r};
        const {error:feeError}=await sb.from("amazon_removal_fees").upsert(feeRow,{onConflict:"order_id,sku,fnsku,source_file"});if(!feeError)feeAccepted++;}
    }await addImport(file,"removal_order_detail",rows.length,accepted);return {accepted,feeAccepted};}
  async function importShipments(rows,file){let accepted=0;for(const r of rows){const order_id=pick(r,"order-id","order id"),tracking_number=pick(r,"tracking-number","tracking number"),fnsku=pick(r,"fnsku"),sku=pick(r,"sku");if(!order_id||!tracking_number)continue;const order=cache.orders.find(o=>o.order_id===order_id&&((fnsku&&o.fnsku===fnsku)||(sku&&o.sku===sku)));const maps=cache.mappings.filter(m=>(fnsku&&norm(m.fnsku)===norm(fnsku))||(!fnsku&&sku&&norm(m.sku)===norm(sku)));const mapKeys=[...new Set(maps.map(m=>`${norm(m.sku)}|${norm(m.material_code)}`))];let mapped=mapKeys.length===1?maps[0]:null;const x={order_id,tracking_number,fnsku,sku:sku||order?.sku||mapped?.sku||null,material_code:order?.material_code||mapped?.material_code||null,mapping_status:mapKeys.length>1?"conflict":(sku||order?.sku||mapped?.sku)?"matched":"unmapped",shipped_qty:num(pick(r,"shipped-quantity","quantity-shipped","quantity")),amazon_status:pick(r,"status","shipment-status"),shipment_date:pick(r,"shipment-date","shipped-date")||null,disposition:pick(r,"disposition"),raw_data:r};const {error}=await sb.from("shipment_lines").insert(x);if(!error)accepted++;}await addImport(file,"removal_shipment_detail",rows.length,accepted);}
  async function importMappings(rows,file){let accepted=0;for(const r of rows){const x={asin:pick(r,"asin","asin.1"),sku:pick(r,"sku"),fnsku:pick(r,"fnsku"),material_code:pick(r,"material code","material-code","material_code"),title:pick(r,"title","product name"),case_size:num(pick(r,"case size","case_size"))||null,channel:pick(r,"channel"),combo_check:pick(r,"combo check","combo_check"),new_mrp:pick(r,"new mrp","new_mrp")?num(pick(r,"new mrp","new_mrp")):null,mapping_version:file.name,raw_data:r};if(!x.fnsku&&!x.sku&&!x.asin)continue;const {error}=await sb.from("product_mappings").insert(x);if(!error)accepted++;}await addImport(file,"product_mapping_master",rows.length,accepted);return accepted;}
  async function importSTR(rows,file){
    const eligible=rows.filter(r=>norm(pick(r,"Transaction Type","transaction-type","transaction type"))==="fc_removal");let accepted=0;
    for(const r of eligible){
      const transaction_id=pick(r,"Transaction Id","transaction-id","transaction id"),sku=pick(r,"Sku","SKU"),asin=pick(r,"Asin","ASIN");
      if(!transaction_id||!sku)continue;
      const mapRows=cache.mappings.filter(m=>(sku&&norm(m.sku)===norm(sku))||(asin&&norm(m.asin)===norm(asin)));const mappedCodes=[...new Set(mapRows.map(m=>m.material_code).filter(Boolean))];const material_code=pick(r,"Material Code","material-code","material code")||(mappedCodes.length===1?mappedCodes[0]:null);
      const x={
        transaction_id,invoice_no:pick(r,"Invoice Number","invoice-number","invoice no"),material_code,
        transaction_date:dateISO(pick(r,"Invoice Date","invoice date","Date","transaction-date")),invoice_date:dateISO(pick(r,"Invoice Date","invoice date")),invoice_value:pick(r,"Invoice Value","invoice value")!==""?num(pick(r,"Invoice Value","invoice value")):null,taxable_value:pick(r,"Taxable Value","taxable value")!==""?num(pick(r,"Taxable Value","taxable value")):null,
        quantity:num(pick(r,"Quantity","transaction-quantity")),hsn:pick(r,"Hsn Code","hsn code","hsn"),gst_rate:num(pick(r,"Igst Rate","igst rate","gst-rate"))||null,
        transaction_type:"FC_REMOVAL",ship_from_fc:pick(r,"Ship From Fc","ship from fc"),ship_from_storage_location:pick(r,"SAP Storage location"),ship_from_plant_code:pick(r,"SAP plant code'"),
        sku,asin,order_id:pick(r,"Order Id","order id"),receiver_gstin:pick(r,"Gstin Of Receiver","gstin of receiver"),
        actual_qty:num(pick(r,"ACTUAL QTY","actual qty")),invoice_quantity:num(pick(r,"Invoice Quantity","invoice quantity")),irn_status:pick(r,"Irn Filing Status","irn filing status"),
        po_number:pick(r,"Po Number","po number"),delivery_no:pick(r,"Delivery No","delivery no"),billing_document:pick(r,"Billing Document","billing document"),migo_no:pick(r,"MIGO No","migo no"),sap_status:pick(r,"Status","status"),raw_data:r
      };
      const {error}=await sb.from("str_records").upsert(x,{onConflict:"transaction_id,sku"});if(!error)accepted++;
    }
    await addImport(file,"str_fc_removal_only",rows.length,accepted);return {eligible:eligible.length,accepted};
  }
  async function importFees(rows,file,cols={}){let accepted=0;for(const r of rows){const order=pick(r,cols.order||"order-id","order-id","removal-order-id","order id"),fee=pick(r,cols.fee||"fee-amount","fee-amount","amount","removal-fee"),currency=pick(r,cols.currency||"currency")||"INR",date=pick(r,cols.date||"posted-date","posted-date","fee-date")||null;if(!order||fee==="")continue;const {data:ord}=await sb.from("removal_orders").select("order_id").eq("order_id",order).limit(1);if(!ord?.length)continue;const {error}=await sb.from("amazon_removal_fees").insert({order_id:order,fee_amount:num(fee),currency,fee_date:date,source_file:file.name,raw_data:r});if(!error)accepted++;}await addImport(file,"amazon_removal_fees_only",rows.length,accepted);return accepted;}
  async function handleAmazonImport(e){e.preventDefault();const fd=new FormData(e.target),file=fd.get("file"),type=fd.get("type");try{const rows=parseCSV(await file.text());let result;if(type==="orders")result=await importOrders(rows,file);if(type==="shipments")await importShipments(rows,file);if(type==="str")result=await importSTR(rows,file);if(type==="fee")result={accepted:await importFees(rows,file)};if(type==="mapping")result={accepted:await importMappings(rows,file)};say("importMsg",`Processed ${rows.length} rows. ${result?.eligible!=null?`FC_REMOVAL rows: ${result.eligible}. `:""}Accepted: ${result?.accepted??"see import log"}.${result?.feeAccepted!=null?` Removal fee lines captured: ${result.feeAccepted}.`:""}`,true);await loadAll();}catch(err){say("importMsg",err.message||String(err));}}
  async function handleFeeUpload(e){e.preventDefault();const fd=new FormData(e.target),file=fd.get("file"),rows=parseCSV(await file.text());try{const accepted=await importFees(rows,file,{order:fd.get("order_col"),fee:fd.get("fee_col"),currency:fd.get("currency_col"),date:fd.get("date_col")});say("feeMsg",`Imported ${accepted} fee rows linked to known Amazon removal orders.`,true);await loadAll();}catch(err){say("feeMsg",err.message);}}

  async function handleReceipt(e){e.preventDefault();const fd=new FormData(e.target),id=fd.get("shipment_id");const s=cache.shipments.find(x=>x.id===id);if(!s){say("receiptMsg","Shipment not found. Use the internal shipment ID shown in the tracker.");return;}const vals={};for(const k of ["good_qty","damaged_qty","near_expiry_qty","expired_qty","short_qty","not_received_qty"])vals[k]=num(fd.get(k));const total=Object.values(vals).reduce((a,b)=>a+b,0),already=receiptTotals(s.id),prev=Object.values(already).reduce((a,b)=>a+b,0);if(total+prev>num(s.shipped_qty)){say("receiptMsg",`Allocation exceeds shipped quantity (${s.shipped_qty}). Existing final allocation: ${prev}.`);return;}const {error}=await sb.from("warehouse_receipts").insert({shipment_line_id:s.id,...vals,receipt_date:fd.get("receipt_date"),remarks:fd.get("remarks")});if(error)say("receiptMsg",error.message);else{say("receiptMsg","Receipt allocation saved.",true);e.target.reset();await audit("receipt_saved","warehouse_receipt",{shipment_id:s.id,...vals});await loadAll();}}
  async function handleClaim(e){e.preventDefault();const fd=new FormData(e.target),x=Object.fromEntries(fd.entries());x.quantity=num(x.quantity);x.amount=x.amount===""?null:num(x.amount);const {error}=await sb.from("amazon_claims").insert(x);if(error)alert(error.message);else{e.target.reset();await loadAll();}}
  function bind(){
    $("#loginBtn").onclick=async()=>{const {error}=await sb.auth.signInWithPassword({email:$("#loginEmail").value,password:$("#loginPassword").value});say("loginMsg",error?.message||"Signing in…",!error);};
    $("#logoutBtn").onclick=()=>sb.auth.signOut();
    $$("#nav [data-page]").forEach(b=>b.onclick=()=>nav(b.dataset.page));$$("[data-goto]").forEach(b=>b.onclick=()=>nav(b.dataset.goto));
    $("#amazonImportForm").onsubmit=handleAmazonImport;$("#receiptForm").onsubmit=handleReceipt;$("#claimForm").onsubmit=handleClaim;
    $("#orderSearch").oninput=renderOrders;$("#shipmentSearch").oninput=renderShipments;$("#monthFilter").onchange=renderAll;$("#fyFilter").onchange=renderAll;
    const month=$("#monthFilter");for(let y=2025;y<=2028;y++)for(let m=1;m<=12;m++){const v=`${y}-${String(m).padStart(2,"0")}`,o=document.createElement("option");o.value=v;o.textContent=new Date(y,m-1,1).toLocaleString("en-IN",{month:"short",year:"numeric"});month.append(o);}
    $("#receiptForm").elements.receipt_date.value=today();$("#receiptShipmentSelect").addEventListener("change",showReceiptShipment);
  }
  document.addEventListener("DOMContentLoaded",()=>{bind();init();});
})();