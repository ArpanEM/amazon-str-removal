from pathlib import Path
p=Path('/mnt/data/strtower_work/index.html'); s=p.read_text()
s=s.replace('<label class="wide">Select by Amazon Order ID or Tracking Number (then choose the SKU line)<select name="shipment_id" id="receiptShipmentSelect" required><option value="">Import shipment details first</option></select></label>\n<label class="wide">Selected shipment details<div id="receiptShipmentInfo" class="notice">Select a shipment line to view its product and shipped quantity.</div></label><label>Good qty<input name="good_qty" type="number" min="0" value="0"></label><label>Damaged qty<input name="damaged_qty" type="number" min="0" value="0"></label><label>Near expiry qty<input name="near_expiry_qty" type="number" min="0" value="0"></label><label>Expired qty<input name="expired_qty" type="number" min="0" value="0"></label><label>Short qty<input name="short_qty" type="number" min="0" value="0"></label><label>Not received qty<input name="not_received_qty" type="number" min="0" value="0"></label><label>Receipt date<input name="receipt_date" type="date" required></label><label class="wide">Remarks<input name="remarks" placeholder="Optional notes"></label><button class="primary">Save receipt</button>', '<label>Find shipment by<select id="receiptLookupType"><option value="order">Amazon Order ID</option><option value="tracking">Tracking Number</option></select></label><label class="wide">Select Order ID / Tracking Number<select id="receiptLookupValue" required><option value="">Import shipment details first</option></select></label><label class="wide">Shipment lines and physical receipt allocation<div id="receiptLineEditor" class="receipt-line-editor"><div class="notice">Select an Order ID or Tracking Number to load all associated material lines and shipped quantities.</div></div></label><label>Receipt date<input name="receipt_date" type="date" required></label><label class="wide">Remarks<input name="remarks" placeholder="Optional notes for all selected lines"></label><button class="primary">Save all receipt lines</button>')
p.write_text(s)

p=Path('/mnt/data/strtower_work/app.js'); s=p.read_text()
s=s.replace('2026-11-12-final-sap-sku','2026-11-12-filter-sort-multiline-receipts')
start=s.index('  function renderOrders(){')
end=s.index('  function emptyRow(n)',start)
new='''  function renderOrders(){
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
    if(!filterRow){filterRow=document.createElement("tr");filterRow.className="column-filters";[...head.cells].forEach((th,i)=>{const cell=document.createElement("th"),input=document.createElement("input");input.className="column-filter";input.placeholder="Filter…";input.setAttribute("aria-label",`Filter ${th.textContent}`);input.value=state.filters[i]||"";input.oninput=()=>{state.filters[i]=input.value;applyTableFilters(tbodyId);};cell.append(input);filterRow.append(cell);th.title="Click to sort";th.style.cursor="pointer";th.onclick=()=>{if(state.sortCol===i)state.asc=!state.asc;else{state.sortCol=i;state.asc=true;}sortTableRows(tbodyId);};});table.tHead.append(filterRow);}
    applyTableFilters(tbodyId);
  }
  function applyTableFilters(id){const body=$("#"+id),state=tableState[id];if(!body||!state)return;[...body.rows].forEach(row=>{const ok=Object.entries(state.filters).every(([i,v])=>!v||norm(row.cells[Number(i)]?.textContent).includes(norm(v)));row.hidden=!ok;});}
  function sortTableRows(id){const body=$("#"+id),state=tableState[id];if(!body||!state||state.sortCol<0)return;const rows=[...body.rows];rows.sort((a,b)=>{const x=a.cells[state.sortCol]?.textContent.trim()||"",y=b.cells[state.sortCol]?.textContent.trim()||"",nx=Number(x.replace(/,/g,"")),ny=Number(y.replace(/,/g,""));const cmp=x!==""&&y!==""&&Number.isFinite(nx)&&Number.isFinite(ny)?nx-ny:x.localeCompare(y,undefined,{numeric:true,sensitivity:"base"});return state.asc?cmp:-cmp;});rows.forEach(r=>body.append(r));applyTableFilters(id);}
'''
s=s[:start]+new+s[end:]
# replace receipt selector/show/render and handler
start=s.index('  function renderReceiptSelector(){')
end=s.index('  function renderFees()',start)
new='''  function renderReceiptSelector(){
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
'''
s=s[:start]+new+s[end:]
# replace handleReceipt
start=s.index('  async function handleReceipt(e){')
end=s.index('  async function handleClaim(e)',start)
new='''  async function handleReceipt(e){
    e.preventDefault();const lines=[...document.querySelectorAll("#receiptLineEditor .receipt-line")];if(!lines.length){say("receiptMsg","Select an Order ID or Tracking Number first.");return;}
    const receipt_date=e.target.elements.receipt_date.value,remarks=e.target.elements.remarks.value;let saved=0,errors=[];
    for(const line of lines){const s=cache.shipments.find(x=>x.id===line.dataset.shipment);if(!s)continue;const vals={};line.querySelectorAll("[data-qty]").forEach(i=>vals[i.dataset.qty]=num(i.value));const total=Object.values(vals).reduce((a,b)=>a+b,0);if(!total)continue;const already=receiptTotals(s.id),prev=Object.values(already).reduce((a,b)=>a+num(b),0);if(total+prev>num(s.shipped_qty)){errors.push(`${s.order_id} / ${s.tracking_number}: allocation exceeds shipped quantity ${s.shipped_qty}`);continue;}const {error}=await sb.from("warehouse_receipts").insert({shipment_line_id:s.id,...vals,receipt_date,remarks});if(error)errors.push(`${s.order_id} / ${s.tracking_number}: ${error.message}`);else{saved++;await audit("receipt_saved","warehouse_receipt",{shipment_id:s.id,...vals});}}
    if(saved){say("receiptMsg",`${saved} shipment line(s) saved.${errors.length?' Issues: '+errors.join(' | '):''}`,!errors.length);await loadAll();}else say("receiptMsg",errors.join(" | ")||"Enter at least one quantity against a shipment line.");
  }
'''
s=s[:start]+new+s[end:]
# bind event changes
s=s.replace('$("#receiptForm").elements.receipt_date.value=today();$("#receiptShipmentSelect").addEventListener("change",showReceiptShipment);','$("#receiptForm").elements.receipt_date.value=today();$("#receiptLookupType").addEventListener("change",renderReceiptSelector);$("#receiptLookupValue").addEventListener("change",renderReceiptLineEditor);')
p.write_text(s)
