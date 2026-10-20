/**
 * CONSTRUCTION COST INTELLIGENCE 4.5 — BILLING ACCOUNT
 * Install as the ONLY invoice-processing code in the billing-account Apps Script project.
 * Existing Script Properties: OPENAI_API_KEY, SPREADSHEET_ID, BILLING_GMAIL_ADDRESS.
 * Google Sheet is initialized by the personal-account script. No deploy required.
 * The script NEVER extracts unlabeled/unverified-job mail.
 */
const C45 = {
  version: '4.5.3', endpoint: 'https://api.openai.com/v1/responses',
  headers: {
    'Price History':['Rate_Key','Cost_Code','Normalized_Item','Cost_Type','Unit','Brand_Spec','Samples','Low','Median','Average','High','Latest','Latest_Date','Recommended','Auto_Use','Source_Documents','Suppliers','Latest_Supplier','Notes'],
    'Estimate':['Project','Cost_Code','Description','Cost_Type','Qty','Unit','Rate_Key','Suggested_Unit_Cost','Override_Unit_Cost','Used_Unit_Cost','Builder_Cost','Markup_Pct','Client_Total','Evidence'],
    'Documents':['Document_ID','Source_Key','Thread_ID','Message_ID','Vendor','Document_No','Doc_Type','Document_Date','Project_ID','Project_Name','Job_Address','Subtotal','Tax','Freight','Discount','Total','Cost_Impact','Status','Confidence','Transaction_ID','Related_Doc_ID','Refs','Scope_Summary','Source_Link','Notes','Recorded_At'],
    'Cost Lines':['Line_ID','Document_ID','Project_ID','Cost_Code','Normalized_Item','Source_Description','Cost_Type','Qty','Unit','Line_Cost','Unit_Cost','Rate_Eligible','Confidence','Brand_Spec','SKU','Manual_Qty','Manual_Unit','Scope_Notes'],
    'Scope Notes':['Scope_ID','Document_ID','Detail','Qty','Unit','Amount','Evidence'],
    'Projects':['Project_ID','Project_Name','Gmail_Label','Job_Address','Aliases','Verified','Last_Seen'],
    'Review Queue':['Review_ID','Document_ID','Issue_Type','Details','Suggested_Action','Status','Created_At'],
    'Cost Codes':['Category','Cost_Code','Aliases'],
    'Label Rules':['Gmail_Label','Role','Project_ID','Verified','Note'],
    'Settings':['Key','Value','Description'],
    'Source Ledger':['Source_Key','Thread_ID','Message_ID','Filename','Attachment_Hash','Document_ID','State','Attempts','Updated_At','Error'],
    'Run Log':['At','Mode','Threads_Checked','Sources_Processed','Docs_Recorded','Reviews','Duration_s','Note']
  },
  ignoreLabels: ['inbox','sent','drafts','important','starred','spam','trash','invoices','paid','unpaid','billing','vendors','vendor','statements','statement','quotes','quote','receipts','receipt','supplier','suppliers','personal','accounting','processed','review'],
  supportTypes: ['Delivery Confirmation','Packing Slip','Order Confirmation','Purchase Order','Statement','Payment Confirmation','Other'],
  costTypes: ['Invoice','Receipt','Credit']
};

// ---------- Sheet I/O: one read per sheet per run, batched appended rows ----------
function c45Book_(){const id=PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID');if(!id)throw Error('Missing SPREADSHEET_ID property');return SpreadsheetApp.openById(id);}
function c45View_(ss,name){const sh=ss.getSheetByName(name);if(!sh)throw Error('Missing '+name+'. Complete v4.5 sheet initialization first.');const h=C45.headers[name];const actual=sh.getRange(1,1,1,h.length).getValues()[0].map(String);if(h.some((x,i)=>x!==actual[i]))throw Error('Wrong v4.5 headers: '+name);const v=sh.getLastRow()>1?sh.getRange(2,1,sh.getLastRow()-1,h.length).getValues():[];const rows=v.map((values,i)=>{const r={_row:i+2};h.forEach((k,j)=>r[k]=values[j]);return r;}).filter(r=>h.some(k=>r[k]!==''&&r[k]!==null));return {sh,h,rows};}
function c45Append_(v,rows){if(!rows.length)return [];const first=v.sh.getLastRow()+1;v.sh.getRange(first,1,rows.length,v.h.length).setValues(rows.map(r=>v.h.map(k=>r[k]===undefined?'':r[k])));const arr=rows.map((r,i)=>Object.assign({_row:first+i},r));v.rows.push(...arr);return arr;}
// Patch only the requested cells. Never round-trip a whole row with getValues():
// getValues() resolves =HYPERLINK() to its visible label and destroys the URL on rewrite.
function c45Patch_(v,row,patch){
  const updates=Object.keys(patch).map(k=>{
    const i=v.h.indexOf(k);
    if(i<0)throw Error('Column '+k+' missing');
    return {key:k,column:i+1,value:patch[k]};
  });
  updates.forEach(x=>{
    v.sh.getRange(row._row,x.column).setValue(x.value);
    row[x.key]=x.value;
  });
}
function c45Set_(settings,key,value){const row=settings.rows.find(r=>r.Key===key);if(!row)throw Error('Missing setting '+key);c45Patch_(settings,row,{Value:value});}
function c45Settings_(v){const o={};v.rows.forEach(r=>o[String(r.Key)]=r.Value);return o;}
function c45State_(){const ss=c45Book_(),v={ss};Object.keys(C45.headers).forEach(k=>v[k]=c45View_(ss,k));v.cfg=c45Settings_(v.Settings);v.docByID=new Map(v.Documents.rows.map(r=>[String(r.Document_ID),r]));v.sourceByKey=new Map(v['Source Ledger'].rows.map(r=>[String(r.Source_Key),r]));v.sourceHashByKey=new Map(v['Source Ledger'].rows.map(r=>[String(r.Source_Key),String(r.Attachment_Hash||'')]));v.linesByDoc=new Map();v['Cost Lines'].rows.forEach(r=>{const key=String(r.Document_ID);if(!v.linesByDoc.has(key))v.linesByDoc.set(key,[]);v.linesByDoc.get(key).push(r);});return v;}
function c45Bool_(x){return String(x||'').trim().toUpperCase()==='YES';}
function c45Int_(x,d,lo,hi){const n=Number(x);return Number.isFinite(n)?Math.max(lo,Math.min(hi,Math.floor(n))):d;}
function c45Num_(x){if(x===null||x===undefined||x==='')return null;const n=Number(x);return Number.isFinite(n)?n:null;}
function c45Money_(x){const n=c45Num_(x);return n===null?null:Math.round((n+Number.EPSILON)*100)/100;}
function c45Norm_(x){return String(x||'').toLowerCase().normalize('NFKD').replace(/[^a-z0-9]+/g,' ').trim();}
function c45Vendor_(x){let s=c45Norm_(x).replace(/\b(the|inc|llc|ltd|corp|corporation|company|co)\b/g,' ').replace(/\s+/g,' ').trim();if(s.length>9&&s.endsWith('co'))s=s.slice(0,-2);return s;}
function c45Number_(x){return c45Norm_(x).replace(/\s+/g,'').replace(/^(\D*)0+(\d+)$/,'$1$2');}
function c45Address_(x){return c45Norm_(x).replace(/\b(north|n)\b/g,'n').replace(/\b(south|s)\b/g,'s').replace(/\b(east|e)\b/g,'e').replace(/\b(west|w)\b/g,'w').replace(/\b(road|rd)\b/g,'rd').replace(/\b(street|st)\b/g,'st').replace(/\b(avenue|ave)\b/g,'ave').replace(/\b(drive|dr)\b/g,'dr');}
function c45Unit_(x){const s=c45Norm_(x).replace(/ /g,'');const a={sf:'SF',sqft:'SF',squarefoot:'SF',squarefeet:'SF',lf:'LF',linearfeet:'LF',linearfoot:'LF',ea:'EA',each:'EA',hr:'HR',hour:'HR',hours:'HR',cy:'CY',cubicyard:'CY',sq:'SQ',square:'SQ',bf:'BF',boardfoot:'BF',tons:'TON',ton:'TON',day:'DAY',load:'LOAD',loads:'LOAD',yd:'YD',feet:'LF',foot:'LF'};return a[s]||String(x||'').toUpperCase().trim();}
function c45CostType_(x){const s=c45Norm_(x);return ['material','labor','subcontractor','equipment','mixed'].includes(s)?s[0].toUpperCase()+s.slice(1):'Mixed';}
function c45Hash_(value){const bytes=typeof value==='string'?Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256,value,Utilities.Charset.UTF_8):Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256,value);return bytes.map(x=>('0'+(x&255).toString(16)).slice(-2)).join('');}
function c45Safe_(x){if(typeof x==='string'&&/^[=+@\-]/.test(x))return "'"+x;return x===null||x===undefined?'':x;}
function c45Median_(xs){const a=xs.slice().sort((x,y)=>x-y);return a.length%2?a[(a.length-1)/2]:(a[a.length/2-1]+a[a.length/2])/2;}

// ---------- Job labels: only verified projects. New labels appear as candidates. ----------
function c45LabelIgnored_(name){const n=c45Norm_(name);return !n||C45.ignoreLabels.includes(n)||/^(cci4|cci45|costagent)\//i.test(name);}
function syncJobLabelsV45(){const st=c45State_();return c45SyncLabels_(st);}
function c45SyncLabels_(st){const rules=st['Label Rules'],projects=st.Projects,known=new Map(rules.rows.map(r=>[String(r.Gmail_Label),r]));const prefix=String(st.cfg.JOB_LABEL_PREFIX||'').trim(),newRows=[];let created=0;
  GmailApp.getUserLabels().forEach(label=>{const name=label.getName();if(c45LabelIgnored_(name)||known.has(name))return;const verified=!!prefix&&name.startsWith(prefix)&&name.length>prefix.length;const id=verified?'P-'+c45Hash_(name).slice(0,12):'';newRows.push({Gmail_Label:name,Role:verified?'Project':'Candidate',Project_ID:id,Verified:verified?'YES':'NO',Note:verified?'Trusted job-label prefix':'Check whether this is a job label'});});
  c45Append_(rules,newRows);const projectByID=new Set(projects.rows.map(x=>String(x.Project_ID)));
  rules.rows.forEach(rule=>{if(rule.Role!=='Project'||!c45Bool_(rule.Verified)||!rule.Gmail_Label)return;const id=String(rule.Project_ID||'P-'+c45Hash_(rule.Gmail_Label).slice(0,12));if(!rule.Project_ID)c45Patch_(rules,rule,{Project_ID:id});if(projectByID.has(id))return;const labelName=prefix&&String(rule.Gmail_Label).startsWith(prefix)?String(rule.Gmail_Label).slice(prefix.length):String(rule.Gmail_Label);c45Append_(projects,[{Project_ID:id,Project_Name:labelName,Gmail_Label:rule.Gmail_Label,Job_Address:'',Aliases:'',Verified:'YES',Last_Seen:new Date()}]);projectByID.add(id);created++;});
  Logger.log(JSON.stringify({newLabelRules:newRows.length,newProjects:created,verifiedRules:rules.rows.filter(r=>r.Role==='Project'&&c45Bool_(r.Verified)).length}));return {newLabels:newRows.length,newProjects:created};}
function c45Project_(thread,st){const labels=new Set(thread.getLabels().map(l=>l.getName())),rules=st['Label Rules'].rows.filter(r=>labels.has(String(r.Gmail_Label))&&r.Role==='Project'&&c45Bool_(r.Verified));const ids=[...new Set(rules.map(r=>String(r.Project_ID)).filter(Boolean))];if(ids.length!==1)return {ok:false,reason:ids.length>1?'Conflicting job labels':'No verified job label'};const project=st.Projects.rows.find(p=>String(p.Project_ID)===ids[0]);return project?{ok:true,project}:{ok:false,reason:'Job label has no Projects row'};}
function previewIntakeV45(){const st=c45State_();c45SyncLabels_(st);const query='has:userlabels newer_than:'+c45Int_(st.cfg.LIVE_LOOKBACK_DAYS,14,1,365)+'d -in:trash -in:spam -in:drafts';const threads=GmailApp.search(query,0,20);const result=threads.map(t=>({subject:t.getFirstMessageSubject(),job:c45Project_(t,st)}));Logger.log(JSON.stringify(result));return result;}

// ---------- Intake scheduling and controlled history ----------
function testBillingV45(){const st=c45State_(),key=PropertiesService.getScriptProperties().getProperty('OPENAI_API_KEY');if(!key)throw Error('OPENAI_API_KEY missing in billing Script Properties');Logger.log(JSON.stringify({version:C45.version,spreadsheet:st.ss.getUrl(),projects:st.Projects.rows.length,verifiedLabels:st['Label Rules'].rows.filter(r=>r.Role==='Project'&&c45Bool_(r.Verified)).length,documents:st.Documents.rows.length,ledger:st['Source Ledger'].rows.length,triggers:ScriptApp.getProjectTriggers().map(t=>t.getHandlerFunction())}));}
function processNewMailV45(){return c45Run_('LIVE');}
function processPilotV45(){return c45Run_('PILOT');}
function hourlyBackfillV45(){return c45Run_('HISTORY');}
function c45Run_(mode){const lock=LockService.getScriptLock();if(!lock.tryLock(1000)){Logger.log('Another run is active; skipped.');return;}const started=Date.now(),stats={mode,threads:0,sources:0,documents:0,reviews:0,notes:[]};
  try{const st=c45State_();c45SyncLabels_(st);const cfg=st.cfg;if(mode==='HISTORY'&&!c45Bool_(cfg.HISTORY_ENABLED)){Logger.log('History disabled');return;}
    let query='',start=0,count=0;if(mode==='PILOT'){query='label:"CCI45/Pilot" -in:trash -in:spam';count=20;}else if(mode==='LIVE'){start=c45Int_(cfg.LIVE_OFFSET,0,0,100000000);query='has:userlabels newer_than:'+c45Int_(cfg.LIVE_LOOKBACK_DAYS,14,1,365)+'d -in:trash -in:spam -in:drafts';count=c45Int_(cfg.LIVE_SCAN_THREADS,40,1,100);}else{query=String(cfg.HISTORY_QUERY||'has:userlabels -in:trash -in:spam -in:drafts');if(!/has:userlabels|label:/i.test(query))query+=' has:userlabels';if(cfg.HISTORY_CUTOFF_DATE)query+=' before:'+String(cfg.HISTORY_CUTOFF_DATE);start=c45Int_(cfg.HISTORY_OFFSET,0,0,100000000);count=c45Int_(cfg.HISTORY_PAGE_SIZE,30,1,100);}
    const threads=GmailApp.search(query,start,count),limit=mode==='HISTORY'?c45Int_(cfg.HISTORY_MAX_SOURCES,2,1,15):c45Int_(cfg.LIVE_MAX_SOURCES,3,1,15);let completed=0,partial=false;
    for(const thread of threads){if(Date.now()-started>c45Int_(cfg.MAX_RUNTIME_MS,210000,30000,270000)||stats.sources>=limit){partial=true;break;}stats.threads++;const matched=c45Project_(thread,st);
      if(!matched.ok){if(matched.reason==='Conflicting job labels')c45Review_(st,'','Conflicting job labels',thread.getFirstMessageSubject(),'Correct Gmail job labels before intake',stats);completed++;continue;}
      const result=c45ProcessThread_(thread,matched.project,st,stats,limit,started);if(!result.finished){partial=true;break;}completed++;
    }
    if(mode==='LIVE')c45Set_(st.Settings,'LIVE_OFFSET',!partial&&threads.length<count?0:start+completed);
    if(mode==='HISTORY'){c45Set_(st.Settings,'HISTORY_OFFSET',start+completed);if(!partial&&threads.length<count){c45Set_(st.Settings,'HISTORY_ENABLED','NO');c45RemoveTrigger_('hourlyBackfillV45');stats.notes.push('History search exhausted: trigger stopped. Re-run from offset 0 for coverage audit.');}}
    if(stats.documents)c45RefreshRates_(st);Logger.log(JSON.stringify(stats));return stats;
  }catch(e){stats.notes.push('ERROR: '+String(e.stack||e));Logger.log(String(e.stack||e));throw e;}
  finally{try{const ss=c45Book_(),log=c45View_(ss,'Run Log');c45Append_(log,[{At:new Date(),Mode:mode,Threads_Checked:stats.threads,Sources_Processed:stats.sources,Docs_Recorded:stats.documents,Reviews:stats.reviews,Duration_s:Math.round((Date.now()-started)/1000),Note:stats.notes.join(' | ').slice(0,700)}]);}finally{lock.releaseLock();}}
}
function c45SourceItems_(msg,settings){const files=[];msg.getAttachments({includeInlineImages:false,includeAttachments:true}).forEach((blob,i)=>{const name=blob.getName()||'attachment',ext=name.toLowerCase();let mime=(blob.getContentType()||'').toLowerCase();if(!/^image\/|application\/pdf|application\/vnd/.test(mime)){if(/\.png$/i.test(ext))mime='image/png';else if(/\.jpe?g$/i.test(ext))mime='image/jpeg';else if(/\.webp$/i.test(ext))mime='image/webp';else if(/\.pdf$/i.test(ext))mime='application/pdf';else if(/\.xlsx$/i.test(ext))mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';else if(/\.xls$/i.test(ext))mime='application/vnd.ms-excel';else if(/\.csv$/i.test(ext))mime='text/csv';}if(/\.(pdf|png|jpe?g|webp|xls|xlsx|csv)$/i.test(ext)||/^(image\/(jpeg|png|webp)|application\/pdf)$/.test(mime)){files.push({name,blob,mime,index:i});}});
  const text=(msg.getSubject()||'')+' '+(msg.getPlainBody()||'').slice(0,1500);
  if(!files.length&&c45Bool_(settings.PROCESS_BODY_ONLY)&&/invoice|receipt|quote|estimate|payment|delivery|order|credit|subtotal|amount due|total due|purchase|job site|project address|site address|work location/i.test(text))files.push({name:'(email body)',blob:null,mime:'text/plain',index:-1});return files;}
function c45ProcessThread_(thread,project,st,stats,limit,started){for(const msg of thread.getMessages()){for(const src of c45SourceItems_(msg,st.cfg)){
    if(stats.sources>=limit||Date.now()-started>c45Int_(st.cfg.MAX_RUNTIME_MS,210000,30000,270000))return {finished:false};
    const bytes=src.blob?src.blob.getBytes():null,hash=bytes?c45Hash_(bytes):c45Hash_((msg.getPlainBody()||'').slice(0,25000)),sourceKey='SRC-'+c45Hash_(msg.getId()+'|'+src.index+'|'+hash).slice(0,24),prev=st.sourceByKey.get(sourceKey);
    if(prev&&['DONE','SUPPORT','COPY','NO_DOC','ERROR'].includes(String(prev.State)))continue; // safe replay of partial/errored source
    stats.sources++;const ledger=prev||c45Append_(st['Source Ledger'],[{Source_Key:sourceKey,Thread_ID:thread.getId(),Message_ID:msg.getId(),Filename:src.name,Attachment_Hash:hash,Document_ID:'',State:'WORKING',Attempts:0,Updated_At:new Date(),Error:''}])[0];st.sourceByKey.set(sourceKey,ledger);st.sourceHashByKey.set(sourceKey,hash);
    c45Patch_(st['Source Ledger'],ledger,{State:'WORKING',Attempts:Number(ledger.Attempts||0)+1,Updated_At:new Date(),Error:''});
    try{if(bytes&&bytes.length>c45Int_(st.cfg.MAX_ATTACHMENT_MB,12,1,25)*1024*1024)throw Error('Attachment exceeds configured file size');
      const data=c45Extract_(msg,src,project,st.cfg);if(!data.is_document){c45EnrichProject_(st,project,data);c45Patch_(st['Source Ledger'],ledger,{State:'NO_DOC',Updated_At:new Date()});continue;}
      const saved=c45Save_(st,data,{msg,src,thread,project,sourceKey,hash},stats);c45Patch_(st['Source Ledger'],ledger,{Document_ID:saved.id,State:saved.state,Updated_At:new Date(),Error:''});stats.documents++;
    }catch(e){c45Patch_(st['Source Ledger'],ledger,{State:'ERROR',Updated_At:new Date(),Error:String(e.message||e).slice(0,500)});c45Review_(st,sourceKey,'Processing error',src.name+': '+String(e.message||e),'Correct source or API problem and set ledger State=RETRY',stats);}
  }}return {finished:true};}
function c45Review_(st,docId,type,detail,action,stats){const q=st['Review Queue'];if(q.rows.some(r=>String(r.Document_ID)===String(docId)&&r.Issue_Type===type&&r.Details===detail&&r.Status==='OPEN'))return;c45Append_(q,[{Review_ID:'R-'+Utilities.getUuid().slice(0,12),Document_ID:docId,Issue_Type:type,Details:c45Safe_(detail),Suggested_Action:action,Status:'OPEN',Created_At:new Date()}]);if(stats)stats.reviews++;}

// ---------- Structured AI extraction (one API call per new relevant source) ----------
function c45Schema_(){const str={type:'string'},num={type:['number','null']},conf={type:'number'};const item={type:'object',additionalProperties:false,properties:{normalized_item:str,description:str,trade:str,cost_type:str,quantity:num,unit:str,line_cost:num,brand_spec:str,sku:str,cost_code_hint:str,confidence:conf},required:['normalized_item','description','trade','cost_type','quantity','unit','line_cost','brand_spec','sku','cost_code_hint','confidence']};const scope={type:'object',additionalProperties:false,properties:{detail:str,quantity:num,unit:str,amount:num,evidence:str},required:['detail','quantity','unit','amount','evidence']};return {type:'object',additionalProperties:false,properties:{is_document:{type:'boolean'},document_type:str,vendor:str,document_number:str,invoice_number:str,order_number:str,receipt_number:str,issue_date:str,job_name:str,job_address:str,subtotal:num,tax:num,freight:num,discount:num,total:num,confidence:conf,scope_summary:str,notes:str,reference_ids:{type:'array',items:str},items:{type:'array',items:item},scope:{type:'array',items:scope}},required:['is_document','document_type','vendor','document_number','invoice_number','order_number','receipt_number','issue_date','job_name','job_address','subtotal','tax','freight','discount','total','confidence','scope_summary','notes','reference_ids','items','scope']};}
function c45Extract_(msg,src,project,cfg){const key=PropertiesService.getScriptProperties().getProperty('OPENAI_API_KEY');if(!key)throw Error('Missing OPENAI_API_KEY');const prompt=[
  'Extract ONE construction business document or relevant email into JSON exactly matching the schema. Never invent monetary values, quantities, numbers, job sites or scope.',
  'DOCUMENT TYPES: Invoice, Receipt, Delivery Confirmation, Packing Slip, Order Confirmation, Purchase Order, Quote, Estimate, Statement, Payment Confirmation, Credit, Other.',
  'Read and extract EVERY documented priced line item regardless of the paperwork title, including quantities and item descriptions. Distinguish an actual invoice/charged retail receipt from merely supporting paperwork. Delivery/shipment/packing/order confirmations are not automatically extra expenses even when they show prices, but if the attachment genuinely contains an invoice/receipt classify it by that financial evidence. A payment confirmation of a previously invoiced charge is not a second expense.',
  'Extract invoice/order/receipt reference numbers separately, and shared references in reference_ids. Do not conflate purchase-order numbers with invoice numbers.',
  'For each priced product/service, output one item and its documented quantity, unit, amount and cost_type (Material, Labor, Subcontractor, Equipment, Mixed). For a lump sum, output ONE priced item, with descriptive scope details separately in scope. Do not price every scope note.',
  'Normalize item descriptions to useful estimating names (e.g. Framing lumber, Excavation, Concrete flatwork) while preserving source description and brand. Do not conflate blended installed cost with materials-only cost.',
  'Find supporting construction details from email body and attached file, including project name/address (but distinguish vendor remit address). The verified Gmail job is '+project.Project_Name+' | '+project.Job_Address+'. The verified label controls project classification.',
  'Confidence is between 0 and 1. Never fabricate quantities or assume invoice is paid. If source is not a real cost-related document, is_document=false.',
  'Email: From '+msg.getFrom()+' | Subject '+msg.getSubject()+' | Body '+String(msg.getPlainBody()||'').slice(0,c45Int_(cfg.MAX_EMAIL_CHARS,12000,500,30000))
].join('\n\n');const content=[{type:'input_text',text:prompt}];if(src.blob){const data='data:'+(src.mime||'application/pdf')+';base64,'+Utilities.base64Encode(src.blob.getBytes());if(/^image\//.test(src.mime))content.push({type:'input_image',image_url:data});else content.push({type:'input_file',filename:src.name,file_data:data});}
  const payload={model:String(cfg.OPENAI_MODEL||'gpt-4.1-mini'),input:[{role:'user',content}],text:{format:{type:'json_schema',name:'construction_cost_v45',strict:true,schema:c45Schema_()}}};const response=UrlFetchApp.fetch(C45.endpoint,{method:'post',contentType:'application/json',headers:{Authorization:'Bearer '+key},payload:JSON.stringify(payload),muteHttpExceptions:true});if(response.getResponseCode()<200||response.getResponseCode()>=300)throw Error('OpenAI HTTP '+response.getResponseCode()+': '+response.getContentText().slice(0,800));const obj=JSON.parse(response.getContentText());let text=obj.output_text||'';(obj.output||[]).forEach(o=>(o.content||[]).forEach(t=>{if(t.type==='output_text'&&t.text&&!text)text=t.text;}));if(!text)throw Error('OpenAI did not return structured JSON: '+String(obj.status||''));return JSON.parse(text);}

// ---------- Purchase identification; 80% = candidate similarity, NOT automatic expense deletion ----------
function c45Type_(type){const s=c45Norm_(type);if(/packing/.test(s))return 'Packing Slip';if(/delivery|shipping|shipment/.test(s))return 'Delivery Confirmation';if(/order confirmation|order acknowledgment/.test(s))return 'Order Confirmation';if(/payment confirmation|remittance|payment receipt/.test(s))return 'Payment Confirmation';if(/purchase order/.test(s))return 'Purchase Order';if(/credit/.test(s))return 'Credit';if(/invoice/.test(s))return 'Invoice';if(/estimate/.test(s))return 'Estimate';if(/quote/.test(s))return 'Quote';if(/statement/.test(s))return 'Statement';if(/receipt/.test(s))return 'Receipt';return 'Other';}
function c45EffectiveType_(doc,src,msg){const filename=c45Norm_(src.name||'');const t=c45Type_(doc.document_type);/* A true PDF invoice or purchase receipt overrides a logistics-related EMAIL filename. */if(['Invoice','Receipt','Credit'].includes(t))return t;if(/packing slip|packing list|packing ticket/.test(filename))return 'Packing Slip';if(/delivery confirmation|proof of delivery|shipment confirmation|tracking notice/.test(filename))return 'Delivery Confirmation';if(/order confirmation|order acknowledgement|order acknowledgment/.test(filename))return 'Order Confirmation';if(t==='Other'&&/delivery confirmation|your shipment has arrived/.test(c45Norm_(msg.getSubject())))return 'Delivery Confirmation';return t;}
function c45Refs_(doc){const out=[];function add(tag,x){const n=c45Number_(x);if(n&&n.length>=3)out.push(tag+':'+n);}add('INV',doc.invoice_number);add('ORD',doc.order_number);add('REC',doc.receipt_number);(doc.reference_ids||[]).forEach(r=>add('REF',r));return [...new Set(out)];}
function c45Tokens_(text){return new Set(c45Norm_(text).split(' ').filter(s=>s.length>2));}
function c45Jaccard_(a,b){const x=c45Tokens_(a),y=c45Tokens_(b);if(!x.size||!y.size)return 0;let n=0;x.forEach(t=>{if(y.has(t))n++;});return n/(x.size+y.size-n);}
function c45ItemSimilarity_(newItems,oldLines){const aa=(newItems||[]).filter(x=>c45Money_(x.line_cost)!==null),bb=(oldLines||[]).filter(x=>c45Money_(x.Line_Cost)!==null);if(!aa.length||!bb.length)return 0;const used=new Set();let score=0;aa.forEach(a=>{let best=0,idx=-1;bb.forEach((b,i)=>{if(used.has(i))return;const sim=c45Jaccard_(a.normalized_item+' '+a.description+' '+a.sku,b.Normalized_Item+' '+b.Source_Description+' '+b.SKU),mon=(Math.abs(Number(a.line_cost||0)-Number(b.Line_Cost||0))<=0.02)?1:0,qty=(a.quantity==null||!b.Qty||Math.abs(Number(a.quantity)-Number(b.Qty))<0.001)?1:0,s=.6*sim+.3*mon+.1*qty;if(s>best){best=s;idx=i;}});if(idx>=0&&best>=.65){score+=best;used.add(idx);}});return score/Math.max(aa.length,bb.length);}
function c45Match_(st,d,project,type,excludeID,sourceHash){const nv=c45Vendor_(d.vendor),refs=c45Refs_(d),total=c45Money_(d.total),threshold=Number(st.cfg.SIMILARITY_REVIEW_THRESHOLD||.8),out=[];
 for(const old of st.Documents.rows){if(old.Document_ID===excludeID||['WRITING','HELD_REVIEW','HELD_DUPLICATE','EXCLUDED'].includes(String(old.Status)))continue;const sameHash=!!sourceHash&&sourceHash===st.sourceHashByKey.get(String(old.Source_Key));if(!old.Vendor||c45Vendor_(old.Vendor)!==nv&&!sameHash)continue;const sameProject=String(old.Project_ID)===String(project.Project_ID);const oldRefs=String(old.Refs||'').split(';').filter(Boolean);const oldVals=new Set(oldRefs.map(x=>x.split(':').slice(1).join(':'))),common=refs.filter(x=>oldRefs.includes(x)||oldVals.has(x.split(':').slice(1).join(':')));const exactInvoice=common.some(x=>x.startsWith('INV:'))||!!(d.invoice_number&&c45Number_(d.invoice_number)===c45Number_(old.Document_No));const commonOrder=common.some(x=>x.startsWith('ORD:'))||oldRefs.some(x=>x.startsWith('ORD:')&&refs.some(y=>y.slice(y.indexOf(':')+1)===x.slice(4)));const amountSame=total!==null&&c45Money_(old.Total)!==null&&Math.abs(total-Number(old.Total))<=0.02;const lineSim=c45ItemSimilarity_(d.items,st.linesByDoc.get(String(old.Document_ID))||[]);const relatedText=c45Jaccard_(d.scope_summary,old.Scope_Summary);
  const docNoDifferent=c45Number_(d.document_number)&&c45Number_(old.Document_No)&&c45Number_(d.document_number)!==c45Number_(old.Document_No);
  let score=common.length?Math.min(1,(exactInvoice?.62:commonOrder?.53:.4)+(amountSame?.17:0)+lineSim*.25+relatedText*.08):(.65*lineSim+.25*(amountSame?1:0)+.10*relatedText);if(sameHash)score=1;if(!sameProject)score=(sameHash||exactInvoice)?.95:0;
  const isCost=C45.costTypes.includes(type),oldCost=C45.costTypes.includes(String(old.Doc_Type));
  // One PO can cover many legitimate invoices. Both invoices with DIFFERENT IDs remain separate.
  const clearlyDifferent=type==='Invoice'&&old.Doc_Type==='Invoice'&&docNoDifferent&&!exactInvoice;
  const anchor=sameProject&&(sameHash||(exactInvoice&&amountSame)||(commonOrder&&amountSame&&lineSim>=.8)||(commonOrder&&amountSame&&(!isCost||!oldCost)));
  if(!clearlyDifferent&&(score>=threshold||anchor))out.push({old,score,anchor,amountSame,exactInvoice,commonOrder,lineSim,possible:score>=threshold});
 }
 out.sort((a,b)=>(Number(b.anchor)-Number(a.anchor))||(Number(b.old.Status==='COUNTED')-Number(a.old.Status==='COUNTED'))||(b.score-a.score));return out[0]||null;}
function c45Priority_(type){return type==='Invoice'?3:type==='Receipt'?2:type==='Credit'?1:0;}
function c45Decision_(type,match,data){if(!C45.costTypes.includes(type))return {status:'SUPPORTING',impact:0,relation:match?'LINKED':'SUPPORTING'};const total=c45Money_(data.total);if(total===null)return {status:'HELD_REVIEW',impact:0,reason:'Missing document total'};const signed=type==='Credit'?-Math.abs(total):Math.abs(total);
 if(!match)return {status:'COUNTED',impact:signed};const old=match.old;
 if(!C45.costTypes.includes(String(old.Doc_Type)))return {status:'COUNTED',impact:signed,relation:'REPLACES_SUPPORT'};
 // Credits adjust costs, but copies of the SAME credit memo must not double-subtract.
 if(type==='Credit'&&old.Doc_Type==='Credit'){if(match.anchor&&match.amountSame)return {status:'COPY',impact:0};if(match.possible)return {status:'HELD_DUPLICATE',impact:0,reason:'Possible duplicate credit memo'};}
 if(type==='Credit'||old.Doc_Type==='Credit')return {status:'COUNTED',impact:signed};
 if(match.anchor&&match.amountSame){if(c45Priority_(type)>c45Priority_(old.Doc_Type))return {status:'COUNTED',impact:signed,replace:old.Document_ID,relation:'UPGRADE_AUTHORITY'};return {status:'COPY',impact:0,relation:'CONFIRMED_COPY'};}
 if(match.possible)return {status:'HELD_DUPLICATE',impact:0,reason:'Similar purchase documentation without a conclusive shared transaction identifier'};
 return {status:'COUNTED',impact:signed};}
function c45CostCode_(st,item){const requested=String(item.cost_code_hint||'').trim(),codes=st['Cost Codes'].rows;if(requested&&codes.some(c=>c.Cost_Code===requested))return requested;const text=c45Norm_(item.normalized_item+' '+item.trade+' '+item.description);let best='',score=0;codes.forEach(c=>{const terms=String(c.Aliases||'').split(/[;,]/).concat([String(c.Cost_Code||'').replace(/^\S+\s*-\s*/,'')]);terms.forEach(t=>{const token=c45Norm_(t);if(token.length>3&&text.includes(token)&&token.length>score){score=token.length;best=c.Cost_Code;}});});return best;}

function c45EnrichProject_(st,project,doc){if(!doc.job_address||Number(doc.confidence||0)<Number(st.cfg.MIN_CLASS_CONFIDENCE||.95))return;if(!project.Job_Address)c45Patch_(st.Projects,project,{Job_Address:c45Safe_(doc.job_address),Last_Seen:new Date()});else if(c45Address_(doc.job_address)!==c45Address_(project.Job_Address))c45Review_(st,'','Project address mismatch',String(project.Project_Name)+': '+doc.job_address+' versus '+project.Job_Address,'Verify location context; never replace an existing verified address automatically');}
// ---------- Linked source records and idempotent write ----------
function c45BillingAddress_(){const p=PropertiesService.getScriptProperties(),email=String(p.getProperty('BILLING_GMAIL_ADDRESS')||Session.getEffectiveUser().getEmail()||'').trim();if(!/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(email))throw Error('Set BILLING_GMAIL_ADDRESS Script Property');return email;}
function c45Link_(msg){try{const hdr=String(msg.getHeader('Message-ID')||'').replace(/^<|>$/g,'').trim();if(!hdr)return '';const url='https://mail.google.com/mail/u/?authuser='+encodeURIComponent(c45BillingAddress_())+'#search/'+encodeURIComponent('in:anywhere rfc822msgid:'+hdr);return '=HYPERLINK("'+url.replace(/"/g,'""')+'","Open billing email")';}catch(e){Logger.log('Source link unavailable: '+e.message);return '';}}
function c45Save_(st,d,ctx,stats){const type=c45EffectiveType_(d,ctx.src,ctx.msg),id='D-'+c45Hash_(ctx.sourceKey).slice(0,20),docs=st.Documents,old=st.docByID.get(id),confidence=Number(d.confidence||0),subtotal=c45Money_(d.subtotal),total=c45Money_(d.total),match=c45Match_(st,d,ctx.project,type,id,ctx.hash),decision=c45Decision_(type,match,d);let status=decision.status,impact=decision.impact,issues=[];
 const essential=C45.costTypes.includes(type)&&(!d.vendor||total===null||confidence<Number(st.cfg.MIN_CLASS_CONFIDENCE||.95));if(essential){status='HELD_REVIEW';impact=0;issues.push('Missing supplier/total or document confidence below threshold');}
 if(decision.reason)issues.push(decision.reason);
 const invoiceNum=String(d.invoice_number||d.document_number||d.receipt_number||'').trim(),refs=c45Refs_(d),txn=(decision.replace&&st.docByID.get(decision.replace)?.Transaction_ID)||(match&&match.old.Transaction_ID)||'TX-'+c45Hash_([c45Vendor_(d.vendor),ctx.project.Project_ID,refs.join('|'),ctx.sourceKey].join('|')).slice(0,16);
 // Some confirmations list prices but are not costs. Never use their price lines as rates.
 const docRow={Document_ID:id,Source_Key:ctx.sourceKey,Thread_ID:ctx.thread.getId(),Message_ID:ctx.msg.getId(),Vendor:c45Safe_(d.vendor),Document_No:c45Safe_(invoiceNum),Doc_Type:type,Document_Date:c45Safe_(d.issue_date),Project_ID:ctx.project.Project_ID,Project_Name:ctx.project.Project_Name,Job_Address:c45Safe_(d.job_address||ctx.project.Job_Address),Subtotal:subtotal,Tax:c45Money_(d.tax),Freight:c45Money_(d.freight),Discount:c45Money_(d.discount),Total:total,Cost_Impact:0,Status:'WRITING',Confidence:confidence,Transaction_ID:txn,Related_Doc_ID:match?match.old.Document_ID:'',Refs:refs.join(';'),Scope_Summary:c45Safe_(d.scope_summary),Source_Link:c45Link_(ctx.msg),Notes:c45Safe_(d.notes),Recorded_At:new Date()};
 const row=old?(c45Patch_(docs,old,docRow),old):c45Append_(docs,[docRow])[0];st.docByID.set(id,row);
 const found=new Set((st.linesByDoc.get(id)||[]).map(r=>r.Line_ID)),priced=(d.items||[]).filter(x=>c45Money_(x.line_cost)!==null);if(!priced.length&&C45.costTypes.includes(type)&&subtotal!==null)priced.push({normalized_item:'Lump-sum scope',description:d.scope_summary||d.notes||'Unitemized charge',trade:'',cost_type:'Mixed',quantity:null,unit:'',line_cost:subtotal,confidence,brand_spec:'',sku:'',cost_code_hint:''});
 const sourceSum=priced.reduce((s,x)=>s+Number(x.line_cost||0),0),unreconciled=subtotal!==null&&priced.length&&Math.abs(sourceSum-subtotal)>Math.max(1,Math.abs(subtotal)*.01);
 if(unreconciled)issues.push('Itemized lines '+sourceSum.toFixed(2)+' differ from documented subtotal '+subtotal.toFixed(2));
 const lineRows=[];priced.forEach((x,i)=>{const lineId=id+'-L'+String(i+1).padStart(3,'0');if(found.has(lineId))return;const qty=c45Num_(x.quantity),amount=c45Money_(x.line_cost),unit=c45Unit_(x.unit),costDocument=status==='COUNTED'&&type!=='Credit',goodConfidence=Number(x.confidence||0)>=Number(st.cfg.MIN_RATE_CONFIDENCE||.95),rateStatus=!costDocument?'NO_SUPPORTING':unreconciled?'NO_MISMATCH':!goodConfidence?'NO_CONFIDENCE':!(qty>0&&unit)?'NO_QTY':'YES';lineRows.push({Line_ID:lineId,Document_ID:id,Project_ID:ctx.project.Project_ID,Cost_Code:c45CostCode_(st,x),Normalized_Item:c45Safe_(x.normalized_item||'Unclassified'),Source_Description:c45Safe_(x.description),Cost_Type:c45CostType_(x.cost_type),Qty:qty,Unit:unit,Line_Cost:amount,Unit_Cost:qty>0&&amount!==null?Math.round(amount/qty*10000)/10000:'',Rate_Eligible:rateStatus,Confidence:Number(x.confidence||0),Brand_Spec:c45Safe_(x.brand_spec),SKU:c45Safe_(x.sku),Manual_Qty:'',Manual_Unit:'',Scope_Notes:c45Safe_(d.scope_summary)});});
 const appended=c45Append_(st['Cost Lines'],lineRows);if(!st.linesByDoc.has(id))st.linesByDoc.set(id,[]);st.linesByDoc.get(id).push(...appended);
 const priorScope=new Set(st['Scope Notes'].rows.filter(r=>r.Document_ID===id).map(r=>r.Scope_ID));const scopeRows=(d.scope||[]).map((x,i)=>({Scope_ID:id+'-S'+String(i+1).padStart(3,'0'),Document_ID:id,Detail:c45Safe_(x.detail),Qty:c45Num_(x.quantity),Unit:c45Unit_(x.unit),Amount:c45Money_(x.amount),Evidence:c45Safe_(x.evidence)})).filter(x=>!priorScope.has(x.Scope_ID));c45Append_(st['Scope Notes'],scopeRows);
 // Carry forward verified project address only when blank; never overwrite a conflicting address automatically.
 if(d.job_address){if(!ctx.project.Job_Address)c45Patch_(st.Projects,ctx.project,{Job_Address:c45Safe_(d.job_address),Last_Seen:new Date()});else if(c45Address_(d.job_address)!==c45Address_(ctx.project.Job_Address))issues.push('Job address differs from verified project record');}
 if(decision.replace&&!essential){const prior=st.docByID.get(decision.replace);if(prior&&prior.Status==='COUNTED')c45Patch_(docs,prior,{Cost_Impact:0,Status:'COPY',Related_Doc_ID:id});}
 // Rate eligibility is checked AGAINST current authoritative document when refreshing the price library.
 c45Patch_(docs,row,{Cost_Impact:impact,Status:status});
 if(issues.length)c45Review_(st,id,status==='HELD_DUPLICATE'?'Possible duplicate':'Document verification',issues.join('; '),'Check only the questionable fields; never invent quantities',stats);
 return {id,state:status==='SUPPORTING'?'SUPPORT':status==='COPY'?'COPY':'DONE'};
}

// ---------- Continuously maintain source-backed unit price history ----------
function refreshPriceHistoryV45(){const lock=LockService.getScriptLock();if(!lock.tryLock(1000))return;try{const st=c45State_();return c45RefreshRates_(st);}finally{lock.releaseLock();}}
// ---------- Confirmed purchase prices + distinctly labeled provisional references ----------
// Still ONE Price History tab. Quotes / delivery details never become booked
// expenses, but unique supported prices with quantities remain visible as
// provisional reference rates. They are NEVER Auto_Use=YES.
function c45RefreshRates_(st){
 const recentMonths=c45Int_(st.cfg.MAX_RATE_AGE_MONTHS,24,1,120),minSamples=c45Int_(st.cfg.MIN_AUTO_SAMPLES,3,1,20);
 const floor=Number(st.cfg.MIN_RATE_CONFIDENCE||.8),buckets=new Map(),refs=new Map(),now=Date.now();
 const counted=st.Documents.rows.filter(d=>d.Status==='COUNTED'&&['Invoice','Receipt'].includes(String(d.Doc_Type)));
 const referenceSeen=new Set();
 function getDate(d){const x=d.Document_Date?new Date(d.Document_Date):new Date(d.Recorded_At);return Number.isFinite(x.getTime())?x:new Date(d.Recorded_At);}
 function refDuplicated(line,doc){
   return counted.some(other=>{
     if(other.Document_ID===doc.Document_ID||String(other.Project_ID)!==String(doc.Project_ID)||c45Vendor_(other.Vendor)!==c45Vendor_(doc.Vendor))return false;
     const shared=c453SharedRefs_(doc,other).length>0||String(doc.Related_Doc_ID)===String(other.Document_ID)||String(doc.Transaction_ID)&&String(doc.Transaction_ID)===String(other.Transaction_ID);
     const comparable=(st.linesByDoc.get(String(other.Document_ID))||[]).some(x=>c453SameLine_(line,x));
     if(!comparable)return false;
     if(shared)return true;
     const dayDelta=Math.abs(getDate(doc)-getDate(other))/86400000;
     return dayDelta<=14&&c45Money_(doc.Total)===c45Money_(other.Total);
   });
 }
 for(const line of st['Cost Lines'].rows){
   const d=st.docByID.get(String(line.Document_ID));if(!d)continue;
   const manualQty=c45Num_(line.Manual_Qty),qty=manualQty>0?manualQty:c45Num_(line.Qty),unit=c45Unit_(manualQty>0?line.Manual_Unit:line.Unit);
   const cost=c45Money_(line.Line_Cost),confidence=Number(line.Confidence||0),status=String(line.Rate_Eligible||'').toUpperCase();
   if(!(qty>0)||!unit||cost===null||cost<0||confidence<floor||status==='NO')continue;
   const verified=d.Status==='COUNTED'&&['Invoice','Receipt'].includes(String(d.Doc_Type))&& (status==='YES'||(status==='NO_QTY'&&manualQty>0));
   const reference=d.Status==='SUPPORTING'&&!verified;
   if(!verified&&!reference)continue;
   if(reference&&refDuplicated(line,d))continue;
   const date=getDate(d);if(now-date.getTime()>recentMonths*31*86400000)continue;
   const name=c45Norm_(line.Normalized_Item),brand=c45Norm_(line.Brand_Spec),type=String(line.Cost_Type||''),code=String(line.Cost_Code||'');if(!name)continue;
   const key=(reference?'REFRATE-':'RATE-')+c45Hash_([code,name,type,unit,brand].join('|')).slice(0,16),map=reference?refs:buckets;
   if(reference){
     const sourceKey=[c45Vendor_(d.Vendor),d.Project_ID,c45Norm_(line.Normalized_Item),unit,cost,qty,c45Norm_(d.Refs||d.Document_No)].join('|');
     // Supporting emails repeated in the same project should not multiply
     // unverified rate observations just because they were forwarded again.
     if(referenceSeen.has(sourceKey))continue;referenceSeen.add(sourceKey);
   }
   const bucket=map.get(key)||{key,code,name:line.Normalized_Item,type,unit,brand:line.Brand_Spec,byDoc:new Map()};
   const observed=bucket.byDoc.get(d.Document_ID)||{amount:0,qty:0,date,doc:d.Document_ID,vendor:String(d.Vendor||''),docType:d.Doc_Type};
   observed.amount+=cost;observed.qty+=qty;bucket.byDoc.set(d.Document_ID,observed);map.set(key,bucket);
 }
 const result=[];
 function finalize(group,reference){
   group.forEach(b=>{
     const obs=[...b.byDoc.values()].filter(x=>x.qty>0).map(x=>({price:x.amount/x.qty,date:x.date,doc:x.doc,vendor:x.vendor,type:x.docType}));
     if(!obs.length)return;
     const values=obs.map(x=>x.price).sort((a,z)=>a-z),n=values.length,median=c45Median_(values),latest=obs.slice().sort((x,y)=>y.date-x.date)[0],average=values.reduce((s,x)=>s+x,0)/n;
     result.push({Rate_Key:b.key,Cost_Code:b.code,Normalized_Item:(reference?'[REFERENCE] ':'')+b.name,Cost_Type:b.type,Unit:b.unit,Brand_Spec:b.brand,Samples:n,Low:values[0],Median:median,Average:average,High:values[n-1],Latest:latest.price,Latest_Date:latest.date,Recommended:median,Auto_Use:reference?'NO':(n>=minSamples&&!!b.code?'YES':'NO'),Source_Documents:obs.map(x=>x.doc).join(';'),Suppliers:[...new Set(obs.map(x=>x.vendor).filter(Boolean))].join('; '),Latest_Supplier:latest.vendor,Notes:reference?'REFERENCE ONLY — priced supporting document; not verified purchased expense. Never auto-use.':!b.code?'Unmapped cost code; reference rate only':n<minSamples?'PROVISIONAL — '+n+' of '+minSamples+' purchased sources':'Based on distinct qualifying invoices/receipts'});
   });
 }
 finalize(buckets,false);finalize(refs,true);
 result.sort((a,b)=>String(a.Cost_Code).localeCompare(String(b.Cost_Code))||String(a.Normalized_Item).localeCompare(String(b.Normalized_Item)));
 const view=st['Price History'],old=view.sh.getLastRow()-1;
 if(old>0)view.sh.getRange(2,1,old,view.h.length).clearContent();
 if(result.length)view.sh.getRange(2,1,result.length,view.h.length).setValues(result.map(r=>view.h.map(k=>r[k]===undefined?'':r[k])));
 Logger.log('Price History refreshed: '+result.length+' groups (including '+refs.size+' reference-only groups)');return result.length;
}

// ---------- Trigger controls; billing and historical processing independent ----------
function removeBillingTriggersV45(){const old=['processNewMailV45','hourlyBackfillV45','processNewMailV4','hourlyBackfillV4','processNewMail','runHistoricalBackfill','scheduledHistoricalBackfillV333'];ScriptApp.getProjectTriggers().forEach(t=>{if(old.includes(t.getHandlerFunction()))ScriptApp.deleteTrigger(t);});Logger.log('Billing triggers paused');}
function installBillingTriggersV45(){removeBillingTriggersV45();ScriptApp.newTrigger('processNewMailV45').timeBased().everyMinutes(15).create();const st=c45State_();if(c45Bool_(st.cfg.HISTORY_ENABLED))ScriptApp.newTrigger('hourlyBackfillV45').timeBased().everyHours(1).create();Logger.log('15-minute live trigger installed; hourly history '+(c45Bool_(st.cfg.HISTORY_ENABLED)?'ON':'OFF'));}
function c45RemoveTrigger_(name){ScriptApp.getProjectTriggers().forEach(t=>{if(t.getHandlerFunction()===name)ScriptApp.deleteTrigger(t);});}
function pauseHistoryV45(){const st=c45State_();c45Set_(st.Settings,'HISTORY_ENABLED','NO');c45RemoveTrigger_('hourlyBackfillV45');}
function startHistoryV45(){const st=c45State_();if(!st.cfg.HISTORY_CUTOFF_DATE){const tomorrow=new Date(Date.now()+86400000);c45Set_(st.Settings,'HISTORY_CUTOFF_DATE',Utilities.formatDate(tomorrow,Session.getScriptTimeZone(),'yyyy/MM/dd'));c45Set_(st.Settings,'HISTORY_OFFSET',0);}c45Set_(st.Settings,'HISTORY_ENABLED','YES');c45RemoveTrigger_('hourlyBackfillV45');ScriptApp.newTrigger('hourlyBackfillV45').timeBased().everyHours(1).create();}
function restartHistoryCoverageV45(){const st=c45State_();c45Set_(st.Settings,'HISTORY_OFFSET',0);Logger.log('History offset reset. Existing source ledger prevents duplicate costs.');}
function retrySourceErrorsV45(){const st=c45State_();let n=0;st['Source Ledger'].rows.filter(r=>r.State==='ERROR').forEach(r=>{c45Patch_(st['Source Ledger'],r,{State:'RETRY',Error:''});n++;});Logger.log('Retriable sources: '+n);}
function billingHealthV45(){const st=c45State_(),report={version:C45.version,verifiedProjects:st.Projects.rows.length,documents:st.Documents.rows.length,counted:st.Documents.rows.filter(x=>x.Status==='COUNTED').length,supporting:st.Documents.rows.filter(x=>x.Status==='SUPPORTING').length,priceGroups:st['Price History'].rows.length,openReviews:st['Review Queue'].rows.filter(x=>x.Status==='OPEN').length,ledger:st['Source Ledger'].rows.length,triggers:ScriptApp.getProjectTriggers().map(t=>t.getHandlerFunction()),historyEnabled:st.cfg.HISTORY_ENABLED};Logger.log(JSON.stringify(report));return report;}

/** OPTIONAL one-time carry-forward of VERIFIED JOB LABEL CONFIGURATION ONLY, not old invoices.
 * Run BEFORE the personal-account reset, AFTER installing billing v4.5 and pausing triggers.
 * Works with v4 Projects/Label Rules headings. Keeps up to 29+ approved labels without retyping.
 */
function captureVerifiedJobLabelsV45(){const ss=c45Book_();function read(name){const sh=ss.getSheetByName(name);if(!sh)return [];const v=sh.getDataRange().getValues(),h=v[0].map(String);return v.slice(1).map(a=>{const r={};h.forEach((k,i)=>r[k]=a[i]);return r;});}
 const projects=read('Projects'),pMap=new Map(projects.map(p=>[String(p.Project_ID),p]));const labels=read('Label Rules').filter(r=>String(r.Role)==='Project'&&String(r.Verified).toUpperCase()==='YES').map(r=>{const p=pMap.get(String(r.Project_ID))||{};return {label:String(r.Gmail_Label||''),name:String(p.Project_Name||r.Gmail_Label||''),address:String(p.Job_Address||p.Address||''),aliases:String(p.Aliases||'')};}).filter(r=>r.label);
 const properties=PropertiesService.getScriptProperties(),raw=JSON.stringify(labels),chunk=5500,count=Math.ceil(raw.length/chunk);for(let i=0;i<count;i++)properties.setProperty('V45_VERIFIED_LABELS_'+i,raw.slice(i*chunk,(i+1)*chunk));properties.setProperty('V45_VERIFIED_LABELS_COUNT',String(count));Logger.log('Saved '+labels.length+' approved job labels in billing Script Properties; not backed-up invoices.');return labels.length;}
function restoreVerifiedJobLabelsV45(){const props=PropertiesService.getScriptProperties(),n=c45Int_(props.getProperty('V45_VERIFIED_LABELS_COUNT'),0,0,50);if(!n)throw Error('No saved label settings. Run captureVerifiedJobLabelsV45 before wiping the prior spreadsheet, or verify labels manually.');let raw='';for(let i=0;i<n;i++)raw+=props.getProperty('V45_VERIFIED_LABELS_'+i)||'';const saved=JSON.parse(raw),st=c45State_(),r=st['Label Rules'],p=st.Projects,known=new Set(r.rows.map(x=>String(x.Gmail_Label)));let restored=0;
 saved.forEach(x=>{if(known.has(x.label))return;const id='P-'+c45Hash_(x.label).slice(0,12);c45Append_(r,[{Gmail_Label:x.label,Role:'Project',Project_ID:id,Verified:'YES',Note:'Previously verified job label (v4 configuration)'}]);c45Append_(p,[{Project_ID:id,Project_Name:x.name||x.label,Gmail_Label:x.label,Job_Address:x.address||'',Aliases:x.aliases||'',Verified:'YES',Last_Seen:new Date()}]);known.add(x.label);restored++;});Logger.log('Restored '+restored+' verified project labels.');return restored;}

/** Rare exception only: verify original source, set Settings REVIEW_DOC_ID and REVIEW_DECISION.
 * COUNT = reviewed separate charge; EXCLUDE = non-expense / duplicate. No Gmail reprocessing.
 */
function resolveReviewedDocumentV45(){const st=c45State_(),id=String(st.cfg.REVIEW_DOC_ID||'').trim(),action=String(st.cfg.REVIEW_DECISION||'').trim().toUpperCase(),row=st.docByID.get(id);
 if(!row||!['HELD_DUPLICATE','HELD_REVIEW'].includes(String(row.Status)))throw Error('Select a HELD_DUPLICATE or HELD_REVIEW Document_ID');if(!['COUNT','EXCLUDE'].includes(action))throw Error('Set REVIEW_DECISION=COUNT or EXCLUDE in Settings');
 if(action==='COUNT'&&!C45.costTypes.includes(row.Doc_Type))throw Error('Supporting or quoted documents cannot become expenses');
 if(action==='COUNT'){
  const existing=st.Documents.rows.find(r=>r.Document_ID!==id&&r.Status==='COUNTED'&&Number(r.Cost_Impact)!==0&&c45Vendor_(r.Vendor)===c45Vendor_(row.Vendor)&&String(r.Project_ID)===String(row.Project_ID)&&c45Number_(r.Document_No)&&c45Number_(r.Document_No)===c45Number_(row.Document_No));
  if(existing)throw Error('A counted document has the same vendor, job and document number: '+existing.Document_ID+'. Do not count twice.');
 }
 const impact=action==='EXCLUDE'?0:row.Doc_Type==='Credit'?-Math.abs(Number(row.Total||0)):Math.abs(Number(row.Total||0));
 c45Patch_(st.Documents,row,{Cost_Impact:impact,Status:action==='EXCLUDE'?'EXCLUDED':'COUNTED',Notes:String(row.Notes||'')+' | Human reviewed '+action+' '+new Date().toISOString()});
 st['Review Queue'].rows.filter(r=>String(r.Document_ID)===id&&r.Status==='OPEN').forEach(r=>c45Patch_(st['Review Queue'],r,{Status:'RESOLVED'}));
 c45Set_(st.Settings,'REVIEW_DOC_ID','');c45Set_(st.Settings,'REVIEW_DECISION','');c45RefreshRates_(st);Logger.log('Reviewed '+id+': '+action+'; cost impact '+impact);return {document:id,decision:action,costImpact:impact};}


// ---------- v4.5.1 source link diagnostics and in-place repair ----------
/** Count blank, plain-text and formula source links without changing data. */
function diagnoseSourceLinksV451(){
  const sh=c45Book_().getSheetByName('Documents');
  if(!sh)throw Error('Documents tab missing');
  const cols=C45.headers.Documents;
  if(sh.getRange(1,1,1,cols.length).getValues()[0].some((h,i)=>h!==cols[i]))
    throw Error('Documents tab does not match v4.5 headers');
  const n=Math.max(0,sh.getLastRow()-1);
  if(!n){Logger.log(JSON.stringify({documents:0}));return;}
  const range=sh.getRange(2,24,n,1),v=range.getValues(),f=range.getFormulas();
  const result={documents:n,workingFormulas:0,plainText:0,blank:0,otherFormulas:0};
  for(let i=0;i<n;i++){
    const formula=String(f[i][0]||''),value=String(v[i][0]||'');
    if(formula){if(/^=HYPERLINK\(/i.test(formula))result.workingFormulas++;else result.otherFormulas++;}
    else if(value.trim())result.plainText++;
    else result.blank++;
  }
  Logger.log(JSON.stringify(result));return result;
}

/** Rebuild saved source links from Documents.Message_ID using the billing Gmail account.
 * Does not re-extract email, reprocess invoices, or alter costs. Up to 30 rows/run.
 * Invoke repeatedly until complete; uses a saved cursor across manual executions.
 */
function repairSourceLinksV451(){
  const lock=LockService.getScriptLock();
  if(!lock.tryLock(1000)){Logger.log('Another billing operation is running; try again.');return;}
  try{
    const sh=c45Book_().getSheetByName('Documents');
    if(!sh)throw Error('Documents tab missing');
    const h=C45.headers.Documents;
    if(sh.getRange(1,1,1,h.length).getValues()[0].some((name,i)=>name!==h[i]))
      throw Error('Documents tab does not match v4.5 headers');
    c45BillingAddress_(); // Fail early if billing account property is missing.
    const props=PropertiesService.getScriptProperties(),last=sh.getLastRow();
    if(last<2){Logger.log('No document rows to repair.');return;}
    let start=Number(props.getProperty('V451_LINK_REPAIR_ROW')||2);
    if(!Number.isInteger(start)||start<2)start=2;
    if(start>last){props.deleteProperty('V451_LINK_REPAIR_ROW');Logger.log('Previous repair already complete.');return;}
    const count=Math.min(30,last-start+1),ids=sh.getRange(start,4,count,1).getValues();
    const existing=sh.getRange(start,24,count,1).getFormulas();
    const stats={rowsChecked:count,repaired:0,alreadyCorrect:0,missingMessageId:0,unavailableMessages:0,noRfcMessageId:0,fromRow:start,toRow:start+count-1};
    for(let i=0;i<count;i++){
      const msgId=String(ids[i][0]||'').trim();
      if(!msgId){stats.missingMessageId++;continue;}
      try{
        const msg=GmailApp.getMessageById(msgId);
        if(!msg){stats.unavailableMessages++;continue;}
        const formula=c45Link_(msg);
        if(!formula){stats.noRfcMessageId++;continue;}
        if(String(existing[i][0]||'')===formula){stats.alreadyCorrect++;continue;}
        sh.getRange(start+i,24).setFormula(formula);
        stats.repaired++;
      }catch(e){stats.unavailableMessages++;Logger.log('Row '+(start+i)+' could not be repaired: '+e.message);}
    }
    const next=start+count;
    if(next>last){props.deleteProperty('V451_LINK_REPAIR_ROW');stats.complete=true;}
    else{props.setProperty('V451_LINK_REPAIR_ROW',String(next));stats.complete=false;stats.nextRow=next;}
    Logger.log(JSON.stringify(stats));return stats;
  }finally{lock.releaseLock();}
}

// ---------- v4.5.2 controlled cost-line re-evaluation (30 existing lines/run) ----------
// This is a rules-based recheck of STORED extraction results, not a fresh OpenAI call.
// It never creates new expense rows, changes Cost_Impact, or promotes supporting
// documentation into booked costs. Questionable priced documents go to Review Queue.
function c452LineDecision_(st,line,doc){
  const prev=String(line.Rate_Eligible||'').trim().toUpperCase();
  // A human-selected NO is treated as a deliberate exclusion.
  if(prev==='NO')return {next:'NO',reason:'Manual exclusion preserved'};
  if(!doc)return {next:'NO_SUPPORTING',reason:'Parent document missing'};
  const status=String(doc.Status||''),type=String(doc.Doc_Type||'');
  if(status!=='COUNTED'||!['Invoice','Receipt'].includes(type))
    return {next:'NO_SUPPORTING',reason:'Parent is '+status+'/'+type};
  const cost=c45Money_(line.Line_Cost),conf=c45Num_(line.Confidence);
  if(cost===null||cost<0)return {next:'NO',reason:'No verifiable nonnegative line cost'};
  const rateFloor=Number(st.cfg.MIN_RATE_CONFIDENCE);
  if(conf===null||conf<rateFloor)return {next:'NO_CONFIDENCE',reason:'Line confidence below '+rateFloor};
  const lines=st.linesByDoc.get(String(doc.Document_ID))||[];
  const subtotal=c45Money_(doc.Subtotal);
  if(subtotal!==null&&lines.length){
    const sum=lines.reduce((total,x)=>total+(c45Money_(x.Line_Cost)||0),0);
    if(Math.abs(sum-subtotal)>Math.max(1,Math.abs(subtotal)*.01))
      return {next:'NO_MISMATCH',reason:'Itemized lines do not reconcile to subtotal'};
  }
  const manualQty=c45Num_(line.Manual_Qty),qty=manualQty>0?manualQty:c45Num_(line.Qty);
  const unit=c45Unit_(manualQty>0?line.Manual_Unit:line.Unit);
  if(!(qty>0)||!unit)return {next:'NO_QTY',reason:'No verified quantity and unit'};
  return {next:'YES',reason:'Verified counted purchase line with quantity'};
}

function c452RunCostLineBatch_(apply){
  const lock=LockService.getScriptLock();
  if(!lock.tryLock(1500))throw Error('Another billing operation is running; retry after it finishes');
  try{
    const st=c45State_(),props=PropertiesService.getScriptProperties(),lines=st['Cost Lines'].rows;
    const rateFloor=Number(st.cfg.MIN_RATE_CONFIDENCE);
    const classFloor=Number(st.cfg.MIN_CLASS_CONFIDENCE);
    if(!Number.isFinite(rateFloor)||rateFloor<0||rateFloor>1||!Number.isFinite(classFloor)||classFloor<0||classFloor>1)
      throw Error('Confidence settings must be decimals between 0 and 1 (for 80%, enter 0.80, not 80).');
    const saved=Number(props.getProperty('V452_COST_LINE_CURSOR')||0);
    const offset=Number.isInteger(saved)&&saved>=0&&saved<lines.length?saved:0;
    const selection=lines.slice(offset,offset+30);
    const result={mode:apply?'APPLY':'PREVIEW',version:'4.5.2',confidenceThreshold:rateFloor,totalLines:lines.length,startingAt:offset,examined:selection.length,changed:0,unchanged:0,pricedSupporting:0,heldReview:0,issuesCreated:0,manualExclusionsPreserved:0,examples:[]};
    const seenReview=new Set();
    for(const line of selection){
      const doc=st.docByID.get(String(line.Document_ID));
      const check=c452LineDecision_(st,line,doc);
      const before=String(line.Rate_Eligible||'').trim().toUpperCase();
      if(before==='NO')result.manualExclusionsPreserved++;
      if(before!==check.next){
        result.changed++;
        if(apply)c45Patch_(st['Cost Lines'],line,{Rate_Eligible:check.next});
        if(result.examples.length<12)result.examples.push({line:line.Line_ID,was:before,now:check.next,why:check.reason});
      }else result.unchanged++;
      if(!doc)continue;
      const docId=String(doc.Document_ID);
      const amount=c45Money_(line.Line_Cost);
      // Only one review entry per document per pass. A priced packing slip may be
      // useful price evidence, but not necessarily a new expense.
      if(doc.Status==='SUPPORTING'&&amount!==null&&!seenReview.has('SUPPORT:'+docId)){
        seenReview.add('SUPPORT:'+docId);result.pricedSupporting++;
        if(apply){
          const prior=st['Review Queue'].rows.length;
          c45Review_(st,docId,'Priced supporting document',
            'Contains documented line amounts; verify whether a distinct actual purchase or a supporting copy. Do not count twice.',
            'Compare original receipt/invoice and transaction references; reclassify only after confirming unique cost');
          if(st['Review Queue'].rows.length>prior)result.issuesCreated++;
        }
      }
      if(doc.Status==='HELD_REVIEW'&&Number(doc.Confidence||0)>=classFloor&&
         c45Money_(doc.Total)!==null&&String(doc.Vendor||'').trim()&&
         !seenReview.has('HELD:'+docId)){
        seenReview.add('HELD:'+docId);result.heldReview++;
        if(apply){
          const prior=st['Review Queue'].rows.length;
          c45Review_(st,docId,'Recheck held document',
            'Source now meets configured confidence/total/vendor checks; other conflicts may still apply.',
            'Review transaction matching before moving to COUNTED');
          if(st['Review Queue'].rows.length>prior)result.issuesCreated++;
        }
      }
    }
    const next=offset+selection.length;
    result.complete=next>=lines.length;
    result.nextOffset=result.complete?0:next;
    if(apply){
      // Cursor advances only after the sheet changes succeed; repeating a batch is safe.
      props.setProperty('V452_COST_LINE_CURSOR',String(result.nextOffset));
      c45RefreshRates_(st);
    }
    Logger.log(JSON.stringify(result));return result;
  }finally{lock.releaseLock();}
}

/** Inspect the next 30 existing Cost Lines without changing the workbook. */
function previewCostLinesBatchV452(){return c452RunCostLineBatch_(false);}

/** Recheck up to 30 stored cost lines, refresh Price History, and save a cursor.
 * Run repeatedly until the log says complete:true. No OpenAI calls or duplicate imports.
 */
function reevaluateCostLinesBatchV452(){return c452RunCostLineBatch_(true);}

/** Restart the 30-line review at the top of Cost Lines; never erases records. */
function restartCostLinesBatchV452(){
  PropertiesService.getScriptProperties().setProperty('V452_COST_LINE_CURSOR','0');
  Logger.log('Cost-line re-evaluation cursor reset to 0. Documents and prices were not deleted.');
}


// ---------- v4.5.3 focused document re-analysis / error diagnostics ----------
// Re-reads original sources for likely misclassified documents in small, restartable
// batches. Changes cost-impact ONLY for well-evidenced purchases with no related
// counted transaction. Priced confirmations remain price REFERENCES, never expenses.
const C453 = {version:'4.5.3',docsPerRun:6,rateMatchThreshold:0.75};

function diagnoseDocumentAnalysisErrorsV453(){
  const st=c45State_();
  const errors=st['Source Ledger'].rows.filter(x=>String(x.State).toUpperCase()==='ERROR');
  const reviews=st['Review Queue'].rows.filter(x=>String(x.Status).toUpperCase()==='OPEN'&&/analysis|processing|extract|openai|error/i.test(String(x.Issue_Type)+' '+String(x.Details)));
  const result={version:C453.version,sourceErrors:errors.length,analysisReviews:reviews.length,
    errors:errors.slice(0,12).map(x=>({source:x.Source_Key,filename:x.Filename,attempts:x.Attempts,error:String(x.Error||'').slice(0,380)})),
    reviews:reviews.slice(0,12).map(x=>({document:x.Document_ID,issue:x.Issue_Type,details:String(x.Details||'').slice(0,380)}))};
  Logger.log(JSON.stringify(result));return result;
}

function c453Candidates_(st){
  return st.Documents.rows.filter(d=>{
    if(['SUPPORTING','HELD_REVIEW','HELD_DUPLICATE'].includes(String(d.Status)))return (st.linesByDoc.get(String(d.Document_ID))||[]).some(l=>c45Money_(l.Line_Cost)!==null);
    return false;
  }).sort((a,b)=>String(a.Document_ID).localeCompare(String(b.Document_ID)));
}

function c453Source_(st,doc){
  const ledger=st.sourceByKey.get(String(doc.Source_Key));
  if(!ledger)throw Error('Source Ledger entry missing');
  const msg=GmailApp.getMessageById(String(doc.Message_ID||ledger.Message_ID));
  if(!msg)throw Error('Original Gmail message is unavailable');
  const sources=c45SourceItems_(msg,st.cfg);
  for(const src of sources){
    const bytes=src.blob?src.blob.getBytes():null;
    const hash=bytes?c45Hash_(bytes):c45Hash_((msg.getPlainBody()||'').slice(0,25000));
    const key='SRC-'+c45Hash_(msg.getId()+'|'+src.index+'|'+hash).slice(0,24);
    if(key===String(doc.Source_Key))return {msg,src};
  }
  throw Error('Original source attachment/body not found for '+doc.Source_Key+'; do not substitute another attachment');
}

function c453ProjectForDoc_(st,doc){
  const project=st.Projects.rows.find(p=>String(p.Project_ID)===String(doc.Project_ID)&&c45Bool_(p.Verified));
  if(!project)throw Error('Verified project no longer matches document');
  return project;
}

function c453Extract_(msg,src,project,cfg){
  const key=PropertiesService.getScriptProperties().getProperty('OPENAI_API_KEY');
  if(!key)throw Error('OPENAI_API_KEY missing');
  const prompt=[
    'Independently RE-READ ONE ORIGINAL source (email plus its named attachment) for construction cost evidence. Return JSON exactly matching the schema. Do not assume the previous category was correct.',
    'Separate purchase evidence from supporting logistics paperwork. An INVOICE demands payment for supplied goods or work, a PURCHASE RECEIPT records a completed sale. Delivery notices/packing slips/order confirmations/quotes can show quantities and prices WITHOUT establishing a new charge. A payment confirmation of an already billed invoice does not establish another charge. Classify based on original text, not the presence of amounts alone.',
    'Do NOT invent a charge, number, paid status, quantity, total, SKU, or job address. Set is_document=false only if there is no business document or meaningful price/quantity data.',
    'Extract ALL individual priced line items, including descriptions, extended line costs, actual units and quantities, brand and SKU. Keep tax, freight and discounts separate. For lump sum costs, extract one monetary line plus scope notes; do not make up a unit cost.',
    'Use a confidence between 0 and 1 reflecting evidence. Differentiate Document_Number, Invoice_Number, Order_Number, Receipt_Number and all reference IDs. If the same document appears to be a revised invoice, preserve the actual revised total; do not treat it as an incremental additional charge.',
    'The verified project is '+String(project.Project_Name||'')+' | '+String(project.Job_Address||'')+'. Do not use the vendor remit address as the project.',
    'Email From: '+msg.getFrom()+' | Subject: '+msg.getSubject()+' | Body: '+String(msg.getPlainBody()||'').slice(0,c45Int_(cfg.MAX_EMAIL_CHARS,12000,500,30000))
  ].join('\n\n');
  const content=[{type:'input_text',text:prompt}];
  if(src.blob){
    const bytes=src.blob.getBytes();
    if(bytes.length>c45Int_(cfg.MAX_ATTACHMENT_MB,12,1,25)*1024*1024)throw Error('Document too large to re-analyze');
    const url='data:'+(src.mime||'application/pdf')+';base64,'+Utilities.base64Encode(bytes);
    if(/^image\//.test(src.mime))content.push({type:'input_image',image_url:url});
    else if(src.mime==='application/pdf')content.push({type:'input_file',filename:src.name,file_data:url});
    else throw Error('Unsupported input file type for re-analysis: '+src.mime+' (not sent to OpenAI)');
  }
  const payload={model:String(cfg.OPENAI_MODEL||'gpt-4.1-mini'),input:[{role:'user',content}],text:{format:{type:'json_schema',name:'construction_cost_v453',strict:true,schema:c45Schema_()}}};
  const response=UrlFetchApp.fetch(C45.endpoint,{method:'post',contentType:'application/json',headers:{Authorization:'Bearer '+key},payload:JSON.stringify(payload),muteHttpExceptions:true});
  if(response.getResponseCode()<200||response.getResponseCode()>=300)throw Error('OpenAI HTTP '+response.getResponseCode()+': '+response.getContentText().slice(0,400));
  const body=JSON.parse(response.getContentText());let txt=body.output_text||'';
  (body.output||[]).forEach(o=>(o.content||[]).forEach(t=>{if(t.type==='output_text'&&t.text&&!txt)txt=t.text;}));
  if(!txt)throw Error('OpenAI analysis returned no structured output (status '+String(body.status||'unknown')+')');
  return JSON.parse(txt);
}

function c453LineText_(x){return c45Norm_([x.normalized_item||x.Normalized_Item,x.description||x.Source_Description,x.sku||x.SKU].join(' '));}
function c453LineCost_(x){return c45Money_(x.line_cost===undefined?x.Line_Cost:x.line_cost);}
function c453LineQty_(x){return c45Num_(x.quantity===undefined?x.Qty:x.quantity);}
function c453SameLine_(a,b){
  const amountA=c453LineCost_(a),amountB=c453LineCost_(b);
  if(amountA===null||amountB===null||Math.abs(amountA-amountB)>0.02)return false;
  const skuA=c45Norm_(a.sku||a.SKU),skuB=c45Norm_(b.sku||b.SKU);
  const qtyA=c453LineQty_(a),qtyB=c453LineQty_(b);
  if(qtyA!==null&&qtyB!==null&&Math.abs(qtyA-qtyB)>0.001)return false;
  if(skuA&&skuB&&skuA===skuB)return true;
  const nameA=c45Norm_(a.normalized_item||a.Normalized_Item),nameB=c45Norm_(b.normalized_item||b.Normalized_Item);
  if(nameA&&nameA===nameB)return true;
  return c45Jaccard_(c453LineText_(a),c453LineText_(b))>=C453.rateMatchThreshold;
}
function c453ChangedLine_(a,b){
  const skuA=c45Norm_(a.sku||a.SKU),skuB=c45Norm_(b.sku||b.SKU);
  if(skuA&&skuB&&skuA!==skuB)return false;
  const identity=(skuA&&skuB&&skuA===skuB)||c45Jaccard_(c453LineText_(a),c453LineText_(b))>=0.86;
  return !!identity&&!c453SameLine_(a,b);
}
function c453SharedRefs_(a,b){
  const ar=String(a.Refs||'').split(';').filter(Boolean),br=String(b.Refs||'').split(';').filter(Boolean);
  const av=new Set(ar.map(x=>x.split(':').slice(1).join(':')));
  return br.filter(x=>av.has(x.split(':').slice(1).join(':')));
}
function c453Relation_(st,doc,extracted){
  const vendor=c45Vendor_(extracted.vendor||doc.Vendor),proj=String(doc.Project_ID),number=c45Number_(extracted.invoice_number||extracted.receipt_number||extracted.document_number||doc.Document_No);
  const total=c45Money_(extracted.total),newLines=(extracted.items||[]).filter(x=>c453LineCost_(x)!==null);
  const others=st.Documents.rows.filter(o=>String(o.Document_ID)!==String(doc.Document_ID)&&o.Status==='COUNTED'&&c45Vendor_(o.Vendor)===vendor);
  let best={kind:'NONE',other:'',score:0};
  for(const old of others){
    const oldLines=st.linesByDoc.get(String(old.Document_ID))||[];
    const sameProject=String(old.Project_ID)===proj;
    const sameAmount=total!==null&&c45Money_(old.Total)!==null&&Math.abs(total-Number(old.Total))<0.03;
    const sameNumber=!!number&&number===c45Number_(old.Document_No);
    const sharedRefs=c453SharedRefs_(doc,old);
    const sharedInvoice=sharedRefs.some(x=>x.startsWith('INV:')||x.startsWith('REC:'));
    const sharedOrder=sharedRefs.some(x=>x.startsWith('ORD:'));
    const matched=newLines.length?newLines.filter(n=>oldLines.some(l=>c453SameLine_(n,l))).length:0;
    const overlap=newLines.length?matched/newLines.length:0;
    const score=(sharedInvoice||sameNumber?0.6:sharedOrder?0.3:0)+(sameAmount?0.2:0)+(overlap*0.2);
    // Identical products and prices alone are NOT proof two purchases are the same.
    const strongCopy=((sharedInvoice||sameNumber)&&sameAmount)||(sharedOrder&&sameAmount&&overlap>=.95);
    const newDate=new Date(extracted.issue_date||doc.Document_Date||doc.Recorded_At),oldDate=new Date(old.Document_Date||old.Recorded_At);
    const within14Days=Number.isFinite(newDate.getTime())&&Number.isFinite(oldDate.getTime())&&Math.abs(newDate-oldDate)<=14*86400000;
    const conflict=((sharedInvoice||sameNumber)&&!sameAmount)||(sharedOrder&&overlap>=.75&&!sameAmount)||(overlap>=.75&&within14Days);
    const acrossJobs=!sameProject&&(sameNumber||sharedInvoice)&&sameAmount;
    const result=acrossJobs?{kind:'CONFLICT',other:old.Document_ID,score}:sameProject&&strongCopy?{kind:'COPY',other:old.Document_ID,score}:sameProject&&conflict?{kind:'CONFLICT',other:old.Document_ID,score}:null;
    if(result&&(result.kind==='COPY'||best.kind==='NONE'))best=result;
    if(best.kind==='COPY')break;
  }
  return best;
}

function c453Comparison_(st,doc,data){
  const oldLines=st.linesByDoc.get(String(doc.Document_ID))||[];
  const incoming=(data.items||[]).filter(x=>c453LineCost_(x)!==null);
  const missing=[],changed=[];
  incoming.forEach(item=>{
    if(oldLines.some(o=>c453SameLine_(item,o)))return;
    if(oldLines.some(o=>c453ChangedLine_(item,o)))changed.push(item);
    else missing.push(item);
  });
  const stale=oldLines.filter(l=>!incoming.some(x=>c453SameLine_(l,x)||c453ChangedLine_(l,x)));
  const oldTotal=c45Money_(doc.Total),newTotal=c45Money_(data.total);
  return {missing,changed,stale,oldTotal,newTotal};
}

function c453AppendMissingLines_(st,doc,data,missing){
  const lines=st['Cost Lines'];let next=(st.linesByDoc.get(String(doc.Document_ID))||[]).length+1;const additions=[];
  for(const x of missing){
    const amount=c453LineCost_(x),qty=c453LineQty_(x),unit=c45Unit_(x.unit),confidence=Number(x.confidence||0);
    if(amount===null)continue;
    let id='';do{id=String(doc.Document_ID)+'-R'+String(next++).padStart(3,'0');}while(lines.rows.some(y=>String(y.Line_ID)===id));
    additions.push({Line_ID:id,Document_ID:doc.Document_ID,Project_ID:doc.Project_ID,Cost_Code:c45CostCode_(st,x),Normalized_Item:c45Safe_(x.normalized_item||'Unclassified'),Source_Description:c45Safe_(x.description),Cost_Type:c45CostType_(x.cost_type),Qty:qty,Unit:unit,Line_Cost:amount,Unit_Cost:qty>0?Math.round(amount/qty*1e4)/1e4:'',Rate_Eligible:'NO_SUPPORTING',Confidence:confidence,Brand_Spec:c45Safe_(x.brand_spec),SKU:c45Safe_(x.sku),Manual_Qty:'',Manual_Unit:'',Scope_Notes:'Re-extracted from original source in document review v4.5.3'});
  }
  const rows=c45Append_(lines,additions);
  if(!st.linesByDoc.has(String(doc.Document_ID)))st.linesByDoc.set(String(doc.Document_ID),[]);
  st.linesByDoc.get(String(doc.Document_ID)).push(...rows);
  return rows.length;
}

function c453ReviewOne_(st,doc){
  const {msg,src}=c453Source_(st,doc),project=c453ProjectForDoc_(st,doc),data=c453Extract_(msg,src,project,st.cfg);
  const type=c45Type_(data.document_type),conf=Number(data.confidence||0),total=c45Money_(data.total),cmp=c453Comparison_(st,doc,data);
  const result={document:doc.Document_ID,was:doc.Doc_Type+'/'+doc.Status,extractedType:type,confidence:conf,missingLines:cmp.missing.length,changedLines:cmp.changed.length,oldTotal:cmp.oldTotal,newTotal:cmp.newTotal,action:'UNCHANGED'};
  if(!data.is_document){result.action='NO_DOCUMENT';return result;}
  const relation=c453Relation_(st,doc,data);result.relation=relation.kind;result.related=relation.other;
  const strongPurchase=['Invoice','Receipt'].includes(type)&&conf>=Number(st.cfg.MIN_CLASS_CONFIDENCE||.8)&&total!==null&&total>=0&&c45Vendor_(data.vendor||doc.Vendor)&&String(data.invoice_number||data.receipt_number||data.document_number||'').trim();
  const addressConflict=!!(project.Job_Address&&data.job_address&&c45Address_(project.Job_Address)!==c45Address_(data.job_address));
  // Changing an already extracted amount might represent a correction, not a new charge.
  const uncertain=cmp.changed.length>0||cmp.stale.length>0||(cmp.oldTotal!==null&&total!==null&&Math.abs(cmp.oldTotal-total)>0.02);
  if(relation.kind==='COPY'){
    // Keep source data and all pricing for traceability, but never add duplicate expense.
    c45Patch_(st.Documents,doc,{Status:'COPY',Cost_Impact:0,Related_Doc_ID:relation.other,Notes:String(doc.Notes||'')+' | Reviewed duplicate of '+relation.other});
    result.action='MARKED_COPY';return result;
  }
  if(!strongPurchase||relation.kind==='CONFLICT'||uncertain||addressConflict){
    // Priced supporting records are retained as provisional price-only observations.
    if(cmp.missing.length&&!uncertain){result.newLines=c453AppendMissingLines_(st,doc,data,cmp.missing);}
    const reason=addressConflict?'Job address conflicts with verified project':!strongPurchase?'No conclusive invoice/receipt evidence':relation.kind==='CONFLICT'?'Possible conflicting transaction':'Changed/stale lines or revised total';
    if(relation.kind==='CONFLICT'||addressConflict||(strongPurchase&&uncertain))c45Review_(st,doc.Document_ID,'Document re-analysis conflict',reason+'; matched '+relation.other+'; check original source and whether this supersedes another charge','Verify transaction, do not sum revised invoice or repeat a charge');
    result.action=relation.kind==='CONFLICT'?'HELD_CONFLICT':(uncertain?'REVIEW_CHANGED':'PRICE_REFERENCE_ONLY');return result;
  }
  // Conclusive, independent purchase: add any genuinely new lines, then require total reconciliation.
  if(cmp.missing.length)result.newLines=c453AppendMissingLines_(st,doc,data,cmp.missing);
  const all=st.linesByDoc.get(String(doc.Document_ID))||[];
  const statedSubtotal=c45Money_(data.subtotal),sum=all.reduce((s,l)=>s+(c45Money_(l.Line_Cost)||0),0);
  if(statedSubtotal!==null&&all.length&&Math.abs(sum-statedSubtotal)>Math.max(1,Math.abs(statedSubtotal)*.01)){
    c45Review_(st,doc.Document_ID,'Document re-analysis mismatch','Re-evaluated item lines '+sum.toFixed(2)+' versus subtotal '+statedSubtotal.toFixed(2),'Inspect line entries; do not fabricate balancing items');
    result.action='REVIEW_SUBTOTAL_MISMATCH';return result;
  }
  const refs=c45Refs_(data),number=String(data.invoice_number||data.receipt_number||data.document_number||doc.Document_No).trim();
  c45Patch_(st.Documents,doc,{Doc_Type:type,Document_No:c45Safe_(number),Vendor:c45Safe_(data.vendor||doc.Vendor),Total:total,Subtotal:statedSubtotal,Tax:c45Money_(data.tax),Freight:c45Money_(data.freight),Discount:c45Money_(data.discount),Confidence:conf,Status:'COUNTED',Cost_Impact:Math.abs(total),Refs:refs.length?refs.join(';'):doc.Refs,Notes:String(doc.Notes||'')+' | Reclassified as purchase on original-document review v4.5.3'});
  all.forEach(line=>{const decision=c452LineDecision_(st,line,doc);if(String(line.Rate_Eligible)!=='NO'&&decision.next!==line.Rate_Eligible)c45Patch_(st['Cost Lines'],line,{Rate_Eligible:decision.next});});
  st['Review Queue'].rows.filter(r=>String(r.Document_ID)===String(doc.Document_ID)&&r.Status==='OPEN'&&['Priced supporting document','Recheck held document'].includes(String(r.Issue_Type))).forEach(r=>c45Patch_(st['Review Queue'],r,{Status:'RESOLVED'}));
  result.action='PROMOTED_PURCHASE';return result;
}

function c453ReviewBatch_(apply){
  const lock=LockService.getScriptLock();if(!lock.tryLock(1500))throw Error('Another billing run is active. Try again after it finishes.');
  try{
    const st=c45State_(),props=PropertiesService.getScriptProperties();
    // Cursor indexes ALL source Document rows, not the shrinking candidate list.
    // Promoted documents therefore cannot cause candidates later in the batch to be skipped.
    const all=st.Documents.rows.slice().sort((a,b)=>a._row-b._row);
    const savedRaw=String(props.getProperty('V453_DOC_CURSOR')||'0');
    const finished=savedRaw==='DONE';
    const saved=Number(savedRaw);
    const offset=finished?all.length:(Number.isInteger(saved)&&saved>=0&&saved<all.length?saved:0);
    const eligible=d=>['SUPPORTING','HELD_REVIEW','HELD_DUPLICATE'].includes(String(d.Status))&&(st.linesByDoc.get(String(d.Document_ID))||[]).some(l=>c45Money_(l.Line_Cost)!==null);
    const selected=[];let cursor=offset;
    for(;cursor<all.length&&selected.length<C453.docsPerRun;cursor++)if(eligible(all[cursor]))selected.push(all[cursor]);
    const result={version:C453.version,mode:apply?'APPLY':'PREVIEW',totalDocuments:all.length,candidateDocuments:all.filter(eligible).length,startingAt:offset,examined:selected.length,reviewed:0,promoted:0,copies:0,newLines:0,errors:0,items:[]};
    let nextIndex=selected.length?all.indexOf(selected[selected.length-1])+1:cursor;
    const started=Date.now();
    for(let i=0;i<selected.length;i++){
      const doc=selected[i];
      if(!apply){const priced=(st.linesByDoc.get(String(doc.Document_ID))||[]).filter(x=>c45Money_(x.Line_Cost)!==null);result.items.push({document:doc.Document_ID,status:doc.Status,type:doc.Doc_Type,pricedLines:priced.length,total:doc.Total});continue;}
      // Leave margin under the Apps Script execution limit.
      if(Date.now()-started>180000){nextIndex=all.indexOf(doc);break;}
      try{
        const item=c453ReviewOne_(st,doc);result.items.push(item);result.reviewed++;result.newLines+=item.newLines||0;if(item.action==='PROMOTED_PURCHASE')result.promoted++;if(item.action==='MARKED_COPY')result.copies++;
      }catch(e){result.errors++;const err=String(e.message||e).slice(0,350);result.items.push({document:doc.Document_ID,action:'ANALYSIS_ERROR',error:err});c45Review_(st,doc.Document_ID,'Document re-analysis error',err,'Inspect extraction error; check Gmail source link and API before retry');}
    }
    result.complete=nextIndex>=all.length;result.nextOffset=result.complete?0:nextIndex;
    if(apply){props.setProperty('V453_DOC_CURSOR',result.complete?'DONE':String(result.nextOffset));if(selected.length)c45RefreshRates_(st);}
    Logger.log(JSON.stringify(result));return result;
  }finally{lock.releaseLock();}
}
function previewDocumentReviewV453(){return c453ReviewBatch_(false);}
function reviewDocumentBatchV453(){return c453ReviewBatch_(true);}
function restartDocumentReviewV453(){PropertiesService.getScriptProperties().setProperty('V453_DOC_CURSOR','0');Logger.log('Document review restarted at cursor 0. No records erased.');}

