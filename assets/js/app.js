const APP_CONFIG = window.APP_CONFIG || {};
const $=id=>document.getElementById(id);let model=null;let generatedHtml='';
const DRAFT_KEY='kommunicateMappingDraftV1';
let mappingState={fieldMeta:{},history:[],activity:[]};
let aiState={suggestions:[],activeSuggestionId:null,lastRun:null};
let pendingSnippet=null;
let extractedSources=[];
let activeSourceId=null;
let lastPreviewSelection={text:'',start:0,end:0};
let pdfModulePromise=null;
const defaultDesign={headingFont:'Arial, Helvetica, sans-serif',bodyFont:'Arial, Helvetica, sans-serif',headlineColor:'#161616',headlineSize:31,bodySize:15,contentWidth:740,companyLogo:'',productLogos:{}};
let designState=JSON.parse(JSON.stringify(defaultDesign));
const demo={issue:{issue_title:'Compliance Newsletter',newsletter_name:'Kommunicate',publication_date:'2026-07-15',volume_number:'06',issue_number:'13',headline:'The social security framework enters a new compliance phase',introduction:'This edition brings together the most important legal and regulatory developments affecting businesses across India.',executive_summary:'The principal theme is the transition to new social security schemes, digital reporting requirements and time-bound regularisation windows.',company_name:'LEXPLOSION',company_tagline:'Innovating Legally',contact_email:'contact@lexplosion.in'},updates:[{display_order:1,category:'Labour & Employment',jurisdiction:'India',regulator:'Ministry of Labour and Employment',short_headline:'New Provident Fund Scheme notified',summary:'The 2026 scheme replaces the legacy framework and introduces electronic reporting, revised disclosures and digital member onboarding.',effective_date:'29 June 2026',compliance_deadline:'14 July 2026',risk_level:'Critical',required_action:'Review payroll, member registration and statutory reporting processes.'},{display_order:2,category:'Corporate Law',jurisdiction:'India',regulator:'Ministry of Corporate Affairs',short_headline:'Compliance facilitation window extended',summary:'The validity of the Companies Compliance Facilitation Scheme has been extended to allow completion of pending statutory filings.',effective_date:'15 July 2026',compliance_deadline:'31 August 2026',risk_level:'High',required_action:'Identify overdue filings and complete submissions before the revised deadline.'},{display_order:3,category:'State Regulation',jurisdiction:'Andhra Pradesh',regulator:'Government of Andhra Pradesh',short_headline:'State social security rules brought into force',summary:'The new state rules introduce gratuity timelines, healthcare-related requirements and mandatory creche obligations for qualifying establishments.',effective_date:'7 July 2026',risk_level:'Medium',required_action:'Validate establishment-level applicability and update local compliance calendars.'}],insights:[{title:'Understanding the Social Security Framework of 2026',summary:'A practical guide to the transition windows, employer obligations and legal exposure.',url:'#',cta_text:'Read analysis',full_analysis:'The new framework changes how employers enrol employees, submit statutory information and manage regularisation windows. Compliance teams should assess payroll configuration, employee master data, filing ownership and evidence retention before the transition deadlines.',key_implications:'Digital reporting becomes the default | Historic omissions may be regularised within defined windows | Governance teams need evidence of timely implementation',affected_entities:'Employers, payroll teams, legal teams and compliance officers',recommended_actions:'Map the new requirements to existing controls | Assign accountable owners | Validate employee and contribution data | Track each transition deadline',important_dates:'29 June 2026: scheme notified | 31 October 2026: enrolment campaign closes',reading_time:'4 min read'},{title:'Why compliance calendar visibility matters',summary:'How proactive visibility reduces operational risk and supports executive accountability.',url:'#',cta_text:'Explore the article',full_analysis:'A central compliance calendar gives leadership one view of obligations, deadlines, evidence and ownership across entities. It reduces dependence on spreadsheets and helps teams identify overdue or unsupported compliance before it becomes an audit issue.',key_implications:'Improved executive visibility | Earlier escalation of missed obligations | Stronger evidence readiness',affected_entities:'Compliance leaders, legal teams, business owners and the C-suite',recommended_actions:'Consolidate regulatory calendars | Define escalation thresholds | Link evidence to each obligation',reading_time:'3 min read'}],promotions:[{product_name:'Komrisk AI',headline:'Move from fragmented tracking to continuous compliance visibility',description:'Bring obligations, evidence, alerts and reporting into one configurable compliance operating layer.',cta_text:'View product details',cta_url:'#',capabilities:'Regulatory obligation tracking | Evidence validation | Compliance calendars | Alerts and escalations | Management reporting',use_cases:'Multi-entity compliance operations | Enterprise evidence management | Executive compliance visibility',target_users:'Compliance, legal, risk and internal audit teams',contact_email:'contact@lexplosion.in'},{product_name:'Komtrakt',headline:'Protect contract obligations as regulations change',description:'Automate template drafting, obligation abstraction and contract lifecycle controls.',cta_text:'View product details',cta_url:'#',capabilities:'Template drafting | Obligation abstraction | Contract review workflows | Renewal and milestone tracking',use_cases:'Vendor agreements | Contract compliance | Regulatory clause management',target_users:'Legal, procurement and contract management teams',contact_email:'contact@lexplosion.in'},{product_name:'Komtrol',headline:'Strengthen disclosure and regulatory event monitoring',description:'Track material events, media and price movements and support disclosure workflows.',cta_text:'View product details',cta_url:'#',capabilities:'Material event monitoring | Media and price movement alerts | Disclosure workflow support | Audit-ready records',use_cases:'SEBI LODR monitoring | Disclosure governance | Regulatory event tracking',target_users:'Company secretarial, legal and compliance teams',contact_email:'contact@lexplosion.in'}]};
function log(msg,type=''){const t=new Date().toLocaleTimeString();$('log').innerHTML=`<div>${t} — ${msg}</div>`+$('log').innerHTML;$('status').textContent=msg;if(type)$('status').className='status '+type}
function isHex(v){return /^#[0-9a-f]{6}$/i.test(String(v||'').trim())}
function refreshPreview(message='Preview updated'){if(!model)return;generatedHtml=renderNewsletter(model);const preview=$('preview');preview.style.width='100%';preview.style.maxWidth=designState.contentWidth+'px';preview.innerHTML=generatedHtml;$('downloadBtn').disabled=$('downloadDocxBtn').disabled=$('copyBtn').disabled=$('emailBtn').disabled=false;$('status').textContent=message}
function bindColour(pickerId,hexId){const picker=$(pickerId),hex=$(hexId);picker.addEventListener('input',()=>{hex.value=picker.value.toUpperCase();refreshPreview('Colours updated')});hex.addEventListener('input',()=>{const value=hex.value.trim();if(isHex(value)){picker.value=value;refreshPreview('Colours updated')}});hex.addEventListener('blur',()=>{if(!isHex(hex.value)){hex.value=picker.value.toUpperCase();log('Enter colours in #RRGGBB format','warn')}})}
function normaliseRow(row){const out={};Object.entries(row||{}).forEach(([k,v])=>out[String(k).trim().toLowerCase().replace(/[^a-z0-9]+/g,'_')]=v);return out}
async function parseExcel(file){const b=await file.arrayBuffer(),wb=XLSX.read(b,{type:'array',cellDates:true});const byName=n=>{const key=wb.SheetNames.find(x=>x.toLowerCase().trim()===n.toLowerCase());return key?XLSX.utils.sheet_to_json(wb.Sheets[key],{defval:''}).map(normaliseRow):[]};const issueRows=byName('Issue');let issue={};if(issueRows.length===1)issue=issueRows[0];else issueRows.forEach(r=>{const k=r.field||r.key||r.name,v=r.value||r.content;if(k)issue[String(k).trim().toLowerCase().replace(/[^a-z0-9]+/g,'_')]=v});return{issue,updates:byName('Regulatory Updates'),insights:byName('Insights'),promotions:byName('Promotions')}}
async function parseWord(file){const b=await file.arrayBuffer(),r=await mammoth.extractRawText({arrayBuffer:b}),lines=r.value.split(/\n+/).map(x=>x.trim()).filter(Boolean),sections={issue:{},updates:[],insights:[],promotions:[]};let bucket='issue',current=null;lines.forEach(line=>{const l=line.toLowerCase();if(l==='regulatory updates'){bucket='updates';current=null;return}if(l==='expert insights'||l==='insights'){bucket='insights';current=null;return}if(l==='products and services'||l==='promotions'){bucket='promotions';current=null;return}const m=line.match(/^([^:]+):\s*(.*)$/);if(m){const key=m[1].trim().toLowerCase().replace(/[^a-z0-9]+/g,'_'),val=m[2].trim();if(bucket==='issue')sections.issue[key]=val;else{if(!current){current={};sections[bucket].push(current)}if((key==='title'||key==='short_headline'||key==='product_name')&&Object.keys(current).length){current={};sections[bucket].push(current)}current[key]=val}}else if(bucket==='issue'&&!sections.issue.introduction)sections.issue.introduction=line});return sections}
function uid(prefix='id'){return prefix+'_'+Date.now().toString(36)+'_'+Math.random().toString(36).slice(2,7)}
function ensureModelIds(){if(!model)return;model.issue=model.issue||{};model.updates=model.updates||[];model.updates.forEach((u,i)=>{if(!u.id)u.id=uid('update');if(!u.display_order)u.display_order=i+1})}
const ISSUE_FIELDS=[['headline','Lead story → Headline'],['introduction','Lead story → Introduction'],['executive_summary','Lead story → Executive summary']];
const UPDATE_FIELDS=[['short_headline','Headline'],['summary','Summary'],['category','Category'],['jurisdiction','Jurisdiction'],['regulator','Regulator'],['effective_date','Effective date'],['compliance_deadline','Compliance deadline'],['risk_level','Risk level'],['required_action','Required action']];
function updateById(id){return model?.updates?.find(x=>x.id===id)}
function targetPathForIssue(field){return `issue.${field}`}
function targetPathForUpdate(id,field){return `updates.${id}.${field}`}
function getTargetValue(path){if(!model)return'';const parts=path.split('.');if(parts[0]==='issue')return model.issue?.[parts[1]]||'';if(parts[0]==='updates')return updateById(parts[1])?.[parts[2]]||'';return''}
function setTargetValue(path,value){const parts=path.split('.');if(parts[0]==='issue'){model.issue=model.issue||{};model.issue[parts[1]]=value;return}if(parts[0]==='updates'){const u=updateById(parts[1]);if(u)u[parts[2]]=value}}
function targetLabel(path){const parts=path.split('.');if(parts[0]==='issue')return ISSUE_FIELDS.find(x=>x[0]===parts[1])?.[1]||path;const u=model?.updates?.find(x=>x.id===parts[1]);const n=Math.max(1,(model?.updates||[]).indexOf(u)+1);const label=UPDATE_FIELDS.find(x=>x[0]===parts[2])?.[1]||parts[2];return `Regulatory update ${n} → ${label}`}
function metaFor(path){return mappingState.fieldMeta[path]||null}
function sourceBadge(path){const m=metaFor(path);if(!m)return'';const cls=m.origin==='excel'?'excel':m.origin==='document'?'document':m.origin==='manual'?'manual':'';const text=m.origin==='excel'?'Excel':m.origin==='document'?(m.sourceName||'Document'):m.origin==='manual'?'Manual':m.origin;return `<span class="source-badge ${cls}">${esc(text)}</span>`}
function initialiseFieldMetadata(origin='excel',sourceName='Structured file'){ensureModelIds();if(!model)return;ISSUE_FIELDS.forEach(([f])=>{const p=targetPathForIssue(f);if(getTargetValue(p)&&!mappingState.fieldMeta[p])mappingState.fieldMeta[p]={origin,sourceName,updatedAt:new Date().toISOString()}});model.updates.forEach(u=>UPDATE_FIELDS.forEach(([f])=>{const p=targetPathForUpdate(u.id,f);if(getTargetValue(p)&&!mappingState.fieldMeta[p])mappingState.fieldMeta[p]={origin,sourceName,updatedAt:new Date().toISOString()}}))}
function saveDraft(silent=true){try{ensureModelIds();localStorage.setItem(DRAFT_KEY,JSON.stringify({model,extractedSources,mappingState,aiState,designState,savedAt:new Date().toISOString()}));if(!silent)log('Draft saved locally','ok')}catch(e){if(!silent)log('Draft could not be saved: '+e.message,'error')}}
function restoreDraft(){try{const raw=localStorage.getItem(DRAFT_KEY);if(!raw)return false;const d=JSON.parse(raw);if(!d.model)return false;model=d.model;extractedSources=d.extractedSources||[];mappingState=d.mappingState||{fieldMeta:{},history:[],activity:[]};aiState=d.aiState||{suggestions:[],activeSuggestionId:null,lastRun:null};if(d.designState)designState={...designState,...d.designState};ensureModelIds();return true}catch(e){return false}}
function recordManualMeta(path){const old=metaFor(path);mappingState.fieldMeta[path]={origin:'manual',sourceName:'Manual edit',updatedAt:new Date().toISOString(),previousOrigin:old?.origin||null}}
function renderMappingActivity(){const el=$('mappingList');if(!el)return;const items=mappingState.activity||[];$('mappingCount').textContent=`${items.length} mapped`;$('undoMappingBtn').disabled=!mappingState.history.length;if(!items.length){el.innerHTML='<div class="empty-source">No document content has been mapped yet.</div>';return}el.innerHTML=items.slice().reverse().slice(0,8).map(x=>`<div class="mapping-item"><strong>${esc(x.targetLabel)}</strong><small>${esc(x.sourceName)} · ${esc(x.mode)} · ${new Date(x.appliedAt).toLocaleTimeString()}</small></div>`).join('')}
function mappingTargets(){const opts=[...ISSUE_FIELDS.map(([f,l])=>({value:targetPathForIssue(f),label:l}))];(model?.updates||[]).forEach((u,i)=>UPDATE_FIELDS.forEach(([f,l])=>opts.push({value:targetPathForUpdate(u.id,f),label:`Regulatory update ${i+1} → ${l}`})));opts.push({value:'__new_update__',label:'Create new regulatory update from selection'});return opts}
const FIELD_RULES={
  'issue.headline':{label:'Lead story headline',recommended:90,max:160,kind:'headline'},
  'issue.introduction':{label:'Introduction',recommended:420,max:900,kind:'body'},
  'issue.executive_summary':{label:'Executive summary',recommended:500,max:1200,kind:'body'},
  short_headline:{label:'Regulatory-update headline',recommended:90,max:180,kind:'headline'},
  summary:{label:'Summary',recommended:500,max:1200,kind:'body'},
  category:{label:'Category',recommended:40,max:80,kind:'short'},
  jurisdiction:{label:'Jurisdiction',recommended:40,max:100,kind:'short'},
  regulator:{label:'Regulator',recommended:80,max:180,kind:'short'},
  effective_date:{label:'Effective date',recommended:30,max:80,kind:'short'},
  compliance_deadline:{label:'Compliance deadline',recommended:40,max:100,kind:'short'},
  risk_level:{label:'Risk level',recommended:10,max:30,kind:'short'},
  required_action:{label:'Required action',recommended:240,max:650,kind:'body'}
};
function fieldRule(path){if(path==='__new_update__')return{label:'New regulatory update',recommended:600,max:1800,kind:'new'};if(FIELD_RULES[path])return FIELD_RULES[path];const field=path.split('.').pop();return FIELD_RULES[field]||{label:'Newsletter field',recommended:500,max:1600,kind:'body'}}
function cleanSourceText(text){return String(text||'').replace(/^SOURCE\s*\d*\s*[:.-]?\s*/i,'').replace(/[-=]{8,}/g,' ').replace(/\s+/g,' ').trim()}
function suggestedForField(text,path){const clean=cleanSourceText(text),rule=fieldRule(path);if(!clean)return'';if(rule.kind==='headline'||rule.kind==='short'){let first=clean.split(/(?<=[.!?])\s+/)[0]||clean;first=first.replace(/[.!?]+$/,'').trim();if(first.length>rule.recommended){const cut=first.slice(0,rule.recommended+1);first=cut.slice(0,Math.max(cut.lastIndexOf(' '),40)).trim()+'…'}return first}if(rule.kind==='new')return clean.slice(0,rule.max);return clean.slice(0,rule.max)}
function updateMappingGuidance(){const path=$('mappingTarget').value,rule=fieldRule(path),value=$('mappingProposed').value.trim(),count=value.length;const counter=$('mappingCounter'),fit=$('fitMappingBtn');counter.textContent=`${count} characters · recommended up to ${rule.recommended} · maximum ${rule.max}`;counter.className='mapping-counter '+(count>rule.max?'over':count>rule.recommended?'warn':'ok');fit.textContent=rule.kind==='headline'||rule.kind==='short'?'Create concise field text':'Fit text to field';fit.classList.toggle('hidden',!pendingSnippet);if(count>rule.max)$('mappingStatus').textContent=`This is too long for ${rule.label}. Shorten it before applying.`;else if($('mappingStatus').textContent.startsWith('This is too long'))$('mappingStatus').textContent=''}
function safeShowDialog(dialog){if(!dialog)return false;try{document.querySelectorAll('dialog[open]').forEach(d=>{if(d!==dialog)try{d.close()}catch(e){d.removeAttribute('open')}});if(typeof dialog.showModal==='function'){dialog.showModal()}else dialog.setAttribute('open','');return true}catch(e){dialog.setAttribute('open','');log('The mapping window was opened in compatibility mode','warn');return true}}
function openMappingDialog(snippet,source){if(!snippet.trim())return log('Select a source passage before mapping','warn');if(!model){model={issue:{issue_title:'Compliance Newsletter',newsletter_name:'Kommunicate'},updates:[],insights:[],promotions:[]};ensureModelIds();show()}pendingSnippet={text:snippet.trim(),sourceId:source?.id||'combined',sourceName:source?.name||'Combined sources'};const select=$('mappingTarget');select.innerHTML=mappingTargets().map(o=>`<option value="${esc(o.value)}">${esc(o.label)}</option>`).join('');$('mappingProposed').value=suggestedForField(pendingSnippet.text,select.value);$('mappingSourceRef').textContent=`Source: ${pendingSnippet.sourceName}`;$('mappingStatus').textContent='';updateMappingComparison();updateMappingGuidance();safeShowDialog($('mappingDialog'))}
function updateMappingComparison(){let path=$('mappingTarget').value;if(path==='__new_update__'){$('mappingCurrent').textContent='A new regulatory update will be created. The selected passage will become its summary and a concise headline will be generated from its opening sentence.';$('mappingConflict').classList.add('hidden');$('mappingMode').value='replace';$('mappingMode').disabled=true;updateMappingGuidance();return}$('mappingMode').disabled=false;const current=getTargetValue(path);$('mappingCurrent').textContent=current||'Empty';const meta=metaFor(path);if(current){$('mappingConflict').classList.remove('hidden');$('mappingConflict').innerHTML=`This field already contains content${meta?` from <b>${esc(meta.sourceName||meta.origin)}</b>`:''}. Review the comparison before replacing or appending.`}else $('mappingConflict').classList.add('hidden');updateMappingGuidance()}
function applyPendingMapping(){if(!pendingSnippet||!model)return;let path=$('mappingTarget').value;let proposed=$('mappingProposed').value.trim();if(!proposed)return $('mappingStatus').textContent='Enter proposed content before applying.';const rule=fieldRule(path);if(proposed.length>rule.max)return $('mappingStatus').textContent=`This content is ${proposed.length} characters. ${rule.label} allows a maximum of ${rule.max}. Use “${$('fitMappingBtn').textContent}” or edit the text.`;if(path==='__new_update__'){const sourceText=cleanSourceText(pendingSnippet.text);const u={id:uid('update'),display_order:model.updates.length+1,category:'Regulatory update',short_headline:suggestedForField(proposed,'updates.new.short_headline'),summary:sourceText.slice(0,FIELD_RULES.summary.max),risk_level:'Medium'};model.updates.push(u);path=targetPathForUpdate(u.id,'short_headline');mappingState.fieldMeta[targetPathForUpdate(u.id,'summary')]={origin:'document',sourceId:pendingSnippet.sourceId,sourceName:pendingSnippet.sourceName,updatedAt:new Date().toISOString(),overrode:false}}else{const current=getTargetValue(path);const mode=$('mappingMode').value;const next=mode==='append'&&current?`${current}\n\n${proposed}`:proposed;const oldMeta=metaFor(path);mappingState.history.push({path,oldValue:current,oldMeta:oldMeta?JSON.parse(JSON.stringify(oldMeta)):null,newValue:next,at:new Date().toISOString()});setTargetValue(path,next);mappingState.fieldMeta[path]={origin:'document',sourceId:pendingSnippet.sourceId,sourceName:pendingSnippet.sourceName,updatedAt:new Date().toISOString(),overrode:Boolean(current)};mappingState.activity.push({path,targetLabel:targetLabel(path),sourceId:pendingSnippet.sourceId,sourceName:pendingSnippet.sourceName,mode:current?mode:'fill',appliedAt:new Date().toISOString()})}saveDraft();show();renderSourceList();renderMappingActivity();$('mappingDialog').close();log(`${targetLabel(path)} mapped from ${pendingSnippet.sourceName}`,'ok');pendingSnippet=null}
function undoLastMapping(){const h=mappingState.history.pop();if(!h)return;setTargetValue(h.path,h.oldValue);if(h.oldMeta)mappingState.fieldMeta[h.path]=h.oldMeta;else delete mappingState.fieldMeta[h.path];mappingState.activity.pop();saveDraft();show();renderMappingActivity();log(`Restored previous value for ${targetLabel(h.path)}`,'ok')}
function validate(m){const errors=[];if(!m.issue||!Object.keys(m.issue).length)errors.push('Issue sheet or issue metadata is missing.');if(!m.updates||!m.updates.length)errors.push('No regulatory updates were found.');m.updates?.forEach((u,i)=>{if(!(u.short_headline||u.regulation_title||u.title))errors.push(`Update ${i+1}: headline missing`);if(!u.summary)errors.push(`Update ${i+1}: summary missing`)});return errors}
function esc(s=''){return String(s).replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]))}
function formatDate(v){if(!v)return'';const d=new Date(v);return isNaN(d)?v:d.toLocaleDateString('en-GB',{day:'numeric',month:'long',year:'numeric'})}
function brandWord(name,sub='',titleColor='#151515',subColor='#757575'){return `<table role="presentation" cellspacing="0" cellpadding="0"><tr><td style="vertical-align:middle;padding-right:8px"><table role="presentation" cellspacing="2" cellpadding="0"><tr><td style="width:4px;height:18px;background:#00a6a6;border-radius:2px"></td><td style="width:4px;height:26px;background:#f58220;border-radius:2px"></td><td style="width:4px;height:13px;background:#f7b733;border-radius:2px"></td></tr></table></td><td><div style="font-weight:900;letter-spacing:.04em;font-size:14px;color:${titleColor}">${esc(name)}</div>${sub?`<div style="font-size:9px;color:${subColor};letter-spacing:.05em">${esc(sub)}</div>`:''}</td></tr></table>`}
function productLogo(name){if(designState.productLogos?.[name])return `<img src="${designState.productLogos[name]}" alt="${esc(name)} logo" style="display:block;max-width:190px;max-height:52px;object-fit:contain">`;const map={"Komrisk AI":['KOMRISK','AI-enabled compliance management'],"Komtrakt":['KOMTRAKT','Contract lifecycle intelligence'],"Komtrol":['KOMTROL','Disclosure and control monitoring'],"Komtrol Plus":['KOMTROL PLUS','Advanced disclosure management']};const x=map[name]||[name,'Lexplosion product'];return `<table role="presentation" cellspacing="0" cellpadding="0"><tr><td style="vertical-align:middle;padding-right:10px"><table role="presentation" cellspacing="2" cellpadding="0"><tr><td style="width:4px;height:21px;background:#00a6a6;border-radius:2px"></td><td style="width:4px;height:28px;background:#f58220;border-radius:2px"></td><td style="width:4px;height:16px;background:#f7b733;border-radius:2px"></td></tr></table></td><td><div style="font-weight:900;letter-spacing:.045em;font-size:16px;color:#141414;line-height:1">${esc(x[0])}</div><div style="font-size:10px;color:#6b7280;letter-spacing:.035em;margin-top:3px">${esc(x[1])}</div></td></tr></table>`}
function listItems(value){return String(value||'').split(/\s*\|\s*|\n+/).map(x=>x.trim()).filter(Boolean)}
function bulletList(value){const items=listItems(value);return items.length?`<ul style="margin:8px 0 0;padding-left:18px;color:#4f5b66;font-size:13px;line-height:1.65">${items.map(v=>`<li style="margin:3px 0">${esc(v)}</li>`).join('')}</ul>`:''}
function detailRow(label,value){return value?`<div style="margin:12px 0"><div style="font-size:10px;letter-spacing:.08em;text-transform:uppercase;font-weight:800;color:#476466">${esc(label)}</div><div style="margin-top:5px;font-size:13px;line-height:1.65;color:#4f5b66">${esc(value)}</div></div>`:''}
function expandable(label,body,p,a){return `<details class="kn-expand" style="margin-top:14px;border-top:1px solid #e6ebe8;padding-top:12px"><summary style="list-style:none;cursor:pointer;display:flex;align-items:center;justify-content:space-between;gap:12px;color:${p};font-size:12px;font-weight:800;outline:none"><span>${esc(label)}</span><span class="kn-chevron" style="width:24px;height:24px;border-radius:50%;background:${a}18;color:${a};display:inline-flex;align-items:center;justify-content:center;font-size:16px;line-height:1">⌄</span></summary><div class="kn-detail" style="margin-top:12px;background:#eef8f6;border-left:4px solid ${p};border-radius:0 12px 12px 0;padding:16px 18px">${body}</div></details>`}
function updateDetail(u,p,a){const body=`${detailRow('What changed',u.what_changed||u.summary)}${detailRow('Why it matters',u.why_it_matters||u.business_impact)}${detailRow('Who is affected',u.affected_entities)}${detailRow('Required action',u.required_action)}${detailRow('Official source',u.official_source_url||u.source_url)}`;return expandable('View full update',body||'<div style="font-size:13px;color:#5f6874">No additional details were provided.</div>',p,a)}
function insightDetail(x,p,a){const body=`${x.full_analysis?`<div style="font-size:13px;line-height:1.7;color:#4f5b66">${esc(x.full_analysis)}</div>`:''}${x.key_implications?`<div style="margin-top:14px"><div style="font-size:10px;letter-spacing:.08em;text-transform:uppercase;font-weight:800;color:#476466">Key implications</div>${bulletList(x.key_implications)}</div>`:''}${detailRow('Who is affected',x.affected_entities)}${x.recommended_actions?`<div style="margin-top:14px"><div style="font-size:10px;letter-spacing:.08em;text-transform:uppercase;font-weight:800;color:#476466">Recommended actions</div>${bulletList(x.recommended_actions)}</div>`:''}${x.important_dates?`<div style="margin-top:14px"><div style="font-size:10px;letter-spacing:.08em;text-transform:uppercase;font-weight:800;color:#476466">Important dates</div>${bulletList(x.important_dates)}</div>`:''}${x.source_url&&x.source_url!=='#'?`<a href="${esc(x.source_url)}" target="_blank" rel="noopener" style="display:inline-block;margin-top:14px;color:${p};font-size:12px;font-weight:800;text-decoration:none">Open official source ↗</a>`:''}`;return expandable(x.cta_text||'Read analysis',body||`<div style="font-size:13px;line-height:1.65;color:#4f5b66">${esc(x.summary||'No further analysis was provided.')}</div>`,p,a)}
function productDetail(x,p,a,contact){const body=`${x.capabilities?`<div><div style="font-size:10px;letter-spacing:.08em;text-transform:uppercase;font-weight:800;color:#476466">Key capabilities</div>${bulletList(x.capabilities)}</div>`:''}${x.use_cases?`<div style="margin-top:14px"><div style="font-size:10px;letter-spacing:.08em;text-transform:uppercase;font-weight:800;color:#476466">Relevant use cases</div>${bulletList(x.use_cases)}</div>`:''}${detailRow('Best suited for',x.target_users)}<a href="mailto:${esc(x.contact_email||contact||'contact@lexplosion.in')}" style="display:inline-block;margin-top:14px;background:${p};color:#fff;border-radius:9px;padding:10px 14px;font-size:12px;font-weight:800;text-decoration:none">Contact Lexplosion</a>`;return expandable(x.cta_text||'View product details',body||`<div style="font-size:13px;line-height:1.65;color:#4f5b66">${esc(x.description||'Contact Lexplosion for more information.')}</div>`,p,a)}
function renderNewsletter(m){const p=$('primaryColor').value,a=$('accentColor').value,i=m.issue||{},hf=designState.headingFont,bf=designState.bodyFont,hc=designState.headlineColor,hs=designState.headlineSize,bs=designState.bodySize,cw=designState.contentWidth,updates=[...(m.updates||[])].sort((x,y)=>(+x.display_order||999)-(+y.display_order||999)),title=$('issueTitle').value||i.issue_title||'Compliance Newsletter',promos=m.promotions||[];return `<!DOCTYPE html><html><head><meta charset="UTF-8"><meta name="viewport" content="width=device-width,initial-scale=1"><title>${esc(title)}</title><style>details>summary::-webkit-details-marker{display:none}.kn-expand[open] .kn-chevron{transform:rotate(180deg)}.kn-chevron{transition:transform .2s ease}.kn-detail{animation:knOpen .2s ease}@keyframes knOpen{from{opacity:0;transform:translateY(-4px)}to{opacity:1;transform:none}}.kn-newsletter,.kn-newsletter table,.kn-newsletter td,.kn-newsletter div,.kn-newsletter p,.kn-newsletter span,.kn-newsletter a,.kn-newsletter summary,.kn-newsletter li{font-family:${bf}!important}.kn-newsletter h1,.kn-newsletter h2,.kn-newsletter h3{font-family:${hf}!important}@media(max-width:600px){.kn-detail{padding:14px!important}}</style></head><body style="margin:0;background:#f1f2ef;font-family:${bf};color:#161616"><table class="kn-newsletter" role="presentation" width="100%" cellspacing="0" cellpadding="0" style="width:100%;background:#f1f2ef;font-family:${bf}"><tr><td align="center" style="padding:28px 12px;font-family:${bf}"><table role="presentation" width="${cw}" cellspacing="0" cellpadding="0" style="width:${cw}px;max-width:100%;background:white;border-radius:22px;overflow:hidden;box-shadow:0 20px 60px rgba(23,63,66,.13);font-family:${bf}">
<tr><td style="padding:20px 40px;border-bottom:1px solid #ecece8;background:#ffffff"><table role="presentation" width="100%"><tr><td>${designState.companyLogo?`<img src="${designState.companyLogo}" alt="Lexplosion logo" style="display:block;max-width:210px;max-height:58px;object-fit:contain">`:brandWord('LEXPLOSION','Innovating Legally')}</td><td align="right" style="font-size:10px;letter-spacing:.12em;text-transform:uppercase;color:#526466;font-weight:800">Kommunicate<br><span style="display:inline-block;margin-top:4px;color:${a};letter-spacing:.08em">Fortnightly Regulatory Update</span></td></tr></table></td></tr>
<tr><td style="background:linear-gradient(135deg,${p} 0%,#0d6969 100%);padding:42px;color:white;position:relative"><div style="display:inline-block;background:${a};padding:7px 12px;border-radius:999px;font-size:10px;font-weight:800;letter-spacing:.12em;text-transform:uppercase">Fortnightly regulatory update</div><table role="presentation" width="100%" style="margin-top:24px"><tr><td style="vertical-align:bottom"><div style="font-family:${hf};font-size:48px;line-height:.95;font-weight:900;letter-spacing:-.045em">Kommunicate<span style="color:${a}">.</span></div><div style="font-size:12px;letter-spacing:.13em;text-transform:uppercase;margin-top:10px;opacity:.82">Lexplosion's regulatory compliance newsletter</div></td><td align="right" style="vertical-align:bottom"><div style="font-size:13px;opacity:.9">${esc(formatDate(i.publication_date))}</div><div style="font-size:12px;opacity:.75;margin-top:5px">Vol ${esc(i.volume_number||'')} · Issue ${esc(i.issue_number||'')}</div></td></tr></table></td></tr>
<tr><td style="padding:34px 42px 10px"><div style="font-size:11px;color:${a};font-weight:800;text-transform:uppercase;letter-spacing:.1em">The lead story</div><h1 style="font-family:${hf};font-size:${hs}px;color:${hc};line-height:1.15;margin:9px 0 14px;letter-spacing:-.02em">${esc(i.headline||title)}</h1><p style="font-size:${bs+1}px;line-height:1.75;color:#4b5563;margin:0 0 12px">${esc(i.introduction||'')}</p><p style="font-size:${bs}px;line-height:1.7;color:#4b5563;margin:0">${esc(i.executive_summary||'')}</p></td></tr>
<tr><td style="padding:30px 42px"><table role="presentation" width="100%"><tr><td><h2 style="font-family:${hf};font-size:21px;margin:0">Key regulatory highlights</h2></td><td align="right"><span style="font-size:11px;background:#e8f7f5;color:#087f7f;padding:7px 10px;border-radius:999px;font-weight:700">${updates.length} updates</span></td></tr></table>${updates.map((u,idx)=>`<table role="presentation" width="100%" cellspacing="0" cellpadding="0" style="margin-top:16px;border:1px solid #e8e8e5;border-radius:16px;overflow:hidden"><tr><td style="width:8px;background:${idx===0?a:p}"></td><td style="padding:20px"><table role="presentation" width="100%"><tr><td><div style="font-size:10px;color:${p};font-weight:800;text-transform:uppercase;letter-spacing:.08em">${esc(u.category||'Regulatory update')} · ${esc(u.jurisdiction||'')}</div></td><td align="right"><span style="font-size:10px;font-weight:800;color:${u.risk_level==='Critical'?'#b42318':u.risk_level==='High'?'#a44e12':'#087f7f'}">${esc(u.risk_level||'Update')}</span></td></tr></table><h3 style="font-family:${hf};font-size:18px;line-height:1.35;margin:8px 0;color:#151515">${esc(u.short_headline||u.regulation_title||u.title||`Update ${idx+1}`)}</h3><p style="font-size:14px;line-height:1.65;color:#56606c;margin:0 0 14px">${esc(u.summary||'')}</p><table role="presentation" width="100%"><tr><td style="font-size:12px;color:#6b7280;padding-right:8px"><b>Effective:</b> ${esc(formatDate(u.effective_date)||'Not specified')}</td><td style="font-size:12px;color:#6b7280"><b>Deadline:</b> ${esc(formatDate(u.compliance_deadline)||'Not specified')}</td></tr></table>${updateDetail(u,p,a)}</td></tr></table>`).join('')}</td></tr>
${m.insights?.length?`<tr><td style="padding:0 42px 30px"><div style="border-top:1px solid #e8e8e5;padding-top:28px"><h2 style="font-family:${hf};font-size:21px;margin:0 0 8px">Insights from our experts</h2><p style="font-size:13px;color:#6b7280;margin:0 0 8px">Practical analysis to help teams interpret and act on regulatory change.</p>${m.insights.map(x=>`<table role="presentation" width="100%" style="border-bottom:1px solid #ecece8"><tr><td style="padding:18px 0"><table role="presentation" width="100%"><tr><td><div style="font-weight:800;font-size:15px">${esc(x.title||'Insight')}</div></td><td align="right" style="font-size:10px;color:#7a858d">${esc(x.reading_time||'')}</td></tr></table><div style="font-size:13px;line-height:1.55;color:#6b7280;margin:6px 0">${esc(x.summary||'')}</div>${insightDetail(x,p,a)}</td></tr></table>`).join('')}</div></td></tr>`:''}
${promos.length?`<tr><td style="background:#f7f6f2;padding:30px 42px"><div style="font-size:11px;color:${a};font-weight:800;text-transform:uppercase;letter-spacing:.1em">How Lexplosion can help</div><h2 style="font-family:${hf};font-size:22px;margin:8px 0 20px">Technology designed around real compliance workflows</h2>${promos.map(x=>`<table role="presentation" width="100%" style="background:white;border-radius:16px;margin-bottom:14px;border:1px solid #e5e6e3;box-shadow:0 8px 20px rgba(23,63,66,.04)"><tr><td style="padding:21px 22px">${productLogo(x.product_name||'Lexplosion')}<h3 style="font-family:${hf};font-size:18px;line-height:1.3;margin:15px 0 8px;letter-spacing:-.01em">${esc(x.headline||'')}</h3><p style="font-size:13.5px;line-height:1.65;color:#5f6874;margin:0">${esc(x.description||'')}</p>${productDetail(x,p,a,i.contact_email)}</td></tr></table>`).join('')}</td></tr>`:''}
<tr><td style="padding:20px 36px;background:#132f31;color:#dce9e7"><table role="presentation" width="100%"><tr><td style="vertical-align:middle">${designState.companyLogo?`<img src="${designState.companyLogo}" alt="Lexplosion logo" style="display:block;max-width:190px;max-height:48px;object-fit:contain;filter:brightness(0) invert(1)">`:brandWord('LEXPLOSION','Innovating Legally','#ffffff','#cfe2df')}</td><td align="right" style="vertical-align:middle;font-size:11px;line-height:1.55;color:#d7e7e5">For products and services<br><a href="mailto:${esc(i.contact_email||'contact@lexplosion.in')}" style="color:#ffffff;text-decoration:none;font-weight:800">${esc(i.contact_email||'contact@lexplosion.in')}</a></td></tr></table></td></tr>
</table></td></tr></table></body></html>`}

function syncDesignControls(){
  $('headingFont').value=designState.headingFont;$('bodyFont').value=designState.bodyFont;
  $('headlineColor').value=designState.headlineColor;$('headlineHex').value=designState.headlineColor.toUpperCase();
  $('headlineSize').value=designState.headlineSize;$('headlineSizeValue').textContent=designState.headlineSize+' px';
  $('bodySize').value=designState.bodySize;$('bodySizeValue').textContent=designState.bodySize+' px';
  $('contentWidth').value=designState.contentWidth;$('contentWidthValue').textContent=designState.contentWidth+' px';
}
function readImage(file){return new Promise((resolve,reject)=>{if(!file)return resolve('');if(file.size>2*1024*1024)return reject(new Error('Logo files must be smaller than 2 MB.'));const r=new FileReader();r.onload=()=>resolve(r.result);r.onerror=()=>reject(new Error('Could not read the image.'));r.readAsDataURL(file)})}
function updateDesignFromControls(){designState.headingFont=$('headingFont').value;designState.bodyFont=$('bodyFont').value;designState.headlineColor=$('headlineColor').value;designState.headlineSize=+$('headlineSize').value;designState.bodySize=+$('bodySize').value;designState.contentWidth=+$('contentWidth').value;refreshPreview('Design updated')}
function bindDesignControls(){
  ['headingFont','bodyFont'].forEach(id=>$(id).addEventListener('change',updateDesignFromControls));
  $('headlineColor').addEventListener('input',()=>{$('headlineHex').value=$('headlineColor').value.toUpperCase();updateDesignFromControls()});
  $('headlineHex').addEventListener('input',()=>{if(isHex($('headlineHex').value)){$('headlineColor').value=$('headlineHex').value;updateDesignFromControls()}});
  [['headlineSize','headlineSizeValue'],['bodySize','bodySizeValue'],['contentWidth','contentWidthValue']].forEach(([id,out])=>$(id).addEventListener('input',()=>{$(out).textContent=$(id).value+' px';updateDesignFromControls()}));
  $('companyLogoInput').addEventListener('change',async e=>{try{designState.companyLogo=await readImage(e.target.files[0]);refreshPreview('Company logo updated')}catch(err){log(err.message,'error')}});
  $('clearCompanyLogoBtn').onclick=()=>{designState.companyLogo='';$('companyLogoInput').value='';refreshPreview('Default company logo restored')};
  $('productLogoInput').addEventListener('change',async e=>{try{const target=$('productLogoTarget').value;designState.productLogos[target]=await readImage(e.target.files[0]);refreshPreview(target+' logo updated')}catch(err){log(err.message,'error')}});
  $('clearProductLogoBtn').onclick=()=>{delete designState.productLogos[$('productLogoTarget').value];$('productLogoInput').value='';refreshPreview('Default product logo restored')};
  $('saveThemeBtn').onclick=()=>{localStorage.setItem('kommunicateDesign',JSON.stringify(designState));log('Design saved in this browser','ok')};
  $('resetThemeBtn').onclick=()=>{designState=JSON.parse(JSON.stringify(defaultDesign));localStorage.removeItem('kommunicateDesign');syncDesignControls();refreshPreview('Design reset')};
}
function populateEditor(){if(!model)return;$('editorEmpty').classList.add('hidden');$('contentEditor').classList.remove('hidden');$('editHeadline').value=model.issue?.headline||'';$('editIntroduction').value=model.issue?.introduction||'';$('editExecutiveSummary').value=model.issue?.executive_summary||'';renderUpdatesEditor()}
function renderUpdatesEditor(){ensureModelIds();const wrap=$('updatesEditor');wrap.innerHTML='';(model.updates||[]).forEach((u,idx)=>{const card=document.createElement('div');card.className='editor-card';const path=f=>targetPathForUpdate(u.id,f);card.innerHTML=`<h4>Regulatory update ${idx+1}</h4><div class="field"><div class="field-meta-line"><label>Headline ${sourceBadge(path('short_headline'))}</label><button class="field-map-action" type="button" data-map-target="${path('short_headline')}">Map source</button></div><input data-field="short_headline" data-id="${u.id}" value="${esc(u.short_headline||u.regulation_title||u.title||'')}"></div><div class="field"><div class="field-meta-line"><label>Summary ${sourceBadge(path('summary'))}</label><button class="field-map-action" type="button" data-map-target="${path('summary')}">Map source</button></div><textarea data-field="summary" data-id="${u.id}">${esc(u.summary||'')}</textarea></div><div class="mini-grid"><div class="field"><label>Category ${sourceBadge(path('category'))}</label><input data-field="category" data-id="${u.id}" value="${esc(u.category||'')}"></div><div class="field"><label>Risk level ${sourceBadge(path('risk_level'))}</label><select data-field="risk_level" data-id="${u.id}"><option ${u.risk_level==='Critical'?'selected':''}>Critical</option><option ${u.risk_level==='High'?'selected':''}>High</option><option ${u.risk_level==='Medium'?'selected':''}>Medium</option><option ${u.risk_level==='Low'?'selected':''}>Low</option></select></div></div><button class="remove-update" type="button" data-remove-id="${u.id}">Remove update</button>`;wrap.appendChild(card)});wrap.querySelectorAll('[data-remove-id]').forEach(b=>b.onclick=()=>{const id=b.dataset.removeId;const linked=Object.keys(mappingState.fieldMeta).some(k=>k.startsWith(`updates.${id}.`)&&mappingState.fieldMeta[k].origin==='document');if(linked&&!confirm('This update contains content mapped from a source document. Remove it anyway?'))return;model.updates=model.updates.filter(x=>x.id!==id);renderUpdatesEditor();refreshPreview('Update removed');saveDraft()});wrap.querySelectorAll('[data-map-target]').forEach(b=>b.onclick=()=>openSourceChooserForTarget(b.dataset.mapTarget))}
function populateEditor(){if(!model)return;$('editorEmpty').classList.add('hidden');$('contentEditor').classList.remove('hidden');$('editHeadline').value=model.issue?.headline||'';$('editIntroduction').value=model.issue?.introduction||'';$('editExecutiveSummary').value=model.issue?.executive_summary||'';const labels=[['editHeadline','headline'],['editIntroduction','introduction'],['editExecutiveSummary','executive_summary']];labels.forEach(([id,f])=>{const input=$(id),label=input.closest('.field')?.querySelector('label');if(label)label.innerHTML=`${f==='headline'?'Lead story headline':f==='introduction'?'Introduction':'Executive summary'} ${sourceBadge(targetPathForIssue(f))}`});renderUpdatesEditor()}
function applyEditorToModel(){if(!model)return;ensureModelIds();model.issue=model.issue||{};[['editHeadline','headline'],['editIntroduction','introduction'],['editExecutiveSummary','executive_summary']].forEach(([id,f])=>{const val=$(id).value.trim(),path=targetPathForIssue(f);if(val!==getTargetValue(path)){model.issue[f]=val;recordManualMeta(path)}});document.querySelectorAll('#updatesEditor [data-field]').forEach(el=>{const u=updateById(el.dataset.id),field=el.dataset.field;if(!u)return;const path=targetPathForUpdate(u.id,field),val=el.value.trim();if(val!==(u[field]||'')){u[field]=val;recordManualMeta(path)}});saveDraft();log('Content edits applied','ok')}
function openSourceChooserForTarget(path){const source=extractedSources.find(x=>x.included&&x.text);if(!source)return log('Upload or paste a source document first','warn');openSourcePreview(source.id);setTimeout(()=>{$('sourcePreviewText').dataset.preferredTarget=path},0)}
function show(){ensureModelIds();populateEditor();refreshPreview('Newsletter generated');renderMappingActivity();const e=validate(model);$('summary').innerHTML=e.length?`<span class="pill error">${e.length} issue(s)</span><br>${e.join('<br>')}`:`<span class="pill ok">Ready</span><br>${model.updates.length} regulatory updates, ${model.insights?.length||0} insights and ${model.promotions?.length||0} products found.`;log('Newsletter generated',e.length?'warn':'ok')}
function setStructuredStatus(message,type=''){
  const el=$('structuredFileStatus');if(!el)return;el.textContent=message||'';el.className='source-upload-status'+(type?' '+type:'');
}
function isStructuredExtension(file){return /\.xlsx$/i.test(file?.name||'')}
function structuredModelText(m,fileName='Structured source'){
  const lines=[`SOURCE: ${fileName}`,''];
  const issue=m?.issue||{};
  const issueFields=[
    ['Issue title',issue.issue_title],['Newsletter name',issue.newsletter_name],['Publication date',formatDate(issue.publication_date)],
    ['Volume',issue.volume_number],['Issue',issue.issue_number],['Headline',issue.headline],
    ['Introduction',issue.introduction],['Executive summary',issue.executive_summary],['Company',issue.company_name],
    ['Tagline',issue.company_tagline],['Contact email',issue.contact_email]
  ].filter(([,value])=>value!==undefined&&value!==null&&String(value).trim());
  if(issueFields.length){lines.push('ISSUE INFORMATION');issueFields.forEach(([label,value])=>lines.push(`${label}: ${value}`));lines.push('')}
  (m?.updates||[]).forEach((u,index)=>{
    lines.push(`REGULATORY UPDATE ${index+1}`);
    [['Headline',u.short_headline||u.regulation_title||u.title],['Category',u.category],['Jurisdiction',u.jurisdiction],['Regulator',u.regulator],['Summary',u.summary],['Effective date',formatDate(u.effective_date)],['Compliance deadline',formatDate(u.compliance_deadline)],['Risk level',u.risk_level],['Required action',u.required_action],['Business impact',u.business_impact],['Source URL',u.source_url||u.official_source_url]].forEach(([label,value])=>{if(value!==undefined&&value!==null&&String(value).trim())lines.push(`${label}: ${value}`)});
    lines.push('');
  });
  (m?.insights||[]).forEach((x,index)=>{
    lines.push(`EXPERT INSIGHT ${index+1}`);
    [['Title',x.title],['Summary',x.summary],['Full analysis',x.full_analysis],['Author',x.author],['URL',x.url]].forEach(([label,value])=>{if(value!==undefined&&value!==null&&String(value).trim())lines.push(`${label}: ${value}`)});
    lines.push('');
  });
  (m?.promotions||[]).forEach((x,index)=>{
    lines.push(`PRODUCT ${index+1}`);
    [['Product name',x.product_name],['Headline',x.headline],['Description',x.description],['Call to action',x.cta_text],['URL',x.cta_url]].forEach(([label,value])=>{if(value!==undefined&&value!==null&&String(value).trim())lines.push(`${label}: ${value}`)});
    lines.push('');
  });
  return cleanExtractedText(lines.join('\n'));
}
async function addStructuredSourceForReview(file,m){
  let text='';
  if(/\.docx$/i.test(file.name)){
    try{const result=await mammoth.extractRawText({arrayBuffer:await file.arrayBuffer()});text=cleanExtractedText(result.value)}catch(e){}
  }
  if(!text)text=structuredModelText(m,file.name);
  extractedSources=extractedSources.filter(source=>source.origin!=='structured');
  extractedSources.unshift({
    id:sourceUid(),name:file.name,type:/\.xlsx$/i.test(file.name)?'STRUCTURED XLSX':'STRUCTURED DOCX',
    text,size:file.size,included:true,aiIncluded:false,status:'ready',origin:'structured',
    note:'Mapped into the newsletter and available here for extraction review. Editing this review copy does not automatically remap the structured newsletter fields.'
  });
  renderSourceList();
  log(`${file.name}: structured source prepared for extraction review`,'ok');
}
async function handle(file){
  if(!file)return;
  if(!isStructuredExtension(file)){
    const message='Upload the standard .xlsx newsletter template here. Use Add supporting content for PDF, Word, TXT or supplementary Excel files.';
    setStructuredStatus(message,'error');log(message,'error');return;
  }
  $('fileName').textContent=file.name;setStructuredStatus('Reading and validating the file…','processing');log('Reading '+file.name);
  try{
    const isExcel=true;
    model=await parseExcel(file);ensureModelIds();mappingState={fieldMeta:{},history:[],activity:[]};aiState={suggestions:[],activeSuggestionId:null,lastRun:null};initialiseFieldMetadata('excel',file.name);
    const hasIssue=model?.issue&&Object.keys(model.issue).length;
    const hasUpdates=Array.isArray(model?.updates)&&model.updates.length;
    if(!hasIssue&&!hasUpdates){
      throw new Error('The workbook does not contain recognised Issue or Regulatory Updates data. Please use the standard newsletter template.');
    }
    setStructuredStatus('File loaded successfully. The newsletter content has been populated.','ok');
    show();
    await addStructuredSourceForReview(file,model);
    const updateCount=Array.isArray(model.updates)?model.updates.length:0;
    const insightCount=Array.isArray(model.insights)?model.insights.length:0;
    const productCount=Array.isArray(model.promotions)?model.promotions.length:0;
    $('summary').innerHTML=`<span class="pill ok">Structured file loaded</span><br><strong>${esc(file.name)}</strong><br>${updateCount} regulatory update${updateCount===1?'':'s'}, ${insightCount} insight${insightCount===1?'':'s'} and ${productCount} product${productCount===1?'':'s'} imported.<br><br><b>Extraction review ready:</b> the structured source is now listed below under Extracted sources. Select <b>Review text</b> to inspect it.`;
    log(`${file.name}: structured content imported successfully`,'ok');
  }catch(err){
    model=null;setStructuredStatus(err.message,'error');log('Could not parse file: '+err.message,'error');$('summary').innerHTML='<span class="pill error">Parsing failed</span><br>'+esc(err.message);
  }
}
(function bindStructuredUpload(){
  const zone=$('dropZone'),input=$('fileInput'),picker=$('chooseStructuredFileBtn');
  if(!zone||!input||!picker)return;
  // The visible control is a <label for="fileInput">, which opens the native
  // file picker without relying on a scripted click. This is more reliable in
  // Chrome, Safari, local-file mode and hosted deployments.
  picker.addEventListener('click',e=>e.stopPropagation());
  zone.addEventListener('click',e=>{
    if(e.target===picker||picker.contains(e.target)||e.target===input)return;
    input.value='';input.click();
  });
  zone.addEventListener('keydown',e=>{
    if(e.key==='Enter'||e.key===' '){e.preventDefault();input.value='';input.click()}
  });
  input.addEventListener('change',async e=>{
    const file=e.target.files&&e.target.files[0];
    if(file)await handle(file);
    input.value='';
  });
  ['dragenter','dragover'].forEach(ev=>zone.addEventListener(ev,e=>{e.preventDefault();e.stopPropagation();zone.classList.add('drag')}));
  ['dragleave','drop'].forEach(ev=>zone.addEventListener(ev,e=>{e.preventDefault();e.stopPropagation();zone.classList.remove('drag')}));
  zone.addEventListener('drop',async e=>{const file=e.dataTransfer?.files?.[0];if(file)await handle(file)});
})();
$('sampleBtn').onclick=()=>{model=JSON.parse(JSON.stringify(demo));mappingState={fieldMeta:{},history:[],activity:[]};aiState={suggestions:[],activeSuggestionId:null,lastRun:null};ensureModelIds();initialiseFieldMetadata('excel','Demo content');show();saveDraft()};$('applyEditsBtn').onclick=()=>{applyEditorToModel();show()};$('addUpdateBtn').onclick=()=>{if(!model)return;model.updates=model.updates||[];model.updates.push({id:uid('update'),display_order:model.updates.length+1,category:'Regulatory update',short_headline:'New regulatory update',summary:'Add the regulatory summary here.',risk_level:'Medium'});renderUpdatesEditor()};$('generateBtn').onclick=()=>{if(!model)return log('Upload a file or load the demo first','warn');applyEditorToModel();show()};
$('downloadBtn').onclick=()=>{const blob=new Blob([generatedHtml],{type:'text/html'}),a=document.createElement('a');a.href=URL.createObjectURL(blob);a.download='kommunicate-newsletter.html';a.click();URL.revokeObjectURL(a.href);log('HTML downloaded')};
function shareLink(){const currentIsHosted=/^https?:$/.test(location.protocol);const configured=(APP_CONFIG.publicAppUrl||'').replace(/\/$/,'');const base=configured||(currentIsHosted?location.origin+location.pathname:'');if(!base)throw new Error('Deploy the app before creating a shareable review link.');const payload=LZString.compressToEncodedURIComponent(JSON.stringify({model,primary:$('primaryColor').value,accent:$('accentColor').value,title:$('issueTitle').value,design:designState,mappingState,aiState}));return base+'#newsletter='+payload}
async function copyText(value){
  try{await navigator.clipboard.writeText(value);return true}catch(e){
    const box=document.createElement('textarea');box.value=value;box.style.position='fixed';box.style.opacity='0';document.body.appendChild(box);box.focus();box.select();const ok=document.execCommand('copy');box.remove();return ok;
  }
}
$('copyBtn').onclick=()=>{
  try{
    const link=shareLink();$('reviewLinkOutput').value=link;$('linkDialogStatus').textContent='The link has been created. Click “Copy link” or “Open link”.';$('linkDialogStatus').className='dialog-status';$('linkDialog').showModal();
  }catch(e){log(e.message,'error')}
};
$('copyLinkConfirmBtn').onclick=async()=>{
  const ok=await copyText($('reviewLinkOutput').value);$('linkDialogStatus').textContent=ok?'Review link copied to your clipboard.':'The browser blocked automatic copying. Select the link above and copy it manually.';$('linkDialogStatus').className='dialog-status '+(ok?'ok':'error');if(ok)log('Review link copied','ok');
};
$('openLinkBtn').onclick=()=>{const link=$('reviewLinkOutput').value;if(link)window.open(link,'_blank','noopener')};
$('emailBtn').onclick=()=>{
  if(!model)return log('Generate the newsletter first','warn');
  $('reviewSubject').value=`${model.issue?.newsletter_name||'Kommunicate'} newsletter review`;
  $('emailDialogStatus').textContent='';$('emailDialogStatus').className='dialog-status';$('emailDialog').showModal();
};
$('sendReviewBtn').onclick=async()=>{
  if(!model)return;
  const recipients=$('reviewRecipients').value.split(/[\n,;]+/).map(x=>x.trim()).filter(Boolean);
  const invalid=recipients.filter(x=>!/^\S+@\S+\.\S+$/.test(x));
  if(!recipients.length){$('emailDialogStatus').textContent='Enter at least one reviewer email address.';$('emailDialogStatus').className='dialog-status error';return}
  if(invalid.length){$('emailDialogStatus').textContent='Check these email addresses: '+invalid.join(', ');$('emailDialogStatus').className='dialog-status error';return}
  const btn=$('sendReviewBtn'),original=btn.textContent;btn.disabled=true;btn.textContent='Sending…';$('emailDialogStatus').textContent='Sending the newsletter…';$('emailDialogStatus').className='dialog-status';
  try{
    let link='';try{link=shareLink()}catch(e){}
    const response=await fetch(APP_CONFIG.emailEndpoint||'/.netlify/functions/send-newsletter',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({to:recipients,subject:$('reviewSubject').value.trim()||`${model.issue?.newsletter_name||'Kommunicate'} newsletter review`,message:$('reviewMessage').value.trim()||'Please review the generated newsletter and share your comments.',html:generatedHtml,reviewLink:link})});
    const result=await response.json().catch(()=>({}));
    if(!response.ok)throw new Error(result.error||'Email service returned an error');
    $('emailDialogStatus').textContent=`Newsletter sent to ${recipients.length} reviewer(s).`;$('emailDialogStatus').className='dialog-status ok';log(`Newsletter emailed to ${recipients.length} reviewer(s)`,'ok');
  }catch(e){$('emailDialogStatus').textContent=`Email could not be sent: ${e.message}. Ask the Netlify administrator to check RESEND_API_KEY and NEWSLETTER_FROM.`;$('emailDialogStatus').className='dialog-status error';log(`Email could not be sent: ${e.message}`,'error')}
  finally{btn.disabled=false;btn.textContent=original}
};

async function downloadDocx(){
  if(!model)return;
  applyEditorToModel();
  refreshPreview('Preparing branded Word document');
  try{
    if(window.htmlDocx&&typeof window.htmlDocx.asBlob==='function'){
      const wordCss=`<style>
        @page{size:A4;margin:18mm 16mm 18mm 16mm}
        body{font-family:${designState.bodyFont};font-size:${designState.bodySize}px;color:#161616;background:#ffffff}
        table{border-collapse:separate}
        p,div,td,span,a,li{font-family:${designState.bodyFont}} h1,h2,h3{font-family:${designState.headingFont}}
        details{display:block} summary{font-weight:bold;margin:8px 0} details>div{display:block!important}
        a{color:${$('primaryColor').value}}
      </style>`;
      const brandedHtml=generatedHtml.replace('</head>',wordCss+'</head>').replace(/<details/g,'<details open');
      const blob=window.htmlDocx.asBlob(brandedHtml,{orientation:'portrait',margins:{top:720,right:720,bottom:720,left:720,header:360,footer:360,gutter:0}});
      const a=document.createElement('a');a.href=URL.createObjectURL(blob);a.download='kommunicate-newsletter-branded.docx';a.click();setTimeout(()=>URL.revokeObjectURL(a.href),1200);log('Branded Word document downloaded','ok');return;
    }
    throw new Error('The branded Word converter did not load');
  }catch(e){
    log(`Branded Word conversion failed: ${e.message}`,'error');
  }
}
$('downloadDocxBtn').onclick=downloadDocx;
(function initDesign(){try{const saved=localStorage.getItem('kommunicateDesign');if(saved)designState={...JSON.parse(JSON.stringify(defaultDesign)),...JSON.parse(saved)};}catch(e){}syncDesignControls();bindDesignControls();})();
(function restore(){const m=location.hash.match(/newsletter=([^&]+)/);if(m)try{const x=JSON.parse(LZString.decompressFromEncodedURIComponent(m[1]));model=x.model;$('primaryColor').value=x.primary||'#173f42';$('primaryHex').value=$('primaryColor').value.toUpperCase();$('accentColor').value=x.accent||'#f58220';$('accentHex').value=$('accentColor').value.toUpperCase();$('issueTitle').value=x.title||'Compliance Newsletter';if(x.mappingState)mappingState=x.mappingState;if(x.aiState)aiState=x.aiState;if(x.design){designState={...JSON.parse(JSON.stringify(defaultDesign)),...x.design,productLogos:x.design.productLogos||{}};syncDesignControls()}show();log('Shared newsletter loaded','ok')}catch(e){log('Shared link could not be read','error')}else{const hasDraft=!!localStorage.getItem(DRAFT_KEY);$('preview').innerHTML=`<div style="padding:70px 46px;text-align:center"><div style="display:inline-block;background:#173f42;color:white;border-radius:18px;padding:28px 34px"><div style="font-size:44px;font-weight:900;letter-spacing:-.05em">Kommunicate<span style="color:#f58220">.</span></div><div style="font-size:11px;letter-spacing:.13em;text-transform:uppercase;margin-top:8px;color:#cfe2df">Lexplosion's regulatory compliance newsletter</div></div><h2 style="font-size:26px;margin:24px 0 8px">Start a new edition or restore a saved draft</h2><p style="color:#6b7280;line-height:1.6">Upload the standard Excel template to generate the newsletter. Supporting documents and pasted text can be added afterward.</p>${hasDraft?'<p style="margin-top:18px;font-weight:700;color:#173f42">A saved draft is available in this browser. It has not been loaded automatically.</p>':''}</div>`;$('summary').innerHTML=hasDraft?'<span class="pill">Saved draft available</span><br>Choose <b>Restore previous draft</b> to load it, or begin with a new source.':'Upload the standard Excel template to generate the newsletter.';if(hasDraft)log('Saved draft available — not loaded automatically','ok')}})();
/* Phase 2: multi-format unstructured source ingestion */
function sourceUid(){return 'src_'+Date.now().toString(36)+'_'+Math.random().toString(36).slice(2,8)}
function humanSize(bytes){if(!Number.isFinite(bytes))return'';if(bytes<1024)return bytes+' B';if(bytes<1048576)return(bytes/1024).toFixed(1)+' KB';return(bytes/1048576).toFixed(1)+' MB'}
function sourceType(fileName){const ext=String(fileName||'').split('.').pop().toLowerCase();return ext||'text'}
function cleanExtractedText(text){return String(text||'').replace(/\u0000/g,'').replace(/[ \t]+\n/g,'\n').replace(/\n{4,}/g,'\n\n\n').trim()}
async function getPdfModule(){
  if(!pdfModulePromise)pdfModulePromise=import('https://cdnjs.cloudflare.com/ajax/libs/pdf.js/4.4.168/pdf.min.mjs').then(mod=>{mod.GlobalWorkerOptions.workerSrc='https://cdnjs.cloudflare.com/ajax/libs/pdf.js/4.4.168/pdf.worker.min.mjs';return mod});
  return pdfModulePromise;
}
async function extractPdf(file){
  const pdfjs=await getPdfModule();
  const data=new Uint8Array(await file.arrayBuffer());
  const pdf=await pdfjs.getDocument({data}).promise;
  const pages=[];
  for(let n=1;n<=pdf.numPages;n++){
    const page=await pdf.getPage(n);const content=await page.getTextContent();
    const line=content.items.map(item=>item.str).join(' ').replace(/\s+/g,' ').trim();
    pages.push(`Page ${n}\n${line}`);
  }
  return{text:cleanExtractedText(pages.join('\n\n')),pages:pdf.numPages};
}
async function extractUnstructuredFile(file){
  const ext=sourceType(file.name);
  if(ext==='pdf')return extractPdf(file);
  if(ext==='docx'){
    const result=await mammoth.extractRawText({arrayBuffer:await file.arrayBuffer()});
    return{text:cleanExtractedText(result.value),sections:(result.value.match(/\n\s*\n/g)||[]).length+1};
  }
  if(ext==='txt')return{text:cleanExtractedText(await file.text())};
  if(ext==='xlsx'){const parsed=await parseExcel(file);return{text:structuredModelText(parsed,file.name)};}
  throw new Error('Unsupported file type. Use PDF, DOCX, TXT or XLSX.');
}
async function addDocumentFiles(files){
  const valid=[...files].filter(Boolean);
  if(!valid.length)return;
  log(`Extracting ${valid.length} source document${valid.length===1?'':'s'}…`);
  for(const file of valid){
    const placeholder={id:sourceUid(),name:file.name,type:sourceType(file.name).toUpperCase(),text:'',size:file.size,included:true,aiIncluded:true,status:'processing',origin:'supporting'};
    extractedSources.push(placeholder);renderSourceList();
    try{
      const result=await extractUnstructuredFile(file);
      placeholder.text=result.text;placeholder.pages=result.pages||null;placeholder.sections=result.sections||null;placeholder.status=result.text?'ready':'empty';
      if(!result.text)placeholder.error='No selectable text was found. The PDF may be scanned and require OCR.';
      log(`${file.name}: text extracted`,result.text?'ok':'warn');
    }catch(error){placeholder.status='error';placeholder.error=error.message;log(`${file.name}: ${error.message}`,'error')}
    renderSourceList();
  }
}
function sourceMeta(source){
  const parts=[source.type,humanSize(source.size)];
  if(source.pages)parts.push(`${source.pages} page${source.pages===1?'':'s'}`);
  if(source.sections)parts.push(`${source.sections} section${source.sections===1?'':'s'}`);
  if(source.text)parts.push(`${source.text.split(/\s+/).filter(Boolean).length.toLocaleString()} words`);
  return parts.filter(Boolean).join(' · ');
}
function renderSourceList(){
  const list=$('sourceList');if(!list)return;
  $('sourceCount').textContent=`${extractedSources.length} source${extractedSources.length===1?'':'s'}`;
  $('combineSourcesBtn').disabled=$('clearSourcesBtn').disabled=$('previewCombinedBtn').disabled=!extractedSources.some(x=>x.text&&x.included);
  if(!extractedSources.length){list.innerHTML='<div class="empty-source">No extracted sources available.</div>';return}
  list.innerHTML='';
  extractedSources.forEach((source,index)=>{
    const card=document.createElement('div');card.className='source-item';
    const status=source.status==='error'?'error':source.status==='processing'?'processing':'ready';const linkedCount=(mappingState.activity||[]).filter(x=>x.sourceId===source.id).length;
    if(source.aiIncluded===undefined)source.aiIncluded=source.origin!=='structured';
    card.innerHTML=`<div class="source-item-top"><div><div class="source-name">${esc(source.name)}</div><div class="source-meta">${esc(sourceMeta(source))}</div></div><div class="source-item-actions"><button class="source-mini" data-action="up" ${index===0?'disabled':''}>↑</button><button class="source-mini" data-action="down" ${index===extractedSources.length-1?'disabled':''}>↓</button><button class="source-mini danger" data-action="remove">×</button></div></div><div class="source-status ${status==='error'?'error':''}">${status==='processing'?'Extracting':status==='error'?'Extraction failed':source.included?'Included':'Excluded'}</div><div class="source-ai-state ${source.aiIncluded?'on':'off'}">${source.aiIncluded?'Selected for AI analysis':'Not selected for AI analysis'}</div>${source.note?`<div class="source-note">${esc(source.note)}</div>`:''}${linkedCount?`<div class="linked-warning">Linked to ${linkedCount} newsletter field${linkedCount===1?'':'s'}</div>`:''}${source.error?`<div class="source-snippet" style="color:#b42318">${esc(source.error)}</div>`:`<div class="source-snippet">${esc(source.text||'No text extracted.')}</div>`}<div class="source-item-actions" style="margin-top:8px"><button class="source-mini" data-action="preview" ${!source.text?'disabled':''}>Review text</button><button class="source-mini ${source.aiIncluded?'ai-active':''}" data-action="ai-toggle" ${!source.text||status==='error'?'disabled':''}>${source.aiIncluded?'Remove from AI':'Use for AI'}</button><button class="source-mini" data-action="toggle" ${!source.text?'disabled':''}>${source.included?'Exclude':'Include'}</button></div>`;
    card.querySelectorAll('[data-action]').forEach(btn=>btn.onclick=()=>handleSourceAction(source.id,btn.dataset.action));list.appendChild(card);
  });
  const aiSources=extractedSources.filter(x=>x.aiIncluded&&x.text&&x.status!=='error');const readyForAi=aiSources.length>0;if($('analyzeSourcesBtn'))$('analyzeSourcesBtn').disabled=!readyForAi;if($('aiSourceSummary'))$('aiSourceSummary').textContent=readyForAi?`${aiSources.length} source${aiSources.length===1?'':'s'} selected: ${aiSources.map(x=>x.name).join(', ')}`:'No supporting source selected for AI analysis.';if(model)saveDraft();
}
function handleSourceAction(id,action){
  const index=extractedSources.findIndex(x=>x.id===id);if(index<0)return;
  if(action==='remove'){const linked=(mappingState.activity||[]).filter(x=>x.sourceId===id).length;if(linked&&!confirm(`This source is linked to ${linked} mapped newsletter field(s). Removing it will retain the newsletter content but remove the review source. Continue?`))return;extractedSources.splice(index,1);}
  if(action==='up'&&index>0)[extractedSources[index-1],extractedSources[index]]=[extractedSources[index],extractedSources[index-1]];
  if(action==='down'&&index<extractedSources.length-1)[extractedSources[index+1],extractedSources[index]]=[extractedSources[index],extractedSources[index+1]];
  if(action==='toggle')extractedSources[index].included=!extractedSources[index].included;
  if(action==='ai-toggle'){extractedSources[index].aiIncluded=!extractedSources[index].aiIncluded;if(aiState.suggestions.length){aiState.suggestions=[];aiState.activeSuggestionId=null;renderAiSuggestions();log('AI source selection changed; previous unapplied suggestions were cleared','warn')}}
  if(action==='preview')openSourcePreview(id);
  renderSourceList();
}
function combinedSourceText(){return extractedSources.filter(x=>x.included&&x.text).map((x,i)=>`SOURCE ${i+1}: ${x.name}\n${'-'.repeat(Math.min(70,x.name.length+10))}\n${x.text}`).join('\n\n\n')}
function openSourcePreview(id='__combined'){
  activeSourceId=id;const combined=id==='__combined';const source=combined?null:extractedSources.find(x=>x.id===id);if(!combined&&!source)return;
  $('sourcePreviewTitle').textContent=combined?'Combined source text':source.name;
  $('sourcePreviewText').value=combined?combinedSourceText():source.text;
  $('sourcePreviewMeta').textContent=combined?`${extractedSources.filter(x=>x.included&&x.text).length} included sources · ${$('sourcePreviewText').value.split(/\s+/).filter(Boolean).length.toLocaleString()} words`:sourceMeta(source);
  $('saveSourceEditBtn').style.display=combined?'none':'';$('sourcePreviewStatus').textContent='';$('sourcePreviewDialog').showModal();
}
function downloadTextFile(name,text){const blob=new Blob([text],{type:'text/plain;charset=utf-8'}),a=document.createElement('a');a.href=URL.createObjectURL(blob);a.download=name;a.click();setTimeout(()=>URL.revokeObjectURL(a.href),800)}
function bindSourceIngestion(){
  const activateSupportMode=mode=>{
    document.querySelectorAll('[data-support-mode]').forEach(tab=>{
      const active=tab.dataset.supportMode===mode;
      tab.classList.toggle('active',active);
      tab.setAttribute('aria-selected',String(active));
    });
    document.querySelectorAll('[data-support-pane]').forEach(pane=>{
      pane.classList.toggle('hidden',pane.dataset.supportPane!==mode);
    });
  };
  document.querySelectorAll('[data-support-mode]').forEach(tab=>{
    tab.addEventListener('click',e=>{e.preventDefault();activateSupportMode(tab.dataset.supportMode)});
  });
  activateSupportMode('upload');
  const zone=$('documentDropZone'),input=$('documentFileInput');zone.onclick=()=>input.click();input.onchange=e=>{addDocumentFiles(e.target.files);input.value=''};
  ['dragenter','dragover'].forEach(ev=>zone.addEventListener(ev,e=>{e.preventDefault();zone.classList.add('drag')}));
  ['dragleave','drop'].forEach(ev=>zone.addEventListener(ev,e=>{e.preventDefault();zone.classList.remove('drag')}));
  zone.addEventListener('drop',e=>addDocumentFiles(e.dataTransfer.files));
  $('addPastedSourceBtn').onclick=()=>{const text=cleanExtractedText($('pastedSourceText').value);if(!text)return log('Paste source text before adding it','warn');extractedSources.push({id:sourceUid(),name:`Pasted text ${extractedSources.filter(x=>x.type==='TEXT').length+1}`,type:'TEXT',text,size:new Blob([text]).size,included:true,aiIncluded:true,status:'ready',origin:'supporting'});$('pastedSourceText').value='';renderSourceList();log('Pasted source text added','ok')};
  $('clearPastedSourceBtn').onclick=()=>$('pastedSourceText').value='';
  $('clearSourcesBtn').onclick=()=>{if(!confirm('Remove all extracted source documents? Pending AI suggestions will also be cleared.'))return;extractedSources=[];aiState={suggestions:[],activeSuggestionId:null,lastRun:null};renderSourceList();renderAiSuggestions();log('All supporting sources and pending AI suggestions cleared')};
  $('combineSourcesBtn').onclick=()=>openSourcePreview('__combined');$('previewCombinedBtn').onclick=()=>openSourcePreview('__combined');
  $('saveSourceEditBtn').onclick=()=>{const source=extractedSources.find(x=>x.id===activeSourceId);if(!source)return;source.text=cleanExtractedText($('sourcePreviewText').value);source.size=new Blob([source.text]).size;renderSourceList();$('sourcePreviewStatus').textContent='Edits saved.';$('sourcePreviewStatus').className='dialog-status ok';log(`${source.name}: extracted text updated`,'ok')};
  const sourcePreviewText=$('sourcePreviewText');
  const rememberPreviewSelection=()=>{const start=sourcePreviewText.selectionStart||0,end=sourcePreviewText.selectionEnd||0;const text=sourcePreviewText.value.slice(start,end).trim();if(text)lastPreviewSelection={text,start,end}};
  ['select','keyup','mouseup','touchend'].forEach(ev=>sourcePreviewText.addEventListener(ev,rememberPreviewSelection));
  $('mapSelectedTextBtn').onclick=e=>{e.preventDefault();rememberPreviewSelection();const selected=lastPreviewSelection.text||'';if(!selected){$('sourcePreviewStatus').textContent='Highlight the exact sentence or paragraph you want to map. Full-document mapping is disabled to prevent oversized newsletter fields.';$('sourcePreviewStatus').className='dialog-status warn';sourcePreviewText.focus();return}const source=activeSourceId==='__combined'?{id:'combined',name:'Combined sources'}:extractedSources.find(x=>x.id===activeSourceId);const preferred=sourcePreviewText.dataset.preferredTarget||'';delete sourcePreviewText.dataset.preferredTarget;$('sourcePreviewStatus').textContent='Opening mapping options…';$('sourcePreviewStatus').className='dialog-status ok';try{$('sourcePreviewDialog').close()}catch(err){$('sourcePreviewDialog').removeAttribute('open')}setTimeout(()=>{openMappingDialog(selected,source);if(preferred){$('mappingTarget').value=preferred;updateMappingComparison()}lastPreviewSelection={text:'',start:0,end:0}},0)};
  $('downloadSourceTextBtn').onclick=()=>downloadTextFile(activeSourceId==='__combined'?'kommunicate-combined-sources.txt':(extractedSources.find(x=>x.id===activeSourceId)?.name||'source')+'.txt',$('sourcePreviewText').value);
  $('mappingTarget').onchange=()=>{if(pendingSnippet)$('mappingProposed').value=suggestedForField(pendingSnippet.text,$('mappingTarget').value);updateMappingComparison()};$('mappingMode').onchange=updateMappingComparison;$('mappingProposed').addEventListener('input',updateMappingGuidance);$('fitMappingBtn').onclick=()=>{if(!pendingSnippet)return;$('mappingProposed').value=suggestedForField(pendingSnippet.text,$('mappingTarget').value);updateMappingGuidance()};$('applyMappingBtn').onclick=applyPendingMapping;$('undoMappingBtn').onclick=undoLastMapping;$('saveDraftBtn').onclick=()=>saveDraft(false);$('restoreDraftBtn').onclick=()=>{if(!localStorage.getItem(DRAFT_KEY))return log('No saved draft is available in this browser','warn');if(model&&!confirm('Restore the previous saved draft? Unsaved changes in the current session will be replaced.'))return;if(restoreDraft()){syncDesignControls();show();renderSourceList();renderAiSuggestions();log('Previous draft restored','ok')}else log('The saved draft could not be restored','error')};$('clearSavedDraftBtn').onclick=()=>{if(!localStorage.getItem(DRAFT_KEY))return log('No saved draft is available to clear','warn');if(!confirm('Clear the saved draft from this browser? This cannot be undone.'))return;localStorage.removeItem(DRAFT_KEY);log('Saved draft cleared from this browser','ok')};$('startEmptyBtn').onclick=()=>{if(model&&!confirm('Start a new empty newsletter? Unsaved changes in the current session will be replaced.'))return;model={issue:{issue_title:'Compliance Newsletter',newsletter_name:'Kommunicate',headline:'',introduction:'',executive_summary:''},updates:[],insights:[],promotions:[]};extractedSources=[];mappingState={fieldMeta:{},history:[],activity:[]};aiState={suggestions:[],activeSuggestionId:null,lastRun:null};ensureModelIds();show();renderSourceList();renderAiSuggestions();saveDraft();log('Empty newsletter started','ok')};
  renderSourceList();renderMappingActivity();
}
bindSourceIngestion();


const AI_FIELDS=['short_headline','summary','category','jurisdiction','regulator','effective_date','compliance_deadline','risk_level','required_action'];
function normaliseAiSuggestion(raw,index=0){
  const fields={};AI_FIELDS.forEach(k=>fields[k]=String(raw?.fields?.[k]||'').trim());
  return {id:raw?.id||uid('ai'),status:'pending',sourceId:String(raw?.source_id||raw?.sourceId||''),sourceName:String(raw?.source_name||raw?.sourceName||'Supporting source'),sourceExcerpt:String(raw?.source_excerpt||raw?.sourceExcerpt||'').trim(),confidence:Math.max(0,Math.min(1,Number(raw?.confidence)||0)),rationale:String(raw?.rationale||'').trim(),fields,createdAt:new Date().toISOString(),targetUpdateId:'__new__',approvedAt:null,appliedAt:null,index};
}
function aiPending(){return aiState.suggestions.filter(x=>x.status==='pending')}
function renderAiSuggestions(){
  const list=$('aiSuggestionList');if(!list)return;const items=aiState.suggestions||[],pending=items.filter(x=>x.status==='pending').length,approved=items.filter(x=>x.status==='approved').length;
  $('aiSuggestionCount').textContent=`${pending} pending${approved?` · ${approved} approved`:''}`;$('clearAiSuggestionsBtn').disabled=!items.length;$('reviewNextAiBtn').disabled=!pending;$('applyAllAiBtn').disabled=!approved;
  if(!items.length){list.innerHTML='<div class="empty-source">Generate suggestions from included supporting sources.</div>';return}
  list.innerHTML=items.map(s=>`<div class="ai-suggestion-item ${esc(s.status)}"><div class="ai-suggestion-top"><div><div class="ai-suggestion-title">${esc(s.fields.short_headline||'Untitled regulatory update')}</div><div class="ai-suggestion-meta">${esc(s.sourceName)} · ${esc(s.fields.regulator||s.fields.jurisdiction||'Source-derived suggestion')} · ${esc(s.status)}</div></div><span class="ai-confidence">${Math.round((s.confidence||0)*100)}% confidence</span></div><div class="ai-suggestion-actions"><button type="button" data-ai-action="review" data-ai-id="${s.id}">Review</button>${s.status==='pending'?`<button type="button" data-ai-action="approve" data-ai-id="${s.id}">Approve</button><button type="button" data-ai-action="reject" data-ai-id="${s.id}">Reject</button>`:''}${s.status==='approved'?`<button type="button" data-ai-action="apply" data-ai-id="${s.id}">Apply</button>`:''}</div></div>`).join('');
  list.querySelectorAll('[data-ai-action]').forEach(b=>b.onclick=()=>handleAiQueueAction(b.dataset.aiId,b.dataset.aiAction));
}
function aiTargetOptions(){const opts=[{value:'__new__',label:'Create a new regulatory update'}];(model?.updates||[]).forEach((u,i)=>opts.push({value:u.id,label:`Regulatory update ${i+1}: ${u.short_headline||'Untitled'}`}));return opts}
function fieldInputHtml(s,field,label){const value=s.fields[field]||'',target=$('aiSuggestionTarget')?.value||'__new__',path=target==='__new__'?'':targetPathForUpdate(target,field),current=path?getTargetValue(path):'',conflict=Boolean(current);const isRisk=field==='risk_level';return `<div class="ai-field-row ${conflict?'has-conflict':''}"><input type="checkbox" data-ai-check="${field}" ${value?'checked':''} ${!value?'disabled':''} aria-label="Include ${esc(label)}"><div class="ai-field-content"><label><span>${esc(label)}</span><span>${value.length} chars</span></label>${isRisk?`<select data-ai-field="${field}"><option value="">Not specified</option>${['Critical','High','Medium','Low'].map(x=>`<option ${value===x?'selected':''}>${x}</option>`).join('')}</select>`:`<textarea data-ai-field="${field}">${esc(value)}</textarea>`}${conflict?`<div class="ai-current-note">Existing value: ${esc(current.slice(0,180))}${current.length>180?'…':''}</div>`:''}</div></div>`}
function renderAiFieldReview(){const s=aiState.suggestions.find(x=>x.id===aiState.activeSuggestionId);if(!s)return;const labels={short_headline:'Headline',summary:'Summary',category:'Category',jurisdiction:'Jurisdiction',regulator:'Regulator',effective_date:'Effective date',compliance_deadline:'Compliance deadline',risk_level:'Risk level',required_action:'Required action'};$('aiFieldReview').innerHTML=AI_FIELDS.map(f=>fieldInputHtml(s,f,labels[f])).join('')}
function openAiSuggestion(id){const s=aiState.suggestions.find(x=>x.id===id);if(!s)return;aiState.activeSuggestionId=id;const select=$('aiSuggestionTarget');select.innerHTML=aiTargetOptions().map(o=>`<option value="${esc(o.value)}">${esc(o.label)}</option>`).join('');select.value=s.targetUpdateId&&aiTargetOptions().some(x=>x.value===s.targetUpdateId)?s.targetUpdateId:'__new__';$('aiSuggestionMeta').innerHTML=`<span>${esc(s.sourceName)}</span><span>${Math.round(s.confidence*100)}% confidence</span><span>${esc(s.status)}</span>`;$('aiSourceExcerpt').textContent=s.sourceExcerpt||'No source excerpt returned.';$('aiSuggestionStatus').textContent='';renderAiFieldReview();safeShowDialog($('aiSuggestionDialog'))}
function syncAiDialogToState(){const s=aiState.suggestions.find(x=>x.id===aiState.activeSuggestionId);if(!s)return null;s.targetUpdateId=$('aiSuggestionTarget').value;document.querySelectorAll('#aiFieldReview [data-ai-field]').forEach(el=>s.fields[el.dataset.aiField]=el.value.trim());return s}
function handleAiQueueAction(id,action){const s=aiState.suggestions.find(x=>x.id===id);if(!s)return;if(action==='review')return openAiSuggestion(id);if(action==='approve'){s.status='approved';s.approvedAt=new Date().toISOString()}if(action==='reject'){s.status='rejected'}if(action==='apply'){openAiSuggestion(id);return}renderAiSuggestions();saveDraft()}
function ensureNewsletterForAi(){if(!model)model={issue:{issue_title:'Compliance Newsletter',newsletter_name:'Kommunicate',headline:'',introduction:'',executive_summary:''},updates:[],insights:[],promotions:[]};ensureModelIds()}
function applyAiSuggestion(id,closeDialog=true){const s=aiState.suggestions.find(x=>x.id===id);if(!s)return false;ensureNewsletterForAi();syncAiDialogToState();const checked=[...document.querySelectorAll('#aiFieldReview [data-ai-check]:checked')].map(x=>x.dataset.aiCheck);if(!checked.length){$('aiSuggestionStatus').textContent='Select at least one proposed field.';return false}let target=s.targetUpdateId,update;if(target==='__new__'){update={id:uid('update'),display_order:model.updates.length+1};model.updates.push(update);target=update.id;s.targetUpdateId=target}else update=updateById(target);if(!update)return false;const mode=$('aiConflictMode').value;checked.forEach(field=>{const value=(s.fields[field]||'').trim();if(!value)return;const path=targetPathForUpdate(target,field),oldValue=getTargetValue(path),oldMeta=metaFor(path);const rule=fieldRule(path);let safeValue=value.slice(0,rule.max);const next=mode==='append'&&oldValue?`${oldValue}\n\n${safeValue}`:safeValue;mappingState.history.push({path,oldValue,oldMeta:oldMeta?JSON.parse(JSON.stringify(oldMeta)):null,newValue:next,at:new Date().toISOString(),origin:'ai'});setTargetValue(path,next);mappingState.fieldMeta[path]={origin:'document',sourceId:s.sourceId,sourceName:`AI · ${s.sourceName}`,sourceExcerpt:s.sourceExcerpt,confidence:s.confidence,updatedAt:new Date().toISOString(),overrode:Boolean(oldValue),aiSuggestionId:s.id};mappingState.activity.push({path,targetLabel:targetLabel(path),sourceId:s.sourceId,sourceName:`AI · ${s.sourceName}`,mode:oldValue?mode:'fill',appliedAt:new Date().toISOString()})});s.status='applied';s.appliedAt=new Date().toISOString();show();renderMappingActivity();renderAiSuggestions();saveDraft();log(`AI suggestion applied to ${targetLabel(targetPathForUpdate(target,'short_headline')).replace(' → Headline','')}`,'ok');if(closeDialog)try{$('aiSuggestionDialog').close()}catch(e){}return true}
function approveAiSuggestion(){const s=syncAiDialogToState();if(!s)return;s.status='approved';s.approvedAt=new Date().toISOString();renderAiSuggestions();saveDraft();$('aiSuggestionStatus').textContent='Suggestion approved and retained in the queue.';$('aiSuggestionStatus').className='dialog-status ok'}
function rejectAiSuggestion(){const s=syncAiDialogToState();if(!s)return;s.status='rejected';renderAiSuggestions();saveDraft();try{$('aiSuggestionDialog').close()}catch(e){}log('AI suggestion rejected')}
function firstMatch(text,patterns){for(const pattern of patterns){const match=text.match(pattern);if(match&&match[1])return cleanExtractedText(match[1])}return ''}
function sentenceList(text){return cleanSourceText(text).replace(/\s+/g,' ').split(/(?<=[.!?])\s+/).map(x=>x.trim()).filter(x=>x.length>30)}
function detectRegulator(text){const pairs=[[/Reserve Bank of India|\bRBI\b/i,'Reserve Bank of India (RBI)'],[/Securities and Exchange Board of India|\bSEBI\b/i,'Securities and Exchange Board of India (SEBI)'],[/Insurance Regulatory and Development Authority of India|\bIRDAI\b/i,'Insurance Regulatory and Development Authority of India (IRDAI)'],[/Ministry of Corporate Affairs|\bMCA\b/i,'Ministry of Corporate Affairs (MCA)'],[/Competition Commission of India|\bCCI\b/i,'Competition Commission of India (CCI)']];for(const [re,name] of pairs)if(re.test(text))return name;return ''}
function detectJurisdiction(text,regulator){if(/India|Reserve Bank of India|SEBI|IRDAI|MCA|CCI/i.test(text+regulator))return 'India';return ''}
function localAiFallback(sources,instruction){const suggestions=[];sources.forEach((src,i)=>{const clean=cleanSourceText(src.text);const sentences=sentenceList(clean);if(!sentences.length)return;const regulator=detectRegulator(clean);const jurisdiction=detectJurisdiction(clean,regulator);const title=firstMatch(clean,[/RBI Issues\s+(.{20,220}?)(?=Detailed instructions)/i,/(.{20,220}?(?:Directions|Regulations|Circular|Notification|Amendment).{0,80}?)(?=Detailed instructions|Accordingly|The comments)/i]);const isDraft=/\bdraft\b/i.test(clean);const deadline=firstMatch(clean,[/on or before\s+([A-Z][a-z]+\s+\d{1,2},\s+\d{4})/i,/by\s+([A-Z][a-z]+\s+\d{1,2},\s+\d{4})/i]);const publicationDate=firstMatch(clean,[/\b([A-Z][a-z]+\s+\d{1,2},\s+\d{4})\b/]);let headline=title?`${isDraft?'RBI proposes ':''}${title.replace(/^Draft\s+/i,'').replace(/[‘’'\"]/g,'').trim()}`:'';if(!headline&&regulator)headline=`${regulator.split(' (')[0]} issues ${isDraft?'draft ':''}regulatory update`;headline=suggestedForField(headline||sentences[0],`updates.local.short_headline`).slice(0,180);const relevant=sentences.filter(x=>!/website|email|phone|press release/i.test(x));const summaryParts=relevant.filter(x=>/recovery|regulated entit|borrower|agent|directions|comments|feedback|conduct|loan/i.test(x)&&x.length<720&&(x.match(/Reserve Bank of India/g)||[]).length<3).slice(0,3);const summary=(summaryParts.length?summaryParts:relevant.slice(0,3)).join(' ').slice(0,1100);let requiredAction='';if(deadline)requiredAction=`Review the draft directions, assess the impact on current policies and recovery-agent arrangements, and submit comments by ${deadline}.`;else if(isDraft)requiredAction='Review the draft, assess likely operational impact and prepare feedback or implementation plans as appropriate.';const excerpt=(summaryParts.length?summaryParts.slice(0,2):relevant.slice(0,2)).join(' ').slice(0,650);const fields={short_headline:headline,summary,category:isDraft?'Draft amendment directions':'Regulatory update',jurisdiction,regulator,effective_date:isDraft?'':publicationDate,compliance_deadline:deadline,risk_level:isDraft?'Medium':'High',required_action:requiredAction};suggestions.push(normaliseAiSuggestion({source_id:src.id,source_name:src.name,source_excerpt:excerpt,confidence:(regulator&&headline.length&&summary.length)?0.78:0.62,rationale:'Local preview extraction based only on the selected source document.',fields},i))});return suggestions.slice(0,8)}
async function requestAiSuggestions(sources,instruction,focusSuggestion=null){const payload={instruction,sources:sources.map(x=>({id:x.id,name:x.name,type:x.type,text:x.text.slice(0,30000)})),existingUpdates:(model?.updates||[]).map(x=>({id:x.id,headline:x.short_headline||'',summary:x.summary||''})),focusSuggestion};if(location.protocol==='file:'||['localhost','127.0.0.1'].includes(location.hostname))return {suggestions:localAiFallback(sources,instruction),mode:'local-preview'};const response=await fetch('/.netlify/functions/analyze-sources',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify(payload)});let data={};try{data=await response.json()}catch(e){}if(!response.ok)throw new Error(data.error||`AI service returned ${response.status}`);return data}
async function generateAiSuggestions(){const sources=extractedSources.filter(x=>x.aiIncluded&&x.text&&x.status!=='error');if(!sources.length)return log('Select at least one source using Use for AI before generating suggestions','warn');const btn=$('analyzeSourcesBtn'),status=$('aiRunStatus');btn.disabled=true;btn.textContent='Analyzing…';status.textContent=`Analyzing only: ${sources.map(x=>x.name).join(', ')}…`;status.className='source-upload-status processing';try{const result=await requestAiSuggestions(sources,$('aiInstruction').value.trim());const selectedById=new Map(sources.map(x=>[x.id,x]));const incoming=(result.suggestions||[]).map(normaliseAiSuggestion).filter(x=>selectedById.has(x.sourceId)&&(x.fields.short_headline||x.fields.summary)).map(x=>{const source=selectedById.get(x.sourceId);x.sourceName=source.name;return x});if(!incoming.length)throw new Error('No usable regulatory updates were identified. Refine the instruction or source content.');aiState.suggestions=incoming;aiState.lastRun={at:new Date().toISOString(),mode:result.mode||'api',instruction:$('aiInstruction').value.trim()};status.textContent=`${incoming.length} suggestion${incoming.length===1?'':'s'} added to the review queue${result.mode==='local-preview'?' in local preview mode':''}.`;status.className='source-upload-status ok';renderAiSuggestions();saveDraft();log(`${incoming.length} AI suggestion${incoming.length===1?'':'s'} ready for review`,'ok')}catch(e){status.textContent=e.message;status.className='source-upload-status error';log(`AI analysis failed: ${e.message}`,'error')}finally{btn.textContent='Generate suggestions';btn.disabled=!extractedSources.some(x=>x.aiIncluded&&x.text&&x.status!=='error')}}
async function regenerateActiveAiSuggestion(){const s=syncAiDialogToState();if(!s)return;const source=extractedSources.find(x=>x.id===s.sourceId)||extractedSources.find(x=>x.included&&x.text);if(!source)return $('aiSuggestionStatus').textContent='The original source is no longer available.';const btn=$('regenerateAiBtn');btn.disabled=true;btn.textContent='Regenerating…';try{const result=await requestAiSuggestions([source],$('aiInstruction').value.trim(),s);const fresh=(result.suggestions||[])[0];if(!fresh)throw new Error('No revised suggestion was returned.');const n=normaliseAiSuggestion(fresh);s.fields=n.fields;s.sourceExcerpt=n.sourceExcerpt;s.confidence=n.confidence;s.rationale=n.rationale;s.status='pending';$('aiSuggestionMeta').innerHTML=`<span>${esc(s.sourceName)}</span><span>${Math.round(s.confidence*100)}% confidence</span><span>regenerated</span>`;$('aiSourceExcerpt').textContent=s.sourceExcerpt;renderAiFieldReview();renderAiSuggestions();saveDraft();$('aiSuggestionStatus').textContent='Suggestion regenerated. Review the revised fields before approval.';$('aiSuggestionStatus').className='dialog-status ok'}catch(e){$('aiSuggestionStatus').textContent=e.message;$('aiSuggestionStatus').className='dialog-status warn'}finally{btn.disabled=false;btn.textContent='Regenerate'}}
function bindAiAssistance(){
  $('analyzeSourcesBtn').onclick=generateAiSuggestions;$('clearAiSuggestionsBtn').onclick=()=>{if(aiState.suggestions.length&&!confirm('Clear all AI suggestions? Applied newsletter content will remain unchanged.'))return;aiState.suggestions=[];aiState.activeSuggestionId=null;renderAiSuggestions();saveDraft();log('AI suggestion queue cleared')};
  $('reviewNextAiBtn').onclick=()=>{const n=aiPending()[0];if(n)openAiSuggestion(n.id)};$('applyAllAiBtn').onclick=()=>{const approved=aiState.suggestions.filter(x=>x.status==='approved');if(!approved.length)return;openAiSuggestion(approved[0].id);$('aiSuggestionStatus').textContent='Apply approved suggestions one at a time so each destination and conflict is reviewed.'};
  $('aiSuggestionTarget').onchange=()=>{const s=aiState.suggestions.find(x=>x.id===aiState.activeSuggestionId);if(s)s.targetUpdateId=$('aiSuggestionTarget').value;renderAiFieldReview()};$('approveAiBtn').onclick=approveAiSuggestion;$('applyAiNowBtn').onclick=()=>applyAiSuggestion(aiState.activeSuggestionId);$('rejectAiBtn').onclick=rejectAiSuggestion;$('regenerateAiBtn').onclick=regenerateActiveAiSuggestion;
  renderAiSuggestions();
}
bindAiAssistance();

// Release-candidate project handoff and diagnostics
const RELEASE_VERSION='1.0.0-rc1';
function safeFileStem(value){return String(value||'kommunicate-project').trim().replace(/[^a-z0-9_-]+/gi,'-').replace(/^-+|-+$/g,'').slice(0,80)||'kommunicate-project'}
function downloadBlobFile(name,content,type='application/octet-stream'){
  const blob=content instanceof Blob?content:new Blob([content],{type});
  const url=URL.createObjectURL(blob),a=document.createElement('a');a.href=url;a.download=name;document.body.appendChild(a);a.click();a.remove();setTimeout(()=>URL.revokeObjectURL(url),1000)
}
function exportProject(){
  if(!model)return log('Generate or start a newsletter before exporting a project','warn');
  const packageData={schema:'kommunicate-project',schemaVersion:1,appVersion:RELEASE_VERSION,exportedAt:new Date().toISOString(),model,extractedSources,mappingState,aiState,designState,ui:{issueTitle:$('issueTitle').value,primaryColor:$('primaryColor').value,accentColor:$('accentColor').value}};
  const stem=safeFileStem(model.issue?.issue_title||$('issueTitle').value);
  downloadBlobFile(`${stem}-project.json`,JSON.stringify(packageData,null,2),'application/json;charset=utf-8');
  log('Project file exported for team handoff','ok');
}
function validateProjectPackage(data){
  if(!data||data.schema!=='kommunicate-project')throw new Error('This is not a Kommunicate project file.');
  if(data.schemaVersion!==1)throw new Error(`Unsupported project schema version: ${data.schemaVersion}`);
  if(!data.model||typeof data.model!=='object')throw new Error('The project does not contain newsletter data.');
  if(data.extractedSources&&!Array.isArray(data.extractedSources))throw new Error('The project source list is invalid.');
  return true;
}
async function importProjectFile(file){
  try{
    const data=JSON.parse(await file.text());validateProjectPackage(data);
    if(model&&!confirm('Import this project and replace the current working session?'))return;
    model=data.model;extractedSources=data.extractedSources||[];mappingState=data.mappingState||{fieldMeta:{},history:[],activity:[]};aiState=data.aiState||{suggestions:[],activeSuggestionId:null,lastRun:null};designState={...defaultDesign,...(data.designState||{})};
    if(data.ui){$('issueTitle').value=data.ui.issueTitle||model.issue?.issue_title||'Compliance Newsletter';if(isHex(data.ui.primaryColor)){$('primaryColor').value=data.ui.primaryColor;$('primaryHex').value=data.ui.primaryColor.toUpperCase()}if(isHex(data.ui.accentColor)){$('accentColor').value=data.ui.accentColor;$('accentHex').value=data.ui.accentColor.toUpperCase()}}
    ensureModelIds();syncDesignControls();show();renderSourceList();renderAiSuggestions();saveDraft();log(`Project imported: ${file.name}`,'ok');
  }catch(error){log(`Project import failed: ${error.message}`,'error')}
}
function csvCell(value){const text=String(value??'');return `"${text.replace(/"/g,'""')}"`}
function downloadProvenanceReport(){
  if(!model)return log('No newsletter is available for a source report','warn');
  const rows=[['Newsletter field','Current value','Origin','Source','Confidence','Updated at','Overrode existing value']];
  ISSUE_FIELDS.forEach(([field])=>{const path=targetPathForIssue(field),value=getTargetValue(path),meta=metaFor(path)||{};if(value)rows.push([targetLabel(path),value,meta.origin||'unknown',meta.sourceName||'',meta.confidence==null?'':Math.round(meta.confidence*100)+'%',meta.updatedAt||'',meta.overrode?'Yes':'No'])});
  (model.updates||[]).forEach(update=>UPDATE_FIELDS.forEach(([field])=>{const path=targetPathForUpdate(update.id,field),value=getTargetValue(path),meta=metaFor(path)||{};if(value)rows.push([targetLabel(path),value,meta.origin||'unknown',meta.sourceName||'',meta.confidence==null?'':Math.round(meta.confidence*100)+'%',meta.updatedAt||'',meta.overrode?'Yes':'No'])}));
  const csv='\ufeff'+rows.map(row=>row.map(csvCell).join(',')).join('\r\n');downloadBlobFile('kommunicate-source-report.csv',csv,'text/csv;charset=utf-8');log('Source report downloaded','ok');
}
function readinessChecks(){
  const hosted=location.protocol==='https:'||location.hostname==='localhost';
  const selectedAi=extractedSources.filter(x=>x.aiIncluded&&x.text&&x.status!=='error');
  const extractionErrors=extractedSources.filter(x=>x.status==='error'||x.error);
  const pending=aiState.suggestions.filter(x=>['pending','approved'].includes(x.status));
  const issues=model?validate(model):['No newsletter has been generated'];
  const checks=[
    {status:model?'pass':'fail',title:'Newsletter data',detail:model?`${model.updates?.length||0} regulatory updates loaded.`:'Upload the standard Excel workbook or start without Excel.'},
    {status:issues.length?'warn':'pass',title:'Content validation',detail:issues.length?`${issues.length} validation item(s): ${issues.slice(0,2).join(' ')}`:'Required newsletter fields pass validation.'},
    {status:extractionErrors.length?'fail':'pass',title:'Source extraction',detail:extractionErrors.length?`${extractionErrors.length} source(s) have extraction errors.`:`${extractedSources.length} source(s) available; no extraction errors.`},
    {status:pending.length?'warn':'pass',title:'AI review queue',detail:pending.length?`${pending.length} suggestion(s) still require final review or application.`:'No unresolved AI suggestions.'},
    {status:hosted?'pass':'warn',title:'Deployment environment',detail:hosted?'Hosted environment detected. Netlify Functions can be used.':'Local file mode detected. AI uses preview logic; email and hosted review links require Netlify.'},
    {status:APP_CONFIG.emailEndpoint||hosted?'pass':'warn',title:'Email endpoint',detail:APP_CONFIG.emailEndpoint||'Default Netlify email function path will be used after deployment.'},
    {status:selectedAi.length<=1?'pass':'warn',title:'AI source selection',detail:selectedAi.length?`${selectedAi.length} source(s) currently selected for AI: ${selectedAi.map(x=>x.name).join(', ')}`:'No source is selected for AI. This is acceptable when AI mapping is not being used.'},
    {status:'pass',title:'Project portability',detail:'Project export/import and source-report download are available for team handoff.'}
  ];return checks
}
function renderReadiness(){
  const checks=readinessChecks(),container=$('readinessResults');container.innerHTML=checks.map(c=>`<div class="readiness-row ${c.status}"><span class="readiness-icon">${c.status==='pass'?'✓':c.status==='warn'?'!':'×'}</span><div><strong>${esc(c.title)}</strong><small>${esc(c.detail)}</small></div></div>`).join('');
  safeShowDialog($('readinessDialog'));
}
function bindReleaseTools(){
  $('exportProjectBtn').onclick=exportProject;$('importProjectInput').onchange=e=>{const file=e.target.files?.[0];if(file)importProjectFile(file);e.target.value=''};$('downloadProvenanceBtn').onclick=downloadProvenanceReport;$('runReadinessBtn').onclick=renderReadiness;$('rerunReadinessBtn').onclick=renderReadiness;
}
bindReleaseTools();
