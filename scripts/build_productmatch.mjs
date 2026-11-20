import fs from 'node:fs/promises';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { parseArgs } from 'node:util';
import { Workbook, SpreadsheetFile } from '@oai/artifact-tool';

const root=path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const {values:args}=parseArgs({options:{'data-root':{type:'string'},'output-dir':{type:'string'}}});
const dataRoot=path.resolve(args['data-root']||root);
const outputDir=path.resolve(args['output-dir']||process.env.PRODUCTMATCH_OUTPUT||path.join(root,'.artifacts/rebuild'));
await fs.mkdir(outputDir,{recursive:true});
const data=JSON.parse(await fs.readFile(path.join(dataRoot,'data/generated/demo_data.json'),'utf8'));
const cand=JSON.parse(await fs.readFile(path.join(dataRoot,'data/generated/candidate_snapshot.json'),'utf8'));
const wb=Workbook.create();
const names=['START HERE','DASHBOARD','FINAL CLEAN CATALOG','EXCEPTIONS','PRODUCT MATCHING','REVIEW DECISIONS','CLEANING','CORRECTIONS','DUPLICATE REVIEW','RESEARCH','MASTER CLEAN','MASTER RAW','SUPPLIER RAW','SETTINGS','DICTIONARY','QA EVALUATION','CANDIDATES'];
const sh=Object.fromEntries(names.map(n=>[n,wb.worksheets.add(n)]));
const C={navy:'#142B45',teal:'#087F8C',pale:'#E8F5F4',ink:'#243A4D',gray:'#EEF2F6',muted:'#5B7083',white:'#FFFFFF',input:'#FFF4D6',red:'#FCE8E7',amber:'#FFF0CA',line:'#D7E0E8'};
const SC=350,MC=300,RS=7,SE=356,ME=306;
const q=n=>`'${n}'!`;
const rg=(n,c,e=SE)=>`${q(n)}$${c}$7:$${c}$${e}`;
const cell=(s,a,v)=>sh[s].getRange(a).values=[[v]];
const formula=(s,a,v)=>sh[s].getRange(a).formulas=[[v]];
const block=(s,a,v)=>sh[s].getRange(a).write(v);
const font={name:'Arial',size:10,color:C.ink};
function base(n,last='S',rows=SE){
 const s=sh[n];s.showGridLines=false;s.tabColor=C.navy;
 s.getRange(`A1:${last}${rows}`).format={font,rowHeight:26,verticalAlignment:'center',columnWidth:17};
 console.log('Initialized '+n);
 s.getRange(`A1:${last}2`).format.fill=C.navy;s.getRange(`A1:${last}2`).format.font={...font,color:C.white,bold:true};
 s.getRange('A1').format.font={...font,size:16,color:C.white,bold:true};
 cell(n,'A1',n);formula(n,'A3',`=HYPERLINK("#'START HERE'!A1","← Start here")`);formula(n,'D3',`=HYPERLINK("#'DASHBOARD'!A1","Dashboard →")`);
 s.getRange('A3:D3').format.font={...font,color:C.teal};s.freezePanes.freezeRows(6);s.freezePanes.freezeColumns(1);
}
function title(n,text,sub){cell(n,'A1',text);cell(n,'A2',sub);}
function table(n,headers,rows,capacity){
 const s=sh[n],end=6+capacity,last=col(headers.length);
 block(n,'A6',[headers]);if(rows?.length)block(n,'A7',rows);
 const t=s.tables.add(`A6:${last}${end}`,true,'T_'+n.replace(/[^A-Za-z0-9]/g,''));t.style='TableStyleMedium2';t.showFilterButton=true;
 s.getRange(`A6:${last}6`).format={fill:C.navy,font:{...font,color:C.white,bold:true},rowHeight:38,wrapText:true,verticalAlignment:'center'};
 s.getRange(`A7:${last}${end}`).format.rowHeight=29;
 return t;
}
function col(n){let s='';while(n){let r=(n-1)%26;s=String.fromCharCode(65+r)+s;n=Math.floor((n-1)/26);}return s;}
function fc(n,c,fn,cap=SC){
 const values=Array.from({length:cap},(_,i)=>[fn(i+7)]);sh[n].getRange(`${c}7:${c}${6+cap}`).formulas=values;
}
function val(n,r,values){sh[n].getRange(r).dataValidation={rule:{type:'list',values}};}
function listref(n,r,ref){sh[n].getRange(r).dataValidation={rule:{type:'list',formula1:`INDIRECT("${ref}")`}};}
function input(n,r){sh[n].getRange(r).format.fill=C.input;sh[n].getRange(r).format.font={...font,color:'#235D86'};}
function cf(n,r,expr,fill=C.red){sh[n].getRange(r).conditionalFormats.addCustom(expr,{fill,font:{color:fill===C.pale?C.teal:'#8B3432'}});}
const trim=x=>`TRIM(CLEAN(SUBSTITUTE(${x},CHAR(160)," ")))`;
const upper=x=>`UPPER(${trim(x)})`;
const raw=(n,c,r)=>`${q(n)}${c}${r}`;
const vlookup=(key,n,last,colidx,cap=SC)=>`VLOOKUP(${key},${q(n)}$A$7:$${last}$${6+cap},${colidx},FALSE)`;
const lookup=(key,n,c,cap=SC,idcol='A')=>`INDEX(${rg(n,c,6+cap)},MATCH(${key},${rg(n,idcol,6+cap)},0))`;
const safeLookup=(key,n,c,cap=SC,idcol='A')=>`IF(COUNTIFS(${rg(n,idcol,6+cap)},${key})=1,IF(${lookup(key,n,c,cap,idcol)}="","",${lookup(key,n,c,cap,idcol)}),"")`;
const corrected=(r,ci,source)=>{const orig=`IF(${source}="","",${source})`;return `IF(COUNTIFS(${rg('CORRECTIONS','A')},$A${r})=1,IF(${vlookup(`$A${r}`,'CORRECTIONS','O',13)}<>"",IF(${vlookup(`$A${r}`,'CORRECTIONS','O',ci)}="",${orig},${vlookup(`$A${r}`,'CORRECTIONS','O',ci)}),${orig}),${orig})`};
const brand=x=>`IF(${trim(x)}="","",IF(COUNTIFS('SETTINGS'!$D$7:$D$56,LOWER(${trim(x)}))>0,VLOOKUP(LOWER(${trim(x)}),'SETTINGS'!$D$7:$E$56,2,FALSE),PROPER(${trim(x)})))`;
const category=x=>`IF(${trim(x)}="","",IF(COUNTIFS('SETTINGS'!$G$7:$G$36,LOWER(${trim(x)}))>0,VLOOKUP(LOWER(${trim(x)}),'SETTINGS'!$G$7:$H$36,2,FALSE),PROPER(${trim(x)})))`;
console.log('Initializing sheets');
for(const n of names)base(n,n==='PRODUCT MATCHING'?'AQ':n==='CLEANING'?'AF':n==='MASTER CLEAN'?'AE':n==='FINAL CLEAN CATALOG'?'X':n==='DASHBOARD'?'W':n==='RESEARCH'?'P':n==='START HERE'?'J':n==='SETTINGS'?'Q':n==='DICTIONARY'?'D':n==='QA EVALUATION'?'M':'S', n==='START HERE'?65:n==='DASHBOARD'?57:n==='RESEARCH'?106:n==='SETTINGS'?70:n==='DICTIONARY'?100:SE);

console.log('Stage: Configuration');
// Configuration and data-backed validation lists.
title('SETTINGS','Settings & matching policy','Amber cells are configurable. Rules and supported capacities are documented below.');
block('SETTINGS','A6',[['Parameter','Value']]);block('SETTINGS','A7',[
 ['Fuzzy suggestion minimum',0.72],['Fuzzy ambiguity gap',0.08],['Master capacity',MC],['Supplier capacity',SC],['Required fields per row',10],['Snapshot generated UTC',cand.generatedUTC],['Dataset disclosure','Synthetic transaction scenario + 15 real models'],['Automatic approval','None; all exports need a reviewer decision'],['Shared identifier tier','Brand-scoped exact MPN; no SKU cross-match'],['GTIN policy','No GTIN supplied as verified. Not used for matching.'],['Duplicate survivor','Earliest raw row within exact retransmission key'],['Supported research fields','Specification (live apply); other fields are evidence only'],['Append behavior','Paste raw rows into blank reserved rows; enter stable IDs in review input sheets'],['Sorting','Sort complete tables; refresh fuzzy after reordering MASTER RAW, even if signatures show Current'],['Fuzzy refresh','Python standard library; scores are stored snapshots'],['Research capacity',100],['No historical audit','Current-state inputs; preserve dated file copies for history']]);
sh.SETTINGS.getRange('A7:A23').format.columnWidth=40;sh.SETTINGS.getRange('B7:B23').format.columnWidth=54;sh.SETTINGS.getRange('A7:B23').format.wrapText=true;sh.SETTINGS.getRange('A7:B23').format.rowHeight=43;input('SETTINGS','B7:B8');sh.SETTINGS.getRange('B7:B8').setNumberFormat('0.00');
block('SETTINGS','D6',[['Brand alias (lower case)','Canonical brand']]);block('SETTINGS','D7',Object.entries(data.aliases));
block('SETTINGS','G6',[['Category alias','Canonical category']]);block('SETTINGS','G7',[...data.categories.map(c=>[c.toLowerCase(),c]),['office misc','Office accessories']]);
block('SETTINGS','J6',[['Categories'],['All'],...data.categories.map(x=>[x])]);
const brands=[...new Set(data.master.map(m=>data.aliases[m[4].trim().toLowerCase()]||m[4]))].sort();
block('SETTINGS','K6',[['Brands'],['All'],...brands.map(x=>[x])]);
const outcomes=['Accepted','Needs review','Unmatched','Rejected','Excluded duplicate'];block('SETTINGS','L6',[['Review outcomes'],['All'],...outcomes.map(x=>[x])]);
block('SETTINGS','N6',[['Valid master IDs']]);fc('SETTINGS','N',r=>`=IF('MASTER RAW'!B${r}="","",'MASTER RAW'!B${r})`,MC);
block('SETTINGS','O6',[['Valid supplier IDs']]);fc('SETTINGS','O',r=>`=IF('SUPPLIER RAW'!A${r}="","",'SUPPLIER RAW'!A${r})`);
block('SETTINGS','A27',[
 ['Matching hierarchy','Rule / effect'],['1. Exact MPN','Brand + nonblank real MPN, retaining all punctuation; variants and other identifiers checked afterward.'],['2. Brand + model','Prefer full variant key; fall back to brand/model candidates for review if variant is missing or different.'],['3. Attributes','Exact normalized name + brand + category + full variant attributes; model may be missing.'],['4. Fuzzy suggestion','Same category and canonical brand (or missing brand). Top three name scores. No automatic approval.'],['Minimum score','Increasing B7 hides weaker snapshot suggestions immediately; it does not recompute scores.'],['Ambiguity gap','A best-minus-second score below B8 flags ambiguity. Applies only to fuzzy tier.'],['Duplicate key (supplier)','Clean SKU + normalized name + brand/model + color/capacity/size/pack + price + currency. All identity fields required.'],['Suspected duplicate','Same full brand/model/variant key across different supplier SKUs. Retain until a reviewer chooses Keep separate or confirms an earlier survivor.'],['Duplicate key (master)','Brand/model + color/capacity/size/pack. Earliest source row survives; genuine variants have separate keys.'],['Required fields','Name, brand, category, model, color, capacity, size/version, pack quantity, price, currency. N/A is an explicit nonapplicable attribute.'],['Completeness','Populated required cells / (10 × supplier rows in scope); not a validity or accuracy metric. After uses final master attributes for accepted rows.'],['Conflicts block export','Conflicting model/MPN/GTIN/brand/category/variant, missing variant, invalid price/pack/currency, ambiguous candidate or stale accepted fuzzy snapshot. Correct data or choose a compatible master.'],['Approved research','Specification only; Verified + Yes + exact model + source/title/note/date + unique field key. Other attributes remain evidence queue.'],['Recalculation','All native formulas, selectors, search and review outputs recalculate automatically in Excel Automatic mode. CSVs are snapshots.'],['Fuzzy staleness','Full supplier signatures and full master signatures are compared to snapshot; any master changes invalidate all fuzzy candidates.'],['Capacity extension','Reserved rows: master 300, supplier 350, research 100. Beyond capacity extend ranges/tables/validations with developer review.']]);
sh.SETTINGS.getRange('A27:B43').format.wrapText=true;sh.SETTINGS.getRange('A27:B43').format.rowHeight=67;sh.SETTINGS.getRange('A6:Q6').format.fill=C.navy;sh.SETTINGS.getRange('A6:Q6').format.font={...font,color:C.white,bold:true};

console.log('Stage: Immutable');
// Immutable raw snapshots and stable-key editable controls.
for(const [n,rows,cap] of [['MASTER RAW',data.master,MC],['SUPPLIER RAW',data.supplier,SC]]){
 title(n,n==='MASTER RAW'?'Master catalog • original snapshot':'Supplier feed • original snapshot','Original values preserved. Append only within reserved capacity; keep source IDs unique. Prices are simulated IDR offers.');
 table(n,data.rawHeaders,rows,cap);sh[n].getRange('A7:M'+(cap+6)).setNumberFormat('@');sh[n].getRange('M7:N'+(cap+6)).setNumberFormat('#,##0');sh[n].getRange('D6:D'+(cap+6)).format.columnWidth=49;sh[n].getRange('R6:S'+(cap+6)).format.columnWidth=43;sh[n].getRange('Q6:Q'+(cap+6)).format.columnWidth=30;
}
title('REVIEW DECISIONS','Review decisions • keyed manual inputs','Seeded decisions are simulated. C stores the approved ID as a value (copy current suggestion for Accept). Edit B:F. Decisions persist across recalculation and fuzzy refresh.');
table('REVIEW DECISIONS',['Supplier row ID','Decision','Approved master ID','Reviewer','Review date','Decision notes','Final selected ID','Current outcome','Current suggestion','Input check'],data.decisions.map(r=>[...r.slice(0,4),r[4]?new Date(r[4]+'T00:00:00Z'):null,r[5]]),SC);
input('REVIEW DECISIONS','A7:F356');val('REVIEW DECISIONS','B7:B356',['Accept suggested match','Select different master','Reject suggestion','Leave pending']);listref('REVIEW DECISIONS','A7:A356',"'SETTINGS'!$O$7:$O$356");listref('REVIEW DECISIONS','C7:C356',"'SETTINGS'!$N$7:$N$306");sh['REVIEW DECISIONS'].getRange('E7:E356').setNumberFormat('yyyy-mm-dd');sh['REVIEW DECISIONS'].getRange('B6:B356').format.columnWidth=29;sh['REVIEW DECISIONS'].getRange('F6:F356').format.columnWidth=68;
title('CORRECTIONS','Cleaning corrections • preserve the raw snapshot','Enter only corrected fields, plus a reason. Blank means use original; source IDs are validated. These are current-state overrides.');
table('CORRECTIONS',['Supplier row ID','Name override','Brand override','Category override','Model override','MPN override','Color override','Capacity override','Size override','Pack override','Price override','Currency override','Reason required','Updated by','Updated date'],data.corrections.map(r=>[...r.slice(0,14),r[14]?new Date(r[14]+'T00:00:00Z'):null]),SC);
input('CORRECTIONS','A7:O356');listref('CORRECTIONS','A7:A356',"'SETTINGS'!$O$7:$O$356");sh.CORRECTIONS.getRange('O7:O356').setNumberFormat('yyyy-mm-dd');sh.CORRECTIONS.getRange('M6:M356').format.columnWidth=52;
title('DUPLICATE REVIEW','Duplicate review • retain or exclude','Automatic exact retransmissions exclude later raw rows. Suspected duplicates need a decision, reason, reviewer and date.');
table('DUPLICATE REVIEW',['Supplier row ID','Duplicate decision','Reason','Reviewer','Review date','Earlier survivor row ID','Detected duplicate state','Current exclusion'],data.duplicates.map(r=>[...r.slice(0,4),r[4]?new Date(r[4]+'T00:00:00Z'):null]),SC);
input('DUPLICATE REVIEW','A7:F356');val('DUPLICATE REVIEW','B7:B356',['Pending','Keep separate','Confirm duplicate']);listref('DUPLICATE REVIEW','A7:A356',"'SETTINGS'!$O$7:$O$356");listref('DUPLICATE REVIEW','F7:F356',"'SETTINGS'!$O$7:$O$356");sh['DUPLICATE REVIEW'].getRange('E7:E356').setNumberFormat('yyyy-mm-dd');sh['DUPLICATE REVIEW'].getRange('C6:C356').format.columnWidth=64;

console.log('Stage: Field-level');
// Field-level research, with a conservative apply gate.
title('RESEARCH','Research & verification • field-level evidence','15 real models researched on 11 Nov 2026. Only eligible, uniquely approved Specification rows feed the catalog.');
table('RESEARCH',['Research ID','Product ID','Exact model','Attribute','Original value','Proposed value','Verification status','Apply approved','Evidence note','Source URL','Page title','Date accessed','Other source URL','Conflict / workflow notes','Apply key (formula)','Apply eligibility (formula)'],data.research.map(r=>[...r.slice(0,11),r[11]?new Date(r[11]+'T00:00:00Z'):null,...r.slice(12)]),100);
input('RESEARCH','A7:N106');val('RESEARCH','G7:G106',['Verified','Not found','Conflicting sources','Pending']);val('RESEARCH','H7:H106',['Yes','No']);listref('RESEARCH','B7:B106',"'SETTINGS'!$N$7:$N$306");sh.RESEARCH.getRange('L7:L106').setNumberFormat('yyyy-mm-dd');
for(const [c,w] of [['C',25],['F',44],['I',80],['J',60],['K',48],['M',60],['N',65]])sh.RESEARCH.getRange(`${c}6:${c}106`).format.columnWidth=w;
sh.RESEARCH.getRange('A7:P106').format.wrapText=true;sh.RESEARCH.getRange('A7:P106').format.rowHeight=76;
fc('RESEARCH','P',r=>`=IF(A${r}="","",IF(AND(G${r}="Verified",H${r}="Yes",D${r}="Specification",F${r}<>"",I${r}<>"",J${r}<>"",K${r}<>"",ISNUMBER(L${r}),COUNTIFS(${rg('MASTER RAW','B',ME)},B${r})=1,UPPER(C${r})=UPPER(${safeLookup(`B${r}`,'MASTER RAW','G',MC,'B')})),"Eligible","Not applied"))`,100);
fc('RESEARCH','O',r=>`=IF(P${r}="Eligible",B${r}&"|"&D${r},"")`,100);
cell('RESEARCH','Q6','Open source');fc('RESEARCH','Q',r=>`=IF(J${r}="","",HYPERLINK(J${r},"Open official source"))`,100);
cf('RESEARCH','G7:H106','$G7="Conflicting sources"');cf('RESEARCH','G7:H106','$G7="Verified"',C.pale);

console.log('Stage: Frozen external');
// Frozen external candidate scores; live signatures expose stale snapshots.
title('CANDIDATES','Fuzzy candidates • external snapshot','Python-generated top three name candidates. Static scores; native matching and reviewer decisions remain live.');
table('CANDIDATES',['Supplier row ID','Supplier signature','Rank 1 ID','Score 1','Rank 2 ID','Score 2','Rank 3 ID','Score 3','Generated UTC'],cand.rows,SC);
block('CANDIDATES','L6',[['Master ID snapshot','Master signature snapshot']]);block('CANDIDATES','L7',cand.masterSnapshot);sh.CANDIDATES.getRange('D7:H356').setNumberFormat('0.000');
sh.CANDIDATES.getRange('B6:B356').format.columnWidth=90;sh.CANDIDATES.getRange('M6:M306').format.columnWidth=90;

console.log('Stage: Master cleaning');
// Master cleaning owns standardized catalog attributes and duplicate eligibility.
const mh=['Source row ID','Master ID','Internal SKU','Raw product name','Standard product name','Canonical brand','Category','Model','MPN','GTIN (text)','Color','Capacity','Size / version','Pack','Price','Currency','Approved specification','Disclosure','Master provenance','Official source URL','Identity key','Duplicate group size','Exclusion reason','Active identity key','Active MPN key','Live fuzzy signature','Fuzzy snapshot state','Required fields missing','Active name / attribute key','Active brand / model key','Quality flags'];
title('MASTER CLEAN','Master cleaning • canonical catalog','Full variant identity prevents collapsing colors, capacity, size/version and pack quantities. Research applies conservatively.');table('MASTER CLEAN',mh,[],MC);
for(const [c,src] of [['A','A'],['B','B'],['C','C'],['D','D'],['R','Q'],['S','R'],['T','S']])fc('MASTER CLEAN',c,r=>`=IF('MASTER RAW'!A${r}="","",'MASTER RAW'!${src}${r})`,MC);
for(const [c,expr] of [['E',r=>`PROPER(${trim(raw('MASTER RAW','D',r))})`],['F',r=>brand(raw('MASTER RAW','E',r))],['G',r=>category(raw('MASTER RAW','F',r))],['H',r=>upper(raw('MASTER RAW','G',r))],['I',r=>upper(raw('MASTER RAW','H',r))],['J',r=>`SUBSTITUTE(SUBSTITUTE(${trim(raw('MASTER RAW','I',r))}," ",""),"-","")`],['K',r=>`PROPER(${trim(raw('MASTER RAW','J',r))})`],['L',r=>`SUBSTITUTE(${upper(raw('MASTER RAW','K',r))}," ","")`],['M',r=>upper(raw('MASTER RAW','L',r))],['N',r=>`IF('MASTER RAW'!M${r}="","",'MASTER RAW'!M${r})`],['O',r=>`IF('MASTER RAW'!N${r}="","",'MASTER RAW'!N${r})`],['P',r=>upper(raw('MASTER RAW','O',r))]])fc('MASTER CLEAN',c,r=>`=IF(A${r}="","",${expr(r)})`,MC);
fc('MASTER CLEAN','Q',r=>`=IF(A${r}="","",IF(COUNTIFS('RESEARCH'!$O$7:$O$106,B${r}&"|Specification")=1,INDEX('RESEARCH'!$F$7:$F$106,MATCH(B${r}&"|Specification",'RESEARCH'!$O$7:$O$106,0)),IF('MASTER RAW'!P${r}="","",'MASTER RAW'!P${r})))`,MC);
fc('MASTER CLEAN','U',r=>`=IF(OR(A${r}="",F${r}="",H${r}="",K${r}="",L${r}="",M${r}="",N${r}=""),"",F${r}&"|"&H${r}&"|"&UPPER(K${r})&"|"&L${r}&"|"&M${r}&"|"&N${r})`,MC);
fc('MASTER CLEAN','V',r=>`=IF(U${r}="","",COUNTIFS(${rg('MASTER CLEAN','U',ME)},U${r}))`,MC);
fc('MASTER CLEAN','W',r=>`=IF(A${r}="","",IF(COUNTIFS($B$7:$B$${ME},B${r})<>1,"Duplicate master ID",IF(AND(U${r}<>"",COUNTIFS($U$7:U${r},U${r})>1),"Exact master duplicate","")))`,MC);
fc('MASTER CLEAN','X',r=>`=IF(W${r}="",U${r},"")`,MC);
fc('MASTER CLEAN','Y',r=>`=IF(AND(A${r}<>"",W${r}="",I${r}<>"",F${r}<>""),F${r}&"|"&I${r},"")`,MC);
fc('MASTER CLEAN','Z',r=>`=IF(A${r}="","",UPPER(E${r})&"|"&F${r}&"|"&G${r}&"|"&H${r}&"|"&I${r}&"|"&UPPER(K${r})&"|"&L${r}&"|"&M${r}&"|"&N${r})`,MC);
fc('MASTER CLEAN','AA',r=>`=IF(A${r}="","",IF(COUNTIFS('CANDIDATES'!$L$7:$L$306,B${r},'CANDIDATES'!$M$7:$M$306,Z${r})=1,"Current","Stale"))`,MC);
fc('MASTER CLEAN','AB',r=>`=IF(A${r}="","",10-((E${r}<>"")+(F${r}<>"")+(G${r}<>"")+(H${r}<>"")+(K${r}<>"")+(L${r}<>"")+(M${r}<>"")+(N${r}<>"")+(O${r}<>"")+(P${r}<>"")))`,MC);
fc('MASTER CLEAN','AC',r=>`=IF(AND(A${r}<>"",W${r}="",E${r}<>"",F${r}<>"",K${r}<>"",L${r}<>"",M${r}<>"",N${r}<>""),F${r}&"|"&G${r}&"|"&UPPER(E${r})&"|"&UPPER(K${r})&"|"&L${r}&"|"&M${r}&"|"&N${r},"")`,MC);
fc('MASTER CLEAN','AD',r=>`=IF(AND(A${r}<>"",W${r}="",F${r}<>"",H${r}<>""),F${r}&"|"&H${r},"")`,MC);
fc('MASTER CLEAN','AE',r=>`=IF(A${r}="","",IF(AB${r}>0,"Missing required fields; ","")&W${r})`,MC);
sh['MASTER CLEAN'].getRange('D6:E306').format.columnWidth=49;sh['MASTER CLEAN'].getRange('Q6:Q306').format.columnWidth=53;

console.log('Stage: Supplier cleaning');
// Supplier cleaning: stable-key overrides, exact retransmissions and suspected variants.
const ch=['Supplier row ID','Supplier SKU','Raw product name','Standard product name','Raw brand','Canonical brand','Category','Model','MPN','GTIN (text)','Color','Capacity','Size / version','Pack','Unit price','Currency','Full identity key','Exact retransmission key','Exact group size','Suspected duplicate','Duplicate decision','Exclusion reason','Required fields missing','Invalid values','Issue queue','Live fuzzy signature','Before complete fields','MPN key','Brand / model key','Name / attribute key','Exact survivor source ID','Correction check'];
title('CLEANING','Cleaning & duplicate detection','Raw and standardized values remain visible together. Filter Issue queue; enter resolutions in CORRECTIONS or DUPLICATE REVIEW.');table('CLEANING',ch,[],SC);
fc('CLEANING','A',r=>`=IF('SUPPLIER RAW'!A${r}="","",'SUPPLIER RAW'!A${r})`);
fc('CLEANING','B',r=>`=IF(A${r}="","",${trim(raw('SUPPLIER RAW','C',r))})`);
fc('CLEANING','C',r=>`=IF(A${r}="","",'SUPPLIER RAW'!D${r})`);fc('CLEANING','E',r=>`=IF(OR(A${r}="",'SUPPLIER RAW'!E${r}=""),"",'SUPPLIER RAW'!E${r})`);
for(const [c,ci,ri,transform] of [['D',2,'D',x=>`PROPER(${trim(x)})`],['F',3,'E',brand],['G',4,'F',category],['H',5,'G',upper],['I',6,'H',upper],['K',7,'J',x=>`PROPER(${trim(x)})`],['L',8,'K',x=>`SUBSTITUTE(${upper(x)}," ","")`],['M',9,'L',upper],['N',10,'M',x=>`IF(${x}="","",${x})`],['O',11,'N',x=>`IF(${x}="","",${x})`],['P',12,'O',upper]])fc('CLEANING',c,r=>`=IF(A${r}="","",${transform(corrected(r,ci,raw('SUPPLIER RAW',ri,r)))})`);
fc('CLEANING','J',r=>`=IF(A${r}="","",IF(OR(UPPER('SUPPLIER RAW'!I${r})="NOT_PROVIDED",UPPER('SUPPLIER RAW'!I${r})="N/A"),"",SUBSTITUTE(SUBSTITUTE(${trim(raw('SUPPLIER RAW','I',r))}," ",""),"-","")))`);
fc('CLEANING','Q',r=>`=IF(OR(A${r}="",F${r}="",H${r}="",K${r}="",L${r}="",M${r}="",N${r}=""),"",F${r}&"|"&H${r}&"|"&UPPER(K${r})&"|"&L${r}&"|"&M${r}&"|"&N${r})`);
fc('CLEANING','R',r=>`=IF(OR(Q${r}="",B${r}="",D${r}="",O${r}="",P${r}=""),"",B${r}&"|"&UPPER(D${r})&"|"&Q${r}&"|"&O${r}&"|"&P${r})`);
fc('CLEANING','S',r=>`=IF(R${r}="","",COUNTIFS($R$7:$R$356,R${r}))`);
fc('CLEANING','T',r=>`=IF(A${r}="","",IF(AND(Q${r}<>"",COUNTIFS($Q$7:$Q$356,Q${r})>COUNTIFS($Q$7:$Q$356,Q${r},$B$7:$B$356,B${r})),"Suspected",""))`);
fc('CLEANING','U',r=>`=IF(A${r}="","",IF(COUNTIFS(${rg('DUPLICATE REVIEW','A')},A${r})=1,${vlookup(`A${r}`,'DUPLICATE REVIEW','H',2)},"Pending"))`);
const dupField=(r,c)=>safeLookup(`A${r}`,'DUPLICATE REVIEW',c);
fc('CLEANING','V',r=>`=IF(A${r}="","",IF(AND(R${r}<>"",COUNTIFS($R$7:R${r},R${r})>1),"Exact retransmission",IF(AND(U${r}="Confirm duplicate",${dupField(r,'C')}<>"",${dupField(r,'D')}<>"",ISNUMBER(${dupField(r,'E')}),${dupField(r,'F')}<>"",${dupField(r,'F')}<A${r},Q${r}<>"",COUNTIFS($A$7:$A$356,${dupField(r,'F')},$Q$7:$Q$356,Q${r})=1),"Confirmed by reviewer","")))`);
fc('CLEANING','W',r=>`=IF(A${r}="","",10-((D${r}<>"")+(F${r}<>"")+(G${r}<>"")+(H${r}<>"")+(K${r}<>"")+(L${r}<>"")+(M${r}<>"")+(N${r}<>"")+(O${r}<>"")+(P${r}<>"")))`);
fc('CLEANING','X',r=>`=IF(A${r}="","",IF(OR(N${r}="",NOT(ISNUMBER(N${r})),N${r}<=0,IF(ISNUMBER(N${r}),MOD(N${r},1),-1)<>0),"Invalid pack; ","")&IF(OR(O${r}="",NOT(ISNUMBER(O${r})),O${r}<=0),"Invalid price; ","")&IF(P${r}<>"IDR","Invalid currency; ","")&IF(COUNTIFS('SETTINGS'!$J$8:$J$16,G${r})=0,"Unmapped category; ",""))`);
fc('CLEANING','Y',r=>`=IF(A${r}="","",IF(W${r}>0,"Missing required; ","")&X${r}&IF(T${r}="Suspected","Suspected duplicate; ","")&IF(S${r}>1,"Exact group; ","")&IF(COUNTIFS($A$7:$A$356,A${r})<>1,"Duplicate source ID; ","")&AF${r})`);
fc('CLEANING','Z',r=>`=IF(A${r}="","",UPPER(D${r})&"|"&F${r}&"|"&G${r}&"|"&H${r}&"|"&I${r}&"|"&UPPER(K${r})&"|"&L${r}&"|"&M${r}&"|"&N${r})`);
const rawRequired=['D','E','F','G','J','K','L','M','N','O'];
fc('CLEANING','AA',r=>`=IF(A${r}="","",${rawRequired.map(c=>`(${trim(raw('SUPPLIER RAW',c,r))}<>"")`).join('+')})`);
fc('CLEANING','AB',r=>`=IF(AND(A${r}<>"",I${r}<>"",F${r}<>""),F${r}&"|"&I${r},"")`);
fc('CLEANING','AC',r=>`=IF(AND(A${r}<>"",F${r}<>"",H${r}<>""),F${r}&"|"&H${r},"")`);
fc('CLEANING','AD',r=>`=IF(AND(A${r}<>"",D${r}<>"",F${r}<>"",K${r}<>"",L${r}<>"",M${r}<>"",N${r}<>""),F${r}&"|"&G${r}&"|"&UPPER(D${r})&"|"&UPPER(K${r})&"|"&L${r}&"|"&M${r}&"|"&N${r},"")`);
fc('CLEANING','AE',r=>`=IF(R${r}="","",INDEX($A$7:$A$356,MATCH(R${r},$R$7:$R$356,0)))`);
fc('CLEANING','AF',r=>`=IF(A${r}="","",IF(COUNTIFS(${rg('CORRECTIONS','A')},A${r})>1,"Duplicate correction key; ",IF(AND(COUNTIFS(${rg('CORRECTIONS','A')},A${r})=1,${safeLookup(`A${r}`,'CORRECTIONS','M')}=""),"Correction reason missing; ","")))`);
for(const [c,w] of [['C',48],['D',48],['Y',62],['V',27],['X',40]])sh.CLEANING.getRange(`${c}6:${c}356`).format.columnWidth=w;
cf('CLEANING','W7:W356','$W7>0');cf('CLEANING','X7:Y356','AND($A7<>"",$X7<>"")');cf('CLEANING','T7:V356','$V7<>""',C.amber);

console.log('Stage: Operational matching');
// Operational matching. Each helper is inspectable; reviewer inputs are looked up by source ID.
const ph=['Supplier row ID','Supplier product','Brand','Category','Suggested master ID','Suggested master product','Method','Name similarity','Candidate count','Suggestion conflicts','Review outcome','Final product ID','Explanation','Decision','Reviewer','Review date','Decision notes','Duplicate state','Required fields missing','Remaining issues','Export ready','MPN count','Full model / variant count','Brand / model count','Attribute count','Native suggested ID','Fuzzy snapshot state','Fuzzy suggested ID','Suggested master position','Selected master position','Before complete fields','After complete fields','Verification status','Export sequence','Exception sequence','Filter included','Search included','Search sequence','Decision input check','Selected conflicts','Candidate ambiguity','Selected ID validity','Duplicate resolution'];
title('PRODUCT MATCHING','Product matching • operational review','One row per supplier record. Suggestions are separate from decisions. Filter outcomes; resolve conflicts before export.');table('PRODUCT MATCHING',ph,[],SC);
for(const [c,src] of [['A','A'],['B','D'],['C','F'],['D','G'],['R','V'],['S','W'],['AE','AA']])fc('PRODUCT MATCHING',c,r=>`=IF('CLEANING'!A${r}="","",'CLEANING'!${src}${r})`);
for(const [c,src,target] of [['V','AB','Y'],['W','Q','X'],['X','AC','AD'],['Y','AD','AC']])fc('PRODUCT MATCHING',c,r=>`=IF(A${r}="","",IF('CLEANING'!${src}${r}="",0,COUNTIFS(${rg('MASTER CLEAN',target,ME)},'CLEANING'!${src}${r})))`);
const matchMaster=(r,suppliercol,mastercol)=>`INDEX(${rg('MASTER CLEAN','B',ME)},MATCH('CLEANING'!${suppliercol}${r},${rg('MASTER CLEAN',mastercol,ME)},0))`;
fc('PRODUCT MATCHING','Z',r=>`=IF(A${r}="","",IF(V${r}>0,${matchMaster(r,'AB','Y')},IF(W${r}>0,${matchMaster(r,'Q','X')},IF(X${r}>0,${matchMaster(r,'AC','AD')},IF(Y${r}>0,${matchMaster(r,'AD','AC')},"")))))`);
fc('PRODUCT MATCHING','AA',r=>`=IF(A${r}="","",IF(AND(COUNTIFS(${rg('CANDIDATES','A')},A${r},${rg('CANDIDATES','B')},'CLEANING'!Z${r})=1,COUNTIFS(${rg('MASTER CLEAN','AA',ME)},"Stale")=0,COUNTIFS('CANDIDATES'!$L$7:$L$306,"<>")=SUMPRODUCT(--(${rg('MASTER CLEAN','A',ME)}<>""))),"Current","Refresh required"))`);
const candget=(r,c)=>safeLookup(`A${r}`,'CANDIDATES',c);
fc('PRODUCT MATCHING','AB',r=>`=IF(A${r}="","",IF(AND(AA${r}="Current",${candget(r,'D')}>='SETTINGS'!$B$7),${candget(r,'C')},""))`);
fc('PRODUCT MATCHING','E',r=>`=IF(A${r}="","",IF(Z${r}<>"",Z${r},AB${r}))`);
fc('PRODUCT MATCHING','AC',r=>`=IF(E${r}="","",MATCH(E${r},${rg('MASTER CLEAN','B',ME)},0))`);
const mi=(r,c,pos='AC')=>`INDEX(${rg('MASTER CLEAN',c,ME)},${pos}${r})`;
fc('PRODUCT MATCHING','F',r=>`=IF(E${r}="","",${mi(r,'E')})`);
fc('PRODUCT MATCHING','G',r=>`=IF(A${r}="","",IF(V${r}>0,"Exact MPN",IF(W${r}>0,"Brand + model + variant",IF(X${r}>0,"Brand + model",IF(Y${r}>0,"Attributes",IF(AB${r}<>"","Fuzzy candidate","No candidate"))))))`);
fc('PRODUCT MATCHING','H',r=>`=IF(G${r}="Fuzzy candidate",${candget(r,'D')},"")`);
fc('PRODUCT MATCHING','I',r=>`=IF(A${r}="","",IF(V${r}>0,V${r},IF(W${r}>0,W${r},IF(X${r}>0,X${r},IF(Y${r}>0,Y${r},IF(AB${r}<>"",(${candget(r,'D')}>='SETTINGS'!$B$7)+(${candget(r,'F')}>='SETTINGS'!$B$7)+(${candget(r,'H')}>='SETTINGS'!$B$7),0))))))`);
function conflict(r,pos){
 const pairs=[['F','F','Brand'],['G','G','Category'],['H','H','Model'],['I','I','MPN'],['J','J','GTIN'],['K','K','Color'],['L','L','Capacity'],['M','M','Size/version'],['N','N','Pack']];
 return pairs.map(([s,m,l])=>`IF(AND('CLEANING'!${s}${r}<>"",${mi(r,m,pos)}<>"",UPPER('CLEANING'!${s}${r}&"")<>UPPER(${mi(r,m,pos)}&"")),"${l}; ","")`).join('&')+`&IF(OR('CLEANING'!K${r}="",'CLEANING'!L${r}="",'CLEANING'!M${r}="",'CLEANING'!N${r}="",'CLEANING'!F${r}=""),"Missing brand/variant; ","")`;
}
fc('PRODUCT MATCHING','J',r=>`=IF(E${r}="","",${conflict(r,'AC')})`);
for(const [c,src] of [['N','B'],['O','D'],['P','E'],['Q','F']])fc('PRODUCT MATCHING',c,r=>`=IF(A${r}="","",${safeLookup(`A${r}`,'REVIEW DECISIONS',src)})`);
fc('PRODUCT MATCHING','L',r=>`=IF(A${r}="","",IF(OR(N${r}="Accept suggested match",N${r}="Select different master"),${safeLookup(`A${r}`,'REVIEW DECISIONS','C')},""))`);
fc('PRODUCT MATCHING','AP',r=>`=IF(L${r}="","",IF(COUNTIFS(${rg('MASTER CLEAN','B',ME)},L${r})<>1,"Invalid master ID",IF(${safeLookup(`L${r}`,'MASTER CLEAN','W',MC,'B')}<>"","Excluded master ID","Valid")))`);
fc('PRODUCT MATCHING','AD',r=>`=IF(AP${r}="Valid",MATCH(L${r},${rg('MASTER CLEAN','B',ME)},0),"")`);
fc('PRODUCT MATCHING','AN',r=>`=IF(AD${r}="","",${conflict(r,'AD')})`);
fc('PRODUCT MATCHING','AO',r=>`=IF(A${r}="","",IF(G${r}="Fuzzy candidate",IF(${candget(r,'D')}-${candget(r,'F')}<'SETTINGS'!$B$8,"Ambiguous fuzzy candidates",""),IF(I${r}>1,"Ambiguous native candidates","")))`);
fc('PRODUCT MATCHING','AQ',r=>`=IF(A${r}="","",IF('CLEANING'!T${r}<>"Suspected","Not required",IF(AND('CLEANING'!U${r}="Keep separate",${dupField(r,'C')}<>"",${dupField(r,'D')}<>"",ISNUMBER(${dupField(r,'E')})),"Resolved","Pending duplicate review")))`);
fc('PRODUCT MATCHING','AM',r=>`=IF(A${r}="","",IF(COUNTIFS(${rg('REVIEW DECISIONS','A')},A${r})>1,"Duplicate decision key; ","")&IF(COUNTIFS(${rg('CLEANING','A')},A${r})<>1,"Duplicate source ID; ","")&IF(OR(N${r}="Accept suggested match",N${r}="Select different master"),IF(OR(O${r}="",NOT(ISNUMBER(P${r}))),"Reviewer/date required; ","")&IF(AP${r}<>"Valid","Valid master required; ","")&IF(AND(N${r}="Accept suggested match",G${r}="Fuzzy candidate",Q${r}=""),"Fuzzy review note required; ","")&IF(AND(N${r}="Accept suggested match",AO${r}<>""),AO${r}&"; ",""),""))`);
fc('PRODUCT MATCHING','K',r=>`=IF(A${r}="","",IF(R${r}<>"","Excluded duplicate",IF(OR(N${r}="Accept suggested match",N${r}="Select different master"),IF(AND(AM${r}="",AN${r}="",'CLEANING'!X${r}="",'CLEANING'!AF${r}="",AQ${r}<>"Pending duplicate review"),"Accepted","Needs review"),IF(N${r}="Reject suggestion","Rejected",IF(E${r}="","Unmatched","Needs review")))))`);
fc('PRODUCT MATCHING','T',r=>`=IF(A${r}="","",IF(K${r}="Accepted","",IF(K${r}="Excluded duplicate",R${r},IF(K${r}="Rejected","Rejected suggestion",AM${r}&IF(AN${r}<>"",AN${r},J${r})&'CLEANING'!X${r}&'CLEANING'!AF${r}&IF(AQ${r}="Pending duplicate review","Duplicate review; ","")&IF(E${r}="",IF(AA${r}="Current","No candidate; ","Refresh fuzzy snapshot; "),"Review pending; ")))))`);
fc('PRODUCT MATCHING','M',r=>`=IF(A${r}="","",IF(G${r}="Exact MPN","Nonblank brand-scoped manufacturer part number.",IF(G${r}="Brand + model + variant","Canonical brand, model and full variant agree.",IF(G${r}="Brand + model","Model candidates; variants need review.",IF(G${r}="Attributes","Exact normalized name and full product attributes.",IF(G${r}="Fuzzy candidate","Snapshot name similarity; not a probability. Human review required.","No qualifying candidate; research or refresh.")))))&IF(AO${r}<>""," "&AO${r},""))`);
fc('PRODUCT MATCHING','U',r=>`=IF(K${r}="Accepted",1,0)`);
fc('PRODUCT MATCHING','AF',r=>`=IF(A${r}="","",IF(K${r}="Accepted",10-INDEX(${rg('MASTER CLEAN','AB',ME)},AD${r}),10-'CLEANING'!W${r}))`);
fc('PRODUCT MATCHING','AG',r=>`=IF(AD${r}="","",IF(COUNTIFS('RESEARCH'!$O$7:$O$106,L${r}&"|Specification")=1,"Verified specification",IF(${mi(r,'R','AD')}="Synthetic","Synthetic / not researched","Pending specification")))`);
fc('PRODUCT MATCHING','AH',r=>`=IF(U${r}=1,SUM($U$7:U${r}),"")`);
fc('PRODUCT MATCHING','AI',r=>`=IF(AND(A${r}<>"",K${r}<>"Accepted",K${r}<>"Excluded duplicate"),SUMPRODUCT(--($A$7:A${r}<>""))-SUM($U$7:U${r})-COUNTIFS($K$7:K${r},"Excluded duplicate"),"")`);
fc('PRODUCT MATCHING','AJ',r=>`=IF(A${r}="",0,IF(AND(OR('DASHBOARD'!$B$5="All",D${r}='DASHBOARD'!$B$5),OR('DASHBOARD'!$D$5="All",C${r}='DASHBOARD'!$D$5),OR('DASHBOARD'!$F$5="All",K${r}='DASHBOARD'!$F$5)),1,0))`);
fc('PRODUCT MATCHING','AK',r=>`=IF(AND(AJ${r}=1,OR('DASHBOARD'!$H$5="",ISNUMBER(SEARCH('DASHBOARD'!$H$5,A${r}&" "&B${r}&" "&C${r}&" "&E${r}&" "&K${r}&" "&T${r}&" "&'CLEANING'!H${r}&" "&'CLEANING'!I${r}&" "&'CLEANING'!B${r}&" "&L${r})))),1,0)`);
fc('PRODUCT MATCHING','AL',r=>`=IF(AK${r}=1,SUM($AK$7:AK${r}),"")`);
for(const [c,w] of [['B',43],['F',43],['G',28],['J',30],['K',22],['M',70],['Q',65],['T',70]])sh['PRODUCT MATCHING'].getRange(`${c}6:${c}356`).format.columnWidth=w;
sh['PRODUCT MATCHING'].getRange('H7:H356').setNumberFormat('0.0%');sh['PRODUCT MATCHING'].getRange('P7:P356').setNumberFormat('yyyy-mm-dd');
cf('PRODUCT MATCHING','J7:K356','AND($A7<>"",$J7<>"")');cf('PRODUCT MATCHING','K7:L356','$K7="Accepted"',C.pale);cf('PRODUCT MATCHING','K7:L356','$K7="Needs review"',C.amber);
console.log('Stage: Review views');
// Review views read the live matching state, not positional manual cells.
for(const [c,src] of [['G','L'],['H','K'],['I','E'],['J','AM']])fc('REVIEW DECISIONS',c,r=>`=IF(A${r}="","",${safeLookup(`A${r}`,'PRODUCT MATCHING',src)})`);
for(const [c,src] of [['G','T'],['H','V']])fc('DUPLICATE REVIEW',c,r=>`=IF(A${r}="","",${safeLookup(`A${r}`,'CLEANING',src)})`);
cf('REVIEW DECISIONS','H7:J356','AND($A7<>"",$H7="Needs review")',C.amber);

console.log('Stage: Export-ready');
// Export-ready compact mapping output; source row ID is the unique row grain.
const fh=['Supplier row ID','Supplier SKU','Master product ID','Internal SKU','Product name','Brand','Category','Model','MPN','GTIN (text)','Color','Capacity','Size / version','Pack quantity','Supplier unit price','Currency','Approved specification','Data disclosure','Supplier provenance','Master provenance','Review status','Verification status','Research / master source URL','Decision notes'];
title('FINAL CLEAN CATALOG','Final clean catalog • accepted mappings','Grain: one accepted, nonduplicate supplier row per mapping. Unique key: Supplier row ID. Multiple supplier SKUs may map to one master.');table('FINAL CLEAN CATALOG',fh,[],SC);
const fpos=r=>`MATCH(ROW()-6,${rg('PRODUCT MATCHING','AH')},0)`;
const pm=(r,c)=>`INDEX(${rg('PRODUCT MATCHING',c)},${fpos(r)})`;
const finalguard=r=>`COUNTIFS(${rg('PRODUCT MATCHING','U')},1)<ROW()-6`;
for(const [c,expr] of [['A',r=>pm(r,'A')],['B',r=>`INDEX(${rg('CLEANING','B')},${fpos(r)})`],['C',r=>pm(r,'L')],['O',r=>`INDEX(${rg('CLEANING','O')},${fpos(r)})`],['P',r=>`INDEX(${rg('CLEANING','P')},${fpos(r)})`],['S',r=>`INDEX(${rg('SUPPLIER RAW','R')},${fpos(r)})`],['U',r=>pm(r,'K')],['V',r=>pm(r,'AG')],['X',r=>pm(r,'Q')]])fc('FINAL CLEAN CATALOG',c,r=>`=IF(${finalguard(r)},"",${expr(r)})`);
for(const [c,src] of [['D','C'],['E','E'],['F','F'],['G','G'],['H','H'],['I','I'],['J','J'],['K','K'],['L','L'],['M','M'],['N','N'],['Q','Q'],['R','R'],['T','S'],['W','T']])fc('FINAL CLEAN CATALOG',c,r=>`=IF(${finalguard(r)},"",IF(INDEX(${rg('MASTER CLEAN',src,ME)},${pm(r,'AD')})="","",INDEX(${rg('MASTER CLEAN',src,ME)},${pm(r,'AD')})))`);
sh['FINAL CLEAN CATALOG'].getRange('E6:E356').format.columnWidth=55;sh['FINAL CLEAN CATALOG'].getRange('Q6:Q356').format.columnWidth=55;sh['FINAL CLEAN CATALOG'].getRange('O7:O356').setNumberFormat('#,##0');sh['FINAL CLEAN CATALOG'].getRange('A7:D356').setNumberFormat('@');
title('EXCEPTIONS','Exceptions • unresolved or rejected records','Grain: one nonaccepted supplier row, excluding confirmed duplicates. All exceptions remain in raw and operational views.');
table('EXCEPTIONS',['Supplier row ID','Product name','Brand','Category','Review outcome','Suggested master','Issue / reason','Duplicate state','Reviewer','Decision notes'],[],SC);
for(const [c,src] of [['A','A'],['B','B'],['C','C'],['D','D'],['E','K'],['F','E'],['G','T'],['H','R'],['I','O'],['J','Q']])fc('EXCEPTIONS',c,r=>`=IF(COUNTIFS(${rg('PRODUCT MATCHING','AI')},ROW()-6)=0,"",INDEX(${rg('PRODUCT MATCHING',src)},MATCH(ROW()-6,${rg('PRODUCT MATCHING','AI')},0)))`);
sh.EXCEPTIONS.getRange('B6:B356').format.columnWidth=54;sh.EXCEPTIONS.getRange('G6:G356').format.columnWidth=78;

console.log('Stage: Dashboard:');
// Dashboard: selectors change the same filter flag consumed by all filtered metrics and charts.
const ds=sh.DASHBOARD;ds.freezePanes.unfreeze();
title('DASHBOARD','ProductMatch | reconciliation workspace','Ridla A. Gani • fictional electronics & office-accessories retailer • 11 Nov 2026');
const widths={A:3,B:21,C:17,D:23,E:16,F:23,G:17,H:31,I:16,J:3};for(const[c,w]of Object.entries(widths))ds.getRange(`${c}1:${c}57`).format.columnWidth=w;
block('DASHBOARD','B4',[['Category','','Brand','','Review outcome','','Search products / issues']]);cell('DASHBOARD','B5','All');cell('DASHBOARD','D5','All');cell('DASHBOARD','F5','All');cell('DASHBOARD','H5','');
listref('DASHBOARD','B5',"'SETTINGS'!$J$7:$J$16");listref('DASHBOARD','D5',`'SETTINGS'!$K$7:$K$${7+brands.length}`);listref('DASHBOARD','F5',"'SETTINGS'!$L$7:$L$12");for(const c of ['B','D','F','H'])input('DASHBOARD',`${c}5`);ds.getRange('B5:I5').format.rowHeight=34;
function card(labelRange,valueRange,label,f,percent=false){ds.getRange(labelRange).merge();ds.getRange(valueRange).merge();ds.getRange(labelRange).values=[[label]];ds.getRange(valueRange).formulas=[[f]];ds.getRange(labelRange).format={fill:C.gray,font:{...font,bold:true,color:C.muted},rowHeight:25};ds.getRange(valueRange).format={fill:C.gray,font:{...font,size:22,bold:true,color:C.navy},rowHeight:28,horizontalAlignment:'center'};if(percent)ds.getRange(valueRange).setNumberFormat('0.0%');}
cell('DASHBOARD','B7','OVERALL TOTALS • unaffected by selectors');ds.getRange('B7:I7').merge();
card('B8:C8','B9:C10','Master rows',`=COUNTIFS(${rg('MASTER RAW','A',ME)},"<>")`);
card('D8:E8','D9:E10','Supplier rows',`=COUNTIFS(${rg('SUPPLIER RAW','A')},"<>")`);
card('F8:G8','F9:G10','Real product models',`=COUNTIFS(${rg('MASTER RAW','Q',ME)},"Real model / simulated offer")`);
card('H8:I8','H9:I10','Research fields verified / total',`=COUNTIFS('RESEARCH'!$G$7:$G$106,"Verified")&" / "&COUNTIFS('RESEARCH'!$A$7:$A$106,"<>")`);
cell('DASHBOARD','B12','FILTERED WORKLOAD • category + brand + review outcome');ds.getRange('B12:I12').merge();
const filterCount=status=>`COUNTIFS(${rg('PRODUCT MATCHING','AJ')},1,${rg('PRODUCT MATCHING','K')},"${status}")`;
card('B13:C13','B14:C15','Accepted mappings',`=${filterCount('Accepted')}`);
card('D13:E13','D14:E15','Awaiting review',`=${filterCount('Needs review')}`);
card('F13:G13','F14:G15','Unmatched products',`=${filterCount('Unmatched')}`);
card('H13:I13','H14:I15','Confirmed duplicate rows',`=${filterCount('Excluded duplicate')}`);
card('B17:C17','B18:C19','Suspected rows still open',`=SUMPRODUCT((${rg('PRODUCT MATCHING','AJ')}=1)*(${rg('PRODUCT MATCHING','AQ')}="Pending duplicate review")*(${rg('PRODUCT MATCHING','R')}=""))`);
const numerator=c=>`SUMPRODUCT(${rg('PRODUCT MATCHING',c)},${rg('PRODUCT MATCHING','AJ')})`;
card('D17:E17','D18:E19','Required completeness • before',`=IF(SUM(${rg('PRODUCT MATCHING','AJ')})=0,"n.a.",${numerator('AE')}/(10*SUM(${rg('PRODUCT MATCHING','AJ')})))`,true);
card('F17:G17','F18:G19','Required completeness • after',`=IF(SUM(${rg('PRODUCT MATCHING','AJ')})=0,"n.a.",${numerator('AF')}/(10*SUM(${rg('PRODUCT MATCHING','AJ')})))`,true);
card('H17:I17','H18:I19','Remaining exception rows',`=COUNTIFS(${rg('PRODUCT MATCHING','AJ')},1,${rg('PRODUCT MATCHING','K')},"<>Accepted",${rg('PRODUCT MATCHING','K')},"<>Excluded duplicate")`);
block('DASHBOARD','L6',[['Outcome','Filtered rows'],...outcomes.map(x=>[x,null])]);for(let i=0;i<5;i++)formula('DASHBOARD',`M${i+7}`,`=${filterCount(outcomes[i])}`);
block('DASHBOARD','O6',[['Category','Open exception rows'],...data.categories.map(x=>[x,null])]);for(let i=0;i<data.categories.length;i++)formula('DASHBOARD',`P${i+7}`,`=COUNTIFS(${rg('PRODUCT MATCHING','AJ')},1,${rg('PRODUCT MATCHING','D')},O${i+7},${rg('PRODUCT MATCHING','K')},"<>Accepted",${rg('PRODUCT MATCHING','K')},"<>Excluded duplicate")`);
const chart=ds.charts.add('bar',ds.getRange('L6:M11'));chart.title='Matching outcomes • filtered rows';chart.hasLegend=false;chart.setPosition('B21','E32');chart.titleTextStyle.typeface='Arial';chart.titleTextStyle.fontSize=13;chart.series.items[0].fill=C.teal;chart.xAxis={axisType:'textAxis',textStyle:{typeface:'Arial',fontSize:11}};chart.yAxis={numberFormatCode:'0',numberFormatSourceLinked:false,textStyle:{typeface:'Arial',fontSize:11}};
const chart2=ds.charts.add('bar',ds.getRange('O6:P15'));chart2.title='Open exceptions by category • filtered';chart2.hasLegend=false;chart2.setPosition('F21','I32');chart2.titleTextStyle.typeface='Arial';chart2.titleTextStyle.fontSize=13;chart2.series.items[0].fill='#54708D';chart2.xAxis={axisType:'textAxis',textStyle:{typeface:'Arial',fontSize:10}};chart2.yAxis={numberFormatCode:'0',numberFormatSourceLinked:false,textStyle:{typeface:'Arial',fontSize:11}};
cell('DASHBOARD','B33','Search results • filters + search text • first 12 rows');formula('DASHBOARD','H33',`=SUM(${rg('PRODUCT MATCHING','AK')})&" total results"`);
block('DASHBOARD','B34',[['Source row','Product','','Brand','Outcome','Suggested ID','Issue / reason','']]);ds.getRange('B34:I34').format={fill:C.navy,font:{...font,color:C.white,bold:true},rowHeight:30};
for(let r=35;r<=46;r++){
 ds.getRange(`C${r}:D${r}`).merge();ds.getRange(`H${r}:I${r}`).merge();
 for(const [dc,pc] of [['B','A'],['C','B'],['E','C'],['F','K'],['G','E'],['H','T']])formula('DASHBOARD',`${dc}${r}`,`=IF(SUM(${rg('PRODUCT MATCHING','AK')})<ROW()-34,"",INDEX(${rg('PRODUCT MATCHING',pc)},MATCH(ROW()-34,${rg('PRODUCT MATCHING','AL')},0)))`);
 ds.getRange(`B${r}:I${r}`).format.fill=r%2?C.white:C.gray;ds.getRange(`B${r}:I${r}`).format.rowHeight=42;ds.getRange(`B${r}:I${r}`).format.wrapText=true;
}
cell('DASHBOARD','B48','Completeness = populated required fields ÷ (10 × filtered supplier rows). After uses accepted master enrichment; N/A is explicit.');
cell('DASHBOARD','B49','Search narrows the result list only. Cards and charts follow the three selectors. “All” includes confirmed duplicate rows.');
cell('DASHBOARD','B50','Research completion counts field-level statuses, not entire-product verification. Export requires a valid reviewer decision.');
formula('DASHBOARD','B52',`="Overall: "&COUNTIFS('MASTER CLEAN'!$W$7:$W$306,"Exact master duplicate")&" master duplicates excluded; "&COUNTIFS('RESEARCH'!$G$7:$G$106,"<>Verified",'RESEARCH'!$A$7:$A$106,"<>")&" research fields unresolved."`);
formula('DASHBOARD','B53',`="Fuzzy snapshot: "&IF(COUNTIFS('PRODUCT MATCHING'!$AA$7:$AA$356,"Refresh required")=0,"current","REFRESH REQUIRED")&". Scores are candidate similarity, not probabilities."`);
ds.getRange('B48:I53').format.font={...font,size:10,color:C.muted};

console.log('Stage: Independent reconciliation');
// Independent reconciliation and deliberately constructed ground truth; no business formula reads QA.
title('QA EVALUATION','QA & evaluation • simulated scenario','Ground truth is deliberately constructed and never feeds operational suggestions or reviewer decisions. Denominators are explicit.');
block('QA EVALUATION','A6',[['Reconciliation check','Observed','Expected / control','Result']]);
const checks=[
 ['Supplier partition',`SUM(COUNTIFS('PRODUCT MATCHING'!$K$7:$K$356,"Accepted"),COUNTIFS('PRODUCT MATCHING'!$K$7:$K$356,"Needs review"),COUNTIFS('PRODUCT MATCHING'!$K$7:$K$356,"Unmatched"),COUNTIFS('PRODUCT MATCHING'!$K$7:$K$356,"Rejected"),COUNTIFS('PRODUCT MATCHING'!$K$7:$K$356,"Excluded duplicate"))`,`COUNTIFS('SUPPLIER RAW'!$A$7:$A$356,"<>")`],
 ['Accepted = exported',`SUMPRODUCT(--('FINAL CLEAN CATALOG'!$A$7:$A$356<>""))`,`COUNTIFS('PRODUCT MATCHING'!$K$7:$K$356,"Accepted")`],
 ['Exceptions reconcile',`SUMPRODUCT(--('EXCEPTIONS'!$A$7:$A$356<>""))`,`COUNTIFS('SUPPLIER RAW'!$A$7:$A$356,"<>")-COUNTIFS('PRODUCT MATCHING'!$K$7:$K$356,"Accepted")-COUNTIFS('PRODUCT MATCHING'!$K$7:$K$356,"Excluded duplicate")`],
 ['Master partition',`SUMPRODUCT(--('MASTER CLEAN'!$A$7:$A$306<>""))-COUNTIFS('MASTER CLEAN'!$W$7:$W$306,"Exact master duplicate")`,`COUNTIFS('MASTER RAW'!$A$7:$A$306,"<>")-COUNTIFS('MASTER CLEAN'!$W$7:$W$306,"Exact master duplicate")`],
 ['Ground truth rows',`COUNTIFS($F$7:$F$356,"<>")`,`COUNTIFS('SUPPLIER RAW'!$A$7:$A$356,"<>")`],
 ['Expected duplicate exclusions',`COUNTIFS('PRODUCT MATCHING'!$K$7:$K$356,"Excluded duplicate")`,`COUNTIFS($I$7:$I$356,"<>")`],
 ['Unique exported source IDs',`SUM($M$7:$M$356)`,`SUMPRODUCT(--('FINAL CLEAN CATALOG'!$A$7:$A$356<>""))`],
 ['Research verified fields',`COUNTIFS('RESEARCH'!$G$7:$G$106,"Verified")`,`15`]
];
for(let i=0;i<checks.length;i++){const r=i+7;cell('QA EVALUATION',`A${r}`,checks[i][0]);formula('QA EVALUATION',`B${r}`,'='+checks[i][1]);formula('QA EVALUATION',`C${r}`,'='+checks[i][2]);formula('QA EVALUATION',`D${r}`,`=IF(B${r}=C${r},"PASS","CHECK")`);}
block('QA EVALUATION','A18',[
 ['Evaluation metric','Value','Denominator / definition'],['Suggested mapping precision',null,'Correct suggested IDs / nonblank suggestions on nonduplicate labeled rows, including NONE negatives'],['Positive suggestion recall',null,'Correct suggested IDs / labeled nonduplicate rows with a valid counterpart'],['Accepted mapping precision',null,'Correct accepted IDs / all accepted nonduplicate labeled rows'],['Accepted positive coverage',null,'Correct accepted IDs / labeled nonduplicate rows with a valid counterpart'],['No-counterpart cases',null,'Nonduplicate ground-truth rows with Expected ID = NONE'],['All evaluation outcomes','SIMULATED','Constructed examples; no business-impact or production accuracy claim'],['Required field definition','10 per supplier row','Name, brand, category, model, color, capacity, size/version, pack, price, currency'],['Engine testing','See docs/verification.md','Headless LibreOffice + programmatic checks; Excel 365 and Google Sheets not available for native verification'],['Known limits','No automatic history','Manual tables retain current state; fuzzy is a refreshable snapshot; fixed reserved capacity']]);
const gtHeaders=['Supplier row ID','Expected master ID','Constructed issue','Duplicate of','Suggested ID','Final accepted ID','Evaluation: correct suggestion','Unique export key'];
block('QA EVALUATION','F6',[gtHeaders]);block('QA EVALUATION','F7',data.truth.map(t=>[t[0],t[1],t[2],t[3]]));
fc('QA EVALUATION','J',r=>`=IF(F${r}="","",${safeLookup(`F${r}`,'PRODUCT MATCHING','E')})`);
fc('QA EVALUATION','K',r=>`=IF(F${r}="","",IF(${safeLookup(`F${r}`,'PRODUCT MATCHING','K')}="Accepted",${safeLookup(`F${r}`,'PRODUCT MATCHING','L')},""))`);
fc('QA EVALUATION','L',r=>`=IF(F${r}="","",IF(AND(I${r}="",G${r}<>"NONE",J${r}=G${r}),1,0))`);
fc('QA EVALUATION','M',r=>`=IF('FINAL CLEAN CATALOG'!A${r}="",0,IF(COUNTIFS('FINAL CLEAN CATALOG'!$A$7:$A$356,'FINAL CLEAN CATALOG'!A${r})=1,1,0))`);
formula('QA EVALUATION','B19',`=IF(SUMPRODUCT(($F$7:$F$356<>"")*($I$7:$I$356="")*($J$7:$J$356<>""))=0,"n.a.",SUM($L$7:$L$356)/SUMPRODUCT(($F$7:$F$356<>"")*($I$7:$I$356="")*($J$7:$J$356<>"")))`);
formula('QA EVALUATION','B20',`=SUM($L$7:$L$356)/SUMPRODUCT(($F$7:$F$356<>"")*($I$7:$I$356="")*($G$7:$G$356<>"NONE"))`);
formula('QA EVALUATION','B21',`=IF(SUMPRODUCT(($K$7:$K$356<>"")*($I$7:$I$356=""))=0,"n.a.",SUMPRODUCT(($K$7:$K$356=$G$7:$G$356)*($K$7:$K$356<>"")*($I$7:$I$356=""))/SUMPRODUCT(($K$7:$K$356<>"")*($I$7:$I$356="")))`);
formula('QA EVALUATION','B22',`=SUMPRODUCT(($K$7:$K$356=$G$7:$G$356)*($K$7:$K$356<>"")*($I$7:$I$356=""))/SUMPRODUCT(($F$7:$F$356<>"")*($I$7:$I$356="")*($G$7:$G$356<>"NONE"))`);
formula('QA EVALUATION','B23',`=SUMPRODUCT(($G$7:$G$356="NONE")*($I$7:$I$356=""))`);sh['QA EVALUATION'].getRange('B19:B22').setNumberFormat('0.0%');sh['QA EVALUATION'].getRange('A6:A29').format.columnWidth=36;sh['QA EVALUATION'].getRange('C6:C29').format.columnWidth=90;sh['QA EVALUATION'].getRange('A18:C29').format.wrapText=true;sh['QA EVALUATION'].getRange('A18:C29').format.rowHeight=50;sh['QA EVALUATION'].getRange('H6:H356').format.columnWidth=70;cf('QA EVALUATION','D7:D14','$D7="CHECK"');

console.log('Stage: Documentation');
// Documentation is part of the workbook rather than external promises.
title('DICTIONARY','Data dictionary & formula notes','Definitions, input ownership and refresh boundaries. Formula helpers are visible for inspection.');
const dict=[
 ['Field / group','Definition','Ownership / behavior','Example / limitation'],
 ['Source row ID','Immutable record key within a feed','Input; primary review lookup key','S0001; never reuse an ID for another product'],
 ['Product ID','Internal canonical product record','Master input; validated manual target','M0001; duplicate master records are excluded'],
 ['Supplier / internal SKU','Different business keys in each system','Text; never used as a cross-system identifier','Leading zeros retained; raw SKU values preserved'],
 ['Model','Exact product model; synthetic models begin DEMO-','UPPER + TRIM + CLEAN + NBSP removal','No punctuation stripped; missing model cannot match blank'],
 ['MPN','Genuine manufacturer part number when researched','Brand-scoped exact key; punctuation preserved','Synthetic products have no fabricated MPN'],
 ['GTIN','Barcode identifier as text','Blank where unavailable; placeholders cleared; no GTIN matching','No invented numeric GTINs; independent GS1 verification required for future GTIN use'],
 ['Variant attributes','Color, capacity, size/hardware version, pack','All retained in identity and conflict checks','N/A explicitly means not applicable, not unknown'],
 ['Price / currency','Simulated supplier offer in IDR','Positive numeric price; allowed currency IDR','Not live retail pricing or manufacturer research'],
 ['Required completeness','10 required populated fields per supplier row','Before raw vs after accepted-master enrichment','Measures presence; invalid values remain separate issues'],
 ['Matching method','Strongest available tier','Exact MPN → brand/model/variant → brand/model → attributes → fuzzy','No automatic acceptance, including exact matches'],
 ['Similarity score','0.6 sorted-token SequenceMatcher + 0.4 token Jaccard','Python-generated snapshot, top 3 per brand/category','0–1 similarity; not a statistical probability'],
 ['Candidate count','Candidates in the selected native tier or qualifying top 3 fuzzy candidates','Live threshold filters the snapshot','Full model/variant tier can resolve a broader family to one variant'],
 ['Suggestion conflicts','Differences between supplier and suggested master','Live formula; textual labels and conditional format','Conflicting model/MPN/GTIN/brand/category/variant trigger review'],
 ['Selected conflicts','Differences against the reviewer-selected master','Live formula; blocks export even after Accept','Correct source-derived data via CORRECTIONS or choose compatible target'],
 ['Current outcome','Accepted / Needs review / Unmatched / Rejected / Excluded duplicate','Derived from decisions, data validity and duplicates','Accepted requires reviewer, date, valid active master, resolved variants'],
 ['Reject suggestion','Removes a record from export','Reviewer input retained by source ID','Record stays in EXCEPTIONS; it is not deleted'],
 ['Duplicate decision','Keep separate / Confirm duplicate / Pending','Reviewer needs reason, name, date; confirm also needs an earlier source ID','Confirm only accepts identical full identity and earlier lexical source ID'],
 ['Exclusion','Automatic retransmission or confirmed same-variant duplicate','Raw retained; operational row retained','Master survivor earliest raw row; supplier exact survivor earliest raw row'],
 ['Research statuses','Field-level Verified / Not found / Conflicting sources / Pending','Manual evidence inputs; access dates are actual research date','Verified means the specific proposed attribute is supported'],
 ['Research apply key','Product ID | Specification when eligible','Exact raw model + Verified + Yes + full evidence; unique key required','Changing evidence / approval updates final output automatically'],
 ['Original / proposed values','Source value and evidence-supported replacement','Current research table; only Specification applied in this version','Other research attributes require developer extension for automatic application'],
 ['Final output grain','One accepted, nonduplicate supplier source row per mapping','Compact live INDEX/MATCH output; unique Source row ID','Not a deduplicated one-row-per-master stock list'],
 ['Exception grain','One pending, rejected or unmatched supplier row','Excludes confirmed duplicates; all raw records remain','Accepting a valid mapping removes it from exceptions'],
 ['Dashboard controls','Category, brand, outcome, search','Three selectors filter cards/charts; search filters result list only','All includes duplicate rows; overall cards remain unchanged'],
 ['Refresh / calculation','Native formulas calculate automatically','Fuzzy script patches CANDIDATES only; then calculate/save Excel','CSV files do not update until exported again'],
 ['Manual persistence','Inputs keyed by stable source IDs','Approved ID is an input value; refresh preserves all four manual tables','Current state, not an automatic historical change log'],
 ['Supported capacity','300 master / 350 supplier / 100 research records','Preformatted tables and bounded formulas','Paste into reserved blanks; capacity extension needs range updates'],
 ['Platform testing','Excel 365 intended; no VBA or macros','Artifact authoring + headless LibreOffice verification','No Excel 365 / native Google Sheets test claimed'],
 ['Source provenance','DEMO-* = generated scenario; OFFICIAL-* = real-model context','15 official sources researched; trading data remains synthetic','Indexed manufacturer evidence used for SanDisk / Crucial; access caveats in evidence notes'],
 ['Evaluation','Constructed ground truth on QA EVALUATION','Never used by suggestions or acceptance logic','No production accuracy, business impact or time saving claim']
];table('DICTIONARY',dict[0],dict.slice(1),dict.length-1);for(const[c,w]of[['A',29],['B',65],['C',72],['D',79]])sh.DICTIONARY.getRange(`${c}6:${c}100`).format.columnWidth=w;sh.DICTIONARY.getRange('A7:D37').format.wrapText=true;sh.DICTIONARY.getRange('A7:D37').format.rowHeight=58;
title('START HERE','ProductMatch','Product Reconciliation & Data Quality Workspace | Author: Ridla A. Gani');sh['START HERE'].freezePanes.unfreeze();
const start=[
 ['Business problem','Reconcile a fictional supplier feed to an internal electronics and office-accessories catalog; deliver accepted mappings and retain unresolved exceptions.'],
 ['Scope','200 master rows, 250 supplier rows, 15 identifiable real product models. All prices, supplier offers, SKUs and trading outcomes are simulated.'],
 ['Provenance','Deterministic synthetic dataset generated by scripts/generate_demo_data.py; real model facts researched from official product pages and specification sheets on 11 Nov 2026.'],
 ['Disclosure','DEMO- model names and brands are fictional. Manufacturer part numbers appear only where sourced. No GTIN has been invented. Raw evidence values remain intact.'],
 ['1. Inspect inputs','Review MASTER RAW and SUPPLIER RAW. Source row IDs link all decisions; SKU systems are independent.'],
 ['2. Clean and resolve duplicates','Filter CLEANING Issue queue. Use CORRECTIONS for overrides and DUPLICATE REVIEW for suspected duplicate decisions. Raw records stay unchanged.'],
 ['3. Review suggestions','Use PRODUCT MATCHING: exact MPN, brand/model, normalized attributes, then Python fuzzy snapshot. Conflicts are visible; fuzzy is never auto-approved.'],
 ['4. Record a decision','In REVIEW DECISIONS choose Accept, Select different master, Reject or Leave pending. Confirm the approved ID in column C and supply reviewer/date; IDs are validated.'],
 ['5. Verify missing details','Review RESEARCH field-level sources, exact model, evidence and date. Only uniquely eligible Specification approvals apply automatically.'],
 ['6. Deliver and monitor','Review FINAL CLEAN CATALOG, EXCEPTIONS and DASHBOARD. Recalculate/save before exporting CSVs with the supplied script.'],
 ['Navigation','Use the links at right; every working sheet has a return link. Table headers have filters; relevant headers and record IDs are frozen.'],
 ['Inputs (amber / blue text)','Editable selector, review, correction, duplicate or research inputs. Seeded approvals and dates outside research are simulated current-state examples.'],
 ['Formulas (dark text)','Calculated cleaning, matching, outcomes, outputs, QA and dashboard. Do not paste over calculated columns.'],
 ['Warnings (amber / red)','Textual status plus conditional formatting. Important conflicts or invalid values block export until resolved.'],
 ['Approved outputs (teal cues)','Only Accepted rows enter the final mapping table. A green research status verifies the field, not every property of that product.'],
 ['Five-minute demo: 0–1 min','Show overall 200/250 inputs, choose Storage, then All. Search M185. Cards/charts follow selectors; search changes the list only.'],
 ['Five-minute demo: 1–2 min','Open CLEANING; inspect original vs standardized text, brand alias and duplicate flags. Show exact retransmissions without deleting raw rows.'],
 ['Five-minute demo: 2–3 min','In REVIEW DECISIONS accept a pending compatible record (e.g. S0170), confirm C=M0170, add reviewer/date, and observe final output. Reject it to remove it from export.'],
 ['Five-minute demo: 3–4 min','Show S0181 overridden to its correct white variant. Contrast S0191 with unavailable Red/pack-2 variant that cannot be exported.'],
 ['Five-minute demo: 4–5 min','Show research R001 and official source. Toggle Apply approved No/Yes and observe the specification disappear/reappear in final output. End on QA.'],
 ['Requirements','Microsoft Excel 365 desktop, macro-free .xlsx, calculation mode Automatic. No VBA, add-ins or Python required for normal review/dashboard use.'],
 ['Refresh instructions','After saving/closing, run python scripts/refresh_productmatch.py --workbook ProductMatch.xlsx --in-place. Then open, calculate and save. Native formulas recalculate automatically.'],
 ['CSV refresh','Run scripts/refresh_productmatch.py --workbook ProductMatch.xlsx --export-cached exports AFTER calculation and saving. Supplied CSVs reflect the delivered baseline only.'],
 ['Capacity / append','300 master, 350 supplier, 100 research rows. Paste into blank reserved raw rows (row 7 onward). Enter new source IDs in editable review tables; do not overwrite formulas. Refresh fuzzy for new records.'],
 ['Manual data safety','Sort complete tables, not individual columns. Source IDs must stay unique. Manual state lives in keyed input tables and survives snapshot refresh; no automatic history log is claimed.'],
 ['Research caveats','SanDisk and Crucial use official indexed specification content; direct page timeout / mirror redirect is recorded. Power/form-factor contradictions remain unresolved. Prices are simulated.'],
 ['Google Sheets','See docs/google_sheets.md and optional scripts/ProductMatch_Sheets.gs. Conversion and Sheets behavior are untested.'],
 ['Test boundary','See docs/verification.md for executed programmatic and headless LibreOffice checks. No native Excel 365 or Google Sheets verification is claimed.'],
 ['Limitations','Fixed reserved capacity; only Specification research applies automatically; no live web scraping; scores require refresh; CSVs are snapshots. All evaluation metrics describe deliberately constructed data.']
];
block('START HERE','A6',[['Workspace guide','Instructions']]);block('START HERE','A7',start);sh['START HERE'].getRange('A6:A40').format.columnWidth=39;sh['START HERE'].getRange('B6:B40').format.columnWidth=112;sh['START HERE'].getRange('A7:B40').format.wrapText=true;sh['START HERE'].getRange('A7:B40').format.rowHeight=58;sh['START HERE'].getRange('D6:D30').format.columnWidth=31;
cell('START HERE','D6','OPEN WORKSPACE');for(let i=1;i<names.length;i++)formula('START HERE',`D${7+i}`,`=HYPERLINK("#'${names[i]}'!A1","${names[i]}")`);sh['START HERE'].getRange('D7:D25').format.font={...font,color:C.teal,bold:true};
for(const n of ['START HERE','DICTIONARY'])sh[n].tabColor=C.teal;
// Make presentation and review headers readable; working grids stay unmerged.
for(const n of names){
 const s=sh[n];s.getRange('A1').format.rowHeight=32;s.getRange('A2').format.rowHeight=29;
 if(n!=='DASHBOARD'&&n!=='START HERE')s.getRange('A4').values=[['Macro-free • stable source IDs • formulas update automatically • fuzzy scores require external refresh']];
}
console.log('Authoring complete; recalculating');
wb.recalculate();
console.log((await wb.inspect({kind:'region',sheetId:'DASHBOARD',range:'B8:I19',maxChars:3500,tableMaxRows:12,tableMaxCols:8})).ndjson);
console.log((await wb.inspect({kind:'match',searchTerm:'#REF!|#DIV/0!|#VALUE!|#NAME\\?|#N/A|#NUM!|#NULL!|#SPILL!|#CALC!',options:{useRegex:true,maxResults:15},maxChars:2500})).ndjson);
const output=await SpreadsheetFile.exportXlsx(wb);await output.save(path.join(outputDir,'ProductMatch.xlsx'));console.log('XLSX exported');
// Render every sheet once for visual QA; three designated ranges are portfolio screenshots.
const ranges={'START HERE':'A1:D14','DASHBOARD':'A1:J33','PRODUCT MATCHING':'A1:M14','RESEARCH':'A1:L11','FINAL CLEAN CATALOG':'A1:H14','EXCEPTIONS':'A1:G14','REVIEW DECISIONS':'A1:J14','CLEANING':'A1:G14','CORRECTIONS':'A1:G12','DUPLICATE REVIEW':'A1:H13','MASTER CLEAN':'A1:H14','MASTER RAW':'A1:H14','SUPPLIER RAW':'A1:H14','SETTINGS':'A1:H14','DICTIONARY':'A1:D13','QA EVALUATION':'A1:D15','CANDIDATES':'A1:I12'};
await fs.mkdir(path.join(outputDir,'screenshots'),{recursive:true});await fs.mkdir(path.join(outputDir,'qa_previews'),{recursive:true});
for(const n of names){try{const image=await wb.render({sheetName:n,range:ranges[n],scale:1.25,format:'png'});const filename=n==='DASHBOARD'?'screenshots/01_dashboard.png':n==='PRODUCT MATCHING'?'screenshots/02_matching_review.png':n==='RESEARCH'?'screenshots/03_research_evidence.png':`qa_previews/${n.replaceAll(' ','_').toLowerCase()}.png`;await fs.writeFile(path.join(outputDir,filename),new Uint8Array(await image.arrayBuffer()));console.log('Rendered '+n);}catch(e){console.log('Render failed '+n+': '+e.message);}}
await fs.writeFile(path.join(outputDir,'build_inspection.json'),JSON.stringify({dashboard:ds.getRange('B8:I19').values,qa:sh['QA EVALUATION'].getRange('A7:D14').values},null,2));
