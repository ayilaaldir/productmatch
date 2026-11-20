#!/usr/bin/env python3
"""Disposable LibreOffice mutation tests, XML feature checks and baseline metrics."""
import argparse, collections, datetime, hashlib, json, pathlib, shutil, subprocess, sys, tempfile, zipfile
from refresh_productmatch import XLSX, NS, ET, colname, export_cached, normalize
ROOT=pathlib.Path(__file__).resolve().parents[1]
SOFFICE='soffice'
def errors(x):
 out=[]
 for sheet in x.sheets:
  root=ET.fromstring(x.files[x.sheets[sheet]])
  for c in root.findall('.//s:sheetData/s:row/s:c',NS):
   if c.attrib.get('t')=='e':out.append((sheet,c.attrib['r'],c.find('s:v',NS).text))
 return out
def recalc(src,dest):
 dest.mkdir(parents=True,exist_ok=True)
 # Discard authoring-engine caches so every formula is evaluated by LibreOffice.
 # FullCalc flags alone can leave imported cached dependents unchanged.
 stage=dest.parent/(dest.name+'_input');stage.mkdir(exist_ok=True);cleaned=stage/src.name;z=XLSX(src)
 for name,part in z.sheets.items():
  root=ET.fromstring(z.files[part])
  for c in root.findall('.//s:sheetData/s:row/s:c',NS):
   if c.find('s:f',NS) is None:continue
   for tag in ['v','is']:
    child=c.find('s:'+tag,NS)
    if child is not None:c.remove(child)
   c.attrib.pop('t',None)
  z.files[part]=ET.tostring(root,encoding='utf-8',xml_declaration=True)
 z.save(cleaned)
 profile=(dest.parent/('lo_profile_'+dest.name)).resolve().as_uri()
 p=subprocess.run([SOFFICE,'-env:UserInstallation='+profile,'--headless','--convert-to','xlsx','--outdir',str(dest),str(cleaned)],capture_output=True,text=True,timeout=240)
 out=dest/src.name
 if p.returncode or not out.exists():raise RuntimeError(p.stdout+p.stderr)
 return out
def metrics(x):
 pm=x.read('PRODUCT MATCHING');dash=x.read('DASHBOARD');qa=x.read('QA EVALUATION');research=x.read('RESEARCH');final=x.read('FINAL CLEAN CATALOG')
 counts=collections.Counter(pm.get(f'K{r}') for r in range(7,357) if pm.get(f'A{r}'))
 return {'outcomes':dict(counts),'exported':sum(bool(final.get(f'A{r}')) for r in range(7,357)),
  'dashboard':{k:dash.get(k) for k in ['B9','D9','F9','H9','B14','D14','F14','H14','B18','D18','F18','H18']},
  'evaluation':{qa.get('A'+str(r)):qa.get('B'+str(r)) for r in range(19,24)},
  'reconciliation':{qa.get('A'+str(r)):qa.get('D'+str(r)) for r in range(7,15)},
  'research_statuses':dict(collections.Counter(research.get(f'G{r}') for r in range(7,107) if research.get(f'A{r}')))}
def inject_caches(original,calculated,out):
 """Retain authoring-native tables/styles/validation; import evaluated cell caches only."""
 a=XLSX(original);b=XLSX(calculated)
 for n,p in a.sheets.items():
  ar=ET.fromstring(a.files[p]);br=ET.fromstring(b.files[b.sheets[n]])
  bmap={c.attrib['r']:c for c in br.findall('.//s:sheetData/s:row/s:c',NS)}
  for c in ar.findall('.//s:sheetData/s:row/s:c',NS):
   if c.find('s:f',NS) is None:continue
   bc=bmap.get(c.attrib['r'])
   if bc is None:continue
   for tag in ['v','is']:
    v=c.find('s:'+tag,NS)
    if v is not None:c.remove(v)
   typ=bc.attrib.get('t')
   if typ=='s':
    c.set('t','str');v=ET.SubElement(c,'{'+NS['s']+'}v');v.text=b.shared[int(bc.find('s:v',NS).text)]
   elif typ=='inlineStr':
    c.set('t','str');ET.SubElement(c,'{'+NS['s']+'}v').text=''.join(t.text or '' for t in bc.iter('{'+NS['s']+'}t'))
   else:
    if typ:c.set('t',typ)
    else:c.attrib.pop('t',None)
    bv=bc.find('s:v',NS)
    if bv is not None:ET.SubElement(c,'{'+NS['s']+'}v').text=bv.text
  a.files[p]=ET.tostring(ar,encoding='utf-8',xml_declaration=True)
 # Refresh source-linked chart caches from calculated summary cells, preserving original chart XML.
 cns={'c':'http://schemas.openxmlformats.org/drawingml/2006/chart'};cn=cns['c']
 for p in list(a.files):
  if not '/charts/chart' in p or not p.endswith('.xml'):continue
  root=ET.fromstring(a.files[p])
  for ref in root.findall('.//c:numRef',cns):
   f=ref.find('c:f',cns)
   if f is None or not f.text or '!' not in f.text:continue
   import re
   sheet,rng=f.text.rsplit('!',1);sheet=sheet.strip("'");rng=rng.replace('$','')
   m=re.fullmatch('([A-Z]+)([0-9]+):([A-Z]+)([0-9]+)',rng)
   if not m or sheet not in b.sheets:continue
   c1,r1,c2,r2=m.groups();vals=b.read(sheet)
   cache=ref.find('c:numCache',cns)
   if cache is None:cache=ET.SubElement(ref,'{'+cn+'}numCache')
   for child in list(cache):cache.remove(child)
   ET.SubElement(cache,'{'+cn+'}formatCode').text='General';ET.SubElement(cache,'{'+cn+'}ptCount',{'val':str(int(r2)-int(r1)+1)})
   for idx,r in enumerate(range(int(r1),int(r2)+1)):
    pt=ET.SubElement(cache,'{'+cn+'}pt',{'idx':str(idx)});ET.SubElement(pt,'{'+cn+'}v').text=str(vals.get(c1+str(r),0))
  a.files[p]=ET.tostring(root,encoding='utf-8',xml_declaration=True)
 a.save(out)
def main():
 global SOFFICE
 p=argparse.ArgumentParser(description=__doc__);p.add_argument('--workbook',default=str(ROOT/'ProductMatch.xlsx'));p.add_argument('--quick',action='store_true',help='Recalculate baseline but skip mutation scenarios');p.add_argument('--cached-only',action='store_true',help='Check saved results and native XML without claiming recalculation');p.add_argument('--soffice',default=shutil.which('soffice'),help='LibreOffice soffice executable');p.add_argument('--report-dir',type=pathlib.Path,default=ROOT/'.artifacts/verification');p.add_argument('--recalculated-output',type=pathlib.Path,help='Optional NEW workbook with evaluated caches, preserving original native parts');p.add_argument('--resume-directory');a=p.parse_args()
 if not a.cached_only and not a.soffice:p.error('LibreOffice was not found. Use --soffice PATH or --cached-only for saved-result checks.')
 if a.cached_only and a.recalculated_output:p.error('--recalculated-output requires LibreOffice recalculation')
 SOFFICE=a.soffice
 source=pathlib.Path(a.workbook).resolve();tmp=pathlib.Path(a.resume_directory) if a.resume_directory else pathlib.Path(tempfile.mkdtemp(prefix='productmatch-qa-'));results=[]
 digest=hashlib.sha256(source.read_bytes()).hexdigest()
 if a.resume_directory:
  assert (tmp/'source.sha256').read_text().strip()==digest,'Source changed since the saved test run'
 else:(tmp/'source.sha256').write_text(digest)
 if a.recalculated_output and a.recalculated_output.resolve()==source:p.error('--recalculated-output must differ from the input workbook')
 calc=source if a.cached_only else (tmp/'baseline'/source.name if a.resume_directory else recalc(source,tmp/'baseline'));x=XLSX(calc);m=metrics(x)
 assert not errors(x),errors(x)[:10]
 assert all(v=='PASS' for v in m['reconciliation'].values()),m
 engine='Saved caches / OOXML; no recalculation' if a.cached_only else 'Headless LibreOffice full recalculation'
 print('Baseline saved-result checks passed' if a.cached_only else 'Baseline full recalculation and reconciliation passed',flush=True)
 results.append({'test':'Baseline formula error scan / eight reconciliation checks','result':'PASS','engine':engine})
 data=json.loads((ROOT/'data/generated/demo_data.json').read_text(encoding='utf-8'))
 for name,key in [('MASTER RAW','master'),('SUPPLIER RAW','supplier')]:
  rawvalues=x.read(name)
  for r,row in enumerate(data[key],7):
   for c,value in enumerate(row,1):assert rawvalues.get(f'{colname(c)}{r}','')==value,('Raw snapshot changed',name,r,c)
 expected={r[1]:normalize(r,data['aliases']) for r in data['master']};f=x.read('FINAL CLEAN CATALOG')
 for r in range(7,357):
  if not f.get('A'+str(r)):continue
  e=expected[f['C'+str(r)]]
  for c,k in [('E','name'),('F','brand'),('G','category'),('H','model'),('I','mpn'),('K','color'),('L','capacity'),('M','size'),('N','pack')]:
   assert str(f.get(c+str(r),''))==str(e[k]),('Final attribute mismatch',r,c,f.get(c+str(r)),e[k])
 results.append({'test':'Every exported mapping attribute agrees with independently normalized selected master','result':'PASS','engine':'Python independent source comparison'})
 original=XLSX(source)
 assert len(original.sheets)==17
 counts={'tables':len([n for n in original.files if n.startswith('xl/tables/table') and n.endswith('.xml')]),'charts':len([n for n in original.files if '/charts/chart' in n and n.endswith('.xml')])}
 assert counts['charts']==2
 assert not any(n.endswith('vbaProject.bin') for n in original.files)
 rroot=ET.fromstring(original.files[original.sheets['REVIEW DECISIONS']]);assert len(rroot.findall('s:dataValidations/s:dataValidation',NS))>=3
 results.append({'test':'Native tables, 2 linked charts, dropdown XML, macro-free archive','result':'PASS','engine':'Programmatic OOXML inspection','features':counts})
 pm=x.read('PRODUCT MATCHING');sraw=x.read('SUPPLIER RAW');mc=x.read('MASTER CLEAN')
 assert pm['K207']=='Excluded duplicate' and pm['K221']=='Excluded duplicate'
 assert mc['W205']=='Exact master duplicate' and mc['W206']=='Exact master duplicate'
 assert pm['K197']=='Rejected' and pm['J197']
 assert pm['G25']!='Exact MPN' # S0019 has blank MPN / model; no blank-to-blank match.
 results.append({'test':'Blank identifiers, duplicate identifiers, conflicting variant and duplicate exclusions','result':'PASS','engine':engine})
 baseline=m['exported']
 def scenario(name,changes,check):
  print('Testing '+name,flush=True)
  inp=tmp/(name+'.xlsx');z=XLSX(source)
  for sheet,updates in changes.items():z.set_cells(sheet,updates)
  z.save(inp);out=tmp/name/inp.name if a.resume_directory else recalc(inp,tmp/name);z=XLSX(out);assert not errors(z),errors(z)[:8]
  details=check(z);results.append({'test':name,'result':'PASS','engine':'Headless LibreOffice full recalculation','detail':details});print('PASS '+name,flush=True)
 if not a.quick and not a.cached_only:
  def selectors(z):
   d=z.read('DASHBOARD');p=z.read('PRODUCT MATCHING')
   assert d['B14']==1 and d['D14']==0 and d['F14']==0
   assert d['B9']==200 and d['D9']==250
   assert sum(p.get(f'AK{r}',0) for r in range(7,357))==1 and d['B35']=='S0001'
   assert d['M7']==1 and sum(d.get('P'+str(r),0) for r in range(7,16))==0
   return ['Category, brand and review-outcome selectors change filtered metrics and chart source cells','Search finds the one qualifying product','Overall totals remain unchanged']
  scenario('selectors_search_and_chart_sources',{'DASHBOARD':{'B5':'Mice','D5':'Logitech','F5':'Accepted','H5':'M185'}},selectors)
  def accepting(z):
   p=z.read('PRODUCT MATCHING');assert p['K176']=='Accepted' and p['L176']=='M0170';assert metrics(z)['exported']==baseline+1
   return {'accepted_source':'S0170','exported_delta':1,'dashboard_accepted_delta':z.read('DASHBOARD')['B14']-m['dashboard']['B14']}
  scenario('accept_pending_mapping',{'REVIEW DECISIONS':{'B176':'Accept suggested match','C176':'M0170','D176':'QA reviewer','E176':46337,'F176':'Disposable QA decision'}},accepting)
  def rejecting(z):
   p=z.read('PRODUCT MATCHING');assert p['K7']=='Rejected' and p['U7']==0;assert metrics(z)['exported']==baseline-1
   f=z.read('FINAL CLEAN CATALOG');assert 'S0001' not in [f.get('A'+str(r)) for r in range(7,132)]
   return {'rejected_source':'S0001','exported_delta':-1}
  scenario('reject_accepted_mapping',{'REVIEW DECISIONS':{'B7':'Reject suggestion'}},rejecting)
  def blocked(z):
   p=z.read('PRODUCT MATCHING')
   assert p['K176']=='Needs review' and p['U176']==0 and 'Valid master' in p['AM176']
   assert p['K197']=='Needs review' and p['U197']==0 and p['AN197']
   assert p['K178']=='Needs review' and 'Reviewer/date' in p['AM178']
   assert p['K187']=='Accepted' and p['L187']=='M0017' and p['E187']!='M0017'
   return ['Invalid pasted master ID blocked','Accepting a conflicting color/pack remains blocked','Reviewer without date remains blocked','Existing explicit override selects a compatible master different from the suggestion']
  scenario('invalid_ID_conflicting_variant_missing_date_and_override',{'REVIEW DECISIONS':{'B176':'Select different master','C176':'M9999','D176':'QA reviewer','E176':46337,'B197':'Accept suggested match','C197':pm['E197'],'D197':'QA reviewer','E197':46337,'B178':'Accept suggested match','C178':'M0172','D178':'QA reviewer','E178':''}},blocked)
  def researching(z):
   mc=z.read('MASTER CLEAN');f=z.read('FINAL CLEAN CATALOG');assert mc['Q7']=='' and f['Q7']=='';assert mc['Q8']=='' and f['Q8']==''
   return ['Revoking research approval removes the final specification','Duplicate approved field keys block application instead of silently first-matching']
  rd=x.read('RESEARCH');dupr={f'{colname(c)}29':rd.get(f'{colname(c)}8','') for c in range(1,15)};dupr.update({'A29':'R023','H7':'No'})
  scenario('research_updates_and_conflicting_approvals',{'RESEARCH':dupr},researching)
  fuzzyrow=next(r for r in range(7,356) if pm.get(f'G{r}')=='Fuzzy candidate' and pm.get(f'H{r}',1)<1)
  def fuzz(z):
   p=z.read('PRODUCT MATCHING');assert p['AA256']=='Refresh required' and p['AB256']=='';assert p[f'AA{fuzzyrow}']=='Current' and p[f'E{fuzzyrow}']==''
   return ['Changed supplier signature hides a stale fuzzy suggestion','Raising the threshold hides a different, current weak suggestion','Fuzzy suggestions remain unapproved']
  scenario('fuzzy_staleness_and_threshold',{'SUPPLIER RAW':{'D256':'Changed product name'},'SETTINGS':{'B7':1}},fuzz)
  def resolved(z):
   p=z.read('PRODUCT MATCHING');assert p['A7']=='S0002' and p['L7']=='M0002' and p['K7']=='Accepted';assert p['A8']=='S0001' and p['L8']=='M0001'
   assert z.read('SUPPLIER RAW')['M50']==0 and z.read('CLEANING')['N50']==1 and p['K50']=='Accepted'
   assert p['K222']=='Excluded duplicate' and p['R222']=='Confirmed by reviewer';assert metrics(z)['exported']==baseline+1
   return ['Keyed decisions survive full-row input reordering','Pack correction preserves raw zero and updates accepted output','Complete same-variant duplicate confirmation excludes the intended row']
  swap={}
  for c in range(1,20):swap[f'{colname(c)}7']=sraw.get(f'{colname(c)}8','');swap[f'{colname(c)}8']=sraw.get(f'{colname(c)}7','')
  scenario('stable_keys_correction_and_duplicate_resolution',{'SUPPLIER RAW':swap,'CORRECTIONS':{'A9':'S0044','J9':1,'M9':'QA corrected from catalog'},'REVIEW DECISIONS':{'B50':'Accept suggested match','C50':'M0044','D50':'QA reviewer','E50':46337},'DUPLICATE REVIEW':{'B222':'Confirm duplicate','C222':'QA confirmed repeated listing','D222':'QA reviewer','E222':46337,'F222':'S0070'}},resolved)
  def appended(z):
   p=z.read('PRODUCT MATCHING');d=z.read('DASHBOARD');assert d['D9']==251 and d['B9']==202
   assert p['E356']=='M0201' and p['K356']=='Needs review';assert p['V11']==2 and p['K11']=='Needs review' and p['U11']==0
   e=z.read('EXCEPTIONS');assert 'S0251' in [e.get('A'+str(r)) for r in range(7,357)]
   return ['Last supported master and supplier rows calculate correctly across blank reserved gaps','New native suggestion appears in exceptions','Duplicate nonblank manufacturer identifier requires review']
  nm=['MR0201','M0201','0010999','Northstar New Demo Mouse','Northstar','Mice','DEMO-NEW-QA','','','Black','N/A','STANDARD',1,99999,'IDR','Synthetic spec','Synthetic','DEMO-QA',''];ns=nm.copy();ns[0]='S0251';ns[1]='';ns[2]='VEN-QA-0001'
  mr=x.read('MASTER RAW');mu={f'{colname(i)}306':v for i,v in enumerate(nm,1)}
  for c in range(1,20):mu[f'{colname(c)}305']=mr.get(f'{colname(c)}11','')
  mu.update({'A305':'MR0202','B305':'M0202','C305':'0099999','J305':'Red'})
  scenario('append_at_capacity_and_duplicate_identifier',{'MASTER RAW':mu,'SUPPLIER RAW':{f'{colname(i)}356':v for i,v in enumerate(ns,1)}},appended)
 refreshed=tmp/'refreshed.xlsx';subprocess.run([sys.executable,str(ROOT/'scripts/refresh_productmatch.py'),'--workbook',str(source),'--output',str(refreshed)],check=True,capture_output=True)
 rz=XLSX(refreshed)
 for n in original.files:
  if n not in [original.sheets['CANDIDATES'],'xl/workbook.xml']:assert original.files[n]==rz.files[n],('Refresh changed unrelated part',n)
 results.append({'test':'External fuzzy refresh preserves manual state and native parts','result':'PASS','engine':'Python byte comparison','detail':'Only CANDIDATES and calculation properties differ'})
 before_refresh=refreshed.read_bytes()
 subprocess.run([sys.executable,str(ROOT/'scripts/refresh_productmatch.py'),'--workbook',str(refreshed),'--in-place'],check=True,capture_output=True)
 assert pathlib.Path(str(refreshed)+'.bak').read_bytes()==before_refresh
 results.append({'test':'Windows-safe in-place refresh and exact backup','result':'PASS','engine':'Python CLI / byte comparison'})
 tableparts=[n for n in original.files if n.startswith('xl/tables/')and n.endswith('.xml')]
 assert sum(ET.fromstring(original.files[n]).find('s:autoFilter',NS)is not None for n in tableparts)==13
 for name in ['PRODUCT MATCHING','REVIEW DECISIONS','RESEARCH']:
  pane=ET.fromstring(original.files[original.sheets[name]]).find('.//s:pane',NS);assert pane is not None and pane.attrib.get('ySplit')=='6'
 results.append({'test':'Final native table filters and frozen working headers','result':'PASS','engine':'OOXML inspection','detail':'13 table auto-filters; six frozen header rows; no native UI click test claimed'})
 # Checks never modify the supplied workbook or the delivered baseline evidence.
 report=a.report_dir.resolve();report.mkdir(parents=True,exist_ok=True)
 if a.recalculated_output:
  a.recalculated_output.parent.mkdir(parents=True,exist_ok=True);inject_caches(source,calc,a.recalculated_output)
 export_cached(x,report/'exports')
 # Exact saved values for portfolio rendering; no formula-based authoring cache is trusted.
 final=x;views={}
 widths={'START HERE':10,'DASHBOARD':23,'FINAL CLEAN CATALOG':24,'EXCEPTIONS':10,'PRODUCT MATCHING':43,'REVIEW DECISIONS':10,'CLEANING':32,'CORRECTIONS':15,'DUPLICATE REVIEW':8,'RESEARCH':17,'MASTER CLEAN':31,'MASTER RAW':19,'SUPPLIER RAW':19,'SETTINGS':17,'DICTIONARY':4,'QA EVALUATION':13,'CANDIDATES':13}
 for n,width in widths.items():
  values=final.read(n);height=max(int(__import__('re').search(r'\d+',a).group()) for a in values)
  views[n]={'range':f'A1:{colname(width)}{height}','values':[[values.get(f'{colname(c)}{r}',None) for c in range(1,width+1)] for r in range(1,height+1)]}
 (report/'preview_values.json').write_text(json.dumps(views),encoding='utf-8')
 assert hashlib.sha256(source.read_bytes()).hexdigest()==digest,'Verification changed the source workbook'
 evidence={'verification_date':datetime.datetime.now(datetime.timezone(datetime.timedelta(hours=7))).date().isoformat(),'python':sys.version.split()[0],'mode':'cached-only' if a.cached_only else ('baseline-recalculation' if a.quick else 'full-recalculation'),'workbook_sha256':digest,'metrics':m,'tests':results,'not_performed':['Native Excel 365 UI','Native Google Sheets UI']+(['LibreOffice recalculation and mutation scenarios'] if a.cached_only else [])}
 (report/'results.json').write_text(json.dumps(evidence,indent=2),encoding='utf-8')
 print(json.dumps({'tests_passed':len(results),'mode':evidence['mode'],'report':str(report/'results.json')},indent=2))
if __name__=='__main__':main()
