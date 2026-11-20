#!/usr/bin/env python3
"""Refresh only fuzzy-candidate snapshot; preserve all manual decisions and workbook parts.
Standard library only. Run after saving/closing the workbook. See README.md.
"""
import argparse, csv, datetime, difflib, json, pathlib, re, shutil, zipfile
import xml.etree.ElementTree as ET
NS={'s':'http://schemas.openxmlformats.org/spreadsheetml/2006/main','r':'http://schemas.openxmlformats.org/officeDocument/2006/relationships'}
S=NS['s']; ET.register_namespace('',S)
def clean(x):return ' '.join(str('' if x is None else x).replace('\xa0',' ').split())
def upper(x):return clean(x).upper()
def colnum(c):
 n=0
 for a in c:n=n*26+ord(a)-64
 return n
def colname(n):
 s=''
 while n:n,r=divmod(n-1,26);s=chr(65+r)+s
 return s
class XLSX:
 def __init__(self,path):
  self.path=path
  # Close the input archive before writing: Windows locks open source files.
  with zipfile.ZipFile(path) as archive:self.files={n:archive.read(n) for n in archive.namelist()}
  self.shared=[]
  if 'xl/sharedStrings.xml' in self.files:
   for si in ET.fromstring(self.files['xl/sharedStrings.xml']).findall('s:si',NS):self.shared.append(''.join(t.text or '' for t in si.iter('{'+S+'}t')))
  rels={r.attrib['Id']:r.attrib['Target'] for r in ET.fromstring(self.files['xl/_rels/workbook.xml.rels'])}
  self.sheets={}
  for sh in ET.fromstring(self.files['xl/workbook.xml']).findall('s:sheets/s:sheet',NS):
   target=rels[sh.attrib['{'+NS['r']+'}id']];self.sheets[sh.attrib['name']]=target.lstrip('/') if target.startswith('/') else 'xl/'+target
 def read(self,name):
  out={}
  for c in ET.fromstring(self.files[self.sheets[name]]).findall('.//s:sheetData/s:row/s:c',NS):
   t=c.attrib.get('t');v=c.find('s:v',NS)
   if t=='inlineStr':val=''.join(x.text or '' for x in c.iter('{'+S+'}t'))
   elif t=='s':val=self.shared[int(v.text)] if v is not None and v.text else ''
   elif v is None or v.text is None:val=''
   elif t in ['str','e']:val=v.text
   else:
    try:val=float(v.text);val=int(val) if val.is_integer() else val
    except ValueError:val=v.text
   out[c.attrib['r']]=val
  return out
 def matrix(self,name,start=7,cap=350,width=19):
  d=self.read(name);return [[d.get(f'{colname(c)}{r}','') for c in range(1,width+1)] for r in range(start,start+cap)]
 def set_cells(self,name,updates):
  p=self.sheets[name];root=ET.fromstring(self.files[p]);sd=root.find('s:sheetData',NS)
  rows={int(r.attrib['r']):r for r in sd}
  for addr,value in updates.items():
   letters,rn=re.match(r'([A-Z]+)(\d+)',addr).groups();rn=int(rn)
   if rn not in rows:rows[rn]=ET.SubElement(sd,'{'+S+'}row',{'r':str(rn)})
   row=rows[rn];cell=next((c for c in row if c.attrib.get('r')==addr),None)
   if cell is None:cell=ET.SubElement(row,'{'+S+'}c',{'r':addr})
   for child in list(cell):cell.remove(child)
   cell.attrib.pop('t',None)
   if value is None or value=='':continue
   if isinstance(value,(int,float)):
    ET.SubElement(cell,'{'+S+'}v').text=str(value)
   else:
    cell.attrib['t']='inlineStr';istr=ET.SubElement(cell,'{'+S+'}is');ET.SubElement(istr,'{'+S+'}t').text=str(value)
  for row in sd:row[:]=sorted(row,key=lambda c:colnum(re.match('[A-Z]+',c.attrib['r']).group()))
  sd[:]=sorted(sd,key=lambda r:int(r.attrib['r']))
  self.files[p]=ET.tostring(root,encoding='utf-8',xml_declaration=True)
 def save(self,path):
  root=ET.fromstring(self.files['xl/workbook.xml']);cp=root.find('s:calcPr',NS)
  if cp is None:cp=ET.SubElement(root,'{'+S+'}calcPr')
  cp.set('calcMode','auto');cp.set('fullCalcOnLoad','1');cp.set('forceFullCalc','1')
  self.files['xl/workbook.xml']=ET.tostring(root,encoding='utf-8',xml_declaration=True)
  with zipfile.ZipFile(path,'w',zipfile.ZIP_DEFLATED) as z:
   for n,b in self.files.items():z.writestr(n,b)
def normalize(raw,aliases,catmap=None,corr=None):
 a=list(raw);corr=corr or {};catmap=catmap or {}
 # Corrections columns: ID, Name, Brand, Category, Model, MPN, Color, Capacity, Size, Pack, Price, Currency.
 for ci,ri in [(1,3),(2,4),(3,5),(4,6),(5,7),(6,9),(7,10),(8,11),(9,12),(10,13),(11,14)]:
  if clean(corr.get(ci,'')):a[ri]=corr[ci]
 brand=aliases.get(clean(a[4]).lower(),clean(a[4]).title())
 cat=catmap.get(clean(a[5]).lower(),clean(a[5]).title())
 if clean(a[5]).lower()=='office misc':cat='Office accessories'
 # Match workbook mapped display names rather than title-casing categories.
 for v in ['Mice','Keyboards','Webcams','Audio','Storage','Networking','Office accessories','Cables','Printing']:
  if cat.lower()==v.lower():cat=v
 d=dict(id=a[0],pid=a[1],sku=clean(a[2]),name=clean(a[3]).title(),brand=brand,category=cat,model=upper(a[6]),mpn=upper(a[7]),gtin=clean(a[8]).replace(' ','').replace('-',''),color=clean(a[9]).title(),capacity=upper(a[10]).replace(' ',''),size=upper(a[11]),pack=a[12],price=a[13],currency=upper(a[14]))
 sig=[upper(d['name']),d['brand'],d['category'],d['model'],d['mpn'],upper(d['color']),d['capacity'],d['size'],str(d['pack'])]
 d['signature']='|'.join(sig)
 d['key']='|'.join([brand,d['model'],upper(d['color']),d['capacity'],d['size'],str(d['pack'])]) if all([brand,d['model'],d['color'],d['capacity'],d['size'],d['pack']]) else ''
 return d
def fuzzy_score(a,b):
 def tokens(s):return re.findall(r'[a-z0-9]+',s.lower())
 ta,tb=tokens(a),tokens(b);sa,sb=set(ta),set(tb)
 j=len(sa&sb)/len(sa|sb) if sa|sb else 0
 seq=difflib.SequenceMatcher(None,' '.join(sorted(ta)),' '.join(sorted(tb)),autojunk=False).ratio()
 return round(0.6*seq+0.4*j,6)
def candidates(master,supplier,aliases,corrections=None,catmap=None):
 corrections=corrections or {};masters=[normalize(m,aliases,catmap) for m in master if clean(m[0])];suppliers=[normalize(s,aliases,catmap,corrections.get(s[0])) for s in supplier if clean(s[0])]
 seen=set();active=[]
 for m in masters:
  if m['key'] and m['key'] in seen:continue
  seen.add(m['key']);active.append(m)
 now=datetime.datetime.now(datetime.timezone.utc).replace(microsecond=0).isoformat()
 rows=[]
 for s in suppliers:
  pool=[m for m in active if m['category']==s['category'] and (not s['brand'] or m['brand']==s['brand'])]
  ranked=sorted([(fuzzy_score(s['name'],m['name']),m['pid']) for m in pool],key=lambda a:(-a[0],a[1]))[:3]
  row=[s['id'],s['signature']]
  for i in range(3):row+=([ranked[i][1],ranked[i][0]] if i<len(ranked) else ['',0])
  rows.append(row+[now])
 return dict(rows=rows,masterSnapshot=[[m['pid'],m['signature']] for m in masters],generatedUTC=now)
def export_cached(x,out):
 out.mkdir(parents=True,exist_ok=True)
 for sheet,filename,width in [('FINAL CLEAN CATALOG','final_clean_catalog.csv',24),('EXCEPTIONS','exceptions.csv',10),('PRODUCT MATCHING','matching_review.csv',20),('RESEARCH','research_current.csv',14)]:
  cells=x.read(sheet);rows=[]
  for r in range(6,357 if sheet!='RESEARCH' else 107):
   if r==6 or cells.get('A'+str(r),'')!='':
    row=[cells.get(f'{colname(c)}{r}','') for c in range(1,width+1)]
    datecols=[11] if sheet=='RESEARCH' else ([15] if sheet=='PRODUCT MATCHING' else [])
    for j in datecols:
     if r>6 and isinstance(row[j],(int,float)) and row[j]>0:
      row[j]=(datetime.datetime(1899,12,30)+datetime.timedelta(days=row[j])).date().isoformat()
    rows.append(row)
  with (out/filename).open('w',encoding='utf-8-sig',newline='') as f:csv.writer(f).writerows(rows)
def main():
 p=argparse.ArgumentParser(description=__doc__);p.add_argument('--data');p.add_argument('--workbook');p.add_argument('--output');p.add_argument('--in-place',action='store_true');p.add_argument('--export-cached');a=p.parse_args()
 if a.data:
  d=json.loads(pathlib.Path(a.data).read_text(encoding='utf-8'));cor={r[0]:dict(enumerate(r)) for r in d['corrections']};c=candidates(d['master'],d['supplier'],d['aliases'],cor)
  output=pathlib.Path(a.output or 'candidate_snapshot.json');output.parent.mkdir(parents=True,exist_ok=True)
  output.write_text(json.dumps(c,indent=2),encoding='utf-8');print('Generated',len(c['rows']),'candidate snapshots.');return
 if not a.workbook:p.error('--workbook or --data required')
 x=XLSX(a.workbook)
 if a.export_cached:export_cached(x,pathlib.Path(a.export_cached));print('Exported saved formula results; recalculate and save in Excel BEFORE this command.');return
 if not a.output and not a.in_place:p.error('Use --output NEW.xlsx or --in-place (creates .bak)')
 if a.output and a.in_place:p.error('Choose --output or --in-place, not both')
 if a.output and pathlib.Path(a.output).resolve()==pathlib.Path(a.workbook).resolve():p.error('Use --in-place to update the input with a backup')
 sets=x.read('SETTINGS');aliases={clean(sets.get(f'D{r}')).lower():sets.get(f'E{r}') for r in range(7,57) if sets.get(f'D{r}')};catmap={clean(sets.get(f'G{r}')).lower():sets.get(f'H{r}') for r in range(7,37) if sets.get(f'G{r}')}
 corrections={r[0]:dict(enumerate(r)) for r in x.matrix('CORRECTIONS',width=15) if r[0]}
 master=x.matrix('MASTER RAW',cap=300);supplier=x.matrix('SUPPLIER RAW',cap=350)
 c=candidates(master,supplier,aliases,corrections,catmap)
 updates={}
 for r in range(7,357):
  row=c['rows'][r-7] if r-7<len(c['rows']) else ['']*9
  for j,v in enumerate(row,1):updates[f'{colname(j)}{r}']=v
 for r in range(7,307):
  row=c['masterSnapshot'][r-7] if r-7<len(c['masterSnapshot']) else ['','']
  for j,v in enumerate(row,12):updates[f'{colname(j)}{r}']=v
 x.set_cells('CANDIDATES',updates)
 path=pathlib.Path(a.workbook if a.in_place else a.output)
 if a.in_place:shutil.copy2(path,str(path)+'.bak')
 x.save(path)
 # All sheets except snapshot and calc properties retain byte-identical XML.
 check=XLSX(path)
 for sh in ['REVIEW DECISIONS','CORRECTIONS','DUPLICATE REVIEW','RESEARCH']:
  assert check.files[check.sheets[sh]]==x.files[x.sheets[sh]],f'Manual sheet changed: {sh}'
 print(f'Refreshed {len(c["rows"])} suppliers; preserved manual decisions. Open in Excel, calculate, save; then export cached CSVs.')
if __name__=='__main__':main()
