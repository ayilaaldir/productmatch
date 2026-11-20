#!/usr/bin/env python3
"""Deterministic synthetic trading scenario; official research values kept separate."""
import argparse, csv, json, pathlib, copy, re
parser=argparse.ArgumentParser(description=__doc__)
parser.add_argument('--output-dir',type=pathlib.Path,default=pathlib.Path(__file__).resolve().parents[1]/'.artifacts/generated',help='Destination tree; defaults to .artifacts/generated without changing delivered data')
ROOT=parser.parse_args().output_dir.resolve()
for folder in ['data/raw','data/reference','data/evaluation','data/generated']:(ROOT/folder).mkdir(parents=True,exist_ok=True)
RAW_HEADERS=['Source row ID','Product ID','SKU','Product name','Brand','Category','Model','MPN','GTIN','Color','Capacity','Size','Pack quantity','Unit price','Currency','Specification','Data disclosure','Provenance reference','Source URL']
REAL=[
 ('Logitech','M185','Wireless Mouse M185','Mice','Blue','N/A','Compact','', 'Wireless optical mouse', 'https://www.logitech.com/id-id/shop/p/m185-wireless-mouse?crid=7','Logitech M185 Compact Wireless Mouse','Official page identifies M185 as a compact wireless mouse and the displayed blue variant.'),
 ('Logitech','K120','K120 Corded Keyboard','Keyboards','Black','N/A','Full size','920-002582','USB wired keyboard with number pad','https://www.logitech.com/id-id/shop/p/k120-usb-standard-computer.920-002582','Logitech K120 USB Standard Computer Keyboard','K120 page identifies the USB connection and integrated number pad; regional part is in page URL.'),
 ('Logitech','C270','C270 HD Webcam','Webcams','Dark Grey','N/A','Standard','','720p at 30 fps','https://www.logitech.com/id-id/shop/p/c270-hd-webcam','Logitech C270 HD Webcam','C270 product description specifies 720p/30fps. Displayed color is dark grey.'),
 ('Logitech','H390','H390 USB Headset','Audio','Black','N/A','Standard','','USB headset with noise-canceling microphone','https://www.logitech.com/id-id/shop/p/h390-usb-headset','Headset USB Logitech H390','Page identifies H390 and microphone feature. USB connector subtype is deliberately not inferred from a multi-variant page.'),
 ('Kingston','DTX/64GB','DataTraveler Exodia 64GB','Storage','Black + Teal','64GB','Standard','DTX/64GB','USB 3.2 Gen 1; 64GB','https://www.kingston.com/id/memory/search/discontinuedmodels?partId=DTX/64GB','Kingston DTX/64GB discontinued model','Exact part lookup identifies DTX/64GB, 64GB, Black + Teal and USB 3.2 Gen 1. Discontinued status does not invalidate specification.'),
 ('TP-Link','TL-SG105','TL-SG105 Desktop Switch','Networking','N/A','N/A','Standard','','5 x 10/100/1000 Mbps ports','https://www.tp-link.com/id/business-networking/unmanaged-switch/tl-sg105/','TL-SG105 | 5-Port Gigabit Desktop Switch','Hardware interface section identifies five gigabit ports. Power is not applied because hardware version is unspecified.'),
 ('TP-Link','TL-SG108','TL-SG108 Desktop Switch V9','Networking','N/A','N/A','V9','','8 x 10/100/1000 Mbps ports','https://www.tp-link.com/id/home-networking/soho-switch/tl-sg108/v9/','TL-SG108 | 8-Port Gigabit Desktop Switch V9','V9 page explicitly identifies eight gigabit ports. Hardware version is retained in Size/version attribute.'),
 ('TP-Link','Archer T2U Nano','Archer T2U Nano Wi-Fi Adapter','Networking','N/A','N/A','Standard','','USB 2.0; dual-band Wi-Fi adapter','https://www.tp-link.com/id/home-networking/adapter/archer-t2u-nano/','Archer T2U Nano | AC600 Nano Wireless USB Adapter','Hardware specifications show USB 2.0; wireless frequencies include 2.4GHz and 5GHz. Signal rates are not throughput promises.'),
 ('TP-Link','UB500','UB500 Bluetooth Adapter V1','Networking','N/A','N/A','V1','','Bluetooth 5.0 (V1)','https://www.tp-link.com/us/home-networking/usb-adapter/ub500/v1/','UB500 | Bluetooth 5.0 Nano USB Adapter V1','Version-specific V1 page specifies Bluetooth 5.0. Not generalized to later hardware revisions.'),
 ('TP-Link','TL-WN725N','TL-WN725N Wireless Nano Adapter','Networking','N/A','N/A','Standard','','USB 2.0; 2.4GHz wireless adapter','https://www.tp-link.com/us/home-networking/usb-adapter/tl-wn725n/','TL-WN725N | 150Mbps Wireless N Nano USB Adapter','Hardware interface is USB 2.0; wireless frequency range is 2.400-2.4835GHz.'),
 ('SanDisk','SDSDXXD-128G-GN4IN','Extreme PRO UHS-I SD 128GB','Storage','N/A','128GB','SD','SDSDXXD-128G-GN4IN','UHS-I; up to 200MB/s read','https://www.sandisk.com/in-id/products/memory-cards/sd-cards/sandisk-extreme-pro-uhs-i-sd?sku=SDSDXXD-128G-GN4IN','128GB SanDisk Extreme PRO SD UHS-I','Official indexed specification block ties SDSDXXD-128G-GN4IN to 128GB and up to 200MB/s read; direct page request timed out. Form-factor label is not applied because page labels conflict.'),
 ('Crucial','CT500MX500SSD1','MX500 500GB SATA SSD','Storage','N/A','500GB','2.5 inch','CT500MX500SSD1','500GB; 2.5-inch SATA 6Gb/s SSD','https://www.crucial.com/content/dam/crucial/ssd-products/mx500/flyer/crucial-mx500-ssd-productflyer-letter-en.pdf','Crucial MX500 SSD product flyer','Manufacturer indexed table identifies CT500MX500SSD1 as 500GB 2.5-inch SATA 6Gb/s. Direct open redirected to a reseller mirror; official indexed table used, not mirror.'),
 ('Samsung','MU-PC1T0T/AM','Portable SSD T7 1TB Gray','Storage','Gray','1TB','Standard','MU-PC1T0T/AM','USB 3.2 Gen 2; 1TB portable SSD','https://www.samsung.com/us/memory-storage/portable-ssd/portable-ssd-t7-usb-3-2-1tb-gray-sku-mu-pc1t0t-am/','Samsung Portable SSD T7 1TB Gray','Page ties exact /AM part to 1TB Gray and USB 3.2 Gen 2. No /WW regional substitution.'),
 ('Epson','L3250','EcoTank L3250 All-in-One Printer','Printing','N/A','N/A','A4','C11CJ67503','Print, scan, copy; Wi-Fi and Wi-Fi Direct','https://www.epson.co.id/Untuk-Rumah/Printer-untuk-Rumah/Home-Office-Printers/Printer-Epson-EcoTank-L3250-A4-Wi-Fi-All-in-One-Ink-Tank/p/C11CJ67503','Epson EcoTank L3250 A4 Wi-Fi All-in-One Ink Tank','Exact Indonesian product page and part C11CJ67503 identify print/scan/copy, Wi-Fi and Wi-Fi Direct.'),
 ('Brother','HL-L2350DW','HL-L2350DW Mono Laser Printer','Printing','N/A','N/A','US','','Automatic duplex monochrome laser printing','https://www.brother-usa.com/p/laser-printers/HLL2350DW','Brother HL-L2350DW Wireless Compact Monochrome Laser Printer','US page explicitly identifies automatic two-sided monochrome laser printing. US variant retained; voltage is not generalized.'),
]
ALIASES={'logitech':'Logitech','logi tech':'Logitech','logitec':'Logitech','kingston':'Kingston','king stone':'Kingston','tp-link':'TP-Link','tp link':'TP-Link','tplink':'TP-Link','sandisk':'SanDisk','san disk':'SanDisk','crucial':'Crucial','samsung':'Samsung','epson':'Epson','brother':'Brother','northstar':'Northstar','north star':'Northstar','vertex':'Vertex','vert ex':'Vertex','luma':'Luma','lumma':'Luma','oakline':'Oakline','oak line':'Oakline'}
CATS=['Mice','Keyboards','Webcams','Audio','Storage','Networking','Office accessories','Cables','Printing']
def clean(x): return ' '.join(str(x or '').replace('\xa0',' ').split())
def canon(x): return ALIASES.get(clean(x).lower(),clean(x).title())
def row(pid,name,brand,cat,model,mpn='',color='Black',cap='N/A',size='Standard',pack=1,price=100000,real=False,url=''):
 return ['MR'+pid[1:],pid,'00'+str(10000+int(pid[1:])),name,brand,cat,model,mpn,'',color,cap,size,pack,price,'IDR','' if real else 'Synthetic specification; no manufacturer claim','Real model / simulated offer' if real else 'Synthetic','OFFICIAL-'+pid if real else 'DEMO-MASTER-20261111',url]
master=[]
for i,p in enumerate(REAL,1):
 b,model,name,cat,col,cap,size,mpn,*_=p
 master.append(row(f'M{i:04}',name,b,cat,model,mpn,col,cap,size,1,50000+i*65000,True,p[9]))
for i in range(16,199):
 family=(i-16)//2
 cat=CATS[family%len(CATS)]; brand=['Northstar','Vertex','Luma','Oakline'][family%4]
 model=f'DEMO-{["MS","KB","WC","AU","ST","NW","OF","CB","PR"][family%9]}-{family+1:03}'
 variant=(i-16)%2
 col=['Black','White'][variant]; cap=['64GB','128GB'][variant] if cat=='Storage' else 'N/A'
 size=['1M','2M'][variant] if cat=='Cables' else ('A4' if cat=='Printing' else 'Standard')
 pack=[1,3][variant] if cat=='Office accessories' else 1
 name=f'{brand} {model} {cat.rstrip("s")} {col} {cap if cap!="N/A" else size}'
 m=row(f'M{i:04}',name,brand,cat,model,'',col,cap,size,pack,75000+family*2500+variant*10000)
 if i%9==0: m[3]='  '+m[3].upper().replace(' ','  ')+' '
 if i%13==0: m[4]={'Northstar':'north star','Vertex':'VERT EX','Luma':'lumma','Oakline':'oak line'}[brand]
 master.append(m)
for i,source in [(199,40),(200,80)]:
 m=copy.deepcopy(master[source-1]);m[0]=f'MR{i:04}';m[1]=f'M{i:04}';m[2]='00'+str(10000+i);master.append(m)
supplier=[];truth=[]
def add(m,target,issue,alter=None):
 n=len(supplier)+1;s=copy.deepcopy(m);s[0]=f'S{n:04}';s[1]='';s[2]=f'VEN-{n:05}';s[17]='DEMO-SUPPLIER-20261111'
 if alter: alter(s)
 supplier.append(s);truth.append([s[0],target,issue,'','Deliberately constructed demonstration ground truth; not used by matching.'])
for i in range(1,181):
 m=master[i-1]
 def alter(s,i=i):
  if i%4==0:s[3]='  '+clean(s[3]).upper().replace(' ','  ')+'  '
  if i%5==0:s[4]={'Logitech':'logi tech','Kingston':'king stone','TP-Link':'tp link','SanDisk':'san disk','Northstar':'north star','Vertex':'vert ex','Luma':'lumma','Oakline':'oak line'}.get(canon(s[4]),s[4].lower())
  if i%7==0:s[3]=' '.join(clean(s[3]).split()[::-1])
  if i in [19,31,47,55,66,81,92,107,119,143,159,171]:s[6]='';s[7]=''
  if i in [22,63,96,131]:s[4]=''
  if i in [28,88,128]:s[9]=''
  if i in [44,94,134]:s[12]=0
  if i in [58,108]:s[13]=-25000
  if i in [68,118]:s[14]='US$'
  if i in [38,78]:s[8]='NOT_PROVIDED'
  if i in [52,112]:s[5]='office misc'
 add(m,m[1],'Whitespace / alias / reorder / selected missing or invalid fields',alter)
# Twenty deliberate wrong-model or variant cases. Some have a correct alternative.
for j in range(20):
 mi=16+j*2;m=master[mi-1]
 if j<10:
  def alt(s,mi=mi):s[6]=master[mi][6];s[9]=master[mi][9];s[10]=master[mi][10];s[11]=master[mi][11];s[12]=master[mi][12];s[6]='' if mi<22 else s[6]
  add(m,master[mi][1],'Same family, different legitimate variant; name remains misleading',alt)
 else:
  def alt(s):s[9]='Red';s[12]=2
  add(m,'NONE','High name similarity but unavailable color / pack variant',alt)
# Fifteen exact retransmissions: identical business fields including supplier SKU and price.
for k in range(15):
 s=copy.deepcopy(supplier[k]);n=len(supplier)+1;s[0]=f'S{n:04}';supplier.append(s)
 truth.append([s[0],truth[k][1],'Exact supplier retransmission',supplier[k][0],'Excluded duplicate; original source row ID is survivor.'])
# Ten suspected duplicates: different SKUs, same identity/variant; need explicit decision.
for k in range(10):add(master[69+k],master[69+k][1],'Same product with another supplier SKU; suspected duplicate')
# Ten no-counterpart synthetic offers.
for j in range(10):
 m=row('M0300',f'Oakline DEMO-NEW-{j+1:03} Presentation Remote','Oakline','Office accessories',f'DEMO-NEW-{j+1:03}',color='Green',size='Pocket',pack=1,price=120000+j*5000)
 add(m,'NONE','No product in master catalog')
# Fifteen fuzzy-only offers; original model omitted and names slightly changed.
for j in range(15):
 m=master[119+j]
 def alt(s):s[6]='';s[7]='';s[3]=clean(s[3]).replace('Standard','Compact').replace('Black','Blk').replace('White','Wht')
 add(m,m[1],'Missing model; fuzzy candidate only',alt)
assert len(master)==200 and len(supplier)==250
research=[]
for i,p in enumerate(REAL,1):
 b,model,name,cat,col,cap,size,mpn,spec,url,title,note=p
 research.append([f'R{i:03}',f'M{i:04}',model,'Specification','',spec,'Verified','Yes',note,url,title,'2026-11-11','','Official source checked; approval is seeded demonstration state'])
# Genuine unresolved source conflicts, not fabricated claims.
research.append(['R016','M0006','TL-SG105','Power','', '', 'Conflicting sources','No','Official pages show 3.0 W versus 2.3 W for different hardware versions. Hardware version absent from this offer; do not combine.','https://www.tp-link.com/id/business-networking/unmanaged-switch/tl-sg105/','TL-SG105 specification variants','2026-11-11','https://www.tp-link.com/us/home-networking/5-port-switch/tl-sg105/v6/','Specify hardware revision before applying power data.'])
research.append(['R017','M0011','SDSDXXD-128G-GN4IN','Form factor','','','Conflicting sources','No','Product page labels a 128GB block SDHC; official compatibility guidance places 64GB-2TB in SDXC. Do not apply the contradictory label.','https://www.sandisk.com/in-id/products/memory-cards/sd-cards/sandisk-extreme-pro-uhs-i-sd?sku=SDSDXXD-128G-GN4IN','SanDisk Extreme PRO 128GB / compatibility guidance','2026-11-11','https://support-en.sandisk.com/app/answers/detailweb/a_id/36273/~/specifications-and-compatibility-for-sd%252Fsdhc%252Fsdxc-memory-cards','Exact part checked through official search index; no form-factor update applied.'])
research.append(['R018','M0001','M185','GTIN','','','Not found','No','Reviewed official page did not expose a GTIN. A page existing is not sufficient verification.','https://www.logitech.com/id-id/shop/p/m185-wireless-mouse?crid=7','Logitech M185 Compact Wireless Mouse','2026-11-11','','Request package barcode for the exact regional variant.'])
for n,pid in enumerate([2,4,7,9],19):
 research.append([f'R{n:03}',f'M{pid:04}',REAL[pid-1][1],'GTIN','','','Pending','No','Research queued; exact regional pack required.','','',None,'','No claim of research completion for this field.'])
decisions=[]
for i,s in enumerate(supplier,1):
 decision='Leave pending';alternate='';reviewer='';date=None;note=''
 if i<=165 and i not in [19,22,28,31,38,44,47,52,55,58,63,66,68,78,81,88,92,94,96,107,108,112,118,119,128,131,134,143,159]:
  decision='Accept suggested match';alternate=f'M{i:04}';reviewer='Demo reviewer';date='2026-11-11';note='Simulated review: identifiers and variant checked.'
 if i in [181,182,183]:
  decision='Select different master';alternate=truth[i-1][1];reviewer='Demo reviewer';date='2026-11-11';note='Simulated override: name misleading; selected matching variant.'
 if i in [191,192,193,194,195]:decision='Reject suggestion';reviewer='Demo reviewer';date='2026-11-11';note='Simulated review: offered variant absent from master.'
 decisions.append([s[0],decision,alternate,reviewer,date,note])
corrections=[]
for sid,b in [('S0022','Northstar'),('S0063',canon(master[62][4]))]:
 # Use intended source brand, never invent a manufacturer fact.
 b=canon(master[int(sid[1:])-1][4]); corrections.append([sid,'',b,'','','','','','','','','','Simulated correction using internal catalog','Demo reviewer','2026-11-11'])
duplicates=[]
for i,s in enumerate(supplier,1):
 status='Keep separate' if i in [221,222,223,224,225] else 'Pending'
 duplicates.append([s[0],status,'Distinct supplier listing retained (simulated decision)' if status=='Keep separate' else '', 'Demo reviewer' if status!='Pending' else '', '2026-11-11' if status!='Pending' else None])
def writecsv(path,headers,rows):
 with open(ROOT/path,'w',encoding='utf-8-sig',newline='') as f:csv.writer(f).writerows([headers]+rows)
writecsv('data/raw/master_raw.csv',RAW_HEADERS,master);writecsv('data/raw/supplier_raw.csv',RAW_HEADERS,supplier)
writecsv('data/reference/research_evidence.csv',['Research ID','Product ID','Exact model','Attribute','Original value','Proposed value','Verification','Apply approved','Evidence note','Source URL','Page title','Date accessed','Other source URL','Conflict / workflow notes'],research)
writecsv('data/evaluation/ground_truth.csv',['Supplier row ID','Expected master ID','Constructed issue','Duplicate of','Evaluation note'],truth)
data=dict(master=master,supplier=supplier,research=research,truth=truth,decisions=decisions,corrections=corrections,duplicates=duplicates,aliases=ALIASES,categories=CATS,rawHeaders=RAW_HEADERS)
(ROOT/'data/generated/demo_data.json').write_text(json.dumps(data,indent=2),encoding='utf-8')
print(json.dumps({'master':len(master),'supplier':len(supplier),'real_models':len(REAL),'research_fields':len(research)}))
if __name__=='__main__':pass
