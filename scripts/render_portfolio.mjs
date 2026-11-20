// Render exact saved, recalculated values with the workbook's own layout.
// Run verify_productmatch.py first; it writes preview_values.json.
import fs from 'node:fs/promises';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { parseArgs } from 'node:util';
import {SpreadsheetFile,FileBlob} from '@oai/artifact-tool';
const root=path.resolve(path.dirname(fileURLToPath(import.meta.url)),'..');
const {values:args}=parseArgs({options:{workbook:{type:'string'},values:{type:'string'},'output-dir':{type:'string'}}});
const dir=path.resolve(args['output-dir']||path.join(root,'.artifacts/previews'));
const input=path.resolve(args.workbook||path.join(root,'ProductMatch.xlsx'));
const values=path.resolve(args.values||path.join(root,'.artifacts/verification/preview_values.json'));
const wb=await SpreadsheetFile.importXlsx(await FileBlob.load(input));
const snapshots=JSON.parse(await fs.readFile(values,'utf8'));
for(const[n,d]of Object.entries(snapshots)){wb.worksheets.getItem(n).getRange(d.range).values=d.values;console.log('Loaded evaluated view '+n);}
// This in-memory visual snapshot is NOT exported as the workbook deliverable.
// The delivered workbook retains all original native formulas.
const ranges={'START HERE':'A1:D14','DASHBOARD':'A1:J33','PRODUCT MATCHING':'A1:L14','RESEARCH':'A1:L11','FINAL CLEAN CATALOG':'A1:H14','EXCEPTIONS':'A1:G14','REVIEW DECISIONS':'A1:J14','CLEANING':'A1:G14','CORRECTIONS':'A1:G12','DUPLICATE REVIEW':'A1:H13','MASTER CLEAN':'A1:H14','MASTER RAW':'A1:H14','SUPPLIER RAW':'A1:H14','SETTINGS':'A1:H14','DICTIONARY':'A1:D13','QA EVALUATION':'A1:D15','CANDIDATES':'A1:I12'};
await fs.mkdir(dir+'/screenshots',{recursive:true});await fs.mkdir(dir+'/qa_previews',{recursive:true});
for(const[n,range]of Object.entries(ranges)){
 const image=await wb.render({sheetName:n,range,scale:1,format:'png'});
 const file=n==='DASHBOARD'?'screenshots/01_dashboard.png':n==='PRODUCT MATCHING'?'screenshots/02_matching_review.png':n==='RESEARCH'?'screenshots/03_research_evidence.png':`qa_previews/${n.toLowerCase().replaceAll(' ','_')}.png`;
 await fs.writeFile(dir+'/'+file,new Uint8Array(await image.arrayBuffer()));console.log('Rendered '+n);
}
