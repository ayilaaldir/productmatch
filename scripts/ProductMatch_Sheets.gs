/** Optional post-conversion repair. No live Google Sheets testing was performed. */
/** @OnlyCurrentDoc */
function onOpen() {
  SpreadsheetApp.getUi().createMenu('ProductMatch')
    .addItem('Repair navigation and validation', 'repairProductMatch')
    .addItem('Check reconciliation cells', 'checkProductMatch')
    .addToUi();
}
function repairProductMatch() {
  const ss = SpreadsheetApp.getActive();
  const required = ['START HERE','DASHBOARD','MASTER RAW','SUPPLIER RAW','SETTINGS','REVIEW DECISIONS','RESEARCH','CORRECTIONS','DUPLICATE REVIEW'];
  required.forEach(n => { if (!ss.getSheetByName(n)) throw new Error('Missing sheet: ' + n); });
  const url = ss.getUrl(), start = ss.getSheetByName('START HERE'), dash = ss.getSheetByName('DASHBOARD');
  ss.getSheets().forEach(s => {
    if (s.getName() !== 'START HERE') s.getRange('A3').setFormula('=HYPERLINK("'+url+'#gid='+start.getSheetId()+'&range=A1","← Start here")');
    s.getRange('D3').setFormula('=HYPERLINK("'+url+'#gid='+dash.getSheetId()+'&range=A1","Dashboard →")');
  });
  const nav = ss.getSheets().filter(s=>s.getName()!=='START HERE');
  nav.forEach((s,i)=>start.getRange(8+i,4).setFormula('=HYPERLINK("'+url+'#gid='+s.getSheetId()+'&range=A1","'+s.getName()+'")'));
  const sets=ss.getSheetByName('SETTINGS');
  const rangeRule=(r)=>SpreadsheetApp.newDataValidation().requireValueInRange(r,true).setAllowInvalid(false).build();
  const listRule=(a)=>SpreadsheetApp.newDataValidation().requireValueInList(a,true).setAllowInvalid(false).build();
  dash.getRange('B5').setDataValidation(rangeRule(sets.getRange('J7:J16')));
  dash.getRange('D5').setDataValidation(rangeRule(sets.getRange('K7:K30')));
  dash.getRange('F5').setDataValidation(rangeRule(sets.getRange('L7:L12')));
  const review=ss.getSheetByName('REVIEW DECISIONS');
  review.getRange('B7:B356').setDataValidation(listRule(['Accept suggested match','Select different master','Reject suggestion','Leave pending']));
  review.getRange('C7:C356').setDataValidation(rangeRule(sets.getRange('N7:N306')));
  ['REVIEW DECISIONS','CORRECTIONS','DUPLICATE REVIEW'].forEach(n=>ss.getSheetByName(n).getRange('A7:A356').setDataValidation(rangeRule(sets.getRange('O7:O356'))));
  const research=ss.getSheetByName('RESEARCH');
  research.getRange('G7:G106').setDataValidation(listRule(['Verified','Not found','Conflicting sources','Pending']));
  research.getRange('H7:H106').setDataValidation(listRule(['Yes','No']));
  research.getRange('B7:B106').setDataValidation(rangeRule(sets.getRange('N7:N306')));
  ss.getSheetByName('DUPLICATE REVIEW').getRange('B7:B356').setDataValidation(listRule(['Pending','Keep separate','Confirm duplicate']));
  ss.getSheetByName('DUPLICATE REVIEW').getRange('F7:F356').setDataValidation(rangeRule(sets.getRange('O7:O356')));
  SpreadsheetApp.flush();
  ss.toast('Navigation and validation repaired. Complete the manual conversion tests before claiming parity.','ProductMatch');
}
function checkProductMatch() {
  const ss=SpreadsheetApp.getActive(), qa=ss.getSheetByName('QA EVALUATION');
  if (!qa) throw new Error('Missing QA EVALUATION sheet');
  SpreadsheetApp.flush();
  const checks=qa.getRange('A7:D14').getDisplayValues();
  const failures=checks.filter(r=>r[3]!=='PASS');
  SpreadsheetApp.getUi().alert(failures.length ? failures.map(r=>r[0]+': '+r[3]).join('\n') : 'Reconciliation cells show PASS. This is not a substitute for selector, decision, variant and research tests.');
}
