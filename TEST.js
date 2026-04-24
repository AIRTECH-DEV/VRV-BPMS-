// // Step 15 TAT is in BR13 (col 70), not BS13 (col 71) — merged cell issue
// // This fixes all Step 15 Planned formulas to reference BR$13

// function fixStep15TATRef() {
//   const ss    = SpreadsheetApp.getActiveSpreadsheet();
//   const bpms  = ss.getSheetByName('BPMS VRV');
//   const tmpl  = ss.getSheetByName('BPMS-Template');
//   const lastRow = bpms.getLastRow();

//   // Correct formula: uses BR$13 (where TAT value 16 actually lives)
//   function makeFormula(row) {
//     return '=IF(OR(ISBLANK(BN' + row + '),ISBLANK(BR$13)),"",LET(' +
//       '_start,BN' + row + ',' +
//       '_tat,BR$13/24,' +
//       '_open,$C$8,_close,$D$8,_wd,$E$8,' +
//       '_hol,Holidays!$A$3:$A,' +
//       '_sd,INT(_start),_st,MOD(_start,1),' +
//       '_iswd,WORKDAY.INTL(_sd-1,1,_wd,_hol)=_sd,' +
//       '_first,IF(_iswd,IF(_st<_open,_sd+_open,IF(_st>=_close,' +
//       'WORKDAY.INTL(_sd,1,_wd,_hol)+_open,_start)),' +
//       'WORKDAY.INTL(_sd,1,_wd,_hol)+_open),' +
//       '_ft,MOD(_first,1),_avail,_close-MAX(_ft,_open),' +
//       'IF(_tat<=_avail,_first+_tat,' +
//       'LET(_rem1,_tat-_avail,_daylen,_close-_open,' +
//       '_k,INT(_rem1/_daylen),_rem2,MOD(_rem1,_daylen),' +
//       '_base,WORKDAY.INTL(INT(_first),1+_k,_wd,_hol),' +
//       'IF(_rem2=0,WORKDAY.INTL(INT(_first),_k,_wd,_hol)+_close,' +
//       '_base+_open+_rem2)))))';
//   }

//   // Fix template
//   tmpl.getRange(2, 71)
//       .setFormula(makeFormula(2))
//       .setNumberFormat('dd/MM/yyyy HH:mm:ss');
//   Logger.log('Template BS2 fixed → now refs BR$13');

//   // Fix all data rows
//   let fixed = 0;
//   for (let row = 15; row <= lastRow; row++) {
//     if (bpms.getRange(row, 1).getValue() === '') continue;
//     bpms.getRange(row, 71)
//         .setFormula(makeFormula(row))
//         .setNumberFormat('dd/MM/yyyy HH:mm:ss');
//     fixed++;
//   }

//   SpreadsheetApp.flush();

//   // Verify row 15
//   const result = bpms.getRange(15, 71).getValue();
//   Logger.log('BS15 after fix: [' + result + ']');

//   ss.toast('Step 15 fixed in ' + fixed + ' rows! BS15 = ' + result, 'BPMS', 5);
//   Logger.log('fixStep15TATRef: done, ' + fixed + ' rows fixed');
// }