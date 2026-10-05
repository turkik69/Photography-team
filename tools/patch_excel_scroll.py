from pathlib import Path
import re

p=Path('index.html')
s=p.read_text(encoding='utf-8')

# 1) Mobile/portal scrolling: let the document own vertical scrolling and reserve space above bottom nav.
css = r'''
/* 2026-10-05: full-page scrolling fix for long reports/tasks on iOS and mobile */
@media (max-width: 900px){
  html,body,#app{height:auto!important;min-height:100%!important;min-height:100dvh!important;max-height:none!important;overflow-x:hidden!important;overflow-y:auto!important;-webkit-overflow-scrolling:touch}
  .portal-layout{height:auto!important;min-height:100dvh!important;max-height:none!important;overflow:visible!important;align-items:stretch}
  .portal-main{height:auto!important;min-height:calc(100dvh - 70px)!important;max-height:none!important;overflow:visible!important;padding-bottom:calc(118px + env(safe-area-inset-bottom))!important}
  .portal-main>.container,.portal-main .container{height:auto!important;max-height:none!important;overflow:visible!important;padding-bottom:24px!important}
  #reportBody,.card-grid{height:auto!important;max-height:none!important;overflow:visible!important}
}
'''
if 'full-page scrolling fix for long reports/tasks' not in s:
    s=s.replace('</style>', css+'\n</style>', 1)

# 2) Use style-capable SheetJS-compatible build so .s cell styles persist in .xlsx files.
s=re.sub(r'https://cdn\.jsdelivr\.net/npm/xlsx(?:@[^/]+)?/dist/xlsx\.full\.min\.js',
         'https://cdn.jsdelivr.net/npm/xlsx-js-style@1.2.0/dist/xlsx.bundle.js', s)
s=re.sub(r'https://unpkg\.com/xlsx(?:@[^/]+)?/dist/xlsx\.full\.min\.js',
         'https://cdn.jsdelivr.net/npm/xlsx-js-style@1.2.0/dist/xlsx.bundle.js', s)

# 3) Reusable Excel theme matching the platform/report colors and RTL direction.
helper = r'''function styleExcelWorksheet(ws,XLSX,headerRowIndex){
  if(!ws||!ws['!ref']) return;
  ws['!views']=[{rightToLeft:true}];
  const range=XLSX.utils.decode_range(ws['!ref']);
  const thin={style:'thin',color:{rgb:'DCE8E1'}};
  const baseBorder={top:thin,bottom:thin,left:thin,right:thin};
  for(let r=range.s.r;r<=range.e.r;r++){
    for(let c=range.s.c;c<=range.e.c;c++){
      const addr=XLSX.utils.encode_cell({r,c});
      const cell=ws[addr];
      if(!cell) continue;
      cell.s={
        font:{name:'Arial',sz:11,color:{rgb:'304A41'}},
        alignment:{horizontal:'right',vertical:'center',wrapText:true,readingOrder:2},
        border:baseBorder
      };
      if(r===0){
        cell.s.font={name:'Arial',sz:15,bold:true,color:{rgb:'FFFFFF'}};
        cell.s.fill={patternType:'solid',fgColor:{rgb:'174A3B'}};
        cell.s.alignment={horizontal:'right',vertical:'center',readingOrder:2};
      }else if(r===1 || (headerRowIndex===4 && r===2)){
        cell.s.font={name:'Arial',sz:11,bold:true,color:{rgb:'174A3B'}};
        cell.s.fill={patternType:'solid',fgColor:{rgb:'EDF5F0'}};
      }else if(r===headerRowIndex){
        cell.s.font={name:'Arial',sz:11,bold:true,color:{rgb:'FFFFFF'}};
        cell.s.fill={patternType:'solid',fgColor:{rgb:'2F6F24'}};
        cell.s.alignment={horizontal:'right',vertical:'center',wrapText:true,readingOrder:2};
      }else if(r>headerRowIndex && ((r-headerRowIndex)%2===0)){
        cell.s.fill={patternType:'solid',fgColor:{rgb:'F7FAF8'}};
      }
    }
  }
  ws['!rows']=Array.from({length:range.e.r+1},(_,i)=>({hpt:i===0?28:(i===headerRowIndex?25:22)}));
  if(range.e.r>headerRowIndex) ws['!autofilter']={ref:XLSX.utils.encode_range({s:{r:headerRowIndex,c:range.s.c},e:{r:range.e.r,c:range.e.c}})};
  ws['!freeze']={xSplit:0,ySplit:headerRowIndex+1,topLeftCell:`A${headerRowIndex+2}`,activePane:'bottomLeft',state:'frozen'};
}

'''
if 'function styleExcelWorksheet(ws,XLSX,headerRowIndex)' not in s:
    marker='function appendSubjectSheetToWorkbook(XLSX,wb){'
    if marker not in s: raise SystemExit('appendSubjectSheetToWorkbook marker missing')
    s=s.replace(marker,helper+marker,1)

# Subject report workbook sheets: header row is row 5 (zero-based index 4).
old="const addSheet=(name,headers,rows,widths)=>{const aoa=[['تقرير حالة موضوع'],['الموضوع',title],['الفترة',start,'إلى',end],[],headers,...rows];const ws=XLSX.utils.aoa_to_sheet(aoa);ws['!cols']=widths.map(w=>({wch:w}));ws['!views']=[{rightToLeft:true}];XLSX.utils.book_append_sheet(wb,ws,name.slice(0,31));};"
new="const addSheet=(name,headers,rows,widths)=>{const aoa=[['تقرير حالة موضوع'],['الموضوع',title],['الفترة',start,'إلى',end],[],headers,...rows];const ws=XLSX.utils.aoa_to_sheet(aoa);ws['!cols']=widths.map(w=>({wch:w}));styleExcelWorksheet(ws,XLSX,4);XLSX.utils.book_append_sheet(wb,ws,name.slice(0,31));};"
if old in s: s=s.replace(old,new,1)
elif 'styleExcelWorksheet(ws,XLSX,4)' not in s: raise SystemExit('subject addSheet pattern missing')

# Standard report workbook sheets: header row is row 4 (zero-based index 3).
old2="const add=(name,headers,rows,widths)=>{const aoa=[[`تقرير فريق التصوير - ${name}`],[`الفترة: ${data.period.label}`],[],headers,...rows];const ws=XLSX.utils.aoa_to_sheet(aoa);ws['!cols']=widths.map(w=>({wch:w}));ws['!views']=[{rightToLeft:true}];XLSX.utils.book_append_sheet(wb,ws,name.slice(0,28));};"
new2="const add=(name,headers,rows,widths)=>{const aoa=[[`تقرير فريق التصوير - ${name}`],[`الفترة: ${data.period.label}`],[],headers,...rows];const ws=XLSX.utils.aoa_to_sheet(aoa);ws['!cols']=widths.map(w=>({wch:w}));styleExcelWorksheet(ws,XLSX,3);XLSX.utils.book_append_sheet(wb,ws,name.slice(0,28));};"
if old2 in s: s=s.replace(old2,new2,1)
elif 'styleExcelWorksheet(ws,XLSX,3)' not in s: raise SystemExit('standard add pattern missing')

p.write_text(s,encoding='utf-8')
