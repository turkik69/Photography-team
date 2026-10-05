from pathlib import Path
import re

p=Path('index.html')
s=p.read_text(encoding='utf-8')

# 1) Make long portal pages scroll completely on desktop and mobile.
marker='/* 2026-10-05: full-page scrolling fix for long reports/tasks on iOS and mobile */'
css='''
/* 2026-10-05: definitive full-page scrolling fix for desktop + mobile */
html,body,#app{height:auto!important;min-height:100%!important;max-height:none!important;overflow-y:auto!important}
body{overflow-x:hidden!important}
#app:has(.portal-layout){height:auto!important;min-height:100vh!important;max-height:none!important;overflow:visible!important}
.portal-layout{height:auto!important;min-height:calc(100vh - 74px)!important;max-height:none!important;overflow:visible!important;align-items:start!important}
.portal-layout .portal-main,.portal-main{height:auto!important;min-height:calc(100vh - 74px)!important;max-height:none!important;overflow:visible!important;padding-bottom:90px!important}
.portal-layout .portal-main .container,.portal-main>.container,.portal-main .container{height:auto!important;min-height:0!important;max-height:none!important;overflow:visible!important;padding-bottom:40px!important}
#reportBody,.card-grid{height:auto!important;max-height:none!important;overflow:visible!important}
@media(max-width:900px){
  .portal-main{padding-bottom:calc(130px + env(safe-area-inset-bottom))!important}
}
'''
if 'definitive full-page scrolling fix for desktop + mobile' not in s:
    pos=s.find(marker)
    if pos<0: raise SystemExit('scroll marker not found')
    s=s[:pos]+css+'\n'+s[pos:]

# 2) Replace Excel styling helper with true Arabic RTL and dynamic row heights.
new_style=r'''function styleExcelWorksheet(ws,XLSX,headerRowIndex){
  if(!ws||!ws['!ref']) return;
  ws['!views']=[{RTL:true,rightToLeft:true}];
  const range=XLSX.utils.decode_range(ws['!ref']);
  const thin={style:'thin',color:{rgb:'DCE8E1'}};
  const baseBorder={top:thin,bottom:thin,left:thin,right:thin};
  const widths=(ws['!cols']||[]).map(x=>Math.max(6,Number(x&&x.wch)||14));
  const rowHeights=[];
  for(let r=range.s.r;r<=range.e.r;r++){
    let maxLines=1;
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
        cell.s.alignment={horizontal:'right',vertical:'center',wrapText:true,readingOrder:2};
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
      const text=String(cell.v==null?'':cell.v);
      const explicitLines=text.split(/\r?\n/).length;
      const colWidth=widths[c]||14;
      const wrappedLines=Math.max(1,Math.ceil(text.length/Math.max(8,colWidth*1.35)));
      maxLines=Math.max(maxLines,explicitLines,wrappedLines);
    }
    let hpt=24;
    if(r===0) hpt=34;
    else if(r===headerRowIndex) hpt=30;
    else hpt=Math.min(110,Math.max(24,18*maxLines+6));
    rowHeights[r]={hpt};
  }
  ws['!rows']=rowHeights;
  if(range.e.r>headerRowIndex) ws['!autofilter']={ref:XLSX.utils.encode_range({s:{r:headerRowIndex,c:range.s.c},e:{r:range.e.r,c:range.e.c}})};
  ws['!freeze']={xSplit:0,ySplit:headerRowIndex+1,topLeftCell:`A${headerRowIndex+2}`,activePane:'bottomRight',state:'frozen'};
}'''
s,n=re.subn(r'function styleExcelWorksheet\(ws,XLSX,headerRowIndex\)\{.*?\n\}',lambda m:new_style,s,count=1,flags=re.S)
if n!=1: raise SystemExit(f'styleExcelWorksheet replacement count={n}')

# 3) Merge report titles across all columns in normal and subject sheets.
s=s.replace("const addSheet=(name,headers,rows,widths)=>{const aoa=[['تقرير حالة موضوع'],['الموضوع',title],['الفترة',start,'إلى',end],[],headers,...rows];const ws=XLSX.utils.aoa_to_sheet(aoa);ws['!cols']=widths.map(w=>({wch:w}));styleExcelWorksheet(ws,XLSX,4);XLSX.utils.book_append_sheet(wb,ws,name.slice(0,31));};",
"const addSheet=(name,headers,rows,widths)=>{const aoa=[['تقرير حالة موضوع'],['الموضوع',title],['الفترة',start,'إلى',end],[],headers,...rows];const ws=XLSX.utils.aoa_to_sheet(aoa);ws['!cols']=widths.map(w=>({wch:w}));ws['!merges']=[{s:{r:0,c:0},e:{r:0,c:Math.max(0,headers.length-1)}}];styleExcelWorksheet(ws,XLSX,4);XLSX.utils.book_append_sheet(wb,ws,name.slice(0,31));};")
s=s.replace("const add=(name,headers,rows,widths)=>{const aoa=[[`تقرير فريق التصوير - ${name}`],[`الفترة: ${data.period.label}`],[],headers,...rows];const ws=XLSX.utils.aoa_to_sheet(aoa);ws['!cols']=widths.map(w=>({wch:w}));styleExcelWorksheet(ws,XLSX,3);XLSX.utils.book_append_sheet(wb,ws,name.slice(0,28));};",
"const add=(name,headers,rows,widths)=>{const aoa=[[`تقرير فريق التصوير - ${name}`],[`الفترة: ${data.period.label}`],[],headers,...rows];const ws=XLSX.utils.aoa_to_sheet(aoa);ws['!cols']=widths.map(w=>({wch:w}));ws['!merges']=[{s:{r:0,c:0},e:{r:0,c:Math.max(0,headers.length-1)}},{s:{r:1,c:0},e:{r:1,c:Math.max(0,headers.length-1)}}];styleExcelWorksheet(ws,XLSX,3);XLSX.utils.book_append_sheet(wb,ws,name.slice(0,28));};")

# 4) Set workbook view RTL too.
s=s.replace("const XLSX=await ensureXLSXLoaded(),data=reportPeriodData(),wb=XLSX.utils.book_new(),selected=selectedReportContentType();",
"const XLSX=await ensureXLSXLoaded(),data=reportPeriodData(),wb=XLSX.utils.book_new(),selected=selectedReportContentType(); wb.Workbook=wb.Workbook||{}; wb.Workbook.Views=[{RTL:true}];")

p.write_text(s,encoding='utf-8')
