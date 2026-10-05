from pathlib import Path
import re

p=Path('index.html')
s=p.read_text(encoding='utf-8')

new_subject = r'''function detailedSubjectReportHtml(){
  if(!S.filters.subjectReportEnabled) return '';
  const {start,end}=subjectReportBounds();
  const items=subjectReportItems();
  const requestItems=items.filter(i=>i.kind==='request');
  const visitItems=items.filter(i=>i.kind==='visit');
  const meetingItems=items.filter(i=>i.kind==='meeting');
  const taskItems=items.filter(i=>i.kind==='task');
  const title=(S.filters.subjectQuery||'').trim()||'جميع الموضوعات';
  let sections='';
  if(requestItems.length){
    const rows=requestItems.map((item,i)=>{ const r=item.data; return [i+1,r.requestNumber||'',r.event?.name||item.title||'',r.event?.date||item.date||'',r.requester?.directorate||'',r.event?.location||'',{text:statusLabel(r.status),statusClass:reportStatusClass(statusLabel(r.status))},priorityLabel(r.priority),empNamesFromIds(r.internal?.assignedPhotographers).join('، '),empNamesFromIds(r.internal?.assignedEditors).join('، '),productionStageLabel(r.production?.stage),r.internal?.teamNotes||r.production?.notes||'']; });
    sections+=reportSheetHtml('طلبات التصوير المرتبطة بالموضوع',`${start} إلى ${end}`,['م','رقم الطلب','الموضوع','التاريخ','الجهة','الموقع','الحالة','الأولوية','التصوير','المونتاج','مرحلة الإنتاج','الملاحظات'],rows);
  }
  if(visitItems.length){
    const rows=[]; let n=1;
    visitItems.forEach(item=>{ const v=item.data; const sites=(v.sites&&v.sites.length)?v.sites:[{}]; sites.forEach(site=>rows.push([n++,v.date||item.date||'',site.projectName||item.title||'',site.location||'',site.taskType||'',site.startTime||v.startTime||'',site.endTime||v.endTime||'',(v.conductedBy||[]).join('، '),site.description||''])); });
    sections+=reportSheetHtml('الزيارات المرتبطة بالموضوع',`${start} إلى ${end}`,['م','تاريخ الزيارة','اسم المشروع / الموقع','الموقع','نوع المهمة','وقت البداية','وقت الانتهاء','من قام بالزيارة','الوصف / الملاحظات'],rows);
  }
  if(meetingItems.length){
    const rows=meetingItems.map((item,i)=>{ const m=item.data; return [i+1,m.title||item.title||'',m.date||item.date||'',m.startTime||'',m.endTime||'',m.location||'',empNamesFromIds(m.attendees).join('، '),m.agenda||'',m.notes||'']; });
    sections+=reportSheetHtml('الاجتماعات المرتبطة بالموضوع',`${start} إلى ${end}`,['م','عنوان الاجتماع','التاريخ','وقت البداية','وقت الانتهاء','المكان','الحضور','جدول الأعمال','الملاحظات / المحضر'],rows);
  }
  if(taskItems.length){
    const rows=taskItems.map((item,i)=>{ const t=item.data; return [i+1,t.title||item.title||'',t.startDate||item.date||'',t.endDate||'',(t.responsible||[]).join('، '),{text:taskStatusLabel(t.status),statusClass:reportStatusClass(taskStatusLabel(t.status))},t.notes||'']; });
    sections+=reportSheetHtml('جدول الأعمال المرتبط بالموضوع',`${start} إلى ${end}`,['م','الموضوع / المشروع','تاريخ البدء','تاريخ الانتهاء','المسؤول','الحالة','التحديات / الملاحظات'],rows);
  }
  return `<section class="report-sheet" style="border-color:var(--gold-500)"><div class="report-sheet-head"><h3>${icon('search')} تقرير حالة موضوع: ${esc(title)}</h3><span>${esc(start)} إلى ${esc(end)}</span></div><div class="section-body"><div class="report-summary-strip" style="margin-bottom:0"><div class="report-summary-item"><strong>${items.length}</strong><span>إجمالي السجلات المرتبطة</span></div><div class="report-summary-item"><strong>${requestItems.length}</strong><span>طلبات تصوير</span></div><div class="report-summary-item"><strong>${visitItems.length}</strong><span>زيارات</span></div><div class="report-summary-item"><strong>${meetingItems.length+taskItems.length}</strong><span>اجتماعات وأعمال</span></div></div></div></section>${sections||'<div class="report-sheet"><div class="section-body"><div class="empty-state"><h3>لا توجد بيانات مطابقة للموضوع خلال الفترة المحددة</h3></div></div></div>'}`;
}
'''

s,n=re.subn(r'function detailedSubjectReportHtml\(\)\{.*?\n\}\n\n\nfunction selectedReportContentType\(\)',new_subject+'\n\nfunction selectedReportContentType()',s,flags=re.S)
if n!=1:
    raise SystemExit(f'detailedSubjectReportHtml replacement count={n}')

new_excel = r'''function appendSubjectSheetToWorkbook(XLSX,wb){
  if(!S.filters.subjectReportEnabled) return;
  const {start,end}=subjectReportBounds();
  const items=subjectReportItems();
  const title=(S.filters.subjectQuery||'').trim()||'جميع الموضوعات';
  const requestItems=items.filter(i=>i.kind==='request'), visitItems=items.filter(i=>i.kind==='visit'), meetingItems=items.filter(i=>i.kind==='meeting'), taskItems=items.filter(i=>i.kind==='task');
  const addSheet=(name,headers,rows,widths)=>{ const aoa=[['تقرير حالة موضوع'],['الموضوع',title],['الفترة',start,'إلى',end],[],headers,...rows]; const ws=XLSX.utils.aoa_to_sheet(aoa); ws['!cols']=widths.map(w=>({wch:w})); ws['!views']=[{rightToLeft:true}]; XLSX.utils.book_append_sheet(wb,ws,name.slice(0,31)); };
  addSheet('حالة الموضوع',['البيان','القيمة'],[['الموضوع',title],['من تاريخ',start],['إلى تاريخ',end],['إجمالي السجلات',items.length],['طلبات التصوير',requestItems.length],['الزيارات',visitItems.length],['الاجتماعات',meetingItems.length],['جدول الأعمال',taskItems.length]],[24,42]);
  if(requestItems.length) addSheet('الموضوع-طلبات',['م','رقم الطلب','الموضوع','التاريخ','الجهة','الموقع','الحالة','الأولوية','التصوير','المونتاج','مرحلة الإنتاج','الملاحظات'],requestItems.map((item,i)=>{const r=item.data;return [i+1,r.requestNumber||'',r.event?.name||item.title||'',r.event?.date||item.date||'',r.requester?.directorate||'',r.event?.location||'',statusLabel(r.status),priorityLabel(r.priority),empNamesFromIds(r.internal?.assignedPhotographers).join('، '),empNamesFromIds(r.internal?.assignedEditors).join('، '),productionStageLabel(r.production?.stage),r.internal?.teamNotes||r.production?.notes||''];}),[6,15,34,14,24,24,16,14,24,24,18,45]);
  if(visitItems.length){const rows=[];let n=1;visitItems.forEach(item=>{const v=item.data;(v.sites&&v.sites.length?v.sites:[{}]).forEach(site=>rows.push([n++,v.date||item.date||'',site.projectName||item.title||'',site.location||'',site.taskType||'',site.startTime||v.startTime||'',site.endTime||v.endTime||'',(v.conductedBy||[]).join('، '),site.description||'']));});addSheet('الموضوع-زيارات',['م','تاريخ الزيارة','اسم المشروع / الموقع','الموقع','نوع المهمة','وقت البداية','وقت الانتهاء','من قام بالزيارة','الوصف / الملاحظات'],rows,[6,14,34,22,20,12,12,28,48]);}
  if(meetingItems.length) addSheet('الموضوع-اجتماعات',['م','عنوان الاجتماع','التاريخ','وقت البداية','وقت الانتهاء','المكان','الحضور','جدول الأعمال','الملاحظات / المحضر'],meetingItems.map((item,i)=>{const m=item.data;return [i+1,m.title||item.title||'',m.date||item.date||'',m.startTime||'',m.endTime||'',m.location||'',empNamesFromIds(m.attendees).join('، '),m.agenda||'',m.notes||''];}),[6,34,14,12,12,24,30,45,45]);
  if(taskItems.length) addSheet('الموضوع-أعمال',['م','الموضوع / المشروع','تاريخ البدء','تاريخ الانتهاء','المسؤول','الحالة','التحديات / الملاحظات'],taskItems.map((item,i)=>{const t=item.data;return [i+1,t.title||item.title||'',t.startDate||item.date||'',t.endDate||'',(t.responsible||[]).join('، '),taskStatusLabel(t.status),t.notes||''];}),[6,38,14,14,28,16,50]);
}
'''

s,n=re.subn(r'function appendSubjectSheetToWorkbook\(XLSX,wb\)\{.*?\n\}\nasync function exportReportToExcel\(\)',new_excel+'\nasync function exportReportToExcel()',s,flags=re.S)
if n!=1:
    raise SystemExit(f'appendSubjectSheetToWorkbook replacement count={n}')

p.write_text(s,encoding='utf-8')
