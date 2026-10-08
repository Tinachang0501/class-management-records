(() => {
 const COLS=['學生座號','姓名','日期','面談方式','主題','輔導內容'];
 const TOPICS=['家庭','性別','情緒','生涯輔導','常規問題','人際關係','課業學習','其它'];
 const METHODS=['面談學生','面談家長','電聯家長','家訪'];
 const STORAGE='class_management_records_data';
 const esc=recordEscape;
 const dateOnly=value=>{const m=String(value||'').match(/^(\d{4})[-/](\d{1,2})[-/](\d{1,2})/);if(!m)throw Error('日期須為年月日');return `${m[1]}-${m[2].padStart(2,'0')}-${m[3].padStart(2,'0')}`;};
 const method=(value,topic,summary)=>{
  if(topic==='常規問題'||summary.includes('【違規情節與背景】'))return '面談學生、電聯家長';
  const s=String(value||'').replaceAll('面談(學生)','面談學生').replaceAll('面談(家長)','面談家長');
  const values=s.split(/[,、]/).map(x=>x.trim());return METHODS.filter(x=>values.includes(x)).join('、')||'面談學生';
 };
 function clean(r){
  const date=dateOnly(r.date??r.日期),seat=Number(r.seat??r.學生座號),name=String(r.name??r.姓名??'').trim(),summary=String(r.summary??r.輔導內容??'').trim();
  if(!Number.isInteger(seat)||seat<1||seat>99||!name||!summary)throw Error('紀錄缺少有效座號、姓名或輔導內容');
  let topic=String(r.topic??r.主題??'其它');if(!TOPICS.includes(topic))topic='其它';
  return {id:`record-${date}-${seat}`,date,time:'',datetime_display:date.replaceAll('-','/'),seat:String(seat).padStart(2,'0'),name,summary,topic,target:method(r.target??r.面談方式,topic,summary)};
 }
 const all=()=>Object.values(recordsData).flat();
 function organize(list){
  const next={};list.sort((a,b)=>a.date.localeCompare(b.date)||Number(a.seat)-Number(b.seat));
  for(const r of list){const m=r.date.slice(0,4)+'年'+r.date.slice(5,7)+'月';(next[m]||=[]).push({...r,no:String(next[m].length+1)});}return next;
 }
 const panel=document.createElement('section');panel.id='record-import-panel';panel.hidden=true;panel.className='no-print';
 panel.innerHTML=`<button id="record-close" style="float:right">關閉</button><h2 style="font-size:22px;font-weight:bold">同步新增班級紀錄</h2><p style="margin:12px 0">選取整理好的 Excel 或 JSON 檔。同一日期及座號會更新原筆，其餘新增，不會重複累加。</p><label style="display:block;padding:12px;border:1px solid #cbd5e1">選取紀錄檔 <input id="record-file" type="file" accept=".json,.xlsx"></label><p id="record-import-status" role="status" style="margin:12px 0">尚未選取檔案。</p><button id="record-backup" class="record-button">下載紀錄備份</button><button id="record-export" class="record-button">下載六欄 Excel</button><p>紀錄保存在本瀏覽器。換電腦或瀏覽器時，請匯入備份；清除網站資料前請先備份。</p><p style="margin-top:12px;color:#64748b">Google 持續自動同步尚待 Apps Script 授權與介面設定。</p>`;
 document.body.append(panel);
 const status=s=>document.getElementById('record-import-status').textContent=s;
 document.getElementById('record-close').onclick=()=>panel.hidden=true;
 openGoogleSheetModal=()=>{panel.hidden=false;};
 document.querySelectorAll('button').forEach(b=>{if(b.textContent.includes('連動 Google 試算表'))b.textContent='同步新增紀錄';});
 const counter=document.createElement('p');counter.id='record-count';counter.className='no-print';counter.style.cssText='max-width:1000px;margin:16px auto;color:#334155';document.getElementById('pages-container').before(counter);
 const originalRender=renderPages;
 renderPages=function(){
  recordsData=organize(all().map(clean));
  students=[...new Map(all().map(r=>[r.seat,{'座號':r.seat,'姓名':r.name}])).values()];
  for(const id of ['view-month-filter','month-select']){const el=document.getElementById(id);if(el){const previous=el.value;el.innerHTML='<option value="ALL">全部月份</option>'+Object.keys(recordsData).map(m=>`<option value="${m}">${m}</option>`).join('');el.value=Object.keys(recordsData).includes(previous)?previous:'ALL';}}
  try{localStorage.setItem(STORAGE,JSON.stringify(recordsData));}catch(e){status('瀏覽器無法儲存，請立即下載備份。');}
  counter.textContent=`共 ${all().length} 筆紀錄，${students.length} 位學生。日期、座號排序；資料存於本瀏覽器。`;
  originalRender();
 };
 formatTargetMethodLine=value=>esc(String(value));
 polishCounselingText=raw=>String(raw||'').trim();
 setSelectedEditTargets=value=>{const s=String(value);for(const [id,token] of [['stu','面談學生'],['parent','面談家長'],['phone','電聯家長'],['home','家訪']]){const el=document.getElementById('edit-chk-'+id);if(el)el.checked=s.includes(token);}const other=document.getElementById('edit-target-other');if(other)other.value='';};
 for(const sel of document.querySelectorAll('select'))if(/topic/.test(sel.id)){const old=sel.value;sel.innerHTML=TOPICS.map(t=>`<option>${t}</option>`).join('');sel.value=TOPICS.includes(old)?old:'課業學習';}
 function download(blob,name){const url=URL.createObjectURL(blob),a=document.createElement('a');a.href=url;a.download=name;a.click();setTimeout(()=>URL.revokeObjectURL(url),2000);}
 const backup=()=>download(new Blob([JSON.stringify({version:2,records:all()},null,2)],{type:'application/json'}),'班級紀錄備份.json');
 document.getElementById('record-backup').onclick=backup;
 exportToExcel=function(){
  if(!all().length)return alert('請先匯入紀錄');if(!window.XLSX)return alert('Excel元件尚未載入');
  const rows=all().map(r=>[Number(r.seat),r.name,(Date.parse(r.date+'T00:00:00Z')-Date.UTC(1899,11,30))/86400000,r.target,r.topic,r.summary]);
  const ws=XLSX.utils.aoa_to_sheet([COLS,...rows]);ws['!cols']=[{wch:9},{wch:12},{wch:14},{wch:26},{wch:14},{wch:100}];ws['!autofilter']={ref:ws['!ref']};
  for(let i=2;i<=rows.length+1;i++){ws['A'+i].z='00';ws['C'+i].z='yyyy/mm/dd';}
  const wb=XLSX.utils.book_new();XLSX.utils.book_append_sheet(wb,ws,'輔導紀錄');XLSX.writeFile(wb,'全班B卡輔導與家長聯絡紀錄_2026年7月起.xlsx');
 };
 document.getElementById('record-export').onclick=exportToExcel;
 exportToWord=function(){
  const content=document.getElementById('pages-container').innerHTML.replace(/<button[\s\S]*?<\/button>/g,'');
  download(new Blob(['\ufeff<html><meta charset="utf-8"><style>table{border-collapse:collapse;width:100%}td,th{border:1px solid black;padding:5px} .print-page{page-break-after:always}.no-print{display:none}</style><body>'+content+'</body></html>'],{type:'application/msword'}),'班級經營紀錄本.doc');
 };
 document.getElementById('record-file').onchange=async event=>{
  try{
   const file=event.target.files[0];if(!file)return;let list;
   if(/\.xlsx$/i.test(file.name)){
    if(!window.XLSX)throw Error('Excel元件尚未載入');const wb=XLSX.read(await file.arrayBuffer(),{type:'array',cellDates:false});const ws=wb.Sheets['輔導紀錄'];if(!ws)throw Error('找不到輔導紀錄分頁');
    list=XLSX.utils.sheet_to_json(ws,{raw:false});
   }else{const input=JSON.parse(await file.text());list=Array.isArray(input)?input:input.records;if(!Array.isArray(list))throw Error('請選擇紀錄匯入檔或備份檔');}
   const incoming=list.map(clean),seen=new Set();for(const r of incoming){if(seen.has(r.id))throw Error('匯入檔有同一學生同日多筆，請先合併');seen.add(r.id);}
   const previous=JSON.stringify(recordsData),map=new Map(all().map(r=>[clean(r).id,clean(r)]));let added=0,updated=0;
   for(const r of incoming){map.has(r.id)?updated++:added++;map.set(r.id,r);}
   const next=organize([...map.values()]);
   localStorage.setItem(STORAGE+'_before_import',previous);localStorage.setItem(STORAGE,JSON.stringify(next));recordsData=next;renderPages();
   status(`匯入完成：新增 ${added} 筆，更新 ${updated} 筆；目前共 ${all().length} 筆。已儲存於本瀏覽器。`);
  }catch(error){status('未匯入：'+error.message);}
  finally{event.target.value='';}
 };
 try{const saved=localStorage.getItem(STORAGE);if(saved){const value=JSON.parse(saved);if(!value||typeof value!=='object')throw Error('資料格式錯誤');recordsData=organize(Object.values(value).flat().map(clean));}}catch(error){status('原有本機紀錄無法讀取，請先匯出備份後再處理。');}
 renderPages();
})();
