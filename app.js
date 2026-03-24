const DB_NAME='zeitapp_db_v8',DB_VERSION=1,STORE_ENTRIES='entries',STORE_SETTINGS='settings',STORE_META='meta';
let db,deferredPrompt=null;
const $=id=>document.getElementById(id);
const installBtn=$('installBtn'),form=$('entryForm'),resetBtn=$('resetBtn'),entriesTable=$('entriesTable'),weekPicker=$('weekPicker'),stats=$('stats'),monthStats=$('monthStats'),pdfBtn=$('pdfBtn'),sheetBtn=$('sheetBtn'),csvBtn=$('csvBtn'),xlsxBtn=$('xlsxBtn'),saveSettingsBtn=$('saveSettingsBtn'),sheetWebhookInput=$('sheetWebhook'),backupBtn=$('backupBtn'),restoreBtn=$('restoreBtn'),restoreFile=$('restoreFile'),siteSelect=$('siteSelect'),vehicleSelect=$('vehicleSelect'),addSiteBtn=$('addSiteBtn'),addVehicleBtn=$('addVehicleBtn'),weeklyTargetInput=$('weeklyTarget'),workDaysPerWeekInput=$('workDaysPerWeek'),monthPicker=$('monthPicker'),refreshMonthBtn=$('refreshMonthBtn'),closeWeekBtn=$('closeWeekBtn'),reopenWeekBtn=$('reopenWeekBtn'),warningsBox=$('warningsBox');

$('date').value=todayISO(); weekPicker.value=currentWeekValue(); monthPicker.value=todayISO().slice(0,7);
boot();

async function boot(){db=await openDb();await seedDefaults();await loadSettings();await fillSelectors();await render();await renderMonth();registerSW();setupInstall();applyEntryTypeMode()}
function openDb(){return new Promise((resolve,reject)=>{const req=indexedDB.open(DB_NAME,DB_VERSION);req.onupgradeneeded=e=>{const db=e.target.result;if(!db.objectStoreNames.contains(STORE_ENTRIES)){const s=db.createObjectStore(STORE_ENTRIES,{keyPath:'id'});s.createIndex('week','week');s.createIndex('month','month')} if(!db.objectStoreNames.contains(STORE_SETTINGS))db.createObjectStore(STORE_SETTINGS,{keyPath:'key'}); if(!db.objectStoreNames.contains(STORE_META))db.createObjectStore(STORE_META,{keyPath:'key'})}; req.onsuccess=()=>resolve(req.result); req.onerror=()=>reject(req.error)})}
function tx(n,m='readonly'){return db.transaction(n,m).objectStore(n)}
function getOne(n,k){return new Promise((res,rej)=>{const r=tx(n).get(k);r.onsuccess=()=>res(r.result);r.onerror=()=>rej(r.error)})}
function putOne(n,v){return new Promise((res,rej)=>{const r=tx(n,'readwrite').put(v);r.onsuccess=()=>res(true);r.onerror=()=>rej(r.error)})}
function deleteOne(n,k){return new Promise((res,rej)=>{const r=tx(n,'readwrite').delete(k);r.onsuccess=()=>res(true);r.onerror=()=>rej(r.error)})}
function getAll(n){return new Promise((res,rej)=>{const r=tx(n).getAll();r.onsuccess=()=>res(r.result||[]);r.onerror=()=>rej(r.error)})}

async function seedDefaults(){
 if(!await getOne(STORE_META,'sites')) await putOne(STORE_META,{key:'sites',values:['Bitte wählen','Unbekannt / wechselnd']});
 if(!await getOne(STORE_META,'vehicles')) await putOne(STORE_META,{key:'vehicles',values:['Bitte wählen','Kein Fahrzeug']});
 const defaults={employee:'',weeklyTarget:40,workDaysPerWeek:5,sheetWebhook:''};
 for(const [k,v] of Object.entries(defaults)) if(!await getOne(STORE_SETTINGS,k)) await putOne(STORE_SETTINGS,{key:k,value:v});
}
async function loadSettings(){
 sheetWebhookInput.value=(await getOne(STORE_SETTINGS,'sheetWebhook'))?.value||'';
 $('employee').value=(await getOne(STORE_SETTINGS,'employee'))?.value||'';
 weeklyTargetInput.value=(await getOne(STORE_SETTINGS,'weeklyTarget'))?.value||40;
 workDaysPerWeekInput.value=(await getOne(STORE_SETTINGS,'workDaysPerWeek'))?.value||5;

}
async function saveSettings(){
 await putOne(STORE_SETTINGS,{key:'sheetWebhook',value:sheetWebhookInput.value.trim()});
 await putOne(STORE_SETTINGS,{key:'employee',value:$('employee').value.trim()});
 await putOne(STORE_SETTINGS,{key:'weeklyTarget',value:Number(weeklyTargetInput.value||40)});
 await putOne(STORE_SETTINGS,{key:'workDaysPerWeek',value:Number(workDaysPerWeekInput.value||5)});

 alert('Einstellungen gespeichert.');
 await render(); await renderMonth();
}
saveSettingsBtn.addEventListener('click',saveSettings);

async function fillSelectors(){
 const sites=(await getOne(STORE_META,'sites'))?.values||[], vehicles=(await getOne(STORE_META,'vehicles'))?.values||[];
 siteSelect.innerHTML=sites.map(v=>`<option value="${escapeHtml(v)}">${escapeHtml(v)}</option>`).join('');
 vehicleSelect.innerHTML=vehicles.map(v=>`<option value="${escapeHtml(v)}">${escapeHtml(v)}</option>`).join('');
}
addSiteBtn.addEventListener('click',async()=>{const v=$('newSite').value.trim(); if(!v) return; const c=await getOne(STORE_META,'sites'); if(!c.values.includes(v)){c.values.push(v); await putOne(STORE_META,c)} $('newSite').value=''; await fillSelectors(); siteSelect.value=v;});
addVehicleBtn.addEventListener('click',async()=>{const v=$('newVehicle').value.trim(); if(!v) return; const c=await getOne(STORE_META,'vehicles'); if(!c.values.includes(v)){c.values.push(v); await putOne(STORE_META,c)} $('newVehicle').value=''; await fillSelectors(); vehicleSelect.value=v;});

$('entryType').addEventListener('change',applyEntryTypeMode);
function applyEntryTypeMode(){
 const type=val('entryType'); const ids=['workStart','workEnd','driveMinutes','breakMinutes','kilometers','fuelLiters','fuelAmount','odometer','breakMode'];
 const disable=type!=='arbeit'; ids.forEach(id=>$(id).disabled=disable);
 $('dayFraction').disabled=(type==='arbeit');
 if(disable){$('workStart').value='08:00'; $('workEnd').value='16:00'; ['driveMinutes','breakMinutes','kilometers','fuelLiters','fuelAmount','odometer'].forEach(id=>$(id).value=0);}
}

form.addEventListener('submit',async e=>{
 e.preventDefault();
 const week=weekPicker.value||currentWeekValue();
 if(await isWeekClosed(week)){alert('Diese Woche ist abgeschlossen und kann nicht bearbeitet werden.'); return;}
 const warnings=validateEntry(); showWarnings(warnings); if(warnings.some(w=>w.blocking)){return}
 const type=val('entryType'), breakMode=val('breakMode'), manualBreak=num('breakMinutes');
 const autoBreak=type==='arbeit'?computeBreakMinutes(val('workStart'),val('workEnd'),breakMode,manualBreak):0;
 const existingId=$('editId').value||crypto.randomUUID(), existing=await getOne(STORE_ENTRIES,existingId);
 let receipt=await readAndCompressReceipt(); if(!receipt && existing?.receipt) receipt=existing.receipt;
 const entry={id:existingId,date:val('date'),week:weekFromDate(val('date')),month:val('date').slice(0,7),employee:val('employee'),
   entryType:type,dayFraction:Number(val('dayFraction')||1),site:siteSelect.value,vehicle:vehicleSelect.value,
   workStart:val('workStart'),workEnd:val('workEnd'),driveMinutes:type==='arbeit'?num('driveMinutes'):0,breakMinutes:autoBreak,breakMode,
   kilometers:type==='arbeit'?num('kilometers'):0,fuelLiters:type==='arbeit'?num('fuelLiters'):0,fuelAmount:type==='arbeit'?num('fuelAmount'):0,
   odometer:type==='arbeit'?num('odometer'):0,notes:val('notes'),receipt,exported:false,updatedAt:new Date().toISOString()};
 if(existing && existing.exported) entry.exported=false;
 await putOne(STORE_ENTRIES,entry); await putOne(STORE_SETTINGS,{key:'employee',value:entry.employee});
 form.reset(); $('date').value=todayISO(); $('employee').value=entry.employee; $('breakMode').value='autoMixed'; $('entryType').value='arbeit'; $('dayFraction').value='1'; $('editId').value=''; warningsBox.classList.add('hidden');
 weekPicker.value=entry.week; applyEntryTypeMode(); await render(); await renderMonth();
});
resetBtn.addEventListener('click',async()=>{form.reset(); $('date').value=todayISO(); $('employee').value=(await getOne(STORE_SETTINGS,'employee'))?.value||''; $('breakMode').value='autoMixed'; $('entryType').value='arbeit'; $('dayFraction').value='1'; $('editId').value=''; warningsBox.classList.add('hidden'); applyEntryTypeMode()});

weekPicker.addEventListener('change',render); monthPicker.addEventListener('change',renderMonth); refreshMonthBtn.addEventListener('click',renderMonth);
pdfBtn.addEventListener('click',printWeeklyReport); sheetBtn.addEventListener('click',exportToGoogleSheets); csvBtn.addEventListener('click',exportCsv); xlsxBtn.addEventListener('click',exportXlsx);
backupBtn.addEventListener('click',exportBackup); restoreBtn.addEventListener('click',()=>restoreFile.click()); restoreFile.addEventListener('change',importBackup);
const syncNowBtn=document.getElementById('syncNowBtn'); if(syncNowBtn) syncNowBtn.addEventListener('click', ()=>exportToGoogleSheets());
const syncReadBtn=document.getElementById('syncReadBtn'); if(syncReadBtn) syncReadBtn.addEventListener('click', ()=>syncFromGoogleSheets());
closeWeekBtn.addEventListener('click',()=>setWeekClosed(true)); reopenWeekBtn.addEventListener('click',()=>setWeekClosed(false));

function val(id){return $(id).value.trim()} function num(id){const v=parseFloat($(id).value); return Number.isFinite(v)?v:0}
async function readAndCompressReceipt(){
 const file=$('receipt').files[0]; if(!file) return null;
 return new Promise((resolve,reject)=>{
   const reader=new FileReader(); reader.onload=()=>{const img=new Image(); img.onload=()=>{const max=1280; let {width,height}=img;
     if(width>max||height>max){const ratio=Math.min(max/width,max/height); width=Math.round(width*ratio); height=Math.round(height*ratio)}
     const canvas=document.createElement('canvas'); canvas.width=width; canvas.height=height; const ctx=canvas.getContext('2d'); ctx.drawImage(img,0,0,width,height);
     resolve(canvas.toDataURL('image/jpeg',0.72));
   }; img.onerror=reject; img.src=reader.result;}; reader.onerror=reject; reader.readAsDataURL(file);
 });
}
function computeBreakMinutes(start,end,mode,manual){if(mode==='manual') return manual; const raw=Math.max(0,timeToMinutes(end)-timeToMinutes(start)), hours=raw/60; if(mode==='auto30') return hours>=6?30:0; if(mode==='auto45') return hours>=9?45:0; if(mode==='autoMixed'){if(hours>=9) return 45; if(hours>=6) return 30; return 0} return manual}
function timeToMinutes(v){const [h,m]=(v||'00:00').split(':').map(Number); return h*60+m}
function calcWorkMinutes(e){ if(['urlaub','krank','feiertag'].includes(e.entryType)) return 0; const raw=Math.max(0,timeToMinutes(e.workEnd)-timeToMinutes(e.workStart)); return Math.max(0,raw-(e.breakMinutes||0))}
function minutesToHours(mins){const sign=mins<0?'-':''; const abs=Math.abs(mins),h=Math.floor(abs/60),m=Math.round(abs%60); return `${sign}${h}h ${String(m).padStart(2,'0')}m`}
function todayISO(){return new Date().toISOString().slice(0,10)} function formatDate(iso){const [y,m,d]=iso.split('-'); return `${d}.${m}.${y}`} function currentWeekValue(){return weekFromDate(todayISO())}
function weekFromDate(isoDate){const date=new Date(isoDate+'T00:00:00'); const temp=new Date(date.getTime()); temp.setDate(temp.getDate()+3-((temp.getDay()+6)%7)); const week1=new Date(temp.getFullYear(),0,4); const week=1+Math.round(((temp-week1)/86400000-3+((week1.getDay()+6)%7))/7); return `${temp.getFullYear()}-W${String(week).padStart(2,'0')}`}
function escapeHtml(v){return String(v??'').replaceAll('&','&amp;').replaceAll('<','&lt;').replaceAll('>','&gt;').replaceAll('"','&quot;').replaceAll("'",'&#039;')}

async function isWeekClosed(week){return !!((await getOne(STORE_META,'closedWeeks'))?.values||[]).includes(week)}
async function setWeekClosed(closed){
 const week=weekPicker.value||currentWeekValue(); let obj=await getOne(STORE_META,'closedWeeks'); if(!obj) obj={key:'closedWeeks',values:[]};
 const has=obj.values.includes(week);
 if(closed && !has) obj.values.push(week);
 if(!closed && has) obj.values=obj.values.filter(v=>v!==week);
 await putOne(STORE_META,obj); alert(closed?'Woche abgeschlossen.':'Woche wieder geöffnet.'); await render();
}

function workingDayCreditMinutes(targetHours, workDays){return targetHours*60/workDays}
function absenceCreditMinutes(entries,targetHours,workDays){
 const perDay=workingDayCreditMinutes(targetHours,workDays); let credit=0;
 entries.forEach(e=>{ if(['urlaub','krank','feiertag'].includes(e.entryType)) credit += perDay*(Number(e.dayFraction||1)); });
 return credit;
}
async function getWeekEntries(){const all=await getAll(STORE_ENTRIES),week=weekPicker.value||currentWeekValue(); return all.filter(e=>e.week===week).sort((a,b)=>a.date.localeCompare(b.date))}
async function render(){
 const entries=await getWeekEntries(); entriesTable.innerHTML='';
 const closed=await isWeekClosed(weekPicker.value||currentWeekValue()); closeWeekBtn.disabled=closed; reopenWeekBtn.disabled=!closed;
 const target=Number(weeklyTargetInput.value||40), workDays=Number(workDaysPerWeekInput.value||5), soll=target*60;
 for(const e of entries){
   const typeClass=`type-${e.entryType||'arbeit'}`;
   const tr=document.createElement('tr');
   tr.innerHTML=`<td>${formatDate(e.date)}</td><td><span class="type-pill ${typeClass}">${escapeHtml(e.entryType||'arbeit')}</span></td><td>${escapeHtml(e.employee)}</td><td>${escapeHtml(e.site||'')}</td><td>${minutesToHours(calcWorkMinutes(e))}</td><td>${minutesToHours(e.driveMinutes||0)}</td><td>${minutesToHours(e.breakMinutes||0)}</td><td>${Number(e.kilometers||0).toFixed(1)}</td><td>${e.exported?'<span class="export-flag export-yes">exportiert</span>':'<span class="export-flag export-no">offen</span>'}</td><td><button class="small-btn secondary" data-action="edit" data-id="${e.id}" ${closed?'disabled':''}>Bearbeiten</button> <button class="small-btn" data-action="delete" data-id="${e.id}" ${closed?'disabled':''}>Löschen</button></td>`;
   entriesTable.appendChild(tr);
 }
 entriesTable.querySelectorAll('button[data-action="edit"]').forEach(b=>b.addEventListener('click',()=>editEntry(b.dataset.id)));
 entriesTable.querySelectorAll('button[data-action="delete"]').forEach(b=>b.addEventListener('click',()=>deleteEntry(b.dataset.id)));
 const totals=entries.reduce((a,e)=>{a.work+=calcWorkMinutes(e); a.drive+=e.driveMinutes||0; a.breaks+=e.breakMinutes||0; a.km+=e.kilometers||0; a.fuel+=e.fuelAmount||0; return a;},{work:0,drive:0,breaks:0,km:0,fuel:0});
 const credit=absenceCreditMinutes(entries,target,workDays), overtime=totals.work+credit-soll;
 const feiertage = entries.reduce((sum,e)=>sum + (e.entryType==='feiertag' ? Number(e.dayFraction||1) : 0), 0);
 stats.innerHTML=`<div class="stat"><span>Soll</span><strong>${minutesToHours(soll)}</strong></div><div class="stat"><span>Ist</span><strong>${minutesToHours(totals.work)}</strong></div><div class="stat"><span>Urlaub/Krank/FT</span><strong>${minutesToHours(credit)}</strong></div><div class="stat"><span>Überstunden</span><strong>${minutesToHours(overtime)}</strong></div><div class="stat"><span>Fahrzeit</span><strong>${minutesToHours(totals.drive)}</strong></div><div class="stat"><span>Pause</span><strong>${minutesToHours(totals.breaks)}</strong></div><div class="stat"><span>Kilometer</span><strong>${totals.km.toFixed(1)}</strong></div><div class="stat"><span>Feiertage gebucht</span><strong>${feiertage}</strong></div>`;
}
async function renderMonth(){
 const month=monthPicker.value||todayISO().slice(0,7), all=await getAll(STORE_ENTRIES); const entries=all.filter(e=>e.month===month);
 const target=Number(weeklyTargetInput.value||40), workDays=Number(workDaysPerWeekInput.value||5);
 const totals=entries.reduce((a,e)=>{a.work+=calcWorkMinutes(e); a.drive+=e.driveMinutes||0; a.breaks+=e.breakMinutes||0; a.km+=e.kilometers||0; a.fuel+=e.fuelAmount||0; if(e.entryType==='urlaub')a.urlaub+=Number(e.dayFraction||1); if(e.entryType==='krank')a.krank+=Number(e.dayFraction||1); if(e.entryType==='feiertag')a.ft+=Number(e.dayFraction||1); return a;},{work:0,drive:0,breaks:0,km:0,fuel:0,urlaub:0,krank:0,ft:0});
 const credit=absenceCreditMinutes(entries,target,workDays);
 monthStats.innerHTML=`<div class="stat"><span>Arbeitszeit</span><strong>${minutesToHours(totals.work)}</strong></div><div class="stat"><span>Fahrzeit</span><strong>${minutesToHours(totals.drive)}</strong></div><div class="stat"><span>Pause</span><strong>${minutesToHours(totals.breaks)}</strong></div><div class="stat"><span>Kilometer</span><strong>${totals.km.toFixed(1)}</strong></div><div class="stat"><span>Tankkosten</span><strong>${totals.fuel.toFixed(2)} €</strong></div><div class="stat"><span>Urlaubstage</span><strong>${totals.urlaub}</strong></div><div class="stat"><span>Kranktage</span><strong>${totals.krank}</strong></div><div class="stat"><span>Feiertage</span><strong>${totals.ft}</strong></div>`;
}

function validateEntry(){
 const warnings=[]; const type=val('entryType');
 const start=timeToMinutes(val('workStart')), end=timeToMinutes(val('workEnd')), raw=end-start;
 if(type==='arbeit'){
   if(raw<0) warnings.push({text:'Arbeitsende liegt vor Arbeitsbeginn.',blocking:true});
   if(raw>16*60) warnings.push({text:'Mehr als 16 Stunden an einem Tag wirken unplausibel.',blocking:false});
   if(num('kilometers')>0 && vehicleSelect.value==='Bitte wählen') warnings.push({text:'Kilometer ohne Fahrzeug.',blocking:false});
   if(num('fuelAmount')>0 && num('fuelLiters')===0) warnings.push({text:'Tankkosten ohne Literangabe.',blocking:false});
   if(num('fuelLiters')>0 && num('fuelAmount')===0) warnings.push({text:'Literangabe ohne Tankkosten.',blocking:false});
 }
 if(['urlaub','krank','feiertag'].includes(type)){
   if(num('kilometers')>0||num('driveMinutes')>0||num('fuelAmount')>0) warnings.push({text:'Bei Urlaub/Krank/Feiertag sollten Fahrt-, km- und Tankwerte normalerweise 0 sein.',blocking:false});
 }
 return warnings;
}
function showWarnings(warnings){
 if(!warnings.length){warningsBox.classList.add('hidden'); warningsBox.innerHTML=''; return;}
 warningsBox.classList.remove('hidden');
 warningsBox.innerHTML='<strong>Hinweise:</strong><ul>'+warnings.map(w=>`<li>${escapeHtml(w.text)}</li>`).join('')+'</ul>';
}

async function editEntry(id){ if(await isWeekClosed(weekFromDate((await getOne(STORE_ENTRIES,id)).date))){alert('Woche ist abgeschlossen.'); return;} const e=await getOne(STORE_ENTRIES,id); if(!e) return;
 $('date').value=e.date||''; $('employee').value=e.employee||''; $('entryType').value=e.entryType||'arbeit'; $('dayFraction').value=String(e.dayFraction||1); $('workStart').value=e.workStart||'08:00'; $('workEnd').value=e.workEnd||'16:00'; $('driveMinutes').value=e.driveMinutes??0; $('breakMinutes').value=e.breakMinutes??0; $('breakMode').value=e.breakMode||'autoMixed'; $('kilometers').value=e.kilometers??0; $('fuelLiters').value=e.fuelLiters??0; $('fuelAmount').value=e.fuelAmount??0; $('odometer').value=e.odometer??0; $('notes').value=e.notes||''; siteSelect.value=e.site||'Bitte wählen'; vehicleSelect.value=e.vehicle||'Bitte wählen'; $('editId').value=e.id; applyEntryTypeMode(); window.scrollTo({top:0,behavior:'smooth'})}
async function deleteEntry(id){const e=await getOne(STORE_ENTRIES,id); if(await isWeekClosed(e.week)){alert('Woche ist abgeschlossen.'); return;} if(!confirm('Eintrag wirklich löschen?')) return; await deleteOne(STORE_ENTRIES,id); await render(); await renderMonth()}

async function exportRows(){const entries=await getWeekEntries(); return entries.map(e=>({Datum:formatDate(e.date),Typ:e.entryType||'arbeit',Anteil:e.dayFraction||1,Mitarbeiter:e.employee,Baustelle:e.site,Fahrzeug:e.vehicle,Arbeitsbeginn:e.workStart,Arbeitsende:e.workEnd,Arbeitsminuten:calcWorkMinutes(e),FahrzeitMin:e.driveMinutes||0,PauseMin:e.breakMinutes||0,Kilometer:e.kilometers||0,Liter:e.fuelLiters||0,Tankkosten:e.fuelAmount||0,Kilometerstand:e.odometer||0,Notizen:e.notes||'',Exportiert:e.exported?'ja':'nein'}))}
async function exportCsv(){const rows=await exportRows(); if(!rows.length) return alert('Für diese Woche gibt es keine Einträge.'); const headers=Object.keys(rows[0]), lines=[headers.join(';')]; for(const row of rows) lines.push(headers.map(h=>`"${String(row[h]).replaceAll('"','""')}"`).join(';')); downloadBlob(new Blob([lines.join('\n')],{type:'text/csv;charset=utf-8;'}),`wochenbericht_${weekPicker.value||currentWeekValue()}.csv`)}
async function exportXlsx(){const rows=await exportRows(); if(!rows.length) return alert('Für diese Woche gibt es keine Einträge.'); const ws=XLSX.utils.json_to_sheet(rows), wb=XLSX.utils.book_new(); XLSX.utils.book_append_sheet(wb,ws,'Wochenbericht'); XLSX.writeFile(wb,`wochenbericht_${weekPicker.value||currentWeekValue()}.xlsx`)}


function setSyncStatus(text, kind='info'){
  const el = document.getElementById('syncStatus');
  if(!el) return;
  el.textContent = text;
  el.dataset.kind = kind;
}
async function autoSyncSingleEntry(entry){
  const webhook = sheetWebhookInput.value.trim();
  if(!webhook) return false;
  try{
    setSyncStatus('Synchronisiere …','info');
    const payload = {
      week: entry.week,
      weeklyTarget: Number(weeklyTargetInput.value||40),
      workDaysPerWeek: Number(workDaysPerWeekInput.value||5),
      entries: [{...entry, workMinutes: calcWorkMinutes(entry)}]
    };
    const response = await fetch(webhook,{
      method:'POST',
      headers:{'Content-Type':'text/plain;charset=utf-8'},
      body:JSON.stringify(payload)
    });
    if(!response.ok) throw new Error('HTTP '+response.status);
    const result = await response.json().catch(()=>({ok:true}));
    entry.exported = true;
    await putOne(STORE_ENTRIES, entry);
    setSyncStatus(`Auto-Sync ok (${result.updated??0} aktualisiert, ${result.inserted??1} neu)`,'ok');
    return true;
  }catch(err){
    setSyncStatus('Auto-Sync fehlgeschlagen – lokal gespeichert','warn');
    return false;
  }
}
async function syncFromGoogleSheets(){
  const webhook = sheetWebhookInput.value.trim();
  if(!webhook) return alert('Bitte zuerst den Google-Sheets-Webhook eintragen.');
  try{
    setSyncStatus('Lese Daten aus Google Sheets …','info');
    const url = webhook + (webhook.includes('?') ? '&' : '?') + 'mode=read';
    const response = await fetch(url);
    if(!response.ok) throw new Error('HTTP '+response.status);
    const data = await response.json();
    const entries = data.entries || [];
    let imported = 0, updated = 0;
    for(const item of entries){
      const existing = await getOne(STORE_ENTRIES, item.id);
      if(!existing){
        await putOne(STORE_ENTRIES, {...item, exported:true});
        imported++;
      }else{
        const localTs = new Date(existing.updatedAt || 0).getTime();
        const remoteTs = new Date(item.updatedAt || item.exportedAt || 0).getTime();
        if(remoteTs > localTs){
          await putOne(STORE_ENTRIES, {...existing, ...item, exported:true});
          updated++;
        }
      }
    }
    await render();
    await renderMonth();
    setSyncStatus(`Rücklesen abgeschlossen (${imported} neu, ${updated} aktualisiert)`,'ok');
    alert(`Rücklesen abgeschlossen. Neu: ${imported}, aktualisiert: ${updated}`);
  }catch(err){
    setSyncStatus('Rücklesen fehlgeschlagen','warn');
    alert('Rücklesen fehlgeschlagen: ' + err.message);
  }
}

async function exportToGoogleSheets(){
 const webhook=sheetWebhookInput.value.trim(); if(!webhook) return alert('Bitte zuerst den Google-Sheets-Webhook eintragen.');
 setSyncStatus('Synchronisiere offene Einträge …','info');
 const allEntries=await getWeekEntries(), entries=allEntries.filter(e=>!e.exported); if(!entries.length) return alert('Keine offenen Einträge für den Export.');
 const payload={week:weekPicker.value||currentWeekValue(),weeklyTarget:Number(weeklyTargetInput.value||40),workDaysPerWeek:Number(workDaysPerWeekInput.value||5),entries:entries.map(e=>({...e,workMinutes:calcWorkMinutes(e)}))};
 try{ const response=await fetch(webhook,{method:'POST',headers:{'Content-Type':'text/plain;charset=utf-8'},body:JSON.stringify(payload)}); if(!response.ok) throw new Error('HTTP '+response.status); const result=await response.json().catch(()=>({ok:true}));
   for(const e of entries){e.exported=true; await putOne(STORE_ENTRIES,e)} await render(); setSyncStatus(`Sync ok (${result.updated??0} aktualisiert, ${result.inserted??entries.length} neu)`,'ok'); alert(`Export abgeschlossen. ${result.updated??0} aktualisiert, ${result.inserted??entries.length} neu.`)
 }catch(err){setSyncStatus('Sync fehlgeschlagen','warn'); alert('Export fehlgeschlagen: '+err.message)}
}

async function printWeeklyReport(){
 const entries=await getWeekEntries(), week=weekPicker.value||currentWeekValue(); if(!entries.length) return alert('Für diese Woche gibt es keine Einträge.');
 const target=Number(weeklyTargetInput.value||40), workDays=Number(workDaysPerWeekInput.value||5), soll=target*60;
 const totals=entries.reduce((a,e)=>{a.work+=calcWorkMinutes(e); a.drive+=e.driveMinutes||0; a.breaks+=e.breakMinutes||0; a.km+=e.kilometers||0; a.fuel+=e.fuelAmount||0; return a;},{work:0,drive:0,breaks:0,km:0,fuel:0});
 const credit=absenceCreditMinutes(entries,target,workDays), overtime=totals.work+credit-soll;
 const html=`<html><head><title>Wochenbericht ${week}</title><style>body{font-family:Arial,sans-serif;padding:24px;color:#111}table{width:100%;border-collapse:collapse;margin-top:12px}th,td{border:1px solid #ccc;padding:8px;text-align:left;vertical-align:top}th{background:#f3f4f6}.sign{margin-top:60px;display:grid;grid-template-columns:1fr 1fr;gap:40px}.line{border-top:1px solid #111;padding-top:8px}</style></head><body><h1>Wochenbericht</h1><div>Kalenderwoche: ${week}</div><table><thead><tr><th>Datum</th><th>Typ</th><th>Anteil</th><th>Mitarbeiter</th><th>Baustelle</th><th>Arbeitszeit</th><th>Fahrzeit</th><th>Pause</th><th>km</th><th>Tankkosten</th></tr></thead><tbody>${entries.map(e=>`<tr><td>${formatDate(e.date)}</td><td>${escapeHtml(e.entryType||'arbeit')}</td><td>${e.dayFraction||1}</td><td>${escapeHtml(e.employee)}</td><td>${escapeHtml(e.site||'')}</td><td>${minutesToHours(calcWorkMinutes(e))}</td><td>${minutesToHours(e.driveMinutes||0)}</td><td>${minutesToHours(e.breakMinutes||0)}</td><td>${Number(e.kilometers||0).toFixed(1)}</td><td>${Number(e.fuelAmount||0).toFixed(2)} €</td></tr>${e.notes?`<tr><td colspan="10"><strong>Notiz:</strong> ${escapeHtml(e.notes)}</td></tr>`:''}`).join('')}</tbody></table><p><strong>Sollstunden:</strong> ${minutesToHours(soll)}</p><p><strong>Arbeitszeit:</strong> ${minutesToHours(totals.work)}</p><p><strong>Urlaub/Krank/Feiertag angerechnet:</strong> ${minutesToHours(credit)}</p><p><strong>Überstunden:</strong> ${minutesToHours(overtime)}</p><div class="sign"><div class="line">Unterschrift Mitarbeiter</div><div class="line">Unterschrift Verantwortlicher</div></div><script>window.onload=()=>window.print();</script></body></html>`;
 const win=window.open('','_blank'); win.document.open(); win.document.write(html); win.document.close();
}

async function exportBackup(){const entries=await getAll(STORE_ENTRIES), settings=await getAll(STORE_SETTINGS), meta=await getAll(STORE_META); const payload={app:'ZeitApp',version:7,exportedAt:new Date().toISOString(),entries,settings,meta}; downloadBlob(new Blob([JSON.stringify(payload,null,2)],{type:'application/json;charset=utf-8;'}),`zeitapp_backup_${todayISO()}.json`)}
async function importBackup(event){const file=event.target.files[0]; if(!file) return; try{const text=await file.text(), data=JSON.parse(text); if(!data||data.app!=='ZeitApp') return alert('Ungültige Backup-Datei.'); if(!confirm('Vorhandene Daten werden ergänzt oder überschrieben. Fortfahren?')) return; for(const item of(data.settings||[])) await putOne(STORE_SETTINGS,item); for(const item of(data.meta||[])) await putOne(STORE_META,item); for(const item of(data.entries||[])) await putOne(STORE_ENTRIES,item); await loadSettings(); await fillSelectors(); await render(); await renderMonth(); alert('Backup erfolgreich importiert.')}catch(err){alert('Backup konnte nicht importiert werden: '+err.message)} finally{restoreFile.value=''}}
function downloadBlob(blob,fileName){const url=URL.createObjectURL(blob),a=document.createElement('a'); a.href=url; a.download=fileName; a.click(); URL.revokeObjectURL(url)}
async function registerSW(){if('serviceWorker' in navigator){try{await navigator.serviceWorker.register('./service-worker.js')}catch(e){console.warn('SW registration failed',e)}}}
function setupInstall(){window.addEventListener('beforeinstallprompt',e=>{e.preventDefault(); deferredPrompt=e; installBtn.classList.remove('hidden')}); installBtn.addEventListener('click',async()=>{if(!deferredPrompt) return; deferredPrompt.prompt(); await deferredPrompt.userChoice; deferredPrompt=null; installBtn.classList.add('hidden')})}
