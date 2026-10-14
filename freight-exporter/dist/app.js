import {uid,emptyCell,value,display,isBlank,numberFormat,escapeHTML,parseCell,createTable,recalculateAll,filterRows,groupRows,safeName,parseTSV} from './core.js';
import {LocalFileStore} from './storage.js';
import {readWorkbook,sheetToTable,mergeTables,writeWorkbook} from './excel.js';
const $=id=>document.getElementById(id),h=escapeHTML,store=new LocalFileStore();
let removedTables=new Set();
let active=null,files=[],revision=0,savedRevision=0,savePromise=null,saveError='',lastSaved='',conflict=false;
let conditions=[],join='all',groupField='',selectedGroups=null,exportColumns=new Set(),selectedRows=new Set(),view='edit',page=1,pageSize=30,search='',groupSearch='',mode='sheets';
let editorScrollY=0,rowDraftDirty=false,fitColumns=true;
try{fitColumns=localStorage.getItem('glc-column-view')!=='wide';}catch{}
let modalHandler=null,modalCancel=null,toastTimer,autoTimer,processing=false,selectingDirectory=false;
const dirty=()=>!!active&&revision!==savedRevision;
function toast(text,error=false){clearTimeout(toastTimer);$('toast').textContent=text;$('toast').className='show'+(error?' error':'');toastTimer=setTimeout(()=>$('toast').className='',6000);}
function busy(on,text='正在处理表格'){processing=on;$('loading').classList.toggle('hidden',!on);$('loading').querySelector('strong').textContent=text;}
function setStatus(){const el=$('save-status');el.className='save-status';let txt='未关联文件';if(savePromise){txt='正在写入文件…';el.classList.add('saving');}else if(saveError){txt='保存失败 · 点击重试';el.classList.add('error');}else if(active){txt=dirty()?'有未保存修改':lastSaved?'已保存 '+lastSaved:'已读取本地文件';if(dirty())el.classList.add('dirty');}el.innerHTML='<i></i>'+h(txt);el.title=saveError||store.fileName;$('manual-save').disabled=!active||!!savePromise;$('table-title').textContent=active?.name||'我的表格';document.title=window.workspaceMode==='stamp'?'印章工作区 · GLC工作台':(dirty()?'● ':'')+(active?.name?active.name+' · ':'')+'GLC工作台';for(const link of document.querySelectorAll('.table-link'))link.classList.toggle('dirty',link.dataset.name===store.fileName&&dirty());}
function changed(){revision++;if(!conflict)saveError='';setStatus();}
function resetView(){setEditorExpanded(!!active&&fitColumns&&active.columns.length>12);conditions=[];join='all';groupField=active?.columns.find(c=>c.name==='货运公司')?.id||'';selectedGroups=null;exportColumns=new Set(active?.columns.map(c=>c.id)||[]);selectedRows.clear();view='edit';page=1;search='';groupSearch='';$('search').value='';$('group-search').value='';$('condition-join').value='all';}
function activate(table,handle,file){recalculateAll(table);active=table;store.associate(handle,file);revision=savedRevision=0;lastSaved='';saveError='';conflict=false;resetView();renderAll();updateFileLocation();}
function modal(title,content,submit,handler,cancel=null){$('modal').classList.remove('scan-modal','stamp-modal','import-modal','row-modal','row-readonly');if($('modal').open)closeModal(false);$('modal-title').textContent=title;$('modal-content').oninput=null;$('modal-content').innerHTML=content;$('modal-submit').textContent=submit;$('modal-submit').disabled=false;modalHandler=handler;modalCancel=cancel;$('modal').showModal();}
function closeModal(cancelled=true){if(cancelled&&rowDraftDirty&&$('modal').classList.contains('row-modal')&&!confirm('这条记录的修改尚未保存，确定放弃吗？'))return;rowDraftDirty=false;if(cancelled&&window.scanDraftDirty&&$('modal').classList.contains('scan-modal')&&!confirm('当前草稿尚未完成，确定关闭吗？'))return;$('modal').close();$('modal').classList.remove('scan-modal','stamp-modal','import-modal','row-modal','row-readonly');const fn=modalCancel;modalHandler=null;modalCancel=null;if(cancelled)fn?.();}
$('modal-form').addEventListener('submit',async e=>{e.preventDefault();if(!modalHandler)return;const action=modalHandler;$('modal-submit').disabled=true;$('modal-content').querySelector('.modal-error')?.remove();try{await action();}catch(error){const msg=document.createElement('p');msg.className='modal-error';msg.textContent=error.message;$('modal-content').append(msg);}finally{$('modal-submit').disabled=false;}});
$('cancel-modal').onclick=()=>closeModal();$('close-modal').onclick=()=>closeModal();$('modal').addEventListener('cancel',e=>{e.preventDefault();closeModal();});
async function mayLeave(){
  if(savePromise)await savePromise;if(!dirty())return true;
  return new Promise(resolve=>{modal('还有未保存的修改',`<p class="modal-help">${h(store.fileName)} 中的修改尚未写入本地文件。</p><button type="button" id="discard-changes" class="modal-delete">不保存，继续</button>`,'保存并继续',async()=>{if(await saveCurrent(true)){closeModal(false);resolve(true);}},()=>resolve(false));$('discard-changes').onclick=()=>{closeModal(false);resolve(true);};});
}
window.addEventListener('beforeunload',event=>{if(dirty()||rowDraftDirty||savePromise||window.scanDraftDirty||window.stampDraftDirty||window.stampBusy||window.appUpdateBusy){event.preventDefault();event.returnValue='';}});
function rearmAutosave(){clearInterval(autoTimer);const seconds=Number($('autosave').value);if(seconds)autoTimer=setInterval(()=>{if(dirty()&&!savePromise&&!processing&&!conflict)saveCurrent(false);},seconds*1000);}
async function saveCurrent(manual=false){
  if(!active)return false;if(savePromise)return savePromise;if(!dirty()){if(manual)toast('所有修改已保存在本地文件中。');return true;}
  if(conflict){if(manual)showConflict();return false;}
  const version=revision,table=structuredClone(active);recalculateAll(table);
  savePromise=(async()=>{try{const bytes=await writeWorkbook([{name:table.name,rows:table.rows}],table.columns,{metadata:true,formulas:true});await store.save(bytes,{manual});savedRevision=version;lastSaved=new Date().toLocaleTimeString('zh-CN',{hour12:false});saveError='';if(manual)toast('已保存到 '+store.fileName);return true;}catch(error){saveError=error.message;conflict=error.code==='CONFLICT';if(manual){if(conflict)showConflict();else toast(error.message,true);}renderBanner();return false;}finally{savePromise=null;setStatus();renderBanner();}})();setStatus();return savePromise;
}
function renderBanner(){let el=$('save-banner');if(!el){el=document.createElement('div');el.id='save-banner';el.className='save-banner hidden';document.querySelector('.page-heading').after(el);}el.classList.toggle('hidden',!saveError);el.replaceChildren();if(!saveError)return;const text=document.createElement('span');text.textContent=saveError;const btn=document.createElement('button');btn.className='button small';btn.textContent=conflict?'处理冲突':'重试保存';btn.onclick=()=>conflict?showConflict():saveCurrent(true);el.append(text,btn);}
function showConflict(){modal('本地文件已发生变化','<p class="modal-help">为保留两边的修改，自动保存已暂停。可以将当前修改另存为新文件，或放弃当前修改，重新读取本地文件。</p><button type="button" class="modal-delete" id="reload-external">放弃当前修改，重新读取</button>','另存为新文件',()=>saveAs());$('reload-external').onclick=async()=>{const handle=store.handle;closeModal(false);await openFile(handle,true);};}
async function refreshFiles(){if(!store.directory)return;files=await store.list();renderFiles();}
function renderFiles(){
  const visible=files.map((f,i)=>({f,i})).filter(({f})=>!removedTables.has(f.name)),removed=files.filter(f=>removedTables.has(f.name));
  $('table-count').textContent=visible.length;$('restore-tables').classList.toggle('hidden',!removed.length);$('restore-tables').title=`恢复已移除的 ${removed.length} 张表格`;
  $('folder-name').textContent=store.directory?.name||'尚未选择目录';$('folder-name').title=store.directory?.name||'';
  $('table-list').innerHTML=visible.length?visible.map(({f,i})=>`<div class="table-item"><button class="table-link ${store.fileName===f.name?'active':''}" data-file-index="${i}" data-name="${h(f.name)}" title="${h(f.name)}"><span class="table-symbol"><svg class="ui-icon" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.65" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><rect x="3" y="4" width="18" height="16" rx="2"/><path d="M3 9h18M9 9v11"/></svg></span><span class="table-name">${h(f.name.replace(/\.xlsx$/i,''))}</span></button><button class="table-remove" data-remove-file="${i}" title="从工作台移除，保留本地文件" aria-label="从工作台移除 ${h(f.name)}">×</button></div>`).join(''):'<div class="nav-empty">'+(store.directory?(removed.length?'表格已从工作台移除<br>点击“恢复”可重新显示。':'目录中还没有表格<br>新建一张，或导入已有 Excel。'):'选择目录后，表格显示在这里。')+'</div>';setStatus();
}
async function removeWorkspaceTable(file){
  if(!file||!store.directory)return;const directory=store.directory,current=store.fileName===file.name;
  if(current&&!await mayLeave())return;
  if(store.directory!==directory)return;
  removedTables.add(file.name);const persisted=await saveRemovedTables(directory,removedTables);
  if(store.directory!==directory)return;
  if(current&&store.fileName===file.name){active=null;store.handle=null;store.fileName='';store.stamp=null;revision=savedRevision=0;saveError='';lastSaved='';conflict=false;resetView();renderAll();}else renderFiles();
  toast(persisted?'已从工作台移除，本地 Excel 文件保留。':'已从本次工作台移除，本地文件保留；浏览器未能记住此设置。');
}
function restoreWorkspaceTables(){
  const removed=files.filter(f=>removedTables.has(f.name));if(!removed.length)return;
  modal('恢复表格',`<p class="modal-help">选择要重新显示在工作台的本地表格。</p><div class="restore-table-list">${removed.map((f,i)=>`<label><input type="checkbox" data-restore-file="${i}" checked><span>${h(f.name)}</span></label>`).join('')}</div>`,'恢复到工作台',async()=>{const chosen=[...document.querySelectorAll('[data-restore-file]:checked')];if(!chosen.length)throw new Error('请至少选择一张表格。');for(const el of chosen)removedTables.delete(removed[Number(el.dataset.restoreFile)].name);await saveRemovedTables(store.directory,removedTables);renderFiles();closeModal(false);toast('已恢复到工作台。');});
}

async function chooseDirectory(){
  if(selectingDirectory)return false;
  if(!window.showDirectoryPicker||!window.isSecureContext){toast('请先将程序解压到本地，再用 Chrome 或 Edge 打开 HTML 文件，以选择目录并保存表格。',true);return false;}
  selectingDirectory=true;$('new-table').disabled=true;
  try{
    if(!await mayLeave())return false;
    const previous=store.directory,directory=await window.showDirectoryPicker({id:'sheet-workbench',mode:'readwrite'});
    store.directory=directory;
    try{files=await store.list();}catch(e){store.directory=previous;throw e;}
    removedTables=await loadRemovedTables(directory);store.directoryPath=await restoreDirectoryPath(directory);store.handle=null;store.fileName='';store.stamp=null;active=null;revision=savedRevision=0;saveError='';conflict=false;resetView();$('folder-gate').classList.add('hidden');renderAll();updateFileLocation();return true;
  }catch(e){if(e.name==='AbortError')toast('已取消目录选择。');else toast(e.message,true);return false;}
  finally{selectingDirectory=false;$('new-table').disabled=false;}
}
function requireFolder(){if(store.directory)return true;toast('请先选择本地目录。',true);return false;}
function fileField(defaultName){return `<label class="modal-field"><span>表格名称</span><input id="new-name" required maxlength="100" value="${h(defaultName)}" autocomplete="off"></label><div class="filename-hint">保存位置：${h(store.directory.name)} / <span id="name-preview">${h(defaultName)}.xlsx</span></div>`;}
function bindName(){const input=$('new-name');input.oninput=()=>$('name-preview').textContent=validFileName(input.value)+'.xlsx';input.focus();input.select();}
function validFileName(name){return safeName(name.trim().replace(/\.xlsx$/i,''),new Set(),100).replace(/[. ]+$/,'')||'未命名表格';}
async function commitNew(table){const filename=validFileName(table.name)+'.xlsx';table.name=filename.slice(0,-5);const bytes=await writeWorkbook([{name:table.name,rows:table.rows}],table.columns,{metadata:true,formulas:true});const handle=await store.create(filename,bytes);if(removedTables.delete(filename))await saveRemovedTables(store.directory,removedTables);activate(table,handle,await handle.getFile());lastSaved=new Date().toLocaleTimeString('zh-CN',{hour12:false});await refreshFiles();closeModal(false);toast('已创建本地文件：'+filename);}
async function newTable(){if(!store.directory&&!await chooseDirectory())return;if(!await mayLeave())return;modal('新建表格',fileField('新运费表')+'<div class="modal-help">使用当前总表的 22 个字段，创建后可增加或修改。</div>','创建并打开',async()=>{const name=$('new-name').value.trim();if(!name)throw new Error('请输入表格名称。');await commitNew(createTable(name));});bindName();}
function saveAs(){if(!active||!requireFolder())return;const copy=structuredClone(active);copy.id=uid();modal('另存为新文件',fileField(active.name+'_副本'),'保存副本',async()=>{copy.name=$('new-name').value.trim();if(!copy.name)throw new Error('请输入表格名称。');await commitNew(copy);});bindName();}
async function openFile(handle,skipLeave=false){
  if(!skipLeave&&!await mayLeave())return;
  busy(true,'正在读取本地 Excel');
  try{
    const result=await store.open(handle),book=await readWorkbook(result.bytes);
    if(book.sheets.length===1){
      const sheet=await book.loadSheet(0),table=sheetToTable(sheet,handle.name.replace(/\.xlsx$/i,''),undefined,{includeAppendedRows:true});
      if(table.unsupported)throw new Error('此文件含有工作台暂不支持的公式，请通过“导入 Excel”创建数值副本。');
      applyTemplateDefaults(table);activate(table,handle,result.file);return;
    }
    await importBook(book,handle.name);
  }catch(e){toast(e.message,true);}finally{busy(false);}
}
async function importBook(book,filename){
  const multi=book.sheets.length>1,name=filename.replace(/\.xlsx$/i,'')+(multi?'_总表':'_工作表');
  const entries=await Promise.all(book.sheets.map(async(meta,index)=>{
    try{const sheet=await book.loadSheet(index);const empty=!sheet.rows.some(row=>row?.some(c=>!isBlank(c)||c?.rawFormula!==undefined));return{sheet,index,name:meta.name,empty,selected:!empty,header:sheet.headerIndex};}
    catch(error){return{index,name:meta.name,error:error.message,selected:false};}
  }));
  let prepared=null,prepareError='';
  const sheetControls=multi?`<div class="import-sheet-heading"><strong>选择工作表</strong><div><button type="button" class="text-button" id="import-all">全选</button><button type="button" class="text-button" id="import-none">清空</button></div></div><div class="import-sheet-list"><div class="import-sheet-labels"><span>工作表</span><span>表头行</span><span>记录数</span></div>${entries.map(e=>`<div class="import-sheet-row"><label><input type="checkbox" data-import-sheet="${e.index}" ${e.selected?'checked':''} ${e.empty||e.error?'disabled':''}><span title="${h(e.name)}">${h(e.name)}</span></label><input type="number" data-import-header="${e.index}" aria-label="${h(e.name)} 表头所在行" min="1" max="${e.sheet?.rows.length||1}" value="${(e.header||0)+1}" ${e.empty||e.error?'disabled':''}><small id="import-count-${e.index}">${e.empty?'空表':e.error?'无法读取':''}</small></div>`).join('')}</div>`:`<label class="modal-field"><span>工作表</span><select id="import-sheet"><option>${h(entries[0].name)}</option></select></label><label class="modal-field"><span>表头所在行</span><input id="header-row" type="number" min="1" value="${(entries[0].header||0)+1}"></label>`;
  modal('导入为本地表格',fileField(name)+sheetControls+(multi?`<label class="import-row-option"><input type="checkbox" id="import-skip-formulas" checked> 跳过仅含公式、未录入数据的行 <small id="import-skipped"></small></label><div class="import-source"><label><input type="checkbox" id="import-source-enabled" checked> 将工作表名称填入</label><input id="import-source-field" aria-label="工作表名称存入的列名" maxlength="100" value="货运公司"></div><div class="import-caption">同名列自动对齐；已有内容保留，空白处补入分表名称。</div>`:'')+`<div class="modal-help" id="import-info" aria-live="polite"></div><label id="flatten-wrap" class="modal-alert hidden"><input type="checkbox" id="flatten-formulas"> 不支持的公式转为当前显示值</label>`,'创建并打开',async()=>{
    updateImportInfo();if(!prepared)throw new Error(prepareError||'请检查工作表和表头行。');
    if(prepared.rows.some(r=>r.cells.some(c=>c.rawFormula!==undefined&&c.missingFormulaResult)))throw new Error('部分复杂公式没有已保存的计算结果。请先用 Excel 或 WPS 打开原文件，重新计算并保存，再导入。');
    if(prepared.unsupported&&!$('flatten-formulas').checked)throw new Error('所选工作表含复杂公式，请勾选“转为当前显示值”后继续。');
    if(prepared.unsupported){for(const row of prepared.rows)for(const c of row.cells){delete c.rawFormula;delete c.missingFormulaResult;}prepared.unsupported=0;}
    prepared.name=$('new-name').value.trim();if(!prepared.name)throw new Error('请输入表格名称。');recalculateAll(prepared);await commitNew(prepared);
  });
  if(multi)$('modal').classList.add('import-modal');
  function updateImportInfo(){
    prepared=null;prepareError='';const parts=[];let error='';
    for(const entry of entries){
      if(multi){entry.selected=!!document.querySelector(`[data-import-sheet="${entry.index}"]`).checked;entry.header=Number(document.querySelector(`[data-import-header="${entry.index}"]`).value)-1;}
      else entry.header=Number($('header-row').value)-1;
      if(entry.empty||entry.error){if(!multi)error=entry.error||'这张工作表为空，请选择包含表头的文件。';continue;}
      try{const table=sheetToTable(entry.sheet,name,entry.header,{includeExtraColumns:true,skipFormulaOnlyRows:multi&&$('import-skip-formulas').checked});if(multi)$('import-count-'+entry.index).textContent=numberFormat(table.rows.length);if(entry.selected)parts.push({table,sheetName:entry.name});}
      catch(e){if(multi)$('import-count-'+entry.index).textContent='检查表头';if(entry.selected)error=entry.name+'：'+e.message;}
    }
    try{
      if(error)throw new Error(error);if(!parts.length)throw new Error('请至少选择一张工作表。');
      const sourceField=multi&&$('import-source-enabled').checked?$('import-source-field').value.trim():'';
      if(multi&&$('import-source-enabled').checked&&!sourceField)throw new Error('请输入用于保存工作表名称的列名。');
      prepared=multi?mergeTables(parts,name,{sourceField}):parts[0].table;applyTemplateDefaults(prepared);if(multi){const skipped=parts.reduce((n,p)=>n+p.table.skippedFormulaRows,0);$('import-skipped').textContent=skipped?`已跳过 ${numberFormat(skipped)} 行`:'';}
      $('import-info').textContent=(multi?`已选 ${parts.length} 张表 → 1 张总表 · `:'')+`${prepared.columns.length} 个字段 · ${numberFormat(prepared.rows.length)} 条记录`+(prepared.unsupported?` · ${prepared.unsupported} 个复杂公式需转为显示值`:'');
      $('flatten-wrap').classList.toggle('hidden',!prepared.unsupported);
    }catch(e){prepareError=e.message;$('import-info').textContent=e.message;$('flatten-wrap').classList.add('hidden');}
  }
  if(multi){
    for(const el of document.querySelectorAll('[data-import-sheet]'))el.onchange=updateImportInfo;
    for(const el of document.querySelectorAll('[data-import-header]'))el.oninput=updateImportInfo;
    for(const [id,checked] of [['import-all',true],['import-none',false]])$(id).onclick=()=>{for(const el of document.querySelectorAll('[data-import-sheet]:not(:disabled)'))el.checked=checked;updateImportInfo();};
    $('import-skip-formulas').onchange=updateImportInfo;
    $('import-source-enabled').onchange=()=>{$('import-source-field').disabled=!$('import-source-enabled').checked;updateImportInfo();};$('import-source-field').oninput=updateImportInfo;
  }else $('header-row').oninput=updateImportInfo;
  bindName();updateImportInfo();
}
function applyTemplateDefaults(table){const template=createTable('模板');for(const col of table.columns){const src=template.columns.find(c=>c.name===col.name);if(src){col.width=src.width;if(col.type==='text')col.type=src.type;if(src.computed){const fields=src.computed.fields.map(id=>template.columns.find(c=>c.id===id)?.name).map(name=>table.columns.find(c=>c.name===name)?.id);if(fields.every(Boolean))col.newRowComputed={op:src.computed.op,fields};}}}}
async function importSeed(){if(!$('seed-workbook').textContent.trim()){await newTable();return;}if(!requireFolder()||!await mayLeave())return;try{busy(true,'正在载入运费总表');const bytes=Uint8Array.from(atob($('seed-workbook').textContent.trim()),c=>c.charCodeAt(0));const book=await readWorkbook(bytes);await importBook(book,'2026年运费总表.xlsx');$('new-name').value='2026年运费总表';$('new-name').oninput();}catch(e){toast(e.message,true);}finally{busy(false);}}
function dataset(){if(!active)return{matched:[],groups:[],chosen:[],exported:[]};const matched=filterRows(active,conditions,join);const groups=groupRows(matched,active.columns,groupField);const chosen=selectedGroups===null?groups:groups.filter(g=>selectedGroups.has(g.key));const rowIDs=new Set(chosen.flatMap(g=>g.rows.map(r=>r.id)));const exported=matched.filter(r=>rowIDs.has(r.id));return{matched,groups,chosen,exported};}
function renderAll(){renderFiles();setStatus();renderBanner();$('file-location').textContent=store.directory?store.directory.name+' / '+(store.fileName||'选择或新建表格'):'LOCAL FILE · 本地文件';for(const id of ['new-record','add-row','fit-columns','wide-columns','add-column','rename-table','add-condition','choose-columns','export-button','export-top','group-field','search','reset-filters','print-open','scan-open','sync-now','copy-file-path','configure-file-path','expand-editor'])$(id).disabled=!active;$('new-table').disabled=selectingDirectory;$('new-table').title=store.directory?'在当前目录新建 Excel 表格':'先选择保存目录，再新建 Excel 表格';$('import').disabled=!store.directory;$('manual-save').disabled=!active||!!savePromise;$('locate-file').disabled=!store.directory;renderConditions();renderGroupField();renderData();updateFileLocation();}
function renderGroupField(){$('group-field').innerHTML='<option value="">不分组</option>'+(active?.columns.map(c=>`<option value="${c.id}" ${groupField===c.id?'selected':''}>${h(c.name)}</option>`).join('')||'');}
function renderData(){const data=dataset();$('total-count').textContent=numberFormat(active?.rows.length||0);$('match-count').textContent=numberFormat(data.matched.length);$('export-count').textContent=numberFormat(data.exported.length);$('group-count').textContent=mode==='single'?(data.exported.length?1:0):data.chosen.length;$('preview-badge').textContent=numberFormat(data.exported.length);$('column-count').textContent=active?active.columns.length+' 个字段':'';$('chosen-column-count').textContent=active?`${exportColumns.size} / ${active.columns.length} 个字段`:'';$('export-summary').textContent=data.exported.length?`${numberFormat(data.exported.length)} 条记录 · ${mode==='single'?1:data.chosen.length} 张表`:'暂无可导出的数据';$('export-button').disabled=!data.exported.length||!exportColumns.size;$('export-top').disabled=$('export-button').disabled;renderGroups(data.groups);renderGrid(data);}
function renderGroups(groups){$('group-label').textContent='选择分组 · '+groups.length;$('group-list').innerHTML=groups.filter(g=>g.name.toLowerCase().includes(groupSearch.toLowerCase())).map(g=>`<div class="group-item"><label><input type="checkbox" data-group="${h(g.key)}" ${selectedGroups===null||selectedGroups.has(g.key)?'checked':''}><span title="${h(g.name)}">${h(g.name)}</span></label><small>${numberFormat(g.rows.length)}</small><button class="only-group" data-only="${h(g.key)}">仅此组</button></div>`).join('')||'<span class="muted">暂无分组</span>';}
function renderGrid(data=dataset()){
  $('grid-wrap').classList.toggle('fit-columns',fitColumns);$('fit-columns').setAttribute('aria-pressed',String(fitColumns));$('wide-columns').setAttribute('aria-pressed',String(!fitColumns));$('new-record').classList.toggle('hidden',view!=='edit');
  $('edit-tab').classList.toggle('active',view==='edit');$('preview-tab').classList.toggle('active',view==='preview');$('edit-tab').setAttribute('aria-selected',view==='edit');$('preview-tab').setAttribute('aria-selected',view==='preview');$('add-row').classList.toggle('hidden',view!=='edit');$('add-column').classList.toggle('hidden',view!=='edit');$('delete-rows').classList.toggle('hidden',view!=='edit'||!selectedRows.size);$('edit-hint-text').textContent=view==='edit'?'点击行首「编辑」查看全部字段 · 可直接编辑单元格 · Ctrl / ⌘ + S 保存':'预览包含当前选中分组与字段，导出沿用当前总表格式。';
  if(!active){$('data-grid').innerHTML='';let empty=$('workspace-empty');if(!empty){empty=document.createElement('div');empty.id='workspace-empty';empty.className='empty-state';$('grid-wrap').append(empty);}empty.innerHTML='<div class="empty-icon">▦</div><strong>让每一张表，各就各位。</strong><p>新建一张表，或从已有运费总表开始。</p><div class="welcome-actions"><button id="welcome-new" class="button solid">＋ 新建表格</button><button id="welcome-seed" class="button">'+($('seed-workbook').textContent.trim()?'使用运费总表':'使用空白运费模板')+'</button></div>';$('welcome-new').onclick=newTable;$('welcome-seed').onclick=importSeed;$('row-range').textContent='0 条记录';$('page-label').textContent='1 / 1';return;}
  $('workspace-empty')?.remove();let rows=view==='preview'?data.exported:active.rows;
  if(search.trim()){const q=search.toLowerCase();rows=rows.filter(r=>r.cells.some(c=>display(c).toLowerCase().includes(q)));}
  const columns=view==='preview'?active.columns.filter(c=>exportColumns.has(c.id)):active.columns;const pages=Math.max(1,Math.ceil(rows.length/pageSize));page=Math.max(1,Math.min(page,pages));const start=(page-1)*pageSize,shown=rows.slice(start,start+pageSize);const indexes=columns.map(c=>active.columns.indexOf(c));
  const weights=columns.map(c=>/备注|说明/.test(c.name)?1.4:/单位|公司/.test(c.name)?1.25:c.type==='date'||c.name==='车号'?1.3:c.type==='number'?1.2:1);
  const colgroup=fitColumns?`<colgroup><col style="width:56px">${weights.map(w=>`<col data-column-weight="${w}">`).join('')}</colgroup>`:'';
  $('data-grid').innerHTML=colgroup+`<thead><tr><th class="row-number">${view==='edit'?'<input type="checkbox" id="select-page" aria-label="选择当前页">':'记录'}</th>${columns.map(c=>`<th title="${h(c.name)}"><span class="field-type">${c.type==='number'?'#':c.type==='date'?'◷':'T'}</span>${view==='edit'?`<button class="field-edit field-label" data-field-edit="${c.id}" aria-label="修改字段 ${h(c.name)}">${h(c.name)}<span class="field-pencil" aria-hidden="true">✎</span></button>`:h(c.name)}${c.computed?'<span class="calc-tag">ƒ</span>':''}</th>`).join('')}</tr></thead><tbody>${shown.map((row,ri)=>`<tr data-row="${row.id}"><td class="row-number ${selectedRows.has(row.id)?'checked':''}"><div class="row-tools">${view==='edit'?`<input type="checkbox" data-select-row="${row.id}" aria-label="选择第 ${start+ri+1} 行" ${selectedRows.has(row.id)?'checked':''}>`:''}<button class="row-edit" data-edit-row="${row.id}" aria-label="${view==='edit'?'编辑':'查看'}第 ${start+ri+1} 行完整记录">${view==='edit'?'编辑':'查看'}</button></div></td>${indexes.map(i=>{const c=row.cells[i],col=active.columns[i],computed=(col.computed&&!c?.manual)||c?.formula;return`<td data-row-id="${row.id}" data-col="${i}" ${view==='edit'&&!computed?'contenteditable="true" spellcheck="false" role="textbox"':''} aria-label="${h(col.name)}，第 ${start+ri+1} 行" class="${col.type==='number'?'numeric ':''}${computed?'computed ':''}${c?.t==='e'?'cell-error':''}" title="${h(col.name+'：'+display(c)+(computed?'（自动计算）':''))}"><span class="cell-value">${h(display(c))}</span></td>`;}).join('')}</tr>`).join('')||`<tr><td colspan="${columns.length+1}"><div class="empty-state"><div class="empty-icon">＋</div><strong>${view==='edit'?'开始录入第一条数据':'没有符合条件的数据'}</strong><p>${view==='edit'?'点击「新增记录」，填写全部字段。':'调整右侧筛选条件和选中分组。'}</p></div></td></tr>`}</tbody>`;
  fitGridColumns();
  $('row-range').textContent=rows.length?`${start+1}–${Math.min(start+pageSize,rows.length)} / ${numberFormat(rows.length)} 条`:'0 条记录';$('page-label').textContent=`${page} / ${pages}`;$('prev-page').disabled=page<=1;$('next-page').disabled=page>=pages;if($('select-page')){$('select-page').checked=shown.length>0&&shown.every(r=>selectedRows.has(r.id));$('select-page').onchange=e=>{for(const r of shown)e.target.checked?selectedRows.add(r.id):selectedRows.delete(r.id);renderGrid();};}
}
// Use pixel widths: table columns do not resolve mixed percentage/pixel calc widths reliably.
function fitGridColumns(){
  if(!fitColumns)return;const cols=[...$('data-grid').querySelectorAll('col[data-column-weight]')],width=$('grid-wrap').clientWidth-57;
  if(!cols.length||width<=0)return;const total=cols.reduce((n,c)=>n+Number(c.dataset.columnWeight),0);
  for(const col of cols)col.style.width=(width*Number(col.dataset.columnWeight)/total)+'px';
}
new ResizeObserver(fitGridColumns).observe($('grid-wrap'));
const operators=[['eq','等于'],['ne','不等于'],['contains','包含'],['notcontains','不包含'],['starts','开头是'],['gt','大于 / 晚于'],['gte','大于等于'],['lt','小于 / 早于'],['lte','小于等于'],['between','区间'],['empty','为空'],['notempty','不为空']];
function renderConditions(){$('conditions').innerHTML=conditions.map(c=>{const type=active.columns.find(f=>f.id===c.field)?.type;const inputType=type==='date'?'date':type==='number'?'number':'text';return`<div class="condition" data-condition="${c.id}"><div class="condition-top"><select data-prop="field" aria-label="筛选字段">${active.columns.map(f=>`<option value="${f.id}" ${f.id===c.field?'selected':''}>${h(f.name)}</option>`).join('')}</select><select data-prop="op" aria-label="筛选关系">${operators.map(([op,label])=>`<option value="${op}" ${op===c.op?'selected':''}>${label}</option>`).join('')}</select><button class="remove-condition" data-remove="${c.id}" aria-label="删除条件">×</button></div>${['empty','notempty'].includes(c.op)?'':`<div class="condition-values"><input data-prop="value" type="${inputType}" step="any" aria-label="筛选值" placeholder="输入条件" value="${h(c.value||'')}">${c.op==='between'?`<input data-prop="value2" type="${inputType}" step="any" aria-label="区间结束值" placeholder="结束值" value="${h(c.value2||'')}">`:''}</div>`}</div>`;}).join('');}
function freshRow(){
  const row={id:uid(),cells:active.columns.map(emptyCell)};
  for(let i=0;i<active.columns.length;i++){
    const col=active.columns[i],calc=col.newRowComputed;
    if(col.computed||!calc)continue;
    if(calc.fields.every(id=>active.columns.some(c=>c.id===id))){
      row.cells[i]={v:0,t:'n',z:'#,##0.000',formula:{expr:'@0@'+(calc.op==='multiply'?'*':'+')+'@1@',refs:calc.fields.map(id=>({row:row.id,col:id}))}};
    }
  }
  return row;
}
function addRow(){if(!active)return;view='edit';search='';$('search').value='';const row=freshRow();active.rows.push(row);recalculateAll(active);page=Math.ceil(active.rows.length/pageSize);changed();renderData();document.querySelector(`[data-row-id="${row.id}"][contenteditable]`)?.focus();}
function editRecord(id=null,readOnly=false){
  if(!active)return;if(!id&&active.rows.length>=100000){toast('最多支持 100,000 行。',true);return;}
  const table=active,draftTable=structuredClone(table),existing=id?draftTable.rows.find(r=>r.id===id):null;
  if(id&&!existing)return;const draft=existing||freshRow();if(!existing)draftTable.rows.push(draft);recalculateAll(draftTable);rowDraftDirty=false;
  const definitions=[['基本信息',['货运公司','建单日期','车号','始发地','目的地','卸货日期','到货单位','发货单位']],['运费信息',['吨位','单价','金额','其它费用','总金额','含税金额']],['结算与付款',['结算日期','支付方式','支付日期','发票情况','付款时间','付款金额','余额']],['备注',['备注']]],used=new Set(),sections=[];
  for(const [title,names] of definitions){const fields=table.columns.map((c,i)=>({c,i})).filter(({c})=>names.includes(c.name));if(fields.length){fields.forEach(({i})=>used.add(i));sections.push({title,fields});}}
  const other=table.columns.map((c,i)=>({c,i})).filter(({i})=>!used.has(i));if(other.length)sections.push({title:'其他字段',fields:other});
  const calculated=i=>!!((table.columns[i].computed&&!draft.cells[i]?.manual)||draft.cells[i]?.formula);
  function field({c,i}){const calc=calculated(i),long=/备注|说明/.test(c.name),attributes=`data-record-field="${i}" aria-label="${h(c.name)}" ${readOnly||calc?'readonly':''} ${c.type==='number'?'inputmode="decimal"':''}`;return `<label class="record-field ${long?'record-long':''} ${calc?'record-calculated':''}"><span>${h(c.name)}${calc?'<small>自动计算</small>':''}</span>${long?`<textarea ${attributes} rows="2">${h(display(draft.cells[i]))}</textarea>`:`<input type="text" ${attributes} value="${h(display(draft.cells[i]))}" autocomplete="off">`}</label>`;}
  modal(readOnly?'记录详情':existing?'编辑记录':'新增记录',`<div class="record-context"><span>${h(table.name)}</span><span>${table.columns.length} 个字段${existing?' · 第 '+(table.rows.findIndex(r=>r.id===id)+1)+' 条':''}</span></div><div class="record-sections">${sections.map((s,i)=>`<section class="record-section"><h3><b>${String(i+1).padStart(2,'0')}</b>${s.title}</h3><div class="record-fields">${s.fields.map(field).join('')}</div></section>`).join('')}</div>`,readOnly?'关闭':existing?'保存修改':'保存记录',async()=>{
    if(readOnly){closeModal(false);return;}
    if(active!==table)throw new Error('当前表格已变更，请重新打开记录。');
    if(!existing||rowDraftDirty){
      const target=existing?table.rows.find(r=>r.id===id):null;if(existing&&!target)throw new Error('此记录已不存在，请重新打开。');
      if(target)target.cells=structuredClone(draft.cells);else table.rows.push(structuredClone(draft));
      recalculateAll(table);changed();if(!existing){view='edit';search='';$('search').value='';page=Math.ceil(table.rows.length/pageSize);}renderData();
    }
    rowDraftDirty=false;closeModal(false);if(savePromise)await savePromise;await saveCurrent(true);
  },()=>rowDraftDirty=false);
  $('modal').classList.add('row-modal');$('modal').classList.toggle('row-readonly',readOnly);
  $('modal-content').oninput=e=>{const input=e.target.closest('[data-record-field]');if(!input||readOnly||input.readOnly)return;const i=Number(input.dataset.recordField),col=table.columns[i],before=draft.cells[i];const cell=parseCell(input.value,col.type);if(col.computed)cell.manual=true;if(cell.t==='n')cell.z=before?.z||(col.name==='吨位'?'0.000':'#,##0.000');draft.cells[i]=cell;rowDraftDirty=true;recalculateAll(draftTable);for(const el of document.querySelectorAll('[data-record-field]'))if(calculated(Number(el.dataset.recordField)))el.value=display(draft.cells[Number(el.dataset.recordField)]);};
  if(!readOnly)document.querySelector('[data-record-field]:not([readonly])')?.focus({preventScroll:true});
}
function editColumn(id=null){
  if(!active)return;const col=active.columns.find(c=>c.id===id),hasCalculation=col&&(col.computed||active.rows.some(r=>r.cells[active.columns.indexOf(col)]?.formula));
  modal(col?'修改字段':'添加字段',`<label class="modal-field"><span>列标题</span><input id="field-name" required maxlength="100" value="${h(col?.name||'')}"></label><label class="modal-field"><span>数据类型</span><select id="field-type" ${hasCalculation?'disabled':''}><option value="text">文本</option><option value="number">数字 / 金额</option><option value="date">日期</option></select></label>${hasCalculation?'<p class="modal-help">此列保留原有计算方式，可修改列标题。</p>':''}`,'保存字段',()=>{
    const name=$('field-name').value.trim();if(!name)throw new Error('请输入列标题。');if(active.columns.some(c=>c.name===name&&c.id!==id))throw new Error('列标题不能重复。');
    const target=col||{id:uid(),width:18};target.name=name;target.type=hasCalculation?'number':$('field-type').value;
    if(!col){active.columns.push(target);for(const row of active.rows)row.cells.push(emptyCell());exportColumns.add(target.id);}else if(!hasCalculation){const i=active.columns.indexOf(col);for(const row of active.rows)row.cells[i]=parseCell(display(row.cells[i]),target.type);}
    recalculateAll(active);changed();renderAll();closeModal(false);
  });$('field-type').value=col?.type||'text';$('field-name').focus();
}
function chooseColumns(){if(!active)return;modal('选择导出字段',`<div class="column-actions"><button type="button" id="columns-all" class="text-button">全选</button><button type="button" id="columns-none" class="text-button">清空</button></div><div class="column-picker">${active.columns.map(c=>`<label><input type="checkbox" data-export-column="${c.id}" ${exportColumns.has(c.id)?'checked':''}>${h(c.name)}</label>`).join('')}</div>`,'应用',()=>{const ids=[...document.querySelectorAll('[data-export-column]:checked')].map(el=>el.dataset.exportColumn);if(!ids.length)throw new Error('至少选择一个字段。');exportColumns=new Set(ids);closeModal(false);renderData();});$('columns-all').onclick=()=>document.querySelectorAll('[data-export-column]').forEach(el=>el.checked=true);$('columns-none').onclick=()=>document.querySelectorAll('[data-export-column]').forEach(el=>el.checked=false);}
function download(bytes,name,type){const url=URL.createObjectURL(new Blob([bytes],{type}));const a=document.createElement('a');a.href=url;a.download=name;document.body.append(a);a.click();a.remove();setTimeout(()=>URL.revokeObjectURL(url),60000);}
async function exportData(){if(!active)return;const data=dataset();if(!data.exported.length||!exportColumns.size)return;busy(true,'正在生成分表');try{recalculateAll(active);const cols=active.columns.filter(c=>exportColumns.has(c.id)),indexes=cols.map(c=>active.columns.indexOf(c));const project=rows=>rows.map(r=>({...r,cells:indexes.map(i=>structuredClone(r.cells[i]))}));const groups=mode==='single'?[{name:active.name,rows:project(data.exported)}]:data.chosen.map(g=>({name:g.name,title:g.name+' · '+active.name,rows:project(g.rows)}));const base=safeName(active.name,new Set(),90);if(mode==='zip'){const zip=new JSZip(),used=new Set();for(const g of groups)zip.file(safeName(g.name,used,90)+'.xlsx',await writeWorkbook([g],cols));download(await zip.generateAsync({type:'uint8array',compression:'DEFLATE'}),base+'_分表.zip','application/zip');}else download(await writeWorkbook(groups,cols),base+(mode==='single'?'_筛选结果':'_分表')+'.xlsx','application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');toast('已导出 '+numberFormat(data.exported.length)+' 条记录。');}catch(e){toast(e.message,true);}finally{busy(false);}}
$('table-list').onclick=e=>{const remove=e.target.closest('[data-remove-file]');if(remove){removeWorkspaceTable(files[Number(remove.dataset.removeFile)]).catch(e=>toast(e.message,true));return;}const button=e.target.closest('[data-file-index]');if(button)openFile(files[Number(button.dataset.fileIndex)].handle);};
$('restore-tables').onclick=restoreWorkspaceTables;
$('gate-choose').onclick=chooseDirectory;$('choose-folder').onclick=chooseDirectory;$('refresh-files').onclick=async()=>{try{await refreshFiles();toast('目录已刷新。');}catch(e){toast(e.message,true);}};$('new-table').onclick=newTable;$('manual-save').onclick=()=>saveCurrent(true);$('save-status').onclick=()=>saveError&&saveCurrent(true);$('autosave').onchange=rearmAutosave;$('rename-table').onclick=saveAs;
$('import').onclick=async()=>{if(requireFolder()&&await mayLeave())$('file-input').click();};$('file-input').onchange=async()=>{const file=$('file-input').files[0];if(!file)return;busy(true,'正在读取 Excel');try{await importBook(await readWorkbook(await file.arrayBuffer()),file.name);}catch(e){toast(e.message,true);}finally{$('file-input').value='';busy(false);}};
$('new-record').onclick=()=>editRecord();$('add-row').onclick=addRow;
for(const [id,fit] of [['fit-columns',true],['wide-columns',false]])$(id).onclick=()=>{fitColumns=fit;try{localStorage.setItem('glc-column-view',fit?'fit':'wide');}catch{}$('grid-wrap').scrollLeft=0;if(fit&&active.columns.length>12)setEditorExpanded(true);renderGrid();};$('add-column').onclick=()=>editColumn();$('choose-columns').onclick=chooseColumns;$('export-button').onclick=exportData;$('export-top').onclick=exportData;
$('edit-tab').onclick=()=>{view='edit';page=1;renderGrid();};$('preview-tab').onclick=()=>{view='preview';page=1;renderGrid();};$('search').oninput=e=>{search=e.target.value;page=1;renderGrid();};$('page-size').onchange=e=>{pageSize=Number(e.target.value);page=1;renderGrid();};$('prev-page').onclick=()=>{page--;renderGrid();};$('next-page').onclick=()=>{page++;renderGrid();};
$('delete-rows').onclick=()=>{if(!selectedRows.size)return;modal('删除选中的行',`<p class="modal-help">将删除 ${selectedRows.size} 条记录。保存后会同步到本地文件。</p>`,'删除',()=>{active.rows=active.rows.filter(r=>!selectedRows.has(r.id));selectedRows.clear();recalculateAll(active);changed();renderData();closeModal(false);});};
$('add-condition').onclick=()=>{if(!active)return;conditions.push({id:uid(),field:groupField||active.columns[0].id,op:'eq',value:''});renderConditions();};$('conditions').onclick=e=>{if(e.target.dataset.remove){conditions=conditions.filter(c=>c.id!==e.target.dataset.remove);renderConditions();renderData();}};$('conditions').addEventListener('input',e=>{const row=e.target.closest('[data-condition]'),prop=e.target.dataset.prop;if(!row||!prop)return;const c=conditions.find(c=>c.id===row.dataset.condition);c[prop]=e.target.value;if(prop==='field'||prop==='op')renderConditions();page=1;renderData();});$('condition-join').onchange=e=>{join=e.target.value;renderData();};$('reset-filters').onclick=()=>{conditions=[];selectedGroups=null;groupSearch='';$('group-search').value='';renderConditions();renderData();};$('group-field').onchange=e=>{groupField=e.target.value;selectedGroups=null;renderData();};$('group-search').oninput=e=>{groupSearch=e.target.value;renderGroups(dataset().groups);};$('select-all-groups').onclick=()=>{selectedGroups=null;renderData();};$('clear-groups').onclick=()=>{selectedGroups=new Set();renderData();};$('group-list').onchange=e=>{if(e.target.dataset.group){if(selectedGroups===null)selectedGroups=new Set(dataset().groups.map(g=>g.key));e.target.checked?selectedGroups.add(e.target.dataset.group):selectedGroups.delete(e.target.dataset.group);renderData();}};$('group-list').onclick=e=>{if(e.target.dataset.only){selectedGroups=new Set([e.target.dataset.only]);renderData();}};document.querySelectorAll('[name=mode]').forEach(el=>el.onchange=()=>{mode=el.value;renderData();});
$('data-grid').onclick=e=>{const field=e.target.closest('[data-field-edit]'),row=e.target.closest('[data-edit-row]');if(field)editColumn(field.dataset.fieldEdit);else if(row)editRecord(row.dataset.editRow,view==='preview');};$('data-grid').onchange=e=>{if(e.target.dataset.selectRow){e.target.checked?selectedRows.add(e.target.dataset.selectRow):selectedRows.delete(e.target.dataset.selectRow);$('delete-rows').classList.toggle('hidden',!selectedRows.size);e.target.closest('td').classList.toggle('checked',e.target.checked);}};
function refreshCellDisplay(el,c){const span=document.createElement('span');span.className='cell-value';span.textContent=display(c);el.replaceChildren(span);el.title=active.columns[Number(el.dataset.col)].name+'：'+display(c)+(el.classList.contains('computed')?'（自动计算）':'');el.classList.toggle('cell-error',c?.t==='e');}
function renderComputedCells(){for(const el of document.querySelectorAll('td.computed')){const row=active.rows.find(r=>r.id===el.dataset.rowId);refreshCellDisplay(el,row?.cells[Number(el.dataset.col)]);}}
$('data-grid').addEventListener('input',e=>{const cell=e.target.closest('[contenteditable]');if(!cell||!active)return;const row=active.rows.find(r=>r.id===cell.dataset.rowId),i=Number(cell.dataset.col);row.cells[i]=parseCell(cell.textContent,active.columns[i].type);if(active.columns[i].computed)row.cells[i].manual=true;if(active.columns[i].name==='吨位'&&row.cells[i].t==='n')row.cells[i].z='0.000';recalculateAll(active);renderComputedCells();changed();});
$('data-grid').addEventListener('focusout',e=>{if(!e.target.matches('[contenteditable]')||!active)return;const edited=active.rows.find(r=>r.id===e.target.dataset.rowId);refreshCellDisplay(e.target,edited?.cells[Number(e.target.dataset.col)]);recalculateAll(active);const data=dataset();$('total-count').textContent=numberFormat(active.rows.length);$('match-count').textContent=numberFormat(data.matched.length);$('export-count').textContent=numberFormat(data.exported.length);$('preview-badge').textContent=numberFormat(data.exported.length);$('group-count').textContent=mode==='single'?(data.exported.length?1:0):data.chosen.length;renderGroups(data.groups);for(const el of document.querySelectorAll('td.computed')){const row=active.rows.find(r=>r.id===el.dataset.rowId);refreshCellDisplay(el,row?.cells[Number(el.dataset.col)]);} $('export-button').disabled=!data.exported.length||!exportColumns.size;$('export-top').disabled=$('export-button').disabled;$('export-summary').textContent=`${numberFormat(data.exported.length)} 条记录 · ${mode==='single'?1:data.chosen.length} 张表`;});
$('data-grid').addEventListener('paste',e=>{const cell=e.target.closest('[contenteditable]');if(!cell||!active)return;e.preventDefault();const rows=parseTSV(e.clipboardData.getData('text/plain')),start=active.rows.findIndex(r=>r.id===cell.dataset.rowId),col=Number(cell.dataset.col);if(rows.some(r=>r.length+col>active.columns.length)){toast('粘贴内容超过现有列数，请先添加字段。',true);return;}if(start+rows.length>100000){toast('最多支持 100,000 行。',true);return;}for(let r=0;r<rows.length;r++){if(!active.rows[start+r])active.rows.push(freshRow());for(let c=0;c<rows[r].length;c++){const i=col+c;if(!active.columns[i].computed&&!active.rows[start+r].cells[i]?.formula)active.rows[start+r].cells[i]=parseCell(rows[r][c],active.columns[i].type);}}recalculateAll(active);changed();renderData();});
$('data-grid').addEventListener('keydown',e=>{const cell=e.target.closest('[contenteditable]');if(!cell)return;if(e.key==='Enter'&&!e.shiftKey){e.preventDefault();const row=cell.parentElement.nextElementSibling;const next=row?.querySelector(`[data-col="${cell.dataset.col}"][contenteditable]`);if(next)next.focus();else addRow();}});
document.addEventListener('keydown',e=>{if((e.ctrlKey||e.metaKey)&&e.key.toLowerCase()==='s'){e.preventDefault();if(window.appUpdateOpen)return;if($('modal').open&&$('modal').classList.contains('row-modal')){$('modal-form').requestSubmit();return;}if(window.workspaceMode==='stamp')$('stamp-export').click();else if(active)saveCurrent(true);}if((e.ctrlKey||e.metaKey)&&e.key.toLowerCase()==='k'){e.preventDefault();if(window.workspaceMode!=='stamp')$('search').focus();}});
busy(false);renderAll();rearmAutosave();
if(!window.showDirectoryPicker||!window.isSecureContext){$('gate-message').textContent='请使用 Windows 或 Mac 上的 Chrome / Edge 打开此 HTML 文件，以启用本地目录与文件保存。';$('gate-choose').disabled=true;$('choose-folder').disabled=true;}

function setEditorExpanded(expanded){
  const screen=$('table-screen'),button=$('expand-editor');
  if(expanded&&!screen.classList.contains('editor-expanded'))editorScrollY=window.scrollY;
  screen.classList.toggle('editor-expanded',expanded);button.setAttribute('aria-pressed',String(expanded));
  button.querySelector('span').textContent=expanded?'还原视图':'放大编辑';button.title=expanded?'还原筛选与分表面板（Esc）':'表格铺满右侧工作区';
  if(expanded)window.scrollTo(0,0);else if(editorScrollY){window.scrollTo(0,editorScrollY);editorScrollY=0;}
}
$('expand-editor').onclick=()=>setEditorExpanded(!$('table-screen').classList.contains('editor-expanded'));
document.addEventListener('keydown',e=>{if(e.key==='Escape'&&!window.appUpdateOpen&&!$('modal').open&&window.workspaceMode!=='stamp'&&$('table-screen').classList.contains('editor-expanded')){e.preventDefault();setEditorExpanded(false);$('expand-editor').focus();}});
