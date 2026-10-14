let checkingExternal=false;
async function checkExternalFile(manual=false){
  if(window.appUpdateOpen||window.appUpdateBusy||!active||!store.handle||savePromise||processing||checkingExternal||$('modal').open)return;
  checkingExternal=true;const handle=store.handle,version=revision;
  try{
    const file=await handle.getFile();if(file.size===store.stamp?.size&&file.lastModified===store.stamp?.lastModified){if(manual)toast('当前内容已是本地文件的最新版本。');return;}
    if(dirty()){conflict=true;saveError='检测到外部更新；当前还有未保存修改，自动保存已暂停。';renderBanner();setStatus();return;}
    const book=await readWorkbook(await file.arrayBuffer());if(book.sheets.length!==1)throw new Error('外部文件新增了工作表，请重新导入后编辑。');const sheet=await book.loadSheet(0),table=sheetToTable(sheet,handle.name.replace(/\.xlsx$/i,''),undefined,{includeAppendedRows:true});if(table.unsupported)throw new Error('外部更新包含暂不支持的公式，请通过导入 Excel 创建副本。');
    if(handle!==store.handle)return;if(revision!==version||dirty()){conflict=true;saveError='读取外部更新时，页面也产生了修改。请先处理冲突。';renderBanner();setStatus();return;}
    applyTemplateDefaults(table);activate(table,handle,file);lastSaved=new Date().toLocaleTimeString('zh-CN',{hour12:false});setStatus();$('save-status').innerHTML='<i></i>已同步外部更新 '+lastSaved;toast('已载入本地文件中的最新内容。');
  }catch(e){saveError=e.name==='NotFoundError'?'本地文件已被移动或删除，请另存为新文件。':e.message;conflict=true;renderBanner();setStatus();if(manual)toast(saveError,true);}finally{checkingExternal=false;}
}
window.addEventListener('focus',()=>checkExternalFile());document.addEventListener('visibilitychange',()=>{if(!document.hidden)checkExternalFile();});setInterval(()=>{if(!document.hidden)checkExternalFile();},15000);
$('refresh-files').onclick=async()=>{try{await refreshFiles();await checkExternalFile(true);}catch(e){toast(e.message,true);}};

$('sync-now').onclick=()=>checkExternalFile(true);
