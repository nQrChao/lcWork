import {normalizeDirectoryPath,joinLocalPath} from './core.js';

window.workspaceMode='table';
function showWorkspace(mode){
  window.workspaceMode=mode;
  const stamp=mode==='stamp';
  $('table-screen').classList.toggle('hidden',stamp);$('stamp-screen').classList.toggle('hidden',!stamp);
  $('table-navigation').classList.toggle('hidden',stamp);$('table-storage').classList.toggle('hidden',stamp);$('stamp-navigation').classList.toggle('hidden',!stamp);
  $('table-workspace-open').classList.toggle('active',!stamp);$('stamp-open').classList.toggle('active',stamp);
  $('folder-gate').classList.toggle('hidden',stamp||!!store.directory);
  document.title=stamp?'印章工作区 · GLC工作台':(active?active.name+' · ':'')+'GLC工作台';
}
$('table-workspace-open').onclick=()=>showWorkspace('table');$('stamp-open').onclick=()=>showWorkspace('stamp');

function openPathProfiles(){return new Promise((resolve,reject)=>{const request=indexedDB.open('glc-local-directory-settings',1);request.onupgradeneeded=()=>request.result.createObjectStore('paths',{keyPath:'id'});request.onsuccess=()=>resolve(request.result);request.onerror=()=>reject(request.error);request.onblocked=()=>reject(new Error('目录设置暂不可用'));});}
async function readPathProfiles(){const db=await openPathProfiles();try{return await new Promise((resolve,reject)=>{const r=db.transaction('paths').objectStore('paths').getAll();r.onsuccess=()=>resolve(r.result);r.onerror=()=>reject(r.error);});}finally{db.close();}}
async function restoreDirectoryPath(directory){
  try{if(!directory.isSameEntry)return '';for(const profile of await readPathProfiles())try{if(await directory.isSameEntry(profile.handle))return normalizeDirectoryPath(profile.path);}catch{}}catch{}
  return '';
}
async function rememberDirectoryPath(directory,path){
  try{
    let id=uid();for(const profile of await readPathProfiles())try{if(await directory.isSameEntry(profile.handle)){id=profile.id;break;}}catch{}
    const db=await openPathProfiles();try{await new Promise((resolve,reject)=>{const transaction=db.transaction('paths','readwrite');transaction.oncomplete=resolve;transaction.onerror=()=>reject(transaction.error);transaction.onabort=()=>reject(transaction.error);transaction.objectStore('paths').put({id,handle:directory,path});});}finally{db.close();}
  }catch{} // Directory handles may not persist in private browsing; the current session still works.
}
function updateFileLocation(){
  if(!store.directory)return;
  const path=store.directoryPath?(store.fileName?joinLocalPath(store.directoryPath,store.fileName):store.directoryPath):store.directory.name+' / '+(store.fileName||'选择或新建表格');
  $('file-location').textContent=path;$('file-location').title=path;
}
async function copyTextToClipboard(text){
  try{await navigator.clipboard.writeText(text);return;}catch{}
  const input=document.createElement('textarea');input.value=text;input.style.cssText='position:fixed;left:0;top:0;opacity:0';document.body.append(input);input.select();const copied=document.execCommand('copy');input.remove();if(!copied)throw new Error('浏览器未允许复制，请在路径框中手动复制。');
}
async function copyActivePath(){
  if(!active||!store.fileName)return;
  if(!store.directoryPath){configureDirectoryPath(true);return;}
  try{await copyTextToClipboard(joinLocalPath(store.directoryPath,store.fileName));toast('已复制文件绝对路径。');}catch(e){toast(e.message,true);configureDirectoryPath();}
}
function configureDirectoryPath(copyAfter=false){
  if(!store.directory)return;
  modal('设置目录完整路径',`<p class="modal-help">在 Finder 或资源管理器中复制「${h(store.directory.name)}」目录的完整路径，粘贴到下方。</p><label class="modal-field"><span>当前目录的绝对路径</span><input id="directory-absolute-path" autocomplete="off" spellcheck="false" value="${h(store.directoryPath||'')}" placeholder="C:\\工作表 或 /Users/姓名/Documents"></label><p class="modal-help">路径仅用于展示与复制，不改变实际保存位置。浏览器不会提供真实绝对路径，需要填写一次。</p>`,copyAfter?'保存并复制文件路径':'保存路径',async()=>{
    const path=normalizeDirectoryPath($('directory-absolute-path').value),tail=path.split(/[\\/]/).filter(Boolean).pop()||'/',expected=store.directory.name;
    const windows=!path.startsWith('/'),matches=windows?tail.toLowerCase()===expected.toLowerCase():tail===expected;
    if(!matches&&!/^(?:[a-z]:\\|\/)$/i.test(path))throw new Error('路径末尾应为当前目录「'+expected+'」，请确认没有粘贴其他目录或文件的路径。');
    if(copyAfter)await copyTextToClipboard(joinLocalPath(path,store.fileName));
    store.directoryPath=path;updateFileLocation();closeModal(false);toast(copyAfter?'已保存目录路径，并复制文件绝对路径。':'目录路径已保存。');await rememberDirectoryPath(store.directory,path);
  });$('directory-absolute-path').focus();$('directory-absolute-path').select();
}
$('copy-file-path').onclick=copyActivePath;$('configure-file-path').onclick=()=>configureDirectoryPath();
$('locate-file').onclick=async()=>{if(!store.directory)return;try{await window.showDirectoryPicker({id:'glc-locate-folder',startIn:store.directory,mode:'read'});}catch(e){if(e.name!=='AbortError')toast(e.message,true);}};

// Hidden table lists belong to directory identity, never just a folder name.
const removedTableSessions=[];
function openRemovedTableDB(){return new Promise((resolve,reject)=>{const r=indexedDB.open('glc-workspace-tables',1);r.onupgradeneeded=()=>r.result.createObjectStore('removed',{keyPath:'id'});r.onsuccess=()=>resolve(r.result);r.onerror=()=>reject(r.error);r.onblocked=()=>reject(new Error('表格列表设置暂不可用'));});}
async function sameDirectory(a,b){if(a===b)return true;try{return !!a?.isSameEntry&&await a.isSameEntry(b);}catch{return false;}}
async function removedTableProfile(directory){
  for(const profile of removedTableSessions)if(await sameDirectory(directory,profile.handle))return profile;
  try{const db=await openRemovedTableDB();let entries;try{entries=await new Promise((resolve,reject)=>{const r=db.transaction('removed').objectStore('removed').getAll();r.onsuccess=()=>resolve(r.result);r.onerror=()=>reject(r.error);});}finally{db.close();}for(const profile of entries)if(await sameDirectory(directory,profile.handle)){removedTableSessions.push(profile);return profile;}}catch{}
  const profile={id:uid(),handle:directory,names:[]};removedTableSessions.push(profile);return profile;
}
async function loadRemovedTables(directory){const profile=await removedTableProfile(directory);return new Set(profile.names);}
async function saveRemovedTables(directory,names){
  const profile=await removedTableProfile(directory);profile.names=[...names];
  try{const db=await openRemovedTableDB();try{await new Promise((resolve,reject)=>{const tx=db.transaction('removed','readwrite');tx.oncomplete=resolve;tx.onerror=()=>reject(tx.error);tx.onabort=()=>reject(tx.error);tx.objectStore('removed').put(profile);});}finally{db.close();}return true;}catch{return false;}
}
