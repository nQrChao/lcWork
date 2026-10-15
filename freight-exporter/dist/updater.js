import {UPDATE_MAX_BYTES,validateUpdateManifest,readBuildMetadata,updateDigest,verifyUpdateArtifact} from './update-core.js';
const currentBuild=Object.freeze(JSON.parse($('glc-build').textContent));
const updateUI=$('update-dialog');
let updateCandidate=null,updateBytes=null,updateDirectory=null,updateController=null,updateRunning=false,updateInstalled=false;
window.appUpdateBusy=false;window.appUpdateOpen=false;
$('app-version').textContent='v'+currentBuild.version;
function updateShow(id,show=true){$(id).classList.toggle('hidden',!show);}
function updateError(message){$('update-error').textContent=message;updateShow('update-error');}
function updateProgress(message){$('update-status').textContent=message;}
function updateLock(on,writing=false){updateRunning=on;window.appUpdateBusy=on;for(const id of ['update-install','update-download','update-retry','update-change-dir'])$(id).disabled=on;$('update-close').disabled=writing;}
function closeUpdate(){if($('update-close').disabled)return;updateController?.abort();updateUI.close();window.appUpdateOpen=false;}
function openUpdate(){if(!updateUI.open)updateUI.showModal();window.appUpdateOpen=true;}
function updateDB(){return new Promise((resolve,reject)=>{const r=indexedDB.open('glc-program-updater',1);r.onupgradeneeded=()=>r.result.createObjectStore('settings');r.onsuccess=()=>resolve(r.result);r.onerror=()=>reject(r.error);r.onblocked=()=>reject(new Error('更新设置暂不可用'));});}
async function rememberUpdateDirectory(directory){try{const db=await updateDB();try{await new Promise((resolve,reject)=>{const t=db.transaction('settings','readwrite');t.objectStore('settings').put(directory,'program-directory');t.oncomplete=resolve;t.onerror=()=>reject(t.error);});}finally{db.close();}}catch{}}
(async()=>{try{const db=await updateDB();try{updateDirectory=await new Promise((resolve,reject)=>{const r=db.transaction('settings').objectStore('settings').get('program-directory');r.onsuccess=()=>resolve(r.result);r.onerror=()=>reject(r.error);});}finally{db.close();}}catch{}})();
async function fetchUpdateBytes(url,{limit=128*1024,timeout=20000,progress=null}={}){
  const controller=new AbortController();updateController=controller;const timer=setTimeout(()=>controller.abort(),timeout);
  try{
    const response=await fetch(url,{signal:controller.signal,cache:'no-store',credentials:'omit',referrerPolicy:'no-referrer',redirect:'error'});
    if(!response.ok){if([404,409,422].includes(response.status)){const error=new Error('暂时没有可下载的更新，请稍后重试。');error.updateUnavailable=true;throw error;}if([403,429].includes(response.status))throw new Error('检查过于频繁，请稍后重试。');throw new Error('更新服务器返回错误（'+response.status+'），请稍后重试。');}
    const length=Number(response.headers.get('content-length')||0);if(length>limit)throw new Error('更新文件超出允许大小。');
    const reader=response.body.getReader(),chunks=[];let count=0;
    while(true){const {done,value}=await reader.read();if(done)break;count+=value.length;if(count>limit){await reader.cancel();throw new Error('更新文件超出允许大小。');}chunks.push(value);progress?.(count,length);}
    const bytes=new Uint8Array(count);let offset=0;for(const chunk of chunks){bytes.set(chunk,offset);offset+=chunk.length;}return bytes;
  }catch(e){if(e.name==='AbortError')throw new Error('更新请求已取消或超时，请重试。');if(e instanceof TypeError)throw new Error('无法检查更新，请检查网络后重试。');throw e;}
  finally{clearTimeout(timer);if(updateController===controller)updateController=null;}
}
async function checkProgramUpdate(){
  if(updateRunning)return;if(updateInstalled){openUpdate();return;}updateCandidate=null;updateBytes=null;updateInstalled=false;openUpdate();updateLock(true);$('update-title').textContent='检查更新';
  for(const id of ['update-error','update-install','update-download','update-reload','update-change-dir','update-retry'])updateShow(id,false);
  $('update-hint').textContent='';updateProgress('正在检查更新…');
  try{
    const api='https://api.github.com/repos/'+currentBuild.repository+'/git/ref/heads/'+encodeURIComponent(currentBuild.channel);
    const head=JSON.parse(new TextDecoder().decode(await fetchUpdateBytes(api,{limit:1024*1024})));
    if(head.object?.type!=='commit'||!/^[a-f0-9]{40}$/.test(head.object?.sha||''))throw new Error('更新分支信息无效。');
    const base='https://raw.githubusercontent.com/'+currentBuild.repository+'/'+head.object.sha+'/';
    const manifest=validateUpdateManifest(JSON.parse(new TextDecoder().decode(await fetchUpdateBytes(base+'update.json'))));
    if(!updateUI.open)return;
    if(manifest.buildId===currentBuild.buildId){$('update-title').textContent='已是最新版本';updateProgress('暂无可用更新。');updateShow('update-retry');return;}
    updateCandidate={manifest,base};$('update-title').textContent='发现新版本';
    updateProgress('有新版本可用，点击下方按钮即可更新。');
    const canInstall=location.protocol==='file:'&&!!window.showDirectoryPicker;
    updateShow('update-download',!canInstall);updateShow('update-install',canInstall);updateShow('update-change-dir',!!updateDirectory);
  }catch(e){if(updateUI.open){
    if(e.updateUnavailable){$('update-title').textContent='暂无可用更新';updateProgress('当前暂无可下载的新版本。');}
    else{$('update-title').textContent='检查失败';updateProgress('暂时无法确认是否有更新。');updateError(e instanceof SyntaxError?'更新信息暂不可用，请稍后重试。':e.message);}
    updateShow('update-retry');
  }}
  finally{updateLock(false);}
}
function updateDraftProblem(){if(window.stampBusy||processing)return '当前文件仍在处理，请完成后再更新。';if(rowDraftDirty||window.scanDraftDirty)return '请先保存或关闭当前录入草稿，再更新程序。';if(window.stampDraftDirty)return '当前盖章修改尚未导出，请先导出盖章文件或清空草稿，再更新程序。';return '';}
async function prepareUpdate(){
  const problem=updateDraftProblem();if(problem)throw new Error(problem);
  if(savePromise)await savePromise;
  if(dirty()){
    updateUI.close();window.appUpdateOpen=false;
    const ok=await saveCurrent(true);
    if(!ok){if(!$('modal').open)openUpdate();throw new Error('表格尚未成功保存，已停止更新。请先处理保存错误。');}
    openUpdate();if(dirty())throw new Error('保存期间又产生了修改，请先完成编辑再更新。');
  }
  const remaining=updateDraftProblem();if(remaining)throw new Error(remaining);
}

async function downloadUpdate(){
  if(updateBytes)return updateBytes;if(!updateCandidate)throw new Error('请先检查更新。');
  const {manifest,base}=updateCandidate;
  const bytes=await fetchUpdateBytes(base+encodeURIComponent(manifest.file),{limit:Math.min(UPDATE_MAX_BYTES,manifest.bytes),timeout:120000,progress:n=>updateProgress('正在下载新版… '+Math.min(100,Math.floor(n/manifest.bytes*100))+'%')});
  updateProgress('正在校验新版文件…');await verifyUpdateArtifact(bytes,manifest);updateBytes=bytes;return bytes;
}
async function writeUpdateFile(handle,bytes){let stream;try{stream=await handle.createWritable({mode:'exclusive'});await stream.write(bytes);await stream.close();}catch(e){try{await stream?.abort();}catch{}throw e;}}
function programFileName(){let name;try{name=decodeURIComponent(new URL(location.href).pathname.split('/').pop());}catch{}if(!name||!/^.+\.html?$/i.test(name)||/[\\/]/.test(name))throw new Error('请从本地 HTML 文件打开程序后安装更新。');return name;}
async function requestProgramDirectory(){
  // Called directly from a click: keep the permission request within user activation.
  if(updateDirectory){const directory=updateDirectory;const p=await directory.requestPermission({mode:'readwrite'});if(p!=='granted')throw new Error('未获得程序目录写入权限，请重新选择程序目录。');return directory;}
  return window.showDirectoryPicker({id:'glc-program-update',mode:'readwrite'});
}
async function installProgramUpdate(){
  if(updateRunning||!updateCandidate)return;updateShow('update-error',false);const problem=updateDraftProblem();if(problem){updateError(problem);return;}
  if(dirty()||savePromise){updateLock(true);try{await prepareUpdate();updateProgress('表格已保存，请再次点击「立即更新」授权程序目录。');}catch(e){updateError(e.message);}finally{updateLock(false);}return;}
  updateLock(true);updateProgress('请选择当前程序所在目录。');let backupPath='',written=false;
  try{
    const directory=await requestProgramDirectory();if(!updateUI.open)return;
    const name=programFileName();let handle;try{handle=await directory.getFileHandle(name);}catch(e){if(e.name==='NotFoundError')throw new Error('这个目录里没有「'+name+'」。请重新选择当前程序所在目录。');throw e;}
    const original=await handle.getFile(),originalBytes=new Uint8Array(await original.arrayBuffer()),meta=readBuildMetadata(new TextDecoder().decode(originalBytes));
    if(meta.buildId!==currentBuild.buildId)throw new Error('所选程序版本与本页不一致，可能已被更新。请重新打开该文件，或重新选择正确目录。');
    const beforeHash=await updateDigest(originalBytes);
    updateDirectory=directory;await rememberUpdateDirectory(directory);updateShow('update-change-dir');
    const bytes=await downloadUpdate();if(!updateUI.open)return;await prepareUpdate();
    updateLock(true,true);updateProgress('正在备份旧版并安装…');
    const backupDir=await directory.getDirectoryHandle('LC程序备份',{create:true});const backupName=name.replace(/\.html?$/i,'')+'_'+new Date().toISOString().replace(/[:.]/g,'-')+'_'+uid().slice(0,8)+'.html';
    const backup=await backupDir.getFileHandle(backupName,{create:true});await writeUpdateFile(backup,originalBytes);
    if(await updateDigest(await (await backup.getFile()).arrayBuffer())!==beforeHash)throw new Error('旧版备份校验失败，已停止安装。');backupPath='LC程序备份 / '+backupName;
    if(await updateDigest(await (await handle.getFile()).arrayBuffer())!==beforeHash)throw new Error('程序文件在更新期间被修改，已停止覆盖。请重新打开程序。');
    await writeUpdateFile(handle,bytes);written=true;
    if(await updateDigest(await (await handle.getFile()).arrayBuffer())!==updateCandidate.manifest.sha256)throw new Error('写入后的程序校验未通过。');
    updateInstalled=true;$('update-title').textContent='更新已安装';updateProgress('点击「重新加载新版」启用新功能。');$('update-hint').textContent='旧版备份：'+backupPath+'。重新加载后，选择原表格目录即可继续工作。';
    for(const id of ['update-install','update-download','update-change-dir'])updateShow(id,false);updateShow('update-reload');
  }catch(e){if(updateUI.open){updateError(e.name==='AbortError'?'已取消目录选择，程序文件未修改。':written?e.message+' 旧版备份：'+backupPath+'，可从备份恢复。':e.name==='NotAllowedError'?'没有写入权限，未完成安装。请授权后重试。':e.message);updateShow('update-change-dir');}}
  finally{updateLock(false);$('update-close').disabled=false;}
}
$('check-update').onclick=checkProgramUpdate;$('update-retry').onclick=checkProgramUpdate;$('update-close').onclick=closeUpdate;updateUI.addEventListener('cancel',e=>{e.preventDefault();closeUpdate();});
$('update-change-dir').onclick=()=>{updateDirectory=null;rememberUpdateDirectory(null);updateShow('update-change-dir',false);updateProgress('下次点击「立即更新」时，请选择当前程序所在目录。');};
$('update-install').onclick=installProgramUpdate;
$('update-download').onclick=async()=>{if(updateRunning)return;updateShow('update-error',false);updateLock(true);try{const bytes=await downloadUpdate();if(!updateUI.open)return;download(bytes,'LC工作台_'+updateCandidate.manifest.version+'.html','text/html;charset=utf-8');updateProgress('新版已下载。保存当前工作后，用 Chrome / Edge 打开新文件。');}catch(e){if(updateUI.open)updateError(e.message);}finally{updateLock(false);}};
$('update-reload').onclick=async()=>{if(!updateInstalled||updateRunning)return;try{await prepareUpdate();location.reload();}catch(e){updateError(e.message);}};
