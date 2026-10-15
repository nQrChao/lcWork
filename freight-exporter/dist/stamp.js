import {stampAngle,stampGeometry,constrainStamp,resizeStamp} from './stamp-geometry.js';
async function sealLibrary(root){const directory=await root.getDirectoryHandle('印章库',{create:true}),seals=[];for await(const [name,handle] of directory.entries()){if(handle.kind!=='file'||!name.toLowerCase().endsWith('.png'))continue;const bytes=new Uint8Array(await(await handle.getFile()).arrayBuffer());const bitmap=await createImageBitmap(new Blob([bytes],{type:'image/png'}));seals.push({name,bytes,bitmap,ratio:bitmap.width/bitmap.height,url:objectURL(bytes,'image/png')});}return{directory,seals};}
function pageRange(text,count){const set=new Set();for(const part of(text.trim()||'1-'+count).split(/[,，]/)){const m=part.trim().match(/^(\d+)(?:\s*-\s*(\d+))?$/);if(!m)throw new Error('页码格式示例：1-3,5');const start=Number(m[1]),end=Number(m[2]||m[1]);if(start<1||end>count||end<start)throw new Error('页码超出文件范围。');for(let p=start;p<=end;p++)set.add(p-1);}return[...set].sort((a,b)=>a-b);}
let stampRoot=null,disposeStamp=null,stampExporting=false;
async function mountStampWorkspace(root){
  let library;try{library=await sealLibrary(root);}catch(e){toast('无法打开本地印章库：'+e.message,true);return;}
  disposeStamp?.();stampRoot=root;$('stamp-welcome').classList.add('hidden');$('stamp-content').classList.remove('hidden');$('stamp-library-location').textContent=root.name+' / 印章库';
  let doc=null,frame=null,index=0,selectedSeal=library.seals[0]?.name||'',selectedId=null,closed=false,sealBusy=false,placements=new Map(),sliceCache=new Map(),rotatedSeam=null;
  const seam={enabled:false,name:'',y:.5,width:.24,edge:'right',pages:[],angle:0,blend:'normal'};
  const dispose=()=>{closed=true;doc?.close();for(const s of library.seals){s.bitmap.close();URL.revokeObjectURL(s.url);}window.stampDraftDirty=false;};disposeStamp=dispose;
  $('stamp-content').innerHTML=`<div class="stamp-layout"><aside class="stamp-controls"><div class="section-title"><span>01</span><h3>本地印章库</h3></div><div id="seal-drop-zone" class="seal-drop-zone"><div id="seal-library" class="seal-library"></div><div class="seal-drop-hint">拖入印章图片保存 · 拖动印章到文件上</div></div><label class="button small full" for="seal-upload">＋ 上传并保存印章</label><input type="file" id="seal-upload" accept=".png,.jpg,.jpeg,.webp" multiple hidden><label class="inline-check"><input type="checkbox" id="seal-white" checked>去除印章图片白色背景</label><p class="filename-hint">${h(root.name)} / 印章库</p><div class="section-title"><span>02</span><h3>放置印章</h3></div><button type="button" id="place-seal" class="button small solid full" disabled>在当前页添加印章</button><button type="button" id="remove-seal" class="text-button stamp-remove">删除选中的印章</button><fieldset id="stamp-rotation" class="stamp-rotation" disabled><div class="stamp-angle-label"><label for="stamp-angle" id="stamp-angle-label">旋转印章</label><div><input id="stamp-angle" type="number" min="-180" max="180" step="1" value="0" aria-label="印章旋转角度"><span>°</span></div></div><input id="stamp-angle-range" type="range" min="-180" max="180" step="1" value="0" aria-label="拖动调整印章角度"><div class="stamp-rotation-actions"><button type="button" id="stamp-turn-left" title="向左旋转 90°">↶ 左转 90°</button><button type="button" id="stamp-angle-reset">归正</button><button type="button" id="stamp-turn-right" title="向右旋转 90°">右转 90° ↷</button></div><label class="stamp-blend-control"><input type="checkbox" id="stamp-multiply"><span>正片叠底<small>透出底下的文字和线条</small></span></label></fieldset><div class="section-title"><span>03</span><h3>骑缝章</h3><input type="checkbox" id="seam-enabled" aria-label="启用骑缝章"></div><label class="modal-field"><span>页码范围</span><input id="seam-pages" placeholder="例如 1-3,5"></label><label class="modal-field"><span>盖章边缘</span><select id="seam-edge"><option value="right">右侧</option><option value="left">左侧</option></select></label><label class="modal-field"><span>完整印章宽度</span><input id="seam-size" type="range" min="8" max="45" value="24"></label><p class="modal-help">拖动印章调整位置，右下角缩放，上方手柄旋转。骑缝章先旋转再均分，可上下拖动。</p></aside><div class="stamp-document"><div class="document-toolbar"><label class="button small solid" for="stamp-file">＋ 添加文件</label><button type="button" id="clear-stamp-file" class="button small" disabled>清空文件</button><input type="file" id="stamp-file" accept=".png,.jpg,.jpeg,.webp,.bmp,.pdf,.docx,.doc" hidden><select id="stamp-page" aria-label="盖章页码" disabled><option>第 1 页</option></select><select id="stamp-format" aria-label="盖章导出格式"><option value="pdf">导出 PDF</option><option value="png">导出 PNG</option></select><span id="stamp-info" class="muted"></span></div><p id="stamp-word-notice" class="word-preview-notice hidden">Word 按本机字体排版，请核对分页和内容后导出。</p><p id="stamp-pdf-notice" class="word-preview-notice hidden">PDF 盖章页将按预览合成，文字转为图像；未盖章页保留原内容。</p><div class="stamp-preview-scroll"><div id="stamp-stage" class="stamp-stage" tabindex="0" aria-label="盖章预览，可拖动旋转手柄，选中印章后按 Delete 删除"><div class="empty-state"><div class="empty-icon">▧</div><strong>添加需要盖章的文件</strong><p>图片、PDF 或 Word (.docx)，预览确认后导出。</p></div></div></div></div></div>`;
  $('stamp-export').disabled=true;$('stamp-export').onclick=async()=>{if(stampExporting||sealBusy||!doc)return;stampExporting=true;window.stampBusy=true;$('stamp-content').inert=true;$('stamp-export').disabled=true;$('stamp-choose-folder').disabled=true;try{await exportStamped();}catch(e){toast(e.message,true);}finally{stampExporting=false;window.stampBusy=false;$('stamp-content').inert=false;$('stamp-export').disabled=!doc||stampExporting;$('stamp-choose-folder').disabled=false;}};
  const findSeal=name=>library.seals.find(s=>s.name===name);
  function renderLibrary(){
    $('seal-library').innerHTML=library.seals.length?library.seals.map(s=>`<div class="seal-card ${selectedSeal===s.name?'active':''}" data-seal-name="${h(s.name)}" draggable="true"><button type="button" class="seal-select" aria-label="选择印章 ${h(s.name)}" aria-pressed="${selectedSeal===s.name}"><img draggable="false" src="${s.url}" alt="${h(s.name)}"><span>${h(s.name.replace(/\.png$/i,''))}</span></button><button type="button" class="seal-delete" data-delete-seal="${h(s.name)}" title="从本地印章库删除" aria-label="从本地印章库删除 ${h(s.name)}">×</button></div>`).join(''):'<p class="modal-help">将印章图片拖到这里，或点击下方上传。</p>';
    $('place-seal').disabled=!doc||!frame||!selectedSeal;
  }
  function libraryBusy(on){sealBusy=on;window.stampBusy=on;$('stamp-content').inert=on;$('stamp-choose-folder').disabled=on;$('stamp-export').disabled=on||!doc||!frame;}
  async function deleteLibrarySeal(name){
    if(sealBusy||stampExporting)return;const seal=findSeal(name);if(!seal)return;
    const used=[...placements.values()].some(list=>list.some(p=>p.name===name))||seam.enabled&&seam.name===name;
    if(!confirm(`从本地印章库删除“${name}”？${used?'当前文档中使用此印章的盖章位置也会移除。':'删除后，下次打开也不再显示。'}`))return;
    libraryBusy(true);
    try{
      try{await library.directory.removeEntry(name);}catch(e){if(e.name!=='NotFoundError')throw e;}
      library.seals=library.seals.filter(s=>s.name!==name);
      for(const [page,list] of placements)placements.set(page,list.filter(p=>p.name!==name));
      if(seam.name===name){seam.enabled=false;seam.name='';$('seam-enabled').checked=false;}
      if(!(placements.get(index)||[]).some(p=>p.id===selectedId))selectedId=null;
      if(selectedSeal===name)selectedSeal=library.seals[0]?.name||'';
      seal.bitmap.close();URL.revokeObjectURL(seal.url);sliceCache.clear();rotatedSeam=null;if(used)dirtyStamp();renderLibrary();renderMarks();toast('已从本地印章库删除。');
    }catch(e){toast('印章未删除：'+e.message,true);}finally{libraryBusy(false);}
  }

  function selectedPlacement(){return selectedId==='seam'&&seam.enabled?seam:(placements.get(index)||[]).find(p=>p.id===selectedId);}
  function syncRotationControls(){const p=frame&&selectedPlacement();$('stamp-rotation').disabled=!p;$('stamp-angle').value=p?stampAngle(p.angle):0;$('stamp-angle-range').value=p?stampAngle(p.angle):0;$('stamp-angle-label').textContent=selectedId==='seam'?'骑缝章角度':'旋转印章';$('stamp-multiply').checked=p?.blend==='multiply';}
  function rotateSelection(value){if(sealBusy||stampExporting||!frame)return;const p=selectedPlacement();if(!p)return;const n=Number(value);if(!Number.isFinite(n)){syncRotationControls();return;}p.angle=stampAngle(n);if(p===seam){sliceCache.clear();rotatedSeam=null;}else Object.assign(p,constrainStamp(p,findSeal(p.name).ratio,frame.canvas.width,frame.canvas.height));dirtyStamp();renderMarks();}
  $('stamp-multiply').onchange=e=>{const p=selectedPlacement();if(!p||sealBusy||stampExporting)return;p.blend=e.target.checked?'multiply':'normal';dirtyStamp();renderMarks();};
  $('stamp-angle-range').oninput=e=>rotateSelection(e.target.value);
  $('stamp-angle').onchange=e=>rotateSelection(e.target.value);
  $('stamp-angle').onkeydown=e=>{if(e.key==='Enter'){e.preventDefault();rotateSelection(e.target.value);}};
  $('stamp-turn-left').onclick=()=>rotateSelection((selectedPlacement()?.angle||0)-90);
  $('stamp-turn-right').onclick=()=>rotateSelection((selectedPlacement()?.angle||0)+90);
  $('stamp-angle-reset').onclick=()=>rotateSelection(0);
  function dirtyStamp(){window.stampDraftDirty=true;$('stamp-export').disabled=!doc||stampExporting;}
  async function showPage(initialFrame=null){if(!doc)return;const current=index,currentDoc=doc;frame=null;syncRotationControls();$('place-seal').disabled=true;$('stamp-export').disabled=true;$('stamp-info').textContent='正在打开文件…';const f=initialFrame||await doc.page(index,2200);if(closed||current!==index||currentDoc!==doc)return;frame=f;$('stamp-stage').classList.add('has-document');$('stamp-stage').innerHTML=`<img class="document-sheet" src="${frame.canvas.toDataURL('image/jpeg',.95)}" alt="文档第 ${index+1} 页"><div id="stamp-overlays"></div>`;$('stamp-info').textContent=doc.name;$('place-seal').disabled=!selectedSeal;$('stamp-export').disabled=false;renderMarks();}
  function heightFor(width,seal){return width*frame.canvas.width/(seal.ratio*frame.canvas.height);}
  function seamBitmap(seal){
    const angle=stampAngle(seam.angle);if(!angle)return seal.bitmap;
    if(rotatedSeam?.name===seal.name&&rotatedSeam.angle===angle)return rotatedSeam.canvas;
    const r=angle*Math.PI/180,c=Math.abs(Math.cos(r)),s=Math.abs(Math.sin(r)),canvas=document.createElement('canvas');
    canvas.width=Math.ceil(seal.bitmap.width*c+seal.bitmap.height*s);canvas.height=Math.ceil(seal.bitmap.width*s+seal.bitmap.height*c);
    const ctx=canvas.getContext('2d');ctx.translate(canvas.width/2,canvas.height/2);ctx.rotate(r);ctx.drawImage(seal.bitmap,-seal.bitmap.width/2,-seal.bitmap.height/2);
    rotatedSeam={name:seal.name,angle,canvas};sliceCache.clear();return canvas;
  }
  function seamSlice(seal,n,i){const bitmap=seamBitmap(seal),key=seal.name+':'+stampAngle(seam.angle)+':'+n+':'+i;if(sliceCache.has(key))return sliceCache.get(key);const canvas=document.createElement('canvas'),x0=Math.floor(i*bitmap.width/n),x1=Math.floor((i+1)*bitmap.width/n);canvas.width=Math.max(1,x1-x0);canvas.height=bitmap.height;canvas.getContext('2d').drawImage(bitmap,x0,0,x1-x0,bitmap.height,0,0,canvas.width,canvas.height);sliceCache.set(key,canvas);return canvas;}
  function renderMarks(){
    syncRotationControls();if(!frame||!$('stamp-overlays'))return;
    const normal=(placements.get(index)||[]).map(p=>{const seal=findSeal(p.name);return `<div class="stamp-mark ${selectedId===p.id?'selected':''}" data-mark="${p.id}" data-angle="${stampAngle(p.angle)}" style="left:${p.x*100}%;top:${p.y*100}%;width:${p.w*100}%;height:${heightFor(p.w,seal)*100}%;transform:rotate(${stampAngle(p.angle)}deg);mix-blend-mode:${p.blend==='multiply'?'multiply':'normal'}"><img draggable="false" src="${seal.url}" alt="${h(p.name)}"><button type="button" class="mark-rotate" data-rotate="${p.id}" title="拖动旋转 · 按住 Shift 每次 15°" aria-label="旋转印章，可用左右方向键微调">↻</button><button type="button" class="mark-delete" data-remove-mark="${p.id}" title="删除此处印章" aria-label="删除此处印章">×</button><i data-resize="${p.id}" title="拖动缩放"></i></div>`;}).join('');
    let seamHTML='';if(seam.enabled){const seal=findSeal(seam.name),i=seam.pages.indexOf(index);if(seal&&i>=0){const bitmap=seamBitmap(seal),fullWidth=Math.min(seam.width,frame.canvas.height/frame.canvas.width*bitmap.width/bitmap.height),w=fullWidth/seam.pages.length,h=heightFor(fullWidth,{ratio:bitmap.width/bitmap.height}),x=seam.edge==='right'?1-w:0;seamHTML=`<div class="stamp-mark seam-mark ${selectedId==='seam'?'selected':''}" data-seam="true" style="left:${x*100}%;top:${Math.max(0,Math.min(1-h,seam.y-h/2))*100}%;width:${w*100}%;height:${h*100}%;mix-blend-mode:${seam.blend==='multiply'?'multiply':'normal'}"><img draggable="false" src="${seamSlice(seal,seam.pages.length,i).toDataURL()}" alt="骑缝章第 ${i+1} 部分"><button type="button" class="mark-delete" data-remove-mark="seam" title="删除全部骑缝章" aria-label="删除全部骑缝章">×</button></div>`;}}
    $('stamp-overlays').innerHTML=normal+seamHTML;$('remove-seal').disabled=!selectedId;
  }
  $('seal-library').onclick=e=>{const remove=e.target.closest('[data-delete-seal]');if(remove){deleteLibrarySeal(remove.dataset.deleteSeal);return;}const card=e.target.closest('[data-seal-name]');if(!card||sealBusy)return;selectedSeal=card.dataset.sealName;renderLibrary();if(seam.enabled){seam.name=selectedSeal;dirtyStamp();renderMarks();}};
  async function importSeals(files,point=null){
    if(sealBusy||stampExporting)return;libraryBusy(true);let added=0;const errors=[];
    try{for(const file of files){try{
      if(!/\.(png|jpe?g|webp)$/i.test(file.name))throw new Error('印章请使用 PNG、JPG 或 WebP 图片。');
      if(file.size>30*1024*1024)throw new Error('印章图片请小于 30 MB。');
      const bitmap=await createImageBitmap(file);if(bitmap.width*bitmap.height>16000000){bitmap.close();throw new Error('印章图片过大，请选择小于 1600 万像素的图片。');}
      const canvas=document.createElement('canvas');canvas.width=bitmap.width;canvas.height=bitmap.height;const ctx=canvas.getContext('2d',{willReadFrequently:true});ctx.drawImage(bitmap,0,0);bitmap.close();
      if($('seal-white').checked){const data=ctx.getImageData(0,0,canvas.width,canvas.height);for(let i=0;i<data.data.length;i+=4)if(data.data[i]>235&&data.data[i+1]>235&&data.data[i+2]>235)data.data[i+3]=0;ctx.putImageData(data,0,0);}
      const bytes=await canvasBytes(canvas),base=safeName(file.name.replace(/\.[^.]+$/,''),new Set(),80);let name=base+'.png',n=2;
      while(true){try{await library.directory.getFileHandle(name);name=base+'_'+n+++'.png';}catch(error){if(error.name==='NotFoundError')break;throw error;}}
      const handle=await library.directory.getFileHandle(name,{create:true}),writer=await handle.createWritable();try{await writer.write(bytes);await writer.close();}catch(e){await writer.abort().catch(()=>{});throw e;}
      const converted=await createImageBitmap(new Blob([bytes],{type:'image/png'}));library.seals.push({name,bytes,bitmap:converted,ratio:converted.width/converted.height,url:objectURL(bytes,'image/png')});selectedSeal=name;renderLibrary();if(point&&doc&&frame)placeSeal(name,point.x+added*.03,point.y+added*.03);added++;
    }catch(e){errors.push(file.name+'：'+e.message);}}}finally{libraryBusy(false);}
    if(errors.length)toast((added?`已保存 ${added} 个印章；`:'')+errors.join('；'),true);else if(added)toast(`已保存 ${added} 个印章到本地印章库。`);
  }
  $('seal-upload').onchange=async e=>{const input=e.target;await importSeals([...input.files]);input.value='';};
  const sealDragType='application/x-glc-seal';
  $('seal-library').addEventListener('dragstart',e=>{if(e.target.closest('[data-delete-seal]')){e.preventDefault();return;}const card=e.target.closest('[data-seal-name]');if(!card||sealBusy){e.preventDefault();return;}e.dataTransfer.setData(sealDragType,card.dataset.sealName);e.dataTransfer.effectAllowed='copy';card.classList.add('dragging');});
  $('seal-library').addEventListener('dragend',()=>{$('stamp-stage').classList.remove('drag-over');for(const el of document.querySelectorAll('.seal-card.dragging'))el.classList.remove('dragging');});
  const dropZone=$('seal-drop-zone');
  dropZone.addEventListener('dragover',e=>{if([...e.dataTransfer.types].includes('Files')){e.preventDefault();e.dataTransfer.dropEffect='copy';dropZone.classList.add('drag-over');}});
  dropZone.addEventListener('dragleave',e=>{if(!dropZone.contains(e.relatedTarget))dropZone.classList.remove('drag-over');});
  dropZone.addEventListener('drop',e=>{e.preventDefault();e.stopPropagation();dropZone.classList.remove('drag-over');if(e.dataTransfer.files.length)importSeals([...e.dataTransfer.files]);});
  const stage=$('stamp-stage');
  stage.addEventListener('dragover',e=>{const types=[...e.dataTransfer.types];if(types.includes(sealDragType)||types.includes('Files')){e.preventDefault();e.dataTransfer.dropEffect=doc&&frame?'copy':'none';if(doc&&frame)stage.classList.add('drag-over');}});
  stage.addEventListener('dragleave',e=>{if(!stage.contains(e.relatedTarget))stage.classList.remove('drag-over');});
  stage.addEventListener('drop',e=>{e.preventDefault();e.stopPropagation();stage.classList.remove('drag-over');if(!doc||!frame){toast('请先添加需要盖章的图片、PDF 或 Word。',true);return;}if(sealBusy||stampExporting)return;const rect=stage.getBoundingClientRect(),point={x:(e.clientX-rect.left)/rect.width,y:(e.clientY-rect.top)/rect.height},name=e.dataTransfer.getData(sealDragType);if(name&&findSeal(name)){selectedSeal=name;renderLibrary();placeSeal(name,point.x,point.y);}else if(e.dataTransfer.files.length)importSeals([...e.dataTransfer.files],point);});
  // Keep stray file drops from navigating away from unsaved work.
  $('stamp-content').ondragover=e=>{if([...e.dataTransfer.types].includes('Files'))e.preventDefault();};
  $('stamp-content').ondrop=e=>{e.preventDefault();if(e.dataTransfer.files.length)toast('请将印章图片拖到左侧印章库，或拖到已打开的文件上。');};
  function resetDocument(){doc?.close();doc=null;frame=null;index=0;selectedId=null;placements=new Map();seam.enabled=false;seam.name='';seam.pages=[];seam.angle=0;seam.blend='normal';rotatedSeam=null;sliceCache.clear();syncRotationControls();window.stampDraftDirty=false;$('seam-enabled').checked=false;$('seam-pages').value='';$('stamp-page').innerHTML='<option>第 1 页</option>';$('stamp-page').disabled=true;$('stamp-info').textContent='';$('stamp-word-notice').classList.add('hidden');$('stamp-pdf-notice').classList.add('hidden');$('stamp-file').value='';$('stamp-stage').classList.remove('has-document');$('stamp-stage').innerHTML='<div class="empty-state"><div class="empty-icon">▧</div><strong>添加需要盖章的文件</strong><p>图片、PDF 或 Word (.docx)，预览确认后导出。</p></div>';$('place-seal').disabled=true;$('remove-seal').disabled=true;$('stamp-export').disabled=true;$('clear-stamp-file').disabled=true;}
  $('clear-stamp-file').onclick=()=>{if(sealBusy||stampExporting)return;if(window.stampDraftDirty&&!confirm('当前盖章修改尚未导出，确定清空文件和盖章位置吗？印章库会保留。'))return;resetDocument();toast('已清空当前文件，印章库已保留。');};
  $('stamp-file').onchange=async e=>{const input=e.target,file=input.files[0];if(!file)return;if(window.stampDraftDirty&&!confirm('更换文件会清空当前盖章位置，继续吗？')){input.value='';return;}libraryBusy(true);$('stamp-info').textContent='正在打开文件…';let next;try{next=await openLocalDocument(file);const firstFrame=await next.page(0,2200);if(closed){next.close();return;}resetDocument();doc=next;index=0;$('stamp-page').innerHTML=Array.from({length:doc.count},(_,i)=>`<option value="${i}">第 ${i+1} / ${doc.count} 页</option>`).join('');$('stamp-page').disabled=doc.count<=1;$('seam-pages').value=doc.count>1?'1-'+doc.count:'1';$('stamp-format').value=doc.pdf||doc.word?'pdf':'png';$('stamp-word-notice').classList.toggle('hidden',!doc.word);$('stamp-pdf-notice').classList.toggle('hidden',!doc.pdf);$('clear-stamp-file').disabled=false;await showPage(firstFrame);}catch(e){if(next&&next!==doc)next.close();$('stamp-info').textContent=doc?.name||'';toast(e.message,true);}finally{input.value='';if(!closed)libraryBusy(false);}};
  $('stamp-format').onchange=()=>{ $('stamp-pdf-notice').classList.toggle('hidden',!doc?.pdf||$('stamp-format').value!=='pdf'); };
  $('stamp-page').onchange=e=>{index=Number(e.target.value);selectedId=null;showPage().catch(e=>toast(e.message,true));};
  function placeSeal(name,cx=.66,cy=.41){if(!doc||!frame)return;const seal=findSeal(name);if(!seal)return;const w=Math.min(.22,.8*seal.ratio*frame.canvas.height/frame.canvas.width),height=heightFor(w,seal),p={id:uid(),name,x:Math.max(0,Math.min(1-w,cx-w/2)),y:Math.max(0,Math.min(1-height,cy-height/2)),w,angle:0,blend:'normal'};if(!placements.has(index))placements.set(index,[]);placements.get(index).push(p);selectedId=p.id;dirtyStamp();renderMarks();$('stamp-stage').focus({preventScroll:true});}
  function removeMark(id){if(sealBusy||stampExporting)return;if(id==='seam'){if(!seam.enabled)return;seam.enabled=false;$('seam-enabled').checked=false;}else{const list=placements.get(index)||[];if(!list.some(p=>p.id===id))return;placements.set(index,list.filter(p=>p.id!==id));}if(selectedId===id)selectedId=null;dirtyStamp();renderMarks();}
  $('place-seal').onclick=()=>placeSeal(selectedSeal);$('remove-seal').onclick=()=>removeMark(selectedId);
  $('stamp-stage').onclick=e=>{const btn=e.target.closest('[data-remove-mark]');if(btn){e.stopPropagation();removeMark(btn.dataset.removeMark);}};
  $('stamp-stage').addEventListener('keydown',e=>{if((e.key==='Delete'||e.key==='Backspace')&&selectedId&&!e.target.closest('button,input,textarea,select')){e.preventDefault();removeMark(selectedId);}});
  function updateSeam(){if(!doc||!selectedSeal){$('seam-enabled').checked=false;toast('请先添加文档并选择一个印章。',true);return;}try{seam.enabled=$('seam-enabled').checked;seam.pages=pageRange($('seam-pages').value,doc.count);seam.name=selectedSeal;seam.width=Number($('seam-size').value)/100;seam.edge=$('seam-edge').value;if(seam.enabled&&seam.pages.length<2){seam.enabled=false;$('seam-enabled').checked=false;throw new Error('骑缝章请至少选择两页。');}if(seam.enabled)selectedId='seam';else if(selectedId==='seam')selectedId=null;dirtyStamp();renderMarks();}catch(e){toast(e.message,true);}}
  for(const id of ['seam-enabled','seam-pages','seam-size','seam-edge'])$(id).onchange=updateSeam;
  let drag=null;
  stage.addEventListener('pointerdown',e=>{
    const mark=e.target.closest('.stamp-mark');if(!mark||!frame||sealBusy||stampExporting||e.target.closest('[data-remove-mark]'))return;
    e.preventDefault();stage.focus({preventScroll:true});const rect=stage.getBoundingClientRect();
    if(mark.dataset.seam){selectedId='seam';drag={kind:'seam',y:seam.y,startX:e.clientX,startY:e.clientY,rect};}
    else{const p=(placements.get(index)||[]).find(p=>p.id===mark.dataset.mark);if(!p)return;selectedId=p.id;const g=stampGeometry(p,findSeal(p.name).ratio,rect.width,rect.height),cx=rect.left+g.cx,cy=rect.top+g.cy;drag={kind:e.target.closest('[data-rotate]')?'rotate':e.target.closest('[data-resize]')?'resize':'move',p,original:{...p},startX:e.clientX,startY:e.clientY,rect,cx,cy,startAngle:Math.atan2(e.clientY-cy,e.clientX-cx)};}
    stage.setPointerCapture(e.pointerId);renderMarks();
  });
  stage.addEventListener('pointermove',e=>{
    if(!drag)return;const dx=e.clientX-drag.startX,dy=e.clientY-drag.startY;
    if(drag.kind==='seam')seam.y=Math.min(.95,Math.max(.05,drag.y+dy/drag.rect.height));
    else{const {p,original,rect}=drag,ratio=findSeal(p.name).ratio;
      if(drag.kind==='resize')Object.assign(p,resizeStamp(original,ratio,rect.width,rect.height,dx,dy));
      else if(drag.kind==='rotate'){const delta=(Math.atan2(e.clientY-drag.cy,e.clientX-drag.cx)-drag.startAngle)*180/Math.PI,value=(original.angle||0)+delta,angle=e.shiftKey?Math.round(value/15)*15:Math.round(value);Object.assign(p,constrainStamp({...original,angle},ratio,rect.width,rect.height));}
      else Object.assign(p,constrainStamp({...original,x:original.x+dx/rect.width,y:original.y+dy/rect.height},ratio,rect.width,rect.height));
    }
    dirtyStamp();renderMarks();
  });
  const endDrag=()=>{drag=null;};stage.addEventListener('pointerup',endDrag);stage.addEventListener('pointercancel',endDrag);stage.addEventListener('lostpointercapture',endDrag);
  stage.addEventListener('keydown',e=>{const control=e.target.closest('[data-rotate]');if(!control||!['ArrowLeft','ArrowRight'].includes(e.key))return;e.preventDefault();selectedId=control.dataset.rotate;rotateSelection((selectedPlacement()?.angle||0)+(e.key==='ArrowRight'?1:-1)*(e.shiftKey?15:1));stage.querySelector(`[data-rotate="${selectedId}"]`)?.focus({preventScroll:true});});
  async function exportStamped(){
    if(!doc)throw new Error('请添加需要盖章的文件。');
    if(seam.enabled){seam.pages=pageRange($('seam-pages').value,doc.count);if(seam.pages.length<2)throw new Error('骑缝章至少需要两页。');}
    const format=$('stamp-format').value,base=safeName(doc.name.replace(/\.[^.]+$/,''),new Set(),90);
    $('stamp-info').textContent='正在生成盖章文件…';
    const lib=await pdfWriter(),pdf=format==='pdf'?await lib.PDFDocument.create():null;
    const source=pdf&&doc.pdf?await lib.PDFDocument.load(doc.bytes.slice()):null,zip=new JSZip(),copies=new Map();let singleBytes;
    if(source){const indices=Array.from({length:doc.count},(_,i)=>i).filter(i=>!(placements.get(i)||[]).length&&!(seam.enabled&&seam.pages.includes(i)));const pages=await pdf.copyPages(source,indices);indices.forEach((i,n)=>copies.set(i,pages[n]));}
    for(let p=0;p<doc.count;p++){
      const marks=(placements.get(p)||[]).map(pos=>({pos,seal:findSeal(pos.name)}));
      const hasSeam=seam.enabled&&seam.pages.includes(p);
      if(source&&!marks.length&&!hasSeam){
        // Untouched pages retain their original text and content streams.
        pdf.addPage(copies.get(p));
      }else{
        const pageFrame=await doc.page(p,2600),canvas=pageFrame.canvas,ctx=canvas.getContext('2d');
        if(hasSeam){
          const seal=findSeal(seam.name),bitmap=seamBitmap(seal),fullWidth=Math.min(seam.width,canvas.height/canvas.width*bitmap.width/bitmap.height),w=fullWidth/seam.pages.length,h=fullWidth*canvas.width/(bitmap.width/bitmap.height*canvas.height);
          marks.push({pos:{x:seam.edge==='right'?1-w:0,y:Math.max(0,Math.min(1-h,seam.y-h/2)),w,angle:0,blend:seam.blend},seal,part:seamSlice(seal,seam.pages.length,seam.pages.indexOf(p))});
        }
        // Bake marks into the displayed page. Appending operators to arbitrary PDFs
        // can inherit clipping/alpha or sit underneath opaque annotation appearances.
        for(const {pos,seal,part} of marks){
          const bitmap=part||seal.bitmap,g=stampGeometry(pos,bitmap.width/bitmap.height,canvas.width,canvas.height);
          ctx.save();ctx.globalCompositeOperation=pos.blend==='multiply'?'multiply':'source-over';ctx.translate(g.cx,g.cy);ctx.rotate(g.radians);ctx.drawImage(bitmap,-g.width/2,-g.height/2,g.width,g.height);ctx.restore();
        }
        const bytes=await canvasBytes(canvas);
        if(format==='png'){zip.file(base+'_盖章_第'+(p+1)+'页.png',bytes);singleBytes=bytes;}
        else{
          const size=pageFrame.paperSize||(pageFrame.base?[pageFrame.base.width,pageFrame.base.height]:[canvas.width*.3,canvas.height*.3]);
          const image=await pdf.embedPng(bytes),target=pdf.addPage(size);target.drawImage(image,{x:0,y:0,width:target.getWidth(),height:target.getHeight()});
        }
      }
      $('stamp-info').textContent='正在生成 '+(p+1)+' / '+doc.count+' 页';
    }
    if(format==='pdf')download(await pdf.save(),base+'_盖章.pdf','application/pdf');
    else if(doc.count===1)download(singleBytes,base+'_盖章.png','image/png');
    else download(await zip.generateAsync({type:'uint8array'}),base+'_盖章图片.zip','application/zip');
    window.stampDraftDirty=false;$('stamp-info').textContent='已导出盖章文件';toast('盖章文件已导出。');
  }

  renderLibrary();syncRotationControls();
}
async function chooseStampDirectory(){
  if(stampExporting||window.stampBusy)return;
  if(window.stampDraftDirty&&!confirm('更换印章目录会清空尚未导出的盖章位置，继续吗？'))return;
  try{const options={id:'glc-stamp-library',mode:'readwrite'};if(stampRoot||store.directory)options.startIn=stampRoot||store.directory;const root=await window.showDirectoryPicker(options);await mountStampWorkspace(root);}catch(e){if(e.name!=='AbortError')toast(e.message,true);}
}
$('stamp-choose-folder').onclick=chooseStampDirectory;$('stamp-start').onclick=chooseStampDirectory;
if(!window.showDirectoryPicker){$('stamp-start').disabled=true;$('stamp-choose-folder').disabled=true;}
