let wordReaderPromise;
function wordImageSource(value){
  if(!value?.startsWith('data:application/octet-stream;base64,'))return value;
  const head=atob(value.split(',')[1].slice(0,32));
  const type=head.startsWith('\x89PNG')?'png':head.startsWith('\xff\xd8')?'jpeg':head.startsWith('GIF8')?'gif':head.startsWith('BM')?'bmp':head.startsWith('RIFF')?'webp':/^\s*</.test(head)?'svg+xml':null;
  return type?value.replace('application/octet-stream','image/'+type):value;
}
async function wordReader(){
  return wordReaderPromise??=(async()=>{
    loadEmbeddedScript('docx-preview.min.js');loadEmbeddedScript('html2canvas.min.js');
    if(!window.docx||!window.html2canvas)throw new Error('Word 组件加载失败，请重新打开工作台。');
    return{render:window.docx.renderAsync,capture:window.html2canvas};
  })();
}

// Use an isolated document so imported Word styles cannot change the workbench.
// CSS column fragmentation paginates flowing paragraphs/tables without cutting text
// into arbitrary bitmap strips. Explicit Word page/section breaks remain intact.
function paginateWordSection(section,win){
  const doc=section.ownerDocument,style=win.getComputedStyle(section);
  const width=parseFloat(style.width)||794,height=parseFloat(style.minHeight)||1123;
  section.style.width=width+'px';section.style.height=height+'px';section.style.minHeight='0';
  const viewport=doc.createElement('div'),flow=doc.createElement('div');
  viewport.className='word-page-viewport';flow.className='word-page-flow';
  const content=[...section.children].filter(el=>!['HEADER','FOOTER'].includes(el.tagName));
  section.insertBefore(viewport,content[0]||section.querySelector('footer'));
  for(const el of content){el.style.marginBottom='0';flow.append(el);}
  viewport.append(flow);
  const bodyWidth=viewport.clientWidth,bodyHeight=viewport.clientHeight;
  if(bodyWidth<80||bodyHeight<80)throw new Error('Word 页边距过大，无法排版，请调整后重新导入。');
  flow.style.width=bodyWidth+'px';flow.style.height=bodyHeight+'px';flow.style.columnWidth=bodyWidth+'px';
  for(const image of flow.querySelectorAll('img,svg')){image.style.maxWidth='100%';image.style.maxHeight=bodyHeight+'px';image.style.objectFit='contain';}
  const count=Math.max(1,Math.ceil((flow.scrollWidth-1)/bodyWidth));
  return{section,flow,width,height,bodyWidth,count};
}

async function openWordDocument(file){
  const lib=await wordReader();let zip;
  try{zip=await JSZip.loadAsync(await file.arrayBuffer());}catch{throw new Error('无法读取此 Word 文件，请用 Word/WPS 另存为 .docx 后重试。');}
  if(!zip.file('word/document.xml'))throw new Error('这不是有效的 .docx 文件，请用 Word/WPS 另存后重试。');
  // Resolve only resources embedded in the file; external relationships are not loaded.
  for(const entry of Object.values(zip.files).filter(f=>f.name.endsWith('.rels'))){
    const xml=new DOMParser().parseFromString(await entry.async('string'),'application/xml');
    const external=[...xml.getElementsByTagNameNS('*','Relationship')].filter(el=>el.getAttribute('TargetMode')?.toLowerCase()==='external');
    if(external.length){external.forEach(el=>el.remove());zip.file(entry.name,new XMLSerializer().serializeToString(xml));}
  }
  const host=document.createElement('iframe');host.title='Word 排版';host.setAttribute('aria-hidden','true');host.tabIndex=-1;
  host.setAttribute('sandbox','allow-same-origin');host.style.cssText='position:fixed;left:-20000px;top:0;width:1600px;height:1200px;border:0;pointer-events:none';
  const ready=new Promise(resolve=>host.onload=resolve);
  host.srcdoc='<!doctype html><html><head><meta charset="utf-8"><meta http-equiv="Content-Security-Policy" content="default-src \'none\'; style-src \'unsafe-inline\'; img-src data: blob:; font-src data: blob:"></head><body><main></main></body></html>';
  document.body.append(host);await ready;
  try{
    const doc=host.contentDocument,body=doc.querySelector('main');
    await lib.render(await zip.generateAsync({type:'uint8array'}),body,doc.head,{className:'glc-word',inWrapper:false,ignoreWidth:false,ignoreHeight:false,ignoreFonts:false,breakPages:true,ignoreLastRenderedPageBreak:false,useBase64URL:true,renderHeaders:true,renderFooters:true,renderFootnotes:true,renderEndnotes:true,renderAltChunks:false,renderComments:false,renderChanges:false});
    const rules=doc.createElement('style');rules.textContent=`
      html,body{margin:0;padding:0;background:white} section.glc-word{background:white;flex-shrink:0}
      section.glc-word>header,section.glc-word>footer{flex-shrink:0}
      .word-page-viewport{flex:1 1 0;min-height:0;overflow:hidden;position:relative}
      .word-page-flow{column-gap:0;column-fill:auto;position:relative;orphans:2;widows:2}
      .word-page-flow article{margin:0!important}.word-page-flow table{max-width:100%}
      .word-page-flow tr{break-inside:avoid}.word-page-flow img{max-width:100%}
    `;doc.head.append(rules);
    // docx-preview emits generic data MIME types. html2canvas only treats image/*
    // data URLs as inline images; retaining octet-stream would omit them on export.
    for(const img of doc.images)img.src=wordImageSource(img.src);
    for(const img of doc.querySelectorAll('svg image')){
      const href=img.getAttribute('href')||img.getAttributeNS('http://www.w3.org/1999/xlink','href');
      if(href)img.setAttribute('href',wordImageSource(href));
    }
    await Promise.all([...doc.images].map(img=>img.decode().catch(()=>{})));await doc.fonts.ready;
    const pages=[];
    for(const section of body.querySelectorAll('section.glc-word')){
      const layout=paginateWordSection(section,host.contentWindow);
      if(pages.length+layout.count>200)throw new Error('Word 超过 200 页，请拆分文件后导入。');
      for(let i=0;i<layout.count;i++)pages.push({...layout,column:i});
    }
    if(!pages.length)throw new Error('Word 中没有可显示的页面。');
    let closed=false,queue=Promise.resolve();
    return{name:file.name,word:true,count:pages.length,
      page(index,maxWidth=2400){
        const work=queue.then(async()=>{
          if(closed)throw new Error('文件已关闭，请重新导入。');
          const page=pages[index];if(!page)throw new Error('页码超出文件范围。');
          page.flow.style.transform=`translateX(${-page.column*page.bodyWidth}px)`;
          const canvas=await lib.capture(page.section,{backgroundColor:'#ffffff',scale:Math.min(3,maxWidth/page.width,Math.sqrt(14000000/(page.width*page.height))),width:page.width,height:page.height,logging:false,allowTaint:false,useCORS:false,scrollX:0,scrollY:0});
          return{canvas,paperSize:[page.width*.75,page.height*.75]};
        });queue=work.catch(()=>{});return work;
      },close(){closed=true;host.remove();}};
  }catch(error){host.remove();throw new Error('Word 打开失败：'+error.message);}
}
