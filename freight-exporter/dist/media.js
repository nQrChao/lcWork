let embeddedAssets,pdfReaderPromise,pdfWriterPromise;
const mediaURLs=[];
function assetData(name){embeddedAssets??=JSON.parse($('ocr-assets').textContent);const encoded=embeddedAssets[name];if(!encoded)throw new Error('离线组件缺失：'+name);return Uint8Array.from(atob(encoded),c=>c.charCodeAt(0));}
function assetText(name){return new TextDecoder().decode(assetData(name));}
function objectURL(bytes,type='text/javascript'){const url=URL.createObjectURL(new Blob([bytes],{type}));mediaURLs.push(url);return url;}
function loadEmbeddedScript(name){const script=document.createElement('script');script.textContent=assetText(name);document.head.append(script);script.remove();}
async function pdfReader(){return pdfReaderPromise??=(async()=>{const lib=await import(objectURL(assetData('pdf.min.mjs')));lib.GlobalWorkerOptions.workerSrc=objectURL(assetData('pdf.worker.min.mjs'));return lib;})();}
async function pdfWriter(){return pdfWriterPromise??=(async()=>{loadEmbeddedScript('pdf-lib.min.js');return window.PDFLib;})();}
async function imageCanvas(file,maxWidth=2600){const bitmap=await createImageBitmap(file);const scale=Math.min(2,Math.max(1,maxWidth/bitmap.width),Math.sqrt(14000000/(bitmap.width*bitmap.height)));const canvas=document.createElement('canvas');canvas.width=Math.round(bitmap.width*scale);canvas.height=Math.round(bitmap.height*scale);const ctx=canvas.getContext('2d');ctx.fillStyle='#fff';ctx.fillRect(0,0,canvas.width,canvas.height);ctx.drawImage(bitmap,0,0,canvas.width,canvas.height);bitmap.close();return canvas;}
async function openLocalDocument(file){
  if(file.size>60*1024*1024)throw new Error('请选择小于 60 MB 的文件。');
  if(/\.pdf$/i.test(file.name)||file.type==='application/pdf'){
    const bytes=new Uint8Array(await file.arrayBuffer()),lib=await pdfReader();const pdf=await lib.getDocument({data:bytes.slice(),isEvalSupported:false,useSystemFonts:true,disableFontFace:false,isOffscreenCanvasSupported:false}).promise;
    return{name:file.name,bytes,pdf,count:pdf.numPages,async page(index,maxWidth=2400){const p=await pdf.getPage(index+1),base=p.getViewport({scale:1}),viewport=p.getViewport({scale:Math.min(3,maxWidth/base.width)}),canvas=document.createElement('canvas');canvas.width=Math.ceil(viewport.width);canvas.height=Math.ceil(viewport.height);await p.render({canvasContext:canvas.getContext('2d'),viewport,background:'white'}).promise;return{canvas,pdfPage:p,base,viewport};},close(){pdf.destroy();}};
  }
  if(!/\.(png|jpe?g|webp|bmp)$/i.test(file.name))throw new Error('支持 JPG、PNG、WebP、BMP 图片和 PDF。');
  const canvas=await imageCanvas(file);return{name:file.name,count:1,async page(){const copy=document.createElement('canvas');copy.width=canvas.width;copy.height=canvas.height;copy.getContext('2d').drawImage(canvas,0,0);return{canvas:copy};},close(){}};
}
function canvasBytes(canvas,type='image/png',quality=.95){return new Promise((resolve,reject)=>canvas.toBlob(async b=>b?resolve(new Uint8Array(await b.arrayBuffer())):reject(new Error('图片生成失败。')),type,quality));}
function humanSize(n){return(n/1024/1024).toFixed(1)+' MB';}
