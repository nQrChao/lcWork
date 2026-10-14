let exportFormat='xlsx';
function updateExportFormat(){exportFormat=$('output-format').value;const labels={xlsx:[['一个 Excel，多个工作表','每个选中分组独立为一张工作表'],['每组一个 Excel','打包为 ZIP，一次下载'],['合并为一张表','只导出选中的筛选结果']],png:[['每组独立生成图片','长表自动分页，多张 PNG 打包下载'],['每组独立生成图片','打包为 ZIP，一次下载'],['合并生成图片','选中数据合并，长表自动分页']],pdf:[['一个 PDF，按组分页','每组从新的一页开始'],['每组一个 PDF','打包为 ZIP，一次下载'],['合并为一个 PDF','选中数据连续分页']]};document.querySelectorAll('.mode-option').forEach((el,i)=>{el.querySelector('strong').textContent=labels[exportFormat][i][0];el.querySelector('small').textContent=labels[exportFormat][i][1];el.classList.toggle('hidden',exportFormat==='png'&&i===1);});if(exportFormat==='png'&&mode==='zip'){mode='sheets';document.querySelector('[name=mode][value=sheets]').checked=true;}$('export-button').innerHTML='导出 '+exportFormat.toUpperCase()+' <span>↗</span>';renderData();}
async function* renderTablePages(groups,columns){
  const parts=[];if(columns.length<=10)parts.push(columns);else for(let i=3;i<columns.length;i+=7)parts.push([...columns.slice(0,3),...columns.slice(i,i+7)]);
  const width=2400,height=1697,margin=82,bodyTop=255,bottom=height-95,padding=12,fontSize=25;
  for(let gi=0;gi<groups.length;gi++)for(let pi=0;pi<parts.length;pi++){
    const part=parts[pi],group=groups[gi],weights=part.map(c=>Math.min(c.width||18,35)),total=weights.reduce((a,b)=>a+b,0),widths=weights.map(w=>w/total*(width-margin*2));let start=0,number=0;
    const measure=document.createElement('canvas').getContext('2d');measure.font=`${fontSize}px "Microsoft YaHei",sans-serif`;
    const wrap=(text,max)=>{const lines=[];for(const paragraph of text.split('\n')){let line='';for(const ch of paragraph){if(line&&measure.measureText(line+ch).width>max){lines.push(line);line='';}line+=ch;}lines.push(line);}return lines;};
    while(start<group.rows.length){
      const canvas=document.createElement('canvas');canvas.width=width;canvas.height=height;const ctx=canvas.getContext('2d');ctx.fillStyle='#fff';ctx.fillRect(0,0,width,height);ctx.textBaseline='middle';ctx.fillStyle='#173a5e';ctx.font='600 46px "Microsoft YaHei",sans-serif';ctx.fillText(group.name,margin,90,width-margin*2);ctx.fillStyle='#65758b';ctx.font='24px "Microsoft YaHei",sans-serif';ctx.fillText(`记录数：${group.rows.length}　${parts.length>1?'字段分段 '+(pi+1)+' / '+parts.length+'　':''}第 ${++number} 页`,margin,149);
      let x=margin;for(let i=0;i<part.length;i++){ctx.fillStyle='#254665';ctx.fillRect(x,195,widths[i],60);ctx.strokeStyle='#d6dee7';ctx.lineWidth=1;ctx.strokeRect(x,195,widths[i],60);ctx.fillStyle='#fff';ctx.font='600 24px "Microsoft YaHei",sans-serif';ctx.fillText(part[i].name,x+padding,225,widths[i]-padding*2);x+=widths[i];}
      let y=bodyTop;const initial=start;
      while(start<group.rows.length){const row=group.rows[start];const texts=part.map((c,i)=>{const cell=row.cells[columns.indexOf(c)],value=cell?.t==='n'?cell.v.toLocaleString('zh-CN',{minimumFractionDigits:3,maximumFractionDigits:3}):display(cell);return wrap(value,widths[i]-padding*2);});const rowHeight=Math.max(55,...texts.map(lines=>lines.length*34+padding*2));if(y+rowHeight>bottom&&start>initial)break;
        if(rowHeight>bottom-bodyTop)throw new Error('有单元格内容过长，请减少备注长度或使用 Excel 导出。');
        x=margin;for(let i=0;i<part.length;i++){ctx.strokeStyle='#bbc7d4';ctx.strokeRect(x,y,widths[i],rowHeight);ctx.fillStyle='#263549';ctx.font=`${fontSize}px "Microsoft YaHei",sans-serif`;ctx.textAlign=part[i].type==='number'?'right':'left';texts[i].forEach((line,k)=>ctx.fillText(line,x+(ctx.textAlign==='right'?widths[i]-padding:padding),y+padding+17+k*34));x+=widths[i];}ctx.textAlign='left';y+=rowHeight;start++;
      }
      ctx.fillStyle='#7d8a80';ctx.font='20px "Microsoft YaHei",sans-serif';ctx.fillText(`${initial+1}–${start} / ${group.rows.length}`,margin,height-45);let output=canvas;if(exportFormat==='png'&&y+100<height){output=document.createElement('canvas');output.width=width;output.height=Math.ceil(y+100);const out=output.getContext('2d');out.drawImage(canvas,0,0);out.fillStyle='#7d8a80';out.font='20px \"Microsoft YaHei\",sans-serif';out.fillText(`${initial+1}–${start} / ${group.rows.length}`,margin,output.height-30);}yield{canvas:output,group:gi,segment:pi,page:number,name:group.name};await new Promise(r=>setTimeout(r,0));
    }
  }
}
async function exportVisual(){
  if(!active)return;const data=dataset(),columns=active.columns.filter(c=>exportColumns.has(c.id)),indexes=columns.map(c=>active.columns.indexOf(c));if(!data.exported.length||!columns.length)return;
  recalculateAll(active);const project=rows=>rows.map(r=>({...r,cells:indexes.map(i=>r.cells[i])}));const groups=mode==='single'?[{name:active.name,rows:project(data.exported)}]:data.chosen.map(g=>({name:g.name+' · '+active.name,rows:project(g.rows)}));
  busy(true,'正在排版 '+exportFormat.toUpperCase());try{
    const zip=new JSZip(),used=new Set(),base=safeName(active.name,new Set(),90),names=groups.map(g=>safeName(g.name,used,90));let pngCount=0,onlyPNG=null,onlyName='';const lib=exportFormat==='pdf'?await pdfWriter():null,pdfs=[];
    if(lib){if(mode==='zip')for(const g of groups)pdfs.push(await lib.PDFDocument.create());else pdfs.push(await lib.PDFDocument.create());}
    for await(const result of renderTablePages(groups,columns)){
      const bytes=await canvasBytes(result.canvas);if(exportFormat==='png'){const name=names[result.group]+'_字段'+(result.segment+1)+'_第'+result.page+'页.png';zip.file(name,bytes);pngCount++;onlyPNG=bytes;onlyName=name;}
      else{const pdf=pdfs[mode==='zip'?result.group:0],image=await pdf.embedPng(bytes),page=pdf.addPage([841.89,595.28]);page.drawImage(image,{x:0,y:0,width:841.89,height:595.28});}
      $('loading').querySelector('strong').textContent='正在生成 '+result.name+' · 第 '+result.page+' 页';result.canvas.width=1;result.canvas.height=1;
    }
    if(exportFormat==='png'){if(pngCount===1)download(onlyPNG,onlyName,'image/png');else download(await zip.generateAsync({type:'uint8array'}),base+'_图片.zip','application/zip');}
    else if(mode==='zip'){for(let i=0;i<pdfs.length;i++)zip.file(names[i]+'.pdf',await pdfs[i].save());download(await zip.generateAsync({type:'uint8array'}),base+'_PDF.zip','application/zip');}else download(await pdfs[0].save(),base+'_表格.pdf','application/pdf');
    toast('已生成 '+exportFormat.toUpperCase()+' 文件。');
  }catch(e){toast(e.message,true);}finally{busy(false);}
}
$('output-format').onchange=updateExportFormat;$('export-button').onclick=()=>(exportFormat==='xlsx'?exportData():exportVisual());$('export-top').onclick=$('export-button').onclick;
