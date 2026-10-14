import {uid,emptyCell,safeName,display,value,parseCell} from './core.js';
const MAIN='http://schemas.openxmlformats.org/spreadsheetml/2006/main';
const XML_HEAD='<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';
const xmlEscape=v=>String(v??'').replace(/[\x00-\x08\x0b\x0c\x0e-\x1f]/g,'').replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&apos;'}[c]));
const colName=n=>{let s='';for(n++;n;n=Math.floor((n-1)/26))s=String.fromCharCode(65+(n-1)%26)+s;return s;};
const cellPosition=ref=>{const m=String(ref).replaceAll('$','').match(/^([A-Z]+)(\d+)$/);if(!m)return[0,0];return[Number(m[2])-1,[...m[1]].reduce((n,c)=>n*26+c.charCodeAt(0)-64,0)-1];};
const children=(node,tag)=>[...node.getElementsByTagNameNS('*',tag)];
function parseXML(text){const doc=new DOMParser().parseFromString(text,'application/xml');if(children(doc,'parsererror').length)throw new Error('文件内容无法解析，请重新保存为 .xlsx 后导入。');return doc;}
const firstText=(node,tag)=>children(node,tag)[0]?.textContent??'';
const isoToExcel=iso=>(Date.parse(iso+'T00:00:00Z')-Date.UTC(1899,11,30))/86400000;
const excelToISO=(n,is1904=false)=>new Date(Date.UTC(is1904?1904:1899,is1904?0:11,is1904?1:30)+n*86400000).toISOString().slice(0,10);
function resolvePath(base,target){if(target.startsWith('/'))return target.slice(1);const parts=(base.slice(0,base.lastIndexOf('/')+1)+target).split('/');const out=[];for(const p of parts){if(p==='..')out.pop();else if(p&&p!=='.')out.push(p);}return out.join('/');}
const relsPath=p=>p.slice(0,p.lastIndexOf('/')+1)+'_rels/'+p.slice(p.lastIndexOf('/')+1)+'.rels';

export async function readWorkbook(buffer){
  let zip;try{zip=await JSZip.loadAsync(buffer);}catch{throw new Error('无法打开此文件。请选择未加密的 .xlsx 工作簿。');}
  const read=async path=>{const f=zip.file(path);return f?await f.async('string'):null;};
  let metadata=null;const custom=await read('docProps/custom.xml');if(custom){try{const prop=children(parseXML(custom),'property').find(p=>p.getAttribute('name')==='SheetWorkbench');metadata=prop?JSON.parse(prop.textContent):null;}catch{}}
  if(metadata){const validID=id=>typeof id==='string'&&/^[a-zA-Z0-9_-]{1,100}$/.test(id);if(metadata.version!==1||!Array.isArray(metadata.columns)||metadata.columns.some(c=>!validID(c.id)||typeof c.name!=='string'||!['text','number','date'].includes(c.type))){metadata=null;}else{metadata.columns=metadata.columns.map(c=>{const out={id:c.id,name:c.name,type:c.type};if(Number.isFinite(c.width)&&c.width>=5&&c.width<=80)out.width=c.width;for(const key of ['computed','newRowComputed']){const calc=c[key];if(calc&&['add','multiply'].includes(calc.op)&&Array.isArray(calc.fields)&&calc.fields.length===2&&calc.fields.every(validID))out[key]={op:calc.op,fields:calc.fields};}return out;});}}
  const wbText=await read('xl/workbook.xml');if(!wbText)throw new Error('这不是有效的 .xlsx 工作簿。');
  const wb=parseXML(wbText),rels=parseXML(await read('xl/_rels/workbook.xml.rels'));
  const relationships=new Map(children(rels,'Relationship').map(e=>[e.getAttribute('Id'),e.getAttribute('Target')]));
  const is1904=children(wb,'workbookPr')[0]?.getAttribute('date1904')==='1';
  const sharedText=await read('xl/sharedStrings.xml');
  const strings=sharedText?children(parseXML(sharedText),'si').map(e=>children(e,'t').map(t=>t.textContent).join('')):[];
  const stylesText=await read('xl/styles.xml');
  const formats=new Map([[0,'General'],[1,'0'],[2,'0.00'],[3,'#,##0'],[4,'#,##0.00'],[9,'0%'],[10,'0.00%']]);
  let styles=[];
  if(stylesText){const doc=parseXML(stylesText);for(const e of children(doc,'numFmt'))formats.set(Number(e.getAttribute('numFmtId')),e.getAttribute('formatCode'));const xfs=children(doc,'cellXfs')[0];styles=xfs?[...xfs.children].map(e=>Number(e.getAttribute('numFmtId'))):[];}
  const isDateFormat=id=>(id>=14&&id<=22)||(id>=27&&id<=36)||(id>=45&&id<=47)||(id>=50&&id<=58)||/[ymd]/i.test((formats.get(id)||'').replace(/"[^"]*"|\[[^\]]*\]|\\./g,''));
  const sheets=children(wb,'sheet').map(e=>({name:e.getAttribute('name'),path:resolvePath('xl/workbook.xml',relationships.get(e.getAttribute('r:id')||e.getAttributeNS('http://schemas.openxmlformats.org/officeDocument/2006/relationships','id')))}));
  if(!sheets.length)throw new Error('工作簿中没有工作表。');
  async function loadSheet(index){
    const meta=sheets[index],doc=parseXML(await read(meta.path)),rows=[],sharedFormulas=new Map();let formulaCount=0,missingFormulaResults=0;
    for(const rowNode of children(doc,'row')){
      const r=Number(rowNode.getAttribute('r'))-1;if(r>100000)throw new Error('此工作表超过 100,000 行，请先缩小范围。');
      const row=[];
      for(const c of children(rowNode,'c')){
        const [,col]=cellPosition(c.getAttribute('r'));if(col>=256)throw new Error('此工作表超过 256 列，请先移除无关列。');
        const t=c.getAttribute('t'),raw=firstText(c,'v'),fmtId=styles[Number(c.getAttribute('s')||0)]||0,z=formats.get(fmtId)||'General';
        let cell=emptyCell();
        if(children(c,'f').length){formulaCount++;if(!children(c,'v').length||raw==='')missingFormulaResults++;}
        if(t==='s')cell={v:strings[Number(raw)]??'',t:'s'};
        else if(t==='inlineStr')cell={v:children(c,'t').map(e=>e.textContent).join(''),t:'s'};
        else if(t==='str')cell={v:raw,t:'s'};
        else if(t==='e')cell={v:raw,t:'e'};
        else if(t==='b')cell={v:raw==='1',t:'b'};
        else if(t==='d')cell={v:raw.slice(0,10),t:'d'};
        else if(raw!==''){const n=Number(raw);cell=Number.isFinite(n)?(isDateFormat(fmtId)?{v:excelToISO(n,is1904),t:'d'}:{v:n,t:'n',z}):{v:raw,t:'s'};}
        if(children(c,'f').length){const f=children(c,'f')[0];cell.rawFormula=f.textContent;cell.missingFormulaResult=!children(c,'v').length||(raw===''&&t!=='str');if(f.getAttribute('t')==='shared'){const si=f.getAttribute('si');if(cell.rawFormula)sharedFormulas.set(si,{expr:cell.rawFormula,row:r,col});else cell.sharedFormula={si,row:r,col};}}
        row[col]=cell;
      }
      rows[r]=row;
    }
    for(const row of rows)for(const c of row||[]){if(!c?.sharedFormula)continue;const current=c.sharedFormula,base=sharedFormulas.get(current.si);if(base)c.rawFormula=base.expr.replace(/(\$?)([A-Z]{1,3})(\$?)(\d+)/g,(match,ac,letters,ar,r)=>{const col=cellPosition(letters+'1')[1]+(ac?0:current.col-base.col),rn=Number(r)+(ar?0:current.row-base.row);return col<0||rn<1?'#REF!':ac+colName(col)+ar+rn;});else c.rawFormula='UNRESOLVED_SHARED_FORMULA';delete c.sharedFormula;}
    // Shared financial values remain on their original row.

    const merged=children(doc,'mergeCell').map(e=>e.getAttribute('ref'));
    let tableRange=null;
    const relText=await read(relsPath(meta.path));
    if(relText){const rs=children(parseXML(relText),'Relationship');for(const rel of rs){if(rel.getAttribute('Type')?.endsWith('/table')){const tx=await read(resolvePath(meta.path,rel.getAttribute('Target')));if(tx){const t=parseXML(tx).documentElement;tableRange={ref:t.getAttribute('ref'),totals:Number(t.getAttribute('totalsRowCount')||0)};break;}}}}
    let headerIndex=0;
    if(tableRange)headerIndex=cellPosition(tableRange.ref.split(':')[0])[0];
    else{let best=-1;for(let r=0;r<Math.min(40,rows.length);r++){const cells=(rows[r]||[]).filter(c=>value(c)!==null&&display(c)!=='');const textCells=cells.filter(c=>c.t==='s').length;const known=cells.filter(c=>/^(货运公司|建单日期|车号|单价|金额|名称|日期|备注|到货单位|发货单位)$/.test(display(c))).length;const score=textCells*2+known*3-(cells.length-textCells)*2;if(score>best){best=score;headerIndex=r;}}}
    return{rows,merged,headerIndex,tableRange,formulaCount,missingFormulaResults,name:meta.name,metadata};
  }
  return{sheets,loadSheet};
}
export function sheetToTable(sheet,name,headerIndex=sheet.headerIndex,{includeAppendedRows=false,includeExtraColumns=false,skipFormulaOnlyRows=false}={}){
  if(!Number.isInteger(headerIndex)||headerIndex<0||headerIndex>=sheet.rows.length)throw new Error('请选择有效的表头行。');
  const sourceHeaders=sheet.rows[headerIndex]||[];let last=sourceHeaders.length-1;while(last>=0&&!display(sourceHeaders[last]).trim())last--;
  if(last<0)throw new Error('这一行没有列名，请选择正确的表头行。');
  if(includeExtraColumns)for(const row of sheet.rows.slice(headerIndex+1))for(let i=last+1;i<(row?.length||0);i++)if(value(row[i])!==null&&display(row[i])!==''||row[i]?.rawFormula!==undefined)last=i;
  const names=new Set();
  const columns=Array.from({length:last+1},(_,i)=>{const base=display(sourceHeaders[i]).trim()||`未命名列 ${i+1}`;let name=base,n=2;while(names.has(name))name=`${base} (${n++})`;names.add(name);return{id:uid(),name,type:/日期|时间/.test(name)?'date':'text'};});
  let end=sheet.rows.length-1;
  if(sheet.tableRange&&headerIndex===sheet.headerIndex)end=cellPosition(sheet.tableRange.ref.split(':')[1]||sheet.tableRange.ref)[0]-sheet.tableRange.totals;
  const tableEnd=end;if(sheet.metadata||includeAppendedRows)for(let r=end+1;r<sheet.rows.length;r++)if((sheet.rows[r]||[]).some(c=>value(c)!==null&&display(c)!==''||c?.rawFormula))end=r;
  const records=[];let skippedFormulaRows=0;
  for(let r=headerIndex+1;r<=end;r++){
    const cells=Array.from({length:columns.length},(_,i)=>structuredClone(sheet.rows[r]?.[i]||emptyCell()));
    if(skipFormulaOnlyRows&&cells.some(c=>c.rawFormula!==undefined)&&cells.every(c=>c.rawFormula!==undefined?!c.missingFormulaResult:value(c)===null||display(c).trim()==='')){skippedFormulaRows++;continue;}
    if((!sheet.metadata||r>tableEnd)&&cells.every(c=>!c?.rawFormula&&(value(c)===null||display(c)==='')))continue;
    if(!sheet.metadata&&cells.every(c=>c.rawFormula===undefined&&(value(c)===null||typeof value(c)==='string'&&/^(合计|小计|总计)[：:]?$/.test(value(c)))))continue;
    records.push({id:uid(),cells,sourceRow:r+1});
  }
  for(let i=0;i<columns.length;i++){if(columns[i].type==='date')continue;const vals=records.map(r=>r.cells[i]).filter(c=>value(c)!==null);if(vals.length&&vals.every(c=>c.t==='n'||c.t==='e'))columns[i].type='number';}
  if(sheet.metadata?.columns)for(let i=0;i<columns.length;i++){const col=columns[i],name=col.name,saved=sheet.metadata.columns.find(c=>c.name===name)||(sheet.metadata.columns.length===columns.length?sheet.metadata.columns[i]:null);if(saved)Object.assign(col,saved,{name});}
  const rowMap=new Map(records.map(r=>[r.sourceRow,r.id]));let unsupported=0;
  for(const row of records)for(let i=0;i<row.cells.length;i++){
    const cell=row.cells[i];if(cell.rawFormula===undefined){if(columns[i].computed)cell.manual=true;continue;}
    if(columns[i].computed){const calc=columns[i].computed,indices=calc.fields.map(id=>columns.findIndex(c=>c.id===id));const [a,b]=indices.map(j=>colName(j)+row.sourceRow);const expected=calc.op==='multiply'?`IF(OR(${a}="",${b}=""),"",${a}*${b})`:`IF(${a}="","",${a}+${b})`;if(cell.rawFormula.replace(/\s/g,'').toUpperCase()===expected.toUpperCase()){delete cell.rawFormula;continue;}cell.manual=true;}
    const refs=[];let unresolved=false;
    const expr=cell.rawFormula.replace(/(?:('(?:[^']|'')+'|[\w\u3400-\u9fff]+)!)?(\$?)([A-Z]{1,3})(\$?)(\d+)/g,(match,sheetName,ac,letters,ar,r)=>{
      const decoded=sheetName?.replace(/^'|'$/g,'').replaceAll("''", "'");
      const j=cellPosition(letters+'1')[1],id=rowMap.get(Number(r));
      if((decoded&&decoded!==sheet.name)||!id||!columns[j]){unresolved=true;return match;}
      refs.push({row:id,col:columns[j].id,absCol:!!ac,absRow:!!ar});return '@'+(refs.length-1)+'@';
    });
    if(!unresolved&&/^(?:@\d+@|[\d.()+*/\s-]|#REF!)+$/.test(expr)){cell.formula={expr,refs};delete cell.rawFormula;}else{unsupported++;}
  }
  return{id:uid(),name,columns,rows:records,unsupported,skippedFormulaRows,updatedAt:Date.now()};
}

// Align by names, and remap both axes of formula references for each source.
// Workbooks saved by this app may reuse column IDs across different sheets.
export function mergeTables(parts,name,{sourceField=''}={}){
  if(!parts.length)throw new Error('请至少选择一张工作表。');
  const columns=[],byName=new Map(),maps=[];
  function addColumn(source){let col=byName.get(source.name);if(!col){col={id:uid(),name:source.name,type:source.type||'text'};if(source.width)col.width=source.width;columns.push(col);byName.set(col.name,col);}else if(col.type!==source.type)col.type='text';return col;}
  if(sourceField)addColumn({name:sourceField,type:'text'});
  for(const {table} of parts){const cols=new Map();for(const col of table.columns)cols.set(col.id,addColumn(col).id);maps.push({cols,rows:new Map(table.rows.map(row=>[row.id,uid()]))});}
  for(let p=0;p<parts.length;p++)for(const col of parts[p].table.columns){
    const dest=byName.get(col.name);
    for(const key of ['computed','newRowComputed'])if(col[key]){
      const calc={op:col[key].op,fields:col[key].fields.map(id=>maps[p].cols.get(id))};
      if(calc.fields.some(id=>!id))throw new Error(`字段“${col.name}”的计算规则缺少引用字段。`);
      if(dest[key]&&JSON.stringify(dest[key])!==JSON.stringify(calc))throw new Error(`字段“${col.name}”在各分表中的计算规则不同，请分别导入。`);
      dest[key]=calc;
    }
  }
  const rows=[],positions=new Map(columns.map((c,i)=>[c.id,i]));let unsupported=0;
  for(let p=0;p<parts.length;p++){
    const {table,sheetName}=parts[p],map=maps[p];
    for(const source of table.rows){
      const cells=columns.map(col=>({...emptyCell(),...(col.computed?{manual:true}:{})}));
      for(let i=0;i<table.columns.length;i++){
        const col=table.columns[i],j=positions.get(map.cols.get(col.id)),cell=structuredClone(source.cells[i]||emptyCell());
        if(columns[j].computed&&!col.computed)cell.manual=true;
        if(cell.formula)cell.formula.refs=cell.formula.refs.map(ref=>({...ref,row:map.rows.get(ref.row)||ref.row,col:map.cols.get(ref.col)||ref.col}));
        if(cell.rawFormula!==undefined)unsupported++;
        cells[j]=cell;
      }
      if(sourceField){const j=positions.get(byName.get(sourceField).id);if(value(cells[j])===null||display(cells[j]).trim()==='')cells[j]={v:sheetName,t:'s',manual:true};}
      rows.push({id:map.rows.get(source.id),cells,sourceRow:source.sourceRow});
    }
  }
  if(sourceField)byName.get(sourceField).type='text';
  return{id:uid(),name,columns,rows,unsupported,updatedAt:Date.now()};
}

export async function writeWorkbook(groups,columns,options={}){
  if(!groups.length||!columns.length)throw new Error('请选择需要导出的数据和列。');
  const zip=new JSZip(),used=new Set();
  const names=groups.map(g=>safeName(g.name,used));
  const formats=new Map([['yyyy-mm-dd',164],['0.000',165],['#,##0.000',166]]);
  for(const g of groups)for(const row of g.rows)for(const c of row.cells){if(c?.t==='n'&&c.z&&c.z!=='General'&&!formats.has(c.z))formats.set(c.z,164+formats.size);}
  const dateStyle=5,errorStyle=4;
  const customStyle=new Map([...formats.keys()].map((f,i)=>[f,5+i]));
  const xf=(numFmtId=0,fontId=0,fillId=0,align='left')=>`<xf numFmtId="${numFmtId}" fontId="${fontId}" fillId="${fillId}" borderId="1" xfId="0" applyBorder="1" applyAlignment="1" applyNumberFormat="1"><alignment vertical="center" horizontal="${align}" wrapText="1"/></xf>`;
  zip.file('xl/styles.xml',XML_HEAD+`<styleSheet xmlns="${MAIN}"><numFmts count="${formats.size}">${[...formats].map(([f,id])=>`<numFmt numFmtId="${id}" formatCode="${xmlEscape(f)}"/>`).join('')}</numFmts><fonts count="5"><font><sz val="11"/><color rgb="FF263549"/><name val="Microsoft YaHei"/></font><font><b/><sz val="11"/><color rgb="FFFFFFFF"/><name val="Microsoft YaHei"/></font><font><sz val="11"/><color rgb="FFB42318"/><name val="Microsoft YaHei"/></font><font><b/><sz val="18"/><color rgb="FF173A5E"/><name val="Microsoft YaHei"/></font><font><sz val="10"/><color rgb="FF65758B"/><name val="Microsoft YaHei"/></font></fonts><fills count="3"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill><fill><patternFill patternType="solid"><fgColor rgb="FF254665"/><bgColor indexed="64"/></patternFill></fill></fills><borders count="2"><border><left/><right/><top/><bottom/><diagonal/></border><border>${['left','right','top','bottom'].map(s=>`<${s} style="thin"><color rgb="FFBBC7D4"/></${s}>`).join('')}<diagonal/></border></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="${7+formats.size}">${xf()}${xf(0,1,2,'center')}${xf(0,0,0,'right')}${xf(14)}${xf(0,2)}${[...formats].map(([,id])=>xf(id,0,0,'right')).join('')}<xf numFmtId="0" fontId="3" fillId="0" borderId="0" xfId="0" applyAlignment="1"><alignment vertical="center"/></xf><xf numFmtId="0" fontId="4" fillId="0" borderId="0" xfId="0" applyAlignment="1"><alignment vertical="center"/></xf></cellXfs><cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles></styleSheet>`);
  const contentTypes=[['/xl/workbook.xml','application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml'],['/xl/styles.xml','application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml']];
  const relType='http://schemas.openxmlformats.org/officeDocument/2006/relationships';
  const relRoot='http://schemas.openxmlformats.org/package/2006/relationships';
  const bodyCell=(c,address,formula='')=>{
    if(formula){const t=c?.t==='e'?'e':c?.t==='b'?'b':c?.t==='s'||value(c)===null?'str':null;return `<c r="${address}" s="${c?.t==='e'?errorStyle:customStyle.get(c?.z)??2}"${t?` t="${t}"`:''}><f>${xmlEscape(formula)}</f><v>${xmlEscape(c?.t==='b'?(c.v?1:0):c?.v??'')}</v></c>`;}
    if(!c||value(c)===null)return`<c r="${address}" s="0"/>`;
    if(c.t==='n'&&Number.isFinite(c.v))return`<c r="${address}" s="${customStyle.get(c.z)??2}"><v>${c.v}</v></c>`;
    if(c.t==='d'&&Number.isFinite(isoToExcel(c.v)))return`<c r="${address}" s="${dateStyle}"><v>${isoToExcel(c.v)}</v></c>`;
    if(c.t==='b')return`<c r="${address}" s="0" t="b"><v>${c.v?1:0}</v></c>`;
    if(c.t==='e')return`<c r="${address}" s="${errorStyle}" t="e"><v>${xmlEscape(c.v)}</v></c>`;
    return`<c r="${address}" s="0" t="inlineStr"><is><t xml:space="preserve">${xmlEscape(c.v)}</t></is></c>`;
  };
  for(let gi=0;gi<groups.length;gi++){
    const g=groups[gi],id=gi+1,lastColumn=colName(columns.length-1),lastRow=g.rows.length+5;
    const rowLookup=new Map(g.rows.map((r,i)=>[r.id,i+6])),colLookup=new Map(columns.map((c,i)=>[c.id,i]));
    function formulaFor(row,j,rowNumber){
      if(!options.formulas||(row.cells[j]?.manual&&!row.cells[j]?.formula))return '';
      const c=row.cells[j],calc=c?.manual?null:columns[j].computed;
      if(calc){const indexes=calc.fields.map(id=>colLookup.get(id));if(indexes.some(i=>i===undefined))return '';const [a,b]=indexes.map(i=>colName(i)+rowNumber);return calc.op==='multiply'?`IF(OR(${a}="",${b}=""),"",${a}*${b})`:`IF(${a}="","",${a}+${b})`;}
      if(c?.formula){let valid=true;const expr=c.formula.expr.replace(/@(\d+)@/g,(_,i)=>{const ref=c.formula.refs[Number(i)],r=rowLookup.get(ref.row),j=colLookup.get(ref.col);if(!r||j===undefined){valid=false;return '#REF!';}return(ref.absCol?'$':'')+colName(j)+(ref.absRow?'$':'')+r;});return valid?expr:expr.includes('#REF!')?expr:'';}
      return '';
    }
    const widths=columns.map((c,i)=>{let max=Math.max([...c.name].length*2+4,12);for(const r of g.rows.slice(0,100)){const s=display(r.cells[i]);max=Math.max(max,Math.min(48,[...s].reduce((n,ch)=>n+(ch.charCodeAt(0)>255?2:1),0)+2));}return c.width||Math.min(48,max);});
    const header=`<row r="5" ht="30" customHeight="1">${columns.map((c,i)=>`<c r="${colName(i)}5" s="1" t="inlineStr"><is><t>${xmlEscape(c.name)}</t></is></c>`).join('')}</row>`;
    const body=g.rows.map((row,i)=>{let lines=1;row.cells.forEach((c,j)=>{const s=display(c);lines=Math.max(lines,s.split('\n').reduce((sum,line)=>sum+Math.max(1,Math.ceil([...line].reduce((n,ch)=>n+(ch.charCodeAt(0)>255?2:1),0)/(widths[j]-2))),0));});return`<row r="${i+6}" ht="${Math.max(26,lines*15+6)}" customHeight="1">${columns.map((c,j)=>bodyCell(row.cells[j],`${colName(j)}${i+6}`,formulaFor(row,j,i+6))).join('')}</row>`;}).join('');
    zip.file(`xl/worksheets/sheet${id}.xml`,XML_HEAD+`<worksheet xmlns="${MAIN}" xmlns:r="${relType}"><dimension ref="A1:${lastColumn}${lastRow}"/><sheetViews><sheetView workbookViewId="0" showGridLines="0"><pane xSplit="${Math.min(3,columns.length)}" ySplit="5" topLeftCell="${colName(Math.min(3,columns.length))}6" activePane="bottomRight" state="frozen"/></sheetView></sheetViews><sheetFormatPr defaultRowHeight="24"/><cols>${widths.map((width,i)=>`<col min="${i+1}" max="${i+1}" width="${width}" customWidth="1"/>`).join('')}</cols><sheetData><row r="1" ht="20" customHeight="1"><c r="A1" s="${5+formats.size}" t="inlineStr"><is><t>${xmlEscape(g.title||g.name)}</t></is></c></row><row r="2" ht="20" customHeight="1"/><row r="3" ht="24" customHeight="1"><c r="A3" s="${6+formats.size}" t="str"><f>"记录数："&amp;MAX(0,ROWS(Data${id}[#All])-1)</f><v>记录数：${g.rows.length}</v></c></row><row r="4" ht="8" customHeight="1"/>${header}${body}</sheetData><autoFilter ref="A5:${lastColumn}${lastRow}"/><mergeCells count="${columns.length>1?2:1}"><mergeCell ref="A1:${colName(Math.min(5,columns.length-1))}2"/>${columns.length>1?`<mergeCell ref="A3:${colName(Math.min(2,columns.length-1))}3"/>`:""}</mergeCells><tableParts count="1"><tablePart r:id="rId1"/></tableParts></worksheet>`);
    zip.file(`xl/tables/table${id}.xml`,XML_HEAD+`<table xmlns="${MAIN}" id="${id}" name="Data${id}" displayName="Data${id}" ref="A5:${lastColumn}${lastRow}" totalsRowShown="0"><autoFilter ref="A5:${lastColumn}${lastRow}"/><tableColumns count="${columns.length}">${columns.map((c,i)=>`<tableColumn id="${i+1}" name="${xmlEscape(c.name)}"/>`).join('')}</tableColumns><tableStyleInfo name="TableStyleLight1" showFirstColumn="0" showLastColumn="0" showRowStripes="0" showColumnStripes="0"/></table>`);
    zip.file(`xl/worksheets/_rels/sheet${id}.xml.rels`,XML_HEAD+`<Relationships xmlns="${relRoot}"><Relationship Id="rId1" Type="${relType}/table" Target="../tables/table${id}.xml"/></Relationships>`);
    contentTypes.push([`/xl/worksheets/sheet${id}.xml`,'application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml'],[`/xl/tables/table${id}.xml`,'application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml']);
  }
  zip.file('xl/workbook.xml',XML_HEAD+`<workbook xmlns="${MAIN}" xmlns:r="${relType}"><bookViews><workbookView activeTab="0"/></bookViews><sheets>${names.map((name,i)=>`<sheet name="${xmlEscape(name)}" sheetId="${i+1}" r:id="rId${i+1}"/>`).join('')}</sheets><calcPr calcId="191029" fullCalcOnLoad="1"/></workbook>`);
  zip.file('xl/_rels/workbook.xml.rels',XML_HEAD+`<Relationships xmlns="${relRoot}">${names.map((_,i)=>`<Relationship Id="rId${i+1}" Type="${relType}/worksheet" Target="worksheets/sheet${i+1}.xml"/>`).join('')}<Relationship Id="styles" Type="${relType}/styles" Target="styles.xml"/></Relationships>`);
  let customRel='';
  if(options.metadata){zip.file('docProps/custom.xml',XML_HEAD+`<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/custom-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"><property fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}" pid="2" name="SheetWorkbench"><vt:lpwstr>${xmlEscape(JSON.stringify({version:1,columns}))}</vt:lpwstr></property></Properties>`);contentTypes.push(['/docProps/custom.xml','application/vnd.openxmlformats-officedocument.custom-properties+xml']);customRel=`<Relationship Id="custom" Type="${relType}/custom-properties" Target="docProps/custom.xml"/>`;}
  zip.file('_rels/.rels',XML_HEAD+`<Relationships xmlns="${relRoot}"><Relationship Id="rId1" Type="${relType}/officeDocument" Target="xl/workbook.xml"/>${customRel}</Relationships>`);
  zip.file('[Content_Types].xml',XML_HEAD+`<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>${contentTypes.map(([PartName,ContentType])=>`<Override PartName="${PartName}" ContentType="${ContentType}"/>`).join('')}</Types>`);
  return zip.generateAsync({type:'uint8array',compression:'DEFLATE',compressionOptions:{level:6}});
}
