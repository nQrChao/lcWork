export const uid=()=>globalThis.crypto?.randomUUID?.()||`id-${Date.now()}-${Math.random().toString(36).slice(2)}`;
export const emptyCell=()=>({v:null,t:'s'});
export const value=c=>c?.v??null;
export const display=c=>value(c)===null?'':String(value(c));
export const isBlank=c=>value(c)===null||String(value(c)).trim()==='';
export const toNumber=v=>typeof v==='number'&&Number.isFinite(v)?v:(typeof v==='string'&&/^[+-]?(?:\d+\.?\d*|\.\d+)$/.test(v.trim())?Number(v):null);
export const numberFormat=n=>Number(n).toLocaleString('zh-CN');
export const escapeHTML=v=>String(v??'').replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
export function normalizeDirectoryPath(input){
  let path=String(input||'').trim().replace(/^(["'])(.*)\1$/,'$2');
  if(/[\r\n\0]/.test(path))throw new Error('目录路径不能包含换行。');
  if(/^[a-z]:[\\/]/i.test(path)){path=path.replaceAll('/','\\').replace(/\\+$/,'');return /^[a-z]:$/i.test(path)?path+'\\':path;}
  if(/^(?:\\\\|\/\/)[^\\/]+[\\/][^\\/]+/.test(path))return path.replaceAll('/','\\').replace(/\\+$/,'');
  if(path.startsWith('/')&&!path.startsWith('//'))return path.replace(/\/+$/,'')||'/';
  throw new Error('请输入完整目录路径，例如 C:\\工作表 或 /Users/姓名/Documents。');
}
export function joinLocalPath(directory,name){const base=normalizeDirectoryPath(directory),sep=base.startsWith('/')?'/':'\\';return base.endsWith(sep)?base+name:base+sep+name;}
export function parseCell(text,type='text'){
  if(text===''||text===null||text===undefined)return emptyCell();
  if(type==='number') {const n=toNumber(String(text).replaceAll(',','')); return n===null?{v:String(text),t:'s'}:{v:n,t:'n',z:'#,##0.000'};}
  if(type==='date'){
    const m=String(text).trim().match(/^(\d{4})[.\-/年](\d{1,2})[.\-/月](\d{1,2})日?$/);
    if(m){const iso=`${m[1]}-${m[2].padStart(2,'0')}-${m[3].padStart(2,'0')}`;const d=new Date(iso+'T00:00:00Z');if(Number.isFinite(d.getTime())&&d.toISOString().slice(0,10)===iso)return{v:iso,t:'d'};}
  }
  return{v:String(text),t:'s'};
}
export const freightFields=['货运公司','建单日期','车号','始发地','目的地','卸货日期','吨位','单价','金额','其它费用','备注','总金额','到货单位','结算日期','支付方式','支付日期','发货单位','含税金额','发票情况','付款时间','付款金额','余额'];
export function createTable(name,preset='freight'){
  const names=freightFields;
  const widths=[17.22,12.22,13.33,13.89,14.44,12.44,12.22,12.22,13.89,19.44,40,17.22,27.22,15.56,14.44,14.67,12.78,14.44,14.44,14.67,15,13.89];
  const columns=names.map((name,i)=>({id:uid(),name,width:widths[i],type:/日期|时间/.test(name)?'date':/吨位|单价|金额|费用|余额|数量/.test(name)?'number':'text'}));
  {
    columns.find(c=>c.name==='金额').computed={op:'multiply',fields:['吨位','单价'].map(n=>columns.find(c=>c.name===n).id)};
    columns.find(c=>c.name==='总金额').computed={op:'add',fields:['金额','其它费用'].map(n=>columns.find(c=>c.name===n).id)};
  }
  return{id:uid(),name,columns,rows:[],updatedAt:Date.now()};
}
export function recalculate(table,row){
  for(let pass=0;pass<2;pass++)for(let i=0;i<table.columns.length;i++){
    const calc=table.columns[i].computed;if(!calc||row.cells[i]?.manual)continue;
    const vals=calc.fields.map(id=>value(row.cells[table.columns.findIndex(c=>c.id===id)]));
    const nums=vals.map(toNumber);
    let n=null;
    if(calc.op==='multiply'&&nums.every(v=>v!==null))n=nums[0]*nums[1];
    if(calc.op==='add'&&nums[0]!==null&&(nums[1]!==null||vals[1]===null||vals[1]===''))n=nums[0]+(nums[1]??0);
    row.cells[i]=n===null?emptyCell():{v:Math.round((n+Number.EPSILON)*1e8)/1e8,t:'n',z:'#,##0.000'};
  }
  return row;
}
export function matches(row,columns,condition){
  const index=columns.findIndex(c=>c.id===condition.field);if(index<0)return true;
  const cell=row.cells[index],v=value(cell),q=condition.value??'',q2=condition.value2??'';
  const text=String(v??'').toLocaleLowerCase(),query=String(q).toLocaleLowerCase();
  if(condition.op==='empty')return isBlank(cell);
  if(condition.op==='notempty')return!isBlank(cell);
  if(q===''||q===null)return true;
  if(condition.op==='eq')return text===query;
  if(condition.op==='ne')return text!==query;
  if(condition.op==='contains')return text.includes(query);
  if(condition.op==='notcontains')return !text.includes(query);
  if(condition.op==='starts')return text.startsWith(query);
  if(isBlank(cell))return false;
  const isDate=columns[index].type==='date';
  const a=isDate?Date.parse(String(v)):toNumber(v), b=isDate?Date.parse(q):toNumber(q),d=isDate?Date.parse(q2):toNumber(q2);
  if(a===null||b===null||!Number.isFinite(a)||!Number.isFinite(b))return false;
  if(condition.op==='gt')return a>b;if(condition.op==='gte')return a>=b;
  if(condition.op==='lt')return a<b;if(condition.op==='lte')return a<=b;
  if(condition.op==='between')return q2!==''&&Number.isFinite(d)&&d!==null&&a>=b&&a<=d;
  return true;
}
export function filterRows(table,conditions,join='all'){
  conditions=conditions.filter(c=>['empty','notempty'].includes(c.op)||(c.value!==''&&c.value!==null&&c.value!==undefined));
  if(!conditions.length)return table.rows;
  return table.rows.filter(r=>join==='any'?conditions.some(c=>matches(r,table.columns,c)):conditions.every(c=>matches(r,table.columns,c)));
}
export function recalculateAll(table){
  const rows=new Map(table.rows.map(r=>[r.id,r])),cols=new Map(table.columns.map((c,i)=>[c.id,i])),done=new Set(),visiting=new Set();
  function calc(row,i){
    let cell=row.cells[i];const key=row.id+':'+i,computed=!cell?.manual&&table.columns[i].computed;
    if((!cell?.formula&&!computed)||done.has(key))return cell;
    if(visiting.has(key))return{v:'#REF!',t:'e'};visiting.add(key);
    try{
      if(computed&&!cell?.formula){
        const vals=computed.fields.map(id=>{const j=cols.get(id);if(j===undefined)throw new Error('#REF!');const c=calc(row,j);if(c?.t==='e')throw new Error(c.v);return value(c);}),nums=vals.map(toNumber);
        let n=null;
        if(computed.op==='multiply'&&nums.every(v=>v!==null))n=nums[0]*nums[1];
        if(computed.op==='add'&&nums[0]!==null&&(nums[1]!==null||vals[1]===null||vals[1]===''))n=nums[0]+(nums[1]??0);
        cell=row.cells[i]=n===null?emptyCell():{v:Math.round((n+Number.EPSILON)*1e8)/1e8,t:'n',z:'#,##0.000'};
        return cell;
      }
      if(cell.formula.expr.includes('#REF!'))throw new Error('#REF!');
      const vals=cell.formula.refs.map(ref=>{const r=rows.get(ref.row),j=cols.get(ref.col);if(!r||j===undefined)throw new Error('#REF!');const c=calc(r,j);if(c?.t==='e')throw new Error(c.v);if(isBlank(c))return 0;const n=toNumber(value(c));if(n===null)throw new Error('#VALUE!');return n;});
      const tokens=cell.formula.expr.match(/@\d+@|\d+(?:\.\d*)?|\.\d+|[()+*/-]/g)||[];let p=0;
      function atom(){const t=tokens[p++];if(t==='+'||t==='-')return(t==='-'?-1:1)*atom();if(t==='('){const n=sum();if(tokens[p++]!==')')throw new Error('#VALUE!');return n;}if(/^@\d+@$/.test(t||''))return vals[Number(t.slice(1,-1))];if(t!==undefined&&Number.isFinite(Number(t)))return Number(t);throw new Error('#VALUE!');}
      function product(){let n=atom();while(tokens[p]==='*'||tokens[p]==='/'){const op=tokens[p++],b=atom();if(op==='/'&&b===0)throw new Error('#DIV/0!');n=op==='*'?n*b:n/b;}return n;}
      function sum(){let n=product();while(tokens[p]==='+'||tokens[p]==='-'){const op=tokens[p++],b=product();n=op==='+'?n+b:n-b;}return n;}
      const result=sum();if(p!==tokens.length||!Number.isFinite(result))throw new Error('#VALUE!');cell.v=Math.round((result+Number.EPSILON)*1e8)/1e8;cell.t='n';
    }catch(e){cell.v=e.message.startsWith('#')?e.message:'#VALUE!';cell.t='e';}
    finally{visiting.delete(key);done.add(key);}
    return cell;
  }
  for(const row of table.rows)for(let i=0;i<table.columns.length;i++)calc(row,i);
}
export function groupRows(rows,columns,field){
  const index=columns.findIndex(c=>c.id===field);
  if(index<0)return rows.length?[{key:'all',name:'筛选结果',rows}]:[];
  const map=new Map();
  for(const row of rows){const cell=row.cells[index];const blank=isBlank(cell);const key=blank?'empty':JSON.stringify([cell?.t??'s',value(cell)]);if(!map.has(key))map.set(key,{key,name:blank?'（空白）':display(cell),rows:[]});map.get(key).rows.push(row);}
  return [...map.values()];
}
export function safeName(name,used=new Set(),max=31){
  let base=String(name||'未命名').replace(/[\x00-\x1f\\/:*?"<>|\[\]]/g,'_').replace(/^'+|'+$/g,'').trim()||'未命名';
  if(/^(con|prn|aux|nul|com\d|lpt\d)$/i.test(base))base='_'+base;
  let result=base.slice(0,max),n=2;
  while(used.has(result.toLowerCase())){const suffix=`_${n++}`;result=base.slice(0,max-suffix.length)+suffix;}
  used.add(result.toLowerCase());return result;
}
export function parseTSV(text){
  const rows=[];let row=[],field='',quoted=false;
  for(let i=0;i<text.length;i++){
    const ch=text[i];
    if(ch==='"'){if(quoted&&text[i+1]==='"'){field+='"';i++;}else if(quoted||field==='')quoted=!quoted;else field+=ch;}
    else if(ch==='\t'&&!quoted){row.push(field);field='';}
    else if((ch==='\n'||ch==='\r')&&!quoted){if(ch==='\r'&&text[i+1]==='\n')i++;row.push(field);rows.push(row);row=[];field='';}
    else field+=ch;
  }
  row.push(field);if(row.length>1||row[0]!==''||!rows.length)rows.push(row);return rows;
}
