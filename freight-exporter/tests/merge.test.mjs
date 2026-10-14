import test from 'node:test';
import assert from 'node:assert/strict';
import {mergeTables,sheetToTable} from '../dist/excel.js';
import {recalculateAll} from '../dist/core.js';
const cell=v=>({v,t:typeof v==='number'?'n':'s'});
const column=(id,name,type='number')=>({id,name,type});
test('merge aligns reordered and missing columns; remaps colliding row/column IDs and formula references',()=>{
  const a={columns:[column('q','数量'),column('p','单价'),column('v','金额')],rows:[{id:'r',cells:[cell(3),cell(10),{...cell(30),formula:{expr:'@0@*@1@',refs:[{row:'r',col:'q'},{row:'r',col:'p'}]}}]}]};
  const b={columns:[column('p','单价'),column('v','金额'),column('q','数量'),column('note','备注','text')],rows:[{id:'r',cells:[cell(5),{...cell(10),formula:{expr:'@0@*@1@',refs:[{row:'r',col:'q'},{row:'r',col:'p'}]}},cell(2),cell('不丢失')]}]};
  const snapshots=JSON.stringify([a,b]);const result=mergeTables([{table:a,sheetName:'甲'},{table:b,sheetName:'乙'}],'总表',{sourceField:'货运公司'});
  assert.deepEqual(result.columns.map(c=>c.name),['货运公司','数量','单价','金额','备注']);
  recalculateAll(result);assert.deepEqual(result.rows.map(r=>r.cells.map(c=>c.v)),[['甲',3,10,30,null],['乙',2,5,10,'不丢失']]);
  result.rows[1].cells[1]=cell(7);recalculateAll(result);assert.equal(result.rows[1].cells[3].v,35);assert.equal(result.rows[0].cells[3].v,30);assert.equal(JSON.stringify([a,b]),snapshots);
});
test('source column preserves existing names, fills blanks, supports custom name and omission',()=>{
  const table={columns:[column('c','项目','text'),column('q','数量')],rows:[{id:'a',cells:[cell('已有'),cell(1)]},{id:'b',cells:[cell(''),cell(2)]}]};
  const parts=[{table,sheetName:'表名'}];const merged=mergeTables(parts,'总表',{sourceField:'项目'});assert.equal(merged.columns.length,2);assert.deepEqual(merged.rows.map(r=>r.cells[0].v),['已有','表名']);assert.equal(mergeTables(parts,'总表').rows[1].cells[0].v,'');assert.throws(()=>mergeTables([],'空表'),/至少/);
});
test('metadata computations apply to the correct rows and retain manual/blank values from other sheets',()=>{
  const columns=[column('a','吨位'),column('b','单价'),{...column('c','金额'),computed:{op:'multiply',fields:['a','b']}}];
  const a={columns,rows:[{id:'r',cells:[cell(2),cell(10),cell(20)]}]};
  const b={columns:columns.map(({computed,...col})=>col),rows:[{id:'r',cells:[cell(3),cell(10),cell(999)]},{id:'blank',cells:[cell(3),cell(10),cell(null)]}]};
  const table=mergeTables([{table:a,sheetName:'甲'},{table:b,sheetName:'乙'}],'总表');recalculateAll(table);assert.deepEqual(table.rows.map(r=>r.cells[2].v),[20,999,null]);table.rows[0].cells[0]=cell(4);recalculateAll(table);assert.equal(table.rows[0].cells[2].v,40);
});
test('unsupported formulas and missing cached results remain visible for import confirmation',()=>{
  const sheet={columns:[column('v','金额')],rows:[{id:'r',cells:[{...cell(null),rawFormula:'SUM(A1:A3)',missingFormulaResult:true}]}]};const result=mergeTables([{table:sheet,sheetName:'甲'}],'总表');assert.equal(result.unsupported,1);assert.equal(result.rows[0].cells[0].missingFormulaResult,true);assert.equal(result.rows[0].cells[0].rawFormula,'SUM(A1:A3)');
});
test('header validation rejects empty, fractional and out-of-range rows',()=>{const sheet={rows:[[cell('字段')],[cell(1)]],headerIndex:0};for(const index of [-1,0.5,2,NaN])assert.throws(()=>sheetToTable(sheet,'测试',index),/表头/);assert.equal(sheetToTable(sheet,'测试').rows.length,1);});
test('a formula without a cached result is retained rather than discarded as an empty row',()=>{const sheet={name:'测试',headerIndex:0,rows:[[cell('金额')],[{...cell(null),rawFormula:'SUM(B2:C2)',missingFormulaResult:true}]]};const table=sheetToTable(sheet,'测试');assert.equal(table.rows.length,1);assert.equal(table.unsupported,1);assert.equal(table.rows[0].cells[0].missingFormulaResult,true);});
test('import can skip formula-only rows while retaining uncached formulas and unnamed extra data',()=>{
 const sheet={name:'测试',headerIndex:0,rows:[[cell('数量'),cell('金额')],[cell(2),cell(4),cell('表头外附注')],[cell(null),{...cell(0),rawFormula:'A3*2'}],[cell(null),{...cell(null),rawFormula:'SUM(A4)',missingFormulaResult:true}]]};
 const all=sheetToTable(sheet,'全部',0,{includeExtraColumns:true});assert.equal(all.rows.length,3);assert.equal(all.columns.length,3);assert.equal(all.rows[0].cells[2].v,'表头外附注');
 const filtered=sheetToTable(sheet,'过滤',0,{includeExtraColumns:true,skipFormulaOnlyRows:true});assert.equal(filtered.rows.length,2);assert.equal(filtered.skippedFormulaRows,1);assert.equal(filtered.rows[1].cells[1].missingFormulaResult,true);
});
