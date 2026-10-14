export const UPDATE_APP_ID='glc-workbench';
export const UPDATE_REPOSITORY='nQrChao/lcgwork';
export const UPDATE_MAX_BYTES=40*1024*1024;
export function validateUpdateManifest(input){
  if(!input||input.schema!==1||input.appId!==UPDATE_APP_ID||input.repository!==UPDATE_REPOSITORY||input.sourceBranch!=='main')throw new Error('更新信息不属于当前工作台。');
  if(!/^[a-f0-9]{40}$/.test(input.sourceSha||'')||input.buildId!==input.sourceSha)throw new Error('更新缺少有效的提交版本。');
  if(!/^\d+\.\d+\.\d+(?:[-+][a-zA-Z0-9.-]+)?$/.test(input.version||'')||!Number.isFinite(Date.parse(input.builtAt)))throw new Error('更新版本或发布时间无效。');
  if(input.file!=='GLC工作台.html'||!Number.isInteger(input.bytes)||input.bytes<1000||input.bytes>UPDATE_MAX_BYTES||!/^[a-f0-9]{64}$/.test(input.sha256||''))throw new Error('更新文件信息无效。');
  if(typeof input.notes!=='string'||input.notes.length>12000)throw new Error('更新说明无效。');
  return {schema:1,appId:input.appId,repository:input.repository,sourceBranch:input.sourceBranch,sourceSha:input.sourceSha,buildId:input.buildId,version:input.version,builtAt:input.builtAt,file:input.file,bytes:input.bytes,sha256:input.sha256,notes:input.notes};
}
export function readBuildMetadata(html){
  const match=String(html).match(/<script\b[^>]*\bid=["']glc-build["'][^>]*>([\s\S]*?)<\/script>/i);
  if(!match||match[1].length>8000)throw new Error('所选文件不是支持更新的 GLC工作台。');
  let value;try{value=JSON.parse(match[1]);}catch{throw new Error('程序版本信息损坏。');}
  if(value.appId!==UPDATE_APP_ID||value.repository!==UPDATE_REPOSITORY||!/^[a-f0-9]{40,64}$/.test(value.buildId||''))throw new Error('程序身份不匹配。');
  return value;
}
export async function updateDigest(bytes){return [...new Uint8Array(await crypto.subtle.digest('SHA-256',bytes))].map(b=>b.toString(16).padStart(2,'0')).join('');}
export async function verifyUpdateArtifact(bytes,manifest){
  if(bytes.byteLength!==manifest.bytes||await updateDigest(bytes)!==manifest.sha256)throw new Error('新版文件校验失败，未安装更新。请重新检查更新。');
  const html=new TextDecoder('utf-8',{fatal:true}).decode(bytes),meta=readBuildMetadata(html);
  if(meta.buildId!==manifest.buildId||meta.version!==manifest.version||meta.seeded!==false||meta.channel!=='updates')throw new Error('新版程序与更新说明不匹配。');
  const seed=html.match(/<script\b[^>]*\bid=["']seed-workbook["'][^>]*>([\s\S]*?)<\/script>/i);
  if(!seed||seed[1].trim())throw new Error('公开更新包包含非空内置数据，已停止安装。');
  return meta;
}
