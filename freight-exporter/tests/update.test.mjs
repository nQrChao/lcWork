import {test} from 'node:test';import assert from 'node:assert/strict';
import {validateUpdateManifest,readBuildMetadata,updateDigest,verifyUpdateArtifact} from '../dist/update-core.js';
const base={schema:1,appId:'glc-workbench',repository:'nQrChao/lcWork',sourceBranch:'main',sourceSha:'b'.repeat(40),buildId:'b'.repeat(40),version:'1.1.0',builtAt:'2026-10-14T13:00:00Z',notes:'更新录入体验',file:'lcWork.html',bytes:1024,sha256:'a'.repeat(64)};
test('update identity, file path, lengths and hashes reject malformed manifests',()=>{assert.equal(validateUpdateManifest(base).version,'1.1.0');for(const change of [{schema:2},{repository:'other/repo'},{file:'../../work.xlsx'},{bytes:-1},{bytes:50*1024*1024},{sha256:'invalid'},{sourceSha:'main'},{buildId:'a'.repeat(40)},{notes:{}},{version:'<img>'},{builtAt:'invalid'}])assert.throws(()=>validateUpdateManifest({...base,...change}));});
test('verified artifact matches the commit and does not contain embedded business data',async()=>{const meta={appId:base.appId,repository:base.repository,buildId:base.buildId,version:base.version,seeded:false,channel:'updates'};const make=seed=>new TextEncoder().encode('<!doctype html><script id="glc-build" type="application/json">'+JSON.stringify(meta)+'</script><script id="seed-workbook" type="application/octet-stream">'+seed+'</script>'+' '.repeat(1024));const bytes=make(''),manifest={...base,bytes:bytes.length,sha256:await updateDigest(bytes)};assert.equal((await verifyUpdateArtifact(bytes,manifest)).buildId,base.buildId);const corrupt=bytes.slice();corrupt[0]^=1;await assert.rejects(()=>verifyUpdateArtifact(corrupt,manifest),/校验失败/);await assert.rejects(()=>verifyUpdateArtifact(bytes,{...manifest,buildId:'c'.repeat(40)}),/不匹配/);const seeded=make('secret');await assert.rejects(async()=>verifyUpdateArtifact(seeded,{...manifest,bytes:seeded.length,sha256:await updateDigest(seeded)}),/内置数据/);});
test('program identity is required before any local file is overwritten',()=>{assert.throws(()=>readBuildMetadata('<html>other program</html>'));assert.throws(()=>readBuildMetadata('<script id="glc-build">{}</script>'));assert.throws(()=>readBuildMetadata('<script id="glc-build">bad JSON</script>'));});

test('repository migration accepts old local programs but rejects old download manifests and artifacts',async()=>{
  const legacy={appId:base.appId,repository:'nQrChao/lcgwork',buildId:base.buildId,version:base.version,seeded:false,channel:'updates'};
  const html='<script id="glc-build">'+JSON.stringify(legacy)+'</script><script id="seed-workbook"></script>'+' '.repeat(1024);
  assert.equal(readBuildMetadata(html).repository,legacy.repository);
  assert.throws(()=>validateUpdateManifest({...base,repository:legacy.repository}));
  const bytes=new TextEncoder().encode(html);
  const sha256=await updateDigest(bytes);
  await assert.rejects(()=>verifyUpdateArtifact(bytes,{...base,bytes:bytes.length,sha256}),/不匹配/);
});
