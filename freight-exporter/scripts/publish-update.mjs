import fs from 'node:fs';
import path from 'node:path';
import {execFileSync,spawnSync} from 'node:child_process';
import {fileURLToPath} from 'node:url';
import {validateUpdateManifest,verifyUpdateArtifact} from '../dist/update-core.js';
const project=path.resolve(path.dirname(fileURLToPath(import.meta.url)),'../..');
if(process.env.GITHUB_ACTIONS!=='true'||process.env.GITHUB_REPOSITORY!=='nQrChao/lcgwork')throw new Error('仅允许在配置的 GitHub 仓库工作流内发布。');
const manifest=validateUpdateManifest(JSON.parse(fs.readFileSync(path.join(project,'release/update.json'),'utf8')));
const artifact=fs.readFileSync(path.join(project,'release',manifest.file));await verifyUpdateArtifact(artifact,manifest);
if(manifest.sourceSha!==process.env.GITHUB_SHA)throw new Error('发布产物与当前提交不匹配。');
const env={...process.env,GIT_AUTHOR_NAME:'GLC 发布',GIT_AUTHOR_EMAIL:'41898282+github-actions[bot]@users.noreply.github.com',GIT_COMMITTER_NAME:'GLC 发布',GIT_COMMITTER_EMAIL:'41898282+github-actions[bot]@users.noreply.github.com'};
const git=(args,input)=>execFileSync('git',args,{cwd:project,env,input,encoding:'utf8',stdio:['pipe','pipe','inherit']}).trim();
const latest=git(['ls-remote','--heads','origin','refs/heads/main']).split(/\s/)[0];
if(latest!==manifest.sourceSha){console.log('main 已有更新提交，跳过这次过期构建。');process.exit(0);}
const existing=git(['ls-remote','--heads','origin','refs/heads/updates']).split(/\s/)[0];let parent='';
if(existing){git(['fetch','--no-tags','origin','refs/heads/updates']);parent=git(['rev-parse','FETCH_HEAD']);}
const entries=[['update.json',fs.readFileSync(path.join(project,'release/update.json'))],[manifest.file,artifact]];
const tree=git(['mktree'],entries.map(([name,bytes])=>'100644 blob '+git(['hash-object','-w','--stdin'],bytes)+'\t'+name).join('\n')+'\n');
const commit=git(['commit-tree',tree,...(parent?['-p',parent]:[])],`发布 GLC工作台 ${manifest.version} · ${manifest.sourceSha.slice(0,8)}\n`);
// No force push: concurrent publications cannot overwrite each other.
git(['push','origin',commit+':refs/heads/updates']);
console.log('已发布 '+manifest.version+' · '+manifest.sourceSha.slice(0,8));
