# GLC工作台

在 Windows 或 macOS 的 Chrome / Edge 中打开 `GLC工作台.html` 即可使用。新建表格前选择本地目录，录入内容实际写入 Excel；支持自动保存、筛选分表、打印、PNG / PDF、外部表格同步及独立印章工具。

公开程序只提供空白运费模板，包含 22 个初始字段；不会内置真实运费记录、单据或印章。已有 Excel 可以自行导入。运行时无需服务器，只有点击「检查更新」才会连接 GitHub。

选中已放置的印章后，可拖动上方手柄旋转，或在左侧设置角度、左右转 90°、一键归正。骑缝章先旋转再分片；PNG 和 PDF 导出保留旋转效果。

## 首次上传到仓库

仓库：<https://github.com/nQrChao/lcgwork>，源码分支：`main`。

将此发布目录里的 `.github`、`.gitignore`、`README.md` 和 `freight-exporter` 提交到仓库根目录，不要把它们套在额外一层文件夹里。不要上传原工作目录的 `work`、`outputs`、真实表格、单据、印章或旧的完整项目 ZIP。

如使用 Git，可在此目录执行：

```sh
git init -b main
git add .
git commit -m "初始化 GLC工作台，支持在线更新"
git remote add origin https://github.com/nQrChao/lcgwork.git
git push -u origin main
```

如果仓库已有提交，请先拉取并合并现有内容，不要强制覆盖。

上传后，在仓库 Actions 中查看「发布工作台更新」。首次成功运行后，会自动创建 `updates` 分支，其中只有 `GLC工作台.html` 和 `update.json`。从该分支下载 HTML，或使用本地附带的 HTML 检查更新。

如果 Actions 被仓库或组织策略禁用，需要在仓库设置中启用；工作流声明了发布所需的 `contents: write` 权限。程序不需要填写任何 GitHub Token。

## 日常更新

1. 修改 `freight-exporter/dist` 中的程序、`index.template.html` 中的页面，必要时修改 `package.json` 的版本号。
2. 用清楚的中文写提交说明，并提交到 `main`。
3. Actions 自动运行测试、打包并发布更新；构建失败时保留上一个可用版本。
4. 本地程序点击左下角「检查更新」，查看是否有新版本。
5. 有更新时点击「立即更新」，选择当前 HTML 所在目录，下载、校验、备份旧版并安装。点击「重新加载新版」，再选择原表格目录继续工作。

即使版本号未改变，新的提交也会被识别。更新锁定一次构建的提交，下载中出现另一笔提交不会混用两个版本。安装前后均检查文件校验值；旧版保存在程序目录的 `GLC程序备份` 中。

未保存的表格需要先保存成功；盖章和扫描草稿需要先完成或清空。没有网络时继续使用本地版本。不支持直接安装的浏览器会显示「下载更新」，下载后手动打开。程序更新只写程序文件及其备份，不改业务文件；确认更新前的表格保存除外。

首次安装需授权程序目录。浏览器可能要求再次授权；若选错目录，点击「重新选择程序目录」。浏览器无法核实不同目录中两个完全相同副本的系统路径，请选择当前页面实际打开的文件所在目录。程序更新后仍按浏览器权限规则重新关联业务目录。

## 开发

需要 Node.js 24。离线依赖已随源代码提供，构建无需安装 npm 依赖：

```sh
node --test freight-exporter/tests/*.test.mjs
node freight-exporter/build.mjs
```

默认构建不读取任何本地业务数据。`--release` 供 GitHub Actions 使用，生成 `release` 中的更新产物和校验信息。`--with-local-data` 仅供原项目的私有回归测试，不允许与 `--release` 同时使用。

浏览器集成测试需在 `freight-exporter` 内安装开发依赖，并安装 Chrome：

```sh
npm install
npm run build
npm run test:update
npm run test:extended
npm run test:stamp-rotation
```

`dist/update-core.js` 负责更新协议和校验，`dist/updater.js` 负责界面、下载和本地安装，`scripts/publish-update.mjs` 在 GitHub Actions 中发布到更新分支。下载仅访问固定仓库的 GitHub API 与 raw.githubusercontent.com。

依赖许可证位于 `freight-exporter/dist/vendor`，并保留在打包后的 HTML 内。扫描 OCR 的结果仍需人工核对。
