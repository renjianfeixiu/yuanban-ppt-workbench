# 原版 · 计划书与 PPT 工作台

从审核过的大纲生成计划书，在已有 PPTX 模板中填入内容，下载可编辑的演示文稿。项目在本机运行。

## 下载与启动

安装 [Node.js 22 或更高版本](https://nodejs.org/en/download/)。克隆仓库后，在仓库根目录执行：

```sh
git clone https://github.com/renjianfeixiu/yuanban-ppt-workbench.git
cd yuanban-ppt-workbench
npm ci
npm start
```

打开终端显示的地址，默认是 [127.0.0.1:5186](http://127.0.0.1:5186/)。保持终端运行，按 Ctrl+C 停止。仓库访问权限由仓库所有者管理。

如果下载的是 [Releases 中附带依赖的 ZIP](https://github.com/renjianfeixiu/yuanban-ppt-workbench/releases/latest)，解压后直接执行 `npm start`。Windows 可以双击 `start-windows.cmd`，macOS 可以双击 `product-demo/启动产品Demo.command`。启动器会使用 PATH 中的 Node.js。

## 跑一次样例

1. 使用预填的项目大纲。打开人工审核，勾选确认后生成计划书草稿。
2. 确认草稿，选择“使用样例模板”。工作台会把项目字段填入三页中的对应位置。
3. 点击“检查与生成”，再点击“生成 PPT 副本”。在逐页审核界面下载候选 PPT，或用 WPS 查看。
4. 如需添加政策材料，在相应页面查找并选用官方出处，然后重新生成。出处写入对应页的备注。

## 使用自己的模板

在第四步上传普通 `.pptx` 文件。支持最多 200 页、压缩文件不超过 160 MiB。工作台按文件里的页面顺序读取已有文字框和图片。可以逐页编辑文字、选择计划书字段、替换嵌入图片或添加备注。

要重复套用同一模板，可以在文字框里写占位标记。常用标记有 `{{topic}}`（题目）、`{{name}}`（简称）、`{{audience}}`（服务对象）、`{{problem}}`（问题）、`{{solution}}`（方案）、`{{facts}}`（已知事实），以及 `{{chapter.1.title}}` 和 `{{chapter.1.points}}`（第一章标题和内容）。这些标记也可放在一段文字中。未加标记的文字由用户在编辑器中选择或修改。已识别的平安货运模板仍有第二页五栏专用适配。

容量检查使用模板文本框的尺寸或原文长度估算。网页提供位置预览，完整背景、主题、裁切、字体和动画请在 WPS 核对。图表、表格、SmartArt 和图片内文字保留原件。项目草稿基于规则生成。政策查询可联网按关键词检索，无可靠结果时展示本地资料库匹配。上传的模板、图片、草稿和生成文件保存在本地。

## 平台与配置

- macOS、Windows 和 Linux 使用同一份 Node.js 源码和 npm 锁文件。
- 需要用 WPS 的用户可先下载候选文件，再在 WPS 中打开。工作台的打开按钮默认查找 `wpsoffice`（macOS）、`wpp.exe`（Windows）或 `wpp`（Linux）。可通过 `PPT_DEMO_WPS_PATH` 指定应用或可执行文件。
- 端口占用时可用 `PPT_DEMO_PORT` 选择端口。`PPT_DEMO_DATA_DIR` 可指定独立数据目录，未设置时使用 `product-demo/.work/` 和 `product-demo/output/`。
- 运行 `npm test` 验证从空目录启动、内置样例导出、上传不同模板、长度拦截、图片与备注以及复核状态。测试使用独立临时目录。

`product-demo/` 包含界面、本地服务和通用 PPTX 处理代码。`examples/` 提供可分发的样例模板。`demo-prototype/` 保留 XML 补丁工具、命令行演示和概念配图。配图为方案示意。
