# PPTX 替换命令行示例

先在仓库根目录运行 `npm ci`，然后运行：

```sh
node demo-prototype/run-demo.mjs
```

命令会使用内置样例生成可编辑的 PPTX，保存在 `demo-prototype/output/`。可以用 `--template` 指定模板路径，用 `--output-dir` 指定输出目录。普通模板中有字段标记的位置会填入样例数据。对于已识别的平安货运模板，命令会套用第二页示例文案和五张概念图。

需要审核大纲、填写自己的计划书、逐页修订和引用政策时，使用[工作台](../README.md)。
