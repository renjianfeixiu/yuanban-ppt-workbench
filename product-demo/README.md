# 计划书与 PPT 工作台

运行、模板格式和复核说明见[仓库首页](../README.md)。

界面在 `public/`。`server.mjs` 管理上传、生成和本地状态，`workflow.mjs` 管理大纲与计划书，`content.mjs` 将字段分配到模板文字框并检查容量。`pptx.mjs` 按演示文稿的页面关系读取文字与图片，用局部 XML 修改生成新文件。`policies.mjs` 提供中国全国官方政策查询。
