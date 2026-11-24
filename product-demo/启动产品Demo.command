#!/bin/zsh
TASK_PPT_SERVER="${0:A:h}/server.mjs"
if [[ ! -f "$TASK_PPT_SERVER" && "${1:t}" == '启动产品Demo.command' && -f "${1:A:h}/server.mjs" ]]; then
  TASK_PPT_SERVER="${1:A:h}/server.mjs"
fi
if [[ ! -f "$TASK_PPT_SERVER" ]]; then
  TASK_PPT_SERVER='/Users/wangdajin/我的文稿/agent task/codex 任务尝试/ppt-template-study/product-demo/server.mjs'
fi
exec '/Users/wangdajin/.cache/codex-runtimes/codex-primary-runtime/dependencies/node/bin/node' "$TASK_PPT_SERVER" --open-browser
