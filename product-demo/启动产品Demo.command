#!/bin/zsh
set -e
export PATH="/opt/homebrew/bin:/usr/local/bin:$PATH"
cd -- "${${(%):-%x}:A:h:h}"
if ! command -v node >/dev/null; then
  print '请先安装 Node.js 22 或更新版本：https://nodejs.org/'
  exit 1
fi
if [[ ! -f node_modules/yauzl/package.json ]]; then
  npm ci --omit=dev
fi
exec node product-demo/server.mjs --open-browser
