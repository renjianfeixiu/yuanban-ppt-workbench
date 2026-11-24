#!/bin/zsh
TASK_DEMO_DIR="${0:A:h}"
exec '/Users/wangdajin/.cache/codex-runtimes/codex-primary-runtime/dependencies/node/bin/node' "$TASK_DEMO_DIR/run-demo.mjs" "$@"
