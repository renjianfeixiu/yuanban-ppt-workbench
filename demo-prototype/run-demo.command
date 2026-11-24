#!/bin/zsh
set -e
export PATH="/opt/homebrew/bin:/usr/local/bin:$PATH"
entry="${${(%):-%x}:A:h}/run-demo.mjs"
exec node "$entry" "$@"
