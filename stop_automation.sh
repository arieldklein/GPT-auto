#!/bin/zsh
set -euo pipefail
pkill -f 'chatgpt_real_chrome_runner.py' 2>/dev/null || true
pkill -f 'Google Chrome.*--remote-debugging-port=9222' 2>/dev/null || true
echo "Stopped GPT automation runner and automation Chrome, if they were running."
