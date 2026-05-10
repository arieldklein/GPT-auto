#!/bin/zsh
set -euo pipefail
cd "$(dirname "$0")"
echo "Progress:"
if [ -f progress.json ]; then
  cat progress.json
else
  echo "No progress.json yet."
fi
echo
echo "Processes:"
pgrep -fl 'chatgpt_real_chrome_runner.py|Google Chrome --remote-debugging-port=9222' || true
echo
if [ -f logs/latest.log ]; then
  echo "Latest log:"
  echo "$PWD/logs/latest.log"
  tail -40 logs/latest.log
fi
