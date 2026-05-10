#!/bin/zsh
set -euo pipefail
cd "$(dirname "$0")"
mkdir -p logs
mkdir -p "$PWD/.chrome_automation_profile"
log_file="logs/run_$(date +%Y%m%d_%H%M%S).log"
ln -sfn "$(basename "$log_file")" logs/latest.log
echo "Logging to: $PWD/$log_file"
.venv/bin/python -u chatgpt_real_chrome_runner.py \
  --chrome-data-dir "$PWD/.chrome_automation_profile" \
  --prompts-per-chat 10 \
  --resume 2>&1 | tee -a "$log_file"
runner_status=${pipestatus[1]}
exit "$runner_status"
