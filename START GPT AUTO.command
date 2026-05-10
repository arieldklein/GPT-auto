#!/bin/zsh
cd "$(dirname "$0")" || exit 1
clear
echo "GPT AUTO"
echo "Folder: $PWD"
echo
echo "Starting/resuming automation..."
echo "Close this Terminal window or press Ctrl+C to stop the runner."
echo
./start_automation.sh
runner_status=$?
echo
echo "Runner exited with status $runner_status."
echo "Press Enter to close this window."
read -r _
exit $runner_status
