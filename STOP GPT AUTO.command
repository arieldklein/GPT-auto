#!/bin/zsh
cd "$(dirname "$0")" || exit 1
clear
echo "GPT AUTO STOP"
echo "Folder: $PWD"
echo
./stop_automation.sh
echo
./status_automation.sh
echo
echo "Press Enter to close this window."
read -r _
