#!/bin/zsh
cd "$(dirname "$0")" || exit 1
clear
echo "GPT AUTO STATUS"
echo "Folder: $PWD"
echo
./status_automation.sh
echo
echo "Press Enter to close this window."
read -r _
