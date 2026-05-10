# GPT Auto

Browser automation for filling an Excel workbook one prompt at a time through an AI chat website.

The current runner is built for ChatGPT web. It:

- opens a real Chrome automation profile,
- uploads the latest canonical workbook,
- sends `fill prompt N and audit`,
- waits for the model to finish,
- downloads the returned `.xlsx`,
- merges only the requested prompt row back into the canonical workbook,
- verifies the exact requested row was filled before advancing,
- starts a fresh chat every 10 prompts by default,
- saves progress in `progress.json`.

Private files are intentionally ignored by git: Chrome profiles, logs, progress, downloaded outputs, backups, and workbooks.

## Setup

```bash
python3 -m venv .venv
.venv/bin/python -m pip install -r requirements.txt
.venv/bin/python -m playwright install chrome
```

Put your source workbook in this folder and name it:

```text
ceo_fitness_tracker_v2_sequential.xlsx
```

Or run the Python script with explicit paths:

```bash
.venv/bin/python chatgpt_real_chrome_runner.py \
  --initial-excel "/path/to/source.xlsx" \
  --canonical-excel "/path/to/source.xlsx" \
  --output-dir chatgpt_outputs \
  --backup-dir canonical_backups \
  --progress progress.json \
  --prompts-per-chat 10 \
  --resume
```

## Run

On macOS, double-click:

```text
START GPT AUTO.command
```

The first time Chrome opens, sign in manually and choose the model/account you want. The script does not store credentials in git.

Useful commands:

```text
STATUS GPT AUTO.command
STOP GPT AUTO.command
```

## Important Safety Behavior

The runner refuses to advance when:

- ChatGPT returns the wrong prompt workbook,
- the requested row is still `Not Started`,
- required evidence/audit fields are blank,
- a human verification or true blocking notice appears.

If ChatGPT returns the previous prompt by mistake, the runner asks it to correct the workbook once before stopping.

## Using Gemini With Codex

This repo ships the ChatGPT web runner because ChatGPT and Gemini use different web UIs, selectors, upload flows, and download behavior.

To configure a Gemini version, open this repo in Codex and ask Codex:

```text
Create a Gemini version of this runner.

Use chatgpt_real_chrome_runner.py as the base, but create gemini_real_chrome_runner.py.
Keep the workbook merge, progress, validation, backup, and prompt-rotation logic unchanged.
Replace only the browser-specific pieces:
- URL should be https://gemini.google.com/app
- composer detection should target Gemini's prompt textarea/contenteditable
- send-button detection should target Gemini's send button
- upload_file should target Gemini's file attachment flow
- wait_until_response_done should detect Gemini's finished-response state
- click_download_excel should capture the Excel workbook Gemini returns, or document the manual download fallback if Gemini cannot attach .xlsx files reliably

Then create START GEMINI AUTO.command, STOP GEMINI AUTO.command, and STATUS GEMINI AUTO.command.
Do not commit Chrome profiles, progress files, logs, downloads, or workbooks.
```

After Codex creates the Gemini runner, test with `--end-prompt 1` or another tiny range before running the full workbook.
