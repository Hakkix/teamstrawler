# TeamsTrawler

> Scroll back through Microsoft Teams chat history and save every message to a plain-text file.

TeamsTrawler connects to a running Chrome browser via the DevTools protocol, auto-detects the Teams chat scroll container, and progressively scrolls upward — extracting, deduplicating, and persisting messages as it goes. If the script is interrupted it picks up from where it left off.

---

## Features

- **Auto-detection** — finds the chat container automatically using known Teams DOM selectors; falls back to hover-and-click selection when auto-detection fails.
- **Deduplication** — each message is hashed by author + timestamp + content; the same message is never written twice even if it appears in multiple scroll passes.
- **Crash recovery** — a `.checkpoint.json` file is updated every few loops so a restart resumes from the last saved position rather than from scratch.
- **Smart waiting** — after each scroll the script polls the DOM until new content loads instead of sleeping for a fixed number of seconds, making it both faster and more reliable on slow connections.
- **Zero driver setup** — uses Selenium's built-in Selenium Manager (included since Selenium 4.6) to automatically download and cache the correct ChromeDriver for your installed Chrome version. No extra packages or manual driver management needed.
- **Overwrite protection** — if the output file already exists you are prompted before it is replaced.

---

## Requirements

- Python 3.10 or newer
- Google Chrome with remote debugging enabled (see setup below)

Install Python dependencies:

```bash
pip install -r requirements.txt
```

---

## Setup

Launch Chrome with the remote debugging port open **before** running the script. Close any existing Chrome windows first, then run:

**macOS**
```bash
/Applications/Google\ Chrome.app/Contents/MacOS/Google\ Chrome \
  --remote-debugging-port=9222 \
  --user-data-dir=/tmp/chrome-debug
```

**Windows**
```cmd
"C:\Program Files\Google\Chrome\Application\chrome.exe" ^
  --remote-debugging-port=9222 ^
  --user-data-dir=C:\tmp\chrome-debug
```

**Linux**
```bash
google-chrome --remote-debugging-port=9222 --user-data-dir=/tmp/chrome-debug
```

Then log in to Microsoft Teams in that browser window and open the chat you want to export.

---

## Usage

```bash
python teamstrawler.py chat.txt
```

The script will auto-detect the chat container, highlight it in green, and begin scrolling. Progress is printed to the terminal and shown as an overlay inside the browser window.

When finished, messages are written to `chat.txt` in chronological order (oldest first), one message per line:

```
[9:15 AM] Alice Johnson: morning everyone
[9:16 AM] Bob Smith: hey! ready for the standup?
[9:17 AM] Alice Johnson: give me 2 min ☕
```

### Options

| Flag | Description |
|---|---|
| `--no-resume` | Ignore any existing checkpoint and start from scratch |
| `--chrome-path PATH` | Path to the Chrome executable (see below) |

```bash
# Start fresh even if a checkpoint exists
python teamstrawler.py chat.txt --no-resume

# Specify Chrome location explicitly (Windows example)
python teamstrawler.py chat.txt --chrome-path "C:\Program Files\Google\Chrome\Application\chrome.exe"
```

### Specifying the Chrome path

Selenium Manager auto-detects Chrome on most systems, so `--chrome-path` is usually not needed. Pass it when Chrome is installed in a non-standard location or when auto-detection fails.

| Platform | Default Chrome path |
|---|---|
| macOS | `/Applications/Google Chrome.app/Contents/MacOS/Google Chrome` |
| Windows | `C:\Program Files\Google\Chrome\Application\chrome.exe` |
| Linux | `/usr/bin/google-chrome` |

---

## How it works

1. Connects to Chrome on `127.0.0.1:9222` via Selenium's DevTools bridge.
2. Scans the page for a scrollable element matching known Teams selectors.
3. Enters a loop: extract visible messages → deduplicate → scroll up → wait for DOM update → repeat.
4. Stops when no new messages appear for 12 consecutive loops, or when the container is confirmed to be at the top and scroll position stops changing.
5. Writes the final message list in chronological order and deletes the checkpoint file.

---

## Tuning

Constants at the top of `teamstrawler.py` can be adjusted if needed:

| Constant | Default | Description |
|---|---|---|
| `EMPTY_LOOP_LIMIT` | `12` | Consecutive empty loops before stopping |
| `MAX_SCROLL_STALLS` | `6` | Consecutive no-movement scrolls before stopping |
| `AUTOSAVE_INTERVAL` | `5` | Save + checkpoint every N loops that found new messages |
| `SCROLL_WAIT_TIMEOUT` | `5.0` | Max seconds to wait for DOM update after each scroll |
| `SCROLL_STEP_RATIO` | `0.6` | Fraction of the container height scrolled per step |

---

## Limitations

- Requires Chrome; other browsers are not supported.
- The DOM selectors target Microsoft Teams' current web interface. A Teams update that changes internal element attributes may require selector adjustments.
- Reactions, file attachments, and inline images are not captured — only the text content of messages.
- This tool is for personal archival use. Make sure your use complies with your organisation's data and communication policies.

---

## License

MIT
