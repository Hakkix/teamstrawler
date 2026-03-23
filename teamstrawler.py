"""
TeamsTrawler — Microsoft Teams chat history scraper.

Connects to an already-running Chrome instance (--remote-debugging-port=9222),
auto-detects the chat scroll container, scrolls upward, extracts and
deduplicates messages, and writes them to a plain-text file.

Selenium 4.6+ ships its own Selenium Manager which downloads and caches the
correct ChromeDriver automatically in ~/.cache/selenium — no external driver
packages or manual setup required.

Usage:
    python teamstrawler.py chat.txt
    python teamstrawler.py chat.txt --no-resume
"""

import time
import sys
import hashlib
import argparse
import os
import json
import logging

from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.common.exceptions import WebDriverException, TimeoutException


# ---------------------------------------------------------------------------
# Constants
# ---------------------------------------------------------------------------

SCROLL_STEP_RATIO    = 0.6   # fraction of clientHeight scrolled per step
EMPTY_LOOP_LIMIT     = 12    # consecutive empty loops before stopping
MAX_SCROLL_STALLS    = 6     # consecutive no-movement scrolls before stopping
AUTOSAVE_INTERVAL    = 5     # save + checkpoint every N *productive* loops
SCROLL_WAIT_TIMEOUT  = 5.0   # max seconds to wait for DOM update after scroll
SCROLL_WAIT_POLL     = 0.3   # polling interval for DOM-ready check


# ---------------------------------------------------------------------------
# Logging
# ---------------------------------------------------------------------------

def setup_logging():
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s | %(levelname)s | %(message)s",
        datefmt="%H:%M:%S",
    )


# ---------------------------------------------------------------------------
# In-page status overlay
# ---------------------------------------------------------------------------

def inject_status_box(driver, text, color="red"):
    js = """
    var box = document.getElementById('py-status');
    if (!box) {
        box = document.createElement('div');
        box.id = 'py-status';
        box.style.position = 'fixed';
        box.style.top = '10px';
        box.style.right = '10px';
        box.style.zIndex = '1000000';
        box.style.padding = '10px';
        box.style.fontSize = '16px';
        box.style.fontWeight = 'bold';
        box.style.color = 'white';
        box.style.borderRadius = '5px';
        box.style.boxShadow = '0 2px 8px rgba(0,0,0,0.3)';
        document.body.appendChild(box);
    }
    box.style.backgroundColor = arguments[1];
    box.innerText = arguments[0];
    """
    try:
        driver.execute_script(js, text, color)
    except Exception as exc:
        logging.debug("Status box injection failed: %s", exc)


# ---------------------------------------------------------------------------
# Hashing — stable fields only, no DOM-derived extras
# ---------------------------------------------------------------------------

def make_hash(content: str, timestamp: str, author: str) -> str:
    """Hash only stable, human-visible fields to avoid phantom duplicates."""
    raw = f"{author}|{timestamp}|{content}"
    return hashlib.sha256(raw.encode("utf-8")).hexdigest()


# ---------------------------------------------------------------------------
# Checkpoint (crash recovery)
# ---------------------------------------------------------------------------

def checkpoint_path(output_file: str) -> str:
    base, _ = os.path.splitext(output_file)
    return base + ".checkpoint.json"


def checkpoint_exists(output_file: str) -> bool:
    return os.path.exists(checkpoint_path(output_file))


def load_checkpoint(output_file: str) -> tuple[set, list]:
    cp = checkpoint_path(output_file)
    if not os.path.exists(cp):
        return set(), []
    try:
        with open(cp, "r", encoding="utf-8") as f:
            data = json.load(f)
        seen  = set(data.get("seen_hashes", []))
        saved = data.get("ordered_list", [])
        logging.info("Resumed from checkpoint: %d hashes, %d messages.", len(seen), len(saved))
        return seen, saved
    except Exception as exc:
        logging.warning("Could not read checkpoint, starting fresh: %s", exc)
        return set(), []


def save_checkpoint(output_file: str, seen_hashes: set, ordered_list: list):
    cp = checkpoint_path(output_file)
    try:
        with open(cp, "w", encoding="utf-8") as f:
            json.dump(
                {"seen_hashes": list(seen_hashes), "ordered_list": ordered_list},
                f,
                ensure_ascii=False,
            )
    except Exception as exc:
        logging.warning("Checkpoint save failed: %s", exc)


def delete_checkpoint(output_file: str):
    cp = checkpoint_path(output_file)
    try:
        if os.path.exists(cp):
            os.remove(cp)
            logging.info("Checkpoint removed after successful run.")
    except Exception:
        pass


# ---------------------------------------------------------------------------
# Chrome connection — Selenium Manager handles driver automatically
# ---------------------------------------------------------------------------

def connect_to_existing_chrome(chrome_binary: str | None = None) -> webdriver.Chrome:
    """
    Connect to a Chrome instance already running with --remote-debugging-port=9222.

    Passing Service() with no arguments tells Selenium to invoke its built-in
    Selenium Manager, which downloads and caches the correct ChromeDriver
    version for the detected Chrome installation. No external packages or
    manual driver management needed.

    chrome_binary: optional path to the Chrome executable. Useful on Windows
    or when Chrome is installed in a non-standard location. If omitted,
    Selenium Manager auto-detects the system Chrome.

    Platform defaults (only needed when auto-detection fails):
      macOS:   /Applications/Google Chrome.app/Contents/MacOS/Google Chrome
      Windows: C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe
      Linux:   /usr/bin/google-chrome
    """
    chrome_options = Options()
    chrome_options.add_experimental_option("debuggerAddress", "127.0.0.1:9222")
    if chrome_binary:
        if not os.path.exists(chrome_binary):
            raise RuntimeError(f"Chrome binary not found at: {chrome_binary}")
        chrome_options.binary_location = chrome_binary
        logging.info("Using Chrome binary: %s", chrome_binary)
    try:
        return webdriver.Chrome(service=Service(), options=chrome_options)
    except Exception as exc:
        raise RuntimeError(
            f"Could not connect to Chrome on port 9222: {exc}"
        ) from exc


# ---------------------------------------------------------------------------
# DOM utilities
# ---------------------------------------------------------------------------

def enable_mouse_tracking(driver):
    driver.execute_script("""
        window.lastHovered = null;
        document.addEventListener('mouseover', function(e) {
            window.lastHovered = e.target;
        }, true);
    """)


def find_scrollable_ancestor(driver, element):
    js = """
    var el = arguments[0];
    while (el) {
        var style = window.getComputedStyle(el);
        var ov = style.overflowY;
        if ((ov === 'auto' || ov === 'scroll') && el.scrollHeight > el.clientHeight)
            return el;
        el = el.parentElement;
        if (!el || el.tagName === 'BODY') break;
    }
    return null;
    """
    return driver.execute_script(js, element)


def try_find_chat_container_automatically(driver):
    """
    Try known Teams selectors in priority order; includes a generic
    aria fallback for resilience against Teams UI updates.
    """
    selectors = [
        '[data-tid="chat-pane-list"]',
        '[data-tid="message-pane-list-viewport"]',
        '[role="log"]',
        '[aria-label*="Chat"]',
        '[role="main"]',
    ]
    for selector in selectors:
        try:
            for el in driver.find_elements(By.CSS_SELECTOR, selector):
                scrollable = find_scrollable_ancestor(driver, el) or el
                sh = driver.execute_script("return arguments[0].scrollHeight", scrollable)
                ch = driver.execute_script("return arguments[0].clientHeight", scrollable)
                if sh and ch and sh > ch:
                    return scrollable
        except Exception:
            continue
    return None


def get_hover_target(driver):
    return driver.execute_script("return window.lastHovered;")


def highlight_element(driver, element, color="red"):
    driver.execute_script(
        "arguments[0].style.outline = '4px solid %s';"
        "arguments[0].style.outlineOffset = '-2px';" % color,
        element,
    )


# ---------------------------------------------------------------------------
# Message extraction
# ---------------------------------------------------------------------------

def extract_messages(driver, scroll_element) -> list[dict]:
    """
    Extract visible messages from the scroll container.
    Returns dicts with 'author', 'timestamp', 'content' only —
    no DOM ids or indices, which change between scroll positions.
    Results are returned in DOM order (top-to-bottom = older-to-newer).
    """
    js = """
    const root = arguments[0];

    function clean(text) {
        return (text || '').replace(/\\s+/g, ' ').trim();
    }

    const itemSelectors = [
        '[data-tid="chat-pane-item"]',
        '[data-tid="message-body-content"]',
        '[class*="ChatItem"]',
        '[role="listitem"]',
    ];

    let items = [];
    for (const sel of itemSelectors) {
        items = Array.from(root.querySelectorAll(sel));
        if (items.length) break;
    }

    const bodySelectors = [
        '[data-tid="chat-pane-message"]',
        '[data-tid="message-body-content"]',
        '[data-track-module-name="messageBody"]',
        '[class*="messageBody"]',
    ];

    const results = [];

    for (const item of items) {
        try {
            const author =
                clean(item.querySelector('[data-tid="message-author-name"]')?.innerText) ||
                clean(item.querySelector('[data-tid="threadBodyDisplayName"]')?.innerText) ||
                clean(item.querySelector('[aria-label]')?.getAttribute('aria-label')) ||
                'System';

            const timeEl = item.querySelector('time');
            const timestamp =
                clean(timeEl?.getAttribute('title')) ||
                clean(timeEl?.getAttribute('datetime')) ||
                clean(timeEl?.innerText) ||
                '';

            let body = null;
            for (const sel of bodySelectors) {
                body = item.querySelector(sel);
                if (body) break;
            }
            if (!body) continue;

            const content = clean(body.innerText);
            if (!content) continue;

            results.push({ author, timestamp, content });
        } catch (_) {}
    }

    return results;
    """
    return driver.execute_script(js, scroll_element)


# ---------------------------------------------------------------------------
# Scroll helpers
# ---------------------------------------------------------------------------

def get_scroll_state(driver, scroll_element) -> dict:
    return driver.execute_script(
        "return {"
        "  scrollTop:    arguments[0].scrollTop,"
        "  scrollHeight: arguments[0].scrollHeight,"
        "  clientHeight: arguments[0].clientHeight"
        "};",
        scroll_element,
    )


def scroll_up_once(driver, scroll_element):
    driver.execute_script(
        """
        const el   = arguments[0];
        const step = Math.max(200, Math.floor(el.clientHeight * arguments[1]));
        el.scrollTop = Math.max(0, el.scrollTop - step);
        """,
        scroll_element,
        SCROLL_STEP_RATIO,
    )


def wait_for_scroll_update(driver, scroll_element, prev_scroll_top: float) -> dict:
    """
    Poll until scrollTop changes (Teams has loaded older messages)
    or the timeout is reached. Always returns the final scroll state.
    """
    deadline = time.monotonic() + SCROLL_WAIT_TIMEOUT
    while time.monotonic() < deadline:
        state = get_scroll_state(driver, scroll_element)
        if abs(state["scrollTop"] - prev_scroll_top) > 1:
            return state
        time.sleep(SCROLL_WAIT_POLL)
    return get_scroll_state(driver, scroll_element)


# ---------------------------------------------------------------------------
# File I/O
# ---------------------------------------------------------------------------

def format_message(msg: dict) -> str:
    return f"[{msg['timestamp']}] {msg['author']}: {msg['content']}"


def save_messages(filename: str, ordered_list: list):
    """Write messages to file. ordered_list must already be in chronological order."""
    with open(filename, "w", encoding="utf-8") as f:
        f.write("\n".join(ordered_list))


def resolve_start_mode(output_file: str, no_resume: bool) -> bool:
    """
    Decide whether to resume from checkpoint or start fresh.

    FIX: checkpoint state is evaluated *before* asking about overwriting the
    output file, so a resumable run is not unnecessarily blocked by the
    overwrite prompt.

    Returns True if resuming, False if starting fresh.
    """
    if no_resume:
        logging.info("Starting fresh (--no-resume).")
        # Still warn if output file exists and will be overwritten at the end.
        if os.path.exists(output_file):
            answer = input(
                f"\nWARNING: '{output_file}' already exists and --no-resume was set. "
                f"Overwrite at end of run? [y/N] "
            ).strip().lower()
            if answer != "y":
                print("Aborted.")
                sys.exit(0)
        return False

    if checkpoint_exists(output_file):
        answer = input(
            f"\nCheckpoint found for '{output_file}'. Resume previous run? [Y/n] "
        ).strip().lower()
        if answer in ("", "y"):
            return True
        # User declined resume — ask about overwrite before proceeding.
        if os.path.exists(output_file):
            answer2 = input(
                f"'{output_file}' already exists. Overwrite? [y/N] "
            ).strip().lower()
            if answer2 != "y":
                print("Aborted.")
                sys.exit(0)
        return False

    # No checkpoint — only ask about overwrite if the file already exists.
    if os.path.exists(output_file):
        answer = input(
            f"\nWARNING: '{output_file}' already exists. Overwrite? [y/N] "
        ).strip().lower()
        if answer != "y":
            print("Aborted.")
            sys.exit(0)
    return False


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

def main():
    setup_logging()

    parser = argparse.ArgumentParser(
        description="TeamsTrawler — scroll Teams chat history and save to a file."
    )
    parser.add_argument("output_file", help="Output file path, e.g. chat.txt")
    parser.add_argument(
        "--no-resume",
        action="store_true",
        help="Ignore any existing checkpoint and start from scratch.",
    )
    parser.add_argument(
        "--chrome-path",
        default=None,
        metavar="PATH",
        help=(
            "Path to the Chrome executable. Only needed when auto-detection fails. "
            "macOS default: '/Applications/Google Chrome.app/Contents/MacOS/Google Chrome'  "
            "Windows default: 'C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe'  "
            "Linux default: '/usr/bin/google-chrome'"
        ),
    )
    args = parser.parse_args()

    output_file = args.output_file

    # FIX: resolve resume/overwrite before connecting to Chrome.
    resuming = resolve_start_mode(output_file, args.no_resume)

    print("--- TeamsTrawler ---")

    try:
        driver = connect_to_existing_chrome(chrome_binary=args.chrome_path)
    except RuntimeError as exc:
        print(f"ERROR: {exc}")
        sys.exit(1)

    try:
        WebDriverWait(driver, 10).until(
            lambda d: d.execute_script("return document.readyState") == "complete"
        )
    except TimeoutException:
        logging.warning("Page did not finish loading in time — continuing anyway.")

    enable_mouse_tracking(driver)

    scroll_element = try_find_chat_container_automatically(driver)

    if scroll_element:
        highlight_element(driver, scroll_element, "green")
        inject_status_box(driver, "AUTO-FOUND CHAT CONTAINER", "green")
        print("Chat container found automatically.")
        time.sleep(1)
    else:
        inject_status_box(driver, "HOVER OVER CHAT AND PRESS ENTER", "blue")
        input("\n[STEP 1] Hover over the chat area and press ENTER…")

        target = get_hover_target(driver)
        if not target:
            print("Error: no hover target found.")
            sys.exit(1)

        scroll_element = find_scrollable_ancestor(driver, target)
        if not scroll_element:
            print("Error: no scrollable container found from hover target.")
            sys.exit(1)

        highlight_element(driver, scroll_element, "red")

    # FIX: load checkpoint only when resuming; otherwise start fresh.
    # ordered_list stores dicts so ordering can be fixed up before formatting.
    if resuming:
        seen_hashes, ordered_list = load_checkpoint(output_file)
        # checkpoint stores pre-formatted strings for backward compat;
        # convert to dicts if needed (new runs will store dicts directly)
        if ordered_list and isinstance(ordered_list[0], str):
            logging.info(
                "Checkpoint contains pre-formatted strings from an older run. "
                "Ordering correction cannot be applied retrospectively."
            )
    else:
        seen_hashes: set   = set()
        ordered_list: list = []  # list of formatted strings, newest-to-oldest during crawl

    empty_loops      = 0
    scroll_stalls    = 0
    loop             = 0
    productive_loops = 0  # FIX: track loops that actually found new messages
    prev_state       = get_scroll_state(driver, scroll_element)

    while True:
        loop += 1
        new_found = 0

        try:
            messages = extract_messages(driver, scroll_element)
        except WebDriverException as exc:
            logging.error("Message extraction failed on loop %d: %s", loop, exc)
            messages = []

        if not messages and loop > 1:
            logging.warning(
                "Loop %d: no messages extracted — Teams DOM selectors may need updating.",
                loop,
            )

        # FIX: collect new messages from this viewport as an ordered batch.
        # extract_messages returns items in DOM order (top=older, bottom=newer).
        # Since we're scrolling upward, this batch is older than anything
        # appended in previous loops. Prepend the whole batch so that
        # ordered_list stays in newest-to-oldest order throughout the crawl
        # (it is reversed to chronological order just once at the very end).
        new_batch = []
        for msg in messages:
            try:
                content   = msg["content"]
                timestamp = msg["timestamp"]
                author    = msg["author"]
                msg_hash  = make_hash(content, timestamp, author)

                if msg_hash not in seen_hashes:
                    seen_hashes.add(msg_hash)
                    new_batch.append(format_message(msg))
                    new_found += 1
            except Exception as exc:
                logging.debug("Single message parse failed: %s", exc)

        if new_batch:
            # new_batch is in DOM order (older → newer within this viewport).
            # Prepend it so ordered_list remains newest-first overall.
            ordered_list = new_batch + ordered_list

        if new_found == 0:
            empty_loops += 1
            status_color = "orange"
        else:
            empty_loops      = 0
            productive_loops += 1  # FIX: only count loops that found something
            status_color     = "red"

        status = (
            f"Loop {loop} | Saved: {len(ordered_list)} | "
            f"Empty: {empty_loops}/{EMPTY_LOOP_LIMIT} | "
            f"Stalls: {scroll_stalls}/{MAX_SCROLL_STALLS}"
        )
        print(f"{status:<100}", end="\r")
        inject_status_box(driver, status, status_color)

        # FIX: autosave every N *productive* loops, not every N total loops.
        if productive_loops > 0 and productive_loops % AUTOSAVE_INTERVAL == 0:
            try:
                # ordered_list is already oldest-first (prepend strategy keeps it that way).
                save_messages(output_file, ordered_list)
                save_checkpoint(output_file, seen_hashes, ordered_list)
            except Exception as exc:
                logging.warning("Autosave failed: %s", exc)

        # Exit condition 1: too many consecutive loops with no new content
        if empty_loops >= EMPTY_LOOP_LIMIT:
            print("\n\nFinished: no new messages after too many consecutive loops.")
            break

        prev_scroll_top = prev_state.get("scrollTop", 0)

        try:
            scroll_up_once(driver, scroll_element)
        except Exception as exc:
            logging.error("Scroll error on loop %d: %s", loop, exc)

        try:
            new_state = wait_for_scroll_update(driver, scroll_element, prev_scroll_top)
        except Exception as exc:
            logging.error("Scroll state read failed: %s", exc)
            new_state = prev_state

        if abs(new_state.get("scrollTop", 0) - prev_scroll_top) >= 2:
            scroll_stalls = 0
        else:
            scroll_stalls += 1

        # Exit condition 2: at the top and scroll is no longer moving
        at_top = new_state.get("scrollTop", 0) <= 2
        if at_top and scroll_stalls >= MAX_SCROLL_STALLS:
            print("\n\nFinished: reached top of chat and scroll is no longer moving.")
            break

        prev_state = new_state

    # Final save in chronological order (oldest first).
    # ordered_list is already oldest-first: prepending older batches throughout
    # the crawl maintains that order naturally. No reversal needed.
    save_messages(output_file, ordered_list)
    delete_checkpoint(output_file)

    inject_status_box(driver, f"DONE! Saved {len(ordered_list)} messages.", "green")
    print(f"\nSaved {len(ordered_list)} messages to '{output_file}'.")


if __name__ == "__main__":
    main()
