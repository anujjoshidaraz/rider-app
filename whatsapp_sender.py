"""
WhatsApp Web automation using Playwright (async).
Sends messages to WhatsApp groups/contacts from a persistent browser profile.
Uses Playwright instead of Selenium for better reliability with WhatsApp's React UI.
"""

import asyncio
import os
import re
import subprocess
import threading
import time
from pathlib import Path

# Persistent profile directory
PROFILE_DIR = os.path.join(os.path.expanduser("~"), ".qoderwork", "whatsapp_profile")

# Global browser context
_playwright = None
_context = None
_page = None
_lock = threading.Lock()
_loop = None
_loop_thread = None


def _kill_stale_browsers():
    """Kill any stale Playwright Chromium processes that might block launch."""
    try:
        # Kill Playwright's chromium processes (not Edge/Chrome user browsers)
        subprocess.run(
            ['powershell', '-Command',
             "Get-Process chromium -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue"],
            capture_output=True, timeout=10
        )
    except Exception as e:
        print(f"Warning: Could not kill stale browsers: {e}")

    # Clean up lock files that prevent launch
    for lock_name in ['SingletonLock', 'SingletonSocket', 'SingletonCookie']:
        lock_path = os.path.join(PROFILE_DIR, lock_name)
        try:
            if os.path.exists(lock_path):
                os.remove(lock_path)
        except:
            pass


def _start_event_loop():
    """Run an event loop in a dedicated thread for async Playwright."""
    global _loop
    _loop = asyncio.new_event_loop()
    asyncio.set_event_loop(_loop)
    _loop.run_forever()


def _ensure_loop():
    """Ensure the event loop thread is running."""
    global _loop_thread
    if _loop_thread is None or not _loop_thread.is_alive():
        _loop_thread = threading.Thread(target=_start_event_loop, daemon=True)
        _loop_thread.start()
        time.sleep(0.5)  # Give loop time to start


def _run_async(coro):
    """Run an async coroutine in the dedicated loop thread and wait for result."""
    _ensure_loop()
    future = asyncio.run_coroutine_threadsafe(coro, _loop)
    return future.result(timeout=90)


async def _get_browser():
    """Get or create the Playwright browser with persistent context."""
    global _playwright, _context, _page

    if _context is not None:
        try:
            # Check if still alive by accessing pages
            _ = _context.pages
            if _page and not _page.is_closed():
                return _playwright, _context, _page
        except:
            pass
        # Context is dead, clean up
        try:
            await _context.close()
        except:
            pass
        _playwright = None
        _context = None
        _page = None

    # Kill stale browsers and clean locks
    _kill_stale_browsers()

    # Ensure profile directory exists
    os.makedirs(PROFILE_DIR, exist_ok=True)

    from playwright.async_api import async_playwright

    _playwright = await async_playwright().start()

    # Launch with persistent context — NOT headless (WhatsApp detects headless)
    # Window is positioned off-screen so user doesn't see it
    _context = await _playwright.chromium.launch_persistent_context(
        user_data_dir=PROFILE_DIR,
        headless=False,
        args=[
            '--no-sandbox',
            '--disable-dev-shm-usage',
            '--disable-gpu',
            '--disable-extensions',
            '--disable-background-networking',
            '--disable-sync',
            '--no-first-run',
            '--window-position=-10000,0',
            '--window-size=1200,900',
        ]
    )

    # Get or create page
    if _context.pages:
        _page = _context.pages[0]
    else:
        _page = await _context.new_page()

    return _playwright, _context, _page


async def _async_is_logged_in():
    """Async check if WhatsApp Web is logged in."""
    try:
        _, _, page = await _get_browser()
        await page.goto('https://web.whatsapp.com', wait_until='networkidle', timeout=30000)

        # Wait for either login (pane-side) or QR code
        try:
            await page.wait_for_selector('#pane-side', timeout=15000)
            return True
        except:
            # Check if we're on QR login page
            try:
                qr = await page.query_selector('[data-icon="qr-code"]')
                if qr:
                    return False
            except:
                pass
            return False
    except Exception as e:
        print(f"WhatsApp check failed: {e}")
        return False


def is_logged_in():
    """Check if WhatsApp Web is logged in (sync wrapper)."""
    print("[WA] is_logged_in called")
    try:
        result = _run_async(_async_is_logged_in())
        print(f"[WA] is_logged_in result: {result}")
        return result
    except Exception as e:
        print(f"[WA] is_logged_in exception: {e}")
        return False


def _is_phone_query(query):
    """Detect if query is a phone number (vs group name)."""
    digits = re.sub(r'\D', '', query)
    if len(digits) >= 8:
        digit_ratio = len(digits) / len(query.replace(' ', ''))
        if digit_ratio >= 0.6:
            return True
    return False


def _normalize_name(name):
    """Normalize group name for comparison."""
    return re.sub(r'[^a-z0-9]', '', name.lower())


async def _async_send_message_to_group(group_name, message):
    """Async send a message to a WhatsApp group or contact."""
    try:
        _, _, page = await _get_browser()

        # Navigate to WhatsApp Web
        await page.goto('https://web.whatsapp.com', wait_until='networkidle', timeout=30000)

        # Check if logged in
        try:
            await page.wait_for_selector('#pane-side', timeout=10000)
        except:
            return {"success": False, "error": "Not logged in to WhatsApp Web. Please run setup first."}

        # Click on search box
        search_box = page.locator('div[contenteditable="true"][data-tab="3"]')
        await search_box.click()
        await asyncio.sleep(0.5)

        # Type the group/contact name
        await search_box.fill(group_name)
        await asyncio.sleep(2)  # Wait for search results

        is_phone = _is_phone_query(group_name)

        if is_phone:
            # For phone numbers, click first search result with force=True
            try:
                first_result = page.locator('span[title]').first
                await first_result.click(force=True, timeout=5000)
            except:
                return {"success": False, "error": f"Phone number '{group_name}' not found"}
        else:
            # For group names, find exact or close match
            normalized_target = _normalize_name(group_name)
            found = False

            # Try exact title match first
            try:
                group_item = page.locator(f'span[title="{group_name}"]').first
                await group_item.click(force=True, timeout=5000)
                found = True
            except:
                pass

            if not found:
                # Try partial match
                try:
                    spans = await page.locator('span[title]').all()
                    for span in spans:
                        title = await span.get_attribute('title')
                        if title and (normalized_target in _normalize_name(title) or
                                     _normalize_name(title) in normalized_target):
                            await span.click(force=True)
                            found = True
                            break
                except:
                    pass

            if not found:
                return {"success": False, "error": f"Group '{group_name}' not found"}

        await asyncio.sleep(1)

        # Find message input box
        message_box = page.locator('div[contenteditable="true"][data-tab="10"]')
        await message_box.wait_for(timeout=10000)

        # Clear existing text
        await message_box.fill('')
        await asyncio.sleep(0.3)

        # Type message, handling newlines
        lines = message.split('\n')
        for i, line in enumerate(lines):
            await message_box.type(line, delay=10)
            if i < len(lines) - 1:
                await page.keyboard.press('Shift+Enter')
        await asyncio.sleep(0.5)

        # Send message (Enter)
        await page.keyboard.press('Enter')
        await asyncio.sleep(1)

        return {"success": True, "error": None}

    except Exception as e:
        return {"success": False, "error": str(e)}


def send_message_to_group(group_name, message):
    """Send a message to a WhatsApp group (sync wrapper)."""
    return _run_async(_async_send_message_to_group(group_name, message))


async def _async_setup_login():
    """Async setup: open visible browser for QR code scan."""
    _kill_stale_browsers()
    os.makedirs(PROFILE_DIR, exist_ok=True)

    from playwright.async_api import async_playwright

    pw = await async_playwright().start()

    # Launch visible browser for QR scan (centered on screen)
    context = await pw.chromium.launch_persistent_context(
        user_data_dir=PROFILE_DIR,
        headless=False,
        args=[
            '--no-sandbox',
            '--disable-dev-shm-usage',
            '--disable-gpu',
            '--disable-extensions',
            '--no-first-run',
            '--window-position=200,100',
            '--window-size=1000,800',
        ]
    )

    page = context.pages[0] if context.pages else await context.new_page()
    await page.goto('https://web.whatsapp.com')

    print("\n" + "="*60)
    print("WHATSAPP WEB LOGIN")
    print("="*60)
    print("1. Scan the QR code with your WhatsApp mobile app")
    print("2. Wait until you see your chats loaded")
    print("3. Then close this browser window")
    print("="*60 + "\n")

    # Wait for login
    try:
        await page.wait_for_selector('#pane-side', timeout=120000)
        print("\nLogin successful! You can close the browser now.")
        print("  The session will be saved for future use.\n")
        await asyncio.sleep(3)
    except:
        print("\nLogin timeout. Please try again.\n")

    await context.close()
    await pw.stop()


def setup_login():
    """Open WhatsApp Web for initial login (QR code scan)."""
    asyncio.run(_async_setup_login())


if __name__ == "__main__":
    import sys
    if len(sys.argv) > 1 and sys.argv[1] == "setup":
        setup_login()
    elif len(sys.argv) > 2 and sys.argv[1] == "send":
        group = sys.argv[2]
        msg = sys.argv[3] if len(sys.argv) > 3 else "Test message from Barcode Tracker"
        result = send_message_to_group(group, msg)
        print(f"Result: {result}")
    else:
        print("Usage:")
        print("  python whatsapp_sender.py setup          # Login to WhatsApp Web")
        print("  python whatsapp_sender.py send <group> <message>  # Send message")
