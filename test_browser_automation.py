#!/usr/bin/env python3
"""
Test 5: Browser Automation for Google Docs Comments

Uses Playwright to automate the Google Docs UI via keyboard shortcuts.

WARNINGS:
- This is fragile and may break with Google Docs updates
- May violate Google's Terms of Service
- Requires manual Google login (can't automate OAuth easily)
- Use at your own risk for testing/research purposes only

SETUP:
    pip install playwright
    playwright install chromium

USAGE:
    python test_browser_automation.py <google_doc_url> <target_text> <comment_text>

WORKFLOW:
    1. Opens Google Doc in browser (you may need to log in manually first time)
    2. Uses Cmd+F / Ctrl+F to find target text
    3. Uses keyboard to select the found text
    4. Uses Cmd+Option+M / Ctrl+Alt+M to open comment dialog
    5. Types comment and submits
"""

import sys
import time
import platform
from playwright.sync_api import sync_playwright, TimeoutError as PlaywrightTimeout


def get_modifier_key():
    """Return the correct modifier key for the OS."""
    if platform.system() == "Darwin":
        return "Meta"  # Cmd on Mac
    return "Control"  # Ctrl on Windows/Linux


def add_comment_to_google_doc(doc_url: str, target_text: str, comment_text: str, headless: bool = False):
    """
    Automate adding a comment to a Google Doc.

    Args:
        doc_url: Full URL to the Google Doc
        target_text: Text to find and comment on
        comment_text: The comment to add
        headless: Run without visible browser (default False for debugging)
    """
    mod = get_modifier_key()

    with sync_playwright() as p:
        # Launch browser with persistent context to reuse login
        # User data dir stores cookies/session
        user_data_dir = "/tmp/playwright-google-session"

        print(f"\n=== Browser Automation Test ===")
        print(f"Document: {doc_url}")
        print(f"Target: '{target_text}'")
        print(f"Comment: '{comment_text}'")
        print(f"Modifier key: {mod}")
        print()

        browser = p.chromium.launch_persistent_context(
            user_data_dir,
            headless=headless,
            slow_mo=100,  # Slow down for visibility
        )

        page = browser.pages[0] if browser.pages else browser.new_page()

        try:
            # Navigate to the document
            print("1. Opening document...")
            page.goto(doc_url, wait_until="networkidle", timeout=60000)

            # Wait for the document to be interactive
            # Google Docs takes a while to fully load
            print("2. Waiting for document to load...")
            time.sleep(3)

            # Click somewhere in the document to focus it
            print("3. Focusing document...")
            # Try to click on the document canvas/editor area
            try:
                page.click(".kix-appview-editor", timeout=5000)
            except:
                # Fallback: click in the middle of the page
                page.mouse.click(500, 400)

            time.sleep(0.5)

            # Open Find dialog
            print(f"4. Opening Find dialog ({mod}+F)...")
            page.keyboard.press(f"{mod}+f")
            time.sleep(0.5)

            # Type the search text
            print(f"5. Searching for '{target_text}'...")
            page.keyboard.type(target_text, delay=50)
            time.sleep(0.5)

            # Press Enter to find
            page.keyboard.press("Enter")
            time.sleep(0.5)

            # Close find dialog with Escape
            print("6. Closing Find dialog...")
            page.keyboard.press("Escape")
            time.sleep(0.3)

            # The found text should now be selected/highlighted
            # But we need to actually select it for commenting
            # Try using Shift+End or selecting via Find & Replace

            # Alternative approach: Use Find & Replace which keeps selection
            print("7. Selecting found text...")

            # On Mac, the shortcut to add comment is Cmd+Option+M
            # On Windows/Linux, it's Ctrl+Alt+M
            if mod == "Meta":
                comment_shortcut = "Meta+Alt+m"
            else:
                comment_shortcut = "Control+Alt+m"

            print(f"8. Opening comment dialog ({comment_shortcut})...")
            page.keyboard.press(comment_shortcut)
            time.sleep(1)

            # Type the comment
            print(f"9. Typing comment...")
            page.keyboard.type(comment_text, delay=30)
            time.sleep(0.3)

            # Submit the comment
            print("10. Submitting comment...")
            # Try Cmd/Ctrl+Enter to submit, or just click the button
            page.keyboard.press(f"{mod}+Enter")
            time.sleep(1)

            print("\n✓ Comment sequence completed!")
            print("Check the document to see if the comment was added.")

            # Keep browser open for inspection
            print("\nPress Enter to close browser...")
            input()

        except PlaywrightTimeout as e:
            print(f"\n✗ Timeout error: {e}")
            print("The page may not have loaded correctly.")
            input("Press Enter to close browser...")

        except Exception as e:
            print(f"\n✗ Error: {e}")
            input("Press Enter to close browser...")

        finally:
            browser.close()


def main():
    if len(sys.argv) < 4:
        print("Usage: python test_browser_automation.py <google_doc_url> <target_text> <comment_text>")
        print()
        print("Example:")
        print('  python test_browser_automation.py "https://docs.google.com/document/d/xxx/edit" "quick brown fox" "Test comment"')
        print()
        print("NOTE: You may need to log into Google the first time.")
        print("The script uses a persistent browser profile to remember your session.")
        sys.exit(1)

    doc_url = sys.argv[1]
    target_text = sys.argv[2]
    comment_text = sys.argv[3]

    add_comment_to_google_doc(doc_url, target_text, comment_text)


if __name__ == "__main__":
    main()
