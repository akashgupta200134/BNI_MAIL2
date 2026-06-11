"""
core/email_sender.py — Orchestrates the full send campaign using Playwright.
Selectors verified via Playwright codegen on this exact Outlook account.
Image is embedded inline using base64 JS injection — no toolbar clicks needed.
"""

import time
import os
import base64
import mimetypes
from playwright.sync_api import sync_playwright, TimeoutError as PWTimeout, Page
from config import Config
from utils.validator import is_valid_email
from utils.excel_handler import ExcelHandler
from utils.state_manager import StateManager


class EmailSender:

    def __init__(self, excel: ExcelHandler, state_manager: StateManager, logger, dry_run: bool = False):
        self.excel         = excel
        self.state_manager = state_manager
        self.logger        = logger
        self.dry_run       = dry_run

    # ─────────────────────────────────────────────────────────────────────────
    # Public entry point
    # ─────────────────────────────────────────────────────────────────────────

    def run(self, pending_rows: list, subject: str, template: str, image_path: str = None):
        state       = self.state_manager.load()
        sent_today  = state["count"]
        batch_count = 0

        with sync_playwright() as p:
            browser = p.chromium.launch(headless=Config.HEADLESS)
            context = browser.new_context()
            page    = context.new_page()
            page.set_default_timeout(Config.PAGE_TIMEOUT)

            self._do_login(page)

            print(f"\n🚀 Starting campaign. {len(pending_rows)} emails to process.\n")

            for row in pending_rows:
                if sent_today >= Config.DAILY_LIMIT:
                    print(f"\n🚫 Daily limit of {Config.DAILY_LIMIT} reached. Stopping.")
                    self.logger.warning("Daily limit reached mid-campaign.")
                    break

                name_cell   = row[Config.COL_NAME]
                email_cell  = row[Config.COL_EMAIL]
                status_cell = row[Config.COL_STATUS]
                remark_cell = row[Config.COL_REMARK]

                name  = str(name_cell.value or "").strip()
                email = str(email_cell.value or "").strip()

                print(f"📧 [{sent_today + 1}/{Config.DAILY_LIMIT}] Processing: {name} <{email}>")

                if not is_valid_email(email):
                    status_cell.value = Config.STATUS_INVALID
                    self.excel.save()
                    self.logger.warning(f"INVALID_EMAIL | {name} | {email}")
                    print(f"   ⚠️  Invalid email format — skipped, marked 'Invalid Email'")
                    continue

                body = template.replace("{{Name}}", name).replace("{{Email}}", email)

                success = self._send_with_retry(page, email, subject, body, name, image_path)

                if success:
                    remark_cell.value = Config.REMARK_DONE
                    sent_today += 1
                    batch_count += 1
                    state["count"] = sent_today
                    self.state_manager.save(state)
                    self.excel.save()
                    self.logger.info(f"SENT | {name} | {email}")
                    print(f"   ✅ Sent successfully ({sent_today} today)")
                else:
                    status_cell.value = Config.STATUS_NOT_FOUND
                    self.excel.save()
                    self.logger.error(f"FAILED | {name} | {email}")
                    print(f"   ❌ Failed — marked 'Not Found'")

                if batch_count > 0 and batch_count % Config.BATCH_SIZE == 0:
                    print(f"\n⏸️  Batch of {Config.BATCH_SIZE} complete. Pausing {Config.BATCH_PAUSE}s...\n")
                    self.logger.info(f"Batch pause after {batch_count} emails")
                    time.sleep(Config.BATCH_PAUSE)
                elif sent_today < Config.DAILY_LIMIT:
                    remaining = len(pending_rows) - (pending_rows.index(row) + 1)
                    if remaining > 0:
                        print(f"   ⏳ Waiting {Config.DELAY_BETWEEN_EMAILS}s before next email...")
                        time.sleep(Config.DELAY_BETWEEN_EMAILS)

            browser.close()

    # ─────────────────────────────────────────────────────────────────────────
    # Login
    # ─────────────────────────────────────────────────────────────────────────

    def _do_login(self, page: Page):
        print("\n🌐 Opening Outlook Web...")
        page.goto(Config.OUTLOOK_URL)

        try:
            page.wait_for_selector('[aria-label="New mail"], [name="New mail"]', timeout=5000)
            print("✅ Already logged in.\n")
            return
        except PWTimeout:
            pass

        if not Config.OUTLOOK_EMAIL or not Config.OUTLOOK_PASSWORD:
            raise RuntimeError(
                "Outlook credentials not found. "
                "Make sure OUTLOOK_EMAIL and OUTLOOK_PASSWORD are set in your .env file."
            )

        print(f"🔐 Auto-logging in as: {Config.OUTLOOK_EMAIL}")

        try:
            page.get_by_role("textbox", name="Enter your email, phone, or").click()
            page.get_by_role("textbox", name="Enter your email, phone, or").fill(Config.OUTLOOK_EMAIL)
            page.get_by_role("button", name="Next").click()

            page.get_by_role("textbox", name="Enter the password for").wait_for(timeout=15000)
            page.get_by_role("textbox", name="Enter the password for").fill(Config.OUTLOOK_PASSWORD)
            page.get_by_role("textbox", name="Enter the password for").press("Enter")

            try:
                page.get_by_role("button", name="No").wait_for(timeout=8000)
                page.get_by_role("button", name="No").click()
                self.logger.info("Dismissed 'Stay signed in?' prompt")
            except PWTimeout:
                pass

            page.get_by_role("button", name="New mail").wait_for(timeout=30000)
            print("✅ Login successful. Starting campaign...\n")
            self.logger.info(f"Logged in as {Config.OUTLOOK_EMAIL}")

        except PWTimeout:
            self._take_screenshot(page, "login_failed")
            raise RuntimeError(
                "Auto-login timed out.\n"
                "  • Check OUTLOOK_EMAIL and OUTLOOK_PASSWORD in .env\n"
                "  • MFA/2FA may be blocking the login\n"
                "  Screenshot saved to screenshots/login_failed.png"
            )
        except RuntimeError:
            raise
        except Exception as e:
            self._take_screenshot(page, "login_error")
            raise RuntimeError(f"Unexpected login error: {e}")

    # ─────────────────────────────────────────────────────────────────────────
    # Send with retry
    # ─────────────────────────────────────────────────────────────────────────

    def _send_with_retry(self, page: Page, to_email: str, subject: str, body: str, name: str, image_path: str = None) -> bool:
        for attempt in range(1, Config.COMPOSE_RETRIES + 2):
            try:
                self._compose_and_send(page, to_email, subject, body, image_path)
                return True
            except PWTimeout as e:
                self.logger.warning(f"TIMEOUT attempt {attempt} | {name} | {to_email} | {e}")
                print(f"   ⚠️  Timeout on attempt {attempt}/{Config.COMPOSE_RETRIES + 1}")
                if attempt <= Config.COMPOSE_RETRIES:
                    self._take_screenshot(page, f"timeout_{to_email}_{attempt}")
                    time.sleep(5)
                    self._recover_to_inbox(page)
            except Exception as e:
                self.logger.error(f"ERROR attempt {attempt} | {name} | {to_email} | {e}")
                print(f"   ⚠️  Error on attempt {attempt}: {e}")
                if attempt <= Config.COMPOSE_RETRIES:
                    self._take_screenshot(page, f"error_{to_email}_{attempt}")
                    time.sleep(5)
                    self._recover_to_inbox(page)
        return False

    # ─────────────────────────────────────────────────────────────────────────
    # Compose and send
    # ─────────────────────────────────────────────────────────────────────────

    def _compose_and_send(self, page: Page, to_email: str, subject: str, body: str, image_path: str = None):
        self._recover_to_inbox(page)

        # ── Open compose ──────────────────────────────────────────────────
        page.get_by_role("button", name="New mail").click()
        page.wait_for_timeout(1500)

        # ── To field ──────────────────────────────────────────────────────
        page.get_by_label("To", exact=True).click()
        page.wait_for_timeout(400)
        page.get_by_label("To", exact=True).fill(to_email)
        page.wait_for_timeout(1000)
        page.keyboard.press("Enter")
        page.wait_for_timeout(600)

        # ── Subject ───────────────────────────────────────────────────────
        page.get_by_placeholder("Add a subject").click()
        page.wait_for_timeout(300)
        page.get_by_placeholder("Add a subject").fill(subject)
        page.wait_for_timeout(400)

        # ── Body text ─────────────────────────────────────────────────────
        page.get_by_label("Message body").click()
        page.wait_for_timeout(500)

        for line in body.split("\n"):
            page.keyboard.type(line, delay=10)
            page.keyboard.press("Enter")

        page.wait_for_timeout(600)

        # ── Embed image inline via base64 JS injection ────────────────────
        # Does NOT click any toolbar buttons — injects <img> directly into
        # the contenteditable body div via JavaScript. This is the only
        # reliable method that works without OS file picker dialogs.
        if image_path and os.path.isfile(image_path):
            self._embed_image_via_js(page, image_path)
        elif image_path:
            print(f"   ⚠️  Image not found: {image_path} — sending without image")
            self.logger.warning(f"Image not found: {image_path}")

        # ── Send or discard ───────────────────────────────────────────────
        if self.dry_run:
            print("   🧪 DRY RUN — compose filled, not sending")
            try:
                page.get_by_role("button", name="Discard").click()
                page.wait_for_timeout(800)
            except Exception:
                page.keyboard.press("Escape")
                page.wait_for_timeout(800)
        else:
            page.get_by_label("Send", exact=True).click()
            page.wait_for_timeout(Config.SEND_WAIT)

    # ─────────────────────────────────────────────────────────────────────────
    # Embed image inline via JavaScript — NO toolbar clicks
    # ─────────────────────────────────────────────────────────────────────────

    def _embed_image_via_js(self, page: Page, filepath: str):
        """
        Injects a base64 <img> tag directly into Outlook's contenteditable
        body div using JavaScript. No toolbar buttons, no file picker dialogs.

        Flow:
          1. Read image file → encode as base64 data URI
          2. Find the body div via JS
          3. Insert <br><img src="data:..."> at the end of the body
          4. Fire an 'input' event so Outlook registers the change
        """
        filename = os.path.basename(filepath)
        print(f"   🖼️  Embedding image inline: {filename}")

        try:
            # Read and encode image
            with open(filepath, "rb") as f:
                image_bytes = f.read()

            mime_type, _ = mimetypes.guess_type(filepath)
            if not mime_type:
                mime_type = "image/png"

            b64_string = base64.b64encode(image_bytes).decode("utf-8")
            data_uri   = f"data:{mime_type};base64,{b64_string}"

            # Inject into Outlook body via JS
            success = page.evaluate("""
                (dataUri) => {
                    // Find Outlook's contenteditable body div
                    const body =
                        document.querySelector('[aria-label="Message body"][contenteditable="true"]') ||
                        document.querySelector('[contenteditable="true"][aria-multiline="true"]') ||
                        document.querySelector('[role="textbox"][contenteditable="true"]');

                    if (!body) return false;

                    // Move cursor to end
                    body.focus();
                    const range = document.createRange();
                    range.selectNodeContents(body);
                    range.collapse(false);
                    const sel = window.getSelection();
                    sel.removeAllRanges();
                    sel.addRange(range);

                    // Create line break + image node
                    const br  = document.createElement('br');
                    const img = document.createElement('img');
                    img.src            = dataUri;
                    img.style.maxWidth = '100%';
                    img.style.display  = 'block';
                    img.style.marginTop = '8px';

                    // Append to end of body
                    body.appendChild(br);
                    body.appendChild(img);

                    // Tell Outlook the content changed
                    body.dispatchEvent(new InputEvent('input', { bubbles: true }));
                    body.dispatchEvent(new Event('change', { bubbles: true }));

                    return true;
                }
            """, data_uri)

            page.wait_for_timeout(1000)

            if success:
                print(f"   ✅ Image embedded inline: {filename}")
                self.logger.info(f"Inline image embedded via JS: {filename}")
            else:
                raise RuntimeError("JS injection returned false — body element not found")

        except Exception as e:
            self.logger.error(f"JS image embed failed for {filename}: {e}")
            print(f"   ❌ Image embed failed: {e} — sending without image")

    # ─────────────────────────────────────────────────────────────────────────
    # Helpers
    # ─────────────────────────────────────────────────────────────────────────

    def _recover_to_inbox(self, page: Page):
        try:
            page.get_by_role("button", name="New mail").wait_for(timeout=3000)
        except PWTimeout:
            self.logger.info("Recovering — navigating back to inbox")
            page.goto(f"{Config.OUTLOOK_URL}/mail/0/")
            page.get_by_role("button", name="New mail").wait_for(timeout=Config.PAGE_TIMEOUT)

    def _take_screenshot(self, page: Page, label: str):
        try:
            safe_label = label.replace("@", "_at_").replace(".", "_")
            path = os.path.join(Config.SCREENSHOT_DIR, f"{safe_label}.png")
            page.screenshot(path=path)
            self.logger.info(f"Screenshot saved: {path}")
        except Exception as e:
            self.logger.warning(f"Could not save screenshot: {e}")
            


                       