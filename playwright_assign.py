import argparse
import json
import random
import re
from datetime import datetime
from pathlib import Path

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment
from playwright.sync_api import TimeoutError as PlaywrightTimeoutError
from playwright.sync_api import sync_playwright


BASE_DIR = Path(__file__).resolve().parent
CONTACTS_FILE = BASE_DIR / "contacts.xlsx"
PROFILE_DIR = BASE_DIR / "whatsapp_profile"
SCREENSHOT_DIR = BASE_DIR / "screenshots"

SEARCH_SELECTORS = (
    '[aria-label="Search input textbox"]',
    '[aria-label="Search or start a new chat"]',
    'div[contenteditable="true"][data-tab="3"]',
)
COMPOSER_SELECTORS = (
    'footer [aria-label="Type a message"]',
    'footer div[contenteditable="true"][data-tab="10"]',
    'footer div[contenteditable="true"][role="textbox"]',
)


def human_pause(page):
    page.wait_for_timeout(random.randint(2000, 5000))


def first_visible(page, selectors):
    for selector in selectors:
        locator = page.locator(selector).first
        try:
            if locator.is_visible(timeout=1000):
                return locator
        except PlaywrightTimeoutError:
            continue
    return None


def wait_for_first_visible(page, selectors, timeout=3000):
    for selector in selectors:
        locator = page.locator(selector).first
        try:
            locator.wait_for(state="visible", timeout=timeout)
            return locator
        except PlaywrightTimeoutError:
            continue
    return None


def wait_for_login(page):
    print("WhatsApp Web is open. Scan the QR code if this is the first run.")
    deadline = datetime.now().timestamp() + 180
    while datetime.now().timestamp() < deadline:
        for selector in SEARCH_SELECTORS:
            try:
                page.wait_for_selector(selector, state="visible", timeout=1000)
                return
            except PlaywrightTimeoutError:
                continue
        page.wait_for_timeout(500)
    raise TimeoutError("WhatsApp login was not detected within 3 minutes.")


def handle_post_login_popup(page):
    dialog_selector = '[role="dialog"][aria-modal="true"]'
    continue_pattern = re.compile(r"^\s*continue\s*$", re.IGNORECASE)

    for _ in range(3):
        dialog = page.locator(dialog_selector).last
        try:
            dialog.wait_for(state="visible", timeout=1500)
        except PlaywrightTimeoutError:
            return

        continue_button = dialog.get_by_role("button", name=continue_pattern).last
        try:
            continue_button.wait_for(state="visible", timeout=1000)
        except PlaywrightTimeoutError:
            popup_text = dialog.inner_text().strip().replace("\n", " ")
            print(
                f"WhatsApp popup needs a manual click: {popup_text[:250]}",
                flush=True,
            )
            try:
                dialog.wait_for(state="hidden", timeout=120000)
            except PlaywrightTimeoutError as error:
                raise RuntimeError(
                    "WhatsApp popup was not dismissed within 2 minutes."
                ) from error
            continue

        continue_button.click()
        print("Clicked Continue in WhatsApp popup.", flush=True)
        try:
            dialog.wait_for(state="hidden", timeout=5000)
        except PlaywrightTimeoutError:
            continue

    remaining_dialog = page.locator(dialog_selector).last
    try:
        remaining_dialog.wait_for(state="visible", timeout=1000)
    except PlaywrightTimeoutError:
        return
    raise RuntimeError("WhatsApp popup is still open after Continue was clicked.")


def search_contact(page, phone):
    phone_digits = re.sub(r"\D", "", phone)
    number_pattern = re.compile(r"\D*".join(re.escape(digit) for digit in phone_digits))
    result_selectors = (
        '#pane-side [data-testid="cell-frame-container"]',
        '#pane-side [role="listitem"]',
        '#pane-side [role="row"]',
    )

    for search_term in dict.fromkeys((phone, phone_digits)):
        if not search_term:
            continue
        handle_post_login_popup(page)
        search_box = wait_for_first_visible(page, SEARCH_SELECTORS)
        if search_box is None:
            raise RuntimeError("WhatsApp search box did not appear after login and popup handling.")

        search_box.click(timeout=5000)
        search_box.fill(search_term)
        human_pause(page)

        for result_selector in result_selectors:
            matching_result = page.locator(result_selector).filter(
                has_text=number_pattern
            ).first
            try:
                matching_result.wait_for(state="visible", timeout=2000)
            except PlaywrightTimeoutError:
                continue

            try:
                matching_result.click(timeout=5000)
            except PlaywrightTimeoutError:
                continue

            human_pause(page)
            if wait_for_first_visible(page, COMPOSER_SELECTORS):
                return True

        matching_text = page.locator("#pane-side").get_by_text(number_pattern).last
        try:
            matching_text.wait_for(state="visible", timeout=2000)
            matching_text.click(timeout=5000)
            human_pause(page)
            if wait_for_first_visible(page, COMPOSER_SELECTORS):
                return True
        except PlaywrightTimeoutError:
            continue
    return False


def extract_last_incoming_messages(page):
    page.wait_for_timeout(1000)
    selectors = (
        'div.message-in span.selectable-text',
        'div.message-in [data-testid="selectable-text"]',
        'div.message-in div.copyable-text',
    )
    for selector in selectors:
        messages = page.locator(selector)
        count = messages.count()
        if count:
            extracted = []
            for index in range(max(0, count - 3), count):
                text = messages.nth(index).inner_text().strip()
                if text:
                    extracted.append(text)
            return extracted

    bubbles = page.locator("#main div.message-in")
    count = bubbles.count()
    if count:
        extracted = []
        for index in range(max(0, count - 3), count):
            text = bubbles.nth(index).inner_text().strip()
            if text:
                extracted.append(text)
        return extracted
    return []


def read_contacts():
    if not CONTACTS_FILE.exists():
        raise FileNotFoundError(
            f"{CONTACTS_FILE.name} must be in the same folder as this script."
        )

    workbook = load_workbook(CONTACTS_FILE, read_only=True, data_only=True)
    try:
        rows = workbook.active.iter_rows(values_only=True)
        headers = next(rows, None)
        if not headers:
            raise ValueError("contacts.xlsx has no header row.")

        columns = {
            str(value).strip().lower(): index
            for index, value in enumerate(headers)
            if value is not None
        }
        if "name" not in columns or "phone" not in columns:
            raise ValueError("contacts.xlsx must contain Name and Phone columns.")

        contacts = []
        for row in rows:
            name_value = row[columns["name"]] if columns["name"] < len(row) else None
            phone_value = row[columns["phone"]] if columns["phone"] < len(row) else None
            message_value = (
                row[columns["message"]]
                if "message" in columns and columns["message"] < len(row)
                else None
            )
            name = str(name_value or "").strip()
            phone = str(phone_value or "").strip()
            if name or phone:
                contacts.append(
                    {
                        "name": name,
                        "phone": phone,
                        "message_template": str(message_value or "").strip(),
                    }
                )
        return contacts
    finally:
        workbook.close()


def new_contact_result(contact):
    name = contact["name"]
    template = contact["message_template"] or "Hello {name}, I hope you're doing well."
    return {
        "name": name,
        "phone": contact["phone"],
        "message": template.replace("{name}", name),
        "status": "not_sent",
        "confirmation_method": None,
        "sent_at": None,
        "screenshot": None,
        "last_messages": [],
        "extraction_status": "not_started",
        "error": None,
    }


def send_to_contact(page, contact, index):
    result = new_contact_result(contact)
    try:
        if not search_contact(page, result["phone"]):
            result["error"] = "Contact was not found in WhatsApp search."
            return result

        composer = first_visible(page, COMPOSER_SELECTORS)
        if composer is None:
            result["error"] = "Message composer was not found."
            return result

        human_pause(page)
        composer.fill(result["message"])
        outgoing = page.locator("div.message-out").filter(has_text=result["message"])
        previous_count = outgoing.count()
        composer.press("Enter")

        send_confirmed = False
        try:
            outgoing.nth(previous_count).wait_for(state="visible", timeout=10000)
            send_confirmed = True
            result["confirmation_method"] = "outgoing_message_visible"
        except PlaywrightTimeoutError:
            if not composer.inner_text().strip():
                send_confirmed = True
                result["confirmation_method"] = "composer_cleared"
            else:
                result["error"] = "Could not confirm the message was sent in the chat."

        if send_confirmed:
            result["status"] = "sent"
            result["sent_at"] = datetime.now().isoformat(timespec="seconds")
            page.wait_for_timeout(1000)
            SCREENSHOT_DIR.mkdir(parents=True, exist_ok=True)
            safe_name = re.sub(
                r"[^A-Za-z0-9_-]+", "_", result["name"] or result["phone"]
            ).strip("_")
            screenshot_path = SCREENSHOT_DIR / f"{index:03d}_{safe_name}.png"
            page.screenshot(path=str(screenshot_path))
            result["screenshot"] = str(screenshot_path.relative_to(BASE_DIR))

        result["last_messages"] = extract_last_incoming_messages(page)
        result["extraction_status"] = (
            "extracted" if result["last_messages"] else "no_incoming_messages"
        )
    except Exception as error:
        result["error"] = f"{type(error).__name__}: {error}"
    return result


def save_reports(results):
    report_date = datetime.now().strftime("%Y-%m-%d")
    json_path = BASE_DIR / f"whatsapp_report_{report_date}.json"
    excel_path = BASE_DIR / f"whatsapp_report_{report_date}.xlsx"
    report = {
        "generated_at": datetime.now().isoformat(timespec="seconds"),
        "total_contacts": len(results),
        "sent": sum(item["status"] == "sent" for item in results),
        "failed": sum(item["status"] != "sent" for item in results),
        "contacts": results,
    }
    json_path.write_text(
        json.dumps(report, indent=2, ensure_ascii=False), encoding="utf-8"
    )

    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "Summary"
    sheet.append(
        [
            "Name",
            "Phone",
            "Status",
            "Confirmation Method",
            "Sent Message",
            "Sent At",
            "Screenshot",
            "Last 3 Messages",
            "Extraction Status",
            "Error",
        ]
    )
    for item in results:
        sheet.append(
            [
                item["name"],
                item["phone"],
                item["status"],
                item["confirmation_method"],
                item["message"],
                item["sent_at"],
                item["screenshot"],
                "\n".join(item["last_messages"]),
                item["extraction_status"],
                item["error"],
            ]
        )
    sheet.freeze_panes = "A2"
    sheet.auto_filter.ref = sheet.dimensions
    column_widths = (22, 20, 14, 26, 32, 22, 38, 46, 22, 55)
    for index, width in enumerate(column_widths, start=1):
        sheet.column_dimensions[sheet.cell(row=1, column=index).column_letter].width = width
    for row in sheet.iter_rows():
        for cell in row:
            cell.alignment = Alignment(vertical="top", wrap_text=True)
    workbook.save(excel_path)
    print(f"JSON report:  {json_path}")
    print(f"Excel report: {excel_path}")


def main():
    parser = argparse.ArgumentParser(description="WhatsApp Web message automation")
    parser.add_argument(
        "--dry-run",
        action="store_true",
        help="Open WhatsApp Web and wait without sending messages or writing reports.",
    )
    args = parser.parse_args()
    print(f"Running script: {Path(__file__).resolve()}", flush=True)

    contacts = read_contacts()
    if not contacts:
        raise ValueError("No contacts found in contacts.xlsx.")

    results = []
    try:
        with sync_playwright() as playwright:
            print("Launching Chromium...", flush=True)
            context = playwright.chromium.launch_persistent_context(
                user_data_dir=str(PROFILE_DIR), headless=False
            )
            try:
                print("Chromium launched; opening WhatsApp Web...", flush=True)
                page = context.pages[0] if context.pages else context.new_page()
                page.goto("https://web.whatsapp.com", wait_until="domcontentloaded")
                print("WhatsApp Web loaded; waiting for login...", flush=True)
                wait_for_login(page)
                handle_post_login_popup(page)

                if args.dry_run:
                    print("DRY RUN: no messages will be sent and reports will not be changed.", flush=True)
                else:
                    for index, contact in enumerate(contacts, start=1):
                        label = contact["name"] or contact["phone"]
                        print(f"[{index}/{len(contacts)}] Processing {label}...", flush=True)
                        results.append(send_to_contact(page, contact, index))
                        print(f"  Status: {results[-1]['status']}", flush=True)
                        if results[-1]["error"]:
                            print(f"  Error: {results[-1]['error']}", flush=True)
                        if index < len(contacts):
                            human_pause(page)

                    if args.dry_run:
                        print("Dry run finished. Close the browser window when you are done.", flush=True)
                        while not page.is_closed():
                            page.wait_for_timeout(500)
                    else:
                        print("Run finished. Closing WhatsApp Web.", flush=True)
            finally:
                context.close()
    except Exception as error:
        print(f"Browser run stopped: {type(error).__name__}: {error}", flush=True)
        if not args.dry_run:
            for contact in contacts[len(results):]:
                failed_result = new_contact_result(contact)
                failed_result["error"] = (
                    f"Run stopped before contact was processed: {type(error).__name__}: {error}"
                )
                results.append(failed_result)
    finally:
        if not args.dry_run:
            save_reports(results)


if __name__ == "__main__":
    try:
        main()
    except Exception as error:
        print(f"Run could not be completed: {type(error).__name__}: {error}")