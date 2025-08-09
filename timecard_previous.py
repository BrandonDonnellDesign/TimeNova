import os
import time
import csv
from datetime import datetime
from playwright.sync_api import sync_playwright
import json
import re
from dotenv import load_dotenv

# --- Load environment variables from .env ---
load_dotenv()
NOVATIME_USERNAME = os.getenv("NOVATIME_USERNAME")
NOVATIME_PASSWORD = os.getenv("NOVATIME_PASSWORD")
LOGIN_URL = os.getenv("LOGIN_URL")
TIMESHEET_SELECTOR = os.getenv("TIMESHEET_SELECTOR")
API_PREFIX = os.getenv("API_PREFIX")

def sanitize_folder_name(name):
    """Sanitizes a string to be a valid folder name."""
    name = re.sub(r'[\\/:*?"<>|]', '_', name)
    name = name.replace(' ', '_')
    return name

def login_and_grab_timesheet():
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=False)
        context = browser.new_context()
        page = context.new_page()
        page.set_viewport_size({"width": 2560, "height": 1440})

        # Create output directory
        script_dir = os.path.dirname(os.path.abspath(__file__))
        base_output_dir = os.path.join(script_dir, "timeCard")
        os.makedirs(base_output_dir, exist_ok=True)
        print(f"📁 Base output directory: {base_output_dir}")
        now_str = datetime.now().strftime("%Y%m%d")

        captured_json_data = {"data": None, "found": False}

        def handle_response(response):
            url = response.url
            if API_PREFIX in url:
                print(f"✅ Found matching JSON API request: {url}")
                try:
                    captured_json_data["data"] = response.body().decode("utf-8")
                    captured_json_data["found"] = True
                except Exception as e:
                    print(f"❌ Failed to capture JSON response body: {e}")

        page.on("response", handle_response)

        # 1) Go to login page
        page.goto(LOGIN_URL)
        page.wait_for_selector("#txtUserName")

        # 2) Login
        page.fill("#txtUserName", NOVATIME_USERNAME)
        page.fill("#txtPassword", NOVATIME_PASSWORD)
        page.click("input[value='Employee Web']")
        page.wait_for_load_state("networkidle")
        print("✅ Logged in!")

        # 3) Go to Timesheet
        page.wait_for_selector("h4:has-text('Timesheet')")
        page.click("h4:has-text('Timesheet')")
        page.wait_for_load_state("networkidle")
        # Extra wait to ensure Angular bindings are ready
        page.wait_for_timeout(3000) 
        print(f"📄 Page URL after Timesheet click: {page.url}")

        # 4) Select the dropdown and choose the option labeled "Last Pay Period"
        try:
            page.wait_for_selector("#MainContentFrame", state="attached", timeout=10000)
            print("✅ MainContentFrame loaded")
            dropdown_locator = (page.locator("#MainContentFrame")
                              .content_frame
                              .locator("#iFrame")
                              .content_frame
                              .locator("#titleRightDiv")
                              .get_by_role("combobox"))

            # Wait for dropdown to be visible and have options
            dropdown_locator.wait_for(state="visible", timeout=10000)
            page.wait_for_timeout(5000)  # Additional wait for Angular to populate options

            # Get all available options
            options = dropdown_locator.locator("option").all()
            option_values = [option.get_attribute("value") for option in options if option.get_attribute("value")]
            option_texts = [option.inner_text() for option in options]
            print(f"📋 Available dropdown options: {list(zip(option_texts, option_values))}")

            # Find the option with label "Last Pay Period"
            desired_label = "Last Pay Period"
            selected_value = None
            for option in options:
                if option.inner_text().strip() == desired_label:
                    selected_value = option.get_attribute("value")
                    break

            if selected_value:
                dropdown_locator.select_option(selected_value)
                print(f"✅ Selected option: {desired_label} (value: {selected_value})")
            else:
                print(f"⚠️ Option '{desired_label}' not found in dropdown.")
                if option_values:
                    fallback_option = option_values[0]
                    dropdown_locator.select_option(fallback_option)
                    print(f"✅ Selected fallback option: {option_texts[0]} (value: {fallback_option})")
                else:
                    print("❌ No options available in dropdown.")
                    browser.close()
                    return

            page.wait_for_timeout(5000)  # Wait for page to react to selection

        except Exception as e:
            print(f"❌ Failed to interact with dropdown: {e}")
            browser.close()
            return

        # 5) Find the iframe containing the timesheet
        max_wait_time = 120  # seconds
        start_time = time.time()
        timesheet_frame = None

        print("🔎 Searching for timesheet iframe...")

        while time.time() - start_time < max_wait_time:
            for frame in page.frames:
                try:
                    # Check if "TimesheetSection" is in the frame's URL or content
                    if "TimesheetSection" in frame.url or "TimesheetSection" in frame.content():
                        print(f"✅ Found timesheet in frame: {frame.url}")
                        timesheet_frame = frame
                        break
                except Exception as e:
                    pass  # Silently ignore frames that can't be accessed
            if timesheet_frame:
                break
            time.sleep(1)

        if not timesheet_frame:
            print("❌ Could not find timesheet iframe after waiting.")
            browser.close()
            return

        # 6) Navigate directly to the iframe URL
        iframe_url = timesheet_frame.url
        print(f"🌐 Navigating directly to timesheet iframe URL: {iframe_url}")
        page.goto(iframe_url)
        page.wait_for_load_state("networkidle")

        # Wait additional time for data and network requests
        print("⏳ Waiting 10 seconds for data and API request to finalize...")
        time.sleep(10)  # Increased wait time to ensure all data is loaded and captured

        # --- Process and save JSON/CSV data only once here ---
        if captured_json_data["found"] and captured_json_data["data"]:
            try:
                json_data = captured_json_data["data"]

                # Parse JSON to extract WeekGroupString for folder naming
                timesheet_json = json.loads(json_data)
                records = timesheet_json.get("DataList", [])

                # --- Determine date range for folder name ---
                pay_period_start = pay_period_end = None
                for rec in records:
                    if rec.get("dPayPeriodStart") and rec.get("dPayPeriodEnd"):
                        pay_period_start = rec["dPayPeriodStart"]
                        pay_period_end = rec["dPayPeriodEnd"]
                        break
                if not (pay_period_start and pay_period_end):
                    work_dates = [rec.get("dWorkDate") for rec in records if rec.get("dWorkDate")]
                    if work_dates:
                        pay_period_start = min(work_dates)
                        pay_period_end = max(work_dates)
                # If no pay period or work dates found, skip processing
                if not pay_period_start or not pay_period_end:
                    print("No timecard data available yet")
                    browser.close()
                    return
                def fmt(dtstr):
                    if not dtstr:
                        return "unknown"
                    try:
                        dt = datetime.strptime(dtstr.split()[0], "%m/%d/%Y")
                        return dt.strftime("%m-%d-%y")
                    except Exception:
                        return "unknown"
                start_str = fmt(pay_period_start)
                end_str = fmt(pay_period_end)
                folder_name = f"{start_str}_to_{end_str}"
                folder_name = sanitize_folder_name(folder_name)
                # --- END date range for folder name ---

                # Create the new weekly output directory
                weekly_output_dir = os.path.join(base_output_dir, folder_name)
                os.makedirs(weekly_output_dir, exist_ok=True)
                print(f"📁 Output files will be saved in: {weekly_output_dir}")

                # Define file paths using the new weekly_output_dir
                json_filename = "timesheet.json"
                json_path = os.path.join(weekly_output_dir, json_filename)
                
                csv_filename = "timesheet.csv"
                csv_path = os.path.join(weekly_output_dir, csv_filename)

                screenshot_filename = "timesheet.png"
                screenshot_path = os.path.join(weekly_output_dir, screenshot_filename)

                # Save JSON
                with open(json_path, "w", encoding="utf-8") as f:
                    f.write(json_data)
                print(f"✅ Timesheet JSON data saved at {json_path}")

                columns = [
                    "Date",
                    "Pay Code",
                    "In",
                    "Out",
                    "Reg",
                    "OT-1",
                    "OT-2",
                    "Daily Hours *",
                    "Shift Exp",
                    "Schedule",
                    "Total Hours\xa0*",
                    "Account",
                    "ActShortCode",
                    "Facility",
                ]

                column_field_map = {
                    "Date": "DateKey",
                    "Pay Code": "cPayCodeDescription",
                    "In": "dIn",
                    "Out": "dOut",
                    "Reg": "nWorkHours",
                    "OT-1": "nOT1Hours",
                    "OT-2": "nOT2Hours",
                    "Daily Hours *": "nDailyHours",
                    "Shift Exp": "cShiftExpression",
                    "Schedule": "cSchedule",
                    "Total Hours\xa0*": "nWeeklyHours",
                }

                # Save CSV
                with open(csv_path, "w", newline="", encoding="utf-8") as csvfile:
                    writer = csv.writer(csvfile)
                    writer.writerow(columns)
                    for rec in records:
                        row_data = []
                        account_value = ""
                        act_short_code_value = ""
                        facility_value = ""

                        for group in rec.get("GroupingList", []) + rec.get("GroupValueList", []):
                            if group.get("iGroupNumber") == 3:
                                account_value = group.get("cGroupValueDescription", "")
                                act_short_code_value = group.get("cGroupValue", "")
                            elif group.get("iGroupNumber") == 17:
                                facility_value = group.get("cGroupValueDescription", "")
                            elif group.get("iGroupNumber") == 16 and not facility_value:
                                facility_value = group.get("cGroupValueDescription", "")

                        for col in columns:
                            if col == "Account":
                                row_data.append(account_value)
                            elif col == "ActShortCode":
                                row_data.append(act_short_code_value)
                            elif col == "Facility":
                                row_data.append(facility_value)
                            else:
                                row_data.append(rec.get(column_field_map.get(col, col), ""))

                        writer.writerow(row_data)
                print(f"✅ Timesheet CSV file saved at {csv_path}")

                # 7) Locate the timesheet table element
                timesheet_element = page.query_selector(TIMESHEET_SELECTOR)
                if not timesheet_element:
                    print("❌ Could not find timesheet element on iframe page for screenshot.")
                else:
                    # 8) Save screenshot
                    timesheet_element.screenshot(path=screenshot_path)
                    print(f"📸 Screenshot saved at {screenshot_path}")

                print("✅ Script completed successfully with JSON and CSV saved.")

            except Exception as e:
                print(f"❌ Failed to process captured JSON data and save files: {e}")
        else:
            print("❌ Did not detect any JSON API requests matching the prefix. No files saved.")

        browser.close()

if __name__ == "__main__":
    if not NOVATIME_USERNAME or not NOVATIME_PASSWORD:
        print("❌ Missing NOVATIME_USERNAME or NOVATIME_PASSWORD in your .env file.")
    else:
        login_and_grab_timesheet()