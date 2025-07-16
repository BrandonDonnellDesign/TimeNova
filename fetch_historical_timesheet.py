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
# TIMESHEET_SELECTOR is no longer needed as screenshot functionality is removed
# API_PREFIX is no longer needed as we are directly calling the URL

# Define the new URL to fetch timesheet details
NEW_TIMESHEET_API_URL = "https://online7.timeanywhere.com/novatimeservicesV2/api/16c135d5-ca06-4522-b863-569e1c67c565/timesheetdetail?AccessSeq=1142&EmployeeSeq=26462&StartDate=Tue%20Mar%2001%202022&EndDate=Fri%20Jul%2004%202025&UserSeq=0&CustomDateRange=true&ShowOneMoreDay=false&EmployeeSeqList=&DailyDate=Sun%20Mar%2013%202022&ForceAbsent=false&PolicyGroup="

def sanitize_folder_name(name):
    """Sanitizes a string to be a valid folder name."""
    # Replace invalid characters with an underscore
    name = re.sub(r'[\\\\/:*?"<>|]', '_', name)
    # Replace spaces with underscores
    name = name.replace(' ', '_')
    return name

def login_and_grab_timesheet():
    """
    Logs into the Novatime system, then directly fetches timesheet data
    from a specified API URL, saves it as JSON and CSV in the current directory.
    Screenshot functionality has been removed.
    """
    with sync_playwright() as p:
        # Launch a Chromium browser instance in headless mode
        browser = p.chromium.launch(headless=True)
        context = browser.new_context()
        # Set a large viewport size for consistency
        page = context.new_page()
        page.set_viewport_size({"width": 2560, "height": 1440})

        # Files will be saved directly in the current working directory
        print(f"Files will be saved in the current working directory: {os.getcwd()}")

        # 1) Navigate to login page and perform login
        print(f"Navigating to login page: {LOGIN_URL}")
        try:
            page.goto(LOGIN_URL, wait_until="networkidle")
            page.fill("#txtUserName", NOVATIME_USERNAME)
            page.fill("#txtPassword", NOVATIME_PASSWORD)
            page.click("input[value='Employee Web']")
            # Wait for the page to load after login, allowing redirects
            page.wait_for_load_state("networkidle")
            print("✅ Successfully logged in.")
        except Exception as e:
            print(f"❌ Failed to log in: {e}")
            browser.close()
            return

        # --- New Step: Directly fetch JSON data from the specified API URL ---
        captured_json_data = None
        try:
            print(f"Attempting to fetch timesheet data from: {NEW_TIMESHEET_API_URL}")
            # Navigate directly to the API URL. Playwright will fetch its content.
            # Increased timeout for page.goto to 60 seconds (60000 ms)
            response = page.goto(NEW_TIMESHEET_API_URL, wait_until="domcontentloaded", timeout=60000)

            if response and response.ok:
                # Get the response body as text and parse it as JSON
                json_data_str = response.text()
                captured_json_data = json.loads(json_data_str)
                print("✅ Successfully fetched JSON data from the provided URL.")

                # --- Debugging additions ---
                if isinstance(captured_json_data, dict):
                    print(f"Keys in captured JSON data: {list(captured_json_data.keys())}")
                    records = captured_json_data.get('DataList', [])
                    print(f"Number of 'DataList' records found: {len(records)}")
                else:
                    print("Captured JSON data is not a dictionary. Cannot inspect keys.")
                # --- End debugging additions ---

            else:
                print(f"❌ Failed to fetch data from {NEW_TIMESHEET_API_URL}. "
                      f"Status: {response.status if response else 'No response'}")
                browser.close()
                return # Exit if data fetching failed

        except json.JSONDecodeError as e:
            print(f"❌ Failed to decode JSON from the URL response: {e}")
            # Print only a snippet of the response text to avoid excessively long output
            print(f"Response text that caused error (first 500 chars): {json_data_str[:500] if 'json_data_str' in locals() else 'Not available'}")
            browser.close()
            return
        except Exception as e:
            print(f"❌ An error occurred while fetching data from the URL: {e}")
            browser.close()
            return

        if captured_json_data:
            try:
                # Files will be saved directly in the current directory
                json_path = os.path.join(os.getcwd(), "historical_timesheet.json")
                csv_path = os.path.join(os.getcwd(), "historical_timesheet.csv")

                # 4) Save the captured JSON data to a file
                with open(json_path, 'w') as f:
                    json.dump(captured_json_data, f, indent=4)
                print(f"✅ JSON data saved at {json_path}")

                # 5) Process JSON data and save to CSV with new format
                # Defined the exact column headers as per your request, including empty columns
                # Re-added \xa0 (non-breaking space) to match the other program's expectation
                column_headers = [
                    "Date", "Pay Code", "In", "", "Out", "", "Reg", "OT-1", "OT-2",
                    "Daily Hours\xa0*", "Shift Exp", "Schedule", "Total Hours\xa0*",
                    "Account", "ActShortCode", "Facility", "", "", "", ""
                ]

                # Mapping for direct fields from the JSON record
                column_field_map = {
                    "Pay Code": "cPayCodeDescription",
                    "Schedule": "cSchedule",
                    "Shift Exp": "cExpCode", # Using cExpCode for this field
                }

                records_to_process = captured_json_data.get('DataList', [])

                # Sort records by dWorkDate and then by dOut (or dIn if dOut is often None)
                # This helps identify the 'last' entry for a given day for Daily Hours calculation
                def sort_key(rec):
                    work_date_str = rec.get('dWorkDate', '').split(' ')[0]
                    out_time_str = rec.get('dOut', '')
                    try:
                        work_date_obj = datetime.strptime(work_date_str, '%m/%d/%Y')
                        # Use a placeholder time if dOut is None or invalid to ensure consistent sorting
                        out_time_obj = datetime.strptime(out_time_str, '%m/%d/%Y %H:%M:%S') if out_time_str else datetime.min
                        return (work_date_obj, out_time_obj)
                    except ValueError:
                        return (datetime.min, datetime.min) # Fallback for unparseable dates/times

                records_to_process.sort(key=sort_key)

                daily_hours_totals = {}
                last_record_on_date = {} # Stores the last record object for each date

                # First pass to calculate daily totals and identify last record per day
                for rec in records_to_process:
                    date_str = rec.get('dWorkDate')
                    total_hours_for_punch = rec.get('nTotalHours', 0.0)
                    if date_str:
                        date_key = date_str.split(' ')[0] # e.g., "03/16/2022"
                        daily_hours_totals[date_key] = daily_hours_totals.get(date_key, 0.0) + float(total_hours_for_punch)
                        last_record_on_date[date_key] = rec # Update with the current record (last one encountered so far for this date)

                with open(csv_path, 'w', newline='') as csvfile:
                    writer = csv.writer(csvfile)
                    writer.writerow(column_headers)

                    if not records_to_process:
                        print("⚠️ No 'DataList' records found in the JSON data to write to CSV. The CSV will only contain headers.")

                    for rec in records_to_process:
                        row_data = []
                        current_date_key = rec.get('dWorkDate', '').split(' ')[0]

                        for col in column_headers:
                            if col == "Date":
                                date_value = rec.get('dWorkDate')
                                if date_value is not None:
                                    try:
                                        # Parse "MM/DD/YYYY HH:MM:SS" and format to "DayOfWeek MM/DD/YYYY"
                                        date_obj = datetime.strptime(date_value.split(' ')[0], '%m/%d/%Y')
                                        row_data.append(date_obj.strftime('%a %m/%d/%Y'))
                                    except ValueError:
                                        row_data.append(str(date_value)) # Fallback if parsing fails
                                else:
                                    row_data.append("") # Append empty string if None
                            elif col == "In":
                                in_time_value = rec.get('dIn')
                                if in_time_value is not None:
                                    try:
                                        # Parse "MM/DD/YYYY HH:MM:SS" and format to "HH:MM AM/PM"
                                        time_obj = datetime.strptime(in_time_value, '%m/%d/%Y %H:%M:%S')
                                        row_data.append(time_obj.strftime('%I:%M %p'))
                                    except ValueError:
                                        row_data.append(str(in_time_value)) # Fallback if parsing fails
                                else:
                                    row_data.append("") # Append empty string if None
                            elif col == "": # Handle the empty columns after "In" and "Out" and at the end
                                row_data.append("")
                            elif col == "Out":
                                out_time_value = rec.get('dOut')
                                if out_time_value is not None:
                                    try:
                                        # Parse "MM/DD/YYYY HH:MM:SS" and format to "HH:MM AM/PM"
                                        time_obj = datetime.strptime(out_time_value, '%m/%d/%Y %H:%M:%S')
                                        row_data.append(time_obj.strftime('%I:%M %p'))
                                    except ValueError:
                                        row_data.append(str(out_time_value)) # Fallback if parsing fails
                                else:
                                    row_data.append("") # Append empty string if None
                            elif col == "Reg":
                                reg_pay = rec.get('nWorkHours', 0.0)
                                row_data.append(f"{float(reg_pay):.2f}") # Format to 2 decimal places
                            elif col == "OT-1":
                                ot1_pay = rec.get('nOT1Pay', 0.0)
                                row_data.append(f"{float(ot1_pay):.2f}") # Format to 2 decimal places
                            elif col == "OT-2":
                                ot2_pay = rec.get('nOT2Pay', 0.0)
                                row_data.append(f"{float(ot2_pay):.2f}") # Format to 2 decimal places
                            elif col == "Daily Hours\xa0*": # Re-added non-breaking space
                                # Populate Daily Hours only if this is the last record for the day
                                if rec == last_record_on_date.get(current_date_key):
                                    row_data.append(f"{daily_hours_totals.get(current_date_key, 0.0):.1f}") # Format to 1 decimal place
                                else:
                                    row_data.append("") # Leave empty for other entries on the same day
                            elif col == "Total Hours\xa0*": # Re-added non-breaking space
                                # This is the nDailyTotalHours for the individual punch, formatted to 1 decimal place
                                total_hours_punch = rec.get('nDailyTotalHours', 0.0)
                                row_data.append(f"{float(total_hours_punch):.1f}")
                            elif col == "Account":
                                account_value = ""
                                # Iterate through GroupValueList to find Account (iGroupNumber 12)
                                for group_val in rec.get('GroupValueList', []):
                                    if group_val.get('iGroupNumber') == 12:
                                        c_group_value = group_val.get('cGroupValue', '')
                                        c_group_value_description = group_val.get('cGroupValueDescription', '')
                                        account_value = f"{c_group_value} [{c_group_value_description}]"
                                        break
                                row_data.append(account_value)
                            elif col == "ActShortCode":
                                act_short_code_value = ""
                                # Iterate through GroupValueList to find Activity ShortCode (iGroupNumber 3)
                                for group_val in rec.get('GroupValueList', []):
                                    if group_val.get('iGroupNumber') == 3:
                                        c_group_value = group_val.get('cGroupValue', '')
                                        c_group_value_description = group_val.get('cGroupValueDescription', '')
                                        act_short_code_value = f"{c_group_value} [{c_group_value_description}]"
                                        break
                                row_data.append(act_short_code_value)
                            elif col == "Facility":
                                facility_value = ""
                                # Iterate through GroupValueList to find FacilityName (iGroupNumber 1)
                                for group_val in rec.get('GroupValueList', []):
                                    if group_val.get('iGroupNumber') == 1:
                                        c_group_value = group_val.get('cGroupValue', '')
                                        c_group_value_description = group_val.get('cGroupValueDescription', '')
                                        facility_value = f"{c_group_value} [{c_group_value_description}]"
                                        break
                                row_data.append(facility_value)
                            elif col == "Shift Exp":
                                value = rec.get(column_field_map.get(col, col), "")
                                row_data.append(str(value))
                            else:
                                # For other columns, use the direct field map
                                value = rec.get(column_field_map.get(col, col), "")
                                # Ensure numeric values are formatted to two decimal places if they are floats
                                if isinstance(value, (int, float)):
                                    row_data.append(f"{value:.2f}") # Default to 2 for non-hour numbers
                                else:
                                    row_data.append(str(value))

                        writer.writerow(row_data)
                print(f"✅ Timesheet CSV file saved at {csv_path}")

                print("✅ Script completed successfully with JSON and CSV saved.")

            except Exception as e:
                print(f"❌ Failed to process captured JSON data and save files: {e}")
        else:
            print("❌ No JSON data was captured from the API URL. No files saved.")

        browser.close()

if __name__ == "__main__":
    # Check if necessary environment variables are set
    if not NOVATIME_USERNAME or not NOVATIME_PASSWORD:
        print("❌ Missing NOVATIME_USERNAME or NOVATIME_PASSWORD in your .env file.")
    elif not LOGIN_URL:
        print("❌ Missing LOGIN_URL in your .env file.")
    else:
        login_and_grab_timesheet()
