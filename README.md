# NovaTime Timecard Exporter

This script automates logging into NovaTime, downloading your timesheet data, and saving it as JSON, CSV, and a screenshot.

## Setup

1. **Clone the repository** and install dependencies:
    ```sh
    pip install -r requirements.txt
    ```

2. **Create your `.env` file** (see `.env.example` for required variables):
    ```
    NOVATIME_USERNAME=your_username
    NOVATIME_PASSWORD=your_password
    LOGIN_URL=https://online7.timeanywhere.com/novatime/ewsfunctionkey.aspx?CID=your_company_id
    TIMESHEET_SELECTOR=#TimesheetSection > div.row.visible-lg-block.visible-md-block.visible-sm-block.hidden-xs.table-al-change
    API_PREFIX=https://online7.timeanywhere.com/novatimeservicesV2/api/your_company_id/timesheetdetail
    ```

3. **Run the script:**
    ```sh
    python timecard.py
    ```

## Output

- Timesheet data is saved in a folder named like `07-13-25_to_07-19-25` under `timeCard/`.
- Each folder contains:
    - `timesheet.json`
    - `timesheet.csv`
    - `timesheet.png` (screenshot)

If no timecard data is available, the script will output:  
`No timecard data available yet`

## Notes

- Requires [Playwright](https://playwright.dev/python/) and [python-dotenv](https://pypi.org/project/python-dotenv/).
- Use the provided `.env.example` as a template for your `.env` file.
