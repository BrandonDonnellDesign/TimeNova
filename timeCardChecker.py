import os
import time
import pandas as pd
from datetime import datetime
from watchdog.observers import Observer
from watchdog.events import FileSystemEventHandler
from dotenv import load_dotenv

# --- Load environment variables ---
load_dotenv()

# --- Make watch folder relative to this script ---
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
WATCH_FOLDER = os.path.join(SCRIPT_DIR, "timeCard")

# Log file inside timeCard folder
LOG_FILE = os.path.join(WATCH_FOLDER, "discrepancy_log.csv")

# Output options
ENABLE_CSV = True  # Always save CSV
ENABLE_EMAIL = os.getenv("EMAIL_SMTP_SERVER") is not None
ENABLE_SLACK = os.getenv("SLACK_WEBHOOK_URL") is not None
ENABLE_DISCORD = os.getenv("DISCORD_WEBHOOK_URL") is not None
ENABLE_INITIAL_SCAN = os.getenv("INITIAL_SCAN", "false").lower() == "true"


def check_discrepancies(df):
    """
    Very simple example:
    Compare 'Daily Hours *' vs 'Total Hours *' and flag mismatches.
    """
    discrepancies = []

    for idx, row in df.iterrows():
        daily = str(row.get("Daily Hours *", "")).strip()
        total = str(row.get("Total Hours *", row.get("Total Hours *", ""))).strip()

        if daily and total and daily != total:
            discrepancies.append({
                "Date": row.get("Date", ""),
                "In": row.get("In", ""),
                "Out": row.get("Out", ""),
                "Daily Hours": daily,
                "Total Hours": total,
                "File": row.get("SourceFile", "")
            })

    return pd.DataFrame(discrepancies)


def log_discrepancies(discrepancies, source_file):
    """Append discrepancies to central log file inside timeCard."""
    if discrepancies.empty:
        print(f"[OK] No discrepancies in {source_file}")
        return

    discrepancies["ProcessedAt"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    discrepancies["SourceFile"] = os.path.basename(source_file)

    if not os.path.exists(LOG_FILE):
        discrepancies.to_csv(LOG_FILE, index=False)
    else:
        discrepancies.to_csv(LOG_FILE, mode="a", index=False, header=False)

    print(f"[!] Logged {len(discrepancies)} discrepancies from {source_file}")


class TimeCardHandler(FileSystemEventHandler):
    def on_created(self, event):
        if not event.is_directory and event.src_path.endswith(".csv"):
            self.process(event.src_path)

    def on_modified(self, event):
        if not event.is_directory and event.src_path.endswith(".csv"):
            self.process(event.src_path)

    def process(self, file_path):
        try:
            print(f"[Processing] {file_path}")
            df = pd.read_csv(file_path)

            # Keep track of which file rows came from
            df["SourceFile"] = os.path.basename(file_path)

            # Run discrepancy check
            discrepancies = check_discrepancies(df)

            # Log results
            if ENABLE_CSV:
                log_discrepancies(discrepancies, file_path)

            # Notifications only if .env values exist
            if ENABLE_EMAIL:
                print("📧 Email notification would be sent here.")
            if ENABLE_SLACK:
                print("💬 Slack notification would be sent here.")
            if ENABLE_DISCORD:
                print("🎮 Discord notification would be sent here.")

        except Exception as e:
            print(f"[Error] Failed to process {file_path}: {e}")


# --- Initial scan of existing files ---
def initial_scan(handler, folder):
    print("[Startup] Scanning existing CSV files...")
    for root, _, files in os.walk(folder):
        for f in files:
            if f.endswith(".csv"):
                file_path = os.path.join(root, f)
                handler.process(file_path)
    print("[Startup] Scan complete.")


if __name__ == "__main__":
    print(f"Monitoring folder: {WATCH_FOLDER}")
    event_handler = TimeCardHandler()

    # 🔹 Run initial scan only if enabled in .env
    if ENABLE_INITIAL_SCAN:
        initial_scan(event_handler, WATCH_FOLDER)

    observer = Observer()
    observer.schedule(event_handler, WATCH_FOLDER, recursive=True)
    observer.start()

    try:
        while True:
            time.sleep(5)
    except KeyboardInterrupt:
        observer.stop()
    observer.join()