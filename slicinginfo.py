import subprocess
import pandas as pd
import re
import os
import time
from watchdog.observers import Observer
from watchdog.events import FileSystemEventHandler

# Define paths
WATCH_FOLDER = "/path/to/your/folder"  # Change this to your actual folder path
EXCEL_FILE = "slicing_info.xlsx"
ORCA_CLI = "/Applications/OrcaSlicer.app/Contents/MacOS/OrcaSlicer"  # Path to OrcaSlicer CLI

def extract_slicing_info(three_mf_file):
    """Runs OrcaSlicer CLI to extract print details from a 3MF file"""
    cmd = [ORCA_CLI, three_mf_file, "--info"]

    try:
        output = subprocess.check_output(cmd, universal_newlines=True)
    except subprocess.CalledProcessError as e:
        print("❌ Error running OrcaSlicer CLI:", e)
        return None

    # Extract details using regex
    filament_color = re.findall(r"Filament Color:\s*(.*)", output)
    filament_amount = re.findall(r"Filament Used:\s*([\d.]+) mm", output)
    print_time = re.findall(r"Estimated Print Time:\s*(.*)", output)

    return {
        "File Name": os.path.basename(three_mf_file),
        "Filament Color": filament_color[0] if filament_color else "N/A",
        "Filament Amount (mm)": filament_amount[0] if filament_amount else "N/A",
        "Estimated Print Time": print_time[0] if print_time else "N/A",
    }

def update_excel(data, excel_file=EXCEL_FILE):
    """Updates or creates an Excel file with new data"""
    try:
        df = pd.read_excel(excel_file) if os.path.exists(excel_file) else pd.DataFrame()
    except Exception:
        df = pd.DataFrame()

    df = pd.concat([df, pd.DataFrame([data])], ignore_index=True)
    df.to_excel(excel_file, index=False)
    print(f"✅ Data added to {excel_file}: {data}")

class Watcher(FileSystemEventHandler):
    """Monitors folder for new .3mf files"""
    def on_created(self, event):
        if event.is_directory or not event.src_path.endswith(".3mf"):
            return

        time.sleep(1)  # Wait for file to fully save
        print(f"📂 New file detected: {event.src_path}")

        slicing_data = extract_slicing_info(event.src_path)
        if slicing_data:
            update_excel(slicing_data)

def start_watching():
    """Starts folder monitoring"""
    print(f"👀 Watching folder: {WATCH_FOLDER}")
    observer = Observer()
    event_handler = Watcher()
    observer.schedule(event_handler, WATCH_FOLDER, recursive=False)
    observer.start()

    try:
        while True:
            time.sleep(1)  # Keeps script running
    except KeyboardInterrupt:
        observer.stop()
        print("\n🛑 Stopping folder watcher...")

    observer.join()

# Start watching folder
start_watching()
