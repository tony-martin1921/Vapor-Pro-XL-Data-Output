import serial
import pandas as pd
import shutil
import os
import re
from openpyxl import load_workbook
import time

# ============================
# Configuration
# ============================
SERIAL_PORT = "COM3"
BAUD_RATE = 9600
XLSX_FOLDER = r"C:\Data\MoistureLogs"
NETWORK_SHARE = r"\\fileserver\quality\MoistureAnalyzerDataOutput"


# ============================
# Helper Functions
# ============================

def connect_serial():
    while True:
        try:
            ser = serial.Serial(
                port=SERIAL_PORT,
                baudrate=BAUD_RATE,
                bytesize=serial.EIGHTBITS,
                parity=serial.PARITY_NONE,
                stopbits=serial.STOPBITS_ONE,
                timeout=1,
                xonxoff=False,
                rtscts=False,
                dsrdtr=False
            )
            print(f"Connected to {SERIAL_PORT}")
            return ser
        except serial.SerialException as e:
            print(f"Serial connection failed: {e}. Retrying in 5s...")
            time.sleep(5)


def read_serial_line(ser):
    try:
        return ser.readline().decode(errors='ignore').strip()
    except serial.SerialException:
        print("Serial connection lost. Attempting to reconnect...")
        return None


def autofit_excel_columns(filepath):
    wb = load_workbook(filepath)
    ws = wb.active
    for col in ws.columns:
        max_length = 0
        col_letter = col[0].column_letter
        for cell in col:
            if cell.value:
                max_length = max(max_length, len(str(cell.value)))
        ws.column_dimensions[col_letter].width = max_length + 2
    wb.save(filepath)
    wb.close()


def save_record_xlsx(record):
    # Extract timestamp from data fields
    date_str = record.get("DATE", None)
    time_str = record.get("TIME OF DAY", None)

    if date_str and time_str:
        try:
            dt = pd.to_datetime(f"{date_str} {time_str}", dayfirst=True)
            timestamp_for_filename = dt.strftime("%d-%m-%Y-%H-%M-%S")
        except:
            timestamp_for_filename = time.strftime("%d-%m-%Y-%H-%M-%S")
    else:
        timestamp_for_filename = time.strftime("%d-%m-%Y-%H-%M-%S")

    # Extract fields for filename
    job_id = record.get("ID", "UNKNOWN")
    lot_number = record.get("LOT NUMBER", "UNKNOWN")
    sample_name = record.get("SAMPLE NAME", "UNKNOWN")

    # Filename
    filename = f"{job_id}-{lot_number}-{sample_name}-{timestamp_for_filename}.xlsx"
    filename = re.sub(r'[\\/*?:"<>|]', "_", filename)

    # Convert to vertical DataFrame
    rows = [{"Key": k, "Data": v} for k, v in record.items()]
    df = pd.DataFrame(rows)

    os.makedirs(XLSX_FOLDER, exist_ok=True)
    local_path = os.path.join(XLSX_FOLDER, filename)
    df.to_excel(local_path, index=False)
    autofit_excel_columns(local_path)

    try:
        dest_path = os.path.join(NETWORK_SHARE, filename)
        shutil.copy(local_path, dest_path)
        print(f"Saved Excel to {local_path} and copied to network share as {filename}")
    except Exception as e:
        print(f"Error copying to network share: {e}")

# ============================
# Main Program
# ============================

if __name__ == "__main__":
    ser = connect_serial()
    current_record = {}
    collecting = False
    skipping = False

    print("Reading serial data... Press Ctrl+C to stop.")

    try:
        while True:
            line = read_serial_line(ser)
            if line is None:
                ser = connect_serial()
                continue
            if not line:
                continue

            # Detect start of test
            if "VAPOR PRO XL TEST RESULT REPORT" in line.upper():
                collecting = True
                skipping = False
                current_record = {}
                continue

            # Detect end of test
            if "USER NAME" in line.upper():
                if collecting and current_record:
                    save_record_xlsx(current_record)
                collecting = False
                skipping = True
                current_record = {}
                continue

            if skipping or not collecting:
                continue

            # Capture key:value pairs with normalized keys
            if ":" in line:
                key, value = line.split(":", 1)
                key = key.strip().upper()
                value = value.strip()
                current_record[key] = value
            else:
                current_record.setdefault("NOTES", "")
                current_record["NOTES"] += f" {line}"

    except KeyboardInterrupt:
        print("\nStopped by user.")
        if collecting and current_record:
            save_record_xlsx(current_record)

    ser.close()
