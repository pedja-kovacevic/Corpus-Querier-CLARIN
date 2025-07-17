import requests
import urllib.parse
import pandas as pd
import time
import os
import sys
import threading
from datetime import datetime
import winsound  # Windows-specific; for sound ping

# === SETTINGS ===
QUERY_DELAY_SECONDS = 5
TIMEOUT_SECONDS = 300  # 5 minutes

# === Query Function ===
def get_hits(cql_query, corpus="ENTER CORPUS NAME"):
    base_url = "https://www.clarin.si/ske/bonito/run.cgi/first"
    params = {
        "corpname": corpus,
        "queryselector": "cqlrow",
        "default_attr": "word",
        "cql": cql_query,
        "pagesize": 1,
        "format": "json"
    }

    full_url = base_url + "?" + urllib.parse.urlencode(params)
    print(f"\n🔗 Querying:\n{full_url}")

    try:
        response = requests.get(full_url, timeout=TIMEOUT_SECONDS)
        response.raise_for_status()
        data = response.json()
        hits = data.get("fullsize", 0)
        print(f"✅ Total hits for {cql_query}: {hits}")
        return hits
    except Exception as e:
        print(f"❌ Error while querying '{cql_query}': {e}")
        raise  # re-raise the error so it can be caught outside

# === Save Function ===
def save_output(df, output_file):
    df.to_excel(output_file, index=False)
    print(f"💾 Saved output to {output_file}")

# === Main Loop ===
def process_file(input_path, output_path, start_row, end_row, columns):
    try:
        df = pd.read_excel(input_path)
    except Exception as e:
        print(f"❌ Could not load input file: {e}")
        return

    try:
        for row in range(start_row - 1, end_row):
            print(f"\n➡️ Processing row {row + 1}")
            for col in columns:
                try:
                    cql = str(df.at[row, col])
                    if not cql or not cql.startswith("["):
                        print(f"⏭️ Skipping invalid or empty CQL in row {row+1}, column {col}")
                        continue
                    hits = get_hits(cql)
                    df.at[row, col] = hits
                    time.sleep(QUERY_DELAY_SECONDS)
                except Exception as e:
                    print(f"⚠️ Error in row {row+1}, column {col}: {e}")
                    save_output(df, output_path)
                    print("⏸️ Pausing due to error. Exiting.")
                    return

    except KeyboardInterrupt:
        print("\n🛑 Interrupted by user (CTRL+C). Saving progress...")
        save_output(df, output_path)
        return
    except Exception as e:
        print(f"❌ Unexpected failure: {e}")
        save_output(df, output_path)
        return

    # ✅ If everything completes
    save_output(df, output_path)
    winsound.Beep(1000, 1000)  # beep: frequency=1000Hz, duration=1000ms
    print("✅ All done! ✨")

# === Entry Point ===
if __name__ == "__main__":
    print("📥 Spreadsheet Query Processor for CLARIN.si")

    try:
        input_file = input("📝 Enter full path to input Excel file: ").strip().strip('"')
        output_file = input("💾 Enter full path to save output Excel file: ").strip().strip('"')
        start_row = int(input("🔢 Start from row (1-based): ").strip())
        end_row = int(input("🔢 End at row (1-based): ").strip())
        columns_input = input("🔠 Enter comma-separated column names to scan (e.g., B,C,D): ").strip()
        columns = [col.strip().upper() for col in columns_input.split(",")]
    except Exception as e:
        print(f"❌ Invalid input: {e}")
        sys.exit(1)

    process_file(input_file, output_file, start_row, end_row, columns)
