import requests
import urllib.parse
import time
import pandas as pd
import threading
import signal
import sys

# Timeout limit (in seconds) per request
TIMEOUT_LIMIT = 300

def get_hits(cql, corpus="srwac"):
    base_url = "https://www.clarin.si/ske/bonito/run.cgi/view"
    results = []

    def fetch_in_thread(attempt_num, result_container):
        try:
            params = {
                "corpname": corpus,
                "q": f"q{cql}",
                "pagesize": 1,
                "ctxattrs": "word,tag",
                "format": "json"
            }
            full_url = f"{base_url}?{urllib.parse.urlencode(params)}"
            print(f"\n🔁 Attempt {attempt_num}")
            print(f"🔗 URL: {full_url}")
            response = requests.get(full_url, timeout=30)
            response.raise_for_status()
            data = response.json()
            result_container.append(data.get("fullsize", 0))
            print(f"📈 Hits on attempt {attempt_num}: {result_container[-1]}")
        except Exception as e:
            print(f"❌ Error on attempt {attempt_num} for CQL={cql}: {e}")
            result_container.append(None)

    print(f"\n📊 Running query: {cql}")

    for attempt in [1, 2]:
        result_holder = []
        thread = threading.Thread(target=fetch_in_thread, args=(attempt, result_holder))
        thread.start()
        thread.join(timeout=TIMEOUT_LIMIT)

        if thread.is_alive():
            print(f"⏰ Timeout on attempt {attempt} for CQL={cql}")
            results.append(None)
        else:
            results.append(result_holder[0] if result_holder else None)

        time.sleep(1)

    time.sleep(5)  # Cooldown after both attempts

    valid_results = [r for r in results if r is not None]
    return max(valid_results) if valid_results else None

# === Main Execution ===
if __name__ == "__main__":
    try:
        input_file = input("📂 Enter full path to Excel file with CQLs: ").strip()
        output_file = input("💾 Enter full path to save output Excel file: ").strip()
        corpus_name = input("📚 Enter the corpus name (e.g., srwac): ").strip()

        df = pd.read_excel(input_file, header=None)
        print(f"📥 Loaded spreadsheet with {df.shape[0]} rows and {df.shape[1]} columns.")

        col_input = input("🔤 Enter one or more column letters to scan (e.g., A,C,E): ").strip().upper()
        col_indices = [ord(col.strip()) - ord('A') for col in col_input.split(",")]

        row_start = int(input("🔢 Enter starting row number (e.g., 2): ")) - 1
        row_end = int(input("🔢 Enter ending row number (inclusive): "))

        print(f"\n⏳ Press Ctrl+C to stop at any time — progress will be saved.")
        print(f"▶️ Scanning columns {col_input} and rows {row_start + 1}–{row_end}")

        for row in range(row_start, row_end):
            for col in col_indices:
                original_value = str(df.iat[row, col]).strip()
                if original_value:
                    print(f"\n📍 Cell R{row + 1}C{col + 1} → CQL: {original_value}")
                    token_count = get_hits(original_value, corpus=corpus_name)
                    df.iat[row, col] = token_count if token_count is not None else "ERROR"

        df.to_excel(output_file, index=False, header=False)
        print(f"\n✅ Final output saved to: {output_file}")

    except KeyboardInterrupt:
        print("\n🛑 Interrupted by user. Saving current progress...")
        df.to_excel(output_file, index=False, header=False)
        print(f"💾 Progress saved to: {output_file}")
        sys.exit(0)
