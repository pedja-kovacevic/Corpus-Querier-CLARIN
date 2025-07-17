# Corpus-Querier-CLARIN
# 📊 Corpus Querier for CLARIN.si

This Python script queries the [CLARIN.si](https://www.clarin.si/) corpus infrastructure to retrieve total hit counts for CQL (Corpus Query Language) queries stored in an Excel spreadsheet. The results replace the original queries in the spreadsheet with their corresponding frequency counts and save the output as a new Excel file.

---

## ✅ Features

- 📥 Reads queries from specified columns and row ranges in an Excel file.
- 🔍 Sends queries to the CLARIN.si `ske/bonito/run.cgi/first` API endpoint.
- 🧠 Supports complex CQL expressions such as `[lemma="tata"]`, `[word="bio"]`, etc.
- 🕓 Enforces a delay (default: 5 seconds) between queries to respect server stability.
- 🧯 Gracefully handles keyboard interruption (`CTRL+C`) and saves progress.
- 💥 Automatically catches and logs errors and saves partial output if a failure occurs.
- 🔔 Emits a completion ping when finished (on supported platforms like Windows).

---

## ⚙️ How It Works

### Constants and Configuration

- **API Endpoint:** `https://www.clarin.si/ske/bonito/run.cgi/first`
- **Corpus Name:** `srwac` (can be modified in the script).
- **Excel Input/Output:**
  - User is prompted to provide file paths for both input and output files.
- **Query Delay:** 5 seconds between queries by default.
- **Columns/Rows:** User specifies which columns and rows to scan via prompts.

---

### Querying CLARIN.si (`get_hits` function)

- Accepts a single CQL query string.
- Sends a GET request to the CLARIN.si Bonito API.
- Retrieves the total hit count from the `fullsize` field in the JSON response.
- Returns the count or `None` if the request fails or times out.

---

### Main Script Logic

1. Loads the input Excel file into a `pandas` DataFrame.
2. Prompts the user to enter:
   - Full path to the input file
   - Full path for saving the output
   - Start and end row numbers (1-based)
   - Columns to scan (e.g., `B,C,D`)
3. Iterates over each specified cell:
   - Skips empty or invalid CQL expressions
   - Calls `get_hits` to retrieve hit count
   - Replaces the query string in the DataFrame with the numeric result
   - Waits 5 seconds between requests
4. Displays progress in the terminal.
5. On completion or interruption, writes the collected data to the output file.
6. Emits a system beep if supported by the OS.

---

## 🧪 Usage

Run the script from terminal:

```bash
python clarin_query.py

Follow the prompts:

📥 Spreadsheet Query Processor for CLARIN.si
📝 Enter full path to input Excel file: C:\path\to\queries.xlsx
💾 Enter full path to save output Excel file: C:\path\to\output.xlsx
🔢 Start from row (1-based): 1
🔢 End at row (1-based): 100
🔠 Enter comma-separated column names to scan (e.g., B,C,D): B,C

 Output
The resulting Excel file mirrors the original structure.

All queried cells are overwritten with their frequency counts (integers).

🛡️ Error Handling and Safety
Saves output immediately on error or CTRL+C.

If a request fails or hangs, the script will timeout after a defined wait period and write all progress so far.

You can re-run the script on skipped/failed parts.

📝 Notes
This script uses the CLARIN.si Bonito API (ske/bonito/run.cgi/first), not the deprecated noske interface.

CQL syntax must be valid — malformed queries may return 0 or error.

Keep a backup of your original Excel file, as queries are overwritten in-place.

The script uses a conservative delay to avoid overloading the public infrastructure.

Designed and tested on the srWaC corpus.

🧰 Requirements
Python 3.8+

Dependencies:

pandas

requests

openpyxl
