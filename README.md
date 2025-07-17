# 🧠 Corpus Querier for CLARIN.SI Corpora

This Python script automates querying [CLARIN.SI's Sketch Engine](https://www.clarin.si) corpora using **CQL (Corpus Query Language)** expressions from an Excel spreadsheet. It fetches token frequencies for each CQL pattern and writes them back to the spreadsheet.



---

## ✨ Features

- 📥 **Reads CQLs from Excel (.xlsx)**: Choose columns and row ranges interactively.
- 🔁 **Runs each query twice**: Ensures reliable results by returning the **maximum** count of both attempts.
- 📚 **Custom corpus support**: Enter any available CLARIN.SI corpus name (e.g., `srwac`, `hrwac22_rft1`).
- 🕐 **Timeout safety**: Each query attempt has a 5-minute timeout. Unfinished queries return `"ERROR"`.
- ⏳ **Rate-limited requests**: Adds a 5-second delay between CQLs to avoid server overload.
- 💾 **Auto-save on Ctrl+C**: Interrupting the script mid-run will save progress and exit cleanly.
- 📈 **Writes token counts** into the spreadsheet in place of the original CQLs.

---

## 🚀 How to Use

### 1. 🔧 Requirements

Install required packages:

```bash
pip install requests pandas openpyxl

2. 🗂 Prepare Your Excel File
Save your .xlsx file with one CQL query per cell.

Empty or irrelevant cells will be ignored.

3. ▶️ Run the Script
Launch the script in your terminal:

python your_script_name.py

You will be prompted for:

📂 Input Excel file path

💾 Output file path for the results

📚 Corpus name (e.g., srwac)

🔤 Columns to process (e.g., A,C,E)

🔢 Row range (e.g., from row 2 to 100)

The script will then:

Query each CQL string twice

Wait 5 seconds between queries

Replace the CQL in the Excel file with the resulting frequency

📌 Example
Input Excel Cell:

[lemma="udruženje" & !tag="N.*s.*"] [tag=".*g" & !tag="S.*"]

Resulting Query:
https://www.clarin.si/ske/bonito/run.cgi/view?corpname=srwac&q=q[lemma="udruženje" & !tag="N.*s.*"] [tag=".*g" & !tag="S.*"]

Output Cell Value:
6

💡 Notes
If a query fails or times out, it will be replaced with "ERROR".

Output is saved as an Excel file with CQLs replaced by token counts.

You can safely abort the script mid-run using Ctrl+C.

📜 License
This project is open-source and released under the MIT License.

🤝 Acknowledgments
Special thanks to CLARIN.SI for providing open access to powerful corpus querying tools.



