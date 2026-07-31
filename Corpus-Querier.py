import sys
import threading
import time
import urllib.parse

import pandas as pd
import requests


# === SETTINGS ===

QUERY_DELAY_SECONDS = 5
REQUEST_TIMEOUT_SECONDS = 30
THREAD_TIMEOUT_SECONDS = 300
MAX_ATTEMPTS = 2

BASE_URL = "https://www.clarin.si/ske/bonito/run.cgi/view"


def fetch_once(cql, corpus, attempt_number, result_container):
    """Run one CLARIN corpus query and store the hit count."""

    try:
        params = {
            "corpname": corpus,
            "q": f"q{cql}",
            "pagesize": 1,
            "ctxattrs": "word,tag",
            "format": "json",
        }

        full_url = f"{BASE_URL}?{urllib.parse.urlencode(params)}"

        print(f"\nAttempt {attempt_number}")
        print(f"Querying: {full_url}")

        response = requests.get(
            full_url,
            timeout=REQUEST_TIMEOUT_SECONDS,
        )

        response.raise_for_status()
        data = response.json()

        hits = data.get("fullsize", 0)
        result_container.append(hits)

        print(f"Hits on attempt {attempt_number}: {hits}")

    except Exception as error:
        print(
            f"Error on attempt {attempt_number} "
            f"for CQL '{cql}': {error}"
        )

        result_container.append(None)


def get_hits(cql, corpus):
    """Run a CQL query several times and return the highest valid result."""

    results = []

    print(f"\nRunning query: {cql}")

    for attempt_number in range(1, MAX_ATTEMPTS + 1):
        result_holder = []

        thread = threading.Thread(
            target=fetch_once,
            args=(
                cql,
                corpus,
                attempt_number,
                result_holder,
            ),
        )

        thread.start()
        thread.join(timeout=THREAD_TIMEOUT_SECONDS)

        if thread.is_alive():
            print(
                f"Timeout on attempt {attempt_number} "
                f"for CQL '{cql}'"
            )

            results.append(None)

        elif result_holder:
            results.append(result_holder[0])

        else:
            results.append(None)

        if attempt_number < MAX_ATTEMPTS:
            time.sleep(1)

    time.sleep(QUERY_DELAY_SECONDS)

    valid_results = [
        result
        for result in results
        if result is not None
    ]

    if valid_results:
        return max(valid_results)

    return None


def save_output(dataframe, output_file):
    """Save the current spreadsheet state."""

    try:
        dataframe.to_excel(
            output_file,
            index=False,
            header=False,
        )

        print(f"Saved output to: {output_file}")

    except Exception as error:
        print(f"Could not save output file: {error}")


def column_letters_to_indices(column_input):
    """Convert Excel column letters such as A,C,E into zero-based indices."""

    indices = []

    for column in column_input.split(","):
        column = column.strip().upper()

        if not column:
            continue

        if len(column) != 1 or not column.isalpha():
            raise ValueError(
                f"Invalid column: '{column}'. "
                "This version supports single-letter columns A-Z."
            )

        index = ord(column) - ord("A")
        indices.append(index)

    return indices


def process_file(
    input_file,
    output_file,
    corpus,
    column_indices,
    start_row,
    end_row,
):
    """Process the selected spreadsheet cells."""

    try:
        dataframe = pd.read_excel(
            input_file,
            header=None,
        )

    except Exception as error:
        print(f"Could not load input file: {error}")
        return

    print(
        f"Loaded spreadsheet with "
        f"{dataframe.shape[0]} rows and "
        f"{dataframe.shape[1]} columns."
    )

    if start_row < 1:
        print("Starting row must be at least 1.")
        return

    if end_row < start_row:
        print("Ending row must not precede the starting row.")
        return

    if start_row > len(dataframe):
        print("Starting row is outside the spreadsheet.")
        return

    end_row = min(end_row, len(dataframe))

    for column_index in column_indices:
        if column_index >= len(dataframe.columns):
            print(
                f"Column index {column_index + 1} "
                "is outside the spreadsheet."
            )
            return

    print()
    print("Press Ctrl+C at any time to save progress and stop.")
    print(
        f"Processing rows {start_row}-{end_row} "
        f"in columns "
        f"{[index + 1 for index in column_indices]}"
    )

    try:
        for row_index in range(start_row - 1, end_row):
            print(f"\nProcessing row {row_index + 1}")

            for column_index in column_indices:
                original_value = dataframe.iat[
                    row_index,
                    column_index,
                ]

                if pd.isna(original_value):
                    print(
                        f"Skipping empty cell "
                        f"R{row_index + 1}C{column_index + 1}"
                    )
                    continue

                cql = str(original_value).strip()

                if not cql:
                    continue

                if not cql.startswith("["):
                    print(
                        f"Skipping invalid CQL "
                        f"in R{row_index + 1}"
                        f"C{column_index + 1}: {cql}"
                    )
                    continue

                print(
                    f"Cell R{row_index + 1}"
                    f"C{column_index + 1}: {cql}"
                )

                hits = get_hits(
                    cql,
                    corpus,
                )

                if hits is None:
                    dataframe.iat[
                        row_index,
                        column_index,
                    ] = "ERROR"

                    print(
                        "Query failed. Saving progress "
                        "and stopping."
                    )

                    save_output(
                        dataframe,
                        output_file,
                    )

                    return

                dataframe.iat[
                    row_index,
                    column_index,
                ] = hits

        save_output(
            dataframe,
            output_file,
        )

        # Terminal bell; unlike winsound, this works on Linux terminals.
        print("\a")
        print("All queries completed.")

    except KeyboardInterrupt:
        print("\nInterrupted by user. Saving progress...")

        save_output(
            dataframe,
            output_file,
        )


def main():
    print("Spreadsheet Query Processor for CLARIN.si")

    try:
        input_file = input(
            "Enter full path to the Excel file with CQLs: "
        ).strip().strip('"')

        output_file = input(
            "Enter full path for the output Excel file: "
        ).strip().strip('"')

        corpus = input(
            "Enter the corpus name, for example srwac: "
        ).strip()

        columns_input = input(
            "Enter column letters to scan, "
            "separated by commas, for example A,C,E: "
        ).strip()

        column_indices = column_letters_to_indices(
            columns_input
        )

        if not column_indices:
            print("No valid columns were supplied.")
            return

        start_row = int(
            input(
                "Enter the starting row number: "
            ).strip()
        )

        end_row = int(
            input(
                "Enter the ending row number, inclusive: "
            ).strip()
        )

    except ValueError as error:
        print(f"Invalid input: {error}")
        sys.exit(1)

    except KeyboardInterrupt:
        print("\nCancelled.")
        sys.exit(0)

    process_file(
        input_file=input_file,
        output_file=output_file,
        corpus=corpus,
        column_indices=column_indices,
        start_row=start_row,
        end_row=end_row,
    )


if __name__ == "__main__":
    main()
