import re
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


def column_letter_to_index(column):
    """Convert one Excel column label into a zero-based index."""

    column = column.strip().upper()

    if not column or not column.isalpha():
        raise ValueError(f"Invalid Excel column: '{column}'")

    index = 0

    for character in column:
        index = index * 26 + ord(character) - ord("A") + 1

    return index - 1


def column_letters_to_indices(column_input):
    """Convert comma-separated Excel column labels into indices."""

    indices = []

    for column in column_input.split(","):
        column = column.strip()

        if column:
            indices.append(column_letter_to_index(column))

    return indices


def index_to_column_letter(index):
    """Convert a zero-based column index into an Excel column label."""

    index += 1
    letters = ""

    while index:
        index, remainder = divmod(index - 1, 26)
        letters = chr(ord("A") + remainder) + letters

    return letters


def escape_cql_value(value):
    """Escape characters that could break a quoted CQL value."""

    return (
        str(value)
        .strip()
        .replace("\\", "\\\\")
        .replace('"', '\\"')
    )


def template_has_placeholder(template):
    """Check whether a template contains a supported placeholder."""

    return bool(
        re.search(
            r"\{(?:WORD|LEMMA)\}|\b(?:WORD|LEMMA)\b",
            template,
        )
    )


def generate_cql(template, input_value):
    """Replace template placeholders with an input word or lemma."""

    escaped_value = escape_cql_value(input_value)

    generated_cql = template

    # Recommended explicit placeholders
    generated_cql = generated_cql.replace(
        "{WORD}",
        escaped_value,
    )
    generated_cql = generated_cql.replace(
        "{LEMMA}",
        escaped_value,
    )

    # Also support bare WORD and LEMMA placeholders
    generated_cql = re.sub(
        r"\bWORD\b",
        lambda match: escaped_value,
        generated_cql,
    )
    generated_cql = re.sub(
        r"\bLEMMA\b",
        lambda match: escaped_value,
        generated_cql,
    )

    return generated_cql


def get_templates():
    """Ask the user for any number of CQL templates."""

    templates = []

    print()
    print("Enter one CQL template at a time.")
    print("Use {WORD} or {LEMMA} as the placeholder.")
    print('Example: [lemma="{LEMMA}"]')
    print(
        'Example: [lemma="{LEMMA}"] '
        '[tag="N.*"]'
    )
    print("Press Enter without typing anything when finished.")

    while True:
        template_number = len(templates) + 1

        template = input(
            f"Template {template_number}: "
        ).strip()

        if not template:
            break

        if not template_has_placeholder(template):
            print(
                "That template does not contain {WORD}, "
                "{LEMMA}, WORD or LEMMA."
            )

            use_anyway = input(
                "Add it anyway? (y/n): "
            ).strip().lower()

            if use_anyway not in {"y", "yes"}:
                print("Template skipped.")
                continue

        templates.append(template)
        print("Template added.")

    return templates


def get_words_from_keyboard():
    """Read comma-separated input words from the keyboard."""

    print()
    print("Enter words or lemmas separated by commas.")
    print("Example: large, small, long, young")

    word_input = input("Words or lemmas: ").strip()

    words = [
        word.strip()
        for word in word_input.split(",")
        if word.strip()
    ]

    return words


def get_words_from_spreadsheet():
    """Load words or lemmas from a selected spreadsheet column."""

    input_file = input(
        "Enter full path to the Excel file containing "
        "the words or lemmas: "
    ).strip().strip('"')

    column = input(
        "Enter the column containing the words or lemmas, "
        "for example A: "
    ).strip()

    column_index = column_letter_to_index(column)

    start_row = int(
        input(
            "Enter the first row containing a word or lemma: "
        ).strip()
    )

    end_row_input = input(
        "Enter the final row, or press Enter to use "
        "the rest of the column: "
    ).strip()

    try:
        dataframe = pd.read_excel(
            input_file,
            header=None,
        )

    except Exception as error:
        print(f"Could not load input file: {error}")
        return []

    if column_index >= len(dataframe.columns):
        print(
            f"Column {column.upper()} is outside the spreadsheet."
        )
        return []

    if start_row < 1 or start_row > len(dataframe):
        print("The starting row is outside the spreadsheet.")
        return []

    if end_row_input:
        end_row = int(end_row_input)
        end_row = min(end_row, len(dataframe))
    else:
        end_row = len(dataframe)

    if end_row < start_row:
        print("The ending row precedes the starting row.")
        return []

    words = []

    for row_index in range(start_row - 1, end_row):
        value = dataframe.iat[row_index, column_index]

        if pd.isna(value):
            continue

        value = str(value).strip()

        if value:
            words.append(value)

    print(f"Loaded {len(words)} words or lemmas.")

    return words


def create_generated_dataframe(words, templates):
    """Create a spreadsheet containing generated CQLs."""

    rows = []

    # The first row identifies the source and template columns.
    header_row = ["INPUT"] + templates
    rows.append(header_row)

    for word in words:
        generated_queries = [
            generate_cql(template, word)
            for template in templates
        ]

        rows.append([word] + generated_queries)

    return pd.DataFrame(rows)


def process_dataframe(
    dataframe,
    output_file,
    corpus,
    column_indices,
    start_row,
    end_row,
):
    """Process selected cells in an already loaded dataframe."""

    print(
        f"Spreadsheet has "
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
                f"Column {index_to_column_letter(column_index)} "
                "is outside the spreadsheet."
            )
            return

    print()
    print("Press Ctrl+C at any time to save progress and stop.")
    print(
        f"Processing rows {start_row}-{end_row} "
        f"in columns "
        f"{[index_to_column_letter(i) for i in column_indices]}"
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
                        f"R{row_index + 1}"
                        f"C{column_index + 1}"
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

        print("\a")
        print("All queries completed.")

    except KeyboardInterrupt:
        print("\nInterrupted by user. Saving progress...")

        save_output(
            dataframe,
            output_file,
        )


def process_existing_cql_spreadsheet():
    """Run the original spreadsheet-processing workflow."""

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

    try:
        dataframe = pd.read_excel(
            input_file,
            header=None,
        )

    except Exception as error:
        print(f"Could not load input file: {error}")
        return

    process_dataframe(
        dataframe=dataframe,
        output_file=output_file,
        corpus=corpus,
        column_indices=column_indices,
        start_row=start_row,
        end_row=end_row,
    )


def process_automatically_generated_cqls():
    """Generate CQLs from words or lemmas and execute them."""

    print()
    print("How would you like to supply the words or lemmas?")
    print("1 - Load them from an Excel spreadsheet")
    print("2 - Enter them directly")

    input_method = input("Choose 1 or 2: ").strip()

    if input_method == "1":
        words = get_words_from_spreadsheet()

    elif input_method == "2":
        words = get_words_from_keyboard()

    else:
        print("Invalid input method.")
        return

    if not words:
        print("No words or lemmas were supplied.")
        return

    templates = get_templates()

    if not templates:
        print("No CQL templates were supplied.")
        return

    corpus = input(
        "Enter the corpus name, for example srwac: "
    ).strip()

    output_file = input(
        "Enter full path for the output Excel file: "
    ).strip().strip('"')

    dataframe = create_generated_dataframe(
        words,
        templates,
    )

    print()
    print(
        f"Generated {len(words) * len(templates)} CQL queries "
        f"from {len(words)} input items and "
        f"{len(templates)} templates."
    )

    # Column A contains the original words.
    # Generated queries begin in column B.
    query_column_indices = list(
        range(1, len(templates) + 1)
    )

    # Row 1 contains column descriptions.
    # Queries therefore begin in row 2.
    process_dataframe(
        dataframe=dataframe,
        output_file=output_file,
        corpus=corpus,
        column_indices=query_column_indices,
        start_row=2,
        end_row=len(dataframe),
    )


def main():
    print("Spreadsheet Query Processor for CLARIN.si")
    print()
    print("Choose an operating mode:")
    print("1 - I already have a spreadsheet containing CQLs")
    print("2 - Generate CQLs automatically from words or lemmas")

    try:
        mode = input("Choose 1 or 2: ").strip()

        if mode == "1":
            process_existing_cql_spreadsheet()

        elif mode == "2":
            process_automatically_generated_cqls()

        else:
            print("Invalid mode. Enter either 1 or 2.")

    except ValueError as error:
        print(f"Invalid input: {error}")
        sys.exit(1)

    except KeyboardInterrupt:
        print("\nCancelled.")
        sys.exit(0)


if __name__ == "__main__":
    main()