# 🧠 Corpus Querier for CLARIN.SI Corpora

Corpus Querier is a Python program for automatically querying corpora hosted by [CLARIN.SI’s Sketch Engine](https://www.clarin.si).

The program can either:

1. Process complete **CQL (Corpus Query Language)** expressions stored in an Excel spreadsheet.
2. Automatically generate CQL expressions by inserting words or lemmas into reusable CQL templates.

It retrieves token frequencies and saves the results in an Excel file.

---

## ✨ Features

- 📥 **Reads complete CQLs from Excel (`.xlsx`)**
- 🧩 **Generates CQLs automatically from templates**
- ⌨️ **Accepts words or lemmas entered directly**
- 📊 **Accepts words or lemmas from an Excel column**
- ♾️ **Supports any number of CQL templates**
- 🔤 **Supports Excel columns from `A` to `Z` and beyond**, such as `AA` and `AB`
- 🔁 **Runs each query twice** and returns the highest valid result
- 📚 **Supports any available CLARIN.SI corpus**
- 🕐 **Uses request and thread timeouts**
- ⏳ **Waits five seconds between CQLs** to reduce server load
- 💾 **Saves progress when interrupted with `Ctrl+C`**
- 🚨 **Writes `ERROR` and stops if both attempts fail**

---

## 🔧 Requirements

Install Python 3 and the required packages:

```bash
pip install requests pandas openpyxl
```

---

## ▶️ Running the Program

Open a terminal in the project directory and run:

```bash
python "Corpus Querier.py"
```

The program first asks you to choose an operating mode:

```text
1 - I already have a spreadsheet containing CQLs
2 - Generate CQLs automatically from words or lemmas
```

---

## Mode 1: Process Existing CQLs

Choose this mode if your Excel spreadsheet already contains complete CQL expressions.

### Preparing the spreadsheet

Save the spreadsheet as an `.xlsx` file with one CQL expression per cell.

Example:

```cql
[lemma="udruženje" & !tag="N.*s.*"] [tag=".*g" & !tag="S.*"]
```

Empty cells and cells that do not begin with `[` are skipped.

### Required information

The program asks for:

- the input Excel file path;
- the output Excel file path;
- the CLARIN.SI corpus name;
- the columns containing CQLs;
- the first row to process;
- the final row to process.

Multiple columns can be entered using commas:

```text
A,C,E
```

Columns beyond `Z` are also supported:

```text
A,AA,AB
```

### Output

Each processed CQL is replaced by its token frequency.

For example:

```cql
[lemma="udruženje" & !tag="N.*s.*"] [tag=".*g" & !tag="S.*"]
```

may be replaced by:

```text
6
```

---

## Mode 2: Generate CQLs Automatically

Choose this mode when you have words or lemmas but have not created the individual CQL expressions.

The program can obtain the input items from:

1. an Excel spreadsheet; or
2. words or lemmas entered directly in the terminal.

---

### Entering Words Directly

Enter words or lemmas separated by commas:

```text
large, small, long, young
```

Leading and trailing spaces are removed automatically.

---

### Loading Words from Excel

If the words are stored in Excel, the program asks for:

- the Excel file path;
- the column containing the words;
- the first row containing a word;
- the final row.

Press Enter instead of providing a final row to read the remainder of the selected column.

Empty cells are ignored.

---

## 🧩 CQL Templates

After loading the words, enter one or more CQL templates.

Use either `{WORD}` or `{LEMMA}` as the placeholder for the input item.

Recommended examples:

```cql
[word="{WORD}"]
```

```cql
[lemma="{LEMMA}"]
```

```cql
[lemma="{LEMMA}"] [tag="N.*"]
```

```cql
[word="very"] [lemma="{LEMMA}"]
```

The program also accepts bare `WORD` and `LEMMA` placeholders, but the forms with braces are recommended because they are more explicit.

Enter one template at a time:

```text
Template 1: [lemma="{LEMMA}"]
Template 2: [word="very"] [lemma="{LEMMA}"]
Template 3:
```

Press Enter without typing another template to finish.

There is no fixed limit on the number of templates.

---

## 🔄 CQL Generation

For every input item, the program replaces the placeholder in every template.

For example, given the input:

```text
large
```

and the template:

```cql
[lemma="{LEMMA}"] [tag="N.*"]
```

the generated CQL is:

```cql
[lemma="large"] [tag="N.*"]
```

If four words and three templates are supplied, the program generates and executes twelve queries.

---

## 📊 Automatically Generated Output

In automatic mode:

- column `A` contains the original words or lemmas;
- each additional column corresponds to one CQL template;
- the first row records the templates;
- the remaining cells contain the resulting token frequencies.

Example:

| INPUT | `[lemma="{LEMMA}"]` | `[word="very"] [lemma="{LEMMA}"]` |
|---|---:|---:|
| large | 154321 | 8241 |
| small | 176543 | 9123 |
| long | 198765 | 7456 |

---

## 🔁 Query Behaviour

Each generated or supplied CQL is queried twice.

The program:

1. sends the query to CLARIN.SI;
2. records the result of each attempt;
3. ignores failed attempts;
4. returns the highest valid hit count;
5. waits five seconds before processing the next CQL.

If both attempts fail, the affected cell is replaced with:

```text
ERROR
```

The program then saves the current results and stops.

---

## 💾 Interrupting and Saving

Press:

```text
Ctrl+C
```

at any time to interrupt processing.

The program catches the interruption and saves all results collected up to that point.

---

## 🌐 Query Endpoint

Queries are sent to:

```text
https://www.clarin.si/ske/bonito/run.cgi/view
```

A generated request includes:

- the corpus name;
- the URL-encoded CQL;
- JSON output format;
- a page size of one, because only the total frequency is needed.

---

## 💡 Notes

- Input and output files must use the `.xlsx` format.
- The output file should normally have a different name from the input file.
- The corpus name must match a corpus available through CLARIN.SI.
- Generated CQL values are escaped if they contain quotation marks or backslashes.
- Cells containing invalid CQLs are skipped.
- Zero is treated as a valid query result.
- Internet access is required.

---

## 📜 License

This project is open source and released under the MIT License.

---

## 🤝 Acknowledgments

Special thanks to [CLARIN.SI](https://www.clarin.si) for providing open access to corpus resources and corpus-querying infrastructure.