# Corpus Querier 2.0

Corpus Querier is a desktop application for running **CQL (Corpus Query Language)** queries against corpora hosted by [CLARIN.SI](https://www.clarin.si/). It can calculate total query frequencies, extract the word forms returned by a query, and report the frequency of every extracted form.

The application works with existing CQL spreadsheets or generates queries automatically from words and reusable `{WORD}` and `{LEMMA}` templates.

## Download and install

Download the appropriate file from the [latest GitHub release](https://github.com/pedja-kovacevic/Corpus-Querier-CLARIN/releases/latest):

- **Windows:** `Corpus-Querier-Setup.exe`
- **Ubuntu, Debian or Linux Mint:** `corpus-querier_2.0.0_amd64.deb`
- **Other x86-64 Linux distributions:** `Corpus-Querier-x86_64.AppImage`

Python, PowerShell, VS Code and manual dependency installation are not required when using these packages.

### Windows

Run `Corpus-Querier-Setup.exe` and follow the installation wizard. The installer:

- installs the application for the current user;
- creates a Start Menu shortcut;
- optionally creates a Desktop shortcut;
- adds Corpus Querier to Windows **Installed apps**;
- includes Python and all runtime dependencies.

Windows SmartScreen may display a warning because the installer is not commercially code-signed.

### Ubuntu, Debian and Linux Mint

Open the `.deb` file with the system Software Installer, or run:

```bash
sudo apt install ./corpus-querier_2.0.0_amd64.deb
```

### AppImage

The AppImage is portable and does not install anything:

```bash
chmod +x Corpus-Querier-x86_64.AppImage
./Corpus-Querier-x86_64.AppImage
```

Some Linux distributions require the FUSE 2 compatibility package to run AppImages.

## What is new in version 2.0

- Four selectable output modes: total frequencies, forms, forms with frequencies, or totals together with forms and frequencies.
- Automatic extraction of surface forms returned by CQL queries.
- Configurable minimum-attestation threshold.
- Optional capitalization-sensitive or capitalization-insensitive grouping.
- Structured Excel output for form-based research.
- A reproducible `Run Information` sheet recording corpus and query settings.
- Full backward compatibility with the original frequency-count workflow.
- Every configured query attempt is always executed; the default is two attempts.

## Input modes

### Existing CQL spreadsheet

Choose **Existing CQL spreadsheet** when the workbook already contains complete CQL expressions. Select the CQL columns and row range in the interface.

Example:

```cql
[lemma="čitati" & tag="Vmr3s"]
```

### Generate CQLs

Choose **Generate CQLs** to create queries from words or lemmas. Input items can be loaded from an Excel column or pasted directly into the application.

Example input items:

```text
čitati, učiti, putovati
```

Example templates:

```cql
[lemma="{LEMMA}"]
[lemma="{LEMMA}" & tag="Vmr3s"]
[word="{WORD}"]
```

Every template is applied to every input item.

## Output modes

### Total frequency

Returns only the total number of matches for each CQL. This mode preserves the original spreadsheet structure: every processed CQL cell is replaced by its highest valid hit count.

| INPUT | `[lemma="{LEMMA}"]` | `[lemma="{LEMMA}" & tag="Vmr3s"]` |
|---|---:|---:|
| čitati | 125000 | 42000 |
| putovati | 38000 | 11500 |

No additional worksheets are created in this mode.

### Forms only

Extracts distinct surface forms but does not include their token frequencies in the `Forms` sheet.

### Forms with frequencies

Extracts every qualifying form and places its frequency in a separate numeric Excel cell.

| QUERY ID | INPUT | FORM | FREQUENCY | RANK |
|---|---|---|---:|---:|
| Q000001 | čitati | čita | 42000 | 1 |
| Q000001 | čitati | Čita | 850 | 2 |

### Total + forms with frequencies

Provides the complete result: the overall number of CQL matches, each qualifying surface form, its individual frequency, its rank and its percentage of the total.

## Structured workbooks

Every forms-based mode creates three worksheets:

1. **Queries** — one row per executed query, including the input, template, generated CQL, status, attempts and duration. The combined mode also includes total hits.
2. **Forms** — one row per extracted form, with separate cells for every value.
3. **Run Information** — corpus, output mode, capitalization setting, threshold, attempts, delay, timeout and timestamps.

The worksheets have frozen headers, filters and adjusted column widths.

## Capitalization

When **Capitalization matters** is enabled, forms such as `čita`, `Čita` and `ČITA` remain separate.

When it is disabled, capitalization variants are combined into one form and their frequencies are added before the minimum-attestation threshold is applied.

## Minimum attestations

The minimum-attestation setting determines how frequently a surface form must occur to appear in the output.

- `1` includes every attested form.
- `10` excludes forms occurring fewer than ten times.
- Larger values can be used to reduce spelling variation and annotation noise.

The total CQL frequency is not changed by this threshold.

## Repeated queries

Corpus Querier performs every configured attempt. With the default setting of two attempts, each CQL is queried twice even when the first attempt succeeds.

- Total-frequency mode saves the highest valid total.
- Forms-based modes retain the highest observed frequency for each form.
- Failed attempts are ignored if another attempt succeeds.

This behaviour is intentional because some corpora may return an incomplete result on an initial request.

## Progress and saving

The interface displays:

- the current CQL and spreadsheet cell;
- completed and failed queries;
- elapsed time and estimated time remaining;
- a progress bar and detailed live log.

Results are saved periodically. **Stop & save** finishes the current request and preserves the results collected so far.

## Running from source

Python 3.10 or newer is recommended.

```bash
python -m pip install -r requirements.txt
python Corpus-Querier.py
```

## Building installers

GitHub Actions builds the Windows installer, Linux AppImage and Debian package automatically when a version tag is pushed.

For a local Windows build, install Python 3 and [Inno Setup 6](https://jrsoftware.org/isdl.php), then run:

```powershell
powershell -ExecutionPolicy Bypass -File .\build_installer.ps1
```

The installer is written to `installer-output\Corpus-Querier-Setup.exe`.

## License

Corpus Querier is released under the [MIT License](LICENSE).

## Acknowledgments

Corpus Querier uses the corpus-querying infrastructure provided by [CLARIN.SI](https://www.clarin.si/).
