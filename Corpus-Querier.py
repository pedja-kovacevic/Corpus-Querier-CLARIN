from __future__ import annotations

import queue
import re
import threading
import time
import urllib.parse
from dataclasses import dataclass
from pathlib import Path
from tkinter import filedialog, messagebox

import customtkinter as ctk
import pandas as pd
import requests


APP_NAME = "Corpus Querier"
APP_VERSION = "1.0.0"
BASE_URL = "https://www.clarin.si/ske/bonito/run.cgi/view"
COMMON_CORPORA = ["srwac", "hrwac22_rft1", "bswac", "slwac", "mk_wac"]
PLACEHOLDER_RE = re.compile(r"\{(?:WORD|LEMMA)\}|\b(?:WORD|LEMMA)\b")

ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("blue")


def column_letter_to_index(column: str) -> int:
    column = column.strip().upper()
    if not column or not column.isalpha():
        raise ValueError(f"Invalid Excel column: {column!r}")
    index = 0
    for character in column:
        index = index * 26 + ord(character) - ord("A") + 1
    return index - 1


def column_letters_to_indices(value: str) -> list[int]:
    columns = [item.strip() for item in value.split(",") if item.strip()]
    if not columns:
        raise ValueError("Enter at least one query column.")
    return [column_letter_to_index(item) for item in columns]


def index_to_column_letter(index: int) -> str:
    index += 1
    result = ""
    while index:
        index, remainder = divmod(index - 1, 26)
        result = chr(65 + remainder) + result
    return result


def escape_cql_value(value: object) -> str:
    return str(value).strip().replace("\\", "\\\\").replace('"', '\\"')


def generate_cql(template: str, value: object) -> str:
    escaped = escape_cql_value(value)
    result = template.replace("{WORD}", escaped).replace("{LEMMA}", escaped)
    result = re.sub(r"\bWORD\b", lambda _: escaped, result)
    return re.sub(r"\bLEMMA\b", lambda _: escaped, result)


def format_duration(seconds: float | None) -> str:
    if seconds is None or seconds < 0:
        return "—"
    seconds = int(seconds)
    hours, remainder = divmod(seconds, 3600)
    minutes, seconds = divmod(remainder, 60)
    return f"{hours:d}:{minutes:02d}:{seconds:02d}" if hours else f"{minutes:02d}:{seconds:02d}"


@dataclass
class JobConfig:
    mode: str
    input_file: Path | None
    output_file: Path
    corpus: str
    columns: list[int]
    start_row: int
    end_row: int | None
    words_text: str
    word_source: str
    word_column: int
    templates: list[str]
    delay: float
    timeout: float
    attempts: int
    save_every: int


class QueryError(RuntimeError):
    pass


class QueryWorker(threading.Thread):
    def __init__(self, config: JobConfig, events: queue.Queue, stop_event: threading.Event):
        super().__init__(daemon=True)
        self.config = config
        self.events = events
        self.stop_event = stop_event
        self.session = requests.Session()
        self.session.headers.update(
            {"User-Agent": f"Corpus-Querier/{APP_VERSION} (CLARIN.SI research client)"}
        )

    def emit(self, kind: str, **payload) -> None:
        self.events.put((kind, payload))

    def log(self, text: str) -> None:
        self.emit("log", text=text)

    def fetch_hits(self, cql: str) -> tuple[int, int, float]:
        params = {
            "corpname": self.config.corpus,
            "q": f"q{cql}",
            "pagesize": 1,
            "ctxattrs": "word,tag",
            "format": "json",
        }
        full_url = f"{BASE_URL}?{urllib.parse.urlencode(params)}"
        last_error: Exception | None = None
        started = time.monotonic()

        for attempt in range(1, self.config.attempts + 1):
            if self.stop_event.is_set():
                raise InterruptedError
            self.log(f"Attempt {attempt}/{self.config.attempts}")
            self.log(f"Querying: {full_url}")
            try:
                response = self.session.get(full_url, timeout=self.config.timeout)
                response.raise_for_status()
                data = response.json()
                if "fullsize" not in data:
                    raise QueryError("Response does not contain 'fullsize'.")
                hits = int(data["fullsize"])
                duration = time.monotonic() - started
                self.log(f"Hits: {hits:,} ({duration:.1f} s)")
                return hits, attempt, duration
            except Exception as error:
                last_error = error
                self.log(f"Error: {error}")
                if attempt < self.config.attempts:
                    if self.stop_event.wait(min(2 ** attempt, 10)):
                        raise InterruptedError

        raise QueryError(str(last_error or "Query failed."))

    def load_job(self) -> tuple[pd.DataFrame, list[tuple[int, int, str]]]:
        cfg = self.config
        if cfg.mode == "existing":
            dataframe = pd.read_excel(cfg.input_file, header=None)
            end_row = min(cfg.end_row or len(dataframe), len(dataframe))
            if cfg.start_row > len(dataframe):
                raise ValueError("The starting row is outside the spreadsheet.")
            tasks = []
            for row in range(cfg.start_row - 1, end_row):
                for column in cfg.columns:
                    if column >= len(dataframe.columns):
                        raise ValueError(f"Column {index_to_column_letter(column)} is outside the spreadsheet.")
                    value = dataframe.iat[row, column]
                    if pd.isna(value) or not str(value).strip():
                        continue
                    cql = str(value).strip()
                    if not cql.startswith("["):
                        self.log(f"Skipped invalid CQL at {index_to_column_letter(column)}{row + 1}: {cql}")
                        continue
                    tasks.append((row, column, cql))
            return dataframe, tasks

        if cfg.word_source == "spreadsheet":
            source = pd.read_excel(cfg.input_file, header=None)
            if cfg.word_column >= len(source.columns):
                raise ValueError("The selected word column is outside the spreadsheet.")
            end_row = min(cfg.end_row or len(source), len(source))
            values = source.iloc[cfg.start_row - 1:end_row, cfg.word_column].tolist()
            words = [str(value).strip() for value in values if not pd.isna(value) and str(value).strip()]
        else:
            words = [item.strip() for item in re.split(r"[,;\n]+", cfg.words_text) if item.strip()]

        if not words:
            raise ValueError("No words or lemmas were supplied.")
        if not cfg.templates:
            raise ValueError("Enter at least one CQL template.")

        dataframe = pd.DataFrame([["INPUT", *cfg.templates]])
        tasks = []
        for word in words:
            row = len(dataframe)
            queries = [generate_cql(template, word) for template in cfg.templates]
            dataframe.loc[row] = [word, *queries]
            for offset, cql in enumerate(queries, start=1):
                tasks.append((row, offset, cql))
        return dataframe, tasks

    def save(self, dataframe: pd.DataFrame) -> None:
        self.config.output_file.parent.mkdir(parents=True, exist_ok=True)
        dataframe.to_excel(self.config.output_file, index=False, header=False)
        self.log(f"Saved: {self.config.output_file}")

    def run(self) -> None:
        started = time.monotonic()
        completed = failed = 0
        dataframe: pd.DataFrame | None = None
        try:
            dataframe, tasks = self.load_job()
            total = len(tasks)
            if not total:
                raise ValueError("No valid CQL queries were found.")
            self.emit("started", total=total)
            self.log(f"Prepared {total:,} queries for corpus '{self.config.corpus}'.")

            for position, (row, column, cql) in enumerate(tasks, start=1):
                if self.stop_event.is_set():
                    break
                cell = f"{index_to_column_letter(column)}{row + 1}"
                self.emit("current", position=position, cell=cell, cql=cql)
                self.log(f"\n[{position:,}/{total:,}] Cell {cell}\nRunning query: {cql}")
                try:
                    hits, attempts, duration = self.fetch_hits(cql)
                    dataframe.iat[row, column] = hits
                    completed += 1
                    self.emit("result", success=True, hits=hits, attempts=attempts, duration=duration)
                except QueryError as error:
                    dataframe.iat[row, column] = "ERROR"
                    failed += 1
                    self.log(f"Query failed: {error}")
                    self.emit("result", success=False, error=str(error))

                processed = completed + failed
                elapsed = time.monotonic() - started
                eta = (elapsed / processed) * (total - processed) if processed else None
                self.emit("progress", completed=completed, failed=failed, processed=processed,
                          total=total, elapsed=elapsed, eta=eta)
                if processed % self.config.save_every == 0:
                    self.save(dataframe)
                if position < total and self.stop_event.wait(self.config.delay):
                    break

            self.save(dataframe)
            stopped = self.stop_event.is_set()
            self.emit("finished", stopped=stopped, completed=completed, failed=failed,
                      elapsed=time.monotonic() - started, output=str(self.config.output_file))
        except Exception as error:
            if dataframe is not None:
                try:
                    self.save(dataframe)
                except Exception:
                    pass
            self.emit("fatal", error=str(error))
        finally:
            self.session.close()


class CorpusQuerierApp(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title(f"{APP_NAME} {APP_VERSION} · CLARIN.SI")
        self.geometry("1220x820")
        self.minsize(1060, 720)
        self.events: queue.Queue = queue.Queue()
        self.stop_event = threading.Event()
        self.worker: QueryWorker | None = None
        self.total = 0

        self.grid_columnconfigure(0, weight=0)
        self.grid_columnconfigure(1, weight=1)
        self.grid_rowconfigure(0, weight=1)
        self._build_sidebar()
        self._build_workspace()
        self.after(100, self._poll_events)
        self.protocol("WM_DELETE_WINDOW", self._close)

    def _build_sidebar(self) -> None:
        side = ctk.CTkFrame(self, width=250, corner_radius=0, fg_color=("#E8EEF5", "#111827"))
        side.grid(row=0, column=0, sticky="nsew")
        side.grid_propagate(False)
        ctk.CTkLabel(side, text="CORPUS\nQUERIER", font=ctk.CTkFont(size=26, weight="bold"),
                     justify="left").pack(anchor="w", padx=28, pady=(30, 4))
        ctk.CTkLabel(side, text="CLARIN.SI corpus research", text_color=("#526070", "#94A3B8"),
                     font=ctk.CTkFont(size=13)).pack(anchor="w", padx=28, pady=(0, 30))
        self.status_pill = ctk.CTkLabel(side, text="●  READY", fg_color=("#D9F7E8", "#123D2B"),
                                        text_color=("#087443", "#6EE7B7"), corner_radius=12, height=34)
        self.status_pill.pack(fill="x", padx=24)
        ctk.CTkLabel(side, text="RUN SUMMARY", font=ctk.CTkFont(size=11, weight="bold"),
                     text_color=("#64748B", "#94A3B8")).pack(anchor="w", padx=28, pady=(32, 10))
        self.summary_labels = {}
        for key, title in [("total", "Total queries"), ("done", "Completed"), ("failed", "Failed"),
                           ("elapsed", "Elapsed"), ("eta", "Estimated left")]:
            row = ctk.CTkFrame(side, fg_color="transparent")
            row.pack(fill="x", padx=28, pady=5)
            ctk.CTkLabel(row, text=title, text_color=("#526070", "#94A3B8")).pack(side="left")
            label = ctk.CTkLabel(row, text="0" if key in {"total", "done", "failed"} else "—",
                                 font=ctk.CTkFont(weight="bold"))
            label.pack(side="right")
            self.summary_labels[key] = label
        ctk.CTkLabel(side, text="Appearance", text_color=("#526070", "#94A3B8")).pack(anchor="w", padx=28, pady=(35, 6))
        ctk.CTkOptionMenu(side, values=["Dark", "Light", "System"], command=ctk.set_appearance_mode,
                          height=34).pack(fill="x", padx=24)

    def _build_workspace(self) -> None:
        main = ctk.CTkFrame(self, fg_color="transparent")
        main.grid(row=0, column=1, sticky="nsew", padx=24, pady=22)
        main.grid_columnconfigure(0, weight=1)
        main.grid_rowconfigure(2, weight=1)

        header = ctk.CTkFrame(main, fg_color="transparent")
        header.grid(row=0, column=0, sticky="ew", pady=(0, 14))
        ctk.CTkLabel(header, text="Configure query run", font=ctk.CTkFont(size=25, weight="bold")).pack(side="left")
        self.stop_button = ctk.CTkButton(header, text="Stop & save", fg_color="#B91C1C", hover_color="#991B1B",
                                         state="disabled", command=self._stop)
        self.stop_button.pack(side="right", padx=(10, 0))
        self.start_button = ctk.CTkButton(header, text="▶  Start queries", width=160, command=self._start)
        self.start_button.pack(side="right")

        self.tabs = ctk.CTkTabview(main, height=390)
        self.tabs.grid(row=1, column=0, sticky="ew")
        existing = self.tabs.add("Existing CQL spreadsheet")
        generated = self.tabs.add("Generate CQLs")
        self._build_existing(existing)
        self._build_generated(generated)

        monitor = ctk.CTkFrame(main)
        monitor.grid(row=2, column=0, sticky="nsew", pady=(14, 0))
        monitor.grid_columnconfigure(0, weight=1)
        monitor.grid_rowconfigure(3, weight=1)
        top = ctk.CTkFrame(monitor, fg_color="transparent")
        top.grid(row=0, column=0, sticky="ew", padx=16, pady=(12, 5))
        ctk.CTkLabel(top, text="Live run monitor", font=ctk.CTkFont(size=16, weight="bold")).pack(side="left")
        self.percent_label = ctk.CTkLabel(top, text="0%", font=ctk.CTkFont(weight="bold"))
        self.percent_label.pack(side="right")
        self.progress = ctk.CTkProgressBar(monitor, height=10)
        self.progress.grid(row=1, column=0, sticky="ew", padx=16)
        self.progress.set(0)
        self.current_label = ctk.CTkLabel(monitor, text="No query is running.", anchor="w",
                                          text_color=("#526070", "#94A3B8"))
        self.current_label.grid(row=2, column=0, sticky="ew", padx=16, pady=7)
        self.log_box = ctk.CTkTextbox(monitor, font=ctk.CTkFont(family="Consolas", size=12), wrap="word")
        self.log_box.grid(row=3, column=0, sticky="nsew", padx=12, pady=(0, 12))
        self.log_box.insert("end", "Ready. Configure a run above and select Start queries.\n")
        self.log_box.configure(state="disabled")

    def _entry_row(self, parent, row, label, variable, browse=None, placeholder=""):
        ctk.CTkLabel(parent, text=label).grid(row=row, column=0, sticky="w", padx=(16, 10), pady=7)
        entry = ctk.CTkEntry(parent, textvariable=variable, placeholder_text=placeholder)
        entry.grid(row=row, column=1, sticky="ew", padx=(0, 8), pady=7)
        if browse:
            ctk.CTkButton(parent, text="Browse", width=78, command=browse).grid(row=row, column=2, padx=(0, 16), pady=7)
        return entry

    def _build_existing(self, tab) -> None:
        tab.grid_columnconfigure(1, weight=1)
        self.existing_input = ctk.StringVar()
        self.existing_output = ctk.StringVar()
        self.existing_corpus = ctk.StringVar(value="srwac")
        self.existing_columns = ctk.StringVar(value="A")
        self.existing_start = ctk.StringVar(value="1")
        self.existing_end = ctk.StringVar()
        self._entry_row(tab, 0, "Input Excel file", self.existing_input, lambda: self._choose_open(self.existing_input))
        self._entry_row(tab, 1, "Output Excel file", self.existing_output, lambda: self._choose_save(self.existing_output))
        ctk.CTkLabel(tab, text="Corpus").grid(row=2, column=0, sticky="w", padx=(16, 10), pady=7)
        ctk.CTkComboBox(tab, values=COMMON_CORPORA, variable=self.existing_corpus).grid(row=2, column=1, sticky="ew", padx=(0, 8), pady=7)
        ranges = ctk.CTkFrame(tab, fg_color="transparent")
        ranges.grid(row=3, column=0, columnspan=3, sticky="ew", padx=16, pady=7)
        for i in range(6): ranges.grid_columnconfigure(i, weight=1 if i in {1, 3, 5} else 0)
        ctk.CTkLabel(ranges, text="CQL columns").grid(row=0, column=0, padx=(0, 8))
        ctk.CTkEntry(ranges, textvariable=self.existing_columns).grid(row=0, column=1, sticky="ew", padx=(0, 16))
        ctk.CTkLabel(ranges, text="Start row").grid(row=0, column=2, padx=(0, 8))
        ctk.CTkEntry(ranges, textvariable=self.existing_start, width=90).grid(row=0, column=3, sticky="ew", padx=(0, 16))
        ctk.CTkLabel(ranges, text="End row").grid(row=0, column=4, padx=(0, 8))
        ctk.CTkEntry(ranges, textvariable=self.existing_end, placeholder_text="Last row").grid(row=0, column=5, sticky="ew")
        self._build_advanced(tab, 4, "existing")

    def _build_generated(self, tab) -> None:
        tab.grid_columnconfigure(1, weight=1)
        tab.grid_rowconfigure(4, weight=1)
        self.word_source = ctk.StringVar(value="spreadsheet")
        self.generated_input = ctk.StringVar()
        self.generated_output = ctk.StringVar()
        self.generated_corpus = ctk.StringVar(value="srwac")
        self.word_column = ctk.StringVar(value="A")
        self.generated_start = ctk.StringVar(value="1")
        self.generated_end = ctk.StringVar()
        source = ctk.CTkSegmentedButton(tab, values=["Spreadsheet", "Type or paste words"], command=self._source_changed)
        source.grid(row=0, column=0, columnspan=3, sticky="ew", padx=16, pady=(10, 6))
        source.set("Spreadsheet")
        self.generated_file_label = ctk.CTkLabel(tab, text="Words Excel file")
        self.generated_file_label.grid(row=1, column=0, sticky="w", padx=(16, 10), pady=7)
        self.generated_input_entry = ctk.CTkEntry(tab, textvariable=self.generated_input)
        self.generated_input_entry.grid(row=1, column=1, sticky="ew", padx=(0, 8), pady=7)
        self.generated_browse = ctk.CTkButton(tab, text="Browse", width=78, command=lambda: self._choose_open(self.generated_input))
        self.generated_browse.grid(row=1, column=2, padx=(0, 16), pady=7)
        self.words_box = ctk.CTkTextbox(tab, height=70)
        self.words_box.grid(row=2, column=0, columnspan=3, sticky="ew", padx=16, pady=7)
        self.words_box.insert("1.0", "large, small, long")
        self.words_box.grid_remove()
        details = ctk.CTkFrame(tab, fg_color="transparent")
        details.grid(row=3, column=0, columnspan=3, sticky="ew", padx=16, pady=5)
        for i in range(8): details.grid_columnconfigure(i, weight=1 if i in {1, 3, 5, 7} else 0)
        labels_vars = [("Word column", self.word_column), ("Start row", self.generated_start), ("End row", self.generated_end)]
        for n, (label, variable) in enumerate(labels_vars):
            ctk.CTkLabel(details, text=label).grid(row=0, column=n*2, padx=(0, 6) if n == 0 else (12, 6))
            ctk.CTkEntry(details, textvariable=variable, width=85, placeholder_text="Last" if n == 2 else "").grid(row=0, column=n*2+1, sticky="ew")
        self.word_details = details
        template_frame = ctk.CTkFrame(tab, fg_color="transparent")
        template_frame.grid(row=4, column=0, columnspan=3, sticky="nsew", padx=16, pady=5)
        template_frame.grid_columnconfigure(1, weight=1)
        ctk.CTkLabel(template_frame, text="CQL templates\n(one per line)", justify="left").grid(row=0, column=0, sticky="nw", padx=(0, 10))
        self.templates_box = ctk.CTkTextbox(template_frame, height=86, font=ctk.CTkFont(family="Consolas", size=12))
        self.templates_box.grid(row=0, column=1, sticky="ew")
        self.templates_box.insert("1.0", '[lemma="{LEMMA}"]\n[word="(?i){WORD}"]')
        lower = ctk.CTkFrame(tab, fg_color="transparent")
        lower.grid(row=5, column=0, columnspan=3, sticky="ew", padx=16, pady=(3, 8))
        lower.grid_columnconfigure(1, weight=1)
        lower.grid_columnconfigure(3, weight=1)
        ctk.CTkLabel(lower, text="Corpus").grid(row=0, column=0, padx=(0, 8))
        ctk.CTkComboBox(lower, values=COMMON_CORPORA, variable=self.generated_corpus).grid(row=0, column=1, sticky="ew", padx=(0, 15))
        ctk.CTkLabel(lower, text="Output file").grid(row=0, column=2, padx=(0, 8))
        ctk.CTkEntry(lower, textvariable=self.generated_output).grid(row=0, column=3, sticky="ew", padx=(0, 8))
        ctk.CTkButton(lower, text="Browse", width=78, command=lambda: self._choose_save(self.generated_output)).grid(row=0, column=4)
        self._build_advanced(tab, 6, "generated")

    def _build_advanced(self, parent, row, prefix) -> None:
        frame = ctk.CTkFrame(parent, fg_color="transparent")
        frame.grid(row=row, column=0, columnspan=3, sticky="ew", padx=16, pady=(4, 10))
        variables = {
            "delay": ctk.StringVar(value="5"), "timeout": ctk.StringVar(value="30"),
            "attempts": ctk.StringVar(value="2"), "save": ctk.StringVar(value="25")}
        setattr(self, f"{prefix}_advanced", variables)
        for i, (label, key) in enumerate([("Delay (s)", "delay"), ("Timeout (s)", "timeout"),
                                           ("Attempts", "attempts"), ("Save every", "save")]):
            ctk.CTkLabel(frame, text=label, text_color=("#526070", "#94A3B8")).grid(row=0, column=i*2, padx=(0 if i == 0 else 14, 6))
            ctk.CTkEntry(frame, textvariable=variables[key], width=65).grid(row=0, column=i*2+1)

    def _source_changed(self, value: str) -> None:
        if value == "Spreadsheet":
            self.word_source.set("spreadsheet")
            self.words_box.grid_remove()
            self.generated_file_label.grid()
            self.generated_input_entry.grid()
            self.generated_browse.grid()
            self.word_details.grid()
        else:
            self.word_source.set("typed")
            self.generated_file_label.grid_remove()
            self.generated_input_entry.grid_remove()
            self.generated_browse.grid_remove()
            self.word_details.grid_remove()
            self.words_box.grid()

    def _choose_open(self, variable) -> None:
        path = filedialog.askopenfilename(filetypes=[("Excel workbooks", "*.xlsx"), ("All files", "*.*")])
        if path: variable.set(path)

    def _choose_save(self, variable) -> None:
        path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel workbooks", "*.xlsx")])
        if path: variable.set(path)

    def _number(self, value: str, name: str, integer=False, optional=False):
        value = value.strip()
        if optional and not value: return None
        try: number = int(value) if integer else float(value)
        except ValueError: raise ValueError(f"{name} must be a number.")
        if number <= 0: raise ValueError(f"{name} must be greater than zero.")
        return number

    def _collect_config(self) -> JobConfig:
        generated = self.tabs.get() == "Generate CQLs"
        prefix = "generated" if generated else "existing"
        advanced = getattr(self, f"{prefix}_advanced")
        output = Path(getattr(self, f"{prefix}_output").get().strip().strip('"'))
        if not str(output).strip() or str(output) == ".": raise ValueError("Choose an output Excel file.")
        corpus = getattr(self, f"{prefix}_corpus").get().strip()
        if not corpus: raise ValueError("Enter a corpus name.")
        if generated:
            input_text = self.generated_input.get().strip().strip('"')
            input_file = Path(input_text) if self.word_source.get() == "spreadsheet" else None
            if self.word_source.get() == "spreadsheet" and (not input_text or not input_file.exists()):
                raise ValueError("Choose a valid Excel file containing the words.")
            templates = [line.strip() for line in self.templates_box.get("1.0", "end").splitlines() if line.strip()]
            bad = [template for template in templates if not PLACEHOLDER_RE.search(template)]
            if bad: raise ValueError("Every template must contain {WORD} or {LEMMA}.")
            start = self._number(self.generated_start.get(), "Start row", True) if self.word_source.get() == "spreadsheet" else 1
            end = self._number(self.generated_end.get(), "End row", True, True) if self.word_source.get() == "spreadsheet" else None
            return JobConfig("generated", input_file, output, corpus, [], start, end,
                             self.words_box.get("1.0", "end"), self.word_source.get(),
                             column_letter_to_index(self.word_column.get()), templates,
                             self._number(advanced["delay"].get(), "Delay"), self._number(advanced["timeout"].get(), "Timeout"),
                             self._number(advanced["attempts"].get(), "Attempts", True), self._number(advanced["save"].get(), "Save interval", True))
        input_text = self.existing_input.get().strip().strip('"')
        input_file = Path(input_text)
        if not input_text or not input_file.exists(): raise ValueError("Choose a valid input Excel file.")
        return JobConfig("existing", input_file, output, corpus, column_letters_to_indices(self.existing_columns.get()),
                         self._number(self.existing_start.get(), "Start row", True), self._number(self.existing_end.get(), "End row", True, True),
                         "", "spreadsheet", 0, [], self._number(advanced["delay"].get(), "Delay"),
                         self._number(advanced["timeout"].get(), "Timeout"), self._number(advanced["attempts"].get(), "Attempts", True),
                         self._number(advanced["save"].get(), "Save interval", True))

    def _start(self) -> None:
        try: config = self._collect_config()
        except Exception as error:
            messagebox.showerror(APP_NAME, str(error)); return
        self.stop_event.clear()
        self._reset_monitor()
        self._set_running(True)
        self.worker = QueryWorker(config, self.events, self.stop_event)
        self.worker.start()

    def _stop(self) -> None:
        self.stop_event.set()
        self.status_pill.configure(text="●  STOPPING", fg_color=("#FEF3C7", "#4A3510"), text_color=("#92400E", "#FCD34D"))
        self._append_log("\nStop requested. The current request will finish, then results will be saved.\n")

    def _set_running(self, running: bool) -> None:
        self.start_button.configure(state="disabled" if running else "normal")
        self.stop_button.configure(state="normal" if running else "disabled")
        self.tabs.configure(state="disabled" if running else "normal")
        if running:
            self.status_pill.configure(text="●  RUNNING", fg_color=("#DBEAFE", "#17345C"), text_color=("#1D4ED8", "#93C5FD"))

    def _reset_monitor(self) -> None:
        self.progress.set(0); self.percent_label.configure(text="0%")
        for key in ("total", "done", "failed"): self.summary_labels[key].configure(text="0")
        self.summary_labels["elapsed"].configure(text="—"); self.summary_labels["eta"].configure(text="—")
        self.log_box.configure(state="normal"); self.log_box.delete("1.0", "end"); self.log_box.configure(state="disabled")

    def _append_log(self, text: str) -> None:
        self.log_box.configure(state="normal"); self.log_box.insert("end", text + ("" if text.endswith("\n") else "\n"))
        self.log_box.see("end"); self.log_box.configure(state="disabled")

    def _poll_events(self) -> None:
        try:
            while True:
                kind, payload = self.events.get_nowait()
                if kind == "log": self._append_log(payload["text"])
                elif kind == "started":
                    self.total = payload["total"]; self.summary_labels["total"].configure(text=f"{self.total:,}")
                elif kind == "current":
                    self.current_label.configure(text=f'{payload["position"]:,}/{self.total:,} · {payload["cell"]} · {payload["cql"]}')
                elif kind == "progress":
                    ratio = payload["processed"] / payload["total"]
                    self.progress.set(ratio); self.percent_label.configure(text=f"{ratio:.0%}")
                    self.summary_labels["done"].configure(text=f'{payload["completed"]:,}')
                    self.summary_labels["failed"].configure(text=f'{payload["failed"]:,}')
                    self.summary_labels["elapsed"].configure(text=format_duration(payload["elapsed"]))
                    self.summary_labels["eta"].configure(text=format_duration(payload["eta"]))
                elif kind == "finished":
                    self._set_running(False)
                    status = "STOPPED · SAVED" if payload["stopped"] else "COMPLETE"
                    self.status_pill.configure(text=f"●  {status}", fg_color=("#D9F7E8", "#123D2B"), text_color=("#087443", "#6EE7B7"))
                    self.current_label.configure(text=f'Results saved to {payload["output"]}')
                    messagebox.showinfo(APP_NAME, f'{status.title()}\n\nCompleted: {payload["completed"]:,}\nFailed: {payload["failed"]:,}\nElapsed: {format_duration(payload["elapsed"])}')
                elif kind == "fatal":
                    self._set_running(False)
                    self.status_pill.configure(text="●  ERROR", fg_color=("#FEE2E2", "#4A1515"), text_color=("#B91C1C", "#FCA5A5"))
                    self._append_log(f'\nFatal error: {payload["error"]}')
                    messagebox.showerror(APP_NAME, payload["error"])
        except queue.Empty:
            pass
        self.after(100, self._poll_events)

    def _close(self) -> None:
        if self.worker and self.worker.is_alive():
            if not messagebox.askyesno(APP_NAME, "A run is active. Stop, save, and close when possible?"): return
            self.stop_event.set()
        self.destroy()


if __name__ == "__main__":
    CorpusQuerierApp().mainloop()

