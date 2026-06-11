"""
Oracle SQL Query Manager v2
============================
Production-grade CLI tool for querying Oracle databases using oracledb and pandas.
Python 3.10+ required.
"""

from __future__ import annotations

import logging
import math
import os
import re
import time
import traceback
from datetime import datetime
from pathlib import Path
from typing import Iterator

import oracledb
import pandas as pd

# ─────────────────────────────────────────────────────────────────────────────
# LOGGING SETUP  (query_manager.log in working directory)
# ─────────────────────────────────────────────────────────────────────────────
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s  %(levelname)-8s  %(message)s",
    datefmt="%Y-%m-%d %H:%M:%S",
    handlers=[
        logging.FileHandler("query_manager.log", encoding="utf-8"),
        # Remove the StreamHandler so logs don't clutter the terminal UI
    ],
)
log = logging.getLogger("oracle_query_manager")


# ─────────────────────────────────────────────────────────────────────────────
# CONNECTION CONFIGURATION
# Set either DB_SID or DB_SERVICE_NAME, not both.
# If both are set, SERVICE_NAME takes priority.
# ─────────────────────────────────────────────────────────────────────────────
DB_USER     = os.getenv("DB_USER",     "")
DB_PASSWORD = os.getenv("DB_PASSWORD", "")
DB_HOST     = os.getenv("DB_HOST",     "")
DB_PORT     = int(os.getenv("DB_PORT", ""))
DB_SID          = os.getenv("DB_SID",          "")   # e.g. ORCL
DB_SERVICE_NAME = os.getenv("DB_SERVICE_NAME", "")

# Chunk size for streaming rows from Oracle — keep at 50,000 for memory safety
FETCH_CHUNK_SIZE = 50_000

# Rough byte estimate per row used for export size display
BYTES_PER_ROW_ESTIMATE = 200

# Excel hard row limit (including header row)
XLSX_MAX_DATA_ROWS = 1_048_575


# ─────────────────────────────────────────────────────────────────────────────
# ANSI COLOUR HELPERS
# ─────────────────────────────────────────────────────────────────────────────
class C:
    RESET  = "\033[0m"
    BOLD   = "\033[1m"
    GREEN  = "\033[92m"
    YELLOW = "\033[93m"
    RED    = "\033[91m"
    CYAN   = "\033[96m"
    WHITE  = "\033[97m"


def info(msg: str)   -> None: print(f"{C.CYAN}{msg}{C.RESET}")
def ok(msg: str)     -> None: print(f"{C.GREEN}{msg}{C.RESET}")
def warn(msg: str)   -> None: print(f"{C.YELLOW}⚠  {msg}{C.RESET}")
def error(msg: str)  -> None: print(f"{C.RED}✘  {msg}{C.RESET}")
def header(msg: str) -> None: print(f"\n{C.BOLD}{C.WHITE}{msg}{C.RESET}")
def divider()        -> None: print(f"{C.CYAN}{'─' * 62}{C.RESET}")


# ─────────────────────────────────────────────────────────────────────────────
# ORACLE CONNECTION
# ─────────────────────────────────────────────────────────────────────────────
def build_dsn() -> str:
    """
    Build a DSN string.
    Prefers SERVICE_NAME over SID. Raises ValueError if neither is configured.
    """
    if DB_SERVICE_NAME:
        info(f"  Using Service Name : {DB_SERVICE_NAME}")
        return oracledb.makedsn(DB_HOST, DB_PORT, service_name=DB_SERVICE_NAME)
    elif DB_SID:
        info(f"  Using SID          : {DB_SID}")
        return oracledb.makedsn(DB_HOST, DB_PORT, sid=DB_SID)
    else:
        raise ValueError("Either DB_SID or DB_SERVICE_NAME must be configured.")


def get_connection() -> oracledb.Connection:
    """Open and return a new Oracle connection. Raises on failure."""
    dsn = build_dsn()
    return oracledb.connect(user=DB_USER, password=DB_PASSWORD, dsn=dsn)


def test_connection() -> bool:
    """Verify Oracle connectivity. Returns True on success."""
    try:
        conn = get_connection()
        conn.close()
        ok("✔  Oracle connection successful.")
        log.info("Oracle connection test passed.")
        return True
    except (oracledb.DatabaseError, ValueError) as exc:
        error(f"Oracle connection failed: {exc}")
        log.error("Oracle connection test failed: %s", exc)
        return False


# ─────────────────────────────────────────────────────────────────────────────
# SQL VALIDATION  (stronger — checks anywhere in the query)
# ─────────────────────────────────────────────────────────────────────────────

# Forbidden DML / DDL keywords — detected ANYWHERE in the query
_FORBIDDEN_KEYWORDS: tuple[str, ...] = (
    "INSERT", "UPDATE", "DELETE", "MERGE",
    "DROP", "TRUNCATE", "ALTER", "CREATE",
    "EXECUTE", "CALL", "BEGIN", "GRANT", "REVOKE",
)

# Build a pattern that matches any forbidden keyword as a whole word
_FORBIDDEN_RE = re.compile(
    r"\b(" + "|".join(_FORBIDDEN_KEYWORDS) + r")\b",
    re.IGNORECASE,
)

# Allowed entry-point keywords
_ALLOWED_START_RE = re.compile(r"^\s*(SELECT|WITH)\b", re.IGNORECASE)


def validate_query(sql: str) -> tuple[bool, str]:
    """
    Validate that the SQL is a safe, single read-only SELECT or WITH query.

    Rules:
      1. Must not be empty.
      2. Must start with SELECT or WITH.
      3. Must not contain any forbidden DML/DDL keyword anywhere.
      4. Must not contain multiple statements (semicolons are rejected).

    Returns:
        (True, "")            — query is safe to execute.
        (False, reason_str)   — query was rejected with a human-readable reason.
    """
    stripped = sql.strip()

    if not stripped:
        return False, "Query is empty."

    # Rule 4 — reject multiple statements
    # Strip string literals before checking for semicolons to avoid false positives
    no_strings = re.sub(r"'[^']*'", "''", stripped)
    if ";" in no_strings:
        return False, (
            "Multiple statements separated by semicolons are not allowed. "
            "Please submit a single SELECT or WITH query."
        )

    # Rule 3 — forbidden keywords anywhere in the query
    # Strip single-quoted string literals first to avoid false positives
    # e.g.  WHERE name = 'DROP'  should NOT be rejected
    sql_no_literals = re.sub(r"'[^']*'", "''", stripped)
    match = _FORBIDDEN_RE.search(sql_no_literals)
    if match:
        kw = match.group(0).upper()
        return False, (
            f"Forbidden keyword '{kw}' detected in the query. "
            "Only read-only SELECT / WITH queries are permitted."
        )

    # Rule 2 — must start with SELECT or WITH
    if not _ALLOWED_START_RE.match(stripped):
        return False, "Query must start with SELECT or WITH."

    return True, ""


# ─────────────────────────────────────────────────────────────────────────────
# MULTI-LINE SQL INPUT
# ─────────────────────────────────────────────────────────────────────────────
def read_multiline_sql() -> str | None:
    """
    Collect multi-line SQL from the user.
    The user types their query then types RUN on a new line.

    Returns:
        SQL string, or None if the user typed EXIT / QUIT.
    """
    header("Enter your SQL query (SELECT / WITH only).")
    info("Type RUN on a new line to execute, or EXIT to quit.\n")

    lines: list[str] = []
    while True:
        try:
            line = input()
        except (EOFError, KeyboardInterrupt):
            return None

        upper = line.strip().upper()

        if upper == "RUN":
            sql = "\n".join(lines).strip()
            if not sql:
                warn("No SQL entered. Please type your query first.")
                lines = []
                continue
            return sql

        if upper in ("EXIT", "QUIT"):
            return None

        lines.append(line)


# ─────────────────────────────────────────────────────────────────────────────
# ROW COUNTING
# ─────────────────────────────────────────────────────────────────────────────
def count_rows(conn: oracledb.Connection, user_sql: str) -> tuple[int, float]:
    """
    Wrap *user_sql* in COUNT(*) and return (total_rows, elapsed_seconds).
    """
    count_sql = f"SELECT COUNT(*) FROM (\n{user_sql}\n)"
    t0 = time.perf_counter()
    cursor = conn.cursor()
    try:
        cursor.execute(count_sql)
        total = cursor.fetchone()[0]
    finally:
        cursor.close()
    return int(total), time.perf_counter() - t0


# ─────────────────────────────────────────────────────────────────────────────
# VIEW SAMPLE  (option 1)
# ─────────────────────────────────────────────────────────────────────────────
def view_sample(conn: oracledb.Connection, user_sql: str) -> None:
    """Execute the query limited to 20 rows and display via pandas."""
    limited_sql = f"SELECT * FROM (\n{user_sql}\n) WHERE ROWNUM<= 20"
    header("Sample (up to 20 rows)")
    divider()
    log.info("View sample requested.")

    try:
        t0 = time.perf_counter()
        df = pd.read_sql(limited_sql, conn)
        elapsed = time.perf_counter() - t0

        if df.empty:
            warn("Query returned no rows.")
            return

        with pd.option_context(
            "display.max_columns", None,
            "display.max_rows",    20,
            "display.width",       0,
            "display.max_colwidth", 40,
        ):
            print(df.to_string(index=False))

        divider()
        ok(f"  Rows returned  : {len(df)}")
        ok(f"  Columns        : {len(df.columns)}")
        ok(f"  Execution time : {elapsed:.3f}s")
        log.info("Sample fetched: %d rows, %d cols, %.3fs", len(df), len(df.columns), elapsed)

    except oracledb.DatabaseError as exc:
        error(f"Oracle error during sample fetch: {exc}")
        log.error("Sample fetch failed: %s", exc)
    except Exception as exc:
        error(f"Unexpected error: {exc}")
        log.exception("Unexpected error during sample fetch.")


# ─────────────────────────────────────────────────────────────────────────────
# ROW COUNT ONLY  (option 3)
# ─────────────────────────────────────────────────────────────────────────────
def show_row_count(conn: oracledb.Connection, user_sql: str) -> None:
    """Wrap the query in COUNT(*) and display the total."""
    header("Row Count")
    divider()
    info("Counting rows – this may take a moment for large datasets…")
    log.info("Row count requested.")

    try:
        total, elapsed = count_rows(conn, user_sql)
        ok(f"  Total rows     : {total:,}")
        ok(f"  Execution time : {elapsed:.3f}s")
        log.info("Row count result: %d rows in %.3fs", total, elapsed)
    except oracledb.DatabaseError as exc:
        error(f"Oracle error during COUNT: {exc}")
        log.error("Row count failed: %s", exc)
    except Exception as exc:
        error(f"Unexpected error: {exc}")
        log.exception("Unexpected error during row count.")


# ─────────────────────────────────────────────────────────────────────────────
# EXPORT HELPERS — user prompts
# ─────────────────────────────────────────────────────────────────────────────
def _human_size(byte_count: float) -> str:
    """Convert a byte count to a human-readable string."""
    for unit in ("B", "KB", "MB", "GB", "TB"):
        if byte_count < 1024:
            return f"{byte_count:.1f} {unit}"
        byte_count /= 1024
    return f"{byte_count:.1f} PB"


def _human_time(seconds: float) -> str:
    """Format seconds as Xm Ys or Xs."""
    if seconds >= 60:
        m = int(seconds // 60)
        s = int(seconds % 60)
        return f"{m}m {s}s"
    return f"{seconds:.1f}s"


def ask_export_format() -> str | None:
    """Prompt for CSV or XLSX. Returns 'csv', 'xlsx', or None to go back."""
    header("Export Format")
    print("  1. CSV")
    print("  2. XLSX")
    print("  3. Back to Menu")

    while True:
        choice = input("Select format [1/2/3]: ").strip()
        if choice == "1":
            return "csv"
        if choice == "2":
            return "xlsx"
        if choice == "3":
            return None          # signal to caller: user wants to go back
        warn("Enter 1, 2 or 3.")


def ask_rows_per_file() -> int:
    """Prompt for rows-per-file split threshold. Returns an integer >= 1."""
    header("Rows Per File")
    print("  1. 500,000")
    print("  2. 1,000,000")
    print("  3. Custom")
    while True:
        choice = input("Select [1/2/3]: ").strip()
        if choice == "1":
            return 500_000
        if choice == "2":
            return 1_000_000
        if choice == "3":
            raw = input("Enter custom rows per file: ").strip()
            try:
                n = int(raw.replace(",", "").replace("_", ""))
                if n < 1:
                    raise ValueError
                return n
            except ValueError:
                warn("Enter a valid positive integer.")
        else:
            warn("Enter 1, 2, or 3.")


def ask_output_folder() -> Path:
    """
    Prompt for an output folder path.
    Automatically creates the folder if it does not exist.
    """
    header("Output Folder")
    while True:
        raw = input("Enter output folder path (e.g. D:\\Exports): ").strip()
        path = Path(raw)
        if not path.exists():
            warn(f"Folder does not exist: {path}")
            try:
                path.mkdir(parents=True, exist_ok=True)
                ok(f"Created folder: {path}")
                log.info("Created output folder: %s", path)
                return path
            except OSError as exc:
                error(f"Could not create folder: {exc}")
        elif not path.is_dir():
            warn("That path is a file, not a folder. Try again.")
        else:
            return path


def ask_file_prefix() -> str:
    """Prompt for a file name prefix (no extension). Sanitises unsafe chars."""
    header("File Prefix")
    while True:
        prefix = input("Enter file prefix (e.g. ManagedObject): ").strip()
        if not prefix:
            warn("Prefix cannot be empty.")
            continue
        safe = re.sub(r'[\\/:*?"<>|]', "_", prefix)
        if safe != prefix:
            warn(f"Unsafe characters replaced. Using: {safe}")
        return safe


def _resolve_output_path(folder: Path, filename: str) -> Path:
    """
    Return a safe output path.
    If the file already exists, ask the user to overwrite or auto-timestamp.
    """
    target = folder / filename
    if not target.exists():
        return target

    warn(f"File already exists: {target.name}")
    print("  1. Overwrite")
    print("  2. Auto-append timestamp")
    while True:
        choice = input("Choose [1/2]: ").strip()
        if choice == "1":
            return target
        if choice == "2":
            stem   = target.stem
            suffix = target.suffix
            ts     = datetime.now().strftime("%Y%m%d_%H%M%S")
            return folder / f"{stem}_{ts}{suffix}"
        warn("Enter 1 or 2.")


# ─────────────────────────────────────────────────────────────────────────────
# EXPORT CONFIRMATION
# ─────────────────────────────────────────────────────────────────────────────
def confirm_export(
    user_sql:      str,
    total_rows:    int | None,
    fmt:           str,
    rows_per_file: int,
    output_folder: Path,
    file_prefix:   str,
) -> bool:
    """
    Display a summary of the planned export and ask the user to confirm.

    Returns:
        True if the user typed Y, False otherwise.
    """
    header("Export Confirmation")
    divider()
    # Truncate long queries for display
    display_sql = user_sql if len(user_sql) <= 120 else user_sql[:117] + "…"
    print(f"  Query         : {display_sql}")
    print(f"  Total Rows    : {f'{total_rows:,}' if total_rows is not None else 'Unknown (direct export)'}")
    print(f"  Output Format : {fmt.upper()}")
    print(f"  Rows Per File : {rows_per_file:,}")
    print(f"  Output Folder : {output_folder}")
    print(f"  File Prefix   : {file_prefix}")
    divider()

    while True:
        ans = input("Proceed with export? (Y/N): ").strip().upper()
        if ans == "Y":
            return True
        if ans == "N":
            info("Export cancelled.")
            return False
        warn("Enter Y or N.")


# ─────────────────────────────────────────────────────────────────────────────
# EXPORT SUMMARY REPORT
# ─────────────────────────────────────────────────────────────────────────────
def write_export_summary(
    output_folder:  Path,
    user_sql:       str,
    start_time:     datetime,
    end_time:       datetime,
    fmt:            str,
    rows_per_file:  int,
    total_exported: int,
    files_created:  int,
) -> None:
    """
    Write a plain-text export_summary.txt into *output_folder*.
    Safe to call even if the folder somehow disappeared.
    """
    elapsed = (end_time - start_time).total_seconds()
    summary_path = output_folder / "export_summary.txt"

    try:
        with open(summary_path, "w", encoding="utf-8") as f:
            f.write("Oracle SQL Query Manager — Export Summary\n")
            f.write("=" * 50 + "\n\n")
            f.write(f"Query            :\n{user_sql}\n\n")
            f.write(f"Start Time       : {start_time.strftime('%Y-%m-%d %H:%M:%S')}\n")
            f.write(f"End Time         : {end_time.strftime('%Y-%m-%d %H:%M:%S')}\n")
            f.write(f"Elapsed Time     : {_human_time(elapsed)}\n")
            f.write(f"Output Format    : {fmt.upper()}\n")
            f.write(f"Rows Per File    : {rows_per_file:,}\n")
            f.write(f"Total Rows       : {total_exported:,}\n")
            f.write(f"Files Created    : {files_created}\n")
            f.write(f"Output Folder    : {output_folder}\n")
        ok(f"  Summary written : {summary_path}")
        log.info("Export summary written to %s", summary_path)
    except OSError as exc:
        warn(f"Could not write export summary: {exc}")
        log.warning("Export summary write failed: %s", exc)


# ─────────────────────────────────────────────────────────────────────────────
# PROGRESS REPORTER
# ─────────────────────────────────────────────────────────────────────────────
class ProgressReporter:
    """
    Prints a single updating progress line showing:
      Rows Exported  |  Current File  |  Speed  |  ETA
    """

    def __init__(self, total: int | None) -> None:
        self.total      = total   # None when skipping COUNT
        self.exported   = 0
        self.start_time = time.perf_counter()
        self.current_file: str = ""

    def update(self, rows_just_written: int, current_file: str) -> None:
        """Call after each chunk is written to disk."""
        self.exported    += rows_just_written
        self.current_file = current_file
        elapsed           = time.perf_counter() - self.start_time
        speed             = self.exported / elapsed if elapsed > 0 else 0

        if self.total:
            pct = self.exported / self.total * 100
            remaining_rows = self.total - self.exported
            eta_sec = remaining_rows / speed if speed > 0 else 0
            line = (
                f"\r  Rows Exported : {self.exported:>12,} / {self.total:,}"
                f"  ({pct:.1f}%)"
                f"  |  File: {current_file}"
                f"  |  {speed:,.0f} rows/s"
                f"  |  ETA: {_human_time(eta_sec)}"
            )
        else:
            line = (
                f"\r  Rows Exported : {self.exported:>12,}"
                f"  |  File: {current_file}"
                f"  |  {speed:,.0f} rows/s"
            )

        print(line, end="", flush=True)

    def finish(self) -> None:
        """Print a newline to end the progress line."""
        print()


# ─────────────────────────────────────────────────────────────────────────────
# XLSX WRITER
# ─────────────────────────────────────────────────────────────────────────────
class XlsxWriter:
    """
    Streams pandas chunks into split .xlsx files, respecting Excel's row limit
    and the user-specified rows_per_file threshold (whichever is smaller).
    """

    def __init__(self, folder: Path, prefix: str, rows_per_file: int) -> None:
        self.folder        = folder
        self.prefix        = prefix
        self.rows_per_file = min(rows_per_file, XLSX_MAX_DATA_ROWS)
        self.part_number   = 0
        self.rows_in_file  = 0
        self._writer: pd.ExcelWriter | None = None
        self._header_written = False
        self._open_new_file()

    def _open_new_file(self) -> None:
        """Close the current file (if any) and open the next part."""
        if self._writer is not None:
            self._writer.close()
            ok(f"\n  Part {self.part_number} Completed")

        self.part_number   += 1
        self.rows_in_file   = 0
        self._header_written = False

        filename     = f"{self.prefix}_part_{self.part_number}.xlsx"
        target       = _resolve_output_path(self.folder, filename)
        self._writer = pd.ExcelWriter(str(target), engine="openpyxl")
        self._current_filename = target.name

    @property
    def current_filename(self) -> str:
        return getattr(self, "_current_filename", "")

    def write_chunk(self, df: pd.DataFrame) -> None:
        """Write a chunk, splitting to a new file when the limit is reached."""
        remaining = df
        while not remaining.empty:
            space = self.rows_per_file - self.rows_in_file
            batch = remaining.iloc[:space]
            remaining = remaining.iloc[space:]

            if not self._header_written:
                batch.to_excel(self._writer, sheet_name="Data",
                               index=False, startrow=0, header=True)
                self._header_written = True
            else:
                ws = self._writer.sheets["Data"]
                batch.to_excel(self._writer, sheet_name="Data",
                               index=False, startrow=ws.max_row, header=False)

            self.rows_in_file += len(batch)

            if self.rows_in_file >= self.rows_per_file and not remaining.empty:
                self._open_new_file()

    def close(self) -> None:
        """Flush and close the last open file."""
        if self._writer is not None:
            self._writer.close()
            ok(f"\n  Part {self.part_number} Completed")
            self._writer = None

    @property
    def files_created(self) -> int:
        return self.part_number


# ─────────────────────────────────────────────────────────────────────────────
# CSV STREAMING WRITER  (explicit class for clarity + verification counters)
# ─────────────────────────────────────────────────────────────────────────────
class CsvWriter:
    """
    Streams pandas chunks into split .csv files.

    Verification counters guarantee:
      - No row loss
      - No row duplication
      - Header appears exactly once per file
    """

    def __init__(self, folder: Path, prefix: str, rows_per_file: int) -> None:
        self.folder        = folder
        self.prefix        = prefix
        self.rows_per_file = rows_per_file
        self.part_number   = 0
        self.rows_in_file  = 0
        self.total_written = 0          # verification counter
        self._fh           = None       # file handle
        self._current_filename: str = ""
        self._open_new_file()

    def _open_new_file(self) -> None:
        """Close the current file and open the next part."""
        if self._fh is not None:
            self._fh.close()
            ok(f"\n  Part {self.part_number} Completed  →  {self._current_filename}")
            log.info("CSV part %d completed: %s", self.part_number, self._current_filename)

        self.part_number  += 1
        self.rows_in_file  = 0

        filename = f"{self.prefix}_part_{self.part_number}.csv"
        target   = _resolve_output_path(self.folder, filename)
        self._current_filename = target.name
        self._fh = open(target, "w", encoding="utf-8", newline="", buffering=1 << 20)
        self._write_header_next = True  # flag: write header on very next batch

    @property
    def current_filename(self) -> str:
        return self._current_filename

    def write_chunk(self, df: pd.DataFrame) -> None:
        """
        Write a chunk to the current CSV file(s).
        Automatically splits when rows_per_file is reached.
        """
        remaining = df
        while not remaining.empty:
            space = self.rows_per_file - self.rows_in_file
            batch = remaining.iloc[:space]
            remaining = remaining.iloc[space:]

            batch.to_csv(
                self._fh,
                index=False,
                header=self._write_header_next,  # header only at top of each file
            )

            self._write_header_next = False
            written = len(batch)
            self.rows_in_file  += written
            self.total_written += written

            if self.rows_in_file >= self.rows_per_file and not remaining.empty:
                self._open_new_file()

    def close(self) -> None:
        """Flush and close the last open file."""
        if self._fh is not None:
            self._fh.close()
            ok(f"\n  Part {self.part_number} Completed  →  {self._current_filename}")
            log.info("CSV part %d completed: %s", self.part_number, self._current_filename)
            self._fh = None

    @property
    def files_created(self) -> int:
        return self.part_number


# ─────────────────────────────────────────────────────────────────────────────
# CORE EXPORT RUNNER
# ─────────────────────────────────────────────────────────────────────────────
def _run_export(
    conn:          oracledb.Connection,
    user_sql:      str,
    fmt:           str,
    rows_per_file: int,
    output_folder: Path,
    file_prefix:   str,
    total_rows:    int | None,
) -> tuple[int, int]:
    """
    Stream the query result to disk in chunks.

    Args:
        total_rows: Pre-counted total (None if skipped — progress shows no %).

    Returns:
        (rows_exported, files_created)

    Raises:
        Re-raises oracledb.DatabaseError, OSError on unrecoverable errors.
    """
    progress = ProgressReporter(total_rows)

    if fmt == "csv":
        writer: CsvWriter | XlsxWriter = CsvWriter(output_folder, file_prefix, rows_per_file)
    else:
        writer = XlsxWriter(output_folder, file_prefix, rows_per_file)

    for chunk in pd.read_sql(user_sql, conn, chunksize=FETCH_CHUNK_SIZE):
        writer.write_chunk(chunk)
        progress.update(len(chunk), writer.current_filename)

    writer.close()
    progress.finish()

    return writer.total_written if fmt == "csv" else writer.part_number * rows_per_file, writer.files_created


# ─────────────────────────────────────────────────────────────────────────────
# EXPORT DATA  (option 2 — full guided flow)
# ─────────────────────────────────────────────────────────────────────────────
def export_data(conn: oracledb.Connection, user_sql: str) -> None:
    """
    Guided export flow:
      1. Ask whether to count rows first or export directly.
      2. Collect format / rows-per-file / folder / prefix.
      3. Show confirmation screen.
      4. Stream export with progress.
      5. Verify row counts (CSV).
      6. Write export_summary.txt.
    """

    fmt = ask_export_format()
    if fmt is None:
        info("Returning to menu...")
        return

    header("Export Data")
    divider()

    # ── Step 1: count or skip? ────────────────────────────────────────────────
    print("  1. Count Rows First  (safer — shows total + ETA)")
    print("  2. Export Directly   (faster — skips COUNT(*) on huge tables)")
    total_rows: int | None = None

    while True:
        choice = input("Select [1/2]: ").strip()
        if choice == "1":
            info("Counting rows – please wait…")
            try:
                total_rows, count_elapsed = count_rows(conn, user_sql)
                est_bytes = total_rows * BYTES_PER_ROW_ESTIMATE
                ok(f"  Total rows         : {total_rows:,}")
                ok(f"  Estimated export   : {_human_size(est_bytes)}")
                ok(f"  Count query time   : {count_elapsed:.3f}s")
                log.info("Export pre-count: %d rows in %.3fs", total_rows, count_elapsed)
            except oracledb.DatabaseError as exc:
                error(f"COUNT query failed: {exc}")
                log.error("Export pre-count failed: %s", exc)
                return
            break
        if choice == "2":
            info("Skipping COUNT(*) — progress will show rows exported only.")
            break
        warn("Enter 1 or 2.")

    if total_rows == 0:
        warn("Query returned 0 rows. Nothing to export.")
        return

    # ── Step 2: collect export parameters ────────────────────────────────────
    fmt           = ask_export_format()
    rows_per_file = ask_rows_per_file()
    output_folder = ask_output_folder()
    file_prefix   = ask_file_prefix()

    # ── Step 3: confirmation ──────────────────────────────────────────────────
    if not confirm_export(user_sql, total_rows, fmt, rows_per_file, output_folder, file_prefix):
        return

    # ── Step 4: stream export ─────────────────────────────────────────────────
    header("Exporting…")
    divider()

    start_dt    = datetime.now()
    export_t0   = time.perf_counter()
    rows_exported = 0
    files_created = 0

    log.info(
        "Export started — format=%s rows_per_file=%d folder=%s prefix=%s",
        fmt, rows_per_file, output_folder, file_prefix,
    )

    try:
        if fmt == "csv":
            csv_writer = CsvWriter(output_folder, file_prefix, rows_per_file)
            progress   = ProgressReporter(total_rows)

            for chunk in pd.read_sql(user_sql, conn, chunksize=FETCH_CHUNK_SIZE):
                csv_writer.write_chunk(chunk)
                progress.update(len(chunk), csv_writer.current_filename)

            csv_writer.close()
            progress.finish()
            rows_exported = csv_writer.total_written
            files_created = csv_writer.files_created

        else:
            xlsx_writer = XlsxWriter(output_folder, file_prefix, rows_per_file)
            progress    = ProgressReporter(total_rows)
            _count      = 0

            for chunk in pd.read_sql(user_sql, conn, chunksize=FETCH_CHUNK_SIZE):
                xlsx_writer.write_chunk(chunk)
                _count += len(chunk)
                progress.update(len(chunk), xlsx_writer.current_filename)

            xlsx_writer.close()
            progress.finish()
            rows_exported = _count
            files_created = xlsx_writer.files_created

    except oracledb.DatabaseError as exc:
        error(f"Oracle error during export: {exc}")
        log.error("Export failed (Oracle): %s", exc)
        return
    except PermissionError as exc:
        error(f"Permission denied writing to {output_folder}: {exc}")
        log.error("Export failed (permission): %s", exc)
        return
    except OSError as exc:
        error(f"Disk write error: {exc}")
        log.error("Export failed (OS): %s", exc)
        return
    except Exception as exc:
        error(f"Unexpected error: {exc}")
        log.exception("Unexpected export error.")
        traceback.print_exc()
        return

    end_dt  = datetime.now()
    elapsed = time.perf_counter() - export_t0

    # ── Step 5: row verification (CSV) ────────────────────────────────────────
    divider()
    ok("Export Finished")
    ok(f"  Total Files Created  : {files_created}")
    ok(f"  Total Rows Exported  : {rows_exported:,}")
    ok(f"  Output Folder        : {output_folder}")
    ok(f"  Elapsed Time         : {_human_time(elapsed)}")

    if total_rows is not None:
        print()
        info(f"  Rows Expected        : {total_rows:,}")
        info(f"  Rows Exported        : {rows_exported:,}")
        if rows_exported != total_rows:
            warn(
                f"ROW COUNT MISMATCH — expected {total_rows:,}, "
                f"got {rows_exported:,}. Check query and disk space."
            )
            log.warning(
                "Row count mismatch: expected=%d exported=%d",
                total_rows, rows_exported,
            )
        else:
            ok("  ✔  Row count verified — no rows lost or duplicated.")
            log.info("Row count verified OK: %d rows.", rows_exported)

    log.info(
        "Export complete — rows=%d files=%d elapsed=%.1fs",
        rows_exported, files_created, elapsed,
    )

    # ── Step 6: summary report ────────────────────────────────────────────────
    write_export_summary(
        output_folder, user_sql, start_dt, end_dt,
        fmt, rows_per_file, rows_exported, files_created,
    )

# ─────────────────────────────────────────────────────────────────────────────
# MAIN MENU
# ─────────────────────────────────────────────────────────────────────────────
def show_main_menu(conn: oracledb.Connection, user_sql: str) -> None:
    """Display the post-query action menu and dispatch to the chosen option."""
    while True:
        header("Main Menu")
        print("  1. View Sample (20 rows)")
        print("  2. Export Data")
        print("  3. Row Count Only")
        print("  4. New Query")
        print("  5. Exit")
        divider()

        choice = input("Select option [1-5]: ").strip()

        if choice == "1":
            view_sample(conn, user_sql)
        elif choice == "2":
            export_data(conn, user_sql)
        elif choice == "3":
            show_row_count(conn, user_sql)
        elif choice == "4":
            return
        elif choice == "5":
            log.info("Application exited by user.")
            info("Goodbye.")
            raise SystemExit(0)
        else:
            warn("Enter a number from 1 to 5.")


# ─────────────────────────────────────────────────────────────────────────────
# ENTRY POINT
# ─────────────────────────────────────────────────────────────────────────────
def main() -> None:
    """Application entry point."""
    print()
    header("Oracle SQL Query Manager  v2")
    info(f"  Host : {DB_HOST}:{DB_PORT}")
    if DB_SERVICE_NAME:
        info(f"  Service : {DB_SERVICE_NAME}")
    else:
        info(f"  SID  : {DB_SID}")
    info(f"  User : {DB_USER}")
    divider()

    log.info("Oracle SQL Query Manager v2 started.")

    info("Connecting to Oracle…")
    if not test_connection():
        info("Check DB_USER / DB_PASSWORD / DB_HOST / DB_PORT / DB_SERVICE_NAME (or DB_SID).")
        raise SystemExit(1)

    try:
        conn = get_connection()
    except (oracledb.DatabaseError, ValueError) as exc:
        error(f"Failed to open session: {exc}")
        log.critical("Could not open Oracle session: %s", exc)
        raise SystemExit(1)

    try:
        while True:
            user_sql = read_multiline_sql()
            if user_sql is None:
                info("Exiting. Goodbye.")
                log.info("Application exited by user (no query).")
                break

            valid, reason = validate_query(user_sql)
            if not valid:
                error(f"Query rejected: {reason}")
                log.warning("Query rejected: %s", reason)
                continue

            ok("Query accepted.")
            log.info("Query accepted: %.120s", user_sql.replace("\n", " "))

            try:
                show_main_menu(conn, user_sql)
            except SystemExit:
                break

    finally:
        try:
            conn.close()
            info("Oracle connection closed.")
            log.info("Oracle connection closed.")
        except Exception:
            pass


if __name__ == "__main__":
    main()