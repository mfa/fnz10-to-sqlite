import io
import datetime
import re
from pathlib import Path

import diskcache
import httpx
import openpyxl
import typer
from sqlite_utils import Database

cache = diskcache.Cache(Path(__file__).parent / ".cache")


@cache.memoize()
def download(year: int, month: int) -> bytes:
    url = (
        "https://www.kba.de/SharedDocs/Downloads/DE/Statistik/Fahrzeuge/FZ10/"
        + f"fz10_{year}_{month:02d}.xlsx?__blob=publicationFile&v=3"
    )
    response = httpx.get(url)
    if response.status_code == 200:
        return response.content
    raise NotImplementedError


def parse_xslx(blob: io.BytesIO):
    """
    Parse the downloaded Excel blob and yield rows as dicts in long form:
    - 'marke' and 'modellreihe' from row 9 header
    - 'kategorie' from row 8 header
    - fields: year, month (e.g. 'Juni'), count (single-month total)
    Skips summary rows where marke contains 'INSGESAMT' or 'ZUSAMMEN'.
    """
    wb = openpyxl.load_workbook(blob, read_only=True, data_only=True)
    sheet = wb["FZ 10.1"]
    rows = sheet.iter_rows(values_only=True)
    # advance to header rows (8 & 9)
    for _ in range(7):
        next(rows, None)
    header8 = next(rows, ())
    header9 = next(rows, ())
    # forward-fill merged/group headers in row 8
    filled8 = []
    last = None
    for cell in header8:
        if cell is not None:
            last = cell
        filled8.append(last)
    # locate static columns (Marke, Modellreihe) and define data blocks
    idx_marke = header9.index("Marke")
    idx_modell = header9.index("Modellreihe")
    start = idx_modell + 1
    # each block has 3 columns: month, month-range, percentage
    block = 3
    # categories = distinct row8 values per block
    categories = [filled8[j] for j in range(start, len(header9), block)]

    last_marke = None
    def _to_int(val):
        try:
            return int(val)
        except Exception:
            return None

    for row in rows:
        if not any(cell is not None for cell in row):
            continue
        # fill down Marke
        if row[idx_marke] is not None:
            last_marke = row[idx_marke]
        marke = last_marke
        # skip overall or summary rows in the marke field
        if isinstance(marke, str) and re.search(r"insgesamt|zusammen", marke, re.IGNORECASE):
            continue
        modell = row[idx_modell]
        # skip summary rows
        if modell == "ZUSAMMEN":
            continue

        # pivot each category block into a record with year, month, range, and count
        for bi, cat in enumerate(categories):
            base = start + bi * block
            raw_month = header9[base]
            raw_range = header9[base + 1]
            value_month = row[base]
            value_range = row[base + 1]
            # parse year and month/range text
            # extract month name and year from the single-month header
            raw_month_str = str(raw_month).strip()
            m1 = re.match(r"(.+?)\s+(\d{4})", raw_month_str)
            if m1:
                month_label = m1.group(1).strip()
                year = int(m1.group(2))
            else:
                month_label = raw_month_str
                year = None
            yield {
                "marke": marke,
                "modellreihe": modell,
                "kategorie": cat,
                "year": year,
                "month": month_label,
                "count": _to_int(value_month),
            }


def _previous_month() -> tuple[int, int]:
    """Return (year, month) for the previous month relative to today."""
    today = datetime.date.today()
    first = today.replace(day=1)
    prev = first - datetime.timedelta(days=1)
    return prev.year, prev.month


app = typer.Typer(help="Download KBA FZ10 reports and insert into a SQLite database.")


@app.command()
def main(
    all_months: bool = typer.Option(
        False, "--all", help="Fetch and insert all months of the year up to previous month."
    ),
    year: int | None = typer.Option(
        None, "-y", "--year", help="Year of report, defaults to previous month or year option if --all."
    ),
    month: int | None = typer.Option(
        None, "-m", "--month", help="Month of report (1-12), ignored if --all."
    ),
    db_path: Path = typer.Option(
        Path("data.db"), "-d", "--db-path", help="SQLite database file."
    ),
):
    """
    Download, parse and insert KBA FZ10 report(s) into a SQLite table.
    Use --all to insert every month of the given or current year up to the previous month.
    """
    # determine target table
    db = Database(db_path)
    table_name = "fz10"
    # decide on single vs batch mode
    if all_months:
        # fetch year/month span
        default_year, last_month = _previous_month()
        target_year = year or default_year
        months = list(range(1, last_month + 1))
    else:
        # single-month mode
        target_year, last_month = _previous_month()
        target_year = year or target_year
        target_month = month or last_month
        months = [target_month]

    total = 0
    for m in months:
        blob = io.BytesIO(download(target_year, m))
        rows = list(parse_xslx(blob))
        if not rows:
            typer.echo(f"No data for {target_year}-{m:02d}, skipping.", err=True)
            continue
        db[table_name].insert_all(rows, pk=None, replace=True)
        total += len(rows)
    typer.echo(f"Inserted {total} rows into {db_path}:{table_name}")


if __name__ == "__main__":
    app()
