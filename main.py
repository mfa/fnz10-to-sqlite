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
    - one column per subheader from row 9 (e.g. Juni 2025, Jan. - Juni 2025, Anteil in %)
    Skips summary rows where modellreihe == 'ZUSAMMEN'.
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
    for row in rows:
        if not any(cell is not None for cell in row):
            continue
        # fill down Marke
        if row[idx_marke] is not None:
            last_marke = row[idx_marke]
        marke = last_marke
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
            # detect range header e.g. 'Jan. - Juni 2025' (overrides year)
            month_range = None
            m2 = re.match(r"(.+?)\s*-\s*(.+?)\s+(\d{4})", str(raw_range))
            if m2:
                month_range = f"{m2.group(1).strip()}-{m2.group(2).strip()}"
                year = int(m2.group(3))
            yield {
                "marke": marke,
                "modellreihe": modell,
                "kategorie": cat,
                "year": year,
                "month": month_label,
                "month_range": month_range,
                "count": value_month,
                "count_range": value_range,
            }


def _previous_month() -> tuple[int, int]:
    """Return (year, month) for the previous month relative to today."""
    today = datetime.date.today()
    first = today.replace(day=1)
    prev = first - datetime.timedelta(days=1)
    return prev.year, prev.month


app = typer.Typer(help="Download KBA FZ10 Excel reports and insert into a SQLite database.")


@app.command()
def main(
    year: int | None = typer.Option(
        None, "-y", "--year", help="Year of report, defaults to previous month."
    ),
    month: int | None = typer.Option(
        None, "-m", "--month", help="Month of report (1-12), defaults to previous month."
    ),
    db_path: Path = typer.Option(
        Path("data.db"), "-d", "--db-path", help="SQLite database file."
    ),
):
    """
    Download, parse and insert KBA FZ10 report into a SQLite table.
    """
    if year is None or month is None:
        year, month = _previous_month()
    blob = io.BytesIO(download(year, month))
    rows = list(parse_xslx(blob))
    if not rows:
        typer.echo("No data rows found, aborting.", err=True)
        raise typer.Exit(1)
    db = Database(db_path)
    # use a single table for all months
    table_name = "fz10"
    db[table_name].insert_all(rows, pk=None, replace=True)
    typer.echo(f"Inserted {len(rows)} rows into {db_path}:{table_name}")


if __name__ == "__main__":
    app()
