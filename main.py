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
    Parse the downloaded Excel blob and yield rows as dicts, using rows 8 and 9 as headers.
    """
    wb = openpyxl.load_workbook(blob, read_only=True, data_only=True)
    sheet = wb["FZ 10.1"]
    rows = sheet.iter_rows(values_only=True)
    # skip to header rows (Excel is 1-indexed; header rows are 8 and 9)
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
    # combine with row 9 to form field names

    columns = []
    for h8, h9 in zip(filled8, header9):
        label = None
        if h8 is not None and h9 is not None:
            label = f"{h8}_{h9}"
        else:
            label = h8 or h9
        # normalize to snake_case
        key = re.sub(r"[^0-9a-zA-Z]+", "_", str(label).strip()).lower().strip("_")
        columns.append(key)
    # yield remaining data rows, carrying forward empty Marke values
    last_marke = None
    for row in rows:
        # skip entirely empty rows
        if not any(cell is not None for cell in row):
            continue
        record = dict(zip(columns, row))
        # fill down missing Marke values
        marque = record.get("marke")
        if marque is None:
            record["marke"] = last_marke
        else:
            last_marke = marque
        yield record


def _previous_month() -> tuple[int, int]:
    """Return (year, month) for the previous month relative to today."""
    today = datetime.date.today()
    first = today.replace(day=1)
    prev = first - datetime.timedelta(days=1)
    return prev.year, prev.month


app = typer.Typer(help="Download KBA FZ10 Excel reports and insert into a SQLite database.")


@app.command()
def main(
    db_path: Path = typer.Argument(Path("data.db"), help="SQLite database file."),
    year: int | None = typer.Option(None, "-y", "--year", help="Year of report, defaults to previous month."),
    month: int | None = typer.Option(None, "-m", "--month", help="Month of report (1-12), defaults to previous month."),
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
    table_name = f"fz10_{year}_{month:02d}"
    db[table_name].insert_all(rows, pk=None, replace=True)
    typer.echo(f"Inserted {len(rows)} rows into {db_path}:{table_name}")


if __name__ == "__main__":
    app()
