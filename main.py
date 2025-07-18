import io
from pathlib import Path

import diskcache
import httpx
import openpyxl

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
    workbook = openpyxl.load_workbook(blob, read_only=True)
    sheet = workbook.get_sheet_by_name("FZ 10.1")
    for row in sheet.iter_rows(values_only=True):
        print(row)


def main():
    # FIXME: get year/month for previous month
    blob = io.BytesIO(download(2025, 6))
    parse_xslx(blob)


if __name__ == "__main__":
    main()
