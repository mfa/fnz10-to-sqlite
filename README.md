## kfz-de-to-sqlite

Convert the official Excel export of German Fahrzeugzulassungen (newly registered cars) to SQLite.

### design decisions

- cache the downloads to not hit the server when running again
- the cache is diskcache -- which uses SQLite too, maybe a bit too much for such a simple usecase
- keep all the German names for fields; there is no official translated version

### Usage

```shell
```shell
# load previous month's report into the shared 'fz10' table in data.db
uv run main.py --year 2025 --month 6

# append a different month into the same 'fz10' table in a custom DB file
uv run main.py --year 2025 --month 5 --db-path my.db

# fetch all months for current year up to last month
uv run main.py --all
```
```
