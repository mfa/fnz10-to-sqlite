## kfz-de-to-sqlite

Convert the official Excel export of German Fahrzeugzulassungen (newly registered cars) to SQLite.

### design decisions

- cache the downloads to not hit the server when running again
- the cache is diskcache -- which uses SQLite too, maybe a bit too much for such a simple usecase
- keep all the German names for fields; there is no official translated version

### usage

``` shell
uv run main.py
```
