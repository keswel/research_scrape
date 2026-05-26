# research_scrape

## For the end user (no setup)

Hand them just **`noi_collector.exe`** — it bundles Python, every
dependency, `COLLEGE_DATA.csv`, and `get-session.ps1` into one file.
They double-click it; nothing to install. Usage instructions live in
`dist/READ ME FIRST.txt` (copy it next to the exe when you hand it off).

## Building the exe (developer)

```
pip install -r requirements.txt
pip install pyinstaller
pyinstaller --onefile --console --name noi_collector ^
  --icon noi_collector.ico ^
  --add-data "COLLEGE_DATA.csv;." ^
  --add-data "get-session.ps1;." ^
  main.py
```

The result is `dist/noi_collector.exe`. Rebuild after any code change.
`COLLEGE_DATA.csv` is gitignored, so it must be present locally when you
build (the build is what gets it to the user).

## Running from source (developer)

1. `pip install -r requirements.txt`
2. `python ./main.py`
3. enter credentials
4. search for PID
