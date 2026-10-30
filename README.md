# jobwatch

A job-market watcher that polls company career boards, scores each posting against a rule-based candidate
profile, and maintains an Excel tracker you can edit by hand. One Python file, standard library only, no
API keys. Runs twice a day from launchd on a Mac and has kept two separate searches current since October 2026:
one for silicon performance roles, one for agentic / forward-deployed AI roles.

```
python3 jobwatch.py                       # sweep every configured board, update tracker.json / tracker.xlsx / pipeline.md
python3 jobwatch.py --dir ../agentic      # same engine, a second project folder with its own config and tracker
python3 jobwatch.py --add <url> [...]     # import postings by URL (LinkedIn public view, Greenhouse-backed pages, JSON-LD sites)
python3 jobwatch.py --add-manual "Co" "Title" "Location" <url> text.txt   # sites that block fetching
python3 jobwatch.py --rescore             # refetch every tracked posting and re-apply the fit rules (network)
python3 jobwatch.py --rescore-offline     # re-apply the fit rules from stored posting text (seconds, no network)
python3 jobwatch.py --prune               # apply title/location rules and hard-drop flags offline
python3 jobwatch.py --status <id> Applied --note "..."
```

## What it does

1. **Fetch.** Thirteen hand-written fetchers cover the common applicant-tracking systems: Greenhouse, Ashby, Lever,
   SmartRecruiters and Workday (CXS) through their REST APIs directly (GET and POST, pagination, no client libraries),
   plus Apple, Radancy, Teamtailor, BambooHR, Amazon, sitemap + JSON-LD job pages, server-rendered Google results and
   public LinkedIn job views. Transport is the system `curl` so the tool works behind corporate TLS proxies where
   Python's `urllib` does not. Each source has its own timeout and its errors are isolated: one dead board is reported
   in the run summary and never fails the sweep. 500+ boards across two projects as of October 2026.
2. **Filter.** Title include / exclude / combo rules and a location list, all in `jobwatch_config.json`.
3. **Score.** A rule-based fit: title weights (longest-match, no double counting), capped positive text signals,
   regex negatives that become visible *flags* ("8+ years", "clearance", "sales OTE", "language requirement"),
   and context exceptions so a flag is suppressed when the posting supplies the mitigating context. Density rules
   score how much of the description is about the work you want, with title-based exemptions. `fit.drop_flags` names
   the flags that remove an untouched row outright; dropped rows go to `dropped.json`, reversibly. Compensation is
   extracted from the text and normalised to USD with a fixed FX table so London, Zurich and Auckland compare.
   With `fit.save_jd_text` the posting text of every kept row is stored in `jd_store.json`, so rule changes can be
   re-applied offline in seconds.
4. **Track.** `tracker.json` is the source of truth. `tracker.xlsx` is written from it with a dependency-free OOXML
   writer: a Summary sheet plus one sheet per company or per country, rows sorted by fit then package, status-coloured
   rows, hyperlinks, and frozen headers. Before every rewrite the tool reads your Status and My-notes edits back from
   every sheet and compares them to what it last wrote, so your edits and the tool's never overwrite each other.
   Writes are guarded: the workbook is never rewritten while Excel holds it open, the previous ten copies are kept
   under `backups/`, files are written to a temp path and atomically replaced, and the row count is verified after
   each write. Edits are imported before the sweep and again right before the final write, so a network failure
   mid-run cannot lose a status you typed.
5. **Report.** `pipeline.md` lists what you have marked to act on; `report.md` and `latest.json` describe the last run.
   macOS notifications on new matches.

## Layout

```
jobwatch.py                  engine (1,434 lines, stdlib)
daily.sh                     sweep wrapper used by launchd
com.example.jobwatch.plist   launchd job at 08:30 and 18:30
examples/silicon/            config for a silicon-performance search (company sheets, fit-only sort)
examples/agentic/            config for an agentic-AI search (country sheets, fit-then-comp sort, Comp USD column)
ARCHITECTURE.md              design notes and the decisions behind them
```

A project is a folder holding `jobwatch_config.json`; the engine resolves it from `--dir` or `JOBWATCH_DIR` and keeps
every state file (`tracker.json`, `tracker.xlsx`, `xlsx_state.json`, `jd_store.json`, `dropped.json`, `pipeline.md`,
`report.md`, `latest.json`, `backups/`) inside it.

## Setup

```
mkdir mysearch && cp examples/agentic/jobwatch_config.json mysearch/
# edit the config: sources, title rules, locations, fit weights
python3 jobwatch.py --dir mysearch
```

For scheduled runs, edit the paths in `com.example.jobwatch.plist` and `daily.sh`, then:

```
cp com.example.jobwatch.plist ~/Library/LaunchAgents/
launchctl load ~/Library/LaunchAgents/com.example.jobwatch.plist
```

On macOS, give `/bin/bash` Full Disk Access if the job needs to write under Desktop or Documents.

## Status

Used daily by its author since October 2026 on two searches (silicon performance; agentic AI and automation roles),
scheduled twice a day. No automated test suite yet: correctness is checked by the live sweeps, the per-run source-error
report and the write-time verification above.
Known gaps: Microsoft and Meta boards need a browser session; AMD's iCIMS board is JavaScript-rendered; Qualcomm's
Eightfold API is closed (the sitemap route works instead).

MIT licence.
