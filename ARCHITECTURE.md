# Architecture and decisions

## Shape

```
config ─► fetchers ─► filter ─► detail fetch ─► score ─► tracker.json ─► xlsx / pipeline.md / report.md
                                                              ▲
                                                 read-back of user edits from tracker.xlsx
```

One process, one file, no daemon. State lives in JSON next to the config. The spreadsheet is a view that the user is
allowed to edit in two columns; everything else is regenerated.

## Decisions

**curl as transport.** `urllib` fails TLS verification behind the author's corporate proxy and installing packages
was blocked. Shelling out to `curl` costs nothing and works on every Mac. Each call has an explicit timeout, and
sources can override it (Palantir's Lever feed is 6 MB).

**Rules, not a model, for fit.** A model would be opaque about *why* a posting scored well, and the author needs to
argue with the score. Every point is explained in the Fit-notes cell ("+3 title 'forward deployed'; -2 language
requirement"). Negative rules become flags that stay visible in the workbook, so the reason for a low score is never hidden. Context exceptions (`text_negative_unless`) stop a flag when the posting itself supplies the mitigation, for
example a pay-transparency boilerplate that mentions commission "for sales roles".

**Longest-match title weights.** "applied ai" inside "applied ai engineer" used to count twice. Title keys are sorted
by weight then length and a key is skipped when it is contained in, or contains, one already counted.

**Comp normalisation.** The upper bound of a posted range, converted with a fixed FX table and rounded to the nearest
thousand. Bare `$` is resolved from the location for NZ, AU, SG and CA. Monthly figures are annualised. Precision is
not the point; ordering within a star band is.

**Workbook grouping is configuration.** `workbook.group_by` is `company` or `country`; `workbook.sort` is `fit` or
`fit_comp`. A posting that names several countries is written to each country sheet. Country detection is a keyword
table with buckets for "Remote / Europe-wide" and "Other".

**Read-back by header, not by column index.** The two projects have different column sets. The reader locates
ID / Status / My-notes by header text on every sheet and skips sheets that lack them. It compares each cell against
`xlsx_state.json`, which records what the tool last wrote, so programmatic status changes are not mistaken for user
edits and user edits are never overwritten.

**Dependency-free OOXML.** The .xlsx writer uses inline strings, a small style table, and relationship hyperlinks.
Writing it by hand was cheaper than fighting the package-install block and keeps the tool a single file.

**Imports for sites without feeds.** `--add` handles LinkedIn public job views, JSON-LD job pages, and JavaScript
career sites that are Greenhouse-backed (a `gh_jid` in the URL is resolved against every configured Greenhouse board).
`--add-manual` takes the fields plus a text file for sites behind bot protection.

**Scheduling.** launchd, twice a day. Offline runs are harmless: every source reports a DNS timeout, nothing is
scanned, nothing is pruned, and the log says so.

## Things that did not work

- Qualcomm's Eightfold API returns 403 for unauthenticated clients; the sitemap plus per-page JSON-LD does the job.
- Workday GET probes always return 400, so site names can only be discovered by POSTing to the CXS endpoint.
- Google's careers API v3 is retired; the server-rendered results page parses fine off VPN.
- Microsoft's job APIs refuse connections from the author's network on and off VPN.
- Meta requires a browser session. LinkedIn search scraping is against the terms of service, so the tool only accepts
  pasted public job-view URLs.

## What a second project looks like

`examples/agentic/jobwatch_config.json` is a complete second search: different sources (Ashby and Lever heavy), a
different title vocabulary, country sheets sorted by fit then USD package, and flags tuned to that market (language
requirements, clearance, sales-weighted compensation, "no sponsorship"). Nothing in the engine knows which search it
is running.
