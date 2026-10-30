#!/usr/bin/env python3
"""jobwatch: poll company job boards, filter by title and location, score fit against a candidate profile (rules in jobwatch_config.json),
track everything in an Excel workbook, and notify on new matches.

Usage:
  python3 jobwatch.py                 # fetch, score new matches, update tracker.xlsx, notify
  python3 jobwatch.py --all           # print every tracked match, not only new
  python3 jobwatch.py --dry           # fetch and print, write nothing
  python3 jobwatch.py --only apple,arm
  python3 jobwatch.py --probe         # check each source responds
  python3 jobwatch.py --rescore       # re-fetch details and re-score every tracked posting
  python3 jobwatch.py --status <id-or-url-fragment> "Move ahead" [--note "..."]
  python3 jobwatch.py --pipeline      # print pipeline.md (postings marked Move ahead / Applied / Interview)

Files (all next to this script):
  jobwatch_config.json  sources, queries, include/exclude, locations, fit rules
  tracker.json          ledger of every match ever seen, with fit, compensation, status, notes
  tracker.xlsx          the same ledger as a workbook; edit Status and My notes there, they are read back
  pipeline.md           postings you marked Move ahead / Applied / Interview; Claude reads this at session start
  report.md, latest.json, log.txt
Stdlib only. Transport is the system curl (trusts the corporate proxy certificate).
"""
import argparse
import datetime as dt
import html as htmllib
import json
import os
import re
import subprocess
import sys
import urllib.error
import urllib.parse
import urllib.request
import zipfile
from xml.sax.saxutils import escape

def _project_dir():
    """Folder holding config/tracker/workbook. Default: next to this script. Override with --dir or JOBWATCH_DIR so
    several independent searches (e.g. jobwatch = silicon, agentwatch = agentic delivery) share one engine."""
    for i, a in enumerate(sys.argv):
        if a == "--dir" and i + 1 < len(sys.argv):
            return os.path.abspath(sys.argv[i + 1])
        if a.startswith("--dir="):
            return os.path.abspath(a[len("--dir="):])
    return os.environ.get("JOBWATCH_DIR") or os.path.dirname(os.path.abspath(__file__))


HERE = _project_dir()
P = lambda name: os.path.join(HERE, name)  # noqa: E731
CONFIG, REPORT, LATEST = P("jobwatch_config.json"), P("report.md"), P("latest.json")
TRACKER, XLSX, PIPELINE = P("tracker.json"), P("tracker.xlsx"), P("pipeline.md")
XLSX_STATE = P("xlsx_state.json")
JD_STORE = P("jd_store.json")   # id -> JD text for tracked rows (config fit.save_jd_text); feeds collateral requirement-matching
DROPPED = P("dropped.json")   # archive of rows removed by fit.drop_flags (reversible; never contains Applied/Rejected/Skip rows)
UA = ("Mozilla/5.0 (Macintosh; Intel Mac OS X 14_0) AppleWebKit/537.36 "
      "(KHTML, like Gecko) Chrome/124.0 Safari/537.36")
USE_CURL = os.environ.get("JOBWATCH_TRANSPORT", "curl") == "curl"
STATUSES = ["New", "Interested", "Move ahead", "Applied", "Interview", "Offer", "Rejected", "Skip"]
PIPELINE_STATUSES = {"Move ahead", "Applied", "Interview", "Offer"}


# ----------------------------------------------------------------------------- http
def http(url, method="GET", body=None, headers=None, timeout=30, raw=False):
    h = {"User-Agent": UA, "Accept": "text/html,application/xhtml+xml,*/*;q=0.8" if raw else "application/json, text/plain, */*",
         "Accept-Language": "en-US,en;q=0.9"}
    data = None
    if body is not None:
        h["Content-Type"] = "application/json"
        data = json.dumps(body)
    if headers:
        h.update(headers)
    if USE_CURL:
        cmd = ["curl", "-sS", "-L", "-m", str(timeout), "-X", method, url]
        for k, v in h.items():
            cmd += ["-H", f"{k}: {v}"]
        if data is not None:
            cmd += ["--data", data]
        cmd += ["-w", "\n%{http_code}"]
        p = subprocess.run(cmd, capture_output=True, text=True)
        if p.returncode != 0:
            raise RuntimeError(f"curl failed: {p.stderr.strip()[:160]}")
        text, _, code = p.stdout.rpartition("\n")
        if not code.startswith("2"):
            raise RuntimeError(f"HTTP {code} ({len(text)} bytes)")
    else:
        req = urllib.request.Request(url, data=data.encode() if data else None, method=method, headers=h)
        with urllib.request.urlopen(req, timeout=timeout) as r:
            text = r.read().decode("utf-8", "ignore")
    if raw:
        return text
    try:
        return json.loads(text)
    except json.JSONDecodeError:
        raise RuntimeError(f"non-JSON response ({len(text)} bytes): {text[:80]!r}")


def strip_html(s):
    s = re.sub(r"<(script|style)[^>]*>.*?</\1>", " ", s or "", flags=re.S | re.I)
    s = re.sub(r"<br\s*/?>|</p>|</li>|</div>|</h\d>", "\n", s, flags=re.I)
    s = re.sub(r"<[^>]+>", " ", s)
    s = htmllib.unescape(s)
    return re.sub(r"[ \t\xa0]+", " ", s).strip()


# ----------------------------------------------------------------------------- list fetchers
# Each returns a list of {id, title, location, url, posted, detail?: {...}}
def fetch_greenhouse(src):
    data = http(f"https://boards-api.greenhouse.io/v1/boards/{src['board']}/jobs?content=false", timeout=src.get("timeout", 30))
    return [{"id": f"gh-{src['board']}-{j['id']}", "title": j.get("title", ""),
             "location": (j.get("location") or {}).get("name", ""), "url": j.get("absolute_url", ""),
             "posted": (j.get("updated_at") or "")[:10], "_gh_id": j["id"]} for j in data.get("jobs", [])]


def fetch_ashby(src):
    data = http(f"https://api.ashbyhq.com/posting-api/job-board/{src['board']}?includeCompensation=true", timeout=src.get("timeout", 30))
    out = []
    for j in data.get("jobs", []):
        if j.get("isListed") is False:
            continue
        loc = j.get("location", "") or ""
        extra = [a.get("location", "") for a in (j.get("secondaryLocations") or [])]
        if extra:
            loc = "; ".join([loc] + [e for e in extra if e])
        comp = ((j.get("compensation") or {}).get("compensationTierSummary")
                or (j.get("compensation") or {}).get("scrapeableCompensationSalarySummary") or "")
        out.append({"id": f"ashby-{src['board']}-{j['id']}", "title": j.get("title", ""), "location": loc,
                    "url": j.get("jobUrl", ""), "posted": (j.get("publishedAt") or "")[:10],
                    "detail": {"text": j.get("descriptionPlain") or strip_html(j.get("descriptionHtml", "")),
                               "comp": comp}})
    return out


def fetch_lever(src):
    data = http(f"https://api.lever.co/v0/postings/{src['board']}?mode=json", timeout=src.get("timeout", 30))
    out = []
    for j in data:
        ts = j.get("createdAt")
        posted = dt.datetime.utcfromtimestamp(ts / 1000).date().isoformat() if ts else ""
        out.append({"id": f"lever-{src['board']}-{j['id']}", "title": j.get("text", ""),
                    "location": (j.get("categories") or {}).get("location", "") or "",
                    "url": j.get("hostedUrl", ""), "posted": posted,
                    "detail": {"text": strip_html(j.get("descriptionPlain") or j.get("description", "")) + "\n"
                               + "\n".join(f"{L.get('text', '')}\n{strip_html(L.get('content', ''))}" for L in (j.get("lists") or []))
                               + "\n" + strip_html(j.get("additionalPlain") or j.get("additional", "")), "comp": ""}})
    return out


def fetch_workday(src, queries):
    base = f"https://{src['host']}/wday/cxs/{src['tenant']}/{src['site']}/jobs"
    public = f"https://{src['host']}/{src['site']}"
    out, seen = [], set()
    for q in (src.get("queries") or queries):
        offset = 0
        while True:
            data = http(base, "POST", {"appliedFacets": {}, "limit": 20, "offset": offset, "searchText": q},
                        headers={"Referer": public, "Origin": f"https://{src['host']}"})
            posts = data.get("jobPostings", [])
            for j in posts:
                path = j.get("externalPath", "")
                jid = f"wd-{src['tenant']}-{path.rsplit('/', 1)[-1]}"
                if jid in seen:
                    continue
                seen.add(jid)
                loc = j.get("locationsText", "")
                if not loc and path.count("/") >= 3:          # Accenture-style sites leave locationsText empty; the city is in the path
                    loc = path.split("/")[2].replace("-", " ")
                out.append({"id": jid, "title": j.get("title", "").strip("* "), "location": loc,
                            "url": public + path, "posted": j.get("postedOn", ""),
                            "_wd_detail": f"https://{src['host']}/wday/cxs/{src['tenant']}/{src['site']}{path}"})
            offset += 20
            if len(posts) < 20 or offset >= src.get("max", 100):
                break
    return out


def fetch_apple(src, queries):
    out, seen = [], set()
    for q in queries:
        phrase = f'"{q}"' if " " in q and not q.startswith('"') else q
        page = 1
        while True:
            body = {"query": phrase, "filters": {"postingpostLocation": src.get("locations", [])}, "page": page,
                    "locale": "en-us", "sort": "relevance",
                    "format": {"longDate": "MMMM D, YYYY", "mediumDate": "MMM D, YYYY"}}
            data = http("https://jobs.apple.com/api/v1/search", "POST", body,
                        headers={"Referer": "https://jobs.apple.com/en-us/search", "Origin": "https://jobs.apple.com"})
            res = data.get("res", data)
            total = res.get("totalRecords") or 0
            hits = res.get("searchResults", [])
            for j in hits:
                pid = j.get("positionId") or j.get("id", "").split("-")[-1]
                jid = f"apple-{pid}"
                if jid in seen:
                    continue
                seen.add(jid)
                locs = "; ".join(l.get("name", "") for l in (j.get("locations") or []))
                slug = j.get("transformedPostingTitle") or re.sub(r"[^a-z0-9]+", "-", (j.get("postingTitle") or "").lower())
                out.append({"id": jid, "title": j.get("postingTitle", ""), "location": locs,
                            "url": f"https://jobs.apple.com/en-us/details/{pid}/{slug}",
                            "posted": (j.get("postingDate") or "")[:12], "_apple_pid": pid})
            page += 1
            if not hits or page * 20 > min(total, src.get("max", 200)):
                break
    return out


def fetch_radancy(src, queries):
    out, seen = [], set()
    base = src["base"].rstrip("/")
    anchor_re = re.compile(r'<a[^>]+href="([^"]+)"[^>]*data-job-id="([^"]+)"[^>]*>(.*?)</a>', re.S)
    for q in queries:
        params = {"CurrentPage": 1, "RecordsPerPage": src.get("per_page", 50), "Keywords": q,
                  "SearchResultsModuleName": "Search Results", "SearchFiltersModuleName": "Search Filters",
                  "SortCriteria": 5, "SortDirection": 1}
        data = http(f"{base}/search-jobs/results?" + urllib.parse.urlencode(params))
        blob = data.get("results", "")
        matches = list(anchor_re.finditer(blob))
        for k, m in enumerate(matches):
            href, jid, inner = m.group(1), m.group(2), m.group(3)
            key = f"radancy-{src['name'].lower()}-{jid}"
            if key in seen:
                continue
            seen.add(key)
            title = strip_html(inner).replace("\n", " ").strip()
            chunk = blob[m.end(): matches[k + 1].start() if k + 1 < len(matches) else m.end() + 1500]
            lm = re.search(r'job-location[^>]*>(.*?)</', chunk, re.S)
            if lm:
                loc = strip_html(lm.group(1))
            else:
                parts = href.strip("/").split("/")
                loc = parts[1].replace("-", " ").title() if len(parts) > 2 and parts[0] == "job" else ""
            out.append({"id": key, "title": title, "location": loc, "url": base + href, "posted": ""})
    return out


def fetch_google(src, queries):  # retired API, kept for reference; disabled in config
    out = []
    for q in queries:
        data = http("https://careers.google.com/api/v3/search/?" + urllib.parse.urlencode({"q": q}))
        for j in data.get("jobs", []):
            out.append({"id": f"google-{j.get('id', '').rsplit('/', 1)[-1]}", "title": j.get("title", ""),
                        "location": "; ".join(l.get("display", "") for l in (j.get("locations") or [])),
                        "url": j.get("apply_url", ""), "posted": (j.get("publish_date") or "")[:10]})
    return out



def fetch_teamtailor(src, queries):
    """Teamtailor career sites (e.g. careers.huaweirc.ch): parse the /jobs listing HTML. One page, no API needed."""
    base = src["base"].rstrip("/")
    out, seen = [], set()
    pages = [f"{base}/jobs"]  # Teamtailor lists every opening on one page; the site can be slow, so one request only
    for url in pages:
        page = http(url, raw=True, timeout=src.get("timeout", 60))
        for li in re.findall(r"<li[^>]*>(.*?)</li>", page, re.S):
            m = re.search(r'href="((?:https?://[^"/]+)?/jobs/(\d+)-[^"]+)"[^>]*>(.*?)</a>', li, re.S)
            if not m:
                continue
            href, jid, inner = m.group(1), m.group(2), m.group(3)
            key = f"tt-{src['name'].lower().replace(' ', '')}-{jid}"
            if key in seen:
                continue
            seen.add(key)
            title = re.split(r"\s{2,}", strip_html(inner).replace("\n", "  ").strip())[0].strip()  # drop department/location suffix
            spans = [strip_html(x) for x in re.findall(r"<span[^>]*>(.*?)</span>", li[m.end():], re.S)]
            spans = [x for x in spans if x and x not in ("·", "&middot;")]
            loc = spans[-1] if spans else src.get("default_location", "")
            if not href.startswith("http"):
                href = base + href
            out.append({"id": key, "title": title, "location": loc, "url": href, "posted": "", "_html_detail": True})
    return out


def fetch_bamboohr(src):
    """BambooHR hosted careers (e.g. codasip.bamboohr.com): /careers/list is public JSON."""
    data = http(f"https://{src['host']}/careers/list")
    out = []
    for j in data.get("result", []):
        loc = j.get("location") or {}
        loc_s = ", ".join(x for x in [loc.get("city", ""), loc.get("state", ""), loc.get("country", "")] if x) if isinstance(loc, dict) else str(loc)
        out.append({"id": f"bamboo-{src['host'].split('.')[0]}-{j.get('id')}", "title": j.get("jobOpeningName", ""),
                    "location": loc_s, "url": f"https://{src['host']}/careers/{j.get('id')}", "posted": "",
                    "_bamboo": (src["host"], j.get("id"))})
    return out

def fetch_amazon(src, queries):
    """amazon.jobs public search JSON. Descriptions and qualifications come inline, so no detail call is needed."""
    out, seen = [], set()
    for q in queries:
        params = {"base_query": q, "loc_query": src.get("loc_query", ""), "offset": 0,
                  "result_limit": src.get("limit", 100), "sort": "recent"}
        data = http("https://www.amazon.jobs/en/search.json?" + urllib.parse.urlencode(params))
        for j in data.get("jobs", []):
            jid = f"amazon-{j.get('id_icims') or j.get('id')}"
            if jid in seen:
                continue
            seen.add(jid)
            text = "\n".join(strip_html(j.get(k, "") or "") for k in ("description", "basic_qualifications", "preferred_qualifications"))
            out.append({"id": jid, "title": j.get("title", ""), "location": j.get("normalized_location") or j.get("location", ""),
                        "url": "https://www.amazon.jobs" + (j.get("job_path") or ""), "posted": j.get("posted_date", ""),
                        "detail": {"text": text, "comp": find_comp(text)}})
    return out


def fetch_sitemap(src, queries):
    """Career sites whose API is closed but whose sitemap lists every job URL with the title in the slug (e.g. Qualcomm on Eightfold).
    Title and location are read from the slug for matching; the job page's JSON-LD fills in exact fields at detail time."""
    xml = http(src["sitemap"], raw=True, timeout=src.get("timeout", 60))
    out = []
    for loc in re.findall(r"<loc>(.*?)</loc>", xml):
        m = re.search(src.get("pattern", r"/job/(\d+)-([a-z0-9-]+)"), loc)
        if not m:
            continue
        jid, slug = m.group(1), m.group(2)
        words = slug.replace("-", " ")
        out.append({"id": f"sitemap-{src['name'].lower()}-{jid}", "title": words, "location": words,
                    "url": loc.split("?")[0] if src.get("strip_query") else loc, "posted": "", "_jsonld": True})
    return out


def fetch_google_html(src, queries):
    """google.com/about/careers results pages render server-side (off corporate VPN), 20 cards per page.
    Each card ends with a 'Learn more about <title>' link; the location sits between a 'place' icon and a 'bar_chart' icon before it."""
    out, seen = [], set()
    base = "https://www.google.com/about/careers/applications/jobs/results"
    link_re = re.compile(r'<a[^>]+href="(?:https://www\.google\.com/about/careers/applications/)?jobs/results/(\d{10,})-([a-z0-9-]+)[^"]*"[^>]*aria-label="Learn more about ([^"]+)"', re.S)
    # Google returns an empty page when several location filters are combined, so query one location at a time.
    for q in src.get("queries") or queries:
        for L in (src.get("locations") or [None]) + ([None] if src.get("also_global", True) else []):
            for page in range(1, src.get("pages", 2) + 1):
                params = [("q", q), ("page", page)] + ([("location", L)] if L else [])
                html_page = http(base + "?" + urllib.parse.urlencode(params), raw=True)
                ms = list(link_re.finditer(html_page))
                if not ms:
                    break
                new_here = 0
                for k, m in enumerate(ms):
                    jid, slug, title = m.group(1), m.group(2), htmllib.unescape(m.group(3)).strip()
                    key = f"google-{jid}"
                    if key in seen:
                        continue
                    seen.add(key)
                    prev_end = ms[k - 1].end() if k else max(0, m.start() - 6000)
                    card_text = re.sub(r"\s*\|\s*", "|", re.sub(r"<[^>]+>", "|", html_page[prev_end:m.start()]))
                    lm = re.search(r"\|place\|+(.*?)\|+bar_chart", card_text, re.S)
                    loc = ""
                    if lm:
                        loc = "; ".join(x.strip(" ;") for x in lm.group(1).split("|") if x.strip(" ;"))
                    out.append({"id": key, "title": title, "location": htmllib.unescape(loc), "url": f"{base}/{jid}-{slug}",
                                "posted": "", "_html_detail": True})
    return out


def fetch_smartrecruiters(src):
    """SmartRecruiters public postings API; details fetched lazily via _sr_detail."""
    out, offset = [], 0
    while True:
        data = http(f"https://api.smartrecruiters.com/v1/companies/{src['board']}/postings?limit=100&offset={offset}", timeout=src.get("timeout", 30))
        posts = data.get("content", [])
        for j in posts:
            loc = j.get("location") or {}
            loc_s = ", ".join(x for x in [loc.get("city", ""), loc.get("region", ""), loc.get("country", "").upper()] if x)
            if loc.get("remote"):
                loc_s = (loc_s + "; Remote").strip("; ")
            out.append({"id": f"sr-{src['board']}-{j['id']}", "title": j.get("name", ""), "location": loc_s,
                        "url": f"https://jobs.smartrecruiters.com/{src['board']}/{j['id']}", "posted": (j.get("releasedDate") or "")[:10],
                        "_sr_detail": f"https://api.smartrecruiters.com/v1/companies/{src['board']}/postings/{j['id']}"})
        offset += 100
        if len(posts) < 100 or offset >= src.get("max", 1000):
            break
    return out


FETCHERS = {"greenhouse": fetch_greenhouse, "ashby": fetch_ashby, "lever": fetch_lever, "smartrecruiters": fetch_smartrecruiters,
            "workday": fetch_workday, "apple": fetch_apple, "google": fetch_google, "radancy": fetch_radancy,
            "teamtailor": fetch_teamtailor, "bamboohr": fetch_bamboohr,
            "amazon": fetch_amazon, "sitemap": fetch_sitemap, "google_html": fetch_google_html}
QUERY_TYPES = {"workday", "apple", "google", "radancy", "teamtailor", "amazon", "sitemap", "google_html"}


# ----------------------------------------------------------------------------- detail fetchers
COMP_PATTERNS = [
    r"\$\s?\d{2,3},\d{3}(?:\.\d{2})?\s*(?:-|–|—|to|and)\s*\$?\s?\d{2,3},\d{3}(?:\.\d{2})?",
    r"(?:£|€|\$|CHF|USD|EUR|GBP|NZD|AUD|SGD|AED|CAD|SEK)\s?\d{2,3}(?:\.\d)?\s?K\s*(?:-|–|—|to)\s*(?:£|€|\$|CHF|USD|EUR|GBP|NZD|AUD|SGD|AED|CAD|SEK)?\s?\d{2,3}(?:\.\d)?\s?K",
    r"(?:£|€|NZD|AUD|SGD|AED|CAD|CHF)\s?\d{2,3},\d{3}(?:\.\d{2})?\s*(?:-|–|—|to|and)\s*(?:£|€|NZD|AUD|SGD|AED|CAD|CHF)?\s?\d{2,3},\d{3}(?:\.\d{2})?",
    r"\$\s?\d{2,3}(?:\.\d)?K\s*(?:-|–|—|to)\s*\$?\s?\d{2,3}(?:\.\d)?K",
    r"\d{2,3},\d{3}\s*(?:USD|EUR|CHF|GBP)\s*(?:-|–|—|to)\s*\d{2,3},\d{3}\s*(?:USD|EUR|CHF|GBP)",
    r"(?:USD|EUR|CHF|GBP|£|€)\s?\d{2,3},\d{3}\s*(?:-|–|—|to)\s*(?:USD|EUR|CHF|GBP|£|€)?\s?\d{2,3},\d{3}",
]


def find_comp(text):
    for pat in COMP_PATTERNS:
        m = re.search(pat, text, re.I)
        if m:
            return re.sub(r"\s+", " ", m.group(0)).strip()
    return ""


FX_TO_USD = {"USD": 1.0, "$": 1.0, "EUR": 1.08, "€": 1.08, "GBP": 1.27, "£": 1.27, "CHF": 1.12, "NZD": 0.60, "AUD": 0.65,
             "SGD": 0.74, "AED": 0.272, "CAD": 0.73, "SEK": 0.095, "INR": 0.012, "JPY": 0.0067}
LOC_CCY = [("united kingdom", "GBP"), ("london", "GBP"), ("manchester", "GBP"), ("uk", "GBP"), ("ireland", "EUR"), ("dublin", "EUR"),
           ("germany", "EUR"), ("france", "EUR"), ("spain", "EUR"), ("italy", "EUR"), ("netherlands", "EUR"), ("europe", "EUR"),
           ("switzerland", "CHF"), ("zurich", "CHF"), ("zürich", "CHF"), ("new zealand", "NZD"), ("auckland", "NZD"), ("wellington", "NZD"),
           ("australia", "AUD"), ("sydney", "AUD"), ("melbourne", "AUD"), ("singapore", "SGD"), ("dubai", "AED"), ("abu dhabi", "AED"),
           ("uae", "AED"), ("canada", "CAD"), ("toronto", "CAD"), ("sweden", "SEK"), ("stockholm", "SEK"), ("india", "INR"), ("bangalore", "INR")]


def comp_usd(comp, location=""):
    """Upper end of a salary range converted to USD per year (int), or None. Fixed FX table; good enough for sorting."""
    if not comp:
        return None
    c = comp.replace(",", "")
    m = re.search(r"\b(USD|EUR|GBP|CHF|NZD|AUD|SGD|AED|CAD|SEK|INR|JPY)\b", c, re.I)
    ccy = m.group(1).upper() if m else next((sym for sym in ("£", "€", "$") if sym in c), None)
    if ccy == "$":   # bare dollar: trust the location for NZ/AU/SG/CA, otherwise USD
        loc = (location or "").lower()
        ccy = next((cc for key, cc in LOC_CCY if key in loc and cc in ("NZD", "AUD", "SGD", "CAD")), "USD")
    if not ccy:
        loc = (location or "").lower()
        ccy = next((cc for key, cc in LOC_CCY if key in loc), "USD")
    nums = [(float(v), bool(k)) for v, k in re.findall(r"(\d+(?:\.\d+)?)\s?([kK])?", c)]
    vals = [v * 1000 if k or v < 1000 else v for v, k in nums if (k or v >= 20)]
    vals = [v for v in vals if 15000 <= v <= 2000000]
    if not vals:
        return None
    hi = max(vals)
    if re.search(r"per month|/month|monthly|p\.?m\.?\b", comp, re.I):
        hi *= 12
    return int(round(hi * FX_TO_USD.get(ccy, 1.0) / 1000.0) * 1000)


# ----------------------------------------------------------------------------- country grouping (workbook sheet per country)
COUNTRY_KEYS = [
    ("United Kingdom", ["united kingdom", "london", "manchester", "cambridge", "bristol", "edinburgh", "leeds", "england", "scotland", " uk", "uk,", "(uk", "uk)"]),
    ("Ireland", ["ireland", "dublin", "cork"]),
    ("Germany", ["germany", "munich", "berlin", "hamburg", "cologne", "stuttgart", "frankfurt", "münchen", "deutschland"]),
    ("France", ["france", "paris", "marseille", "lyon", "sophia"]),
    ("Spain", ["spain", "madrid", "barcelona", "españa"]),
    ("Italy", ["italy", "milan", "milano", "rome", "roma", "italia"]),
    ("Switzerland", ["switzerland", "zurich", "zürich", "lausanne", "geneva", "basel", "schweiz"]),
    ("Netherlands", ["netherlands", "amsterdam", "eindhoven", "rotterdam", "the hague"]),
    ("Belgium", ["belgium", "brussels", "leuven", "antwerp"]),
    ("Luxembourg", ["luxembourg"]),
    ("Austria", ["austria", "vienna", "wien", "linz", "graz"]),
    ("Sweden", ["sweden", "stockholm", "lund", "gothenburg"]),
    ("Denmark", ["denmark", "copenhagen"]),
    ("Norway", ["norway", "oslo"]),
    ("Finland", ["finland", "helsinki", "tampere"]),
    ("Poland", ["poland", "warsaw", "krakow", "kraków", "wroclaw", "gdansk"]),
    ("Czechia", ["czech", "prague", "brno"]),
    ("Portugal", ["portugal", "lisbon", "porto"]),
    ("Serbia", ["serbia", "belgrade"]),
    ("Hungary", ["hungary", "budapest"]),
    ("Romania", ["romania", "bucharest", "cluj"]),
    ("Greece", ["greece", "athens"]),
    ("New Zealand", ["new zealand", "auckland", "wellington", "christchurch"]),
    ("Australia", ["australia", "sydney", "melbourne", "brisbane", "perth", "queensland", "canberra"]),
    ("Singapore", ["singapore"]),
    ("Japan", ["japan", "tokyo"]),
    ("Hong Kong", ["hong kong"]),
    ("UAE", ["uae", "united arab emirates", "dubai", "abu dhabi"]),
    ("Qatar", ["qatar", "doha"]),
    ("India", ["india", "bangalore", "bengaluru", "hyderabad", "pune", "noida", "gurgaon", "chennai", "mumbai"]),
    ("Canada", ["canada", "toronto", "vancouver", "montreal", "markham", "ottawa", "waterloo"]),
    ("USA", ["united states", "usa", ", us", "us)", "california", "new york", "seattle", "austin", "texas", "san francisco", "washington",
             "colorado", "boston", "massachusetts", "illinois", "oregon", "virginia", "georgia", "atlanta", "chicago", "santa clara",
             "san diego", "san jose", "cupertino", "sunnyvale", "mountain view", "pennsylvania", "north carolina", "florida", "arizona",
             "nevada", "michigan", "new jersey", "maryland", "connecticut", "delaware", "alabama", "south carolina", "dallas", "phoenix", "portland"]),
    ("China", ["china", "shanghai", "beijing", "shenzhen", "guangdong", "hangzhou"]),
]


EUROPE = {"United Kingdom", "Ireland", "Germany", "France", "Spain", "Italy", "Switzerland", "Netherlands", "Belgium", "Luxembourg",
          "Austria", "Sweden", "Norway", "Denmark", "Finland", "Poland", "Czechia", "Portugal", "Serbia", "Hungary", "Romania", "Greece"}


def countries_of(location):
    loc = " " + (location or "").lower() + " "
    found = [name for name, keys in COUNTRY_KEYS if any(k in loc for k in keys)]
    # a posting open to "Europe" / "EMEA" that names no European city (or only non-European ones) also belongs on the
    # Europe-wide sheet, otherwise remote-Europe roles vanish from every European tab
    if re.search(r"\beurope\b|\bemea\b|europe-wide|remote.{0,12}(eu|europe)", loc) and not any(c in EUROPE for c in found):
        found.append("Remote / Europe-wide")
    if found:
        return found
    if re.search(r"remote|europe|emea|anywhere|global|worldwide", loc):
        return ["Remote / Europe-wide"]
    return ["Other"]


def fetch_detail(job, src):
    """Return {text, comp}. Best effort; failures return empty text and an error note."""
    try:
        if "detail" in job:
            d = job["detail"]
            return {"text": d.get("text", ""), "comp": d.get("comp") or find_comp(d.get("text", ""))}
        if "_apple_pid" in job:
            d = http(f"https://jobs.apple.com/api/v1/jobDetails/PIPE-{job['_apple_pid']}?locale=en-us",
                     headers={"Referer": job["url"]})
            r = d.get("res", d)
            parts = [r.get("jobSummary", ""), r.get("description", ""), r.get("minimumQualifications", ""),
                     r.get("preferredQualifications", ""), r.get("keyQualifications", ""),
                     r.get("educationAndExperience", ""), r.get("additionalRequirements", "")]
            footers = r.get("postingFooters") or []
            if isinstance(footers, list):
                parts += [f.get("text", "") if isinstance(f, dict) else str(f) for f in footers]
            text = "\n".join(strip_html(p) for p in parts if p)
            return {"text": text, "comp": find_comp(text) or find_comp(strip_html(json.dumps(r)))}
        if "_wd_detail" in job:
            d = http(job["_wd_detail"])
            text = strip_html((d.get("jobPostingInfo") or {}).get("jobDescription", ""))
            return {"text": text, "comp": find_comp(text)}
        if "_sr_detail" in job:
            d = http(job["_sr_detail"])
            secs = (d.get("jobAd") or {}).get("sections") or {}
            text = "\n".join(strip_html((secs.get(k) or {}).get("text", "")) for k in ("companyDescription", "jobDescription", "qualifications", "additionalInformation"))
            return {"text": text, "comp": find_comp(text)}
        if "_gh_id" in job:
            d = http(f"https://boards-api.greenhouse.io/v1/boards/{src['board']}/jobs/{job['_gh_id']}")
            text = strip_html(d.get("content", ""))
            return {"text": text, "comp": find_comp(text)}
        if job.get("_jsonld"):
            page = http(job["url"], raw=True, timeout=src.get("timeout", 60))
            for blk in re.findall(r'<script type="application/ld\+json">(.*?)</script>', page, re.S):
                try:
                    d = json.loads(blk)
                except json.JSONDecodeError:
                    continue
                if isinstance(d, dict) and d.get("@type") == "JobPosting":
                    job["title"] = d.get("title") or job["title"]
                    locs = d.get("jobLocation") or []
                    locs = locs if isinstance(locs, list) else [locs]
                    parts = []
                    for L in locs:
                        a = (L or {}).get("address") or {}
                        fields = []
                        for key in ("addressLocality", "addressRegion", "addressCountry"):
                            val = a.get(key, "")
                            if isinstance(val, dict):
                                val = val.get("name", "")
                            if val and val not in fields:
                                fields.append(str(val))
                        parts.append(", ".join(fields))
                    if parts:
                        job["location"] = "; ".join(parts)
                    job["posted"] = (d.get("datePosted") or "")[:10]
                    text = strip_html(d.get("description", ""))
                    sal = d.get("baseSalary") or {}
                    v = sal.get("value") if isinstance(sal, dict) else None
                    comp = ""
                    if isinstance(v, dict) and (v.get("minValue") or v.get("maxValue")):
                        comp = f"{sal.get('currency', '')} {v.get('minValue', '')} - {v.get('maxValue', '')}".strip()
                    return {"text": text, "comp": comp or find_comp(text)}
            flat = strip_html(page)
            return {"text": flat[:20000], "comp": find_comp(flat)}
        if job.get("_html_detail"):
            flat = strip_html(http(job["url"], raw=True, timeout=src.get("timeout", 60)))
            return {"text": flat[:20000], "comp": find_comp(flat)}
        if "_bamboo" in job:
            host, jid = job["_bamboo"]
            d = http(f"https://{host}/careers/{jid}/detail")
            r = d.get("result", d)
            text = strip_html((r.get("jobOpening") or r).get("description", "") if isinstance(r, dict) else "")
            return {"text": text, "comp": find_comp(text)}
        if job["id"].startswith("radancy-"):
            page = http(job["url"], raw=True)
            m = re.search(r'(?is)<div[^>]+class="[^"]*ats-description[^"]*"[^>]*>(.*?)</section>', page)
            if m:
                text = strip_html(m.group(1))
            else:
                flat = strip_html(page)
                m2 = re.search(r"(?is)(Job Overview|Job Description|About the role)(.*?)(Apply Now|Share this job|Related Jobs)", flat)
                text = m2.group(2) if m2 else flat
            return {"text": text[:20000], "comp": find_comp(text)}
    except Exception as e:  # noqa: BLE001 - best effort
        return {"text": "", "comp": "", "error": str(e)[:120]}
    return {"text": "", "comp": ""}


# ----------------------------------------------------------------------------- matching and fit
def norm(s):
    return re.sub(r"\s+", " ", (s or "").lower())


def match(job, cfg):
    t, loc = norm(job["title"]), norm(job["location"])
    hits = [k for k in cfg["title_include"] if k in t]
    for a, b in cfg.get("title_include_combo", []):
        if a in t and (not b or b in t):
            hits.append(a if not b else f"{a}+{b}")
    if not hits:
        return None
    excluded = [k for k in cfg["title_exclude"] if k in t]
    # "Senior/Principal Researcher" style titles span two levels; do not reject them for the word principal
    if any(x in t for x in cfg.get("title_exclude_unless", [])):
        excluded = [k for k in excluded if k not in ("principal", "staff ", "senior staff")]
    # title_rescue: phrases that name the role family we want outright ("AI Operations Engineer", "AI Platform Engineer");
    # a title containing one of them ignores generic excludes such as "operations" or "platform engineer"
    if excluded and any(x in t for x in cfg.get("title_rescue", [])):
        excluded = []
    if excluded:
        return None
    loc_hit = [k for k in cfg["location_include"] if k in loc]
    loc_low = [k for k in cfg.get("location_low_priority", []) if k in loc]
    if cfg.get("require_location", True) and not loc_hit and not loc_low:
        return None
    return {"keywords": hits, "loc_hit": loc_hit, "low": bool(loc_low and not loc_hit)}


def score_fit(job, text, cfg):
    """Rule-based fit against a candidate profile (rules in jobwatch_config.json). Returns (stars 1-5, notes string, flags list)."""
    rules = cfg["fit"]
    t, loc, body = norm(job["title"]), norm(job["location"]), norm(text)
    score, why, flags = 0.0, [], []
    hits = sorted(((k, w) for k, w in rules["title_weights"].items() if k in t), key=lambda kw: (-kw[1], -len(kw[0])))
    counted = []
    for k, w in hits:
        if rules.get("title_overlap", "sum") == "longest" and any(k in c or c in k for c in counted):
            continue   # "applied ai" inside "applied ai engineer": count the stronger key only
        counted.append(k)
        score += w
        why.append(f"+{w:g} title '{k}'")
    if rules.get("include_location", False):   # off by default: fit = JD vs profile only; location is shown, not scored
        for k, w in rules["location_weights"].items():
            if k in loc:
                score += w
                why.append(f"+{w:g} {k}")
                break
    if body:
        pos = 0
        not_title = rules.get("text_positive_not_title", {})   # label -> title regex: positive does not apply to those titles
        for k, (w, label) in rules["text_positive"].items():
            if re.search(k, body):
                if label in not_title and re.search(not_title[label], t):
                    continue   # e.g. "internal customer" wording inside a Forward Deployed JD is not evidence of an internal role
                pos += w
                why.append(f"+{w:g} {label}")
        score += min(pos, rules.get("text_positive_cap", 4))
        unless = rules.get("text_negative_unless", {})
        unless2 = rules.get("text_negative_unless_internal", {})   # label -> regex: second exception, only for non-delivery titles
        igate = rules.get("internal_title_gate")                   # title regex naming delivery-type roles (same shape as company_negative_title)
        for k, (w, label) in rules["text_negative"].items():
            if re.search(k, body):
                if label in unless and re.search(unless[label], body):
                    why.append(f"(0 {label}: context exception, not penalised)")
                    continue
                if label in unless2 and re.search(unless2[label], body) and not (igate and re.search(igate, t)):
                    why.append(f"(0 {label}: internal-customer role, not penalised)")
                    continue
                score -= w
                flags.append(label)
                why.append(f"-{w:g} {label}")
        # density rules: how much of the description is actually about the work we want (config "text_density")
        for d in rules.get("text_density", []):
            n = len(re.findall(d["pattern"], body))
            if "over_at" in d and n >= d["over_at"]:   # too much of something we lack (e.g. a cloud/web delivery stack)
                score -= d["over_penalty"]; flags.append(d["over_label"]); why.append(f"-{d['over_penalty']:g} {d['over_label']} ({n} mentions)")
            if "min" in d and n < d["min"] and not (d.get("exempt_title") and re.search(d["exempt_title"], t)):
                score -= d["penalty"]; flags.append(d["label"]); why.append(f"-{d['penalty']:g} {d['label']} ({n} mentions)")
            elif "bonus_at" in d and n >= d["bonus_at"]:
                score += d["bonus"]; why.append(f"+{d['bonus']:g} {d.get('bonus_label', d['label'])} ({n} mentions)")
            if "against" in d and n < d.get("against_min", 6) and len(re.findall(d["against"], body)) >= d.get("against_at", 4):
                score -= d["against_penalty"]; flags.append(d["against_label"]); why.append(f"-{d['against_penalty']:g} {d['against_label']}")
    else:
        why.append("no description fetched, title-only score")
    tgate = rules.get("company_negative_title")   # only for delivery-type titles; a platform SWE at the same employer is a different job
    for k, (w, label) in rules.get("company_negative", {}).items():   # employer tier, independent of JD wording
        if re.search(k, norm(job.get("company", ""))) and (not tgate or re.search(tgate, norm(job.get("title", "")))):
            score -= w; flags.append(label); why.append(f"-{w:g} {label}")
            break
    stars = 1 + sum(score > th for th in rules["star_thresholds"])
    for flag, cap in rules.get("cap_stars_on_flags", {}).items():   # e.g. a JD with little agent content can never be 4★+
        if flag in flags:
            stars = min(stars, cap)
    return max(1, min(5, stars)), f"score {score:g}: " + "; ".join(why[:9]), flags


def import_url(url, cfg, tracker, today):
    """Fetch one job posting by URL (LinkedIn public job view, Eightfold/JSON-LD sites, or generic HTML), score it,
    and add it to the tracker as New. Returns the tracker entry."""
    title, company, loc, text, comp, posted = "", "", "", "", "", ""
    page = ""
    gh = re.search(r"gh_jid=(\d+)", url) or re.search(r"greenhouse\.io/[^/]+/jobs/(\d+)", url)
    if gh:  # JS-rendered career pages: ask the Greenhouse API directly, trying each configured board token
        srcs = cfg.get("sources", {}); srcs = list(srcs.values()) if isinstance(srcs, dict) else srcs
        for src in srcs:
            if src.get("type") != "greenhouse" or not src.get("board"):
                continue
            try:
                d = json.loads(http(f"https://boards-api.greenhouse.io/v1/boards/{src['board']}/jobs/{gh.group(1)}?questions=false", raw=True, timeout=30))
            except Exception:
                continue
            if d.get("title"):
                title = d["title"]; company = src.get("name") or src.get("company") or src["board"].title()
                loc = (d.get("location") or {}).get("name", ""); text = strip_html(htmllib.unescape(d.get("content", "")))
                posted = (d.get("updated_at") or "")[:10]
                break
    if not title:
        page = http(url, raw=True, timeout=60)
    for blk in re.findall(r'<script type="application/ld\+json"[^>]*>(.*?)</script>', page, re.S):
        try:
            d = json.loads(blk)
        except json.JSONDecodeError:
            continue
        if isinstance(d, dict) and d.get("@type") == "JobPosting":
            title = d.get("title", ""); company = (d.get("hiringOrganization") or {}).get("name", "")
            locs = d.get("jobLocation") or []; locs = locs if isinstance(locs, list) else [locs]
            parts = []
            for L in locs:
                a = (L or {}).get("address") or {}
                fields = []
                for key in ("addressLocality", "addressRegion", "addressCountry"):
                    v = a.get(key, ""); v = v.get("name", "") if isinstance(v, dict) else v
                    if v and v not in fields:
                        fields.append(str(v))
                parts.append(", ".join(fields))
            loc = "; ".join(x for x in parts if x); text = strip_html(d.get("description", "")); posted = (d.get("datePosted") or "")[:10]
            sal = d.get("baseSalary") or {}; v = sal.get("value") if isinstance(sal, dict) else None
            if isinstance(v, dict) and (v.get("minValue") or v.get("maxValue")):
                comp = f"{sal.get('currency', '')} {v.get('minValue', '')} - {v.get('maxValue', '')}".strip()
            break
    if not title and "linkedin.com" in url:
        t = re.search(r"<title>(.*?)</title>", page, re.S)
        tt = htmllib.unescape(t.group(1)).strip() if t else ""
        m = re.match(r"\s*(.*?) hiring (.*?) in (.*?) \| LinkedIn", tt)
        m2 = re.match(r"\s*(.*?) at (.*?) (?:—|-|–) (.*?) \| LinkedIn", tt)
        if m:
            company, title, loc = m.group(1), m.group(2), m.group(3)
        elif m2:
            title, company, loc = m2.group(1), m2.group(2), m2.group(3)
        d = re.search(r'class="[^"]*(?:show-more-less-html__markup|description__text)[^"]*"[^>]*>(.*?)</div>\s*</div>', page, re.S)
        text = strip_html(d.group(1)) if d else strip_html(page)[:12000]
    if not title:
        t = re.search(r"<title>(.*?)</title>", page, re.S)
        title = htmllib.unescape(t.group(1)).strip()[:120] if t else url
        text = text or strip_html(page)[:12000]
    comp = comp or find_comp(text)
    job = {"title": title, "location": loc}
    stars, notes, flags = score_fit(job, text, cfg)
    m = re.search(r"/(?:view|job)/(\d{6,})", url) or re.search(r"(\d{6,})", url)
    key = f"manual-{(company or 'web').lower().split()[0]}-{m.group(1) if m else abs(hash(url)) % 10**8}"
    entry = tracker.get(key, {"found": today, "status": "New", "notes": "added by URL"})
    entry.update({"company": company or "Unknown", "title": title, "location": loc, "posted": posted, "url": url, "source": "manual",
                  "comp": comp, "fit": stars, "fit_notes": notes, "flags": flags, "low_priority_location": False, "last_seen": today})
    tracker[key] = entry
    return key, entry, text


# ----------------------------------------------------------------------------- xlsx
def col_letter(i):
    s = ""
    while i >= 0:
        s = chr(65 + i % 26) + s
        i = i // 26 - 1
    return s


XCOLS = ["Found", "Company", "Title", "Location", "Same JD as", "Posted", "Compensation", "Fit", "Fit notes",
         "Flags", "Status", "My notes", "Link", "Folder", "Source", "ID"]
XWIDTH = [11, 12, 48, 24, 22, 13, 26, 7, 60, 24, 13, 30, 8, 9, 11, 30]
WB = {"group_by": "company", "sort": "fit", "comp_usd": False}   # overridden from config["workbook"] in main()


def configure_workbook(cfg):
    global XCOLS, XWIDTH
    WB.update(cfg.get("workbook", {}))
    if WB.get("comp_usd") and "Comp USD" not in XCOLS:
        i = XCOLS.index("Compensation") + 1
        XCOLS.insert(i, "Comp USD"); XWIDTH.insert(i, 11)
    if WB.get("group_by") == "country" and "Country" not in XCOLS:
        i = XCOLS.index("Location") + 1
        XCOLS.insert(i, "Country"); XWIDTH.insert(i, 14)
        XCOLS.insert(i + 1, "Also on"); XWIDTH.insert(i + 1, 22)   # other country sheets the same posting appears on


def _sheet_name(name, used):
    base = re.sub(r"[\[\]\*\?/\\:]", " ", name).strip()[:28] or "Sheet"
    n, k = base, 2
    while n in used:
        n = f"{base[:25]} {k}"; k += 1
    used.add(n)
    return n


STATUS_STYLE = {"Applied": 4, "Interview": 5, "Offer": 5, "Move ahead": 6, "Interested": 6, "Rejected": 7, "Skip": 7}


def _sheet_xml(rows, hyperlinks_out):
    """Build one worksheet's XML for the given rows; appends (ref, url) to hyperlinks_out."""
    def cell(ref, val, style=0):
        st = f' s="{style}"' if style else ""
        if isinstance(val, (int, float)) and not isinstance(val, bool):
            return f'<c r="{ref}"{st}><v>{val}</v></c>'
        return f'<c r="{ref}"{st} t="inlineStr"><is><t xml:space="preserve">{escape(str(val or ""))}</t></is></c>'

    xml_rows = ['<row r="1" ht="20" customHeight="1">' + "".join(cell(f"{col_letter(i)}1", c, 1) for i, c in enumerate(XCOLS)) + "</row>"]
    for r_i, row in enumerate(rows, start=2):
        base = STATUS_STYLE.get(row.get("Status", "New"), 0)
        cells = []
        for c_i, name in enumerate(XCOLS):
            ref = f"{col_letter(c_i)}{r_i}"
            if name == "Link":
                cells.append(cell(ref, "open", 2 if not base else base + 10))
                if row.get("Link"):
                    hyperlinks_out.append((ref, row["Link"]))
            elif name == "Folder":
                if row.get("Folder"):
                    cells.append(cell(ref, "folder", 2 if not base else base + 10))
                    hyperlinks_out.append((ref, "file://" + urllib.parse.quote(row["Folder"].rstrip("/") + "/")))
                else:
                    cells.append(cell(ref, "", base))
            elif name == "Fit":
                cells.append(cell(ref, "★" * int(row["Fit"] or 0), 3 if not base else base + 20))
            else:
                cells.append(cell(ref, row.get(name, ""), base))
        xml_rows.append(f'<row r="{r_i}">' + "".join(cells) + "</row>")
    last = f"{col_letter(len(XCOLS) - 1)}{len(rows) + 1}"
    cols = "".join(f'<col min="{i+1}" max="{i+1}" width="{w}" customWidth="1"/>' for i, w in enumerate(XWIDTH))
    status_col = col_letter(XCOLS.index("Status"))
    hl = hyperlinks_out
    return (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            f'<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
            f'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
            f'<sheetViews><sheetView workbookViewId="0"><pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/></sheetView></sheetViews>'
            f'<cols>{cols}</cols><sheetData>{"".join(xml_rows)}</sheetData>'
            f'<autoFilter ref="A1:{last}"/>'
            f'<dataValidations count="1"><dataValidation type="list" allowBlank="1" sqref="{status_col}2:{status_col}5000">'
            f'<formula1>"{",".join(STATUSES)}"</formula1></dataValidation></dataValidations>'
            + (('<hyperlinks>' + "".join(f'<hyperlink ref="{ref}" r:id="rIdH{i}"/>' for i, (ref, _) in enumerate(hl)) + '</hyperlinks>') if hl else "")
            + '</worksheet>')


def _summary_xml(groups, sheet_names, label="Company"):
    hdr = [label, "Sheet", "Tracked", "4-5 stars", "Applied", "Interview / Offer", "Move ahead / Interested", "New"]
    def cell(ref, val, style=0):
        st = f' s="{style}"' if style else ""
        if isinstance(val, (int, float)) and not isinstance(val, bool):
            return f'<c r="{ref}"{st}><v>{val}</v></c>'
        return f'<c r="{ref}"{st} t="inlineStr"><is><t xml:space="preserve">{escape(str(val))}</t></is></c>'
    rows = ['<row r="1">' + "".join(cell(f"{col_letter(i)}1", c, 1) for i, c in enumerate(hdr)) + "</row>"]
    links = []
    for r_i, (company, rws) in enumerate(groups, start=2):
        st = [r["Status"] for r in rws]
        vals = [company, sheet_names[company], len(rws), sum(1 for r in rws if int(r["Fit"] or 0) >= 4), st.count("Applied"),
                st.count("Interview") + st.count("Offer"), st.count("Move ahead") + st.count("Interested"), st.count("New")]
        rows.append(f'<row r="{r_i}">' + "".join(cell(f"{col_letter(i)}{r_i}", v, 2 if i == 1 else 0) for i, v in enumerate(vals)) + "</row>")
        links.append((f"B{r_i}", f"#'{sheet_names[company]}'!A1"))
    cols = "".join(f'<col min="{i+1}" max="{i+1}" width="{w}" customWidth="1"/>' for i, w in enumerate([16, 18, 9, 10, 9, 16, 22, 7]))
    hl = "".join(f'<hyperlink ref="{ref}" location="{escape(loc[1:])}" display="{escape(loc[2:].split(chr(39))[1])}"/>' for ref, loc in links)
    return (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            f'<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
            f'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
            f'<sheetViews><sheetView workbookViewId="0" tabSelected="1"><pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/></sheetView></sheetViews>'
            f'<cols>{cols}</cols><sheetData>{"".join(rows)}</sheetData><hyperlinks>{hl}</hyperlinks></worksheet>')


def _row_sort_key(r):
    if WB.get("sort") == "fit_comp":   # fit first, then package (USD upper bound) high to low, undisclosed last within the band
        return (-int(r["Fit"] or 0), -(r.get("Comp USD") or 0), r["Title"])
    return (-int(r["Fit"] or 0), r["Title"])


def _backup_xlsx():
    """Keep the last 10 workbooks under backups/ so an overwritten edit can always be recovered."""
    if not os.path.exists(XLSX):
        return
    bdir = os.path.join(HERE, "backups"); os.makedirs(bdir, exist_ok=True)
    import shutil
    shutil.copy2(XLSX, os.path.join(bdir, dt.datetime.now().strftime("tracker_%Y%m%d_%H%M%S.xlsx")))
    old = sorted(f for f in os.listdir(bdir) if f.startswith("tracker_") and f.endswith(".xlsx"))
    for f in old[:-10]:
        os.remove(os.path.join(bdir, f))


def write_xlsx(rows):
    """Workbook: a Summary sheet plus one sheet per company (default) or per country (workbook.group_by = "country";
    a posting listing several countries appears on each). Rows sorted by fit, or fit then package. Status colours the row.
    GATE 2: never rewrite a workbook that Excel has open (lock file ~$tracker.xlsx); always back up before overwriting;
    verify afterwards that every non-New status in the tracker is present in the file."""
    if os.path.exists(os.path.join(HERE, "~$tracker.xlsx")):
        print("GATE: skipped workbook rewrite (open in Excel)")
        return
    _backup_xlsx()
    _write_xlsx(rows)
    want = sum(1 for r in rows if r.get("Status") and r["Status"] != "New")
    got = sum(1 for e in read_xlsx_edits_raw().values() if e["status"] and e["status"] != "New")
    print(f"GATE: workbook written; {got} non-New status rows in file ({want} unique in tracker); backup kept under backups/")


def _write_xlsx(rows):
    by = {}
    by_country = WB.get("group_by") == "country"
    for r in rows:
        if by_country:
            cs = countries_of(r["Location"])
            for c in cs:
                rc = dict(r); rc["Country"] = c
                rc["Also on"] = ", ".join(x for x in cs if x != c)   # same posting, other sheets
                by.setdefault(c, []).append(rc)
        else:
            by.setdefault(r["Company"], []).append(r)
    groups = sorted(by.items(), key=lambda kv: (-len(kv[1]), kv[0]))
    for _, rws in groups:
        rws.sort(key=_row_sort_key)
    used = {"Summary"}
    sheet_names = {c: _sheet_name(c, used) for c, _ in groups}
    sheets = [("Summary", _summary_xml(groups, sheet_names, "Country" if by_country else "Company"), [])]
    for company, rws in groups:
        hl = []
        sheets.append((sheet_names[company], _sheet_xml(rws, hl), hl))

    fills = ['<fill><patternFill patternType="none"/></fill>', '<fill><patternFill patternType="gray125"/></fill>',
             f'<fill><patternFill patternType="solid"><fgColor rgb="FF1F3A68"/></patternFill></fill>',   # 2 header navy
             '<fill><patternFill patternType="solid"><fgColor rgb="FFE2F0D9"/></patternFill></fill>',   # 3 applied light green
             '<fill><patternFill patternType="solid"><fgColor rgb="FFC6E0B4"/></patternFill></fill>',   # 4 interview/offer green
             '<fill><patternFill patternType="solid"><fgColor rgb="FFDDEBF7"/></patternFill></fill>',   # 5 move ahead light blue
             '<fill><patternFill patternType="solid"><fgColor rgb="FFEDEDED"/></patternFill></fill>']   # 6 skip/rejected grey
    fonts = ['<font><sz val="11"/><name val="Calibri"/></font>',
             '<font><b/><sz val="11"/><color rgb="FFFFFFFF"/><name val="Calibri"/></font>',
             '<font><u/><sz val="11"/><color rgb="FF1F3A68"/><name val="Calibri"/></font>',
             '<font><sz val="14"/><color rgb="FFB8691A"/><name val="Calibri"/></font>']                 # 3 stars: bigger, amber
    wrap = '<alignment vertical="top" wrapText="1"/>'
    xfs = ['<xf fontId="0" fillId="0" borderId="1" applyBorder="1" applyAlignment="1">' + wrap + '</xf>',                 # 0 normal
           '<xf fontId="1" fillId="2" borderId="1" applyBorder="1" applyFont="1" applyFill="1"/>',                        # 1 header
           '<xf fontId="2" fillId="0" borderId="1" applyBorder="1" applyFont="1"/>',                                      # 2 link
           '<xf fontId="3" fillId="0" borderId="1" applyBorder="1" applyFont="1" applyAlignment="1"><alignment horizontal="center" vertical="top"/></xf>',  # 3 stars
           '<xf fontId="0" fillId="3" borderId="1" applyBorder="1" applyFill="1" applyAlignment="1">' + wrap + '</xf>',   # 4 applied
           '<xf fontId="0" fillId="4" borderId="1" applyBorder="1" applyFill="1" applyAlignment="1">' + wrap + '</xf>',   # 5 interview/offer
           '<xf fontId="0" fillId="5" borderId="1" applyBorder="1" applyFill="1" applyAlignment="1">' + wrap + '</xf>',   # 6 move ahead/interested
           '<xf fontId="0" fillId="6" borderId="1" applyBorder="1" applyFill="1" applyAlignment="1">' + wrap + '</xf>']   # 7 skip/rejected
    while len(xfs) < 14:
        xfs.append(xfs[0])
    for fill in (3, 4, 5, 6):   # 14..17 link cells on coloured rows
        xfs.append(f'<xf fontId="2" fillId="{fill}" borderId="1" applyBorder="1" applyFont="1" applyFill="1"/>')
    while len(xfs) < 24:
        xfs.append(xfs[0])
    for fill in (3, 4, 5, 6):   # 24..27 star cells on coloured rows
        xfs.append(f'<xf fontId="3" fillId="{fill}" borderId="1" applyBorder="1" applyFont="1" applyFill="1" applyAlignment="1"><alignment horizontal="center" vertical="top"/></xf>')
    styles = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
              '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
              f'<fonts count="{len(fonts)}">{"".join(fonts)}</fonts><fills count="{len(fills)}">{"".join(fills)}</fills>'
              '<borders count="2"><border/><border><left style="thin"><color rgb="FFBFBFBF"/></left><right style="thin"><color rgb="FFBFBFBF"/></right><top style="thin"><color rgb="FFBFBFBF"/></top><bottom style="thin"><color rgb="FFBFBFBF"/></bottom></border></borders><cellStyleXfs count="1"><xf/></cellStyleXfs>'
              f'<cellXfs count="{len(xfs)}">{"".join(xfs)}</cellXfs></styleSheet>')
    wb = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
          '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
          'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets>'
          + "".join(f'<sheet name="{escape(n)}" sheetId="{i+1}" r:id="rId{i+1}"/>' for i, (n, _, _) in enumerate(sheets))
          + '</sheets></workbook>')
    wb_rels = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
               '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
               + "".join(f'<Relationship Id="rId{i+1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet{i+1}.xml"/>' for i in range(len(sheets)))
               + f'<Relationship Id="rId{len(sheets)+1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>'
               '</Relationships>')
    ct = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
          '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
          '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
          '<Default Extension="xml" ContentType="application/xml"/>'
          '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'
          + "".join(f'<Override PartName="/xl/worksheets/sheet{i+1}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>' for i in range(len(sheets)))
          + '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/></Types>')
    rels = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>'
            '</Relationships>')
    tmp = XLSX + ".tmp"
    with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", ct)
        z.writestr("_rels/.rels", rels)
        z.writestr("xl/workbook.xml", wb)
        z.writestr("xl/_rels/workbook.xml.rels", wb_rels)
        z.writestr("xl/styles.xml", styles)
        for i, (name, xml, hl) in enumerate(sheets):
            z.writestr(f"xl/worksheets/sheet{i+1}.xml", xml)
            if hl:
                z.writestr(f"xl/worksheets/_rels/sheet{i+1}.xml.rels",
                           '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                           '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
                           + "".join(f'<Relationship Id="rIdH{k}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" '
                                     f'Target="{escape(url)}" TargetMode="External"/>' for k, (_, url) in enumerate(hl))
                           + '</Relationships>')
    os.replace(tmp, XLSX)
    with open(XLSX_STATE, "w") as f:
        json.dump({r["ID"]: {"status": r.get("Status", "New"), "notes": r.get("My notes", "")} for r in rows}, f)


def import_workbook_edits(tracker, today, label="GATE"):
    """Apply Status / My-notes edits from the workbook to the tracker, log them, persist at once."""
    imported = []
    for jid, e in read_xlsx_edits().items():
        if jid in tracker:
            if e["status"] and e["status"] != tracker[jid].get("status"):
                imported.append(f"{tracker[jid]['company']} | {tracker[jid]['title'][:50]} | {tracker[jid]['location'][:30]}: {tracker[jid].get('status')} -> {e['status']}")
                tracker[jid]["status"] = e["status"]
                tracker[jid]["status_date"] = today
            if e["notes"] and e["notes"] != tracker[jid].get("notes", ""):
                tracker[jid]["notes"] = e["notes"]
    if imported:
        save(TRACKER, tracker)
        print(f"{label}: imported {len(imported)} status edit(s) from the workbook:")
        for line in imported:
            print("  " + line)
    else:
        print(f"{label}: no new status edits in the workbook")
    return imported


def read_xlsx_edits_raw():
    """Status/notes per ID from the workbook regardless of what we last wrote (used to verify a write)."""
    return read_xlsx_edits(raw=True)


def read_xlsx_edits(raw=False):
    """Read Status and My notes per ID from every sheet of the workbook (edits made in Excel)."""
    if not os.path.exists(XLSX):
        return {}
    try:
        z = zipfile.ZipFile(XLSX)
        shared = []
        if "xl/sharedStrings.xml" in z.namelist():
            ss = z.read("xl/sharedStrings.xml").decode("utf8", "ignore")
            shared = [htmllib.unescape(re.sub(r"<[^>]+>", "", si)) for si in re.findall(r"<si>(.*?)</si>", ss, re.S)]
        sheets = [n for n in z.namelist() if re.match(r"xl/worksheets/sheet\d+\.xml$", n)]
    except (zipfile.BadZipFile, KeyError):
        return {}
    edits = {}
    written = {} if raw else load(XLSX_STATE, {})   # what we last wrote; a cell counts as edited only if it differs from this
    for sh in sheets:
        sheet = z.read(sh).decode("utf8", "ignore")
        all_rows = re.findall(r"<row[^>]*>(.*?)</row>", sheet, re.S)
        if not all_rows:
            continue
        def cell_values(row):
            out = {}
            for ref, attrs, inner in re.findall(r'<c r="([A-Z]+)\d+"([^>]*?)(?:/>|>(.*?)</c>)', row, re.S):
                col = 0
                for ch in ref:
                    col = col * 26 + ord(ch) - 64
                col -= 1
                if 't="s"' in attrs:          # Excel-saved file: shared string reference
                    v = re.search(r"<v>(.*?)</v>", inner or "")
                    val = shared[int(v.group(1))] if v and int(v.group(1)) < len(shared) else ""
                elif 't="inlineStr"' in attrs:  # our own writer
                    val = htmllib.unescape(re.sub(r"<[^>]+>", "", inner or ""))
                else:
                    v = re.search(r"<v>(.*?)</v>", inner or "")
                    val = htmllib.unescape(v.group(1)) if v else ""
                out[col] = val.strip()
            return out
        hdr_vals = cell_values(all_rows[0])
        idx = {name: i for i, name in hdr_vals.items()}
        if not {"ID", "Status", "My notes"} <= set(idx):
            continue   # Summary sheet or foreign layout
        for row in all_rows[1:]:
            vals = cell_values(row)
            jid = vals.get(idx["ID"], "")
            if not jid:
                continue
            cur = {"status": vals.get(idx["Status"], ""), "notes": vals.get(idx["My notes"], "")}
            was = written.get(jid)
            if was is None or cur["status"] != was.get("status", "") or cur["notes"] != was.get("notes", ""):
                edits[jid] = cur
    return edits


# ----------------------------------------------------------------------------- outputs
def load(path, default):
    if os.path.exists(path):
        with open(path) as f:
            return json.load(f)
    return default


def save(path, obj):
    tmp = path + ".tmp"
    with open(tmp, "w") as f:
        json.dump(obj, f, indent=1, sort_keys=True, ensure_ascii=False)
    os.replace(tmp, path)


def jd_hash(text):
    """MinHash signature (48 values) over word 3-grams of the description, so postings that differ only by a city
    sentence or a pay range still match. Stored per row; same_jd_map() groups rows whose estimated Jaccard >= 0.9."""
    import hashlib
    words = re.findall(r"[a-z]+", (text or "").lower())
    if len(words) < 20:
        return ""
    grams = {" ".join(words[i:i + 3]) for i in range(len(words) - 2)}
    sig = []
    for seed in range(48):
        m = min(int(hashlib.blake2b(f"{seed}|{g}".encode(), digest_size=8).hexdigest(), 16) for g in grams)
        sig.append(format(m, "016x")[:8])
    return ",".join(sig)


def _jd_similar(a, b):
    if not a or not b or "," not in a or "," not in b:
        return a == b and bool(a)
    xa, xb = a.split(","), b.split(",")
    return sum(1 for x, y in zip(xa, xb) if x == y) / len(xa) >= 0.9


def same_jd_map(tracker):
    """jid -> 'City A; City B' for other rows of the same company whose descriptions are near-identical."""
    by_co = {}
    for jid, t in tracker.items():
        if t.get("jd_hash"):
            by_co.setdefault(t.get("company"), []).append(jid)
    out = {}
    for ids in by_co.values():
        for jid in ids:
            others = [o for o in ids if o != jid and _jd_similar(tracker[jid]["jd_hash"], tracker[o]["jd_hash"])]
            if others:
                cities = [tracker[o].get("location", "").split(";")[0].split(",")[0].strip() for o in others]
                out[jid] = "; ".join(dict.fromkeys(cities))
    return out


def tracker_rows(tracker):
    rows = []
    dup = same_jd_map(tracker)
    for jid, t in tracker.items():
        rows.append({"Found": t.get("found", ""), "Company": t.get("company", ""), "Title": t.get("title", ""),
                     "Location": t.get("location", ""), "Posted": t.get("posted", ""), "Compensation": t.get("comp", ""),
                     "Comp USD": comp_usd(t.get("comp", ""), t.get("location", "")),
                     "Country": ", ".join(countries_of(t.get("location", ""))),
                     "Same JD as": dup.get(jid, ""),
                     "Fit": t.get("fit", 0), "Fit notes": t.get("fit_notes", ""), "Flags": ", ".join(t.get("flags", [])),
                     "Status": t.get("status", "New"), "My notes": t.get("notes", ""), "Link": t.get("url", ""),
                     "Folder": t.get("folder", ""), "Source": t.get("source", ""), "ID": jid})
    order = {s: i for i, s in enumerate(STATUSES)}
    rows.sort(key=lambda r: (order.get(r["Status"], 99), -int(r["Fit"] or 0), r["Found"], r["Company"]))
    return rows


def write_pipeline(tracker):
    lines = ["# Job pipeline (written by jobwatch; Claude reads this at session start)", "",
             f"Updated {dt.date.today().isoformat()}. A status here means you want to act: "
             "tailor a resume, draft a cover letter, or prep for an interview. Open the JD link before acting.", ""]
    active = [(jid, t) for jid, t in tracker.items() if t.get("status") in PIPELINE_STATUSES]
    if not active:
        lines.append("Nothing marked Move ahead / Applied / Interview / Offer right now.")
    for status in ["Interview", "Offer", "Applied", "Move ahead"]:
        items = [(jid, t) for jid, t in active if t.get("status") == status]
        if not items:
            continue
        lines += [f"## {status}", ""]
        for jid, t in sorted(items, key=lambda x: -int(x[1].get("fit", 0))):
            comp = f", {t['comp']}" if t.get("comp") else ""
            lines.append(f"- **{t['company']}: {t['title']}** ({t['location']}) fit {'★' * int(t.get('fit', 0))}{comp}. {t['url']}")
            if t.get("notes"):
                lines.append(f"  - Note: {t['notes']}")
            if t.get("flags"):
                lines.append(f"  - Flags from JD: {', '.join(t['flags'])}")
            lines.append(f"  - id `{jid}`, found {t.get('found', '')}, status set {t.get('status_date', '')}")
        lines.append("")
    with open(PIPELINE, "w") as f:
        f.write("\n".join(lines) + "\n")


def notify(new, name="jobwatch"):
    if not new or sys.platform != "darwin":
        return
    best = sorted(new, key=lambda r: -r["fit"])[:3]
    msg = f"{len(new)} new: " + "; ".join(f"{r['company']} {'★' * r['fit']} {r['title'][:38]}" for r in best)
    script = f'display notification {json.dumps(msg)} with title {json.dumps(name)} subtitle "tracker.xlsx updated"'
    subprocess.run(["osascript", "-e", script], capture_output=True)


# ----------------------------------------------------------------------------- main
def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--all", action="store_true")
    ap.add_argument("--dry", action="store_true")
    ap.add_argument("--only", default="")
    ap.add_argument("--probe", action="store_true")
    ap.add_argument("--rescore", action="store_true")
    ap.add_argument("--rescore-offline", action="store_true", help="re-score every tracked row from jd_store.json text under the current rules; no network")
    ap.add_argument("--status", nargs=2, metavar=("ID_OR_URL", "STATUS"))
    ap.add_argument("--note", default=None)
    ap.add_argument("--pipeline", action="store_true")
    ap.add_argument("--prune", action="store_true", help="drop untouched rows the current rules reject; no network")
    ap.add_argument("--add-manual", nargs=5, metavar=("COMPANY", "TITLE", "LOCATION", "URL", "TEXTFILE"),
                    help="import a posting whose site blocks fetching: pass fields plus a file holding the description text")
    ap.add_argument("--add", nargs="+", metavar="URL", help="import one or more job URLs (LinkedIn job view, Eightfold, generic) into the tracker")
    ap.add_argument("--dir", default=None, help="project folder with jobwatch_config.json/tracker (default: this script's folder)")
    args = ap.parse_args()

    cfg = load(CONFIG, None)
    if cfg is None:
        sys.exit(f"missing {CONFIG}")
    configure_workbook(cfg)
    tracker = load(TRACKER, {})
    today = dt.date.today().isoformat()

    # ---- GATE 1: pick up edits made in Excel before anything else, log them, and persist them at once so a crash or
    #      a network failure later in the run can never lose a status typed into the workbook.
    import_workbook_edits(tracker, today)
    if os.path.exists(os.path.join(HERE, "~$tracker.xlsx")):
        print("GATE: tracker.xlsx is OPEN in Excel; the workbook will NOT be rewritten this run (tracker.json and pipeline.md still update). Save and close it before the next run.")

    # ---- import pasted URLs
    if args.add_manual:
        company, title, loc, url, textfile = args.add_manual
        text = open(textfile, encoding="utf-8").read()
        job = {"title": title, "location": loc}
        stars, notes, flags = score_fit(job, text + "\nCOMP: " + find_comp(text), cfg)
        key = f"manual-{company.lower().split()[0]}-{abs(hash(url)) % 10**8}"
        e = tracker.get(key, {"found": today, "status": "New", "notes": "added manually"})
        e.update({"company": company, "title": title, "location": loc, "posted": "", "url": url, "source": "manual",
                  "comp": find_comp(text), "fit": stars, "fit_notes": notes, "flags": flags, "low_priority_location": False, "last_seen": today,
                  "manual_text": os.path.abspath(textfile)})
        tracker[key] = e
        print(f"{'★' * e['fit']} | {company} | {title} | {loc} | {e['comp'] or 'no salary'} | flags: {', '.join(flags) or 'none'}")
        print(f"   {notes[:300]}")
        save(TRACKER, tracker); write_xlsx(tracker_rows(tracker)); write_pipeline(tracker)
        return

    if args.add:
        for url in args.add:
            try:
                key, e, text = import_url(url, cfg, tracker, today)
                print(f"{'★' * e['fit']} | {e['company']} | {e['title']} | {e['location']} | {e.get('comp') or 'no salary'} | flags: {', '.join(e['flags']) or 'none'}")
                print(f"   {e['fit_notes'][:300]}")
                mq = re.search(r"(Minimum Qualifications|Required Qualifications|Requirements|What We.re Looking For|Qualifications)(.{0,900})", text, re.S | re.I)
                if mq:
                    print("   quals:", re.sub(r"\s+", " ", mq.group(0))[:900])
            except Exception as ex:  # noqa: BLE001
                print("FAILED", url, ex)
        save(TRACKER, tracker); write_xlsx(tracker_rows(tracker)); write_pipeline(tracker)
        return

    # ---- status update from the terminal
    if args.status:
        key, status = args.status
        if status not in STATUSES:
            sys.exit(f"status must be one of: {', '.join(STATUSES)}")
        hits = [j for j, t in tracker.items()
                if key in j or key in t.get("url", "") or key.lower() in t.get("title", "").lower()]
        if len(hits) != 1:
            sys.exit(f"{len(hits)} postings match {key!r}: " + ", ".join(hits[:6]))
        tracker[hits[0]].update({"status": status, "status_date": today})
        if args.note is not None:
            tracker[hits[0]]["notes"] = args.note
        save(TRACKER, tracker)
        write_xlsx(tracker_rows(tracker))
        write_pipeline(tracker)
        print(f"{tracker[hits[0]]['company']}: {tracker[hits[0]]['title']} -> {status}")
        return

    if args.pipeline:
        write_pipeline(tracker)
        print(open(PIPELINE).read())
        return

    # ---- prune: drop untouched (status New, no notes) rows that the current exclude list would now reject
    for jid in [j for j, t in tracker.items()
                if t.get("status", "New") == "New" and not t.get("notes")
                and match({"title": t.get("title", ""), "location": t.get("location", "")}, cfg) is None]:
        del tracker[jid]
    # ---- fit.drop_flags (opt-in per config): untouched New rows carrying a hard-reject flag move to dropped.json
    dropped = load(DROPPED, {})
    jd_store = load(JD_STORE, {}) if cfg["fit"].get("save_jd_text") else None
    drop_flags = set(cfg["fit"].get("drop_flags", []))
    if drop_flags:
        for jid, t in list(tracker.items()):
            if (t.get("status", "New") == "New" and not t.get("notes") and t.get("source") not in ("manual", "linkedin")
                    and drop_flags & set(t.get("flags", []))):
                dropped[jid] = {k: t.get(k, "") for k in ("company", "title", "location", "url", "fit", "flags", "fit_notes")}
                dropped[jid]["dropped_on"] = dt.date.today().isoformat()
                del tracker[jid]
    if args.rescore_offline:
        store = load(JD_STORE, {})
        n, moved = 0, 0
        for jid, t in list(tracker.items()):
            text = store.get(jid) or (open(t["manual_text"], encoding="utf-8").read() if t.get("manual_text") and os.path.exists(t["manual_text"]) else "")
            if not text:
                continue
            stars, notes, flags = score_fit({"title": t["title"], "location": t["location"], "company": t["company"]},
                                            text + ("\nCOMP: " + t["comp"] if t.get("comp") else ""), cfg)
            moved += int(t.get("fit", 0)) != stars
            if drop_flags & set(flags) and t.get("status", "New") == "New" and not t.get("notes") and t.get("source") not in ("manual", "linkedin"):
                dropped[jid] = {k: t.get(k, "") for k in ("company", "title", "location", "url")} | {"fit": stars, "flags": flags, "fit_notes": notes, "dropped_on": today}
                del tracker[jid]
                continue
            t.update({"fit": stars, "fit_notes": notes, "flags": flags}); n += 1
        save(DROPPED, dropped); save(TRACKER, tracker); write_xlsx(tracker_rows(tracker)); write_pipeline(tracker)
        print(f"re-scored {n} rows offline ({moved} changed stars); tracked total now {len(tracker)}; dropped archive {len(dropped)}")
        return
    if args.prune:
        save(DROPPED, dropped)
        save(TRACKER, tracker); write_xlsx(tracker_rows(tracker)); write_pipeline(tracker)
        print(f"pruned; tracked total now {len(tracker)}")
        return

    # ---- fetch listings
    only = {s.strip() for s in args.only.split(",") if s.strip()}
    results, errors, counts, src_of = [], [], {}, {}
    for key, src in cfg["sources"].items():
        if (only and key not in only) or not src.get("enabled", True):
            continue
        fn = FETCHERS[src["type"]]
        try:
            jobs = fn(src, cfg["search_queries"]) if src["type"] in QUERY_TYPES else fn(src)
        except (urllib.error.HTTPError, urllib.error.URLError, RuntimeError, KeyError, ValueError) as e:
            errors.append(f"{key}: {e}")
            continue
        counts[key] = len(jobs)
        for j in jobs:
            if src.get("resolve_locations") and re.fullmatch(r"\d+ Locations?", (j.get("location") or "").strip()) \
                    and any(k in norm(j["title"]) for k in cfg["title_include"]) and "_wd_detail" in j:
                try:
                    info = http(j["_wd_detail"]).get("jobPostingInfo") or {}
                    locs = [info.get("location", "")] + [a for a in (info.get("additionalLocations") or []) if a]
                    j["location"] = "; ".join(x for x in locs if x)
                except RuntimeError:
                    pass
            m = match(j, cfg)
            if m:
                j.update({"company": src["name"], "source": key, **m})
                results.append(j)
                src_of[j["id"]] = src

    if args.probe:
        for k, n in counts.items():
            print(f"{k:14s} ok  {n} postings fetched")
        for e in errors:
            print("ERROR", e)
        return

    # ---- score new (or everything with --rescore) and update the tracker
    new = []
    for j in results:
        jid = j["id"]
        if jid in dropped and not args.rescore:
            continue   # already judged and dropped; --rescore re-judges against the current rules
        fresh = jid not in tracker
        if fresh or args.rescore:
            det = fetch_detail(j, src_of[jid])
            stars, notes, flags = score_fit(j, det.get("text", "") + ("\nCOMP: " + det["comp"] if det.get("comp") else ""), cfg)
            old = tracker.get(jid, {"status": "New", "notes": ""})
            if drop_flags & set(flags) and old.get("status", "New") == "New" and not old.get("notes"):
                tracker.pop(jid, None)
                dropped[jid] = {"company": j["company"], "title": j["title"], "location": j["location"], "url": j["url"],
                                "fit": stars, "flags": flags, "fit_notes": notes, "dropped_on": today}
                continue
            dropped.pop(jid, None)   # judged fit under the current rules
            if jd_store is not None and det.get("text"):
                jd_store[jid] = det["text"][:20000]
            entry = tracker.get(jid, {"found": today, "status": "New", "notes": ""})
            entry.update({"company": j["company"], "title": j["title"], "location": j["location"],
                          "posted": j["posted"], "url": j["url"], "source": j["source"],
                          "comp": det.get("comp", ""), "fit": stars, "fit_notes": notes, "flags": flags,
                          "low_priority_location": j["low"], "last_seen": today, "jd_hash": jd_hash(det.get("text", ""))})
            if det.get("error"):
                entry["detail_error"] = det["error"]
            tracker[jid] = entry
            if fresh:
                new.append({"company": j["company"], "title": j["title"], "fit": stars, "id": jid})
        else:
            tracker[jid]["last_seen"] = today

    # ---- manual / URL-imported rows are not in any source; re-import them on --rescore
    if args.rescore:
        for jid, t in list(tracker.items()):
            if t.get("manual_text") and os.path.exists(t["manual_text"]):   # pasted-JD rows (Cloudflare-blocked sites): re-score the saved text
                text = open(t["manual_text"], encoding="utf-8").read()
                t["fit"], t["fit_notes"], t["flags"] = score_fit({"title": t["title"], "location": t["location"], "company": t["company"]},
                                                                  text + "\nCOMP: " + find_comp(text), cfg)
                t["last_seen"] = today
                continue
            if t.get("source") in ("manual", "linkedin") and t.get("url", "").startswith("http") and t.get("status") not in ("Skip", "Rejected"):
                try:
                    key, e, _ = import_url(t["url"], cfg, tracker, today)
                    if key != jid:
                        for f in ("status", "notes", "found", "folder", "status_date"):
                            if f in t:
                                e[f] = t[f]
                        tracker.pop(jid, None)
                except Exception as ex:  # noqa: BLE001
                    tracker[jid]["detail_error"] = str(ex)[:120]

    if args.rescore:   # rows no longer in any feed keep their text score but still get the employer-tier rule
        reached = {j["id"] for j in results}
        tgate = cfg["fit"].get("company_negative_title")
        for jid, t in tracker.items():
            if jid in reached or t.get("source") in ("manual", "linkedin"):
                continue
            for k, (w, label) in cfg["fit"].get("company_negative", {}).items():
                if re.search(k, norm(t.get("company", ""))) and (not tgate or re.search(tgate, norm(t.get("title", "")))) and label not in t.get("flags", []):
                    t["flags"] = t.get("flags", []) + [label]
                    cap = cfg["fit"].get("cap_stars_on_flags", {}).get(label)
                    if cap:
                        t["fit"] = min(int(t.get("fit", 1)), cap)
                    t["fit_notes"] = t.get("fit_notes", "") + f"; -{w:g} {label} (posting not in feed today; employer tier applied)"
                    break

    # ---- report
    new_ids = {n["id"] for n in new}
    shown = [t for j, t in tracker.items() if (args.all or j in new_ids)]
    shown.sort(key=lambda t: (-int(t.get("fit", 0)), t["company"]))
    lines = [f"# jobwatch report, {today}", "",
             f"Sources OK: {len(counts)}. Postings scanned: {sum(counts.values())}. Matches: {len(results)}. "
             f"New: {len(new)}. Tracked total: {len(tracker)}.", ""]
    if shown:
        lines += ["| Fit | Company | Title | Location | Comp | Flags | Link |", "|---|---|---|---|---|---|---|"]
        for t in shown:
            lines.append(f"| {'★' * int(t.get('fit', 0))} | {t['company']} | {t['title']} | {t['location']} | "
                         f"{t.get('comp', '')} | {', '.join(t.get('flags', []))} | [open]({t['url']}) |")
    else:
        lines.append("No new matches.")
    if errors:
        lines += ["", "## Source errors", ""] + [f"- {e}" for e in errors]
    report = "\n".join(lines) + "\n"
    print(report)

    if args.dry:
        return
    save(DROPPED, dropped)
    if jd_store is not None:
        save(JD_STORE, {k: v for k, v in jd_store.items() if k in tracker})
    save(TRACKER, tracker)
    import_workbook_edits(tracker, today, label="GATE (re-check before write)")
    write_xlsx(tracker_rows(tracker))
    write_pipeline(tracker)
    with open(REPORT, "w") as f:
        f.write(report)
    save(LATEST, {"date": today, "new": new, "errors": errors, "counts": counts})
    if cfg.get("notify", True):
        notify(new, cfg.get("name", "jobwatch"))


if __name__ == "__main__":
    main()
