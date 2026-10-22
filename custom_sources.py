#!/usr/bin/env python3
"""
Custom fetchers for big employers that don't use a standard ATS feed.
Each returns a list of jobs in the SAME shape as poller.norm() so the poller can
classify/filter/dedupe them exactly like everything else.

Currently: Amazon (public amazon.jobs search.json API).
Google / Meta / Microsoft gate their APIs and need browser automation — TODO.
"""
import json
import ssl
from urllib.parse import quote
from urllib.request import Request, urlopen

try:
    import certifi
    SSL_CTX = ssl.create_default_context(cafile=certifi.where())
except ImportError:
    SSL_CTX = ssl.create_default_context()

UA = {"User-Agent": "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36",
      "Accept": "application/json"}
TIMEOUT = 15


def _norm(company, provider, slug, uid, title, location, url, posted, comp=""):
    return {"company": company, "provider": provider, "slug": slug,
            "jid": str(uid), "uid": f"{company}:{uid}", "title": title,
            "location": location or "", "url": url, "posted": posted or "", "comp": comp or ""}


def _get(url):
    try:
        with urlopen(Request(url, headers=UA), timeout=TIMEOUT, context=SSL_CTX) as r:
            return json.loads(r.read().decode("utf-8", "replace"))
    except Exception:
        return None


# Broad queries; the poller's classify() narrows to Michelle's role families.
AMAZON_QUERIES = ["software engineer", "machine learning", "applied scientist",
                  "product manager", "forward deployed", "new grad",
                  "university graduate", "graduate 2026", "graduate 2027"]


def from_amazon(_company="Amazon", _slug="amazon"):
    seen, out = set(), []
    for q in AMAZON_QUERIES:
        d = _get(f"https://www.amazon.jobs/en/search.json?base_query={quote(q)}"
                 f"&normalized_country_code=US&result_limit=100&sort=recent")
        if not isinstance(d, dict):
            continue
        for j in d.get("jobs", []):
            jid = j.get("id_icims") or j.get("id") or j.get("job_path")
            if not jid or jid in seen:
                continue
            seen.add(jid)
            path = j.get("job_path", "")
            url = f"https://www.amazon.jobs{path}" if path.startswith("/") else path
            out.append(_norm("Amazon", "amazon", "amazon", jid, j.get("title", ""),
                             j.get("normalized_location", "") or j.get("location", ""),
                             url, j.get("posted_date", ""), ""))
    return out


SIMPLIFY_URL = ("https://raw.githubusercontent.com/SimplifyJobs/"
                "New-Grad-Positions/dev/.github/scripts/listings.json")


def from_simplify(*_):
    """Community-maintained new-grad job list (SimplifyJobs). Every entry is
    already curated as a NEW-GRAD role, so we mark prefiltered_newgrad=True and
    the poller classifies these by role family only (no title-signal needed)."""
    import datetime as _dt
    d = _get(SIMPLIFY_URL)
    out = []
    if not isinstance(d, list):
        return out
    for j in d:
        if not j.get("active") or not j.get("is_visible"):
            continue
        locs = j.get("locations") or []
        loc = ", ".join(locs) if isinstance(locs, list) else str(locs)
        ts = j.get("date_posted") or 0
        posted = _dt.datetime.fromtimestamp(ts).strftime("%Y-%m-%d") if ts else ""
        row = _norm(j.get("company_name", ""), "simplify", "newgrad", j.get("id", ""),
                    j.get("title", ""), loc, j.get("url", ""), posted, "")
        row["prefiltered_newgrad"] = True
        out.append(row)
    return out


# provider name -> fetcher. Poller checks this map for non-standard providers.
CUSTOM_FETCHERS = {"amazon": from_amazon, "simplify": from_simplify}

# Startup boards (Sequoia, a16z, Accel, YC, Wellfound, Built In NYC)
from startup_boards import STARTUP_FETCHERS
CUSTOM_FETCHERS.update(STARTUP_FETCHERS)


if __name__ == "__main__":
    jobs = from_amazon()
    print(f"Amazon: fetched {len(jobs)} US postings")
    for j in jobs[:8]:
        print("  ", j["title"][:55], "|", j["location"])
