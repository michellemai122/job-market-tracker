#!/usr/bin/env python3
"""
Discover which Applicant Tracking System (ATS) each target company uses.

For every company we generate a few plausible "slug" variants and probe the
public job-board JSON endpoints of the four biggest ATS providers. Whatever
returns a non-empty job list wins. Output: a verified registry (registry.json)
mapping company -> (provider, slug, job_count) that the poller then uses.
"""
import json
import re
import ssl
import sys
import time
from concurrent.futures import ThreadPoolExecutor, as_completed
from urllib.request import Request, urlopen
from urllib.error import HTTPError, URLError

try:
    import certifi
    SSL_CTX = ssl.create_default_context(cafile=certifi.where())
except ImportError:
    SSL_CTX = ssl.create_default_context()

TIMEOUT = 12
UA = {"User-Agent": "Mozilla/5.0 (job-tracker; personal use)"}


def slug_variants(name: str):
    """Generate candidate slugs for a company name."""
    n = name.strip().lower()
    n = n.replace("&", "and")
    base = re.sub(r"[^a-z0-9]+", "", n)              # "de shaw" -> "deshaw"
    hyphen = re.sub(r"[^a-z0-9]+", "-", n).strip("-")  # "de shaw" -> "de-shaw"
    nospace_words = n.split()
    variants = {base, hyphen}
    # drop trailing " ai", " inc", " group", " capital", " trading", " partners"
    stripped = re.sub(r"\b(ai|inc|llc|group|capital|trading|partners|company|technologies)\b", "", n)
    variants.add(re.sub(r"[^a-z0-9]+", "", stripped))
    variants.add(re.sub(r"[^a-z0-9]+", "-", stripped).strip("-"))
    # first word only (helps "jane street" -> "janestreet" already covered; "expedia group" -> "expedia")
    if nospace_words:
        variants.add(nospace_words[0])
    return [v for v in variants if v]


def fetch_json(url):
    try:
        req = Request(url, headers=UA)
        with urlopen(req, timeout=TIMEOUT, context=SSL_CTX) as r:
            if r.status != 200:
                return None
            return json.loads(r.read().decode("utf-8", "replace"))
    except (HTTPError, URLError, TimeoutError, ValueError, ConnectionError):
        return None
    except Exception:
        return None


def try_greenhouse(slug):
    d = fetch_json(f"https://boards-api.greenhouse.io/v1/boards/{slug}/jobs")
    if isinstance(d, dict) and d.get("jobs"):
        return len(d["jobs"])
    return 0


def try_lever(slug):
    d = fetch_json(f"https://api.lever.co/v0/postings/{slug}?mode=json")
    if isinstance(d, list) and d:
        return len(d)
    return 0


def try_ashby(slug):
    d = fetch_json(f"https://api.ashbyhq.com/posting-api/job-board/{slug}")
    if isinstance(d, dict) and d.get("jobs"):
        return len(d["jobs"])
    return 0


def try_smartrecruiters(slug):
    d = fetch_json(f"https://api.smartrecruiters.com/v1/companies/{slug}/postings")
    if isinstance(d, dict) and isinstance(d.get("content"), list) and d["content"]:
        return d.get("totalFound", len(d["content"]))
    return 0


PROVIDERS = [
    ("greenhouse", try_greenhouse),
    ("lever", try_lever),
    ("ashby", try_ashby),
    ("smartrecruiters", try_smartrecruiters),
]


def discover_company(name):
    """Return best (provider, slug, count) or None."""
    best = None
    for slug in slug_variants(name):
        for provider, fn in PROVIDERS:
            count = fn(slug)
            if count and (best is None or count > best[2]):
                best = (provider, slug, count)
    if best:
        return {"company": name, "provider": best[0], "slug": best[1], "count": best[2]}
    return None


def main():
    with open("companies.txt") as f:
        companies = [ln.strip() for ln in f if ln.strip() and not ln.startswith("#")]

    found, missing = [], []
    with ThreadPoolExecutor(max_workers=12) as ex:
        futs = {ex.submit(discover_company, c): c for c in companies}
        for fut in as_completed(futs):
            c = futs[fut]
            res = fut.result()
            if res:
                found.append(res)
                print(f"  OK   {c:<28} -> {res['provider']:<15} {res['slug']:<20} ({res['count']} jobs)")
            else:
                missing.append(c)
                print(f"  --   {c:<28} -> no public ATS feed found")

    found.sort(key=lambda r: r["company"].lower())
    with open("registry.json", "w") as f:
        json.dump({"found": found, "missing": missing}, f, indent=2)

    print(f"\nFOUND {len(found)} / {len(companies)} companies with a live feed.")
    print(f"MISSING {len(missing)}: {', '.join(missing)}")


if __name__ == "__main__":
    main()
