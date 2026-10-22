#!/usr/bin/env python3
"""
Startup job boards: VC portfolio boards (Sequoia, a16z, Accel), YC, Wellfound
and Built In NYC. Each fetcher returns jobs in the same shape as poller.norm(),
and custom_sources.CUSTOM_FETCHERS hands them to the poller.

Each board only gets queried for entry-level / role-family slices, never its
full 10k+ listing, so a 10-minute cycle stays cheap. A job is marked
prefiltered_newgrad only when the BOARD itself says it is entry-level (YC's
"new grads ok", Getro's entry_level, Consider's junior-only,
Wellfound descriptions that welcome new grads). Everything
else goes through the poller's strict classify(), which needs a new-grad signal
in the title.

Where the board exposes the posting's own ATS link (Greenhouse, Ashby, ...) we
store that as the URL, so prune_closed.py can ask the ATS whether it's closed
and the poller's URL dedupe catches jobs we already have from registry.json.

Not covered: Ventureloop (Cloudflare blocks scripted requests) and hackNY
(a fellowship program, not a job board).
"""
import html
import json
import re
import ssl
from http.cookiejar import CookieJar
from urllib.parse import quote
from urllib.request import HTTPCookieProcessor, HTTPSHandler, Request, build_opener

try:
    import certifi
    SSL_CTX = ssl.create_default_context(cafile=certifi.where())
except ImportError:
    SSL_CTX = ssl.create_default_context()

UA = "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/126 Safari/537.36"
TIMEOUT = 20
NEWGRAD_TEXT = re.compile(r"new\s*grad|recent\s*grad|new\s*college\s*grad|"
                          r"entry[\s-]*level|early[\s-]*career|0\s*[-–]\s*[12]\s*years", re.I)


def _norm(company, provider, uid, title, location, url, posted="", comp="", newgrad=False):
    row = {"company": (company or "").strip(), "provider": provider, "slug": provider,
           "jid": str(uid), "uid": f"{provider}:{uid}", "title": (title or "").strip(),
           "location": location or "", "url": url, "posted": posted or "", "comp": comp or ""}
    if newgrad:
        row["prefiltered_newgrad"] = True
    return row


def _opener():
    return build_opener(HTTPCookieProcessor(CookieJar()), HTTPSHandler(context=SSL_CTX))


def _get(url, opener=None, headers=None):
    req = Request(url, headers={"User-Agent": UA, **(headers or {})})
    try:
        with (opener or _opener()).open(req, timeout=TIMEOUT) as r:
            return r.read().decode("utf-8", "replace")
    except Exception:
        return ""


def _post_json(url, body, opener=None, headers=None):
    req = Request(url, data=json.dumps(body).encode(),
                  headers={"User-Agent": UA, "Content-Type": "application/json",
                           "Accept": "application/json", **(headers or {})})
    try:
        with (opener or _opener()).open(req, timeout=TIMEOUT) as r:
            return json.loads(r.read().decode("utf-8", "replace"))
    except Exception:
        return None


def _text(fragment):
    return html.unescape(re.sub(r"<[^>]+>", "", fragment or "")).strip()


def _dedupe(jobs):
    seen, out = set(), []
    for j in jobs:
        if j["uid"] not in seen:
            seen.add(j["uid"])
            out.append(j)
    return out


# ---------------------------------------------------------------- Consider (Sequoia)
JUNIOR = ["junior"]
CONSIDER_QUERIES = [  # titlePrefix is Consider's title search; junior is its entry tier
    {"titlePrefix": "new grad"}, {"titlePrefix": "university"}, {"titlePrefix": "early career"},
    {"titlePrefix": "associate product manager"}, {"titlePrefix": "forward deployed"},
    {"titlePrefix": "software engineer", "jobSeniority": JUNIOR},
    {"titlePrefix": "product manager", "jobSeniority": JUNIOR},
    {"titlePrefix": "machine learning", "jobSeniority": JUNIOR},
    {"titlePrefix": "ai engineer", "jobSeniority": JUNIOR},
]


def from_consider(host, board, provider, pages=3):
    """Consider-hosted VC boards. The search API wants the CSRF token and session
    cookie that the /jobs page hands out. A job tagged junior AND mid is not
    clearly entry-level, so only junior-only jobs skip strict matching."""
    op = _opener()
    m = re.search(r'"csrfToken":"([^"]+)"', _get(f"https://{host}/jobs", op))
    if not m:
        return []
    out = []
    for q in CONSIDER_QUERIES:
        out += _consider_query(host, board, provider, q, pages, op, m.group(1))
    return _dedupe(out)


def _consider_query(host, board, provider, query, pages, op, token):
    out, meta = [], {"size": 50}
    for _ in range(pages):
        d = _post_json(f"https://{host}/api-boards/search-jobs",
                       {"meta": meta, "board": {"id": board, "isParent": True},
                        "query": query, "grouped": False},
                       op, {"x-csrf-token": token})
        if not isinstance(d, dict):
            break
        for j in d.get("jobs", []):
            tiers = {s.get("id") for s in j.get("jobSeniorities") or []}
            sal = j.get("salary") or {}
            comp = ""
            if sal.get("minValue") and sal.get("maxValue") and (sal.get("currency") or {}).get("value", "USD") == "USD":
                comp = f"${int(sal['minValue']):,}–${int(sal['maxValue']):,}"
            out.append(_norm(j.get("companyName"), provider, j.get("jobId"), j.get("title"),
                             "; ".join(j.get("locations") or []), j.get("url") or j.get("applyUrl", ""),
                             (j.get("timeStamp") or "")[:10], comp, newgrad=tiers <= {"junior", "intern"}))
        meta = d.get("meta") or {}
        if not meta.get("sequence"):
            break
    return out


def from_sequoia(*_):
    return from_consider("jobs.sequoiacap.com", "sequoia-capital", "sequoia")


# ---------------------------------------------------------------- Getro (Accel)
GETRO_QUERIES = [  # (search text, entry-level filter?)
    ("software engineer", True), ("product manager", True), ("machine learning", True),
    ("ai engineer", True), ("forward deployed", True), ("new grad", False),
    ("university", False), ("early career", False),
]


def from_getro(collection_id, origin, provider, pages=3):
    out = []
    for q, entry in GETRO_QUERIES:
        for page in range(pages):
            d = _post_json(f"https://api.getro.com/api/v2/collections/{collection_id}/search/jobs",
                           {"hitsPerPage": 20, "page": page, "query": q,
                            "filters": {"seniority": ["entry_level"]} if entry else {}},
                           headers={"Origin": origin, "Referer": origin + "/"})
            jobs = ((d or {}).get("results") or {}).get("jobs") or []
            for j in jobs:
                org = j.get("organization") or {}
                comp = ""
                lo, hi = j.get("compensationAmountMinCents"), j.get("compensationAmountMaxCents")
                if lo and hi and j.get("compensationCurrency") in (None, "USD") and j.get("compensationPeriod") == "year":
                    comp = f"${lo // 100:,}–${hi // 100:,}"
                out.append(_norm(org.get("name"), provider, j.get("id"), j.get("title"),
                                 "; ".join(j.get("locations") or []), j.get("url", ""), "", comp,
                                 newgrad=j.get("seniority") == "entry_level"))
            if len(jobs) < 20:
                break
    return _dedupe(out)


def from_accel(*_):
    return from_getro(8672, "https://jobs.accel.com", "accel")


# ---------------------------------------------------------------- a16z (custom Next.js board)
A16Z_QUERIES = ["new grad", "new graduate", "university", "early career", "entry level",
                "associate product manager", "product manager", "forward deployed",
                "software engineer", "machine learning engineer", "ai engineer", "rotational"]


def from_a16z(*_):
    """Server-rendered search page: 25 newest matches per query, which is plenty
    for a poller that runs every 10 minutes and remembers what it has seen."""
    out = []
    for q in A16Z_QUERIES:
        page = _get(f"https://jobs.a16z.com/jobs?q={quote(q)}")
        for card in page.split("<article")[1:]:
            m = re.search(r'href="/jobs/([^/"]+)/([^"]+)">(.*?)</a></h2>', card, re.S)
            if not m:
                continue
            jid = m.group(2).split("--")[0]
            company = _text((re.search(r'data-nosnippet[^>]*>\s*<a[^>]*>(.*?)</a>', card, re.S) or [None, m.group(1)])[1])
            spans = [_text(s) for s in re.findall(r'<span>([^<]+)</span>', card)]
            loc = next((s for s in spans if not s.startswith(("USD", "Posted"))), "")
            comp = next((s.split(" (")[0].replace("USD ", "$") for s in
                         [_text(x) for x in re.findall(r'</svg>(USD[^<]+)</span>', card)]), "")
            posted = (re.search(r'dateTime="(\d{4}-\d\d-\d\d)', card) or [None, ""])[1]
            apply = re.search(r'<a href="(https?://[^"]+)" target="_blank"', card)
            url = html.unescape(apply.group(1)) if apply else f"https://jobs.a16z.com/jobs/{m.group(1)}/{m.group(2)}"
            out.append(_norm(company, "a16z", jid, _text(m.group(3)), loc, url, posted, comp))
    return _dedupe(out)


# ---------------------------------------------------------------- YC (ycombinator.com/jobs)
YC_PAGES = ["software-engineer", "software-engineer/new-york", "software-engineer/san-francisco",
            "software-engineer/remote", "product-manager", "product-manager/new-york",
            "product-manager/san-francisco", "machine-learning-engineer", "ai-engineer",
            "full-stack-engineer", "backend-engineer", "frontend-engineer"]


def from_yc(*_):
    """YC role pages embed their postings as JSON. YC tags each posting's minimum
    experience; "Any (new grads ok)" is the founders saying a new grad fits."""
    out = []
    for p in YC_PAGES:
        m = re.search(r'data-page="([^"]+)"', _get(f"https://www.ycombinator.com/jobs/role/{p}"))
        if not m:
            continue
        try:
            props = json.loads(html.unescape(m.group(1)))["props"]
        except (ValueError, KeyError):
            continue
        for j in props.get("jobPostings") or []:
            if (j.get("type") or "Full-time") != "Full-time":
                continue
            ok = "new grad" in (j.get("minExperience") or "").lower()
            out.append(_norm(j.get("companyName"), "yc", j.get("id"), j.get("title"),
                             j.get("location", ""), "https://www.ycombinator.com" + j.get("url", ""),
                             "", j.get("salaryRange", ""), newgrad=ok))
    return _dedupe(out)


# ---------------------------------------------------------------- Wellfound
WELLFOUND_PAGES = ["software-engineer/new-york", "software-engineer/san-francisco",
                   "software-engineer", "product-manager/new-york", "product-manager",
                   "machine-learning-engineer", "artificial-intelligence-engineer"]


def from_wellfound(*_):
    """Role pages carry an Apollo cache in __NEXT_DATA__. Titles there rarely say
    'new grad', so a description that welcomes new grads counts as the signal."""
    out = []
    for p in WELLFOUND_PAGES:
        m = re.search(r'<script id="__NEXT_DATA__"[^>]*>(.*?)</script>',
                      _get(f"https://wellfound.com/role/l/{p}" if "/" in p else f"https://wellfound.com/role/{p}"), re.S)
        if not m:
            continue
        try:
            ap = json.loads(m.group(1))["props"]["pageProps"]["apolloState"]["data"]
        except (ValueError, KeyError):
            continue
        owner = {}
        for v in ap.values():
            if v.get("__typename") == "StartupResult":
                for ref in v.get("highlightedJobListings") or []:
                    owner[ref.get("__ref")] = v.get("name", "")
        for key, j in ap.items():
            if j.get("__typename") != "JobListingSearchResult" or j.get("jobType") not in (None, "full-time"):
                continue
            locs = list(j.get("locationNames") or []) + (["Remote"] if j.get("remote") else [])
            ymin = j.get("yearsExperienceMin")
            ok = (ymin is not None and ymin <= 1) or bool(NEWGRAD_TEXT.search(j.get("description") or ""))
            out.append(_norm(owner.get(key, ""), "wellfound", j.get("id"), j.get("title"), "; ".join(locs),
                             f"https://wellfound.com/jobs/{j.get('id')}-{j.get('slug', '')}",
                             "", j.get("compensation", ""), newgrad=ok))
    return _dedupe([j for j in out if j["company"]])


# ---------------------------------------------------------------- Built In NYC
BUILTIN_PAGES = ["dev-engineering/entry-level", "product/entry-level", "data-analytics/entry-level"]


def from_builtin_nyc(*_):
    """Built In's entry-level pages. Their 'entry-level' tag is loose (Group PMs and
    founding roles show up), so these go through the poller's strict matching."""
    out = []
    for p in BUILTIN_PAGES:
        for page in (1, 2):
            body = _get(f"https://www.builtinnyc.com/jobs/{p}?page={page}")
            cards = body.split('data-id="job-card"')[1:]
            for card in cards:
                t = re.search(r'href="(/job/[^"]+/(\d+))"[^>]*data-id="job-card-title"[^>]*>(.*?)</a>', card, re.S)
                if not t:
                    continue
                company = _text((re.search(r'data-id="company-title"[^>]*><span>(.*?)</span>', card, re.S) or [None, ""])[1])
                loc = _text((re.search(r'fa-location-dot.*?<span[^>]*>(.*?)</span>', card, re.S) or [None, "New York, NY"])[1])
                out.append(_norm(company, "builtin", t.group(2), _text(t.group(3)), loc or "New York, NY",
                                 "https://www.builtinnyc.com" + t.group(1)))
            if len(cards) < 20:
                break
    return _dedupe(out)


STARTUP_FETCHERS = {"sequoia": from_sequoia, "a16z": from_a16z, "accel": from_accel,
                    "yc": from_yc, "wellfound": from_wellfound, "builtin_nyc": from_builtin_nyc}


if __name__ == "__main__":
    import sys
    sys.path.insert(0, __import__("os").path.dirname(__file__))
    import poller
    for name, fn in STARTUP_FETCHERS.items():
        jobs = fn()
        hits = [j for j in jobs if (poller.classify_role if j.get("prefiltered_newgrad") else poller.classify)(j["title"])
                and poller.location_wanted(j["location"])]
        print(f"{name:<12} fetched {len(jobs):>4}  new-grad matches {len(hits):>3}")
        for j in hits[:6]:
            print(f"    {j['company'][:18]:<18} {j['title'][:52]:<52} {j['location'][:28]}")
