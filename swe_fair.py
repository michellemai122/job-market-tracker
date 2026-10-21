#!/usr/bin/env python3
"""
WE26 (SWE annual conference, Boston, Nov 5-7 2026) career-fair sweep.

For every recruiting company on the fair list, pull its live early-career /
new-grad (class of 2027) full-time roles straight from its careers backend,
open each posting to read the pay range, and write:
  data/we26_fair.xlsx  -> "Roles" (one row per posting) + "Companies" (one row per booth)
  Google Sheet tabs    -> "SHPE26 Fair Roles" + "SHPE26 Fair Companies" (overwritten each run)

Unlike poller.py this is NOT limited to SWE/PM/AI titles: the fair is
all-engineering, so every engineering/tech/business early-career role is kept and
tagged with a Discipline. Re-run any time before the fair:  python3 swe_fair.py
"""
import json
import os
import re
import ssl
import socket
import time
from datetime import datetime
from concurrent.futures import ThreadPoolExecutor
from urllib.parse import quote
from urllib.request import Request, urlopen
from urllib.error import HTTPError
from zoneinfo import ZoneInfo

try:
    import certifi
    SSL_CTX = ssl.create_default_context(cafile=certifi.where())
except ImportError:
    SSL_CTX = ssl.create_default_context()

from poller import location_ok

socket.setdefaulttimeout(60)
HERE = os.path.dirname(os.path.abspath(__file__))
OUT_XLSX = os.path.join(HERE, "data", "we26_fair.xlsx")
MANUAL_PATH = os.path.join(HERE, "we26_manual.json")  # hand-checked rows for no-feed sites
SIMPLIFY_URL = ("https://raw.githubusercontent.com/SimplifyJobs/"
                "New-Grad-Positions/dev/.github/scripts/listings.json")
ET = ZoneInfo("America/New_York")
TIMEOUT = 25
UA = {"User-Agent": "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 "
                    "(KHTML, like Gecko) Chrome/124 Safari/537.36",
      "Accept": "application/json", "Content-Type": "application/json"}

# ---------------------------------------------------------------- fair list
# (company, source, config). source=None means no public feed we can read;
# the Companies tab still lists it with its careers URL so nothing is dropped.
WD = "workday"
FAIR = [
    ("Abbott", WD, ("abbott", 5, "abbottcareers")),
    ("Accenture", WD, ("accenture", 103, "AccentureCareers")),
    ("Air Products", WD, ("airproducts", 5, "AP0001")),
    ("Amsted Industries", None, "https://www.amsted.com/careers"),
    ("Apple", "simplify", "https://jobs.apple.com/en-us/search?team=internships-STDNT-INTRN+apple-students-STDNT"),
    ("Arup", None, "https://www.arup.com/careers/"),
    ("ASML", None, "https://www.asml.com/en/careers/find-your-job"),
    ("Bank of America", WD, ("ghr", 1, "Lateral-US")),
    ("Bechtel", None, "https://jobs.bechtel.com/"),
    ("Blue Origin", WD, ("blueorigin", 5, "BlueOrigin")),
    ("Boston Scientific", None, "https://jobs.bostonscientific.com/"),
    ("Burns & McDonnell", None, "https://www.burnsmcd.com/careers"),
    ("Cadence Design Systems", WD, ("cadence", 1, "External_Careers")),
    ("Capital One", WD, ("capitalone", 12, "Capital_One")),
    ("Cargill", None, "https://careers.cargill.com/en"),
    ("Caterpillar", WD, ("cat", 5, "CaterpillarCareers")),
    ("Constellation Energy", None, "https://www.constellationenergy.com/our-company/careers.html"),
    ("Danaher", WD, ("danaher", 1, "DanaherJobs")),
    ("Datadog", "greenhouse", "datadog"),
    ("Delta Air Lines", None, "https://www.delta.com/us/en/careers/overview"),
    ("Digital Realty", "oracle", ("hdep.fa.us2.oraclecloud.com", "CX")),
    ("Eaton", None, "https://www.eaton.com/us/en-us/company/careers.html"),
    ("Ecolab", WD, ("ecolab", 1, "Ecolab_External")),
    ("Edwards Lifesciences", None, "https://www.edwards.com/careers"),
    ("Exponent", None, "https://www.exponent.com/careers"),
    ("ExxonMobil", None, "https://jobs.exxonmobil.com/"),
    ("Fish & Richardson", None, "https://www.fr.com/careers/"),
    ("Ford Motor Company", None, "https://www.careers.ford.com/"),
    ("Freeport-McMoRan", None, "https://jobs.fcx.com/"),
    ("GE HealthCare", WD, ("gehc", 5, "GEHC_ExternalSite")),
    ("General Motors", WD, ("generalmotors", 5, "Careers_GM")),
    ("Gilbane Building Co.", None, "https://www.gilbaneco.com/careers/"),
    ("GlobalFoundries", WD, ("globalfoundries", 1, "External")),
    ("Goldman Sachs", "simplify", "https://higher.gs.com/"),
    ("H-E-B", None, "https://careers.heb.com/"),
    ("Granite Construction", WD, ("granite", 1, "careers")),
    ("HDR", None, "https://www.hdrinc.com/careers"),
    ("Henkel", None, "https://www.henkel.com/careers"),
    ("Hensel Phelps", None, "https://www.henselphelps.com/careers/"),
    ("Hewlett Packard Enterprise", WD, ("hpe", 5, "Jobsathpe")),
    ("Honeywell", "oracle", ("ibqbjb.fa.ocs.oraclecloud.com", "CX_1")),
    ("Jabil", WD, ("jabil", 5, "Jabil_Careers")),
    ("Jacobs", None, "https://careers.jacobs.com/"),
    ("Kiewit", None, "https://www.kiewit.com/careers/"),
    ("KLA", WD, ("kla", 1, "Search")),
    ("L'Oreal", None, "https://careers.loreal.com/en_US/jobs"),
    ("Lam Research", None, "https://careers.lamresearch.com/careers"),
    ("Lenovo", None, "https://jobs.lenovo.com/en_US/careers"),
    ("Lockheed Martin", None, "https://www.lockheedmartinjobs.com/"),
    ("Lutron Electronics", "simplify", "https://www.lutron.com/en-US/Company/Pages/Careers/Overview.aspx"),
    ("McDonald's", None, "https://careers.mcdonalds.com/corporate-jobs"),
    ("Medtronic", WD, ("medtronic", 1, "MedtronicCareers")),
    ("Merck & Co.", WD, ("msd", 5, "SearchJobs")),
    ("Micron Technology", WD, ("micron", 1, "External")),
    ("Microsoft", "simplify", "https://careers.microsoft.com/v2/global/en/universityinternship"),
    ("MIT Lincoln Laboratory", None, "https://www.ll.mit.edu/careers"),
    ("Mitsubishi Power Americas", None, "https://power.mhi.com/regions/amer/careers"),
    ("Morgan Stanley", "simplify", "https://www.morganstanley.com/people-opportunities/students-graduates"),
    ("Mortenson", None, "https://www.mortenson.com/careers"),
    ("Navy Federal Credit Union", None, "https://www.navyfederal.org/careers/"),
    ("Nintendo of America", None, "https://careers.nintendo.com/"),
    ("Nucor", None, "https://nucor.com/careers"),
    ("onsemi", None, "https://www.onsemi.com/careers"),
    ("Procter & Gamble", WD, ("pg", 5, "1000")),
    ("PG&E", None, "https://www.pge.com/en/about/careers.html"),
    ("Pacific Northwest National Lab", None, "https://www.pnnl.gov/jobs"),
    ("PepsiCo", "simplify", "https://www.pepsicojobs.com/main"),
    ("Qnity Electronics", None, "https://www.qnityelectronics.com/careers"),
    ("Regeneron", WD, ("regeneron", 1, "Careers")),
    ("Relativity Space", "simplify", "https://www.relativityspace.com/careers"),
    ("Rockwell Automation", WD, ("rockwellautomation", 1, "External_Rockwell_Automation")),
    ("Schneider Electric", None, "https://www.se.com/us/en/about-us/careers/overview.jsp"),
    ("Shield AI", "lever", "shieldai"),
    ("SpaceX", "greenhouse", "spacex"),
    ("Texas Instruments", "oracle", ("edbz.fa.us2.oraclecloud.com", "CX")),
    ("Thornton Tomasetti", None, "https://www.thorntontomasetti.com/careers"),
    ("Trimble", WD, ("trimble", 1, "TrimbleCareers")),
    ("Trane Technologies", "simplify", "https://careers.tranetechnologies.com/"),
    ("UL Solutions", "oracle", ("fa-eups-saasfaprod1.fa.ocs.oraclecloud.com", "ULSolutionsCareers")),
    ("Wells Fargo", WD, ("wf", 1, "WellsFargoJobs")),
    ("World Wide Technology", None, "https://www.wwt.com/corporate/who-we-are/careers"),
    ("WSP USA", None, "https://www.wsp.com/en-us/careers"),
]
# Recruiting for grad school / programs, not full-time jobs.
SCHOOLS = ["Break Through Tech", "Carle Illinois College of Medicine", "Carnegie Mellon University",
           "Colorado School of Mines", "Cornell University", "Dartmouth Engineering",
           "Embry-Riddle", "Iowa State", "MIT Engineering", "Northeastern University",
           "NYU Tandon", "Stanford Engineering", "Tufts University", "Virginia Tech"]

# Simplify company_name -> fair name (only fair companies; Simplify is a cross-check for all)
SIMPLIFY_ALIASES = {
    "Accenture": "Accenture", "Apple": "Apple", "Blue Origin": "Blue Origin",
    "Cadence Design Systems": "Cadence Design Systems", "Capital One": "Capital One",
    "Caterpillar": "Caterpillar", "Danaher": "Danaher", "Eaton": "Eaton",
    "ExxonMobil": "ExxonMobil", "General Motors": "General Motors",
    "GlobalFoundries": "GlobalFoundries", "Goldman Sachs": "Goldman Sachs",
    "Hewlett Packard Enterprise": "Hewlett Packard Enterprise",
    "Hewlett Packard": "Hewlett Packard Enterprise", "Honeywell": "Honeywell",
    "KLA": "KLA", "KLA Corporation": "KLA", "Lutron Electronics": "Lutron Electronics",
    "Micron Technology": "Micron Technology", "Microsoft": "Microsoft",
    "Morgan Stanley": "Morgan Stanley", "PepsiCo": "PepsiCo",
    "Procter & Gamble": "Procter & Gamble", "Relativity Space": "Relativity Space",
    "Shield AI": "Shield AI", "SpaceX": "SpaceX", "Texas Instruments": "Texas Instruments",
    "Trane Technologies": "Trane Technologies",
}

# ---------------------------------------------------------------- filters
SEARCHES = ["new grad", "new college graduate", "university graduate", "recent graduate",
            "early career", "entry level", "graduate program", "development program",
            "rotational", "campus", "2027", "associate engineer", "engineer I", "college"]
EARLY = re.compile(r"(new\s*(college\s*)?grad|graduate|university|campus|early[\s-]*career|"
                   r"entry[\s-]*level|recent|class\s*of|rotation|development\s*program|"
                   r"leadership\s*program|\bldp\b|\bedp\b|\btdp\b|2027|\bjr\.?\b|junior|"
                   r"associate\s+(engineer|software|scientist|developer|analyst|consultant|"
                   r"product)|engineer\s*(i|1)\b|\b(i|1)\s*-|\blevel\s*1\b|analyst\s*(i|1)\b|"
                   r"scientist\s*(i|1)\b|developer\s*(i|1)\b|technologist\s*(i|1)\b|"
                   r"college|fresh)", re.I)
NOT_FT = re.compile(r"\b(intern|internship|co-?op|summer|apprentice\w*|fellowship|student\s+worker|"
                    r"part[\s-]*time|seasonal|temporary|placement|ausbildung|dhbw|m/w/d|"
                    r"technician|chef|cook|laborer|operator|assembler|driver|maintenance|"
                    r"warehouse|custodian|janitor|cashier|financial\s+advisor|military|veteran|machinist)\b", re.I)
# Always disqualifying, even inside a "program" title.
SENIOR_HARD = re.compile(r"\b(senior|sr\.?|staff|principal|director|vp|head\s+of|supervisor|"
                         r"postdoc|post-doctoral|experienced|expert|mba)\b", re.I)
# Level/lead/manager words: disqualifying unless the title is clearly a grad program.
SENIOR = re.compile(r"\b(lead|manager|architect|ii|iii|iv|2|3|4)\b", re.I)
PROGRAM_OK = re.compile(r"(program|rotation|graduate|new grad|early career|product manager)", re.I)

DISCIPLINES = [
    ("Product/Program", r"product\s*(manager|management|owner)|program\s*manag|apm\b|technical\s*program"),
    ("AI/ML & Data", r"machine\s*learning|\bml\b|\bai\b|data\s*(scien|engineer|analy)|analytics"),
    ("Software", r"software|developer|swe\b|full[\s-]*stack|front[\s-]*end|back[\s-]*end|cloud|devops|"
                 r"cyber|security\s*engineer|technology\s*(analyst|associate|program)|\bit\b|"
                 r"application|platform|sre|site\s*reliab|tech"),
    ("Electrical/Hardware", r"electric|hardware|asic|fpga|rf\b|circuit|silicon|semiconductor|"
                            r"analog|digital\s*design|verification|validation|signal|power|"
                            r"embedded|firmware|avionics|controls?|photonic|optic|test\s*engineer|"
                            r"product\s*engineer|layout|device|yield|lithograph|wafer|vlsi|dft|"
                            r"design\s*for\s*test|packaging|mixed[\s-]*signal|failure\s*analysis|physical\s*design|"
                            r"\bsoc\b|timing|electrical|computer\s*engineer"),
    ("Finance/Business", r"financ|accounting|consult|sales|marketing|business|strategy|"
                         r"banking|risk|audit|investment|wealth|trade\s*compliance|treasury"),
    ("Mechanical/Mfg/Process", r"mechanic|manufactur|process|industrial|quality|reliability|"
                               r"production|operations|supply|materials|chemical|propulsion|"
                               r"structur|civil|construct|field|plant|project\s*engineer|"
                               r"design\s*engineer|systems?\s*engineer|aerospace|biomedical|automation|"
                               r"scientist|r&d|research"),
    ("Finance/Business", r"analyst|associate"),
]


def discipline(title):
    for name, pat in DISCIPLINES:
        if re.search(pat, title, re.I):
            return name
    return "Other Engineering"


def is_early(title):
    if NOT_FT.search(title):
        return False
    if not EARLY.search(title):
        return False
    if SENIOR_HARD.search(title) or (SENIOR.search(title) and not PROGRAM_OK.search(title)):
        return False
    return True


# ---------------------------------------------------------------- salary
NUM = r"\d{1,3}(?:,\d{3})+(?:\.\d{2})?|\d{2,3}(?:\.\d+)?\s?[kK]|\d{2,3}\.\d{2}|\d{4,6}(?:\.\d{2})?"
RANGE = re.compile(r"(?:\$|USD\s?)\s?(" + NUM + r")\s*(?:USD)?\s*(?:-|–|—|to|and)\s*(?:\$|USD\s?)?\s?(" + NUM + r")"
                   r"(\s*(?:per|/|an|a)\s*(?:hour|hr|year|yr|annum|annually))?", re.I)
SINGLE = re.compile(r"(?:base\s+(?:pay|salary)|salary|pay)[^$]{0,80}\$\s?(" + NUM + r")"
                    r"(\s*(?:per|/|an|a)\s*(?:hour|hr|year|yr))?", re.I)


def _val(s):
    s = s.replace(",", "").strip()
    if s[-1:] in "kK":
        return float(s[:-1]) * 1000
    return float(s)


def _fmt(v, hourly):
    return f"${v:,.2f}/hr" if hourly else f"${v / 1000:,.0f}K"


def extract_salary(text):
    """Return a normalized pay range like '$85K–$120K' or '$28.00/hr–$34.00/hr'."""
    if not text:
        return ""
    text = re.sub(r"<[^>]+>", " ", text).replace("&nbsp;", " ").replace("\\u0024", "$")
    best = None
    for m in RANGE.finditer(text):
        try:
            lo, hi = _val(m.group(1)), _val(m.group(2))
        except ValueError:
            continue
        unit = (m.group(3) or "").lower()
        hourly = "h" in unit or hi < 300
        if hourly and not (10 <= lo <= 300 and lo <= hi <= 300):
            continue
        if not hourly and not (25_000 <= lo <= 600_000 and lo <= hi <= 900_000):
            continue
        # prefer annual ranges over hourly ones, first annual wins
        if best is None or (best[2] and not hourly):
            best = (lo, hi, hourly)
            if not hourly:
                break
    if best:
        lo, hi, h = best
        return _fmt(lo, h) if lo == hi else f"{_fmt(lo, h)}–{_fmt(hi, h)}"
    m = SINGLE.search(text)
    if m:
        try:
            v = _val(m.group(1))
        except ValueError:
            return ""
        h = bool(m.group(2) and "h" in m.group(2).lower()) or v < 300
        if (h and 10 <= v <= 300) or (not h and 25_000 <= v <= 600_000):
            return _fmt(v, h)
    return ""


def mentions_2027(text):
    return bool(re.search(r"(2027|class of '?27|graduat\w*\s+(?:in|by|between)[^.]{0,40}2027)",
                          text or "", re.I))


# ---------------------------------------------------------------- http
def get(url, data=None, tries=3):
    # Workday throttles bursts (429/5xx), so back off and retry instead of
    # silently dropping a posting's pay range.
    body = json.dumps(data).encode() if data is not None else None
    for attempt in range(tries):
        try:
            with urlopen(Request(url, data=body, headers=UA), timeout=TIMEOUT, context=SSL_CTX) as r:
                raw = r.read().decode("utf-8", "replace")
            try:
                return json.loads(raw)
            except ValueError:
                return raw
        except HTTPError as e:
            if e.code not in (429, 500, 502, 503, 504):
                return None
        except Exception:
            pass
        time.sleep(2 * (attempt + 1))
    return None


def row(company, title, location, url, posted, salary="", desc="", source=""):
    return {"company": company, "title": title.strip(), "location": location or "",
            "url": url, "posted": posted or "", "salary": salary, "desc": desc, "source": source}


# ---------------------------------------------------------------- fetchers
def from_workday(company, cfg):
    tenant, wd, site = cfg
    base = f"https://{tenant}.wd{wd}.myworkdayjobs.com"
    api = f"{base}/wday/cxs/{tenant}/{site}"
    hits = {}
    for q in SEARCHES:
        for offset in range(0, 100, 20):
            d = get(f"{api}/jobs", {"appliedFacets": {}, "limit": 20, "offset": offset, "searchText": q})
            posts = (d or {}).get("jobPostings") if isinstance(d, dict) else None
            if not posts:
                break
            for p in posts:
                if is_early(p.get("title", "")):
                    hits[p["externalPath"]] = p
            if len(posts) < 20:
                break

    def detail(path_post):
        path, p = path_post
        d = get(f"{api}{path}")
        info = (d or {}).get("jobPostingInfo", {}) if isinstance(d, dict) else {}
        desc = info.get("jobDescription", "") or ""
        loc = info.get("location") or p.get("locationsText", "")
        country = (info.get("country") or {}).get("descriptor", "")
        if country and country != "United States of America":
            return None
        extra = info.get("additionalLocations") or []
        if extra:
            loc = f"{loc} (+{len(extra)} more)"
        return row(company, p["title"], loc, f"{base}/en-US/{site}{path}",
                   info.get("startDate") or p.get("postedOn", ""), extract_salary(desc), desc, "Workday")

    with ThreadPoolExecutor(8) as ex:
        return [r for r in ex.map(detail, hits.items()) if r]


def from_oracle(company, cfg):
    host, site = cfg
    api = f"https://{host}/hcmRestApi/resources/latest"
    hits = {}
    for q in SEARCHES:
        d = get(f"{api}/recruitingCEJobRequisitions?onlyData=true&expand=requisitionList"
                f"&finder=findReqs;siteNumber={site},keyword={quote(q)},limit=100,sortBy=POSTING_DATES_DESC")
        items = (d or {}).get("items", []) if isinstance(d, dict) else []
        for it in items:
            for r in it.get("requisitionList", []) or []:
                if is_early(r.get("Title", "")) and r.get("PrimaryLocationCountry", "US") == "US":
                    hits[r["Id"]] = r

    def detail(item):
        rid, r = item
        d = get(f"{api}/recruitingCEJobRequisitionDetails?expand=all&onlyData=true"
                f"&finder=ById;Id=%22{rid}%22,siteNumber={site}")
        info = ((d or {}).get("items") or [{}])[0] if isinstance(d, dict) else {}
        desc = " ".join(str(info.get(k) or "") for k in
                        ("ExternalDescriptionStr", "ExternalQualificationsStr",
                         "ExternalResponsibilitiesStr", "CorporateDescriptionStr", "OrganizationDescriptionStr"))
        sal = extract_salary(desc)
        if not sal and info.get("ExternalSalaryMinimum"):
            sal = extract_salary(f"${info['ExternalSalaryMinimum']} - ${info.get('ExternalSalaryMaximum')}")
        url = f"https://{host}/hcmUI/CandidateExperience/en/sites/{site}/job/{rid}"
        return row(company, r["Title"], r.get("PrimaryLocation", ""), url,
                   r.get("PostedDate", ""), sal, desc, "Oracle")

    with ThreadPoolExecutor(8) as ex:
        return list(ex.map(detail, hits.items()))


def from_greenhouse(company, slug):
    d = get(f"https://boards-api.greenhouse.io/v1/boards/{slug}/jobs?content=true")
    out = []
    for j in (d or {}).get("jobs", []) if isinstance(d, dict) else []:
        t = j.get("title", "")
        if not is_early(t):
            continue
        import html
        desc = html.unescape(j.get("content", ""))
        out.append(row(company, t, (j.get("location") or {}).get("name", ""), j.get("absolute_url", ""),
                       (j.get("updated_at") or "")[:10], extract_salary(desc), desc, "Greenhouse"))
    return out


def from_lever(company, slug):
    d = get(f"https://api.lever.co/v0/postings/{slug}?mode=json")
    out = []
    for j in d if isinstance(d, list) else []:
        t = j.get("text", "")
        if not is_early(t):
            continue
        desc = (j.get("descriptionPlain") or "") + " " + (j.get("additionalPlain") or "")
        sal = extract_salary(desc)
        rng = j.get("salaryRange") or {}
        if not sal and rng.get("min"):
            sal = extract_salary(f"${rng['min']} - ${rng.get('max')}")
        ts = j.get("createdAt")
        posted = datetime.fromtimestamp(ts / 1000, ET).strftime("%Y-%m-%d") if ts else ""
        out.append(row(company, t, (j.get("categories") or {}).get("location", ""),
                       j.get("hostedUrl", ""), posted, sal, desc, "Lever"))
    return out


_SIMPLIFY = None


def simplify_rows():
    """Community-curated new-grad list; every entry is already new-grad."""
    global _SIMPLIFY
    if _SIMPLIFY is None:
        d = get(SIMPLIFY_URL)
        _SIMPLIFY = []
        for j in d if isinstance(d, list) else []:
            fair = SIMPLIFY_ALIASES.get(j.get("company_name", ""))
            if not fair or not j.get("active") or not j.get("is_visible"):
                continue
            t = j.get("title", "")
            if NOT_FT.search(t) or SENIOR_HARD.search(t) or \
                    (SENIOR.search(t) and not PROGRAM_OK.search(t)):
                continue
            locs = j.get("locations") or []
            ts = j.get("date_posted") or 0
            _SIMPLIFY.append(row(fair, j.get("title", ""), "; ".join(locs), j.get("url", ""),
                                 datetime.fromtimestamp(ts).strftime("%Y-%m-%d") if ts else "",
                                 "", "", "Simplify"))
    return _SIMPLIFY


def enrich_page_salary(r):
    """Open a posting's public page (Apple, Microsoft, etc.) to read its pay range."""
    if r["salary"] or not r["url"]:
        return r
    page = get(r["url"])
    if isinstance(page, str):
        r["salary"] = extract_salary(page)
        r["desc"] = r["desc"] or page[:200000]
    elif isinstance(page, dict):
        r["salary"] = extract_salary(json.dumps(page))
    return r


FETCH = {"workday": from_workday, "oracle": from_oracle,
         "greenhouse": from_greenhouse, "lever": from_lever}


def sweep():
    results = {}

    def one(entry):
        company, source, cfg = entry
        try:
            return company, FETCH[source](company, cfg) if source in FETCH else []
        except Exception as e:  # one broken site never sinks the sweep
            print(f"  ! {company}: {e}")
            return company, []

    with ThreadPoolExecutor(4) as ex:
        for company, rows in ex.map(one, FAIR):
            results[company] = [r for r in rows if location_ok(r["location"])]
            print(f"  {company:32} {len(results[company])} early-career roles")

    # Simplify cross-check: add any fair-company posting the direct feeds missed
    extra = []
    for r in simplify_rows():
        have = results.setdefault(r["company"], [])
        key = re.sub(r"\W+", "", r["title"].lower())
        if any(re.sub(r"\W+", "", h["title"].lower()) == key for h in have):
            continue
        if location_ok(r["location"]):
            extra.append(r)
    with ThreadPoolExecutor(10) as ex:
        extra = list(ex.map(enrich_page_salary, extra))
    for r in extra:
        results[r["company"]].append(r)
    print(f"  + {len(extra)} roles from the Simplify new-grad list")
    for company, title, loc, url, posted, pay in manual()["roles"]:
        results.setdefault(company, []).append(
            row(company, title, loc, url, posted, extract_salary(pay), title, "Hand-checked"))
    return results


def manual():
    if os.path.exists(MANUAL_PATH):
        with open(MANUAL_PATH) as f:
            return json.load(f)
    return {"roles": [], "notes": {}}


# ---------------------------------------------------------------- output
ROLE_HEADER = ["Company", "Title", "Discipline", "Location", "Salary Range", "Mentions 2027",
               "Posted", "Source", "URL", "Status / Notes"]
CO_HEADER = ["Company", "Open New-Grad Roles", "Disciplines Hiring", "Salary Span (listed)",
             "Roles w/ 2027 Mention", "Careers Page", "Booth Notes"]


def build_tables(results):
    roles, seen = [], set()
    for company, rows in results.items():
        for r in rows:
            k = (company, r["title"].lower(), r["location"].lower())
            if k in seen:
                continue
            seen.add(k)
            roles.append([company, r["title"], discipline(r["title"]), r["location"], r["salary"],
                          "Yes" if mentions_2027(r["desc"] + " " + r["title"]) else "",
                          str(r["posted"])[:10], r["source"], r["url"], ""])
    roles.sort(key=lambda x: (x[0].lower(), x[2], x[1]))

    def span(rs):
        vals = []
        for s in rs:
            for n in re.findall(r"\$([\d,.]+)K", s):
                vals.append(float(n.replace(",", "")))
        return f"${min(vals):,.0f}K–${max(vals):,.0f}K" if vals else ""

    companies, notes = [], manual()["notes"]
    for company, source, cfg in FAIR:
        mine = [r for r in roles if r[0] == company]
        url = cfg if isinstance(cfg, str) and cfg.startswith("http") else (mine[0][8] if mine else "")
        if source == "workday":
            url = f"https://{cfg[0]}.wd{cfg[1]}.myworkdayjobs.com/en-US/{cfg[2]}"
        elif source == "greenhouse":
            url = f"https://boards.greenhouse.io/{cfg}"
        elif source == "lever":
            url = f"https://jobs.lever.co/{cfg}"
        elif source == "oracle":
            url = f"https://{cfg[0]}/hcmUI/CandidateExperience/en/sites/{cfg[1]}"
        disc = sorted({r[2] for r in mine})
        if company in notes:
            note = notes[company]
        elif mine:
            note = ""
        elif source is None:
            note = "Careers site has no public feed — check site / ask at booth"
        else:
            note = "No early-career posting live right now — ask at booth"
        companies.append([company, len(mine), ", ".join(disc), span([r[4] for r in mine]),
                          sum(1 for r in mine if r[5] == "Yes"), url, note])
    for s in SCHOOLS:
        companies.append([s, 0, "", "", 0, "", "University / program booth — grad school, not jobs"])
    return roles, companies


def write_xlsx(roles, companies):
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill
    from openpyxl.utils import get_column_letter
    wb = Workbook()
    for i, (name, header, data, widths) in enumerate([
            ("Roles", ROLE_HEADER, roles, [24, 52, 20, 30, 22, 10, 12, 11, 60, 24]),
            ("Companies", CO_HEADER, companies, [30, 10, 50, 22, 10, 60, 50])]):
        ws = wb.active if i == 0 else wb.create_sheet()
        ws.title = name
        ws.append(header)
        for r in data:
            ws.append(r)
        for c in ws[1]:
            c.font = Font(bold=True, color="FFFFFF")
            c.fill = PatternFill("solid", fgColor="1F2937")
        ws.freeze_panes = "A2"
        ws.auto_filter.ref = ws.dimensions
        for j, w in enumerate(widths, 1):
            ws.column_dimensions[get_column_letter(j)].width = w
    wb.save(OUT_XLSX)


def write_sheet(roles, companies):
    key = os.path.join(HERE, "gcp-service-account.json")
    cfg = os.path.join(HERE, "sheet_config.json")
    if not (os.path.exists(key) and os.path.exists(cfg)):
        return
    import gspread
    from google.oauth2.service_account import Credentials
    gc = gspread.authorize(Credentials.from_service_account_file(
        key, scopes=["https://www.googleapis.com/auth/spreadsheets"]))
    sh = gc.open_by_key(json.load(open(cfg))["sheet_id"])
    for tab, header, data in [("SHPE26 Fair Companies", CO_HEADER, companies),
                              ("SHPE26 Fair Roles", ROLE_HEADER, roles)]:
        try:
            ws = sh.worksheet(tab)
            # keep anything typed in the Status / Notes column across re-runs
            old = {r[8]: r[9] for r in ws.get_all_values()[1:] if len(r) > 9 and r[9]} \
                if tab == "SHPE26 Fair Roles" else {}
            ws.clear()
        except gspread.WorksheetNotFound:
            ws = sh.add_worksheet(tab, rows=len(data) + 10, cols=len(header))
            old = {}
        if old:
            data = [r[:9] + [old.get(r[8], r[9])] for r in data]
        ws.resize(rows=len(data) + 1, cols=len(header))
        ws.update([header] + data, "A1", value_input_option="USER_ENTERED")
        ws.freeze(rows=1)
    print(f"  Google Sheet updated: tabs 'SHPE26 Fair Companies' + 'SHPE26 Fair Roles'")


def main():
    print(f"[swe_fair] {datetime.now(ET):%Y-%m-%d %H:%M} sweeping {len(FAIR)} fair companies")
    roles, companies = build_tables(sweep())
    write_xlsx(roles, companies)
    print(f"[swe_fair] {len(roles)} roles across "
          f"{sum(1 for c in companies if c[1])} companies -> {OUT_XLSX}")
    write_sheet(roles, companies)


if __name__ == "__main__":
    main()
