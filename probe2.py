#!/usr/bin/env python3
"""Probe hand-picked ATS candidates for high-pay firms the auto-discovery missed."""
import json, ssl
from urllib.request import Request, urlopen
try:
    import certifi
    CTX = ssl.create_default_context(cafile=certifi.where())
except ImportError:
    CTX = ssl.create_default_context()
UA = {"User-Agent": "Mozilla/5.0"}

def get(url):
    try:
        with urlopen(Request(url, headers=UA), timeout=12, context=CTX) as r:
            return json.loads(r.read().decode("utf-8", "replace"))
    except Exception as e:
        return None

def gh(s):
    d = get(f"https://boards-api.greenhouse.io/v1/boards/{s}/jobs")
    return len(d["jobs"]) if isinstance(d, dict) and d.get("jobs") else 0
def lever(s):
    d = get(f"https://api.lever.co/v0/postings/{s}?mode=json")
    return len(d) if isinstance(d, list) and d else 0
def ashby(s):
    d = get(f"https://api.ashbyhq.com/posting-api/job-board/{s}")
    return len(d["jobs"]) if isinstance(d, dict) and d.get("jobs") else 0
def sr(s):
    d = get(f"https://api.smartrecruiters.com/v1/companies/{s}/postings")
    return len(d["content"]) if isinstance(d, dict) and d.get("content") else 0

FN = {"greenhouse": gh, "lever": lever, "ashby": ashby, "smartrecruiters": sr}

# (company, provider, slug) candidates
CAND = [
    ("Citadel", "greenhouse", "citadel"), ("Citadel", "greenhouse", "citadelsecurities"),
    ("Citadel", "smartrecruiters", "Citadel"), ("Citadel", "ashby", "citadel"),
    ("Two Sigma", "greenhouse", "twosigma"), ("Two Sigma", "lever", "twosigma"),
    ("Two Sigma", "ashby", "twosigma"),
    ("Hudson River Trading", "greenhouse", "wearehudsonrivertrading"),
    ("Hudson River Trading", "greenhouse", "hudsonrivertrading"),
    ("Hudson River Trading", "greenhouse", "hrt"), ("Hudson River Trading", "lever", "hudson-trading"),
    ("Optiver", "greenhouse", "optiver"), ("Optiver", "smartrecruiters", "Optiver"),
    ("Optiver", "lever", "optiver"), ("Optiver", "ashby", "optiver"),
    ("SIG", "greenhouse", "susquehanna"), ("SIG", "greenhouse", "sig"),
    ("SIG", "lever", "sig"), ("SIG", "smartrecruiters", "SIG"),
    ("Millennium", "greenhouse", "millennium"), ("Millennium", "ashby", "millennium"),
    ("Balyasny", "greenhouse", "balyasny"), ("Balyasny", "greenhouse", "balyasnyassetmanagement"),
    ("Balyasny", "greenhouse", "bam"), ("Balyasny", "ashby", "balyasny"),
    ("Balyasny", "lever", "balyasny"),
    ("DE Shaw", "greenhouse", "deshaw"), ("DE Shaw", "greenhouse", "deshawresearch"),
    ("DE Shaw", "lever", "deshaw"),
    ("Bridgewater", "greenhouse", "bridgewater"), ("Bridgewater", "greenhouse", "bridgewaterassociates"),
    ("Bridgewater", "lever", "bridgewater"),
    ("Jane Street", "greenhouse", "janestreet"),  # confirm
    # AI labs / high-pay startups worth adding
    ("Anthropic", "greenhouse", "anthropic"), ("OpenAI", "ashby", "openai"),
    ("OpenAI", "greenhouse", "openai"), ("Anysphere", "ashby", "anysphere"),
    ("Anysphere Cursor", "ashby", "cursor"), ("Ramp", "ashby", "ramp"),
    ("Rippling", "greenhouse", "rippling"), ("Mistral", "ashby", "mistralai"),
    ("Jump Trading", "greenhouse", "jumptrading"),  # confirm
]

seen = {}
for co, prov, slug in CAND:
    n = FN[prov](slug)
    tag = "OK " if n else "-- "
    print(f"  {tag} {co:<24} {prov:<15} {slug:<24} {n}")
    if n and co not in seen:
        seen[co] = {"company": co, "provider": prov, "slug": slug, "count": n}

print("\nHITS:", ", ".join(seen.keys()))
with open("probe2_hits.json", "w") as f:
    json.dump(list(seen.values()), f, indent=2)
