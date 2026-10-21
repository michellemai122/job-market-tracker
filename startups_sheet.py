#!/usr/bin/env python3
"""
Writes a 'Startups (NYC-SF)' tab into the same Google Sheet: fast-growing
startups HQ'd in NYC or SF, how much they've raised, valuation, investors, and
LIVE open-role counts. Funding/valuation are curated & as-of Aug 2026
(directional — verify before acting); open-role counts are fetched each run.

Full-overwrite of the tab's data each run (it's reference data, not a worklist).
"""
import json, os, ssl
from datetime import datetime
from zoneinfo import ZoneInfo
from concurrent.futures import ThreadPoolExecutor
from urllib.request import Request, urlopen

import gspread
from google.oauth2.service_account import Credentials
try:
    import certifi
    CTX = ssl.create_default_context(cafile=certifi.where())
except ImportError:
    CTX = ssl.create_default_context()

HERE = os.path.dirname(os.path.abspath(__file__))
UA = {"User-Agent": "Mozilla/5.0"}
ET = ZoneInfo("America/New_York")
SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]
TAB = "Startups (NYC-SF)"

# feed = (provider, slug) for live open-role count + careers link, or None.
# "does" = 3 bullets (rendered on separate lines in the cell).
def _b(*bullets):
    return "\n".join("• " + b for b in bullets)

STARTUPS = [
    # ---- SF ----
    ("OpenAI", "SF", _b("Frontier AI lab behind GPT models & ChatGPT", "Developer APIs + Codex coding agents", "~$2B/mo revenue; most-used consumer AI"), "$122B round (Q1 2026)", "~$852B", "SoftBank, Microsoft, Thrive", ("ashby", "openai")),
    ("Anthropic", "SF", _b("Frontier AI lab behind the Claude models", "Claude Code + enterprise AI APIs", "Safety-focused; most valuable private co"), "$65B (May 2026)", "~$965B", "GIC, Coatue", ("greenhouse", "anthropic")),
    ("Anysphere (Cursor)", "SF", _b("AI-native code editor (Cursor)", "Agents that edit whole codebases", "Fastest B2B SaaS ever to $2B ARR"), "raising $2B+", "~$50B", "Thrive, a16z", ("ashby", "cursor")),
    ("Cognition", "SF", _b("Builds Devin, an autonomous AI software engineer", "Agents that write & ship code end-to-end", "Acquired Windsurf"), "$1B+ Series D", "~$26B", "Founders Fund, 8VC", ("ashby", "cognition")),
    ("Scale AI", "SF", _b("Data labeling & evaluation for AI models", "Training data for labs & enterprises", "Major Meta investment"), "Meta investment", "~$29B", "Meta, Accel", ("greenhouse", "scaleai")),
    ("Perplexity", "SF", _b("AI answer engine / conversational search", "Cited, real-time web answers", "Fast-growing consumer + enterprise"), "recent raise", "~$18B", "IVP, NEA, a16z", ("ashby", "perplexity")),
    ("Sierra", "SF", _b("Conversational AI agents for customer experience", "Branded support agents for enterprises", "By Bret Taylor; $0→$200M ARR in 26mo"), "Series", "~$10B", "Greenoaks, Sequoia", ("ashby", "sierra")),
    ("Mercor", "SF", _b("AI marketplace matching experts to AI-training work", "Vets & staffs specialists for labs", "$1M→~$1.5B run-rate in 17 months"), "Series", "~$10B", "Felicis, General Catalyst", ("ashby", "mercor")),
    ("Harvey", "SF", _b("Generative AI for law firms & pro services", "Drafts, reviews & researches legal work", "Backed by Sequoia, OpenAI Startup Fund"), "Series", "~$8B", "Sequoia, Kleiner, GV", ("ashby", "harvey")),
    ("Vercel", "SF", _b("Frontend cloud platform; creator of Next.js", "Hosts & deploys modern web apps", "v0 AI app-builder"), "Series E", "~$9B", "Accel, GV", ("greenhouse", "vercel")),
    ("Replit", "SF", _b("Browser-based AI coding environment", "Build & deploy apps from a prompt (Agent)", "~$525M ARR"), "Series", "~$3B", "a16z, Coatue", ("ashby", "replit")),
    ("Physical Intelligence", "SF", _b("Foundation models for robotics", "General-purpose control across robot hardware", "Founders from Google/Stanford robotics"), "Series", "~$2.4B", "Thrive, Lux, Bond", ("ashby", "physicalintelligence")),
    ("Baseten", "SF", _b("Infra for deploying & serving ML models fast", "Inference optimization for AI apps", "Series D growth-stage"), "Series D", "~$2.1B", "IVP, Spark, Greenoaks", ("ashby", "baseten")),
    ("Decagon", "SF", _b("AI customer-support agents for enterprises", "Resolves tickets & chat autonomously", "Backed by a16z, Accel"), "Series C", "~$1.5B", "a16z, Accel, Bond", ("ashby", "decagon")),
    # ---- NYC ----
    ("Ramp", "NYC", _b("Corporate cards + spend management", "AI automates finance & expense workflows", "~$22B valuation"), "Series", "~$22B", "Founders Fund, Thrive", ("ashby", "ramp")),
    ("ElevenLabs", "NYC / London", _b("AI voice synthesis & audio models", "Text-to-speech, dubbing, voice agents", "Hit $2B ARR"), "$500M Series D", "~$11B", "a16z, ICONIQ", ("ashby", "elevenlabs")),
    ("AlphaSense", "NYC", _b("AI market-intelligence & research platform", "Searches filings, transcripts, expert calls", "Used across finance & corporates"), "$350M", "~$7.5B", "Viking, BDT, Goldman", ("greenhouse", "alphasense")),
    ("Clay", "NYC", _b("GTM data enrichment + AI outbound", "Automates prospecting & lead research", "~$3B valuation"), "Series", "~$3B", "Sequoia, Meritech", None),
    ("Runway", "NYC", _b("Generative AI for video & film", "Text/image-to-video models (Gen series)", "Backed by Nvidia, General Catalyst"), "Series", "~$3B+", "General Catalyst, Nvidia", None),
    ("Hebbia", "NYC", _b("AI document search & analysis for finance", "Answers over massive doc sets (Matrix)", "Backed by a16z, Index"), "Series B", "~$700M+", "a16z, Index", None),
    ("Modal Labs", "NYC", _b("Serverless GPU / compute for AI workloads", "Run code & models without infra setup", "$355M raised"), "$355M", "n/a", "Lux, Redpoint", ("ashby", "modal")),
]

HEADER = ["Startup", "HQ", "What they do", "Raised (latest)", "Valuation", "Key investors", "Open roles", "Careers"]


def live_count_and_url(feed):
    if not feed:
        return "—", ""
    prov, slug = feed
    if prov == "greenhouse":
        api = f"https://boards-api.greenhouse.io/v1/boards/{slug}/jobs"
        url = f"https://boards.greenhouse.io/{slug}"
    else:
        api = f"https://api.ashbyhq.com/posting-api/job-board/{slug}"
        url = f"https://jobs.ashbyhq.com/{slug}"
    try:
        with urlopen(Request(api, headers=UA), timeout=12, context=CTX) as r:
            d = json.loads(r.read().decode("utf-8", "replace"))
        return len(d.get("jobs", [])), url
    except Exception:
        return "?", url


def main():
    with ThreadPoolExecutor(max_workers=10) as ex:
        live = list(ex.map(live_count_and_url, [s[6] for s in STARTUPS]))

    rows = [HEADER]
    for (name, hq, does, raised, val, inv, _), (cnt, url) in zip(STARTUPS, live):
        rows.append([name, hq, does, raised, val, inv, cnt, url])

    creds = Credentials.from_service_account_file(os.path.join(HERE, "gcp-service-account.json"), scopes=SCOPES)
    sh = gspread.authorize(creds).open_by_key(json.load(open(os.path.join(HERE, "sheet_config.json")))["sheet_id"])
    try:
        ws = sh.worksheet(TAB)
        ws.clear()
    except gspread.WorksheetNotFound:
        ws = sh.add_worksheet(TAB, rows=max(50, len(rows) + 5), cols=len(HEADER))

    ws.update(rows, "A1", value_input_option="USER_ENTERED")
    ws.format(f"A1:{chr(64+len(HEADER))}1", {"textFormat": {"bold": True}})
    ws.format(f"C2:C{len(rows)}", {"wrapStrategy": "WRAP"})           # bullets on their own lines
    ws.format(f"A2:A{len(rows)}", {"textFormat": {"bold": True}})     # startup name bold
    ws.columns_auto_resize(0, len(HEADER)) if hasattr(ws, "columns_auto_resize") else None
    ws.freeze(rows=1)
    stamp = datetime.now(ET).strftime("%Y-%m-%d %H:%M ET")
    ws.update([[f"Funding curated & directional (verify) · open-role counts live as of {stamp}"]],
              f"A{len(rows)+2}")
    print(f"wrote '{TAB}' tab: {len(STARTUPS)} startups (live open-role counts).")


if __name__ == "__main__":
    main()
