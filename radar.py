#!/usr/bin/env python3
"""
Startup Radar — a digest of fast-growing startups: what they do, who runs
eng/product, how much they've raised, valuation, and how aggressively they're
hiring (live open-role count pulled from their ATS feed). Outputs data/radar.md.

Funding/valuation figures are curated and marked with an as-of date; the
software layer can't of itself verify a private valuation, so treat the money
figures as directional and confirm the ones you're about to act on. Hiring
counts ARE live (fetched each run) and the new-grad/AI/FDE subset already flows
into jobs.xlsx.
"""
import json, os, ssl
from datetime import datetime
from zoneinfo import ZoneInfo
from concurrent.futures import ThreadPoolExecutor
from urllib.request import Request, urlopen
try:
    import certifi
    CTX = ssl.create_default_context(cafile=certifi.where())
except ImportError:
    CTX = ssl.create_default_context()

HERE = os.path.dirname(os.path.abspath(__file__))
ET = ZoneInfo("America/New_York")
UA = {"User-Agent": "Mozilla/5.0"}
AS_OF = "Aug 2026 (verify fast-moving figures before acting)"

# Curated intel. 'feed' = (provider, slug) for live open-role count.
STARTUPS = [
    {"name": "Anthropic", "feed": ("greenhouse", "anthropic"),
     "does": "Frontier AI lab; Claude models + Claude Code",
     "raised": "$65B Series (May 2026)", "valuation": "~$965B (most valuable private co)",
     "investors": "GIC, Coatue, others", "leaders": "Dario & Daniela Amodei (co-founders/CEO/President)"},
    {"name": "OpenAI", "feed": ("ashby", "openai"),
     "does": "Frontier AI lab; ChatGPT, GPT models, Codex",
     "raised": "$122B round (Q1 2026) — largest private round ever", "valuation": "~$852B",
     "investors": "SoftBank, Microsoft, Thrive", "leaders": "Sam Altman (CEO), Greg Brockman (Pres.)"},
    {"name": "Anysphere (Cursor)", "feed": ("ashby", "cursor"),
     "does": "AI code editor; fastest B2B SaaS ever to $2B ARR",
     "raised": "reportedly raising $2B+", "valuation": "~$50B",
     "investors": "Thrive, a16z", "leaders": "Michael Truell (CEO) & co-founders"},
    {"name": "Cognition", "feed": ("ashby", "cognition"),
     "does": "Devin, autonomous AI software engineer; acquired Windsurf",
     "raised": "Series (2025)", "valuation": "~$10B",
     "investors": "Founders Fund, 8VC", "leaders": "Scott Wu (CEO)"},
    {"name": "Sierra", "feed": ("ashby", "sierra"),
     "does": "Conversational AI agents for enterprise CX; $0→$200M ARR in 26 mo",
     "raised": "Series (2025)", "valuation": "~$10B",
     "investors": "Greenoaks, Sequoia", "leaders": "Bret Taylor (CEO), Clay Bavor (co-founder)"},
    {"name": "Mercor", "feed": ("ashby", "mercor"),
     "does": "AI hiring / expert-data marketplace; $1M→~$1.5B run-rate in 17 mo",
     "raised": "Series (2026)", "valuation": "~$10B (reported)",
     "investors": "Felicis, General Catalyst", "leaders": "Brendan Foody (CEO), Adarsh Hiremath, Surya Midha"},
    {"name": "ElevenLabs", "feed": ("ashby", "elevenlabs"),
     "does": "AI voice / audio models; hit $2B ARR by Feb 2026",
     "raised": "Series C+", "valuation": "~$6.6B (reported)",
     "investors": "a16z, ICONIQ", "leaders": "Mati Staniszewski (CEO), Piotr Dąbkowski (CTO)"},
    {"name": "Lovable", "feed": ("ashby", "lovable"),
     "does": "AI app builder (vibe-coding); $0→$500M ARR in ~14 mo",
     "raised": "Series A ~$200M", "valuation": "~$1.8B+",
     "investors": "Accel, Creandum", "leaders": "Anton Osika (CEO)"},
    {"name": "Replit", "feed": ("ashby", "replit"),
     "does": "AI-native cloud dev environment / agent; ~$525M ARR",
     "raised": "Series (2025)", "valuation": "~$3B",
     "investors": "a16z, Coatue", "leaders": "Amjad Masad (CEO)"},
    {"name": "Decagon", "feed": ("ashby", "decagon"),
     "does": "AI customer-support agents for enterprise",
     "raised": "Series C (2025)", "valuation": "~$1.5B",
     "investors": "a16z, Accel, Bond", "leaders": "Jesse Zhang (CEO), Ashwin Sreenivas (CTO)"},
    {"name": "Harvey", "feed": ("ashby", "harvey"),
     "does": "Generative AI for legal / professional services",
     "raised": "Series (2025/26)", "valuation": "~$5B+",
     "investors": "Sequoia, Kleiner Perkins, GV", "leaders": "Winston Weinberg (CEO), Gabriel Pereyra (Pres.)"},
    {"name": "Physical Intelligence", "feed": ("ashby", "physicalintelligence"),
     "does": "Foundation models for robotics",
     "raised": "Series (2024/25)", "valuation": "~$2.4B",
     "investors": "Thrive, Lux, Bond", "leaders": "Karol Hausman (CEO), Sergey Levine, Chelsea Finn"},
    {"name": "Baseten", "feed": ("ashby", "baseten"),
     "does": "Infrastructure for fast model inference / deployment",
     "raised": "Series D (2026)", "valuation": "~$2.1B",
     "investors": "IVP, Spark, Greenoaks", "leaders": "Tuhin Srivastava (CEO)"},
    {"name": "Suno", "feed": ("ashby", "suno"),
     "does": "Generative AI music",
     "raised": "Series (2025/26)", "valuation": "~$2B+",
     "investors": "Lightspeed, Menlo", "leaders": "Mikey Shulman (CEO)"},
    {"name": "Vercel", "feed": ("greenhouse", "vercel"),
     "does": "Frontend cloud (Next.js) + v0 AI builder",
     "raised": "Series E (2024)", "valuation": "~$9B",
     "investors": "Accel, GV, Notable", "leaders": "Guillermo Rauch (CEO)"},
]


def live_count(feed):
    prov, slug = feed
    try:
        if prov == "greenhouse":
            u = f"https://boards-api.greenhouse.io/v1/boards/{slug}/jobs"
        else:
            u = f"https://api.ashbyhq.com/posting-api/job-board/{slug}"
        with urlopen(Request(u, headers=UA), timeout=12, context=CTX) as r:
            d = json.loads(r.read().decode("utf-8", "replace"))
        return len(d.get("jobs", []))
    except Exception:
        return None


def main():
    with ThreadPoolExecutor(max_workers=10) as ex:
        counts = list(ex.map(live_count, [s["feed"] for s in STARTUPS]))
    for s, c in zip(STARTUPS, counts):
        s["open"] = c

    now = datetime.now(ET).strftime("%Y-%m-%d %H:%M ET")
    lines = [f"# Startup Radar — {now}",
             f"_Funding/valuation as of {AS_OF}. Hiring counts are live this run; "
             f"their new-grad / AI-eng / FDE roles already appear in jobs.xlsx._\n",
             "## Snapshot\n",
             "| Startup | What they do | Raised | Valuation | Open roles |",
             "|---|---|---|---|---|"]
    for s in STARTUPS:
        opn = "?" if s["open"] is None else str(s["open"])
        lines.append(f"| **{s['name']}** | {s['does']} | {s['raised']} | {s['valuation']} | {opn} |")

    lines.append("\n## Profiles\n")
    for s in STARTUPS:
        opn = "n/a" if s["open"] is None else f"{s['open']} open roles"
        prov, slug = s["feed"]
        url = (f"https://boards.greenhouse.io/{slug}" if prov == "greenhouse"
               else f"https://jobs.ashbyhq.com/{slug}")
        lines += [f"### {s['name']}  ({opn})",
                  f"- **What:** {s['does']}",
                  f"- **Money:** {s['raised']} · valuation {s['valuation']} · investors: {s['investors']}",
                  f"- **Who to know (eng/product/founders):** {s['leaders']}",
                  f"- **Careers:** {url}\n"]

    out = os.path.join(HERE, "data", "radar.md")
    with open(out, "w") as f:
        f.write("\n".join(lines))
    print("wrote", out)


if __name__ == "__main__":
    main()
