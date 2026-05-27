import os, json, time, re, hashlib, sys, urllib.parse
from datetime import datetime, timezone
from typing import Optional
from dataclasses import dataclass, field

os.system("pip install requests openpyxl google-auth google-auth-oauthlib google-auth-httplib2 google-api-python-client 2>/dev/null")

import requests
from openpyxl import Workbook, load_workbook

FIRECRAWL_API_KEY = os.environ.get("FIRECRAWL_API_KEY", "")
GOOGLE_OAUTH_CLIENT_ID = os.environ.get("GOOGLE_OAUTH_CLIENT_ID", "")
GOOGLE_OAUTH_CLIENT_SECRET = os.environ.get("GOOGLE_OAUTH_CLIENT_SECRET", "")

EXCEL_FILE = "opportunities.xlsx"
MEMORY_FILE = "bot_memory.json"
REPORT_DIR = "reports"

@dataclass
class Opportunity:
    id: str = ""
    title: str = ""
    url: str = ""
    category: str = ""
    description: str = ""
    profit_per_day: str = ""
    profit_per_week: str = ""
    profit_per_month: str = ""
    effort_level: str = ""
    source: str = ""
    found_date: str = ""
    tags: list = field(default_factory=list)
    status: str = "new"

def load_excel():
    if not os.path.exists(EXCEL_FILE):
        wb = Workbook()
        ws = wb.active
        ws.title = "Opportunities"
        headers = ["ID", "Title", "URL", "Category", "Description", "Profit/Day", "Profit/Week", "Profit/Month", "Effort", "Source", "Found Date", "Tags", "Status"]
        ws.append(headers)
        for col, w in [("A",8),("B",40),("C",50),("D",20),("E",60),("F",15),("G",15),("H",15),("I",12),("J",20),("K",20),("L",30),("M",10)]:
            ws.column_dimensions[col].width = w
        wb.save(EXCEL_FILE)
        return wb, ws
    return load_workbook(EXCEL_FILE), load_workbook(EXCEL_FILE).active

def opportunity_exists(ws, url: str) -> bool:
    for row in ws.iter_rows(min_row=2, values_only=True):
        if len(row) > 2 and row[2] == url:
            return True
    return False

def extract_ddg_url(u: str) -> str:
    """Extract actual URL from DuckDuckGo redirect URL"""
    if u.startswith("//duckduckgo.com/l/") or u.startswith("http://duckduckgo.com/l/") or u.startswith("https://duckduckgo.com/l/"):
        parsed = urllib.parse.urlparse(u)
        qs = urllib.parse.parse_qs(parsed.query)
        return qs.get("uddg", [u])[0]
    return u

def search_firecrawl(query: str) -> list:
    if not FIRECRAWL_API_KEY:
        print("[Firecrawl] No API key set")
        return []
    try:
        resp = requests.post("https://api.firecrawl.dev/v1/search",
            headers={"Authorization": f"Bearer {FIRECRAWL_API_KEY}", "Content-Type": "application/json"},
            json={"query": query, "maxResults": 10, "scrapeOptions": {"formats": ["markdown"]}}, timeout=30)
        data = resp.json()
        results = [{"title": i.get("title",""), "url": i.get("url",""), "description": (i.get("description","") or i.get("markdown",""))[:300]} for i in data.get("data",[])]
        print(f"[Firecrawl] {len(results)} results")
        return results
    except Exception as e:
        print(f"[Firecrawl] Error: {e}")
        return []

def search_ddg(query: str) -> list:
    try:
        resp = requests.get("https://html.duckduckgo.com/html/", params={"q": query},
            headers={"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36"}, timeout=15)
        import html
        links = re.findall(r'<a[^>]+class="result__a"[^>]+href="([^"]+)"[^>]*>(.*?)</a>', resp.text, re.DOTALL)
        snippets = re.findall(r'<a[^>]+class="result__snippet"[^>]*>(.*?)</a>', resp.text, re.DOTALL)
        if not links:
            # Check if DDG is showing a block/error page
            if "captcha" in resp.text.lower() or "blocked" in resp.text.lower():
                print("[DuckDuckGo] Blocked (captcha)")
                return []
            print(f"[DuckDuckGo] No links found ({len(resp.text)} bytes)")
            return []
        results = [{"title": html.unescape(re.sub(r'<[^>]+>', '', t)).strip(),
                  "url": extract_ddg_url(u),
                  "description": html.unescape(re.sub(r'<[^>]+>', '', snippets[i] if i < len(snippets) else "")).strip()[:300]}
                for i, (u, t) in enumerate(links[:10])]
        print(f"[DuckDuckGo] {len(links)} links")
        return results
    except Exception as e:
        print(f"[DuckDuckGo] Error: {e}")
        return []

def search_bing(query: str) -> list:
    """Bing HTML search as fallback"""
    try:
        resp = requests.get("https://www.bing.com/search", params={"q": query},
            headers={"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36", "Accept-Language": "en-US"}, timeout=15)
        import html
        links = re.findall(r'<a[^>]+href="(https?://[^"]+)"[^>]*><h2>(.*?)</h2>', resp.text, re.DOTALL)
        if not links:
            links = re.findall(r'<cite[^>]*>(.*?)</cite>', resp.text, re.DOTALL)
            titles = re.findall(r'<h2[^>]*><a[^>]+href="(https?://[^"]+)"[^>]*>(.*?)</a></h2>', resp.text, re.DOTALL)
            links = [(h, t) for h, t in titles] if titles else []
        snippets = re.findall(r'<p[^>]*>(.*?)</p>', resp.text, re.DOTALL)
        results = []
        for i, (u, t) in enumerate(links[:10]):
            desc = html.unescape(re.sub(r'<[^>]+>', '', snippets[i] if i < len(snippets) else "")).strip()[:300] if i < len(snippets) else ""
            results.append({"title": html.unescape(re.sub(r'<[^>]+>', '', t)).strip()[:200], "url": u, "description": desc})
        print(f"[Bing] {len(results)} results")
        return results
    except Exception as e:
        print(f"[Bing] Error: {e}")
        return []

def search_all(query: str) -> list:
    seen = set()
    results = []
    for fn in [search_firecrawl, search_ddg, search_bing]:
        try:
            for r in fn(query):
                if r["url"] and r["url"] not in seen:
                    seen.add(r["url"])
                    results.append(r)
        except Exception as e:
            print(f"[Search] {fn.__name__} error: {e}")
        time.sleep(1)
    return results

CLASSIFY_KEYWORDS = [
    (["faucet","free crypto","claim","earn satoshi","btc faucet","eth faucet","crypto faucet","bitcoin faucet"], "Crypto Faucet", ["faucet","crypto","free"]),
    (["airdrop","token distribution","free token","claim airdrop","crypto airdrop"], "Airdrop", ["airdrop","crypto","free"]),
    (["ptc","paid-to-click","bux","click ads","get paid to","paidtoclick","earn per click"], "PTC / GPT", ["ptc","gpt","click"]),
    (["survey","paid survey","market research","opinion","survey site","paid surveys"], "Paid Surveys", ["survey","research"]),
    (["cashback","cash back","rebate","shopping","cashback site"], "Cashback", ["cashback","shopping"]),
    (["affiliate","referral","refer","affiliate program"], "Affiliate", ["affiliate","referral"]),
    (["arbitrage","flip","resell","flipping","buy low sell high"], "Arbitrage", ["arbitrage","flipping"]),
    (["stake","staking","defi","yield","lend","apr","liquidity","pool"], "DeFi / Staking", ["defi","staking","yield"]),
    (["play to earn","p2e","gamefi","nft game","play-to-earn","crypto game"], "Play-to-Earn", ["p2e","gaming","nft"]),
    (["mining","cloud mining","hash","mine","miner","crypto mining"], "Mining", ["mining","crypto"]),
    (["trading bot","auto trade","signal","copy trade","trading platform"], "Trading", ["trading","bot","automation"]),
    (["micro task","microtask","captcha","data entry","freelance","micro job"], "Micro Tasks", ["micro-task","freelance"]),
    (["browser automation","auto earn","auto claim","auto bot","automation bot"], "Automation", ["automation","bot"]),
    (["earn crypto","free crypto","get crypto","crypto earn","crypto reward"], "Crypto Earnings", ["crypto","earn"]),
]

def classify(item: dict) -> Optional[Opportunity]:
    t, u, d = item.get("title",""), item.get("url",""), item.get("description","")
    c = f"{t} {d}".lower()
    for keywords, cat, tags in CLASSIFY_KEYWORDS:
        if any(k in c for k in keywords):
            return Opportunity(id=hashlib.md5(f"{t}{u}{time.time()}".encode()).hexdigest()[:12],
                title=t[:100], url=u[:300], category=cat, description=d[:500],
                profit_per_day="Unknown", profit_per_week="Unknown", profit_per_month="Unknown",
                effort_level="Unknown", source="web_search",
                found_date=datetime.now(timezone.utc).strftime("%Y-%m-%d %H:%M:%S UTC"),
                tags=tags, status="new")
    # Catch-all: if it has a money-related word, tag as General
    money_words = ["earn","money","income","profit","passive","free","reward","bonus","cash","pay","withdraw","crypto","bitcoin","eth","btc"]
    if any(k in c for k in money_words):
        return Opportunity(id=hashlib.md5(f"{t}{u}{time.time()}".encode()).hexdigest()[:12],
            title=t[:100], url=u[:300], category="General", description=d[:500],
            profit_per_day="Unknown", profit_per_week="Unknown", profit_per_month="Unknown",
            effort_level="Unknown", source="web_search",
            found_date=datetime.now(timezone.utc).strftime("%Y-%m-%d %H:%M:%S UTC"),
            tags=["money","earn"], status="new")
    return None
    return Opportunity(id=hashlib.md5(f"{t}{u}{time.time()}".encode()).hexdigest()[:12],
        title=t[:100], url=u[:300], category=cat, description=d[:500],
        profit_per_day="Unknown", profit_per_week="Unknown", profit_per_month="Unknown",
        effort_level="Unknown", source="web_search",
        found_date=datetime.now(timezone.utc).strftime("%Y-%m-%d %H:%M:%S UTC"),
        tags=tags, status="new")

QUERIES = [
    "best crypto faucets 2026 free bitcoin ethereum",
    "new crypto airdrops 2026 free tokens claim",
    "paid to click sites that pay instantly 2026",
    "best GPT sites earn money free 2026",
    "passive income crypto staking defi 2026",
    "play to earn crypto games 2026 free",
    "browser automation earn crypto free",
    "auto claim crypto faucet bot",
    "free bitcoin mining sites 2026 no deposit",
    "earn crypto by completing tasks 2026",
    "best cashback apps 2026 free money",
    "micro task sites pay crypto 2026",
    "affiliate programs crypto 2026 high paying",
    "arbitrage opportunities crypto 2026",
    "trading bots free crypto 2026",
    "solana airdrop 2026 claim free",
    "free eth faucet 2026 working",
    "binance earn free crypto 2026",
    "new ways to earn money online 2026 free",
    "AI tools that pay money 2026",
    "data entry jobs from home 2026 pay daily",
    "referral programs that pay instantly 2026",
    "crypto savings account high apr 2026",
    "nft play to earn 2026 free mint",
    "telegram bot earn crypto 2026",
]

def write_report(wb, total_new, categories):
    os.makedirs(REPORT_DIR, exist_ok=True)
    date_str = datetime.now(timezone.utc).strftime("%Y-%m-%d")
    path = os.path.join(REPORT_DIR, f"report_{date_str}.md")
    ws = wb.active
    total = max(0, ws.max_row - 1)
    with open(path, "w", encoding="utf-8") as f:
        f.write(f"# Opportunity Hunter Report - {date_str}\n\n")
        f.write(f"**Total tracked:** {total}  |  **New today:** {total_new}\n\n")
        f.write("## Categories\n\n")
        for cat, cnt in sorted(categories.items(), key=lambda x: -x[1]):
            f.write(f"- **{cat}**: {cnt}\n")
        f.write("\n## Latest 20\n\n| Title | Category | Profit/Month | Link |\n|---|---|---|---|\n")
        for i, row in enumerate(ws.iter_rows(min_row=2, values_only=True)):
            if i >= 20: break
            if len(row) > 12:
                f.write(f"| {str(row[1] or '')[:40]} | {row[3]} | {row[7]} | {str(row[2] or '')[:50]} |\n")
    return path

def write_google_doc(report_path):
    if not GOOGLE_OAUTH_CLIENT_ID or not GOOGLE_OAUTH_CLIENT_SECRET:
        print("[Google Docs] No OAuth - skipping")
        return False
    try:
        from google.auth.transport.requests import Request as GoogleRequest
        from google.oauth2.credentials import Credentials
        from googleapiclient.discovery import build
        refresh_token = os.environ.get("GOOGLE_REFRESH_TOKEN")
        if not refresh_token:
            print("[Google Docs] No GOOGLE_REFRESH_TOKEN - skipping")
            return False
        creds = Credentials(token=None, client_id=GOOGLE_OAUTH_CLIENT_ID,
            client_secret=GOOGLE_OAUTH_CLIENT_SECRET,
            refresh_token=refresh_token,
            token_uri="https://oauth2.googleapis.com/token")
        creds.refresh(GoogleRequest())
        docs = build("docs", "v1", credentials=creds)
        date_str = datetime.now(timezone.utc).strftime("%Y-%m-%d")
        doc = docs.documents().create(body={"title": f"Opportunity Hunter Report - {date_str}"}).execute()
        doc_id = doc["documentId"]
        with open(report_path, encoding="utf-8") as f:
            content = f.read()
        docs.documents().batchUpdate(documentId=doc_id,
            body={"requests": [{"insertText": {"endOfSegmentLocation": {}, "text": content}}]}).execute()
        print(f"[Google Docs] Created: https://docs.google.com/document/d/{doc_id}")
        return True
    except Exception as e:
        print(f"[Google Docs] {e}")
        return False

def main():
    mem = json.load(open(MEMORY_FILE)) if os.path.exists(MEMORY_FILE) else {"runs":0,"total_found":0,"categories_found":{},"last_run":None,"learning":[]}
    mem["runs"] += 1
    wb, ws = load_excel()
    total_new, categories = 0, {}

    for i, q in enumerate(QUERIES):
        sys.stdout.write(f"[{i+1}/{len(QUERIES)}] {q[:50]}... ")
        sys.stdout.flush()
        items = search_all(q)
        sys.stdout.write(f"{len(items)} items... ")
        sys.stdout.flush()
        new_for_query = 0
        for item in items:
            opp = classify(item)
            if opp and not opportunity_exists(ws, opp.url):
                ws.append([opp.id, opp.title, opp.url, opp.category, opp.description,
                    opp.profit_per_day, opp.profit_per_week, opp.profit_per_month,
                    opp.effort_level, opp.source, opp.found_date, ",".join(opp.tags), opp.status])
                new_for_query += 1
                total_new += 1
                categories[opp.category] = categories.get(opp.category, 0) + 1
                mem["total_found"] += 1
        print(f"+{new_for_query}")
        mem["learning"].append({"date": datetime.now(timezone.utc).isoformat(), "query": q, "results": len(items), "new_opps": new_for_query})
        if len(mem["learning"]) > 100: mem["learning"] = mem["learning"][-100:]
        time.sleep(2)

    wb.save(EXCEL_FILE)
    mem["last_run"] = datetime.now(timezone.utc).isoformat()
    for cat, cnt in categories.items():
        mem["categories_found"][cat] = mem["categories_found"].get(cat, 0) + cnt
    json.dump(mem, open(MEMORY_FILE, "w"), indent=2)

    report_path = write_report(wb, total_new, categories)
    print(f"Run #{mem['runs']}: +{total_new} new, {max(0, ws.max_row-1)} total")
    write_google_doc(report_path)
    print(f"[{datetime.now(timezone.utc).isoformat()}] Done!")

if __name__ == "__main__":
    main()
