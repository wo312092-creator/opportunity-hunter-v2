import os, json, time, re, hashlib
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

def search_firecrawl(query: str) -> list:
    if not FIRECRAWL_API_KEY:
        return []
    try:
        resp = requests.post("https://api.firecrawl.dev/v1/search",
            headers={"Authorization": f"Bearer {FIRECRAWL_API_KEY}", "Content-Type": "application/json"},
            json={"query": query, "maxResults": 10, "scrapeOptions": {"formats": ["markdown"]}}, timeout=30)
        data = resp.json()
        return [{"title": i.get("title",""), "url": i.get("url",""), "description": (i.get("description","") or i.get("markdown",""))[:300]} for i in data.get("data",[])]
    except Exception as e:
        print(f"[Firecrawl] {e}")
        return []

def search_serp(query: str) -> list:
    try:
        resp = requests.get("https://html.duckduckgo.com/html/", params={"q": query},
            headers={"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36"}, timeout=15)
        import html
        links = re.findall(r'<a[^>]+class="result__a"[^>]+href="([^"]+)"[^>]*>(.*?)</a>', resp.text, re.DOTALL)
        snippets = re.findall(r'<a[^>]+class="result__snippet"[^>]*>(.*?)</a>', resp.text, re.DOTALL)
        return [{"title": html.unescape(re.sub(r'<[^>]+>', '', t)).strip(),
                  "url": u,
                  "description": html.unescape(re.sub(r'<[^>]+>', '', snippets[i] if i < len(snippets) else "")).strip()[:300]}
                for i, (u, t) in enumerate(links[:10])]
    except Exception as e:
        print(f"[DuckDuckGo] {e}")
        return []

def search_all(query: str) -> list:
    seen = set()
    results = []
    for fn in [search_firecrawl, search_serp]:
        for r in fn(query):
            if r["url"] and r["url"] not in seen:
                seen.add(r["url"])
                results.append(r)
        time.sleep(1)
    return results

def classify(item: dict) -> Optional[Opportunity]:
    t, u, d = item.get("title",""), item.get("url",""), item.get("description","")
    c = f"{t} {d}".lower()
    if any(k in c for k in ["faucet","free crypto","claim","earn satoshi","btc faucet","eth faucet"]):
        cat, tags = "Crypto Faucet", ["faucet","crypto","free"]
    elif any(k in c for k in ["airdrop","token distribution","free token"]):
        cat, tags = "Airdrop", ["airdrop","crypto","free"]
    elif any(k in c for k in ["ptc","paid-to-click","bux","click ads","get paid to"]):
        cat, tags = "PTC / GPT", ["ptc","gpt","click"]
    elif any(k in c for k in ["survey","paid survey","market research","opinion"]):
        cat, tags = "Paid Surveys", ["survey","research"]
    elif any(k in c for k in ["cashback","cash back","rebate","shopping"]):
        cat, tags = "Cashback", ["cashback","shopping"]
    elif any(k in c for k in ["affiliate","referral","refer"]):
        cat, tags = "Affiliate", ["affiliate","referral"]
    elif any(k in c for k in ["arbitrage","flip","resell","flipping"]):
        cat, tags = "Arbitrage", ["arbitrage","flipping"]
    elif any(k in c for k in ["stake","staking","defi","yield","lend","apr"]):
        cat, tags = "DeFi / Staking", ["defi","staking","yield"]
    elif any(k in c for k in ["play to earn","p2e","gamefi","nft game","play-to-earn"]):
        cat, tags = "Play-to-Earn", ["p2e","gaming","nft"]
    elif any(k in c for k in ["mining","cloud mining","hash","mine"]):
        cat, tags = "Mining", ["mining","crypto"]
    elif any(k in c for k in ["trading bot","auto trade","signal","copy trade"]):
        cat, tags = "Trading", ["trading","bot","automation"]
    elif any(k in c for k in ["micro task","microtask","captcha","data entry","freelance"]):
        cat, tags = "Micro Tasks", ["micro-task","freelance"]
    elif any(k in c for k in ["browser automation","auto earn","auto claim","auto bot"]):
        cat, tags = "Automation", ["automation","bot"]
    elif any(k in c for k in ["crypto","bitcoin","ethereum","solana","bnb","web3"]):
        cat, tags = ("Crypto Earnings", ["crypto","earn"]) if "free" in c or "earn" in c or "get" in c else ("Crypto General", ["crypto"])
    else:
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
        from google.oauth2.credentials import Credentials
        from googleapiclient.discovery import build
        creds = Credentials(token=None, client_id=GOOGLE_OAUTH_CLIENT_ID,
            client_secret=GOOGLE_OAUTH_CLIENT_SECRET,
            refresh_token=os.environ.get("GOOGLE_REFRESH_TOKEN"),
            token_uri="https://oauth2.googleapis.com/token")
        if not creds.valid: creds.refresh(requests.Request())
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
        print(f"[{i+1}/{len(QUERIES)}] {q[:50]}...")
        for item in search_all(q):
            opp = classify(item)
            if opp and not opportunity_exists(ws, opp.url):
                ws.append([opp.id, opp.title, opp.url, opp.category, opp.description,
                    opp.profit_per_day, opp.profit_per_week, opp.profit_per_month,
                    opp.effort_level, opp.source, opp.found_date, ",".join(opp.tags), opp.status])
                total_new += 1
                categories[opp.category] = categories.get(opp.category, 0) + 1
                mem["total_found"] += 1
        mem["learning"].append({"date": datetime.now(timezone.utc).isoformat(), "query": q, "results": 0, "new_opps": 0})
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
