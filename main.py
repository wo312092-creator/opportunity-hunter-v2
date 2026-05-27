import os, json, time, re, hashlib, sys, urllib.parse, html as html_mod
from datetime import datetime, timezone
from typing import Optional
from dataclasses import dataclass, field

os.system("pip install requests openpyxl google-auth google-auth-oauthlib google-auth-httplib2 google-api-python-client playwright google-generativeai 2>/dev/null")
os.system("playwright install chromium 2>/dev/null")

import requests
from openpyxl import Workbook, load_workbook

FIRECRAWL_API_KEY = os.environ.get("FIRECRAWL_API_KEY", "")
GEMINI_API_KEY = os.environ.get("GEMINI_API_KEY", "")
GOOGLE_OAUTH_CLIENT_ID = os.environ.get("GOOGLE_OAUTH_CLIENT_ID", "")
GOOGLE_OAUTH_CLIENT_SECRET = os.environ.get("GOOGLE_OAUTH_CLIENT_SECRET", "")

EXCEL_FILE = "opportunities.xlsx"
MEMORY_FILE = "bot_memory.json"
REPORT_DIR = "reports"
TOP_FINDS_FILE = "top_automatable_finds.json"

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
    automation_potential: int = 0
    automation_reason: str = ""
    source: str = ""
    found_date: str = ""
    tags: list = field(default_factory=list)
    status: str = "new"

def load_excel():
    if not os.path.exists(EXCEL_FILE):
        wb = Workbook()
        ws = wb.active
        ws.title = "Opportunities"
        headers = ["ID","Title","URL","Category","Description","Profit/Day","Profit/Week","Profit/Month","Effort","AutoScore","AutoReason","Source","Found Date","Tags","Status"]
        ws.append(headers)
        for col, w in [("A",8),("B",40),("C",50),("D",20),("E",60),("F",15),("G",15),("H",15),("I",12),("J",10),("K",50),("L",20),("M",20),("N",30),("O",10)]:
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
    if "duckduckgo.com/l/" in u:
        parsed = urllib.parse.urlparse(u)
        qs = urllib.parse.parse_qs(parsed.query)
        return qs.get("uddg", [u])[0]
    return u

def search_bing_html(query: str) -> list:
    try:
        resp = requests.get("https://www.bing.com/search", params={"q": query, "count": 15},
            headers={"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/125.0.0.0 Safari/537.36", "Accept-Language": "en-US"},
            timeout=10)
        links = re.findall(r'<a[^>]+href="(https?://[^"]+)"[^>]*><h2>(.*?)</h2>', resp.text, re.DOTALL)
        if not links:
            links = re.findall(r'<h2[^>]*><a[^>]+href="(https?://[^"]+)"[^>]*>(.*?)</a></h2>', resp.text, re.DOTALL)
        snippets = re.findall(r'<p[^>]*>(.*?)</p>', resp.text, re.DOTALL)
        results = []
        for i, (u, t) in enumerate(links[:12]):
            desc = html_mod.unescape(re.sub(r'<[^>]+>', '', snippets[i] if i < len(snippets) else "")).strip()[:300] if i < len(snippets) else ""
            results.append({"title": html_mod.unescape(re.sub(r'<[^>]+>', '', t)).strip()[:200], "url": u, "description": desc})
        print(f"[Bing HTML] {len(results)} results")
        return results
    except Exception as e:
        print(f"[Bing HTML] Error: {e}")
        return []

def search_ddg(query: str) -> list:
    try:
        resp = requests.get("https://html.duckduckgo.com/html/", params={"q": query},
            headers={"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36"}, timeout=8)
        links = re.findall(r'<a[^>]+class="result__a"[^>]+href="([^"]+)"[^>]*>(.*?)</a>', resp.text, re.DOTALL)
        snippets = re.findall(r'<a[^>]+class="result__snippet"[^>]*>(.*?)</a>', resp.text, re.DOTALL)
        if not links:
            return []
        results = [{"title": html_mod.unescape(re.sub(r'<[^>]+>', '', t)).strip(),
                  "url": extract_ddg_url(u),
                  "description": html_mod.unescape(re.sub(r'<[^>]+>', '', snippets[i] if i < len(snippets) else "")).strip()[:300]}
                for i, (u, t) in enumerate(links[:10]) if u]
        return results
    except:
        return []

class PlaywrightPool:
    def __init__(self):
        self.browser = None
        self.ctx = None
    def start(self):
        from playwright.sync_api import sync_playwright
        self._pw = sync_playwright()
        p = self._pw.start()
        self.browser = p.chromium.launch(headless=True, args=["--no-sandbox","--disable-setuid-sandbox","--window-size=1920,1080"])
        self.ctx = self.browser.new_context(user_agent="Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/125.0.0.0 Safari/537.36", viewport={"width": 1920, "height": 1080}, locale="en-US")
    def search(self, query: str) -> list:
        if not self.browser:
            self.start()
        try:
            page = self.ctx.new_page()
            page.goto(f"https://www.bing.com/search?q={urllib.parse.quote(query)}&count=15", timeout=20000)
            page.wait_for_timeout(1500)
            if "captcha" in page.content().lower():
                page.close()
                return []
            page.evaluate("window.scrollBy(0, 400)")
            page.wait_for_timeout(500)
            items = page.query_selector_all("li.b_algo")
            results = []
            for el in items[:15]:
                try:
                    h2 = el.query_selector("h2")
                    a = h2.query_selector("a[href^='http']") if h2 else None
                    if not a: continue
                    url = a.get_attribute("href") or ""
                    title = h2.inner_text().strip()
                    desc_el = el.query_selector("div.b_caption p, div.b_snippet")
                    desc = desc_el.inner_text()[:400] if desc_el else ""
                    if url and title:
                        results.append({"title": title, "url": url, "description": desc.strip()})
                except: continue
            page.close()
            return results
        except:
            return []
    def close(self):
        if self.browser:
            self.browser.close()

def search_all(query: str, pw: PlaywrightPool) -> list:
    seen = set()
    results = []
    bing_results = search_bing_html(query)
    for r in bing_results:
        if r["url"] and r["url"] not in seen:
            seen.add(r["url"])
            results.append(r)
    time.sleep(0.5)
    if len(bing_results) < 3:
        pw_results = pw.search(query)
        for r in pw_results:
            if r["url"] and r["url"] not in seen:
                seen.add(r["url"])
                results.append(r)
        time.sleep(0.3)
    ddg_results = search_ddg(query)
    for r in ddg_results:
        if r["url"] and r["url"] not in seen:
            seen.add(r["url"])
            results.append(r)
    return results

def score_automation_gemini(title: str, desc: str, url: str) -> tuple:
    if not GEMINI_API_KEY:
        return rule_score_automation(title, desc, url)
    try:
        import google.generativeai as genai
        genai.configure(api_key=GEMINI_API_KEY)
        model = genai.GenerativeModel('gemini-2.0-flash')
        prompt = f"""Rate this money-making opportunity's AUTOMATION POTENTIAL (0-10).

High score (7-10): fully automatable with a GitHub bot - auto-click, auto-claim, auto-faucet, auto-mining, auto-task scripts.
Medium score (4-6): partially automatable - needs occasional captchas or approvals.
Low score (0-3): requires continuous human work - typing, reading, manual trading.

Title: {title[:200]}
Description: {desc[:300]}
URL: {url[:200]}

Respond ONLY with JSON: {{"score": N, "reason": "short reason"}}"""
        resp = model.generate_content(prompt)
        text = resp.text.strip()
        match = re.search(r'\{[^}]+\}', text)
        if match:
            data = json.loads(match.group())
            score = max(0, min(10, int(data.get("score", 5))))
            return score, str(data.get("reason", "Gemini analysis"))[:200]
        return 5, "Gemini: parse error"
    except Exception as e:
        print(f"[Gemini] Score error: {e}")
        return rule_score_automation(title, desc, url)

def rule_score_automation(title: str, desc: str, url: str) -> tuple:
    c = f"{title} {desc} {url}".lower()
    if any(s in c for s in ["dictionary","meaning","definition","wikipedia","encyclopedia","tutorial","course","educational","academic"]):
        return 1, "Educational content"
    if any(s in url for s in ["autotrader","carsforsale","kbb.com"]):
        return 1, "Car site - not earnings"
    bonus, signals = 0, 0
    word_groups = [(["mining","miner","hash","cloud mining","free btc","free eth","free ltc","free doge"], 3),
                   (["faucet","claim","withdraw","crypto faucet"], 3),
                   (["auto","bot","telegram bot","auto claim","auto bot"], 2),
                   (["click","ptc","paid-to-click","earn per"], 2),
                   (["passive income","staking","stake","passive"], 2),
                   (["refer","affiliate","earn crypto","reward","bonus","cash"], 1)]
    for words, pts in word_groups:
        if any(w in c for w in words):
            bonus += pts
            signals += 1
    if "survey" in c or "data entry" in c or "freelance" in c:
        bonus -= 2
    score = min(10, max(0, 5 + bonus // max(1, abs(signals // 2)))) if signals > 0 else 0
    reasons = {7: "High automation: bot-friendly", 4: "Medium: partially automatable", 0: "Low: human needed"}
    reason = next((r for threshold, r in sorted(reasons.items(), reverse=True) if score >= threshold), "Low: human needed")
    return score, reason

CLASSIFY_KEYWORDS = [
    (["faucet","free crypto","claim","btc faucet","eth faucet","crypto faucet","bitcoin faucet"], "Crypto Faucet", ["faucet","crypto","free"]),
    (["airdrop","token distribution","free token","claim airdrop","crypto airdrop"], "Airdrop", ["airdrop","crypto","free"]),
    (["ptc","paid-to-click","bux","click ads","get paid to","paidtoclick","earn per click"], "PTC / GPT", ["ptc","gpt","click"]),
    (["survey","paid survey","market research","survey site","paid surveys"], "Paid Surveys", ["survey","research"]),
    (["cashback","cash back","rebate","shopping","cashback site"], "Cashback", ["cashback","shopping"]),
    (["affiliate","referral","refer","affiliate program"], "Affiliate", ["affiliate","referral"]),
    (["stake","staking","defi","yield","lend","apr","liquidity","pool"], "DeFi / Staking", ["defi","staking","yield"]),
    (["play to earn","p2e","gamefi","nft game","play-to-earn","crypto game"], "Play-to-Earn", ["p2e","gaming","nft"]),
    (["mining","cloud mining","hash","mine","miner","crypto mining","ltc miner","bitcoin miner","free mining","mining pool"], "Mining", ["mining","crypto"]),
    (["trading bot","auto trade","signal","copy trade","trading platform","grid trading"], "Trading", ["trading","bot","automation"]),
    (["micro task","microtask","captcha","data entry","freelance","micro job","gig"], "Micro Tasks", ["micro-task","freelance"]),
    (["browser automation","auto earn","auto claim","auto bot","automation bot","auto click","auto faucet"], "Automation Bot", ["automation","bot"]),
    (["earn crypto","free crypto","get crypto","crypto earn","crypto reward","crypto bonus"], "Crypto Earnings", ["crypto","earn"]),
    (["ltc","litecoin","dogecoin","doge coin","doge mining","ltc mining","free ltc","free doge"], "Altcoin Mining", ["mining","ltc","doge"]),
]

def classify(item: dict) -> Optional[Opportunity]:
    t, u, d = item.get("title",""), item.get("url",""), item.get("description","")
    c = f"{t} {d}".lower()
    for keywords, cat, tags in CLASSIFY_KEYWORDS:
        if any(k in c for k in keywords):
            auto_score, auto_reason = score_automation_gemini(t, d, u)
            return Opportunity(id=hashlib.md5(f"{t}{u}{time.time()}".encode()).hexdigest()[:12],
                title=t[:120], url=u[:300], category=cat, description=d[:500],
                profit_per_day="Unknown", profit_per_week="Unknown", profit_per_month="Unknown",
                effort_level="Unknown", automation_potential=auto_score, automation_reason=auto_reason,
                source="web_search", found_date=datetime.now(timezone.utc).strftime("%Y-%m-%d %H:%M:%S UTC"),
                tags=tags, status="new")
    if any(k in c for k in ["earn","money","income","profit","passive","free","reward","bonus","cash","crypto","bitcoin"]):
        auto_score, auto_reason = score_automation_gemini(t, d, u)
        return Opportunity(id=hashlib.md5(f"{t}{u}{time.time()}".encode()).hexdigest()[:12],
            title=t[:120], url=u[:300], category="General", description=d[:500],
            profit_per_day="Unknown", profit_per_week="Unknown", profit_per_month="Unknown",
            effort_level="Unknown", automation_potential=auto_score, automation_reason=auto_reason,
            source="web_search", found_date=datetime.now(timezone.utc).strftime("%Y-%m-%d %H:%M:%S UTC"),
            tags=["money","earn"], status="new")
    return None

QUERIES = [
    "free ltc mining sites 2026 no deposit instant withdrawal",
    "free dogecoin mining sites 2026 no deposit",
    "free bitcoin mining cloud mining 2026 no deposit withdraw",
    "free ethereum mining sites 2026 no investment",
    "crypto mining faucet earn free btc eth ltc doge 2026",
    "best crypto faucets 2026 free bitcoin ethereum litecoin",
    "auto claim crypto faucet bot 2026",
    "paid to click sites that pay instantly 2026 free registration",
    "best GPT sites earn money free 2026 auto earn",
    "new crypto airdrops 2026 free tokens claim",
    "solana airdrop 2026 claim free",
    "telegram bot airdrop claim 2026",
    "passive income crypto staking defi 2026 no minimum",
    "free crypto staking rewards 2026",
    "browser automation earn crypto free 2026",
    "telegram bot earn crypto 2026 free automated",
    "auto trading crypto bot free 2026",
    "free crypto arbitrage bot 2026",
    "play to earn crypto games 2026 free no investment",
    "micro task sites pay crypto 2026",
    "best cashback apps 2026 free money crypto",
    "affiliate programs crypto 2026 high paying free",
    "earn free crypto no deposit 2026 withdraw instantly",
    "free litecoin mining pool 2026 no deposit required",
    "doge coin faucet free claim every hour 2026",
    "btc mining telegram bot free 2026",
    "automatic crypto earning platform 2026 no investment",
    "free bitcoin earning sites 2026 withdraw to wallet",
    "cloud mining free trial 2026 no deposit btc",
    "web3 earn crypto free 2026 browser mining",
    "free crypto signals telegram 2026 copy trade",
    "defi yield farming 2026 no minimum deposit",
    "passive crypto income 2026 set and forget",
    "telegram mining bot free withdrawal 2026",
    "faucet pay crypto instant 2026 free claim",
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
        f.write(f"**Queries:** {len(QUERIES)}  |  **Run:** #{json.load(open(MEMORY_FILE)).get('runs',0) if os.path.exists(MEMORY_FILE) else 1}\n\n")
        f.write("## Categories\n\n")
        for cat, cnt in sorted(categories.items(), key=lambda x: -x[1]):
            f.write(f"- **{cat}**: {cnt}\n")
        f.write("\n## Top Automatable Finds (AutoScore >= 6)\n\n")
        f.write("| Title | Auto | Category | Link |\n|---|---|---|---|\n")
        for row in ws.iter_rows(min_row=2, values_only=True):
            if len(row) > 9 and row[9] and isinstance(row[9], (int,float)) and row[9] >= 6:
                f.write(f"| {str(row[1] or '')[:40]} | {row[9]}/10 | {row[3]} | {str(row[2] or '')[:50]} |\n")
        f.write("\n## Latest 20\n\n| Title | Category | AutoScore | Link |\n|---|---|---|---|\n")
        for i, row in enumerate(ws.iter_rows(min_row=2, values_only=True)):
            if i >= 20: break
            if len(row) > 9:
                f.write(f"| {str(row[1] or '')[:40]} | {row[3]} | {row[9] or '?'}/10 | {str(row[2] or '')[:50]} |\n")
    return path

def write_top_finds(ws):
    top = []
    for row in ws.iter_rows(min_row=2, values_only=True):
        if len(row) > 9 and row[9] and isinstance(row[9], (int,float)) and row[9] >= 7:
            top.append({"title": row[1] or "", "url": row[2] or "", "category": row[3] or "",
                "description": (row[4] or "")[:200], "automation_score": row[9],
                "automation_reason": row[10] or "", "tags": str(row[13] or "")})
    top = sorted(top, key=lambda x: -x["automation_score"])[:30]
    with open(TOP_FINDS_FILE, "w", encoding="utf-8") as f:
        json.dump({"updated": datetime.now(timezone.utc).isoformat(), "top_finds": top}, f, indent=2)
    print(f"[Top Finds] {len(top)} high-automation opportunities saved")

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
            refresh_token=refresh_token, token_uri="https://oauth2.googleapis.com/token")
        creds.refresh(GoogleRequest())
        docs = build("docs", "v1", credentials=creds)
        date_str = datetime.now(timezone.utc).strftime("%Y-%m-%d")
        doc = docs.documents().create(body={"title": f"Opportunity Hunter Report - {date_str}"}).execute()
        doc_id = doc["documentId"]
        with open(report_path, encoding="utf-8") as f:
            content = f.read()
        clean_lines = []
        for line in content.split('\n'):
            line = re.sub(r'^#{1,6}\s+', '', line)
            line = line.replace('**', '').replace('*', '').replace('__', '').replace('_', '')
            if re.match(r'^\|[\s\-:]+\|$', line): continue
            if line.startswith('|') and '|' in line[1:]:
                cells = [c.strip() for c in line.split('|')[1:-1]]
                line = ' | '.join(cells)
            clean_lines.append(line)
        clean_text = '\n'.join(clean_lines)
        docs.documents().batchUpdate(documentId=doc_id,
            body={"requests": [{"insertText": {"endOfSegmentLocation": {}, "text": clean_text}}]}).execute()
        print(f"[Google Docs] Created: https://docs.google.com/document/d/{doc_id}")
        return True
    except Exception as e:
        print(f"[Google Docs] Error: {e}")
        return False

def main():
    mem = json.load(open(MEMORY_FILE)) if os.path.exists(MEMORY_FILE) else {"runs":0,"total_found":0,"categories_found":{},"last_run":None,"learning":[]}
    mem["runs"] += 1
    wb, ws = load_excel()
    total_new, categories = 0, {}

    pw = PlaywrightPool()
    pw.start()
    start_time = time.time()

    for i, q in enumerate(QUERIES):
        elapsed = time.time() - start_time
        sys.stdout.write(f"[{i+1}/{len(QUERIES)}] ({elapsed:.0f}s) {q[:55]}... ")
        sys.stdout.flush()
        items = search_all(q, pw)
        sys.stdout.write(f"{len(items)} items... ")
        sys.stdout.flush()
        new_for_query = 0
        for item in items:
            opp = classify(item)
            if opp and not opportunity_exists(ws, opp.url):
                ws.append([opp.id, opp.title, opp.url, opp.category, opp.description,
                    opp.profit_per_day, opp.profit_per_week, opp.profit_per_month,
                    opp.effort_level, opp.automation_potential, opp.automation_reason,
                    opp.source, opp.found_date, ",".join(opp.tags), opp.status])
                new_for_query += 1
                total_new += 1
                categories[opp.category] = categories.get(opp.category, 0) + 1
                mem["total_found"] += 1
        print(f"+{new_for_query}")
        mem["learning"].append({"date": datetime.now(timezone.utc).isoformat(), "query": q, "results": len(items), "new_opps": new_for_query})
        if len(mem["learning"]) > 100: mem["learning"] = mem["learning"][-100:]
        time.sleep(1)

    pw.close()
    wb.save(EXCEL_FILE)
    mem["last_run"] = datetime.now(timezone.utc).isoformat()
    for cat, cnt in categories.items():
        mem["categories_found"][cat] = mem["categories_found"].get(cat, 0) + cnt
    json.dump(mem, open(MEMORY_FILE, "w"), indent=2)

    write_top_finds(ws)
    report_path = write_report(wb, total_new, categories)
    elapsed = time.time() - start_time
    print(f"Run #{mem['runs']}: +{total_new} new, {max(0, ws.max_row-1)} total in {elapsed:.0f}s")
    write_google_doc(report_path)
    print(f"[{datetime.now(timezone.utc).isoformat()}] Done!")

if __name__ == "__main__":
    main()
