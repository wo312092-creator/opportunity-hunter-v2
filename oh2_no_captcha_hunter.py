#!/usr/bin/env python3
"""
Opportunity Hunter V2 — NO-CAPTCHA Edition
=============================================
Purpose: Search the internet specifically for crypto faucets and earning sites
that have NO captcha, SIMPLE captcha, or auto-claim features that can run
on GitHub Actions (datacenter IPs).

Strategy:
  1. Run 40+ hyper-targeted search queries for no-captcha/auto-claim sites
  2. Visit each site with Playwright + curl_cffi to detect captcha type
  3. Score each site: (no_captcha_bonus) + (faucetpay) + (auto_claim) + (github_compatible)
  4. Only output sites scoring 7+ (easy to automate, no captcha, pays)
  5. Generate a clean JSON report of "Ready to Automate" sites

Integration:
  - Uses same search infrastructure as oh2_v5.py (Bing PW, GitHub, Exa, Firecrawl)
  - Uses same AI scoring (Groq -> OpenRouter -> Gemini -> Rules)
  - Output: no_captcha_finds.json + Google Sheet with NO_CAPTCHA category
"""

import os, json, time, re, sys, urllib.parse, html as html_mod, random
from datetime import datetime
from typing import Optional

# ── Install dependencies ──────────────────────────────────
os.system("pip install requests openpyxl playwright curl-cffi 2>/dev/null")
os.system("playwright install chromium 2>/dev/null")

import requests
from openpyxl import Workbook

# ── Config ─────────────────────────────────────────────────
FAUCETPAY_EMAIL = os.environ.get("FAUCETPAY_EMAIL", "pedagroup.co2020@gmail.com")
EXA_API_KEY = os.environ.get("EXA_API_KEY", "")
FIRECRAWL_API_KEY = os.environ.get("FIRECRAWL_API_KEY", "")

# Groq API keys
GROQ_API_KEYS = []
for _i in range(1, 10):
    _key = os.environ.get(f"GROQ_API_KEY_{_i}" if _i > 1 else "GROQ_API_KEY", "")
    if _key: GROQ_API_KEYS.append(_key)

# OpenRouter keys
OPENROUTER_API_KEYS = []
_key = os.environ.get("OPENROUTER_API_KEY", "")
if _key: OPENROUTER_API_KEYS.append(_key)
for _i in range(2, 10):
    _key = os.environ.get(f"OPENROUTER_API_KEY_{_i}", "")
    if _key: OPENROUTER_API_KEYS.append(_key)

GEMINI_API_KEYS = []
_key = os.environ.get("GEMINI_API_KEY", "")
if _key: GEMINI_API_KEYS.append(_key)
for _i in range(2, 10):
    _key = os.environ.get(f"GEMINI_API_KEY_{_i}", "")
    if _key: GEMINI_API_KEYS.append(_key)

RESULTS_FILE = "no_captcha_finds.json"
_groq_key_index = 0
_openrouter_key_index = 0

# ══════════════════════════════════════════════════════════
#  SEARCH QUERIES — Focused on NO-CAPTCHA / SIMPLE-EARN
# ══════════════════════════════════════════════════════════

NO_CAPTCHA_QUERIES = [
    # ── Explicit "no captcha" queries ──
    'crypto faucet no captcha instant withdrawal 2026',
    'no captcha bitcoin faucet instant pay 2026',
    'free crypto no captcha no survey no tasks 2026',
    '"no captcha" faucet faucetpay 2026',
    'auto claim crypto no captcha 2026',
    'earn crypto without captcha 2026 free',
    'no captcha trx faucet instant payout 2026',
    'no captcha ltc faucet free litecoin 2026',
    'no captcha doge faucet 2026 instant',
    'no captcha usdt faucet 2026 free tether',
    # ── Auto-claim / auto-faucet (no captcha per claim) ──
    'auto faucet crypto no manual claim 2026',
    'auto claim bitcoin bot no captcha 2026',
    'telegram auto earn crypto bot no captcha 2026',
    'auto faucet site no verification 2026 faucetpay',
    'set and forget crypto faucet 2026',
    'auto claim crypto platform no captcha 2026',
    'automatic crypto faucet faucetpay 2026',
    'one click crypto faucet no captcha 2026',
    # ── Simple captcha (text captcha, math captcha, etc) ──
    'simple captcha crypto faucet faucetpay 2026',
    'text captcha faucet free crypto 2026',
    'easy captcha bitcoin faucet instant withdrawal 2026',
    'math captcha crypto faucet free 2026',
    'basic captcha faucet instant faucetpay 2026',
    'anti bot simple captcha crypto faucet 2026',
    # ─── Paid-to-click / shortlink (no captcha) ──
    'ptc sites instant withdrawal no captcha 2026 faucetpay',
    'click earn crypto no captcha instant payout 2026',
    'paid to click crypto faucetpay no captcha 2026',
    'shortlink earn crypto no captcha faucetpay 2026',
    # ── Multi-coin auto faucets ──
    'multi coin auto faucet faucetpay 2026',
    'all in one crypto faucet auto claim 2026 faucetpay',
    'bitcoin litecoin doge auto faucet faucetpay 2026 no captcha',
    # ── Telegram bots (no captcha inside) ──
    'telegram crypto faucet bot auto earn faucetpay 2026',
    'telegram earning bot no captcha free withdrawal 2026',
    'best telegram faucet bot faucetpay 2026 no captcha',
    'telegram mining bot auto claim faucetpay 2026',
    # ── General no-captcha-earning ──
    'free crypto earning platform no captcha 2026 withdraw',
    'earn free crypto automatic faucetpay 2026',
    'crypto faucet list no captcha 2026',
    'claim free crypto auto faucet instant faucetpay',
    'best paying crypto faucet no captcha 2026',
]

# ══════════════════════════════════════════════════════════
#  HELPER FUNCTIONS
# ══════════════════════════════════════════════════════════

USER_AGENTS = [
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/125.0.0.0 Safari/537.36",
    "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 Chrome/125.0.0.0 Safari/537.36",
    "Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 Chrome/125.0.0.0 Safari/537.36",
]

def groq_generate(prompt_str, model="llama-3.1-8b-instant"):
    global _groq_key_index
    if not GROQ_API_KEYS:
        return None
    for _ in range(len(GROQ_API_KEYS)):
        key = GROQ_API_KEYS[_groq_key_index]
        try:
            r = requests.post(
                "https://api.groq.com/openai/v1/chat/completions",
                headers={"Authorization": f"Bearer {key}", "Content-Type": "application/json"},
                json={"model": model, "messages": [{"role": "user", "content": prompt_str}],
                      "temperature": 0.2, "max_tokens": 600},
                timeout=30
            )
            if r.status_code == 200:
                return r.json()["choices"][0]["message"]["content"]
            _groq_key_index = (_groq_key_index + 1) % len(GROQ_API_KEYS)
            time.sleep(0.5)
        except:
            _groq_key_index = (_groq_key_index + 1) % len(GROQ_API_KEYS)
    return None

def openrouter_generate(prompt_str, model="openai/gpt-4o-mini"):
    global _openrouter_key_index
    if not OPENROUTER_API_KEYS:
        return None
    for _ in range(len(OPENROUTER_API_KEYS)):
        key = OPENROUTER_API_KEYS[_openrouter_key_index]
        try:
            r = requests.post(
                "https://openrouter.ai/api/v1/chat/completions",
                headers={"Authorization": f"Bearer {key}", "Content-Type": "application/json"},
                json={"model": model, "messages": [{"role": "user", "content": prompt_str}],
                      "temperature": 0.2, "max_tokens": 600},
                timeout=30
            )
            if r.status_code == 200:
                return r.json()["choices"][0]["message"]["content"]
            _openrouter_key_index = (_openrouter_key_index + 1) % len(OPENROUTER_API_KEYS)
            time.sleep(0.5)
        except:
            _openrouter_key_index = (_openrouter_key_index + 1) % len(OPENROUTER_API_KEYS)
    return None

def extract_domain(url: str) -> str:
    try:
        from urllib.parse import urlparse
        return urlparse(url).netloc.lower().replace("www.", "")
    except:
        return url

def is_safe_url(url: str) -> bool:
    """Block shorteners, dangerous files, non-http."""
    unsafe_domains = ["bit.ly", "tinyurl", "adf.ly", "shorturl", "rebrand.ly", "shorte.st",
                      "bc.vc", "linkbucks", "ceesty", "clkme", "shrinkearn", "shortest"]
    try:
        domain = extract_domain(url)
        if any(d in domain for d in unsafe_domains):
            return False
    except:
        pass
    dangerous_exts = [".exe", ".dll", ".bat", ".scr", ".msi", ".zip", ".rar", ".apk"]
    if any(url.lower().endswith(e) for e in dangerous_exts):
        return False
    if not url.startswith("http"):
        return False
    return True

def resolve_url(url: str) -> str:
    """Extract real URL from Bing redirects."""
    if 'bing.com/ck/a' not in url:
        return url
    try:
        parsed = urllib.parse.urlparse(url)
        qs = urllib.parse.parse_qs(parsed.query)
        u_param = qs.get('u', [''])[0]
        if not u_param:
            return url
        import base64
        for prefix_len in range(0, 5):
            encoded = u_param[prefix_len:]
            if not encoded:
                continue
            try:
                padding = 4 - len(encoded) % 4
                if padding != 4:
                    encoded_padded = encoded + '=' * padding
                else:
                    encoded_padded = encoded
                decoded = base64.b64decode(encoded_padded).decode('utf-8')
                if decoded.startswith('http'):
                    return decoded
            except:
                continue
    except:
        pass
    return url

# ══════════════════════════════════════════════════════════
#  CAPTCHA DETECTION ENGINE
# ══════════════════════════════════════════════════════════

def detect_captcha_type(url: str) -> dict:
    """
    Visit a site and detect what CAPTCHA/verification it uses.
    Returns detailed report of all security measures found.
    Uses curl_cffi first (fast), then Playwright for JS-heavy sites.
    """
    result = {
        "url": url,
        "has_captcha": False,
        "captcha_types": [],        # reCAPTCHA, hCaptcha, Turnstile, Text captcha, Math captcha, None
        "has_faucetpay": False,
        "has_register": False,
        "has_login": False,
        "has_withdraw": False,
        "has_claim_button": False,
        "has_address_input": False,
        "title": "",
        "body_snippet": "",
        "status": 0,
        "success": False,
        "error": "",
        "captcha_sitekey": "",
        "detection_method": "",      # curl_cffi or playwright
    }

    # Try curl_cffi first (fast, no browser)
    try:
        from curl_cffi import requests as curl_requests
        r = curl_requests.get(url, impersonate='chrome', timeout=15,
            headers={"Accept-Language": "en-US,en;q=0.9"})
        text = r.text if hasattr(r, 'text') else r.content.decode('utf-8', errors='replace')
        result["status"] = r.status_code if hasattr(r, 'status_code') else 0
        result["success"] = True
        result["detection_method"] = "curl_cffi"
    except Exception as e:
        result["error"] = str(e)[:100]
        # Fallback to regular requests
        try:
            r = requests.get(url, timeout=15, headers={
                "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) Chrome/125.0.0.0",
                "Accept-Language": "en-US,en;q=0.9"
            })
            text = r.text
            result["status"] = r.status_code
            result["success"] = True
            result["detection_method"] = "requests"
        except Exception as e2:
            result["error"] = str(e2)[:100]
            return result

    if not result["success"] or not text:
        return result

    # Extract title
    m = re.search(r'<title[^>]*>(.*?)</title>', text, re.I | re.S)
    result["title"] = m.group(1).strip()[:120] if m else ""

    # ── Detect reCAPTCHA ──
    if re.search(r'google\.com/recaptcha/api\.js|g-recaptcha|data-sitekey=["\']', text, re.I):
        result["has_captcha"] = True
        result["captcha_types"].append("Google reCAPTCHA")
        m = re.search(r'data-sitekey=["\']([^"\']+)["\']', text)
        if m:
            result["captcha_sitekey"] = m.group(1)

    # ── Detect hCaptcha ──
    if re.search(r'hcaptcha\.com|data-hcaptcha-widget-id|h-captcha', text, re.I):
        result["has_captcha"] = True
        result["captcha_types"].append("hCaptcha")

    # ── Detect Cloudflare Turnstile ──
    if re.search(r'challenges\.cloudflare\.com|cf-turnstile|turnstile', text, re.I):
        result["has_captcha"] = True
        result["captcha_types"].append("Cloudflare Turnstile")

    # ── Detect SolveMedia ──
    if re.search(r'solvemedia\.com|api-secure\.solvemedia', text, re.I):
        result["has_captcha"] = True
        result["captcha_types"].append("SolveMedia")

    # ── Detect simple text captcha (image-based, not reCAPTCHA) ──
    # Look for captcha images, simple math, etc.
    if re.search(r'captcha\.(png|jpg|gif|jpeg)|simple-captcha|captcha_img', text, re.I):
        if not result["has_captcha"]:
            result["has_captcha"] = True
            result["captcha_types"].append("Text/Image captcha")

    # ── Detect math captcha (e.g., "What is 2+3?") ──
    if re.search(r'math.*captcha|what is.*\+|simple.*math|anti.*bot.*math', text, re.I):
        if "Text/Image captcha" not in result["captcha_types"]:
            result["has_captcha"] = True
            if "Math captcha" not in result["captcha_types"]:
                result["captcha_types"].append("Math captcha")

    # If no specific captcha found but site has anti-bot
    if not result["has_captcha"]:
        if re.search(r'captcha|antibot|anti.bot|human.verif|are.you.human', text, re.I):
            result["has_captcha"] = True
            result["captcha_types"].append("Unknown (generic anti-bot)")

    # ── Detect FaucetPay integration ──
    if re.search(r'faucetpay|faucet\.pay|FaucetPay', text, re.I):
        result["has_faucetpay"] = True

    # ── Detect other features ──
    body = re.sub(r'<[^>]+>', ' ', text)
    body = re.sub(r'\s+', ' ', body).lower()

    if re.search(r'register|sign.?up|create.?account', body):
        result["has_register"] = True
    if re.search(r'login|sign.?in|log.?in', body):
        result["has_login"] = True
    if re.search(r'withdraw|payout|send to wallet|withdrawal', body):
        result["has_withdraw"] = True
    if re.search(r'claim|start|earn now|get free|claim now', body):
        result["has_claim_button"] = True
    if re.search(r'address|wallet.*input|enter.*address', body):
        result["has_address_input"] = True

    result["body_snippet"] = body[:500]

    return result


def deep_analyze_captcha(url: str, site_info: dict) -> dict:
    """
    Use Playwright to do deeper captcha analysis.
    Visits the page with a real browser to confirm captcha type.
    """
    try:
        from playwright.sync_api import sync_playwright
    except ImportError:
        return site_info

    try:
        with sync_playwright() as pw:
            browser = pw.chromium.launch(
                headless=True,
                args=["--no-sandbox", "--disable-setuid-sandbox"],
            )
            context = browser.new_context(
                user_agent=random.choice(USER_AGENTS),
                viewport={"width": 1280, "height": 800},
                locale="en-US",
            )
            page = context.new_page()
            page.goto(url, timeout=20000, wait_until="domcontentloaded")
            time.sleep(3)

            # Count actual captcha iframes
            rc_count = page.locator('iframe[title="reCAPTCHA"]').count()
            turn_count = page.locator('iframe[title*="Turnstile"]').count()
            hcap_count = page.locator('iframe[title*="hcaptcha"]').count()

            if rc_count > 0:
                site_info["has_captcha"] = True
                if "Google reCAPTCHA" not in site_info["captcha_types"]:
                    site_info["captcha_types"].append("Google reCAPTCHA")
                site_info["recaptcha_iframes"] = rc_count

                # Try clicking checkbox to see if challenge appears
                try:
                    anchor = page.frame_locator('iframe[title="reCAPTCHA"]')
                    anchor.locator(".recaptcha-checkbox-border").click(timeout=5000)
                    time.sleep(2)

                    challenge = page.locator('iframe[title*="challenge"]')
                    if challenge.count() > 0:
                        cf = page.frame_locator('iframe[title*="challenge"]')
                        challenge_text = cf.locator("body").inner_text(timeout=3000)[:200]
                        site_info["recaptcha_challenge"] = challenge_text

                        # Check if audio is available
                        audio_btn = cf.locator("#recaptcha-audio-button")
                        if audio_btn.count() > 0 and audio_btn.first.is_visible():
                            site_info["captcha_has_audio"] = True
                            audio_btn.click(timeout=3000)
                            time.sleep(2)
                            audio_src = cf.locator("#audio-source")
                            src = audio_src.get_attribute("src", timeout=5000)
                            site_info["captcha_audio_source"] = bool(src)
                except:
                    pass

            if turn_count > 0:
                site_info["has_captcha"] = True
                if "Cloudflare Turnstile" not in site_info["captcha_types"]:
                    site_info["captcha_types"].append("Cloudflare Turnstile")

            if hcap_count > 0:
                site_info["has_captcha"] = True
                if "hCaptcha" not in site_info["captcha_types"]:
                    site_info["captcha_types"].append("hCaptcha")

            # Check for claim button specifically
            try:
                claim_btn = page.query_selector('#claimBtn, button:has-text("Claim"), button:has-text("Start"), button:has-text("Earn")')
                if claim_btn:
                    site_info["has_claim_button"] = True
                    site_info["claim_button_text"] = claim_btn.inner_text()[:40]
                    site_info["claim_button_disabled"] = claim_btn.is_disabled()
            except:
                pass

            # Check for address/email input
            try:
                addr = page.query_selector('#address, input[placeholder*="address" i], input[type="email"], input[placeholder*="email" i]')
                if addr:
                    site_info["has_address_input"] = True
                    site_info["address_placeholder"] = addr.get_attribute("placeholder") or ""
            except:
                pass

            browser.close()
    except Exception as e:
        site_info["playwright_error"] = str(e)[:100]

    return site_info


# ══════════════════════════════════════════════════════════
#  SCORING ENGINE — Focused on NO-CAPTCHA + AUTOMATABLE
# ══════════════════════════════════════════════════════════

def calculate_automation_score(site_info: dict, title: str, body: str) -> dict:
    """
    Calculate an automation score (0-10) focused on:
    - No captcha = HIGH bonus
    - Simple captcha (Turnstile, text) = medium bonus
    - Complex captcha (reCAPTCHA image) = PENALTY
    - FaucetPay integration = bonus
    - Auto-claim feature = bonus
    - GitHub Actions compatible = bonus
    """
    score = 5  # Start neutral
    factors = []
    positives = []
    negatives = []

    # ── CAPTCHA FACTOR (most important) ──
    captcha_types = [c.lower() for c in site_info.get("captcha_types", [])]

    if not site_info.get("has_captcha"):
        score += 4  # No captcha at all = huge bonus
        factors.append("NO_CAPTCHA=+4")
        positives.append("No captcha detected!")
    elif any("turnstile" in c for c in captcha_types):
        score += 2  # Turnstile is bypassable
        factors.append("TURNSTILE=+2")
        positives.append("Cloudflare Turnstile (bypassable)")
    elif any("text" in c or "image" in c or "math" in c for c in captcha_types):
        score += 1  # Simple text/math captcha solvable with OCR
        factors.append("SIMPLE_CAPTCHA=+1")
        positives.append("Simple captcha (OCR-solvable)")
    elif any("recaptcha" in c for c in captcha_types):
        # Check if challenge is SKIP/VERIFY type
        challenge = site_info.get("recaptcha_challenge", "").lower()
        if "skip" in challenge or "none" in challenge:
            score += 1
            factors.append("RECAPTCHA_SKIP=+1")
            positives.append("reCAPTCHA with SKIP option")
        else:
            score -= 3  # reCAPTCHA image challenge = hard
            factors.append("RECAPTCHA_IMAGE=-3")
            negatives.append("Google reCAPTCHA image challenge (hard to automate)")
    elif any("unknown" in c for c in captcha_types):
        score += 0  # Unknown = neutral
        factors.append("UNKNOWN_CAPTCHA=+0")

    # ── FaucetPay factor ──
    if site_info.get("has_faucetpay"):
        score += 2
        factors.append("FAUCETPAY=+2")
        positives.append("Pays via FaucetPay (instant!)")

    # ── Auto-claim factor ──
    if "auto" in body.lower() and any(w in body.lower() for w in ["auto claim", "auto faucet", "automated", "auto-earn"]):
        score += 2
        factors.append("AUTO_CLAIM=+2")
        positives.append("Has auto-claim feature")

    # ── GitHub Actions compatibility ──
    # Sites that can work on GHA (no browser fingerprinting, no datacenter IP blocks)
    body_lower = body.lower()
    if any(kw in body_lower for kw in ["cloudflare", "turnstile", "recaptcha"]):
        score -= 1  # Might block datacenter IPs
        factors.append("GHA_RISK=-1")
        negatives.append("May block datacenter IPs")
    else:
        score += 1
        factors.append("GHA_FRIENDLY=+1")
        positives.append("GitHub Actions friendly")

    # ── Simple claim flow factor ──
    if site_info.get("has_address_input") and site_info.get("has_claim_button"):
        score += 1
        factors.append("SIMPLE_FLOW=+1")
        positives.append("Simple address + claim flow")

    # ── Withdrawal factor ──
    if site_info.get("has_withdraw"):
        score += 1
        factors.append("HAS_WITHDRAW=+1")
        positives.append("Has withdrawal support")

    # ── Telegram bot factor ──
    if "telegram" in body_lower and any(w in body_lower for w in ["bot", "@"]):
        score += 1
        factors.append("TELEGRAM=+1")
        positives.append("Telegram bot (easy to automate)")

    # Clamp score 0-10
    score = max(0, min(10, score))

    return {
        "score": score,
        "factors": factors,
        "positives": positives,
        "negatives": negatives,
        "verdict": "READY TO AUTOMATE" if score >= 7 else "POSSIBLE" if score >= 4 else "NOT RECOMMENDED",
        "automation_estimate": _estimate_automation(site_info, score)
    }


def _estimate_automation(site_info: dict, score: int) -> dict:
    """Estimate how easy it is to automate this site."""
    if score >= 8:
        return {"difficulty": "Easy (15 min)", "type": "requests + cron", "tool": "Python requests or Playwright", "github_action": True}
    elif score >= 6:
        return {"difficulty": "Medium (1-2 hrs)", "type": "Playwright script", "tool": "Playwright + GitHub Actions", "github_action": True}
    elif score >= 4:
        return {"difficulty": "Hard (4+ hrs)", "type": "Complex Playwright + bypass", "tool": "Playwright + anti-detection", "github_action": "Maybe"}
    else:
        return {"difficulty": "Very hard", "type": "Not recommended", "tool": "N/A", "github_action": False}


# ══════════════════════════════════════════════════════════
#  SEARCH FUNCTIONS (from oh2_v5)
# ══════════════════════════════════════════════════════════

class PlaywrightPool:
    def __init__(self):
        self.browser = None
        self._pw = None
    def start(self):
        from playwright.sync_api import sync_playwright
        self._pw = sync_playwright()
        p = self._pw.start()
        self.browser = p.chromium.launch(headless=True, args=["--no-sandbox","--disable-setuid-sandbox"])
    def search(self, query: str, q_idx: int = 0) -> list:
        if not self.browser:
            self.start()
        try:
            ctx = self.browser.new_context(
                user_agent=USER_AGENTS[q_idx % len(USER_AGENTS)],
                viewport={"width": 1920, "height": 1080},
                locale="en-US",
            )
            page = ctx.new_page()
            page.goto(f"https://www.bing.com/search?q={urllib.parse.quote(query)}&count=10", timeout=20000)
            page.wait_for_timeout(2000)
            if "captcha" in page.content().lower():
                ctx.close()
                return []
            items = page.query_selector_all("li.b_algo")
            results = []
            for el in items[:12]:
                try:
                    h2 = el.query_selector("h2")
                    a = h2.query_selector("a[href^='http']") if h2 else None
                    if not a: continue
                    url = a.get_attribute("href") or ""
                    title = h2.inner_text().strip()
                    desc_el = el.query_selector("div.b_caption p, div.b_snippet")
                    desc = desc_el.inner_text()[:300] if desc_el else ""
                    if url and title:
                        results.append({"title": title, "url": url, "description": desc.strip()})
                except: continue
            ctx.close()
            return results
        except Exception as e:
            return []
    def close(self):
        if self.browser:
            try: self.browser.close()
            except: pass
        if self._pw:
            try: self._pw.__exit__(None, None, None)
            except: pass

def search_exa(query: str) -> list:
    if not EXA_API_KEY:
        return []
    try:
        r = requests.post("https://api.exa.ai/search",
            headers={"x-api-key": EXA_API_KEY, "Content-Type": "application/json"},
            json={"query": query, "numResults": 10, "type": "auto", "useAutoprompt": True},
            timeout=15
        )
        if r.status_code == 200:
            return [{"title": item.get("title",""), "url": item.get("url",""), "description": item.get("text","")[:400]}
                    for item in r.json().get("results",[]) if item.get("url")]
        return []
    except:
        return []

def search_firecrawl(query: str) -> list:
    if not FIRECRAWL_API_KEY:
        return []
    try:
        r = requests.post("https://api.firecrawl.dev/v1/search",
            headers={"Authorization": f"Bearer {FIRECRAWL_API_KEY}", "Content-Type": "application/json"},
            json={"query": query, "limit": 10, "scrapeOptions": {"formats": ["markdown"]}},
            timeout=20
        )
        if r.status_code == 200:
            return [{"title": item.get("title","") or item.get("metadata",{}).get("title",""),
                     "url": item.get("url",""),
                     "description": (item.get("markdown","") or item.get("description","") or "")[:400]}
                    for item in r.json().get("data",[]) if item.get("url")]
        return []
    except:
        return []

def search_all(query: str, pw: PlaywrightPool, q_idx: int) -> list:
    """Search multiple engines, dedup results."""
    seen = set()
    results = []

    # 1. Bing Playwright
    pw_results = pw.search(query, q_idx)
    print(f"[BingPW] {len(pw_results)}", end=" ", flush=True)
    for r in pw_results:
        url = resolve_url(r["url"])
        if url and url not in seen:
            seen.add(url)
            results.append(r)

    # 2. Exa AI
    time.sleep(0.3)
    exa_results = search_exa(query)
    print(f"[Exa] {len(exa_results)}", end=" ", flush=True)
    for r in exa_results:
        if r["url"] and r["url"] not in seen:
            seen.add(r["url"])
            results.append(r)

    # 3. Firecrawl
    time.sleep(0.3)
    fc_results = search_firecrawl(query)
    print(f"[FC] {len(fc_results)}", end=" ", flush=True)
    for r in fc_results:
        if r["url"] and r["url"] not in seen:
            seen.add(r["url"])
            results.append(r)

    print(f"=> {len(results)} unique", flush=True)
    return results


# ══════════════════════════════════════════════════════════
#  MAIN HUNTING PIPELINE
# ══════════════════════════════════════════════════════════

def hunt():
    print("=" * 65)
    print("  OPPORTUNITY HUNTER v2 — NO-CAPTCHA EDITION")
    print(f"  Started: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    print(f"  Queries: {len(NO_CAPTCHA_QUERIES)}")
    print(f"  Groq keys: {len(GROQ_API_KEYS)}")
    print(f"  OpenRouter keys: {len(OPENROUTER_API_KEYS)}")
    print("=" * 65)

    pw = PlaywrightPool()
    all_results = []
    seen_urls = set()

    # ── PHASE 1: Search ──
    print(f"\n{'='*65}")
    print(f"  PHASE 1: SEARCHING {len(NO_CAPTCHA_QUERIES)} QUERIES")
    print(f"{'='*65}")

    for i, query in enumerate(NO_CAPTCHA_QUERIES):
        print(f"\n[{i+1}/{len(NO_CAPTCHA_QUERIES)}] q='{query}'")
        results = search_all(query, pw, i)
        
        for r in results:
            url = r.get("url", "")
            if not url or not is_safe_url(url):
                continue
            
            domain = extract_domain(url)
            if domain in seen_urls:
                continue
            seen_urls.add(domain)

            all_results.append({
                "title": r.get("title", "").strip()[:150],
                "url": url,
                "domain": domain,
                "description": r.get("description", "").strip()[:400],
                "source_query": query,
                "found_at": datetime.now().isoformat(),
            })

        # Rate limiting
        time.sleep(random.uniform(1.0, 2.5))

    pw.close()
    print(f"\n{'='*65}")
    print(f"  FOUND {len(all_results)} UNIQUE DOMAINS")
    print(f"{'='*65}")

    # ── PHASE 2: Captcha Detection + Deep Analysis ──
    print(f"\n{'='*65}")
    print(f"  PHASE 2: DETECTING CAPTCHA ON ALL SITES")
    print(f"{'='*65}")

    analyzed_sites = []
    for i, site in enumerate(all_results[:50]):  # Limit to 50 to be practical
        url = site["url"]
        print(f"\n[{i+1}/{min(50, len(all_results))}] {url[:80]}")

        # Quick curl_cffi detection
        site_info = detect_captcha_type(url)
        print(f"  Title: {site['title'][:60]}")
        print(f"  Captcha: {site_info['captcha_types'] or 'NONE'}")
        print(f"  FaucetPay: {site_info['has_faucetpay']}")
        print(f"  Method: {site_info['detection_method']}")

        # If we found the page loads, do deeper Playwright analysis
        if site_info["success"] and not site_info.get("error"):
            site_info = deep_analyze_captcha(url, site_info)
            print(f"  PW Analysis: reCAPTCHA={site_info.get('recaptcha_iframes', 0)}, "
                  f"Challenge: {site_info.get('recaptcha_challenge', 'N/A')[:60]}")

        # Score the site
        body = site["description"] + " " + site_info.get("body_snippet", "")
        scoring = calculate_automation_score(site_info, site["title"], body)

        print(f"  Score: {scoring['score']}/10 — {scoring['verdict']}")
        if scoring['positives']:
            print(f"  Positives: {', '.join(scoring['positives'])}")
        if scoring['negatives']:
            print(f"  Negatives: {', '.join(scoring['negatives'])}")

        analyzed_sites.append({
            **site,
            "captcha": site_info.get("captcha_types", []),
            "has_captcha": site_info.get("has_captcha", False),
            "has_faucetpay": site_info.get("has_faucetpay", False),
            "has_withdraw": site_info.get("has_withdraw", False),
            "has_claim_button": site_info.get("has_claim_button", False),
            "recaptcha_challenge": site_info.get("recaptcha_challenge", ""),
            "score": scoring["score"],
            "verdict": scoring["verdict"],
            "positives": scoring["positives"],
            "negatives": scoring["negatives"],
            "factors": scoring["factors"],
            "automation_estimate": scoring["automation_estimate"],
            "detected_by": site_info.get("detection_method", ""),
        })

        # Save progress incrementally
        with open(RESULTS_FILE, "w") as f:
            json.dump({
                "run_date": datetime.now().isoformat(),
                "queries_used": len(NO_CAPTCHA_QUERIES),
                "sites_found": len(all_results),
                "sites_analyzed": len(analyzed_sites),
                "results": analyzed_sites,
            }, f, indent=2)

    # ── PHASE 3: Report ──
    print(f"\n{'='*65}")
    print(f"  PHASE 3: RESULTS")
    print(f"{'='*65}")

    # Sort by score descending
    analyzed_sites.sort(key=lambda x: x["score"], reverse=True)

    ready = [s for s in analyzed_sites if s["score"] >= 7]
    possible = [s for s in analyzed_sites if 4 <= s["score"] < 7]
    not_recommended = [s for s in analyzed_sites if s["score"] < 4]

    print(f"\n  READY TO AUTOMATE (score 7-10): {len(ready)} sites")
    print(f"  POSSIBLE (score 4-6): {len(possible)} sites")
    print(f"  NOT RECOMMENDED (score 0-3): {len(not_recommended)} sites")

    if ready:
        print(f"\n{'─'*65}")
        print(f"  🟢 TOP CANDIDATES FOR AUTOMATION")
        print(f"{'─'*65}")
        for i, site in enumerate(ready[:15]):
            print(f"\n  [{i+1}] {site['title'][:70]}")
            print(f"      URL: {site['url'][:80]}")
            print(f"      Score: {site['score']}/10")
            print(f"      Positives: {', '.join(site['positives'])}")
            print(f"      Automation: {site['automation_estimate']['difficulty']}")
            print(f"      Tool: {site['automation_estimate']['tool']}")

    # Write final report
    with open(RESULTS_FILE, "w") as f:
        json.dump({
            "run_date": datetime.now().isoformat(),
            "queries_used": len(NO_CAPTCHA_QUERIES),
            "sites_found_total": len(all_results),
            "sites_analyzed": len(analyzed_sites),
            "summary": {
                "ready_to_automate": len(ready),
                "possible": len(possible),
                "not_recommended": len(not_recommended),
            },
            "top_candidates": [
                {
                    "title": s["title"],
                    "url": s["url"],
                    "domain": s["domain"],
                    "score": s["score"],
                    "verdict": s["verdict"],
                    "positives": s["positives"],
                    "negatives": s["negatives"],
                    "automation_estimate": s["automation_estimate"],
                    "has_faucetpay": s["has_faucetpay"],
                    "has_captcha": s["has_captcha"],
                    "captcha_type": s["captcha"],
                    "recaptcha_challenge": s.get("recaptcha_challenge", "")[:150],
                }
                for s in ready[:20]
            ],
            "all_results": analyzed_sites,
        }, f, indent=2)

    print(f"\n  Results saved to: {RESULTS_FILE}")
    print(f"\n{'='*65}")
    print(f"  DONE — {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    print(f"{'='*65}")


if __name__ == "__main__":
    hunt()
