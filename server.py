"""
PinSearchPro — Flask API Server
================================
No login required — opens Pinterest explore page,
dismisses any popups, then searches normally.

Endpoints:
  POST /search            — start a search job
  GET  /status/<job_id>   — check job status + live results
  GET  /download/<job_id> — download Excel file
  GET  /health            — server health check
"""

import os, re, time, uuid, threading
from datetime import datetime
from flask import Flask, request, jsonify, send_file
from flask_cors import CORS
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from selenium import webdriver
from selenium.webdriver.chrome.options import Options as ChromeOptions
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, WebDriverException

app = Flask(__name__)
CORS(app)

jobs = {}  # in-memory job store

SCROLL_PAUSE   = 2.5
PAGE_LOAD_WAIT = 6
AHREFS_WAIT    = 10
OUTPUT_DIR     = "/tmp/pinsearch_results"
os.makedirs(OUTPUT_DIR, exist_ok=True)


# ─────────────────────────────────────────────
#  DRIVER SETUP
# ─────────────────────────────────────────────
def setup_driver():
    options = ChromeOptions()
    options.add_argument("--headless=new")
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-dev-shm-usage")
    options.add_argument("--disable-gpu")
    options.add_argument("--window-size=1920,1080")
    options.add_argument("--disable-blink-features=AutomationControlled")
    options.add_argument("--disable-extensions")
    options.add_argument("--disable-infobars")
    options.add_argument("--single-process")
    options.add_argument("--disable-setuid-sandbox")
    options.add_experimental_option("excludeSwitches", ["enable-automation"])
    options.add_experimental_option("useAutomationExtension", False)
    options.add_argument(
        "user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/120.0.0.0 Safari/537.36"
    )

    # Find Chrome binary — Railway/Nix installs it in different locations
    import shutil
    chrome_paths = [
        "/usr/bin/google-chrome",
        "/usr/bin/google-chrome-stable",
        "/usr/bin/chromium",
        "/usr/bin/chromium-browser",
        "/run/current-system/sw/bin/chromium",
        shutil.which("google-chrome") or "",
        shutil.which("chromium") or "",
        shutil.which("chromium-browser") or "",
    ]
    for path in chrome_paths:
        if path and os.path.exists(path):
            options.binary_location = path
            print(f"[+] Using Chrome at: {path}")
            break

    # Find chromedriver
    driver_paths = [
        "/usr/bin/chromedriver",
        "/run/current-system/sw/bin/chromedriver",
        shutil.which("chromedriver") or "",
    ]
    service = None
    for path in driver_paths:
        if path and os.path.exists(path):
            from selenium.webdriver.chrome.service import Service
            service = Service(executable_path=path)
            print(f"[+] Using chromedriver at: {path}")
            break

    driver = webdriver.Chrome(service=service, options=options) if service else webdriver.Chrome(options=options)
    driver.execute_script(
        "Object.defineProperty(navigator, 'webdriver', {get: () => undefined})"
    )
    driver.set_page_load_timeout(30)
    driver.set_script_timeout(30)
    return driver


# ─────────────────────────────────────────────
#  OPEN PINTEREST WITHOUT LOGIN
# ─────────────────────────────────────────────
def open_pinterest_no_login(driver, job_id):
    """
    Strategy:
    1. Open pinterest.com/ideas (explore page — works without login)
    2. If a login popup appears, close it
    3. If redirected to login page, try the explore URL directly
    """
    update_progress(job_id, "Opening Pinterest explore page...")

    # Try explore/ideas page first — usually no login required
    explore_urls = [
        "https://www.pinterest.com/ideas/",
        "https://www.pinterest.com/explore/",
        "https://www.pinterest.com/",
    ]

    for url in explore_urls:
        try:
            driver.get(url)
            time.sleep(PAGE_LOAD_WAIT)
            dismiss_all_popups(driver)

            # Check if we're on a login wall
            current = driver.current_url.lower()
            page_text = driver.page_source.lower()

            login_blocked = (
                "login" in current or
                "accounts" in current or
                ("create account" in page_text and "explore" not in current)
            )

            if not login_blocked:
                update_progress(job_id, f"Pinterest loaded via {url} ✓")
                return True

            update_progress(job_id, f"Login wall at {url} — trying next...")

        except Exception as e:
            update_progress(job_id, f"Error loading {url}: {e}")
            continue

    # Last resort — go directly to search URL (sometimes bypasses login)
    update_progress(job_id, "Using direct search URL strategy...")
    return True   # continue anyway — search URL often works


# ─────────────────────────────────────────────
#  POPUP DISMISSAL
# ─────────────────────────────────────────────
def dismiss_all_popups(driver):
    """Try every known Pinterest popup close button."""
    selectors = [
        # Login/signup modal close buttons
        "button[data-test-id='closeup-close-button']",
        "button[aria-label='Close']",
        "button[aria-label='close']",
        "div[data-test-id='interstitial-close-button']",
        "[data-test-id='login-page'] button[aria-label='Close']",
        # Generic close icons
        "button.XiG",
        "[class*='closeButton']",
        "[class*='CloseButton']",
        # 'Not now' / 'Continue as guest' type buttons
        "button[data-test-id='simple-dialog-close-button']",
    ]
    for sel in selectors:
        try:
            btn = driver.find_element(By.CSS_SELECTOR, sel)
            btn.click()
            time.sleep(0.6)
            return True
        except Exception:
            continue

    # Try clicking outside the modal (to dismiss it)
    try:
        overlay = driver.find_element(By.CSS_SELECTOR, "[class*='overlay'], [class*='Overlay']")
        overlay.click()
        time.sleep(0.6)
        return True
    except Exception:
        pass

    # Try pressing Escape
    try:
        from selenium.webdriver.common.keys import Keys
        driver.find_element(By.TAG_NAME, "body").send_keys(Keys.ESCAPE)
        time.sleep(0.5)
    except Exception:
        pass

    return False


# ─────────────────────────────────────────────
#  HELPERS
# ─────────────────────────────────────────────
def update_progress(job_id, msg):
    if job_id in jobs:
        jobs[job_id]["progress"] = msg
        print(f"[{job_id[:8]}] {msg}")

def parse_number(text):
    if not text:
        return None
    t = text.strip().upper().replace(",", "").replace("+", "").replace(" ", "")
    try:
        if "B" in t: return int(float(t.replace("B", "")) * 1_000_000_000)
        if "M" in t: return int(float(t.replace("M", "")) * 1_000_000)
        if "K" in t: return int(float(t.replace("K", "")) * 1_000)
        return int(float(re.sub(r"[^\d.]", "", t)))
    except Exception:
        return None

def fmt(n):
    if n is None: return "N/A"
    if n >= 1_000_000: return f"{n/1_000_000:.1f}M"
    if n >= 1_000: return f"{n/1_000:.1f}K"
    return str(n)


# ─────────────────────────────────────────────
#  GET WEBSITE FROM PROFILE
# ─────────────────────────────────────────────
def get_website_from_profile(driver, profile_url):
    try:
        driver.get(profile_url)
        time.sleep(PAGE_LOAD_WAIT)
        dismiss_all_popups(driver)

        for sel in [
            "a[data-test-id='profile-website']",
            "a[data-test-id='user-website-url']",
            "[data-test-id='profile-website-link'] a",
            "[class*='websiteUrl'] a",
            "[class*='website'] a",
        ]:
            try:
                el = driver.find_element(By.CSS_SELECTOR, sel)
                href = el.get_attribute("href")
                if href and "pinterest.com" not in href and href.startswith("http"):
                    return href.strip()
            except Exception:
                continue

        # Fallback: scan all external links on profile page
        skip_domains = ["pinterest.com","facebook.com","instagram.com",
                        "twitter.com","tiktok.com","youtube.com","google.com"]
        for link in driver.find_elements(By.CSS_SELECTOR, "a[href]"):
            href = link.get_attribute("href") or ""
            if (href.startswith("http") and
                    not any(d in href for d in skip_domains) and
                    re.match(r"https?://[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}", href)):
                return href.strip()
    except Exception:
        pass
    return None


# ─────────────────────────────────────────────
#  GET WEBSITE TRAFFIC VIA AHREFS
# ─────────────────────────────────────────────
def get_website_traffic(driver, website_url):
    try:
        domain = re.sub(r"https?://(www\.)?", "", website_url).split("/")[0].strip()
        if not domain:
            return None

        driver.get(f"https://ahrefs.com/traffic-checker/?target={domain}&mode=subdomains")
        time.sleep(AHREFS_WAIT)

        for sel in [
            "[data-tf='organic-traffic']",
            ".traffic-value",
            "[class*='organicTraffic']",
            "[class*='trafficValue']",
            "p[class*='value']",
            "span[class*='value']",
        ]:
            try:
                els = driver.find_elements(By.CSS_SELECTOR, sel)
                for el in els:
                    val = parse_number(el.text.strip())
                    if val and val > 0:
                        return val
            except Exception:
                continue

        # Fallback: scan full page text
        try:
            body = driver.find_element(By.TAG_NAME, "body").text
            if "cloudflare" in body.lower() or "captcha" in body.lower():
                return None
            for pattern in [
                r"Organic\s+traffic[^\d]*(\d[\d,\.]*[KMB]?)",
                r"(\d[\d,\.]*[KMB]?)\s*organic\s*(?:visitors|traffic|visits)",
                r"Traffic[^\d]*(\d[\d,\.]*[KMB]?)",
                r"(\d[\d,\.]*[KMB]?)\s*visitors",
            ]:
                m = re.search(pattern, body, re.IGNORECASE)
                if m:
                    val = parse_number(m.group(1))
                    if val is not None:
                        return val
        except Exception:
            pass
    except Exception:
        pass
    return None


# ─────────────────────────────────────────────
#  SCRAPE PINS
# ─────────────────────────────────────────────
def scrape_pins(driver, keyword, min_views_m, min_traffic, max_pins, job_id):
    results = []
    seen_profiles = set()
    min_views_raw = int(min_views_m * 1_000_000)

    try:
        # Go directly to Pinterest search — works without login
        search_url = (
            f"https://www.pinterest.com/search/pins/"
            f"?q={keyword.replace(' ', '%20')}&rs=typed"
        )
        update_progress(job_id, f"Searching Pinterest for '{keyword}'...")
        driver.get(search_url)
        time.sleep(PAGE_LOAD_WAIT)
        dismiss_all_popups(driver)

        # Collect pin URLs by scrolling
        pins = []
        attempts = 0
        while len(pins) < max_pins and attempts < 8:
            links = driver.find_elements(By.CSS_SELECTOR, "a[href*='/pin/']")
            for lnk in links:
                href = lnk.get_attribute("href")
                if href and "/pin/" in href and href not in [p["url"] for p in pins]:
                    pins.append({"url": href})
                if len(pins) >= max_pins:
                    break
            if len(pins) < max_pins:
                driver.execute_script("window.scrollBy(0, 1500);")
                time.sleep(SCROLL_PAUSE)
                dismiss_all_popups(driver)
                attempts += 1

        pins = pins[:max_pins]
        total = len(pins)
        update_progress(job_id, f"Found {total} pins — now checking each profile...")

        for i, pin in enumerate(pins):
            update_progress(job_id, f"Checking pin {i+1}/{total}...")

            try:
                try:
                    driver.get(pin["url"])
                except (TimeoutException, WebDriverException):
                    continue
                time.sleep(3)
                dismiss_all_popups(driver)

                # Find profile link on pin page
                profile_url = None
                profile_name = None
                for sel in [
                    "a[data-test-id='creator-profile-link']",
                    "[data-test-id='pin-closeup-user'] a",
                    "a[data-test-id='user-rep-link']",
                ]:
                    try:
                        el = driver.find_element(By.CSS_SELECTOR, sel)
                        profile_url  = el.get_attribute("href")
                        profile_name = el.text.strip()
                        if profile_url:
                            break
                    except Exception:
                        continue

                # Fallback profile detection
                if not profile_url:
                    for lnk in driver.find_elements(By.CSS_SELECTOR, "a[href]"):
                        href = lnk.get_attribute("href") or ""
                        parts = href.rstrip("/").split("/")
                        if (href.startswith("https://www.pinterest.com/") and
                                "/pin/" not in href and len(parts) == 5):
                            profile_url  = href
                            profile_name = parts[-1]
                            break

                if not profile_url or profile_url in seen_profiles:
                    continue

                seen_profiles.add(profile_url)
                clean_profile = profile_url.split("?")[0].rstrip("/")

                # Visit profile page
                try:
                    driver.get(clean_profile)
                except (TimeoutException, WebDriverException):
                    continue
                time.sleep(3)
                dismiss_all_popups(driver)

                # Get monthly views
                view_text = None
                for sel in [
                    "[data-test-id='profile-monthly-views']",
                    "[class*='monthlyViews']",
                    "[class*='monthly']",
                ]:
                    try:
                        el = driver.find_element(By.CSS_SELECTOR, sel)
                        view_text = el.text.strip()
                        if view_text:
                            break
                    except Exception:
                        continue

                if not view_text:
                    try:
                        body = driver.find_element(By.TAG_NAME, "body").text
                        m = re.search(
                            r"([\d,.]+\s*[KMBkmb]?)\s*(?:monthly\s*views|Monthly\s*Views)",
                            body
                        )
                        if m:
                            view_text = m.group(1).strip()
                    except Exception:
                        pass

                view_int = parse_number(view_text)

                if not profile_name:
                    try:
                        profile_name = driver.find_element(
                            By.CSS_SELECTOR, "h1, [data-test-id='profile-name']"
                        ).text.strip()
                    except Exception:
                        profile_name = clean_profile.rstrip("/").split("/")[-1]

                update_progress(job_id,
                    f"Pin {i+1}/{total} — {profile_name} — "
                    f"views: {fmt(view_int)} (need {fmt(min_views_raw)})")

                # FILTER 1: Pinterest views
                if not view_int or view_int < min_views_raw:
                    update_progress(job_id, f"Pin {i+1}/{total} — {profile_name} ✗ views too low")
                    continue

                # Get website
                update_progress(job_id, f"Pin {i+1}/{total} — {profile_name} ✓ views — checking website...")
                website_url = get_website_from_profile(driver, clean_profile)

                if not website_url:
                    update_progress(job_id, f"Pin {i+1}/{total} — {profile_name} ✗ no website")
                    continue

                # FILTER 2: Website traffic
                update_progress(job_id, f"Pin {i+1}/{total} — {profile_name} — checking traffic for {website_url}...")
                traffic = get_website_traffic(driver, website_url)

                if traffic is not None and traffic < min_traffic:
                    update_progress(job_id, f"Pin {i+1}/{total} — {profile_name} ✗ traffic too low ({fmt(traffic)})")
                    continue

                result = {
                    "profile_name":        profile_name,
                    "profile_url":         clean_profile,
                    "pinterest_views":     fmt(view_int),
                    "pinterest_views_int": view_int,
                    "website_url":         website_url,
                    "website_traffic":     fmt(traffic) if traffic else "Unknown",
                    "website_traffic_int": traffic or 0,
                }
                results.append(result)
                jobs[job_id]["results"] = list(results)

                update_progress(job_id,
                    f"✓ {profile_name} qualified! "
                    f"({len(results)} profiles found so far)")

            except Exception as e:
                update_progress(job_id, f"Pin {i+1} error: {e}")
                continue

    except Exception as e:
        update_progress(job_id, f"Search error: {e}")

    return results


# ─────────────────────────────────────────────
#  EXCEL EXPORT
# ─────────────────────────────────────────────
def save_to_excel(results, keyword, job_id):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Pinterest Research"

    RED="E60023"; WHITE="FFFFFF"; PINK="FFF0F0"; GRAY="F5F5F5"; BLUE="0563C1"
    hdr_font  = Font(name="Calibri", bold=True, color=WHITE, size=11)
    hdr_fill  = PatternFill(start_color=RED,  end_color=RED,  fill_type="solid")
    main_fill = PatternFill(start_color=PINK, end_color=PINK, fill_type="solid")
    center    = Alignment(horizontal="center", vertical="center", wrap_text=True)
    left      = Alignment(horizontal="left",   vertical="center", wrap_text=True)
    thin      = Side(style="thin", color="DDDDDD")
    border    = Border(left=thin, right=thin, top=thin, bottom=thin)

    ws.merge_cells("A1:G1")
    ws["A1"].value     = f"Pinterest Research · {keyword.title()} · {datetime.now().strftime('%Y-%m-%d %H:%M')}"
    ws["A1"].font      = Font(name="Calibri", bold=True, size=13, color=RED)
    ws["A1"].alignment = center
    ws.row_dimensions[1].height = 28

    ws.merge_cells("A2:G2")
    ws["A2"].value     = f"Total qualifying profiles: {len(results)}"
    ws["A2"].font      = Font(name="Calibri", italic=True, size=10, color="666666")
    ws["A2"].alignment = center
    ws.row_dimensions[2].height = 18

    headers    = ["Profile Name","Profile URL","Pinterest Views","Website URL","Website Traffic","Keyword"]
    col_widths = [24, 42, 18, 42, 18, 24]
    for col, (h, w) in enumerate(zip(headers, col_widths), 1):
        cell = ws.cell(row=3, column=col, value=h)
        cell.font=hdr_font; cell.fill=hdr_fill
        cell.alignment=center; cell.border=border
        ws.column_dimensions[cell.column_letter].width = w
    ws.row_dimensions[3].height = 22

    for ri, r in enumerate(results, start=4):
        row_data = [
            r["profile_name"], r["profile_url"], r["pinterest_views"],
            r["website_url"], r["website_traffic"], keyword.title(),
        ]
        for col, val in enumerate(row_data, 1):
            cell = ws.cell(row=ri, column=col, value=val)
            cell.font=Font(name="Calibri", size=11)
            cell.fill=main_fill; cell.border=border
            cell.alignment=center if col in (3, 5) else left
            if col in (2, 4) and val and str(val).startswith("http"):
                cell.hyperlink = val
                cell.font=Font(name="Calibri", size=11, color=BLUE, underline="single")
        ws.row_dimensions[ri].height = 18

    ws.freeze_panes = "A4"
    ws.auto_filter.ref = f"A3:F{2+len(results)}"

    path = os.path.join(OUTPUT_DIR, f"{job_id}.xlsx")
    wb.save(path)
    return path


# ─────────────────────────────────────────────
#  BACKGROUND JOB WORKER
# ─────────────────────────────────────────────
def run_search_job(job_id, keyword, min_views_m, min_traffic, max_pins):
    jobs[job_id]["status"] = "running"
    driver = None
    try:
        update_progress(job_id, "Starting headless browser...")
        driver = setup_driver()

        # Open Pinterest without login
        open_pinterest_no_login(driver, job_id)

        # Run the scrape
        results = scrape_pins(driver, keyword, min_views_m, min_traffic, max_pins, job_id)

        # Deduplicate by website URL
        seen, unique = set(), []
        for r in results:
            key = r.get("website_url","").lower().rstrip("/")
            if key and key not in seen:
                seen.add(key)
                unique.append(r)

        unique.sort(key=lambda x: -x["pinterest_views_int"])
        jobs[job_id]["results"] = unique

        if unique:
            path = save_to_excel(unique, keyword, job_id)
            jobs[job_id]["excel_path"] = path
            update_progress(job_id, f"Done! {len(unique)} profiles found. Excel ready to download.")
        else:
            update_progress(job_id, "Done — no qualifying profiles found for this keyword.")

        jobs[job_id]["status"] = "done"

    except Exception as e:
        jobs[job_id]["status"]   = "error"
        jobs[job_id]["error"]    = str(e)
        update_progress(job_id, f"Error: {e}")
    finally:
        if driver:
            try: driver.quit()
            except Exception: pass


# ─────────────────────────────────────────────
#  API ROUTES
# ─────────────────────────────────────────────
@app.route("/health", methods=["GET"])
def health():
    return jsonify({"status": "ok", "message": "PinSearchPro API running"})


@app.route("/search", methods=["POST"])
def start_search():
    """
    POST /search
    {
      "keyword":     "home decor",
      "min_views":   2000000,
      "min_traffic": 200,
      "max_pins":    20
    }
    """
    data = request.get_json()
    if not data or not data.get("keyword"):
        return jsonify({"error": "keyword is required"}), 400

    keyword     = data["keyword"].strip().lower()
    min_views_m = float(data.get("min_views", 2_000_000)) / 1_000_000
    min_traffic = int(data.get("min_traffic", 200))
    max_pins    = max(20, min(int(data.get("max_pins", 20)), 200))

    job_id = str(uuid.uuid4())
    jobs[job_id] = {
        "status":     "queued",
        "progress":   "Queued...",
        "results":    [],
        "excel_path": None,
        "error":      None,
        "keyword":    keyword,
        "created_at": datetime.now().isoformat(),
    }

    threading.Thread(
        target=run_search_job,
        args=(job_id, keyword, min_views_m, min_traffic, max_pins),
        daemon=True
    ).start()

    return jsonify({"job_id": job_id, "status": "queued"})


@app.route("/status/<job_id>", methods=["GET"])
def get_status(job_id):
    job = jobs.get(job_id)
    if not job:
        return jsonify({"error": "Job not found"}), 404
    return jsonify({
        "job_id":       job_id,
        "status":       job["status"],
        "progress":     job["progress"],
        "result_count": len(job["results"]),
        "results":      job["results"],
        "has_excel":    job["excel_path"] is not None,
        "error":        job["error"],
    })


@app.route("/download/<job_id>", methods=["GET"])
def download_excel(job_id):
    job = jobs.get(job_id)
    if not job:
        return jsonify({"error": "Job not found"}), 404
    if not job.get("excel_path") or not os.path.exists(job["excel_path"]):
        return jsonify({"error": "Excel not ready yet"}), 404
    return send_file(
        job["excel_path"],
        as_attachment=True,
        download_name=f"pinterest_{job['keyword'].replace(' ','_')}_results.xlsx",
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port, debug=False)
