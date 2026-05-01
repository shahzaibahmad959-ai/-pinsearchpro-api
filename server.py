"""
PinSearchPro — No-Browser API Server
======================================
Uses Pinterest's internal API + HTTP requests.
No Selenium, no Chrome, no Firefox needed.
Fast, lightweight, works on any server.
"""

import os, re, time, uuid, threading, requests
from datetime import datetime
from flask import Flask, request, jsonify, send_file
from flask_cors import CORS
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

app = Flask(__name__)
CORS(app)

jobs = {}
OUTPUT_DIR = "/tmp/pinsearch_results"
os.makedirs(OUTPUT_DIR, exist_ok=True)

# ── Pinterest session (set in Railway environment variables) ──
# Get this from your browser:
# 1. Open pinterest.com (logged in)
# 2. Press F12 → Network tab → refresh page
# 3. Click any request → Headers → Cookie
# 4. Copy the full cookie string
# Set it as PINTEREST_COOKIE env var in Railway
PINTEREST_COOKIE = os.environ.get("PINTEREST_COOKIE", "")

# Pinterest API base URL
PIN_API = "https://www.pinterest.com/resource"

HEADERS = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36",
    "Accept": "application/json, text/javascript, */*, q=0.01",
    "Accept-Language": "en-US,en;q=0.9",
    "X-Requested-With": "XMLHttpRequest",
    "Referer": "https://www.pinterest.com/",
}


# ─────────────────────────────────────────────
#  HELPERS
# ─────────────────────────────────────────────
def update_progress(job_id, msg):
    if job_id in jobs:
        jobs[job_id]["progress"] = msg
        print(f"[{job_id[:8]}] {msg}")

def parse_number(text):
    if not text: return None
    t = str(text).strip().upper().replace(",","").replace("+","").replace(" ","")
    try:
        if "B" in t: return int(float(t.replace("B","")) * 1_000_000_000)
        if "M" in t: return int(float(t.replace("M","")) * 1_000_000)
        if "K" in t: return int(float(t.replace("K","")) * 1_000)
        return int(float(re.sub(r"[^\d.]","",t)))
    except: return None

def fmt(n):
    if n is None: return "N/A"
    if n >= 1_000_000: return f"{n/1_000_000:.1f}M"
    if n >= 1_000: return f"{n/1_000:.1f}K"
    return str(n)

def get_session():
    """Create a requests session with Pinterest headers."""
    s = requests.Session()
    s.headers.update(HEADERS)
    if PINTEREST_COOKIE:
        s.headers["Cookie"] = PINTEREST_COOKIE
        # Extract CSRF token from cookie
        m = re.search(r"csrftoken=([^;]+)", PINTEREST_COOKIE)
        if m:
            s.headers["X-CSRFToken"] = m.group(1)
    return s


# ─────────────────────────────────────────────
#  PINTEREST API CALLS
# ─────────────────────────────────────────────
def search_pins(session, keyword, max_pins=20):
    """Search Pinterest for pins using their internal API."""
    pins = []
    bookmark = None
    page_size = 25

    update_progress_global = lambda msg: print(f"[search] {msg}")

    while len(pins) < max_pins:
        params = {
            "source_url": f"/search/pins/?q={keyword}&rs=typed",
            "data": '{"options":{"query":"' + keyword + '","scope":"pins","page_size":' + str(page_size) +
                    (',"bookmarks":["' + bookmark + '"]' if bookmark else '') +
                    '},"context":{}}',
            "_": int(time.time() * 1000)
        }

        try:
            resp = session.get(f"{PIN_API}/BaseSearchResource/get/", params=params, timeout=15)
            if resp.status_code != 200:
                break
            data = resp.json()
            resource_response = data.get("resource_response", {})
            pin_data = resource_response.get("data", {}).get("results", [])

            if not pin_data:
                break

            for pin in pin_data:
                if len(pins) >= max_pins:
                    break
                pin_id   = pin.get("id")
                pinner   = pin.get("pinner", {})
                if pin_id and pinner:
                    pins.append({
                        "pin_id":       pin_id,
                        "profile_id":   pinner.get("id"),
                        "profile_name": pinner.get("full_name") or pinner.get("username",""),
                        "username":     pinner.get("username",""),
                        "monthly_views": pinner.get("monthly_views", 0),
                        "website_url":  pinner.get("website_url","") or pinner.get("domain_url",""),
                    })

            # Get next page bookmark
            bookmark = resource_response.get("bookmark")
            if not bookmark:
                break

            time.sleep(0.5)  # be polite

        except Exception as e:
            print(f"Search error: {e}")
            break

    return pins[:max_pins]


def get_profile_details(session, username):
    """Get full profile details including monthly views and website."""
    try:
        params = {
            "source_url": f"/{username}/",
            "data": f'{{"options":{{"username":"{username}"}},"context":{{}}}}',
            "_": int(time.time() * 1000)
        }
        resp = session.get(f"{PIN_API}/UserResource/get/", params=params, timeout=15)
        if resp.status_code != 200:
            return None
        data = resp.json()
        user = data.get("resource_response", {}).get("data", {})
        return {
            "username":     user.get("username",""),
            "full_name":    user.get("full_name",""),
            "monthly_views": user.get("monthly_views", 0),
            "website_url":  user.get("website_url","") or user.get("domain_url",""),
            "profile_url":  f"https://www.pinterest.com/{user.get('username','')}/",
        }
    except Exception as e:
        print(f"Profile error for {username}: {e}")
        return None


def get_website_traffic_api(website_url):
    """
    Check website traffic using SimilarWeb's free public data.
    Falls back to a simple check if unavailable.
    """
    if not website_url:
        return None
    try:
        domain = re.sub(r"https?://(www\.)?","", website_url).split("/")[0].strip()
        if not domain:
            return None

        # Try SimilarWeb public API
        headers = {
            "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36",
            "Accept": "application/json",
        }
        resp = requests.get(
            f"https://data.similarweb.com/api/v1/data?domain={domain}",
            headers=headers, timeout=10
        )
        if resp.status_code == 200:
            data = resp.json()
            visits = data.get("EstimatedMonthlyVisits", {})
            if visits:
                # Get most recent month's visits
                latest = list(visits.values())[-1] if visits else None
                if latest:
                    return int(latest)

        # Fallback — just return None (website exists but traffic unknown)
        return None

    except Exception:
        return None


# ─────────────────────────────────────────────
#  EXCEL EXPORT
# ─────────────────────────────────────────────
def save_to_excel(results, keyword, job_id):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Pinterest Research"

    RED="E60023"; WHITE="FFFFFF"; PINK="FFF0F0"; BLUE="0563C1"
    hdr_font  = Font(name="Calibri", bold=True, color=WHITE, size=11)
    hdr_fill  = PatternFill(start_color=RED,  end_color=RED,  fill_type="solid")
    row_fill  = PatternFill(start_color=PINK, end_color=PINK, fill_type="solid")
    center    = Alignment(horizontal="center", vertical="center", wrap_text=True)
    left      = Alignment(horizontal="left",   vertical="center", wrap_text=True)
    thin      = Side(style="thin", color="DDDDDD")
    border    = Border(left=thin, right=thin, top=thin, bottom=thin)

    # Title row
    ws.merge_cells("A1:F1")
    ws["A1"].value     = f"Pinterest Research · {keyword.title()} · {datetime.now().strftime('%Y-%m-%d %H:%M')}"
    ws["A1"].font      = Font(name="Calibri", bold=True, size=13, color=RED)
    ws["A1"].alignment = center
    ws.row_dimensions[1].height = 28

    ws.merge_cells("A2:F2")
    ws["A2"].value     = f"Total qualifying profiles: {len(results)}"
    ws["A2"].font      = Font(name="Calibri", italic=True, size=10, color="666666")
    ws["A2"].alignment = center
    ws.row_dimensions[2].height = 18

    headers    = ["Profile Name", "Pinterest URL", "Monthly Views", "Website", "Traffic", "Keyword"]
    col_widths = [24, 40, 18, 40, 16, 20]
    for col, (h, w) in enumerate(zip(headers, col_widths), 1):
        cell = ws.cell(row=3, column=col, value=h)
        cell.font=hdr_font; cell.fill=hdr_fill
        cell.alignment=center; cell.border=border
        ws.column_dimensions[cell.column_letter].width = w
    ws.row_dimensions[3].height = 22

    for ri, r in enumerate(results, start=4):
        row_data = [
            r["profile_name"], r["profile_url"],
            r["pinterest_views"], r["website_url"],
            r["website_traffic"], keyword.title(),
        ]
        for col, val in enumerate(row_data, 1):
            cell = ws.cell(row=ri, column=col, value=val)
            cell.font=Font(name="Calibri", size=11)
            cell.fill=row_fill; cell.border=border
            cell.alignment=center if col in (3,5) else left
            if col in (2,4) and val and str(val).startswith("http"):
                cell.hyperlink = val
                cell.font=Font(name="Calibri", size=11, color=BLUE, underline="single")
        ws.row_dimensions[ri].height = 18

    ws.freeze_panes = "A4"
    ws.auto_filter.ref = f"A3:F{2+len(results)}"

    path = os.path.join(OUTPUT_DIR, f"{job_id}.xlsx")
    wb.save(path)
    return path


# ─────────────────────────────────────────────
#  MAIN SEARCH JOB
# ─────────────────────────────────────────────
def run_search_job(job_id, keyword, min_views, min_traffic, max_pins):
    jobs[job_id]["status"] = "running"
    results = []
    seen_profiles = set()

    try:
        session = get_session()
        update_progress(job_id, f"Searching Pinterest for '{keyword}'...")

        # Step 1: Get pins from search
        pins = search_pins(session, keyword, max_pins)
        total = len(pins)
        update_progress(job_id, f"Found {total} pins — checking profiles...")

        if total == 0:
            update_progress(job_id, "No pins found. Try a different keyword or add Pinterest cookie.")
            jobs[job_id]["status"] = "done"
            return

        # Step 2: Check each profile
        for i, pin in enumerate(pins):
            username = pin.get("username","")
            if not username or username in seen_profiles:
                continue
            seen_profiles.add(username)

            update_progress(job_id, f"Checking profile {i+1}/{total}: {username}...")

            # Get full profile details
            profile = get_profile_details(session, username)
            if not profile:
                profile = pin  # fallback to pin data

            monthly_views = profile.get("monthly_views") or pin.get("monthly_views") or 0
            if isinstance(monthly_views, str):
                monthly_views = parse_number(monthly_views) or 0

            update_progress(job_id,
                f"Profile {i+1}/{total}: {username} — "
                f"views: {fmt(monthly_views)} (need {fmt(min_views)})")

            # FILTER 1: Pinterest views
            if monthly_views < min_views:
                continue

            # Get website
            website_url = (profile.get("website_url") or pin.get("website_url","")).strip()
            if not website_url:
                update_progress(job_id, f"{username} ✗ no website")
                continue

            # Clean website URL
            if not website_url.startswith("http"):
                website_url = "https://" + website_url

            # FILTER 2: Traffic (optional — skip if min_traffic = 0)
            traffic = None
            if min_traffic > 0:
                update_progress(job_id, f"{username} ✓ views — checking traffic...")
                traffic = get_website_traffic_api(website_url)
                if traffic is not None and traffic < min_traffic:
                    update_progress(job_id, f"{username} ✗ traffic too low ({fmt(traffic)})")
                    continue

            profile_url = profile.get("profile_url") or f"https://www.pinterest.com/{username}/"
            profile_name = profile.get("full_name") or profile.get("profile_name") or username

            result = {
                "profile_name":        profile_name,
                "profile_url":         profile_url,
                "pinterest_views":     fmt(monthly_views),
                "pinterest_views_int": monthly_views,
                "website_url":         website_url,
                "website_traffic":     fmt(traffic) if traffic else "Unknown",
                "website_traffic_int": traffic or 0,
            }
            results.append(result)
            jobs[job_id]["results"] = list(results)
            update_progress(job_id, f"✓ {profile_name} qualified! ({len(results)} found so far)")

            time.sleep(0.3)  # small delay between requests

        # Sort by views
        results.sort(key=lambda x: -x["pinterest_views_int"])
        jobs[job_id]["results"] = results

        if results:
            path = save_to_excel(results, keyword, job_id)
            jobs[job_id]["excel_path"] = path
            update_progress(job_id, f"Done! {len(results)} profiles found. Excel ready.")
        else:
            update_progress(job_id,
                "Done — no qualifying profiles found. "
                "Try lower min views or add Pinterest cookie for better results.")

        jobs[job_id]["status"] = "done"

    except Exception as e:
        jobs[job_id]["status"] = "error"
        jobs[job_id]["error"]  = str(e)
        update_progress(job_id, f"Error: {e}")


# ─────────────────────────────────────────────
#  API ROUTES
# ─────────────────────────────────────────────
@app.route("/health", methods=["GET"])
def health():
    has_cookie = bool(PINTEREST_COOKIE)
    return jsonify({
        "status": "ok",
        "message": "PinSearchPro API running (no-browser mode)",
        "pinterest_cookie_set": has_cookie,
    })


@app.route("/search", methods=["POST"])
def start_search():
    data = request.get_json()
    if not data or not data.get("keyword"):
        return jsonify({"error": "keyword is required"}), 400

    keyword     = data["keyword"].strip().lower()
    min_views   = int(data.get("min_views", 2_000_000))
    min_traffic = int(data.get("min_traffic", 0))  # 0 = skip traffic check
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
        args=(job_id, keyword, min_views, min_traffic, max_pins),
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
        return jsonify({"error": "Excel not ready"}), 404
    return send_file(
        job["excel_path"],
        as_attachment=True,
        download_name=f"pinterest_{job['keyword'].replace(' ','_')}_results.xlsx",
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port, debug=False)
