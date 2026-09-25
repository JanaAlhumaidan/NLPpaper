"""
Google Maps Review Scraper (multi-place, one run)
--------------------------------------------------
Reads every place from your Excel sheet and scrapes all of them in a single run.

Setup:
    pip install playwright pandas openpyxl
    playwright install chromium

Usage:
    1. Put your Excel file in the same folder as this script (default name: places.xlsx).
    2. Check the CONFIG section: the column names must match your sheet's headers.
    3. Run:  python gmaps_multi_scraper.py

Output:
    raw_reviews/<place_id>.csv   one raw file per place (never edited afterwards)
    all_reviews.csv              all places combined
    scrape_summary.csv           one row per place: status + number of reviews collected

Resume:
    If the run stops (crash, blocked, laptop sleeps), just run it again.
    Places marked "done" in scrape_summary.csv are skipped.
"""

import csv
import os
import random
import re
import time
from datetime import datetime

import pandas as pd
from playwright.sync_api import sync_playwright

# ============================== CONFIG ==============================
PLACES_FILE = "places.xlsx"
SHEET_NAME = 0                     # first sheet; or use the sheet name, e.g. "الأماكن"

# Column headers in your Excel sheet (must match exactly)
COL_PLACE_ID = "Place ID"          # optional; auto-generated (P01, P02...) if missing
COL_NAME = "اسم المكان"
COL_CITY = "المدينة"
COL_CATEGORY = "النوع"
COL_URL = "google maps link"

OUTPUT_DIR = "raw_reviews"
COMBINED_FILE = "all_reviews.csv"
SUMMARY_FILE = "scrape_summary.csv"

MAX_REVIEWS_PER_PLACE = None       # e.g. 2000 to cap big places, or None for everything
SORT_BY_NEWEST = True
NO_NEW_LIMIT = 8                   # stop a place after this many scrolls with no new reviews
MAX_SCROLLS = 3000                 # safety limit per place
SCROLL_PAUSE = 2.0                 # seconds between scrolls (increase if reviews load slowly)
PAUSE_BETWEEN_PLACES = (8, 15)     # random pause range in seconds, to avoid getting blocked
HEADLESS = False                   # True = run without showing the browser
# ====================================================================

FIELDS = [
    "place_id", "place_name", "city", "category", "place_url",
    "review_id", "reviewer", "reviewer_profile_url",
    "rating", "date_relative", "review_text", "was_translated",
    "original_confirmed", "text_truncated", "scraped_at",
]


# ---------------------------- Places sheet ----------------------------

def load_places():
    df = pd.read_excel(PLACES_FILE, sheet_name=SHEET_NAME)
    df.columns = [str(c).strip() for c in df.columns]

    for col in (COL_NAME, COL_URL):
        if col not in df.columns:
            raise SystemExit(
                f"Column '{col}' not found in {PLACES_FILE}.\n"
                f"Columns found: {list(df.columns)}\n"
                f"Fix the names in the CONFIG section."
            )

    # City cells are merged in the sheet, so fill them down
    if COL_CITY in df.columns:
        df[COL_CITY] = df[COL_CITY].ffill()

    places = []
    for i, row in df.iterrows():
        url = row.get(COL_URL)
        if pd.isna(url) or not str(url).strip().startswith("http"):
            continue  # skip places without a link yet

        pid = row.get(COL_PLACE_ID) if COL_PLACE_ID in df.columns else None
        if pid is None or pd.isna(pid) or str(pid).strip() == "":
            pid = f"P{i + 1:02d}"  # based on row order, so don't reorder the sheet

        places.append({
            "place_id": str(pid).strip(),
            "place_name": clean(row.get(COL_NAME)),
            "city": clean(row.get(COL_CITY)),
            "category": clean(row.get(COL_CATEGORY)),
            "place_url": str(url).strip(),
        })
    return places


def clean(value):
    return "" if value is None or pd.isna(value) else str(value).strip()


# ---------------------------- Summary / resume ----------------------------

def load_summary():
    if not os.path.exists(SUMMARY_FILE):
        return {}
    with open(SUMMARY_FILE, newline="", encoding="utf-8-sig") as f:
        return {r["place_id"]: r for r in csv.DictReader(f)}


def save_summary(summary):
    fields = ["place_id", "place_name", "city", "category", "place_url",
              "reviews_collected", "status", "finished_at"]
    with open(SUMMARY_FILE, "w", newline="", encoding="utf-8-sig") as f:
        writer = csv.DictWriter(f, fieldnames=fields)
        writer.writeheader()
        writer.writerows(summary.values())


# ---------------------------- Page helpers ----------------------------

def force_english(url):
    if "hl=" in url:
        return re.sub(r"hl=[^&]*", "hl=en", url)
    return url + ("&" if "?" in url else "?") + "hl=en"


def accept_cookies_if_present(page):
    for text in ["Accept all", "I agree", "Reject all"]:
        try:
            btn = page.get_by_role("button", name=text)
            if btn.count() > 0:
                btn.first.click(timeout=3000)
                page.wait_for_timeout(1500)
                return
        except Exception:
            pass


def open_place(page, url):
    page.goto(url, wait_until="domcontentloaded", timeout=60000)
    page.wait_for_timeout(3000)
    # Short links redirect to the full URL; reload it with an English interface
    # (the selectors below depend on English labels; review text stays original)
    if "hl=en" not in page.url:
        page.goto(force_english(page.url), wait_until="domcontentloaded", timeout=60000)
        page.wait_for_timeout(3000)
    accept_cookies_if_present(page)


def open_reviews_tab(page):
    attempts = [
        lambda: page.locator('button[aria-label^="Reviews for"]').first.click(timeout=15000),
        lambda: page.get_by_role("tab", name=re.compile("review", re.I)).first.click(timeout=8000),
        lambda: page.locator("text=/[\\d,]+\\s+review/i").first.click(timeout=8000),
    ]
    for attempt in attempts:
        try:
            attempt()
            page.wait_for_timeout(2000)
            return
        except Exception:
            continue
    raise RuntimeError("Could not open the Reviews tab")


def sort_by_newest(page):
    try:
        page.locator('button[aria-label*="Sort"]').first.click(timeout=5000)
        page.wait_for_timeout(1000)
        page.get_by_role("menuitemradio", name=re.compile("Newest", re.I)).first.click(timeout=5000)
        page.wait_for_timeout(2500)
    except Exception:
        print("    (could not sort by newest; using default order)")


def get_scroll_container(page):
    box = page.locator("div.m6QErb.DxyBCb.kA9KIf.dS8AEf").first
    if box.count() == 0:
        box = page.locator("div.jftiEf").first.locator(
            "xpath=ancestor::div[contains(@class,'m6QErb')]"
        ).first
    return box


def safe_text(locator):
    try:
        return locator.first.inner_text(timeout=300).strip()
    except Exception:
        return ""


def safe_attr(locator, name):
    try:
        return locator.first.get_attribute(name, timeout=300) or ""
    except Exception:
        return ""


# ---------------------------- Review extraction ----------------------------

def reveal_original(card):
    """Switch a Google-translated review back to its original language.
    Returns (was_translated, original_confirmed).
    original_confirmed is True only when the 'See translation' button appears,
    which proves the text on screen is now the original."""
    see_original = card.locator("button").filter(has_text=re.compile("see original", re.I))
    if see_original.count() == 0:
        return False, True  # not translated, text is already original

    see_translation = card.locator("button").filter(has_text=re.compile("see translation", re.I))
    for _ in range(3):
        try:
            card.scroll_into_view_if_needed(timeout=1500)
            see_original.first.click(timeout=2000)
        except Exception:
            pass
        try:
            see_translation.first.wait_for(timeout=2500)
            return True, True
        except Exception:
            continue
    return True, False


def expand_more(card):
    """Click 'More' so long reviews are not cut off with '…'."""
    more = card.locator('button[aria-label="See more"], button.w8nwRe')
    for _ in range(2):
        if more.count() == 0:
            return
        try:
            card.scroll_into_view_if_needed(timeout=1500)
            more.first.click(timeout=2000)
            time.sleep(0.4)
        except Exception:
            pass


def extract_review(card, review_id, place, scraped_at):
    was_translated, original_confirmed = reveal_original(card)
    expand_more(card)  # after switching to original, since the card re-renders

    rating_label = safe_attr(card.locator("span.kvMYJc"), "aria-label")
    if not rating_label:
        rating_label = safe_text(card.locator("span.fzvQIb"))  # alternative layout, e.g. "5/5"
    m = re.search(r"\d+(\.\d+)?", rating_label)

    profile = safe_attr(card.locator("button[data-href]"), "data-href")
    if not profile:
        profile = safe_attr(card.locator("a[href*='contrib']"), "href")

    text = safe_text(card.locator("span.wiI7pd"))

    return {
        "place_id": place["place_id"],
        "place_name": place["place_name"],
        "city": place["city"],
        "category": place["category"],
        "place_url": place["place_url"],
        "review_id": review_id,
        "reviewer": safe_text(card.locator("div.d4r55")),
        "reviewer_profile_url": profile,
        "rating": m.group(0) if m else "",
        "date_relative": safe_text(card.locator("span.rsqaWe")),
        "review_text": text,
        "was_translated": was_translated,
        "original_confirmed": original_confirmed,
        "text_truncated": text.rstrip().endswith("…"),
        "scraped_at": scraped_at,
    }


def scrape_place(page, place, out_path):
    open_place(page, place["place_url"])
    open_reviews_tab(page)
    if SORT_BY_NEWEST:
        sort_by_newest(page)
    page.wait_for_selector("div.jftiEf", timeout=15000)
    container = get_scroll_container(page)

    scraped_at = datetime.now().strftime("%Y-%m-%d")
    seen = set()
    count = 0
    no_new = 0

    with open(out_path, "w", newline="", encoding="utf-8-sig") as f:
        writer = csv.DictWriter(f, fieldnames=FIELDS)
        writer.writeheader()

        for scroll in range(MAX_SCROLLS):
            ids = page.eval_on_selector_all(
                "div.jftiEf",
                "els => els.map(e => e.getAttribute('data-review-id') || '')",
            )
            cards = page.locator("div.jftiEf")
            new_this_round = 0

            for i, rid in enumerate(ids):
                key = rid or f"idx-{i}"
                if key in seen:
                    continue
                seen.add(key)
                try:
                    review = extract_review(cards.nth(i), rid, place, scraped_at)
                except Exception:
                    continue
                if review["review_text"] or review["rating"]:
                    writer.writerow(review)
                    count += 1
                    new_this_round += 1
                if MAX_REVIEWS_PER_PLACE and count >= MAX_REVIEWS_PER_PLACE:
                    return count

            f.flush()
            no_new = 0 if new_this_round else no_new + 1
            if scroll % 10 == 0:
                print(f"    scroll {scroll}: {count} reviews")
            if no_new >= NO_NEW_LIMIT:
                break

            try:
                container.evaluate("el => { el.scrollTop = el.scrollHeight; }")
            except Exception:
                page.mouse.wheel(0, 3000)
            time.sleep(SCROLL_PAUSE)

    return count


# ---------------------------- Main ----------------------------

def combine_outputs(places):
    frames = []
    for place in places:
        path = os.path.join(OUTPUT_DIR, f"{place['place_id']}.csv")
        if os.path.exists(path):
            frames.append(pd.read_csv(path, encoding="utf-8-sig", dtype=str))
    if frames:
        combined = pd.concat(frames, ignore_index=True)
        combined.to_csv(COMBINED_FILE, index=False, encoding="utf-8-sig")
        print(f"\nCombined file: {COMBINED_FILE} ({len(combined)} reviews)")


def main():
    os.makedirs(OUTPUT_DIR, exist_ok=True)
    places = load_places()
    summary = load_summary()
    print(f"Found {len(places)} places with links.\n")

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=HEADLESS)
        context = browser.new_context(
            locale="en-US",
            extra_http_headers={"Accept-Language": "en-US,en;q=0.9"},
        )

        for n, place in enumerate(places, 1):
            pid = place["place_id"]
            if summary.get(pid, {}).get("status") == "done":
                print(f"[{n}/{len(places)}] {pid} {place['place_name']}: already done, skipping")
                continue

            print(f"[{n}/{len(places)}] {pid} {place['place_name']}")
            out_path = os.path.join(OUTPUT_DIR, f"{pid}.csv")
            page = context.new_page()  # fresh page per place, so old reviews never leak in
            try:
                count = scrape_place(page, place, out_path)
                status = "done"
                print(f"    -> {count} reviews saved to {out_path}")
            except Exception as e:
                count = 0
                status = f"failed: {str(e)[:100]}"
                page.screenshot(path=os.path.join(OUTPUT_DIR, f"debug_{pid}.png"))
                print(f"    !! {status}")
            finally:
                page.close()

            summary[pid] = {**place, "reviews_collected": count, "status": status,
                            "finished_at": datetime.now().strftime("%Y-%m-%d %H:%M")}
            save_summary(summary)
            time.sleep(random.uniform(*PAUSE_BETWEEN_PLACES))

        browser.close()

    combine_outputs(places)


if __name__ == "__main__":
    main()