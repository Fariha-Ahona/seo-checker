"""
SEO Checker
-----------
Reads a list of website URLs from a CSV file, audits each page for basic
SEO health, and writes the results to an Excel report.

Usage:
    python seo_checker.py urls.csv
    python seo_checker.py urls.csv -o my_report.xlsx --max-links 10
"""

import argparse
import sys
import time
from urllib.parse import urljoin, urlparse

import pandas as pd
import requests
from bs4 import BeautifulSoup

TIMEOUT = 10  # seconds to wait for a page before giving up
HEADERS = {"User-Agent": "Mozilla/5.0 (SEO-Checker; +https://github.com/)"}
TITLE_MAX = 60
META_DESC_MAX = 160


def read_urls(csv_path):
    """Read URLs from the first column of a CSV file."""
    df = pd.read_csv(csv_path)
    urls = df.iloc[:, 0].dropna().astype(str).str.strip().tolist()
    # Add https:// if the user forgot it
    return [u if u.startswith(("http://", "https://")) else "https://" + u for u in urls]


def fetch_page(url):
    """Download a page. Returns (response, load_time_seconds)."""
    start = time.time()
    response = requests.get(url, headers=HEADERS, timeout=TIMEOUT)
    load_time = round(time.time() - start, 2)
    return response, load_time


def extract_seo_data(html):
    """Pull the title, meta description and first H1 from the HTML."""
    soup = BeautifulSoup(html, "html.parser")

    title = soup.title.get_text(strip=True) if soup.title else ""

    meta = soup.find("meta", attrs={"name": "description"})
    meta_desc = meta.get("content", "").strip() if meta else ""

    h1 = soup.find("h1")
    h1_text = h1.get_text(strip=True) if h1 else ""

    return soup, title, meta_desc, h1_text


def check_broken_links(soup, base_url, max_links):
    """Check the first `max_links` links on the page. Returns a list of broken ones."""
    links = []
    for a in soup.find_all("a", href=True):
        full = urljoin(base_url, a["href"])
        if urlparse(full).scheme in ("http", "https") and full not in links:
            links.append(full)
        if len(links) >= max_links:
            break

    broken = []
    for link in links:
        try:
            # HEAD is faster; some servers reject it, so fall back to GET
            r = requests.head(link, headers=HEADERS, timeout=TIMEOUT, allow_redirects=True)
            if r.status_code in (405, 403):
                r = requests.get(link, headers=HEADERS, timeout=TIMEOUT)
            if r.status_code >= 400:
                broken.append(f"{link} ({r.status_code})")
        except requests.RequestException:
            broken.append(f"{link} (no response)")
    return broken


def find_issues(status, title, meta_desc, h1, load_time, broken):
    """Turn the raw data into a human-readable list of problems."""
    issues = []
    if status >= 400:
        issues.append(f"Page returned {status}")
    if not title:
        issues.append("Missing title")
    elif len(title) > TITLE_MAX:
        issues.append(f"Title too long ({len(title)} chars)")
    if not meta_desc:
        issues.append("Missing meta description")
    elif len(meta_desc) > META_DESC_MAX:
        issues.append(f"Meta description too long ({len(meta_desc)} chars)")
    if not h1:
        issues.append("Missing H1")
    if load_time > 3:
        issues.append(f"Slow load ({load_time}s)")
    if broken:
        issues.append(f"{len(broken)} broken link(s)")
    return "; ".join(issues) if issues else "OK"


def audit_url(url, max_links):
    """Run every check on one URL and return a result row."""
    print(f"Checking {url} ...")
    try:
        response, load_time = fetch_page(url)
        soup, title, meta_desc, h1 = extract_seo_data(response.text)
        broken = check_broken_links(soup, url, max_links)
        return {
            "URL": url,
            "Status": response.status_code,
            "Load Time (s)": load_time,
            "Title": title,
            "Title Length": len(title),
            "Meta Description": meta_desc,
            "H1": h1,
            "Broken Links": "\n".join(broken),
            "Issues": find_issues(response.status_code, title, meta_desc, h1, load_time, broken),
        }
    except requests.RequestException as e:
        # Timeout, DNS failure, bad URL... record it and keep going
        return {
            "URL": url,
            "Status": "Failed",
            "Load Time (s)": None,
            "Title": "",
            "Title Length": 0,
            "Meta Description": "",
            "H1": "",
            "Broken Links": "",
            "Issues": f"Could not reach page: {type(e).__name__}",
        }


def save_report(rows, output_path):
    """Write results to Excel with readable column widths."""
    df = pd.DataFrame(rows)
    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="SEO Report")
        sheet = writer.sheets["SEO Report"]
        widths = {"A": 35, "B": 9, "C": 13, "D": 40, "E": 12, "F": 50, "G": 30, "H": 50, "I": 50}
        for col, width in widths.items():
            sheet.column_dimensions[col].width = width
        sheet.freeze_panes = "A2"


def main():
    parser = argparse.ArgumentParser(description="Audit websites for basic SEO issues.")
    parser.add_argument("csv_file", help="CSV file with URLs in the first column")
    parser.add_argument("-o", "--output", default="seo_report.xlsx", help="Output Excel file")
    parser.add_argument("--max-links", type=int, default=15, help="Links to check per page")
    args = parser.parse_args()

    try:
        urls = read_urls(args.csv_file)
    except FileNotFoundError:
        sys.exit(f"File not found: {args.csv_file}")

    if not urls:
        sys.exit("No URLs found in the CSV file.")

    rows = [audit_url(url, args.max_links) for url in urls]
    save_report(rows, args.output)

    ok = sum(1 for r in rows if r["Issues"] == "OK")
    print(f"\nDone. {len(rows)} pages checked, {ok} with no issues.")
    print(f"Report saved to {args.output}")


if __name__ == "__main__":
    main()
