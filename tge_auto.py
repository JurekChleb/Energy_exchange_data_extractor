import argparse
from datetime import datetime
from pathlib import Path
import re

import pandas as pd
from playwright.sync_api import sync_playwright

# ---------- helpers ----------
def parse_pl_number(s: str) -> float:
  """
  Convert Polish-formatted numbers to float:
  "2 276,76" -> 2276.76
  "2 231" (with NBSP) -> 2231.0
  "-" -> None
  """
  if s is None:
    return None
  s = s.strip()
  if s == "-" or s == "":
    return None
  # replace non-breaking spaces and regular spaces used as thousand separators
  s = s.replace("\u00a0", " ").replace(" ", "")
  # comma decimal to dot
  s = s.replace(",", ".")
  # leave digits and dot only
  s = re.sub(r"[^0-9.]", "", s)
  return float(s) if s else None

def write_to_excel(xlsx_path: Path, row_dict: dict, sheet_name: str = "data"):
  """
  Append row_dict to xlsx_path / sheet_name, create if not exists.
  Deduplicate by (date,label) keeping the latest.
  """
  df_new = pd.DataFrame([row_dict])
  if xlsx_path.exists():
    df_old = pd.read_excel(xlsx_path, sheet_name=sheet_name)
    # unify dtypes
    df_old["date"] = pd.to_datetime(df_old["date"]).dt.date
    df_new["date"] = pd.to_datetime(df_new["date"]).dt.date
    df_all = pd.concat([df_old, df_new], ignore_index=True)
    df_all = df_all.sort_values("date").drop_duplicates(subset=["date", "label"], keep="last")
  else:
    df_all = df_new
  xlsx_path.parent.mkdir(parents=True, exist_ok=True)
  with pd.ExcelWriter(xlsx_path, engine="openpyxl", mode="w") as writer:
    df_all.to_excel(writer, sheet_name=sheet_name, index=False)

# ---------- main ----------
def main(date_str: str):
  # Accept both "30-10-2025" and "2025-10-30"
  try:
    dt = datetime.strptime(date_str, "%d-%m-%Y")
  except ValueError:
    dt = datetime.strptime(date_str, "%Y-%m-%d")

  date_param = dt.strftime("%d-%m-%Y")
  url = f"https://tge.pl/prawa-majatkowe-rpm?dateShow={date_param}&dateAction="

  outdir = Path("out") / dt.strftime("%Y/%m/%d")
  outdir.mkdir(parents=True, exist_ok=True)
  img_path = outdir / f"TGEeff_{dt.strftime('%Y-%m-%d')}.png"
  csv_path = outdir / f"TGEeff_{dt.strftime('%Y-%m-%d')}.csv"
  xlsx_path = Path("out") / "tgeeff_history.xlsx"

  with sync_playwright() as p:
    browser = p.chromium.launch(headless=True)
    context = browser.new_context(
      viewport={"width": 1280, "height": 2000},
      device_scale_factor=2,
    )
    page = context.new_page()
    page.goto(url, wait_until="networkidle")
    page.wait_for_timeout(800)

    # Cookie banner (if any)
    if page.locator("button:has-text('Akceptuj')").count():
      page.click("button:has-text('Akceptuj')")

    # Row with TGEeff
    row = page.locator("tr:has(td:has-text('TGEeff'))").first
    row.wait_for(state="visible", timeout=10_000)

    # Screenshot of the row
    row.screenshot(path=str(img_path))

    # Extract cells: [label, kurs, zmiana, wolumen, zmiana]
    tds = row.locator("td")
    label = tds.nth(0).inner_text().strip()
    kurs_txt = tds.nth(1).inner_text().strip()
    wolumen_txt = tds.nth(3).inner_text().strip()

    kurs = parse_pl_number(kurs_txt)
    wolumen = parse_pl_number(wolumen_txt)

    # Save a CSV alongside (optional)
    csv_path.write_text(
      "date,label,kurs_pln_per_toe,wolumen_toe\n"
      f"{dt.date()},{label},{kurs_txt},{wolumen_txt}\n",
      encoding="utf-8"
    )

    # Append to Excel history
    write_to_excel(
      xlsx_path,
      {
        "date": dt.date(),
        "label": label,
        "kurs_pln_per_toe": kurs,
        "wolumen_toe": wolumen,
        "kurs_raw": kurs_txt,
        "wolumen_raw": wolumen_txt,
      },
      sheet_name="data",
    )

    print(f"Saved screenshot → {img_path}")
    print(f"Saved CSV    → {csv_path}")
    print(f"Updated Excel  → {xlsx_path}")

    context.close()
    browser.close()

if __name__ == "__main__":
  parser = argparse.ArgumentParser()
  parser.add_argument("--date", required=True, help="Date like 30-10-2025 or 2025-10-30")
  args = parser.parse_args()
  main(args.date)