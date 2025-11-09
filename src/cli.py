import argparse
from pathlib import Path
from src.portal.scraper import login_y_buscar_ot
from src.io.excel_exporter import exportar
import pandas as pd
from playwright.sync_api import sync_playwright

def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--user", required=True)
    parser.add_argument("--pass", required=True)
    parser.add_argument("--ot", required=True)
    parser.add_argument("--out", default="resultado.xlsx")
    args = parser.parse_args()

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        context = browser.new_context()
        page, meta, hrefs = login_y_buscar_ot(context, args.user, args.pass_, args.ot)
        df = pd.DataFrame([meta])
        exportar(df, Path(args.out))
