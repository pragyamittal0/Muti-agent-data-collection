"""
Dry-run preview — shows how the 183 URLs are distributed across 15 agents
without calling Apify. Run this first to verify the batching before spending credits.

Usage:  python preview.py
"""

import math
import pandas as pd

EXCEL_INPUT     = "../Panera_CA_All_Yelp_URLs_Complete.xlsx"
DATE_FROM       = "2016-07-18"
DATE_TO         = "2018-07-18"
MAX_REVIEWS     = 100
N_OPEN_AGENTS   = 11
N_CLOSED_AGENTS = 4
ACTOR           = "tri_angle/yelp-review-scraper"


def load_urls():
    df = pd.read_excel(EXCEL_INPUT, header=3)
    df.columns = ["drop", "num", "region", "yelp_url", "status"]
    df = df.dropna(subset=["yelp_url"])
    df = df[df["yelp_url"].astype(str).str.startswith("http")].copy()
    df["status"] = df["status"].str.strip().str.title()
    df["region"] = df["region"].ffill()
    return df[df["status"] == "Open"], df[df["status"] == "Closed"]


def split(urls, n):
    size = math.ceil(len(urls) / n)
    return [urls[i:i+size] for i in range(0, len(urls), size)]


def main():
    open_df, closed_df = load_urls()
    print(f"URLs loaded: {len(open_df)} open + {len(closed_df)} closed = {len(open_df)+len(closed_df)} total\n")

    open_batches   = split(open_df["yelp_url"].tolist(),   N_OPEN_AGENTS)
    closed_batches = split(closed_df["yelp_url"].tolist(), N_CLOSED_AGENTS)

    print(f"{'Agent':<8} {'Type':<8} {'URLs':>5}  Sample URL")
    print("-" * 75)
    agent = 1
    for batch in open_batches:
        print(f"A{agent:<7} {'open':<8} {len(batch):>5}  {batch[0]}")
        agent += 1
    for batch in closed_batches:
        print(f"A{agent:<7} {'closed':<8} {len(batch):>5}  {batch[0]}")
        agent += 1

    total_urls = sum(len(b) for b in open_batches + closed_batches)
    est_reviews = total_urls * MAX_REVIEWS
    print("-" * 75)
    print(f"\nTotal agents : {N_OPEN_AGENTS + N_CLOSED_AGENTS}  ({N_OPEN_AGENTS} open + {N_CLOSED_AGENTS} closed)")
    print(f"Total URLs   : {total_urls}")
    print(f"Date range   : {DATE_FROM} → {DATE_TO}")
    print(f"Max reviews  : {MAX_REVIEWS} per outlet")
    print(f"Est. reviews : up to {est_reviews:,}")
    print(f"Actor        : {ACTOR}")
    print(f"\nRun  python scraper.py  (with APIFY_TOKEN set) to execute.")


if __name__ == "__main__":
    main()
