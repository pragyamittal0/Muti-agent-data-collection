"""
Retry the 150 failed URLs with a fresh Apify token.
Same wave structure: 3 actors × 6 URLs per wave.
Appends new reviews to reviews_merged.csv.
"""

import os, sys, math, asyncio, json
from datetime import datetime
from pathlib import Path

import pandas as pd
from apify_client import ApifyClientAsync

APIFY_TOKEN     = os.environ.get("APIFY_TOKEN", "")
ACTOR_ID        = "tri_angle/yelp-review-scraper"

DATE_FROM       = "2019-12-15"
DATE_TO         = "2021-12-15"
MAX_REVIEWS     = 200
URLS_PER_ACTOR  = 6
ACTORS_PER_WAVE = 3
WAVE_PAUSE_S    = 8

BASE        = Path(__file__).parent
INPUT_X     = BASE / "../Dunkin_CA_Yelp_URLs.xlsx"
CSV_OUT     = BASE / "reviews_merged.csv"
PROGRESS_F  = BASE / "progress.json"
RETRY_PROG  = BASE / "retry_progress.json"

def get_failed_urls():
    """Diff: input URLs vs URLs that have reviews already."""
    all_urls = pd.read_excel(INPUT_X, sheet_name="All Yelp URLs")["Yelp URL"].dropna().tolist()
    df = pd.read_csv(CSV_OUT)
    have = set(df["businessUrl"].dropna().unique())
    return [u for u in all_urls if u not in have]

def load_retry_progress():
    if RETRY_PROG.exists():
        return json.loads(RETRY_PROG.read_text())
    return {"done_urls": []}

def save_retry_progress(prog):
    RETRY_PROG.write_text(json.dumps(prog, indent=2))

async def run_actor(client, wave_id, actor_id, urls):
    label = f"Retry Wave {wave_id} / Actor {actor_id}"
    print(f"    [{label}] starting — {len(urls)} URLs", flush=True)
    run_input = {
        "startUrls":        [{"url": u} for u in urls],
        "maxReviewsPerUrl": MAX_REVIEWS,
        "dateFrom":         DATE_FROM,
        "dateTo":           DATE_TO,
        "language":         "en",
    }
    try:
        run = await client.actor(ACTOR_ID).call(run_input=run_input)
        dataset_id = run["defaultDatasetId"]
        items = []
        async for item in client.dataset(dataset_id).iterate_items():
            item["_wave"]      = f"R{wave_id}"
            item["_actor"]     = actor_id
            item["_scraped_at"] = datetime.utcnow().isoformat()
            items.append(item)
        print(f"    [{label}] ✓ {len(items)} reviews  (dataset {dataset_id})", flush=True)
        return items, urls
    except Exception as e:
        print(f"    [{label}] ✗ ERROR: {e}", flush=True)
        return [], []

async def main():
    if not APIFY_TOKEN:
        sys.exit("ERROR: set APIFY_TOKEN")

    print("=" * 60, flush=True)
    print(" Retry scraper — 3 actors × 6 URLs per wave", flush=True)
    print(f" Date range: {DATE_FROM} → {DATE_TO}", flush=True)
    print("=" * 60, flush=True)

    failed = get_failed_urls()
    prog   = load_retry_progress()
    done   = set(prog["done_urls"])
    pending = [u for u in failed if u not in done]
    print(f"  Failed input URLs : {len(failed)}", flush=True)
    print(f"  Already retried   : {len(done)}", flush=True)
    print(f"  Pending           : {len(pending)}\n", flush=True)

    if not pending:
        print("  Nothing to do.", flush=True); return

    chunks = [pending[i:i+URLS_PER_ACTOR] for i in range(0, len(pending), URLS_PER_ACTOR)]
    total_waves = math.ceil(len(chunks) / ACTORS_PER_WAVE)

    client = ApifyClientAsync(APIFY_TOKEN)
    new_reviews_total = 0

    for wave_idx in range(0, len(chunks), ACTORS_PER_WAVE):
        wave_num = wave_idx // ACTORS_PER_WAVE + 1
        wave_chunks = chunks[wave_idx:wave_idx + ACTORS_PER_WAVE]
        print(f"  ── Wave R{wave_num}/{total_waves} ── {len(wave_chunks)} actors ──", flush=True)

        tasks = [run_actor(client, wave_num, i+1, c) for i, c in enumerate(wave_chunks)]
        results = await asyncio.gather(*tasks)

        wave_reviews = []
        for items, urls in results:
            wave_reviews.extend(items)
            prog["done_urls"].extend(urls)

        # Append to CSV
        if wave_reviews:
            wdf = pd.DataFrame(wave_reviews)
            existing_cols = pd.read_csv(CSV_OUT, nrows=0).columns.tolist()
            for c in existing_cols:
                if c not in wdf.columns:
                    wdf[c] = pd.NA
            wdf = wdf[existing_cols]
            wdf.to_csv(CSV_OUT, mode="a", index=False, header=False)
            new_reviews_total += len(wave_reviews)

        save_retry_progress(prog)
        print(f"  Wave R{wave_num} done — {len(wave_reviews)} reviews | total new: {new_reviews_total}\n", flush=True)

        if wave_idx + ACTORS_PER_WAVE < len(chunks):
            print(f"  Pausing {WAVE_PAUSE_S}s...", flush=True)
            await asyncio.sleep(WAVE_PAUSE_S)

    print(f"\n  Done! Added {new_reviews_total} new reviews to CSV.", flush=True)

if __name__ == "__main__":
    asyncio.run(main())
