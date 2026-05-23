"""
Del Taco CA — Yelp Review Scraper
3 actors × 6 URLs per wave, date range 2021-03-08 → 2023-03-08
"""
import os, sys, math, asyncio, argparse, json
from datetime import datetime
from pathlib import Path
import pandas as pd
from apify_client import ApifyClientAsync

APIFY_TOKEN     = os.environ.get("APIFY_TOKEN", "")
ACTOR_ID        = "tri_angle/yelp-review-scraper"

DATE_FROM       = "2021-03-08"
DATE_TO         = "2023-03-08"
MAX_REVIEWS     = 300
URLS_PER_ACTOR  = 6
ACTORS_PER_WAVE = 3
WAVE_PAUSE_S    = 8

BASE        = Path(__file__).parent
URLS_F      = BASE / "yelp_urls.json"
PROGRESS_F  = BASE / "progress.json"
CSV_OUT     = BASE / "reviews_merged.csv"

def load_urls():
    data = json.loads(URLS_F.read_text())
    urls = sorted({r["yelp_url"] for r in data.values() if r.get("yelp_url")})
    print(f"  Loaded {len(urls)} unique Yelp URLs", flush=True)
    return urls

def load_progress():
    if PROGRESS_F.exists():
        return json.loads(PROGRESS_F.read_text())
    return {"done_urls": [], "dataset_ids": {}}

def save_progress(p):
    PROGRESS_F.write_text(json.dumps(p, indent=2))

async def run_actor(client, wave_id, actor_id, urls):
    label = f"Wave {wave_id} / Actor {actor_id}"
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
            item["_wave"] = wave_id
            item["_actor"] = actor_id
            item["_scraped_at"] = datetime.utcnow().isoformat()
            items.append(item)
        print(f"    [{label}] ✓ {len(items)} reviews  (dataset {dataset_id})", flush=True)
        return items, dataset_id
    except Exception as e:
        print(f"    [{label}] ✗ ERROR: {e}", flush=True)
        return [], None

async def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--resume", action="store_true")
    args = parser.parse_args()

    if not APIFY_TOKEN:
        sys.exit("ERROR: set APIFY_TOKEN")

    print("=" * 60, flush=True)
    print(" Del Taco CA Yelp Scraper — 3 actors × 6 URLs/wave", flush=True)
    print(f" Date range: {DATE_FROM} → {DATE_TO}", flush=True)
    print("=" * 60, flush=True)

    all_urls = load_urls()
    prog = load_progress() if args.resume else {"done_urls": [], "dataset_ids": {}}
    done = set(prog["done_urls"])
    pending = [u for u in all_urls if u not in done]

    total_urls   = len(pending)
    total_actors = math.ceil(total_urls / URLS_PER_ACTOR)
    total_waves  = math.ceil(total_actors / ACTORS_PER_WAVE)
    print(f"\n  Pending: {total_urls} | Actors: {total_actors} | Waves: {total_waves}\n", flush=True)

    if not pending:
        print("  Nothing to scrape", flush=True); return

    chunks = [pending[i:i+URLS_PER_ACTOR] for i in range(0, total_urls, URLS_PER_ACTOR)]
    client = ApifyClientAsync(APIFY_TOKEN)

    for wave_idx in range(0, len(chunks), ACTORS_PER_WAVE):
        wave_num = wave_idx // ACTORS_PER_WAVE + 1
        wave_chunks = chunks[wave_idx:wave_idx + ACTORS_PER_WAVE]
        print(f"  ── Wave {wave_num}/{total_waves} ── {len(wave_chunks)} actors ──", flush=True)

        tasks = [run_actor(client, wave_num, i+1, c) for i, c in enumerate(wave_chunks)]
        results = await asyncio.gather(*tasks)

        wave_reviews = []
        for (items, dsid), chunk in zip(results, wave_chunks):
            wave_reviews.extend(items)
            if dsid:
                prog["done_urls"].extend(chunk)
                prog["dataset_ids"][dsid] = chunk

        if wave_reviews:
            wdf = pd.DataFrame(wave_reviews)
            if CSV_OUT.exists():
                # Align columns
                existing = pd.read_csv(CSV_OUT, nrows=0).columns.tolist()
                for c in existing:
                    if c not in wdf.columns:
                        wdf[c] = pd.NA
                wdf = wdf[existing]
                wdf.to_csv(CSV_OUT, mode="a", index=False, header=False)
            else:
                wdf.to_csv(CSV_OUT, index=False)

        save_progress(prog)
        print(f"  Wave {wave_num} done — {len(wave_reviews)} reviews | total scraped: {len(prog['done_urls'])}/{len(all_urls)}\n", flush=True)

        if wave_idx + ACTORS_PER_WAVE < len(chunks):
            print(f"  Pausing {WAVE_PAUSE_S}s...", flush=True)
            await asyncio.sleep(WAVE_PAUSE_S)

    # Total stats
    total = 0
    if CSV_OUT.exists():
        total = len(pd.read_csv(CSV_OUT))
    print(f"\n  Done! Total reviews in CSV: {total}", flush=True)

if __name__ == "__main__":
    asyncio.run(main())
