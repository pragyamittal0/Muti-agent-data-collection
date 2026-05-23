"""
Yahoo-search each Del Taco location to find its Yelp /biz/ URL.
"""
import asyncio, json, re
from pathlib import Path
from playwright.async_api import async_playwright

BASE = Path(__file__).parent
LOCATIONS_F = BASE / "locations.json"
RESULTS_F   = BASE / "yelp_urls.json"

async def search(browser, query):
    page = await browser.new_page()
    try:
        url = f"https://search.yahoo.com/search?p={query.replace(' ', '+').replace(chr(34), '%22')}"
        await page.goto(url, wait_until="commit", timeout=15000)
        await asyncio.sleep(1.2)
        content = await page.content()
        matches = re.findall(r'yelp\.com/biz/([a-zA-Z0-9_-]+)', content)
        # Filter to del-taco slugs
        deltaco = [m for m in set(matches) if 'del-taco' in m.lower()]
        return deltaco
    except Exception:
        return []
    finally:
        await page.close()

async def main():
    locations = json.loads(LOCATIONS_F.read_text())
    results = {}
    if RESULTS_F.exists():
        results = json.loads(RESULTS_F.read_text())

    pending = [L for L in locations if str(L["num"]) not in results]
    print(f"Total: {len(locations)}, Already done: {len(results)}, Pending: {len(pending)}\n", flush=True)

    async with async_playwright() as p:
        browser = await p.chromium.launch(headless=True)

        for i, loc in enumerate(pending):
            addr = loc["address"]
            city = loc["city"]
            q1 = f'del taco "{addr}" {city} california yelp'
            slugs = await search(browser, q1)
            if not slugs:
                q2 = f'del taco {addr} {city} yelp'
                slugs = await search(browser, q2)

            results[str(loc["num"])] = {
                "city": city, "address": addr,
                "slugs": slugs,
                "yelp_url": f"https://www.yelp.com/biz/{slugs[0]}" if slugs else ""
            }
            print(f"[{i+1}/{len(pending)}] #{loc['num']} {city}/{addr[:30]}: {slugs[:2] if slugs else 'NOT FOUND'}", flush=True)

            if (i+1) % 20 == 0:
                RESULTS_F.write_text(json.dumps(results, indent=2))

            await asyncio.sleep(0.5)

        await browser.close()
    RESULTS_F.write_text(json.dumps(results, indent=2))

    found = sum(1 for r in results.values() if r["slugs"])
    unique = len({r["yelp_url"] for r in results.values() if r["yelp_url"]})
    print(f"\nFound: {found}/{len(results)} | Unique URLs: {unique}", flush=True)

if __name__ == "__main__":
    asyncio.run(main())
