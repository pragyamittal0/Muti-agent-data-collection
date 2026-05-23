# Multi-Agent Data Collection Pipeline

> A modular, resumable, fault-tolerant pipeline that collects public data at scale by dividing the work across specialized agents running in parallel.

![Pipeline Architecture](docs/pipeline-architecture.png)

---

## Overview

This project turns a fragile, one-shot scraping script into a production-grade system. Instead of one large program that breaks at the first failure, the work is split across **six specialized agents** that run independently, communicate through shared state, and recover gracefully from interruptions.

The result is a pipeline that is **fast, resumable, observable, and easy to extend**.

---

## Why Multi-Agent

A single scraping script faces three constant problems:

| Problem | Why It Hurts |
|---|---|
| **Speed** | One script handles one URL at a time. |
| **Fragility** | One failure kills the whole run. |
| **Debugging** | Hundreds of mixed responsibilities in one file. |

The multi-agent design solves all three by giving each agent a single, clear job. When one agent fails, the rest keep moving. When the script is killed, progress is saved and the next run continues from where it stopped.

---

## Architecture

The pipeline has six agents, each owning one stage of the workflow.

| # | Agent | File | Responsibility |
|---|---|---|---|
| 1 | **Input Agent** | `agents/input_agent.py` | Reads source files and prepares a clean list of targets. |
| 2 | **Discovery Agent** | `agents/discovery_agent.py` | Resolves business names or addresses into target URLs via search. |
| 3 | **Scraping Agents** (3 in parallel) | `agents/scraper.py` | Fetch data in coordinated waves. |
| 4 | **Recovery Agent** | `agents/recovery_agent.py` | Detects blocks, rotates credentials, retries failed pages. |
| 5 | **Enrichment Agent** | `agents/enrichment_agent.py` | Fills missing fields from secondary sources. |
| 6 | **Output Agent** | `agents/output_agent.py` | Cleans, deduplicates, and exports formatted Excel files. |

---

## How the Wave Structure Works

The scraping stage runs **three agents in parallel** per wave. After each wave finishes, progress is saved to disk.

```
Wave 1:  Agent A (6 items)   Agent B (6 items)   Agent C (6 items)   =  18 items
Wave 2:  Agent A (6 items)   Agent B (6 items)   Agent C (6 items)   =  18 more
Wave 3:  ...
```

If the run is interrupted at any point, the next execution reads the progress file and resumes from the next pending wave. No work is duplicated and nothing is lost.

---

## Tech Stack

| Tool | Purpose |
|---|---|
| **Python 3.11+** | Core language |
| **asyncio** | Coordinates parallel agents |
| **Apify cloud platform** | Hosts scraping agents with rotating proxies |
| **Playwright** | Headless browser for search and enrichment tasks |
| **pandas** | Data manipulation, merging, deduplication |
| **openpyxl** | Excel output with custom formatting |

---

## Project Structure

```
multi-agent-data-pipeline/
├── docs/
│   └── pipeline-architecture.png
├── agents/
│   ├── input_agent.py
│   ├── discovery_agent.py
│   ├── scraper.py
│   ├── recovery_agent.py
│   ├── enrichment_agent.py
│   └── output_agent.py
├── examples/
│   ├── panera_scraper.py
│   ├── panera_collect.py
│   └── dunkin_scraper.py
├── data/
│   ├── input/
│   ├── progress/
│   └── output/
├── requirements.txt
└── README.md
```

The `examples/` folder contains brand-specific implementations that show how the pipeline has been used on real projects.

---

## Quick Start

### 1. Clone and install

```bash
git clone https://github.com/<your-username>/multi-agent-data-pipeline.git
cd multi-agent-data-pipeline
pip install -r requirements.txt
playwright install chromium
```

### 2. Set credentials

```bash
export APIFY_TOKEN=your_apify_token_here
```

### 3. Run the pipeline

```bash
python -m agents.input_agent
python -m agents.discovery_agent
python -m agents.scraper
python -m agents.enrichment_agent
python -m agents.output_agent
```

### 4. Resume after interruption

```bash
python -m agents.scraper --resume
```

The scraper reads `data/progress/progress.json` and picks up from the last completed wave.

---

## Handling Blocks and Failures

Bot protection systems often block scrapers after a few hundred requests. The Recovery Agent watches for this pattern (multiple empty waves in a row) and triggers credential rotation. With a fresh set of credentials, the pipeline continues until the next block.

This rotation strategy is what allows high coverage even when individual sessions get blocked.

---

## Output Format

The Output Agent produces a clean Excel workbook with multiple sheets:

| Sheet | Contents |
|---|---|
| **Summary** | Totals, date range, coverage percentage |
| **Per-Business** | One row per business with name, address, counts, average rating, URL |
| **All Records** | Every individual record with all fields |
| **Failures** | URLs that could not be processed, with reasons |

The format is designed for analysts to open in Excel and start working immediately. No technical setup required.

---

## Key Design Principles

1. **Modular beats monolithic.** Six small agents are easier to build, test, and fix than one giant script.
2. **Save progress constantly.** Anything running longer than five minutes saves state after every meaningful step.
3. **Plan for blocks, not just success.** The recovery strategy is what makes a pipeline production ready.
4. **Use the right tool for each agent.** A scraping agent needs a real browser. A merging agent only needs pandas.

---

## Future Scope

- **Self-Healing Agents** — Automatic credential rotation from a managed pool
- **Smart Scheduling** — Time-of-day awareness to avoid peak bot-protection windows
- **Real-Time Dashboard** — Web UI showing live agent status and wave progress
- **Cross-Source Enrichment** — Pull from multiple backup sources in parallel
- **Quality Scoring** — Per-record confidence and completeness scores
- **Pipeline as a Service** — Configurable tool where users specify only source and target fields
- **AI-Assisted Field Extraction** — Language model agent for sentiment, topics, and structured extraction
- **Automatic Schema Detection** — Input Agent learns column meanings from new file formats

Each of these builds on the existing modular foundation. New capabilities can be added by introducing new agents, not by rewriting the system.

---

## License

MIT License. See [LICENSE](LICENSE) for details.
