# Dot-connect — Outlook Email Network Visualizer

Visualize Outlook email send/receive/CC relationships as an interactive network graph. Automatically detect **CC key persons**, **hub individuals**, and **passive observers** to reveal hidden communication structures within your organization.

> **[Live Demo (sample data)](https://9BwgeBTPG-QH.github.io/Dot-connect/index_en.html)** — Try the interactive visualization with fictional team data.

> **[日本語版 README はこちら](README.ja.md)**

## Features

- **Network Graph** — Obsidian-style dark theme, force-directed layout (vis.js barnesHut)
- **Node Drill-down** — Click a node to zoom into its connections (Workflowy-style)
- **Community Detection** — Louvain algorithm with convex hull boundaries
- **CC Key Person** — Identifies people who appear in CC above a configurable threshold
- **Hub Detection** — Weighted degree + betweenness centrality scoring
- **Passive Observer** — Detects receive-only and CC-only participants
- **Word Cloud** — Name-based word cloud sized by email frequency
- **Export** — PNG screenshot, standalone HTML, CSV analysis results
- **Zero-install for end users** — Embedded Python via network share (no local install needed)

## Quick Start

### Option A: Local PC (simplest)

1. Double-click `setup.bat` (first time only — installs Python + dependencies)
2. Double-click `start.bat` → browser opens automatically
3. Select Outlook folders, set date range, click "Extract & Analyze"

### Option B: File Server (multi-user)

Run the server on a shared folder — users access via browser, no Python install required.

**Server setup (once):**
1. Place Dot-connect on a network share (e.g., `\\SERVER\share\Dot-connect`)
2. Run `setup.bat` → `start.bat`
3. Set `network_share_path` in `config.yaml`
4. Allow port 8000 in Windows Firewall

**Each user:**
1. Open `http://<server>:8000` in browser
2. Download the `.bat` extractor, run it locally
3. Select Outlook folders → auto-extract & upload → results appear in browser

### Option C: CSV Upload

Upload a previously extracted CSV file via the web UI.

### Option D: Docker

```bash
docker compose up
# Open http://localhost:8000 and upload a CSV
```

## Developer Setup

```bash
pip install -r requirements.txt
pip install pywin32               # Windows only (Outlook extraction)
uvicorn app.main:app --reload     # Dev server
```

### CLI Usage

```bash
# Extract emails from Outlook
python extract.py --start 2025-01-01 --end 2025-12-31

# Generate HTML visualization
python generate.py --input output/emails_20250101.csv
```

### Sample Data

A sample CSV with dummy data is included for testing:

```bash
# Via CLI
python generate.py --input sample/emails_sample.csv

# Or upload via web UI at http://localhost:8000
```

## Pipeline

```
Option A (Local):     start.bat → Browser selects Outlook folders → Visualization
Option B (Server):    Browser → .bat download → Local extract → Server analysis → Visualization
Option C (CSV):       start.bat → Browser CSV upload → Visualization
Option D (CLI):       extract.py → CSV → generate.py → index.html
```

## Analysis

| Analysis | Description |
|----------|-------------|
| **CC Key Person** | People with CC appearance rate above threshold (default: 30%) |
| **Hub** | Top nodes by weighted degree + betweenness centrality |
| **Community** | Automatic cluster detection via Louvain algorithm |
| **Passive Observer** | Receive-only or CC-only participants (never send) |

## Configuration

See [`config.yaml`](config.yaml) for all settings:

```yaml
network_share_path: "\\\\SERVER\\share\\Dot-connect"

company_domains:
  - example.co.jp

exclude_patterns:
  - "^no-?reply@"

alias_map:
  canonical@example.co.jp:
    - alias@old-domain.co.jp

thresholds:
  cc_key_person_threshold: 0.30
  min_edge_weight: 1
  hub_degree_weight: 0.5
  hub_betweenness_weight: 0.5
```

## Project Structure

```
Dot-connect/
├── app/
│   ├── core.py              # Analysis engine (shared by CLI & Web)
│   ├── extract.py           # Outlook COM wrapper
│   ├── main.py              # FastAPI application
│   └── models.py            # Pydantic validation models
├── templates/
│   ├── upload.html          # Web UI: upload & extract page
│   └── network.html         # Visualization (vis.js + wordcloud2.js)
├── sample/
│   └── emails_sample.csv    # Sample data for testing
├── extract.py               # CLI: Outlook → CSV extraction
├── generate.py              # CLI: CSV → HTML generation
├── extract_and_upload.py    # Self-contained local extractor + uploader
├── config.yaml              # Configuration
├── setup.bat                # One-click setup (downloads Python)
├── start.bat                # One-click server start
├── Dockerfile               # Docker support
├── docker-compose.yml       # Docker Compose
├── requirements.txt         # Python dependencies
└── requirements-extract.txt # pywin32 for Outlook COM
```

## Privacy & Data Handling

### Data Collected

This tool extracts the following metadata from Outlook via COM automation:

- Sender email address and display name
- To/CC recipient email addresses and display names
- Received date and subject line

**Email body content is never collected or stored.**

### Data Processing & Storage

- All processing happens **entirely on your local machine** (or your own server). No data is sent to any external service
- When using the file server mode (Option B), extracted CSV is uploaded only to your own internal server
- Analysis results are held in memory and discarded when the server stops
- Exported HTML files contain aggregated network data (names, email addresses, communication counts) — **be mindful of who you share them with**

### Recommendations for Admins

- Before deploying this tool, **inform employees** that their email communication patterns will be visualized
- The tool reveals organizational communication structures — treat the output as confidential
- Ensure usage complies with your organization's data policies and applicable privacy regulations (e.g., GDPR, APPI)

## Related Projects

- [slack-mention-map](https://github.com/9BwgeBTPG-QH/slack-mention-map) — Slack mention network visualization tool

## License

[MIT](LICENSE)
