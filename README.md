# Freight Cost Analysis Platform

Automated month-end freight cost intelligence dashboard built with Streamlit.

## Quick Start

```bash
pip install -r requirements.txt
streamlit run workers/code2_dashboard.py
```

## What It Does

- **Net cost formula:** LE Posting + WE Accrual − WE Reverse Accrual
- **MoM & YoY comparison** with KPIs (Total Cost, Shipments, Cost/Shipment, Delta)
- **Cost decomposition:** Volume vs Rate vs Mix
- **Driver analysis:** By Carrier, By Country, By Order Type, By Cost Type
- **Rate card comparison:** 2025 vs 2026 contract rates by carrier/country
- **Auto-generated insights** in plain English

## Data Included (August demo)

- `data/Accruals/` — 3 accrual files (Aug 2026, Jul 2026, Aug 2025)
- `data/Postings/` — 3 posting files (Aug 2026, Jul 2026, Aug 2025)
- `data/rate_cards/` — 2 rate card files (2025 & 2026 contracts)

## Architecture

```
freight-cost-analysis-demo/
├── workers/
│   ├── config_loader.py      ← Loads YAML configs
│   ├── code1_ingest.py        ← BMS download → posting file
│   ├── code2_dashboard.py     ← Streamlit dashboard (RUN THIS)
│   └── code3_rate_cards.py    ← Rate card comparison
├── configs/                  ← All mappings (edit YAML, not Python)
│   ├── warehouse_mapping.yaml
│   ├── carrier_mapping.yaml
│   ├── charge_type_mapping.yaml
│   ├── column_structure.yaml
│   └── formula.yaml
├── data/                     ← Accruals, Postings, Rate cards
├── .streamlit/config.toml
└── requirements.txt
```

## Deploy to Streamlit Community Cloud

1. Push this folder to a public GitHub repo
2. Go to [share.streamlit.io](https://share.streamlit.io)
3. Connect your GitHub repo
4. Set main file path: `workers/code2_dashboard.py`
5. Deploy — you get a public URL

## Team

**Matias Alonso** | SELS Vibe Coding AI Challenge 2026
