# Equity Research AI Agent

A free, classroom-oriented Streamlit app for an auditable equity-research workflow.

## What it teaches

- SEC financial statement retrieval and XBRL mapping
- Five-year historical financial analysis
- Ratio analysis
- Revenue-driver and risk evidence from the latest 10-K
- WACC and five-year UFCF DCF valuation
- Perpetuity-growth and exit EV/EBITDA terminal methods
- Sensitivity analysis and valuation football field
- Excel model generation with formulas
- Constrained AI memo prompting
- Hallucination / claim auditing
- Source provenance and model governance

## Files

- `streamlit_app.py` — the entire web application.
- `requirements.txt` — Python packages Streamlit Community Cloud installs.
- `README.md` — this overview.

## Classroom default

Ticker: `STLD` (Steel Dynamics)

Change the ticker in the sidebar to rerun the workflow for another U.S. public company.

## Important control

The app intentionally gates later modules until the student manually reconciles:
1. Revenue
2. Operating Income
3. Operating Cash Flow
4. Capital Expenditures

against the latest 10-K.

## Free data sources

- SEC EDGAR submissions and Companyfacts/XBRL
- U.S. Treasury Daily Treasury Par Yield Curve Rates
- Damodaran implied U.S. equity risk premium
- yfinance / Yahoo Finance market-price reference and beta cross-check

The accounting, ratio, WACC, DCF, sensitivity, and Excel modules do **not** require a generative-AI API.

## Deployment

Deploy from a public GitHub repository to Streamlit Community Cloud with `streamlit_app.py` as the main file.

Educational use only. Not investment advice.
