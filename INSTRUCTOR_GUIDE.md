# Instructor Guide — Equity Research AI Agent

## 1. Learning architecture

The application enforces this sequence:

**Company identity → SEC filings → XBRL mapping → manual reconciliation → ratios → filing evidence → risks → forecast assumptions → WACC → DCF → sensitivity → Excel → AI memo → human audit**

The critical governance rule is:

**Application = evidence + calculations**  
**AI = synthesis**  
**Student = verification + judgment**

Later modules are gated until Revenue, Operating Income, Operating Cash Flow, and Capital Expenditures are manually reconciled to the latest 10-K.

## 2. Module-to-module information flow

### Module 1 — Company
Input: editable ticker, default STLD.

Outputs:
- SEC company name
- CIK
- latest 10-K metadata and filing URL
- market-price reference
- basic company information

Passed forward:
- CIK → SEC Companyfacts and submissions
- 10-K URL/text → Revenue Drivers and Risks
- ticker → yfinance market data and beta regression
- current price → valuation comparison

### Module 2 — Financial Statements
Input: CIK.

SEC Companyfacts is mapped to candidate US-GAAP tags. The app keeps duration facts separate from instant facts.

Outputs:
- approximately five annual fiscal years
- financial statements in USD millions
- diluted shares in millions
- XBRL mapping table
- CSV download
- reconciliation gate

Passed forward:
- historical financials → Ratio Analysis
- latest revenue, debt, cash, diluted shares → DCF/WACC defaults
- mapping evidence → Excel and source log

### Module 3 — Ratio Analysis
Input: reconciled historical statements.

Outputs:
- growth, margin, liquidity, leverage, efficiency, and return ratios
- average-balance denominators for ROA/ROE/asset turnover where possible
- constrained trend statements

Passed forward:
- ratios → forecasting context and AI memo prompt

The Ratio Agent can say what changed. It cannot establish why from ratios alone.

### Module 4 — Revenue Drivers
Input: latest 10-K text.

The app searches Item 1 / Item 7 evidence candidates for price, volume, shipments, mix, demand, capacity, sales, and operations.

Output:
- source excerpts for student verification
- student notes separating observed trend, management explanation, and analyst interpretation

Passed forward:
- verified evidence → memo prompt and audit

### Module 5 — Risks
Input: latest 10-K Item 1A.

Output:
- risk evidence candidates
- 4–6 student-selected risks
- transmission mechanisms linking disclosure to DCF variables

Passed forward:
- verified risk evidence → memo prompt and audit

### Module 6 — DCF/WACC
Inputs:
- latest historical revenue
- historical margins / D&A / capex as reference points
- Treasury 10-year yield
- beta
- Damodaran implied ERP
- editable cost of debt, tax rate, capital structure
- student forecast assumptions

Outputs:
- CAPM cost of equity
- WACC
- five-year UFCF forecast
- perpetuity-growth valuation
- exit EV/EBITDA valuation
- WACC / terminal-growth sensitivity
- WACC / exit-multiple sensitivity
- terminal-value dependence
- valuation football field

Passed forward:
- assumptions and outputs → Excel and memo

Important: historical ratios inform assumptions but do not mechanically determine them.

### Module 7 — Excel
Input: historical statements, ratios, WACC, DCF assumptions, source list.

Output:
- downloadable `.xlsx`
- Historical sheet
- Ratios sheet with formulas
- WACC sheet with formulas
- DCF sheet with formulas
- XBRL Mapping sheet
- Sources sheet

Control:
Students must change Revenue Growth, EBIT Margin, WACC inputs, and Terminal Growth in Excel and observe valuation recalculate.

### Module 8 — Memo
Input: only supplied evidence and calculations.

Output:
- a structured prompt for any approved AI tool
- no paid API required

The AI may synthesize but is explicitly prohibited from inventing guidance, quotations, market share, peer multiples, benchmarks, catalysts, causal explanations, or forecasts not supplied by the analyst.

### Module 9 — Audit
Input: pasted AI memo.

Output:
- 15–20 candidate material claims
- claim classification
- evidence field
- verification status
- correction field
- causal-language flags
- management-attribution flags
- second-AI preliminary audit prompt
- detailed error log

Final authority remains the student.

---

## 3. GitHub — first-time setup

1. Go to GitHub and create a free account.
2. After signing in, choose **New repository**.
3. Repository name: `equity-research-agent`.
4. Choose **Public** for the classroom prototype.
5. Select **Add a README file**.
6. Create the repository.
7. In the repository, choose **Add file → Upload files**.
8. Upload:
   - `streamlit_app.py`
   - `requirements.txt`
9. Commit the upload.

A **commit** is a saved version of the repository. Think of it as a labeled checkpoint in the project history.

### Edit directly in GitHub

1. Open `streamlit_app.py`.
2. Click the pencil/edit icon.
3. Make the change.
4. Choose **Commit changes**.
5. Enter a short message such as `Clarify DCF warning`.
6. Commit to the `main` branch.

Each commit creates a new recoverable version.

---

## 4. Streamlit Community Cloud deployment

1. Open Streamlit Community Cloud.
2. Sign in with GitHub.
3. Authorize Streamlit to access the public repository.
4. Choose **Create app**.
5. Select repository `equity-research-agent`.
6. Branch: `main`.
7. Main file path: `streamlit_app.py`.
8. Deploy.

No local Python installation is required for this basic path.

### Logs

Deployment/build logs identify:
- package-install failures
- Python errors
- network/API errors
- missing modules

If a module is missing:
1. Add its package name to `requirements.txt`.
2. Commit the change.
3. Streamlit detects the dependency change and redeploys.

Normal code commits also update the deployed app automatically.

---

## 5. Instructor first test — STLD

Run the following sequentially. Stop if an earlier accounting control fails.

1. Enter `STLD`.
2. Confirm Steel Dynamics identification.
3. Confirm SEC CIK.
4. Open latest 10-K.
5. Review five-year financial statements.
6. Reconcile Revenue.
7. Reconcile Operating Income.
8. Reconcile Operating Cash Flow.
9. Reconcile Capital Expenditures.
10. Check Ratio Analysis only after all four reconciliation boxes pass.
11. Manually recompute several ratios.
12. Review Revenue Driver evidence.
13. Review Item 1A Risk evidence.
14. Open DCF/WACC.
15. Verify current 10-year Treasury source/date.
16. Verify beta source/cross-check and ERP source.
17. Change WACC and confirm valuation moves inversely.
18. Change Year 1 Revenue Growth and confirm forecast/valuation changes.
19. Change EBIT Margin and confirm UFCF/valuation changes.
20. Review WACC / terminal-growth sensitivity.
21. Review WACC / exit-multiple sensitivity.
22. Review terminal-value share of EV; discuss dominance warning.
23. Review football field valuation ranges.
24. Download Excel.
25. In Excel, change Revenue Growth.
26. Change EBIT Margin.
27. Change WACC inputs.
28. Change Terminal Growth.
29. Confirm formulas recalculate implied prices.
30. Inspect Ratio worksheet formulas.
31. Generate/copy Memo prompt.
32. Use an approved AI tool to create the first draft.
33. Paste the memo into Audit.
34. Classify 15–20 material claims.
35. Verify every material number.
36. Correct unsupported causal statements.
37. Correct assumptions presented as facts.
38. Document at least three material AI errors when present.
39. Change ticker to another U.S. public company.
40. Repeat from Company identification; do not assume STLD mappings transfer perfectly.

---

## 6. Student submission package

1. **Python notebook or spreadsheet model.** For this version, submit the generated Excel workbook plus the Streamlit source code if required.
2. **Equity research note, 4–6 pages.**
3. **AI audit appendix.** Export the audit table and include the detailed error log.
4. **Source evidence file.** Use the source log, then add filing section/page references and screenshot filenames.
5. **One-page reflection.** Explain where the AI made mistakes, why those errors were plausible, how the student detected them, and where human financial judgment remained necessary.

---

## 7. Suggested grading controls

A useful grading structure is to separate:
- data integrity / reconciliation
- ratio accuracy
- source provenance
- forecast reasoning
- WACC construction
- DCF mechanics
- sensitivity interpretation
- Excel formula functionality
- memo quality
- hallucination audit quality
- reflection / model governance

Do not award full credit for a polished memo when the evidence chain or model controls fail.

---

## 8. Important limitations to teach explicitly

- XBRL tags vary across issuers and years; automatic mapping can fail.
- `N/A` can be a mapping problem, not an economic zero.
- SEC filing HTML does not always provide convenient stable printed-page references.
- yfinance is appropriate for classroom/research convenience but is not a primary accounting source.
- Beta estimates vary by lookback, return frequency, benchmark, and provider.
- Damodaran ERP is an input choice, not an observable fact.
- WACC is an estimate.
- Terminal value can dominate enterprise value.
- A DCF is a scenario framework, not a precise intrinsic-value oracle.
- Generative AI can summarize evidence well and still fabricate causes, guidance, benchmarks, or attribution.
- AI self-review does not replace independent human verification.
