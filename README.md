# Commercial Valuation & Contract Pricing Engine

[![Streamlit App](https://static.streamlit.io/badges/streamlit_badge_black_white.svg)](https://commercial-valuation-engine.streamlit.app/)
![Python](https://img.shields.io/badge/python-3.10%20%7C%203.11%20%7C%203.12-blue)[cite: 2]
![License](https://img.shields.io/badge/license-MIT-green)[cite: 2]

An automated capital budgeting and commercial valuation engine designed for multi-year facility management (FM) service contracts and corporate tenders[cite: 2]. The application translates operational headcounts and shift schedules into fully loaded cost models, evaluates Net Working Capital (NWC) drag, solves for Discounted Cash Flow (DCF) yield metrics, and validates pricing against parametric Class 5 Cost Estimating Relationships (CER)[cite: 2].

---

### [Launch Interactive Web Application](https://commercial-valuation-engine.streamlit.app/)

---

## Core Capabilities

* **Fully Loaded Direct & Indirect Cost Factoring:** Builds monthly operational expenditures from baseline labor tiers (Managers, Supervisors, Workers) and dynamically applies statutory burdens (GOSI, health, taxes, uniforms at 43.13%), dedicated worker transit (41.33%), cleaning consumables (4.58%), equipment servicing (1.67%), SG&A overhead (9.77%), and municipal stamp duties (11.62%)[cite: 2].
* **Shift Schedule Normalization:** Computes standard hourly adjustments across 4 shift regimes: 8h/7d (240h), 8h/6d (208h), 12h/7d (360h), and 12h/6d (312h)[cite: 2].
* **NWC Cash Flow Drag Modeling:** Accurately deducts annual working capital lockup from operating cash flows using Days Sales Outstanding (DSO) client billing cycles and Days Payable Outstanding (DPO) vendor credit lags[cite: 2].
* **Discounted Cash Flow (DCF) Valuation:** Evaluates multi-year project cash flows against hurdle rates (WACC), producing exact Net Present Value (NPV)[cite: 2].
* **Newton-Raphson IRR Solver:** Custom numerical polynomial iteration solving the Internal Rate of Return without external binary dependencies[cite: 2].
* **Closed-Form Breakeven Pricing:** Analytically evaluates the minimum required monthly contract price where $\text{NPV} = 0$[cite: 2].
* **AACE Class 5 CER Benchmarking:** Validates bottom-up tender quotations against empirical parametric power-law regression curves ($P_{50} = 10,089.86 \times \text{Headcount}^{1.0874}$) bounded within $-30\%$ to $+50\%$ confidence intervals[cite: 2].
* **Dynamic In-Memory Excel Pipeline:** Dynamically generates stylized dual-sheet OpenXML templates (`Pricing_Inputs_Template.xlsx`) and structured financial outputs (`Commercial_Valuation_Results.xlsx`) via in-memory byte buffers (`io.BytesIO`)[cite: 2].

---

## Mathematical & Financial Architecture

### 1. Free Cash Flow (FCF) & Working Capital Drag[cite: 2]
$$\text{EBIT} = (\text{Monthly Price} - \text{Opex}_{\text{monthly}} - \text{Depreciation}_{\text{monthly}}) \times 12$$

$$\Delta\text{NWC}_{\text{annual}} = \frac{\left(\text{Revenue}_{\text{annual}} \times \frac{\text{DSO}}{365}\right) - \left(\text{Opex}_{\text{annual}} \times \frac{\text{DPO}}{365}\right)}{\text{Contract Tenor}}$$

$$\text{FCF}_t = \text{EBIT}_t \times (1 - T) + \text{Depreciation}_t - \Delta\text{NWC}_{\text{annual}}$$

### 2. Net Present Value (NPV)[cite: 2]
$$\text{NPV} = \sum_{t=1}^{n} \frac{\text{FCF}_t}{(1 + \text{WACC})^t} - \text{Capex}_0$$

### 3. Newton-Raphson Internal Rate of Return (IRR)[cite: 2]
Finds root $r$ such that $f(r) = 0$:
$$r_{k+1} = r_k - \frac{f(r_k)}{f'(r_k)} = r_k - \frac{\sum_{t=0}^{n} C_t (1+r_k)^{-t}}{-\sum_{t=0}^{n} t \cdot C_t (1+r_k)^{-(t+1)}}$$

### 4. Analytical Breakeven Monthly Price ($\text{NPV} = 0$)[cite: 2]
$$\text{Monthly Price}_{\text{BE}} = \frac{\text{Capex}_0 + \left(\sum_{t=1}^{n} \frac{1}{(1+\text{WACC})^t}\right) \times \text{Cost Burden}}{\left(\sum_{t=1}^{n} \frac{1}{(1+\text{WACC})^t}\right) \times \text{Revenue Coefficient} \times 12}$$

---

## Live Workflow

1. **Access the Engine:** Open the [Streamlit Application](https://commercial-valuation-engine.streamlit.app/).
2. **Download Benchmark Template:** Click **Download Pricing_Inputs_Template.xlsx** to retrieve the structured workbook featuring the `INPUTS` table and the operational parameter `REFERENCE_GUIDE`[cite: 2].
3. **Populate Parameters:** Specify tender identifiers, operational headcount allocations, shift schedules, target commercial markups, WACC discount rates, tax rates, and working capital credit periods (DSO/DPO)[cite: 2].
4. **Execute Valuation Pipeline:** Upload the modified file (or evaluate pre-loaded commercial contract presets)[cite: 2].
5. **Inspect & Export:** Review dynamic valuation tables on screen and download the finalized audit-ready workbook (`Commercial_Valuation_Results.xlsx`)[cite: 2].

---

## Repository Structure[cite: 2]

```text
├── app.py                   # Streamlit interactive application & OpenPyXL pipeline
├── commercial_valuation.py  # Standalone CLI/batch valuation engine
├── requirements.txt         # Package dependencies
└── README.md                # Technical specification and documentation