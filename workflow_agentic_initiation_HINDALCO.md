# Agentic Equity Research Workflow: HINDALCO Case Study

**Analyst:** [Your Name]
**Date:** Feb 11, 2026
**Subject:** Hindalco Industries Ltd (NSE: HINDALCO)

---

## 1. Introduction

This document outlines the **Agentic Workflow** used to initiate coverage on Hindalco Industries. The process leveraged an **AI Investment Agent (Antigravity)** to automate data ingestion, financial modeling, and initial drafting, allowing the analyst to focus on thesis generation, valuation judgment, and final refinement.

---

## 2. Planning Phase
**Objective:** Define scope, directory structure, and deliverables.

*   **Agent Role (Automation):**
    *   Proposed a detailed **Implementation Plan** breaking down the initiation into phases (Data, Model, Note, Deck).
    *   Automatically created a structured project directory: `/data/`, `/models/`, `/notes/`, `/slides/`.
    *   Set up a task tracking system to monitor progress.
*   **Analyst Role (Review & Refine):**
    *   Defined the specific ticker (`HINDALCO`) and sector (`Metals`).
    *   Approved the milestone-based approach (Model first -> Note -> Deck).
    *   **Value Add:** Ensured the scope included critical sector-specific analysis (Novelis, LME trends) rather than just generic financials.

---

## 3. Data Collection & Structuring
**Objective:** Gather clean historical data for the 3-statement model.

*   **Agent Role (Automation):**
    *   Scraped consolidated financial data (P&L, Balance Sheet, Cash Flow) for FY21-FY25 from public sources (Screener.in).
    *   Parsed raw HTML tables into structured CSV files (`IS.csv`, `BS.csv`, `CF.csv`).
    *   Handled data cleaning (removing formatting, handling nulls) using Python scripts.
*   **Analyst Role (Review & Refine):**
    *   Verified the data source reliability.
    *   **Sanity Check:** Confirmed that key figures (Revenue, EBITDA, Net Debt) matched reported Annual Report numbers.

---

## 4. Financial Modeling
**Objective:** Build a dynamic, formula-driven DCF & Comps model.

*   **Agent Role (Automation):**
    *   Wrote Python scripts (`create_model_v2.py`) to generate a fully functioning Excel file (`three_statement_HINDALCO.xlsx`) with the `xlsxwriter` engine.
    *   **Automated Linkages:** Built formula relationships across Input, IS, BS, DCF, and Comps sheets (e.g., Revenue linked to growth drivers, Depreciation to Net Block).
    *   **Valuation Logic:** Automated the WACC calculation, Terminal Value derivation, and Enterprise Value bridge.
    *   **Scenario Table:** Created a dynamic scenario matrix (Bull/Base/Bear) linked to key drivers.
*   **Analyst Role (Review & Refine):**
    *   **Troubleshooting:** Identified file corruption issues with initial `openpyxl` engine; directed agent to switch to `xlsxwriter`.
    *   **Assumption Setting:** Defined key forecast drivers (Revenue Growth: 8%, EBITDA Margin: 14%, WACC: 10%).
    *   **Valuation Judgment:** Selected the appropriate peer set (Vedanta, Tata Steel, JSW) for relative valuation.

---

## 5. Initiation Note Drafting
**Objective:** Synthesize quantitative model outputs into a qualitative investment thesis.

*   **Agent Role (Automation):**
    *   Conducted rapid web research on global aluminium supply/demand, LME price outlook, and analyst consensus.
    *   Drafted the structured **Initiation Note** (`initiation_note_HINDALCO.md`) covering Industry, Company, Financials, and Valuation.
    *   **Synthesis:** Combined model outputs (Target Price ₹1,150) with qualitative drivers (Novelis expansion, Deleveraging).
*   **Analyst Role (Review & Refine):**
    *   **Thesis Sharpening:** Focused the narrative on the "Integrated Model" and "Novelis Downstream Optionality" as the key differentiator.
    *   **Risk Assessment:** Ensured specific risks (Bay Minette Capex overrun) were highlighted prominently.

---

## 6. Deck Creation
**Objective:** Convert the detailed note into a presentation-ready format.

*   **Agent Role (Automation):**
    *   Transformed the long-form note into a **11-slide Thesis Deck** (`initiation_deck_HINDALCO.md`).
    *   Structured content into bullet points suitable for slides.
    *   Extracted key tables (Valuation Summary, Comps) directly from the model logic.
*   **Analyst Role (Review & Refine):**
    *   Approved the slide flow and logic.
    *   **Storytelling:** Verified that the "Buy" recommendation was supported by a cohesive argument across all slides.

---

## Conclusion for Interview
This workflow demonstrates how **IA (Intelligent Automation)** shifts the analyst's role from "Data Processor" to "Investment Straegist." By automating 80% of the manual data entry and formatting, I could dedicate 100% of my intellectual energy to **generating a differentiated investment thesis and validating critical assumptions.**
