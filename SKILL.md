---
name: template-based-business-analysis-report
description: Create, revise, or quality-check business analysis reports covering business problems, market context, operating models, processes, capabilities, financial logic, risks, and action recommendations.
---

**Bundled Template Assets**

Use the **Skill directory** shown above as the base directory. All relative paths below resolve from that directory.

Bundled files:
- Template name: Business Analysis Visual Word Template
- Template file: `./business-analysis-visual-word-template.docx`
- Generator file: `./generate_business_analysis_visual_word_template.py`

When generating or revising a Word document, copy the bundled Python generator and `.docx` template from the Skill directory into the active workspace. Prefer editing and running the copied generator. Write a new Python generator only if the bundled generator cannot support the requested output.

**Target write language**: {{ product_language }}

For every writing artifact, all user-visible text MUST use {{ product_language }} unless the user explicitly requests another language.

# Business Analysis Report

Use this skill to create or improve decision-ready business analysis reports, including market, operational, financial, process, capability, and strategy analysis.

The goal is to convert facts and evidence into clear findings, implications, and actions the intended audience can use.

## 1. Methodology Selection

Choose methods by decision goal, business context, and audience. Use only the methods that improve the report; do not force every framework into every analysis.

| Decision goal | Recommended methods | Output |
|---|---|---|
| Define the business problem | IIBA BACCM + stakeholder analysis + problem statement | Need, context, stakeholders, value, solution boundary |
| Understand market or industry context | PESTEL + Porter's Five Forces + market sizing | External drivers, industry attractiveness, opportunity/risk |
| Analyze the business model | Business Model Canvas + Value Proposition Canvas | Customer, value, channel, revenue, cost, resource logic |
| Diagnose operational or process issues | BPMN/process map + SIPOC + root cause analysis | Current-state process, bottlenecks, root causes |
| Compare current vs. future state | Gap analysis + capability map + requirements traceability | Gap list, capability needs, requirements, priorities |
| Prioritize initiatives | MoSCoW + impact/effort + risk/value scoring | Prioritized roadmap and decision rationale |
| Prove financial or performance logic | KPI tree + cost-benefit/ROI + sensitivity analysis | Metrics, assumptions, scenarios, financial implications |
| Turn analysis into execution | Roadmap + RACI + risk register + benefits tracking | Actions, owners, milestones, risks, success metrics |

Selection rule: start from the user's decision goal, not from a default framework. Prefer 2-4 relevant methods. Every method used must produce a concrete section, table, chart, or decision implication.

**Methodology references:**
- IIBA Business Analysis Standard and BABOK ecosystem: https://www.iiba.org/knowledgehub/the-business-analysis-standard/
- IIBA BABOK standards: https://www.iiba.org/standards-and-resources/babok/
- OMG BPMN standard: https://www.omg.org/bpmn/
- ASQ root cause analysis tools: https://asq.org/quality-resources/root-cause-analysis/tools
- Strategyzer Business Model Canvas: https://www.strategyzer.com/library/what-is-a-business-model
- Harvard Business School Porter's Five Forces: https://www.isc.hbs.edu/strategy/business-strategy/Pages/the-five-forces.aspx
- Balanced Scorecard overview: https://strategymanage.com/resources/about-the-balanced-scorecard/
- McKinsey 7-S framework: https://www.mckinsey.com/business-functions/strategy-and-corporate-finance/our-insights/enduring-ideas-the-7-s-framework

## 2. Business Analysis Drafting Guide

Draft around six failure points:

- **Decision fit:** define the audience, business question, decision, geography, time period, and required output before analyzing.
- **Fact discipline:** extract core variables such as company, product, market, date, geography, stakeholder, process, KPI, currency, unit, and assumption; keep them consistent across all sections.
- **Evidence chain:** distinguish verified facts, user-provided claims, inference, and assumptions; cite key claims with source name, URL/file, date, and confidence.
- **Coverage:** cover every requested question, source, business unit, market, segment, process, KPI, and deliverable; mark missing items as `[TBD]`, `[Not yet analyzed]`, or `[Needs verification]`.
- **Metric logic:** define formulas, units, currency, time period, denominator, geography, and source for key metrics; use code-assisted calculation for financial or numeric analysis.
- **Actionability:** connect findings to implications, recommendations, owners/functions, priority, expected impact, risk, validation metric, and next step.

## 3. Document Structure

**Priority:** If the user gives a required structure, section order, page limit, rubric, board format, investor format, internal memo format, supplied source document, or existing draft, follow it first unless it creates a direct contradiction or materially weakens the analysis.

### Core Modules

Use these modules for a complete business analysis report unless the user's source document requires a different order:

```text
Title / Cover
Executive Summary
Decision Goal And Audience
Business Context And Scope
Methodology And Source Register
Current-State Findings
Market / Customer / Industry Context if relevant
Process / Operating Model / Capability Analysis if relevant
Financial / KPI / Scenario Analysis if relevant
Gap Analysis And Root Causes
Options Or Recommendations
Implementation Roadmap
Risks, Assumptions, And Limitations
Appendices / Source List
```

For shorter documents, keep the same decision logic but compress sections instead of dropping fundamentals.

### Do Not Miss

- Define decision goal, audience, scope, geography, time period, and intended action.
- Extract key facts and variables before drafting; keep names, dates, metrics, units, currencies, and assumptions consistent.
- Provide source and date for key market, company, financial, operational, regulatory, and benchmark claims.
- Use coverage matrices when the user provides multiple questions, sources, business units, markets, or deliverables.
- Separate facts, interpretation, assumptions, and recommendations.
- Convert each major finding into a strategic or operational implication.
- Make recommendations specific enough to execute or validate.
- Preserve confirmed prior content during partial edits and multi-turn revisions.

### Useful Tables

Fact variable register:

| Variable | Value | Source | Used In | Status |
| --- | --- | --- | --- | --- |

Source register:

| Claim / Data Point | Source | Source Type | Date | Confidence | Notes |
| --- | --- | --- | --- | --- | --- |

Coverage matrix:

| Requirement / Question | Source(s) Checked | Covered? | Output Section | Gap / Next Step |
| --- | --- | --- | --- | --- |

Metric dictionary:

| Metric | Definition | Unit / Currency | Period | Formula | Source / Assumption |
| --- | --- | --- | --- | --- | --- |

Findings-to-actions matrix:

| Finding | Evidence | Implication | Recommendation | Owner / Function | Metric |
| --- | --- | --- | --- | --- | --- |

Risk register:

| Risk | Why It Matters | Evidence | Mitigation | Next Validation |
| --- | --- | --- | --- | --- |

## 4. Template Execution

Use `business-analysis-visual-word-template.docx` as the reference layout for formal Word outputs. Use `generate_business_analysis_visual_word_template.py` as the source generator. Copy both files into the active workspace, prefer editing the copied generator, run it from the workspace, and verify the regenerated Word file. Generate another format directly only when the user requests a non-Word artifact.

Execution rules:

- Ask no more than 3 questions. If facts are missing, state assumptions and continue.
- Use `[TBD]` for missing audience, decision goal, business unit, geography, date, source, metric, unit, currency, owner, or recommendation.
- Preserve user-provided reports, slides, spreadsheets, URLs, screenshots, dashboards, financials, rubrics, prior drafts, marked revisions, and requested structures unless instructed otherwise.
- Treat fact registers, source registers, coverage matrices, KPI tables, process maps, findings, recommendations, and executive summaries as modules of the same business analysis.
- If the user asks for a partial edit, keep confirmed content unchanged and only modify the authorized scope.
- If the user requests a specific output format, verify that format before delivery.

## 5. Business Analysis Quality Check And Fixes

Run the mandatory business-analysis quality review every time. For formal files or template-backed outputs, create `quality_check.md` beside the generated document. For lightweight chat answers, include a concise quality-check summary in the response instead.

### Mandatory Business Analysis Quality Check

Check six areas:

- **Fact-variable extraction and document-wide consistency:** company, market, product, stakeholder, process, KPI, date, geography, unit, currency, assumptions, and named entities are extracted and consistent across title, summary, body, tables, charts, appendices, and file names.
- **Research evidence chain, source verification, and recency:** key claims have source name, URL/file, date, source type, and confidence; time-sensitive market, company, regulatory, pricing, financial, and KPI facts are current or marked `[Needs verification]`.
- **Coverage matrix and completeness check:** all user questions, source files, URLs, screenshots, data tables, business units, markets, processes, deliverables, and prior confirmed instructions are covered or explicitly marked as `[TBD]` / `[Not yet analyzed]`.
- **Numeric metrics, finance, and unit consistency:** formulas, units, currencies, denominators, periods, scenarios, rankings, totals, and assumptions are defined and consistent; financial outputs are calculated with a reproducible workflow when needed.
- **Audience, scenario, localization, and tone fit:** structure, depth, terminology, language, examples, geography, compliance context, and tone match the intended audience and use case.
- **Findings-to-action recommendation loop:** findings lead to implications, recommendations, owners/functions, priorities, expected impact, risks, validation metrics, and next steps.

If fixable issues exist, fix the document, regenerate the output if needed, and update `quality_check.md` or the concise quality-check summary.

Use this `quality_check.md` structure for formal outputs:

```markdown
# Business Analysis Quality Check

## Checks
- Fact-variable extraction and document-wide consistency: OK / Fixed / Issues / TBD
- Research evidence chain, source verification, and recency: OK / Fixed / Issues / TBD
- Coverage matrix and completeness check: OK / Fixed / Issues / TBD
- Numeric metrics, finance, and unit consistency: OK / Fixed / Issues / TBD
- Audience, scenario, localization, and tone fit: OK / Fixed / Issues / TBD
- Findings-to-action recommendation loop: OK / Fixed / Issues / TBD

## Material Risks
- Section/Area; risk; current treatment; severity; fix applied.

## Remaining Fixes Needed
- [TBD]
```
