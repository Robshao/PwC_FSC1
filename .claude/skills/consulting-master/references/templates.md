# Templates and rubrics

Contents: 1 Problem statement · 2 Stakeholder and decision frame · 3 MECE check · 4 Hypothesis card · 5 Prioritization table · 6 Analysis plan · 7 Contribution table · 8 Recommendation · 9 Executive summary (SCQA) · 10 Full deliverable outline · 11 Grading rubric (Coach and Review modes)

Render these in the user's language. Keep table columns; adapt wording.

## 1. Problem statement

**Initial (before diagnosis), 3 to 4 sentences:**
> [Metric] for [scope: segment, product, region, channel] moved from [baseline] to [current] over [period] ([+/- x percentage points or x%]). [Who holds which view, stated as views, not facts.] This matters because [business impact tied to an objective]. Success means [analysis deliverable] by [analysis deadline], so that [decision], with the goal of [outcome target] by [outcome deadline].

**Diagnosed (after root cause):** add "primarily driven by [confirmed drivers, with share of the change]" and layered success metrics.

Checks: specific · measurable with baseline and period · who is affected · why it matters · target is achievable · deadlines split (analysis vs outcome) · no cause embedded in the initial version · rates written in percentage points.

## 2. Stakeholder and decision frame

| Stakeholder | Present? | Cares about | Favorite explanation (to test) | Power | Interest | How to engage |
|---|---|---|---|---|---|---|

Decision to support (one sentence, concrete: amount, scope, timing):
Primary success metric (at most two, with reason):

## 3. MECE check (per level)

| Check | Question | Result and rule |
|---|---|---|
| Mutually exclusive | Can any concrete item fit two branches? | |
| Collectively exhaustive | Did 20+ brainstormed items all find a home? Internal and external, demand and supply, measurement? | |
| One dimension | Is every branch on this level cut by the same logic? | |
| Same kind | All causes, or all steps, or all solutions? All totals or all rates? | |
| Scope | Does every branch belong to this problem? | |
| Actionable | Does each branch map to a lever someone owns and data that exists? | |

## 4. Hypothesis card

| Field | Content |
|---|---|
| Driver | We believe ... |
| Outcome | ... is causing [observable outcome] |
| Mechanism | because ... (how the driver produces the outcome, not the evidence) |
| Type | Descriptive / explanatory / solution |
| Expected evidence | If true, the data should show ... |
| Refuted if | If we see ..., the hypothesis is rejected |
| Alternatives | At least two other plausible explanations |
| Confounders | What else could produce the same pattern, and how we control for it |
| If true, then | Action tied to the decision |
| If false, next | Where the search moves next |

## 5. Prioritization table

| Hypothesis | Impact | Likelihood | Speed to test | Already known? | Depends on | Priority and reason |
|---|---|---|---|---|---|---|

Rules: already known means quantify, do not test; dependencies go first; slow but high-impact items get their data request on day one; everything else gets a quick descriptive check.

## 6. Analysis plan

| Hypothesis | Expected evidence | Refuted if | Method | Data fields | Source and owner (named) | Expected output | Timeline |
|---|---|---|---|---|---|---|---|
| Data quality (all) | Clean joins, stable definitions | Definitions differ | Profiling, reconciliation | Keys, dates, codes | IT or data owner | Data issues log | Day 1 to 2 |

## 7. Contribution table

| Driver | Estimation method | Contribution (points or amount) | Share of change | Confidence |
|---|---|---|---|---|
| Other / unexplained | Residual | | | Low |
| Total | | | 100% | |

State the rule for interaction effects (for example, assign the price × volume cross term to price) and explain negative contributions.

## 8. Recommendation

| Driver found | Recommendation | Expected impact | Timeline | Risks and who is affected | Success metric | Tracking owner and early warning |
|---|---|---|---|---|---|---|

Overall call when a go/no-go decision was framed: Proceed / Modify / Stop / Test further, with 2 to 3 sentences linking evidence to the call.

## 9. Executive summary (SCQA)

- Situation: the stable baseline.
- Complication: what changed, with the number.
- Question: what leadership needs answered.
- Answer: the conclusion first, then the two or three drivers with their share, then the recommendation.
Then list what was checked and ruled out.

## 10. Full deliverable outline

1. Working assumptions to confirm
2. Decision and stakeholders
3. Problem statement
4. MECE structure with rules
5. Where the problem sits (bridge and is/is-not)
6. Hypothesis tree (Mermaid) and cards
7. Prioritization
8. Analysis plan
9. Findings and contributions (only if data exists; otherwise "to be completed")
10. Recommendation and next steps

Use Mermaid for the tree. If the user needs an image or a document, produce it with the appropriate file skill.

## 11. Grading rubric (Coach and Review modes)

Score each 0 to 2, total 20. Point out the single most important fix.

| Criterion | 2 points looks like |
|---|---|
| Decision framing | A concrete decision is named and the analysis serves it |
| Metric validation | Definition, sources and traps checked; right first cut for the metric type |
| Problem statement | SMART, who is affected, no embedded cause, points not percent for rates |
| MECE first cut | Identity-based where possible; one dimension per level |
| MECE honesty | Overlaps and gaps admitted with boundary rules |
| Localization | Where / when / is-not used before causes |
| Hypothesis quality | Falsifiable, mechanism stated, alternatives named |
| Prioritization | Impact, likelihood, speed plus dependencies and already-known items |
| Analysis design | Pass/fail set in advance, confounders named, data owner named |
| Synthesis | Contributions to 100%, confidence, recommendation tied to decision |
