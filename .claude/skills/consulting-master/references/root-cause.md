# Root cause and testing reference

Contents: 1 Symptom, contributing factor, root cause · 2 5 Whys discipline · 3 Stop rules · 4 Fishbone categories · 5 Quick tests and red flags · 6 Pseudo root causes · 7 White space · 8 Testing designs and traps · 9 When to use which tool

## 1. Symptom, contributing factor, root cause

| Layer | What it is | Test |
|---|---|---|
| Symptom | A visible outcome: declines, complaints, rework, outages, backlog | You can observe it directly |
| Contributing factor | Explains this instance, but removing it alone does not stop recurrence | "If this had not happened, would the problem still appear, only later or smaller?" If yes, it is contributing |
| Root cause | A specific, controllable driver in the system; fixing it lowers likelihood or impact across cases | Passes the stop rules below |

Good root causes usually read as a missing condition: "no quality check before release", "no handoff template", "incentives reward shipments not sales".

## 2. 5 Whys discipline

- Five is a guideline. Stop by the rules in section 3, not by the count.
- Anchor every "why" in evidence (data, interview, document). Without evidence it is a story.
- At each level ask whether there is more than one answer; a single chain can miss parallel causes. Use MECE on the first "why".
- If an answer leaves the client's control ("customers can't take photos"), step back one level and find the controllable condition ("no real-time image quality check").

## 3. Stop rules (all should hold)

1. Actionable: there is a clear thing to change.
2. Within management control.
3. Systemic: process, design, policy, structure, capability, tool, incentive.
4. Not itself caused by another controllable factor within scope.
Stop signals: further whys add no insight, drift into vague territory ("culture", "people are careless"), or leave the organization's control.

## 4. Fishbone categories (use all seven as a checklist)

People · Process · Technology · Materials / documents · Environment (mostly external: quantify and exclude) · Measurement (definitions, missing codes, changed reporting) · Management and policy (incentives, authority limits, KPIs, budgets).
Fishbone generates hypotheses; it does not declare causes and does not guarantee mutual exclusivity. Measurement and management/policy are the two most often forgotten.

## 5. Quick tests and red flags

Three quick tests (a "no" means you are probably still at symptom level):
1. Outcome or mechanism? A cause describes a process, structure or decision.
2. Control: can the client change it? (An external factor is real but not a fixable root cause.)
3. Recurrence: if fully fixed, would the problem almost certainly not come back for this reason?

Red flags:
- Restating the problem as its cause ("late because timelines are unrealistic").
- Vague labels ("poor communication", "lack of accountability").
- Band-aid fixes: extra approvals, reports, meetings. In regulated industries some controls are mandatory; ask whether a control was required by regulation or added after an incident, and whether the original cause was fixed. Check with compliance before removing any control.
- The same issue reappearing across teams or time after "fixes".

Strong root-cause statement: names a process, rule, structure or capability; observable or measurable; within control; explains several symptoms (supporting evidence, not a requirement). You can say how to test the change, who owns it and which metrics should move.

## 6. Pseudo root causes

Treat "culture", "communication", "mindset", "accountability" as starting points. Ask: where exactly does it break down, between whom, about what, at which moment in the workflow, and which behaviors do we observe? Translate into mechanisms: missing forum, unclear decision rights, no handoff template, incentive design.

## 7. White space

Most operational problems sit in handoffs between departments that nobody owns. On a swimlane map, every arrow crossing lanes is a handoff. For each ask: what information passes and in what form; who tracks it after handoff; how long it waits and whether there is a time limit; how often it comes back and why. Fixes such as more training or headcount cuts miss problems that live here.

## 8. Testing designs and traps

| Design | Good for | Trap |
|---|---|---|
| Trend comparison | Timing of a change | Seasonality; base effects |
| Segment comparison (difference in proportions) | Is the gap real between groups | Selection effect: groups differ for other reasons |
| Within-segment comparison | Controlling a known confounder | Needs enough volume per segment |
| Cohort analysis (before vs after) | Effect of a launch or change | Anything else that changed in the same period |
| Same-period comparison of exposed vs not exposed | Isolating one change | How exposure was assigned (random, by segment, self-selected) |
| Logistic regression | Probability of a yes/no outcome with controls | Controlling for a mediator hides part of the effect; report with and without |
| A/B test | Solution hypotheses | Traffic and sample size; development effort; regulatory review of customer-facing changes |
| Sensitivity model | Trade-offs of an action | Assumptions must be stated |
| Interviews, operational walkthrough | Mechanism, workarounds outside the system | Small samples; pair with data |

Three causal traps to name in every plan: confounding (a third factor drives both), reverse causality (the outcome drives the supposed cause), measurement bias (who responds, how a field is recorded). Check the time order of cause and effect.

## 9. When to use which tool

| Situation | Start with |
|---|---|
| Scope is vague: no agreement yet on what the problem is or how big | MECE structure |
| Problem is clear but the cause is vague and time is short | Hypothesis-driven testing |
| Complex system with many interacting departments or steps | MECE decomposition to map it first |
| Symptoms are clear, cause is not, or fixes keep failing | Root cause analysis |
| Speed matters and data exists | Hypothesis-driven testing |
| Data is slow or unavailable | Qualitative tests (interviews, walkthroughs) while the data request runs |
| Client arrives with an answer and wants it "validated" | Hypothesis-driven, with the refutation condition agreed with the client up front |

"Ambiguous" means two different things; decide which kind before choosing a tool. In practice the tools loop: MECE to structure, root cause thinking to generate candidate causes, hypotheses to test them, root cause again to dig into confirmed branches.
