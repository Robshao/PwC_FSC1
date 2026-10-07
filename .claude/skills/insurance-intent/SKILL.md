---
name: insurance-intent
description: Turn an insurance client's business problem into a strong intent.md, then a spec, a clickable prototype with fake data, business-rule evals and an engineering handoff package, so a business analyst can validate with the client before RD engineering starts. Use when the user describes an insurance (life or non-life / P&C) client problem, a process-improvement idea (underwriting, policy issuance, premium collection, policy servicing, endorsement, renewal, claims, reinsurance), asks for an intent, BRD/FSD-style requirement, prototype, or "handoff to RD", or wants to apply the ai-native-sdlc skill to an insurance engagement.
---

# Insurance intent → prototype → RD handoff

Apply the `ai-native-sdlc` loop to an insurance engagement. The business analyst owns stages 1 to 4 and the client validation gate. RD owns engineering refinement. **The main thing handed over is a sharp intent plus executable business rules.** Prototype code is a reference, not production code.

Write artifact bodies in the user's language (default: verbose Traditional Chinese, Taiwan usage). Write every abbreviation as 縮寫（English Full Name，中文）. Keep life and non-life separate; never treat P&C as "short life insurance".

## Guardrails (check before anything else)
- **No real client data** in prompts, files or prototypes. Use fictional companies and fake records. If the user pastes real policyholder data, stop and ask them to anonymize it.
- Follow the firm's and the client's AI-tool policy. If unknown, list it as an open question in intent.md.
- No employer or client branding on prototypes. Mark every prototype screen **「原型（Prototype）— 非正式系統」**.
- The prototype validates *what* (flow, rules, data needed). RD decides *how* (architecture, integration, security, performance).

## Step 1: intent.md (insurance template)

```
# Intent：<專案名稱>
- 發起人／客戶角色／日期：
- 業別：壽險（Life）｜產險（Non-life / P&C）
- 價值鏈環節：商品｜通路｜核保｜承保發單｜收費｜保全／批改｜續保｜理賠｜再保｜投資 ALM
## 問題（客戶原話）
## 現況基準線（Baseline，附資料來源與期間）
  例：理賠平均結案天數 9 天、補件率 30%、自動核保率（STP，Straight-Through Processing）35%、綜合成本率 104%
## 目標（SMART：具體、可衡量、可達成、相關、有時限）
## 受影響的人與系統（利害關係人、核心保單系統、理賠系統、外部單位）
## 業務規則（已知的先列，未知的放待解問題）
## 法規與限制（公平待客原則 TCF、個人資料保護法、洗錢防制 AML、人工智慧指引、IFRS 17 / TW-ICS 影響）
## 不做的範圍（Out of scope）
## 待解問題（每題附負責人）
## 事實來源（Source of Truth）
```

**Three intent-strength tests.** Do not move to spec until all three pass:
1. **有數字**：baseline and target are numeric, with a period and a data source.
2. **可追溯**：every requirement you expect to write maps to a pain point and a KPI.
3. **陌生人測試**：someone absent from the workshop could build the same thing from this file alone.
Report each test as PASS or FAIL with the missing piece.

## Step 2: spec.md
Numbered requirements (FR-xx), each traced to the intent · To-Be process with swimlanes (actor / system / manual) · **business-rule table** (condition → outcome, e.g. 險種 × 事故 → 必附文件) · exceptions · acceptance criteria · data fields needed · flagged concerns with owner (compliance, actuarial, IT).

## Step 3: prototype
- Single-folder, dependency-light (plain HTML/JS unless RD has agreed a framework in advance; if agreed, use theirs so code is reusable).
- Put business rules in **one module** (for example `rules.js`) shared by the UI and the evals. This is the most reusable code for RD.
- Fake data only. Show the KPI the intent promised to move, so the client can see it.

## Step 4: evals (business rules as tests)
One command, non-zero exit on failure. Each spec rule becomes at least one case. Include edge cases the client named. Client walkthrough = UAT-lite (User Acceptance Testing): log each finding as pass, fix or new requirement.

## Step 5: handoff package to RD (`HANDOFF.md`)
| Item | Status for RD |
|---|---|
| intent.md, spec.md | Contract: what and why |
| rules module + evals | **Reuse**: becomes regression tests |
| prototype UI | Reference for flow and fields; RD may rewrite |
| decisions.md | Why choices were made, who approved |
| open questions | Must close before build |
| NOT covered | auth and permissions, core-system integration, real data migration, security review, performance, audit logging |

Agree this table with RD **before** the prototype is built, not after.

## Output checklist
- [ ] Guardrails checked (fake data, AI policy, prototype label)
- [ ] intent.md passes the three tests
- [ ] spec.md rules table traced to intent
- [ ] prototype + shared rules module
- [ ] evals pass
- [ ] HANDOFF.md with reuse/rewrite split and open questions
