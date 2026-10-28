---
name: 21-strategy-skills
description: Consulting-style strategy toolkit that holds 21 methods and runs only the one (or the short chain) that matches the user's actual situation. Use for business strategy problems — where the business stands, why growth stalled, testing a plan's assumptions, market sizing, competitors, customer segments, value-chain profit, choosing between options, pricing diagnosis, business cases (NPV/IRR), portfolio capital allocation, operating model and decision rights, prioritizing initiatives, transformation roadmaps, war games, strategy risk registers, KPI design, tracking promised benefits, stakeholder buy-in, board storylines, and one-page decision memos. Triggers on 策略、現況診斷、成長停滯、市場規模、競爭分析、客群區隔、利潤池、策略選項、定價診斷、商業論證、投資組合、營運模式、優先排序、轉型路線圖、兵棋推演、風險登記、KPI、效益追蹤、利害關係人、董事會簡報、決策備忘錄.
---

# 21 Strategy Skills (router)

One entry point for 21 consulting methods. Each method lives in `references/<method>.md`. Your job is to **diagnose the situation first, then load and run only the method(s) that fit**. Never run all 21, and never pad an answer with frameworks the situation did not call for.

## How to use this skill

1. **Read the situation.** Identify the real question behind the request: what decision is being made, at what stage of the work, and what the user already has (data, a draft plan, a finished analysis, an audience).
2. **Route.** Match the situation to one method using the routing table below. Use the tie-breakers when two methods look close.
3. **Load the method.** Read `references/<method>.md` in full before doing the work. Follow its steps, hypotheses, quality checks and deliverable. Do not work from memory of the method.
4. **Say what you picked and why** in one or two lines at the top of the answer (e.g. "This is a stalled-growth question, so I'm using Growth Barriers rather than a full situation assessment."). Name one runner-up only if it was a close call.
5. **Chain only when the request truly spans stages.** If the user asks for something end to end, propose the short chain (usually 2–4 methods, see "Common chains") and confirm before running it. Run them in order, with each method's output feeding the next.
6. **Ask, don't guess, when it is ambiguous.** If the request could reasonably go to two methods with different deliverables, ask one short question naming both options. If the user only wants a quick answer, give the quick answer using the method's core logic without the full deliverable.
7. **Work with what exists.** When data is missing, run the method on what is available, mark assumptions and gaps clearly, and list the data that would close them. Don't invent numbers presented as facts.
8. Reply in the user's language (Traditional Chinese if they write in Chinese).

## Routing table

| If the situation is… (signals) | Use | File |
|---|---|---|
| **Diagnosis & framing** | | |
| New engagement or kickoff; "where do we actually stand?"; leadership's story doesn't match the numbers; need a fact baseline before recommending anything | Situation Assessment | `references/situation-assessment.md` |
| Revenue plateaued or decelerating; "why aren't we growing?"; teams blaming each other; need the single binding constraint | Growth Barriers | `references/growth-barriers.md` |
| A plan or business case exists and confidence outruns evidence; "what are we assuming?"; before committing capital; pre-mortem | Assumption Audit | `references/assumption-audit.md` |
| **Market & competitive intelligence** | | |
| "How big is the market?"; TAM/SAM/SOM; market entry; where's the white space | Market Mapping | `references/market-mapping.md` |
| Competitive threat; "what will competitor X do next?"; Five Forces; capability comparison vs. rivals | Competitive Intel | `references/competitive-intel.md` |
| "Who is our customer, really?"; generic personas don't drive trade-offs; choosing which customers to focus on | Customer Segmentation | `references/customer-segmentation.md` |
| Revenue grows but margin doesn't; where is the money in the value chain; vertical integration / channel / unbundling decision | Profit Pool Analysis | `references/profit-pool-analysis.md` |
| **Strategic choice & economics** | | |
| One big decision framed as yes/no; nobody can name the second-best option; need alternatives scored on explicit criteria | Strategic Options | `references/strategic-options.md` |
| Routine discounting, margin erosion with healthy win rates, unsure if price is too high/low/wrongly structured; pricing a new tier | Pricing Strategy | `references/pricing-strategy.md` |
| Must justify an investment in numbers; NPV/IRR/payback; "what does this case depend on?"; build vs. buy economics | Business Case Builder | `references/business-case-builder.md` |
| Several business units/products/bets competing for capital; allocation by habit; grow/hold/fix/exit calls | Portfolio Review | `references/portfolio-review.md` |
| **Operating model & execution** | | |
| Strategy decided but "who owns this?" is unclear; org structure, decision rights (RACI/RAPID), hand-offs | Operating Model Design | `references/operating-model-design.md` |
| Too many initiatives (often 15+); competing priority lists; need a kill list and sequencing within capacity | Initiative Prioritizer | `references/initiative-prioritizer.md` |
| Priorities chosen, now need phases, owners, milestones, and "what happens Monday / first 90 days" | Transformation Roadmap | `references/transformation-roadmap.md` |
| **Risk, performance & value governance** | | |
| Before a major move (entry, M&A, price change): simulate competitor, customer and execution responses over several turns | War Gaming | `references/war-gaming.md` |
| Board asks "what could go wrong and who's watching?"; need a strategy-specific risk register with owners and leading indicators | Risk & Mitigation | `references/risk-mitigation.md` |
| Defining success measures; dashboard full of vanity metrics; north-star, KPI tree, thresholds | KPI Architect | `references/kpi-architect.md` |
| Business case approved/funded; need to prove promised benefits actually land; stage-gate funding on proven value | Value Realization | `references/value-realization.md` |
| **Alignment & executive communication** | | |
| Who must say yes; pre-wiring before a board/ExCo meeting; blockers and swing votes; tailoring the message per person | Stakeholder Alignment | `references/stakeholder-alignment.md` |
| Board deck, executive readout, partner pitch; need a storyline (SCQA, Pyramid Principle) that lands in 60 seconds | Narrative Builder | `references/narrative-builder.md` |
| Analysis done; need a one-page, answer-first memo naming the exact approval requested | Decision Memo | `references/decision-memo.md` |

## Tie-breakers for look-alike methods

- **Situation Assessment vs. Growth Barriers** — Broad "where are we" across the whole business → Situation Assessment. The specific symptom is growth slowing → Growth Barriers.
- **Assumption Audit vs. Risk & Mitigation vs. War Gaming** — Testing the *beliefs* a plan rests on → Assumption Audit. Listing *events* that could hurt a chosen plan, with owners → Risk & Mitigation. Simulating *active responses* by competitors/customers over turns → War Gaming.
- **Competitive Intel vs. War Gaming** — Understanding and predicting rivals (input to strategy) → Competitive Intel. Stress-testing *our* specific move against their reaction → War Gaming.
- **Market Mapping vs. Customer Segmentation** — Sizing a market and finding white space (outside-in, includes non-customers) → Market Mapping. Re-cutting *our* customer base and picking focus segments → Customer Segmentation.
- **Market Mapping vs. Profit Pool Analysis** — Where demand is → Market Mapping. Where profit sits along the value chain → Profit Pool.
- **Strategic Options vs. Initiative Prioritizer vs. Portfolio Review** — Choosing *one direction* among alternatives → Strategic Options. Cutting a *long list of projects* to what fits capacity → Initiative Prioritizer. Reallocating capital across *existing units/bets* → Portfolio Review.
- **Business Case Builder vs. Value Realization** — Before approval (will it pay off?) → Business Case. After approval (is it paying off?) → Value Realization.
- **Initiative Prioritizer vs. Transformation Roadmap** — Deciding *what* to do → Prioritizer. Deciding *when and by whom* for already-chosen initiatives → Roadmap.
- **Operating Model Design vs. Transformation Roadmap** — *Structure and decision rights* to run the strategy → Operating Model. *Phased schedule* of initiatives → Roadmap.
- **Narrative Builder vs. Decision Memo vs. Stakeholder Alignment** — A presentation/storyline → Narrative Builder. A one-page written ask → Decision Memo. Getting the right people on side *before* the meeting → Stakeholder Alignment.
- **KPI Architect vs. Value Realization** — Designing the ongoing *metrics system* → KPI Architect. Tracking the specific *benefits a business case promised* → Value Realization.

## Common chains (propose, then confirm)

- **Full strategy engagement:** Situation Assessment → Strategic Options → Initiative Prioritizer → Transformation Roadmap → Narrative Builder
- **Growth problem:** Growth Barriers → Customer Segmentation or Pricing Strategy (whichever layer the constraint sits in) → KPI Architect
- **Market entry:** Market Mapping → Competitive Intel → Business Case Builder → War Gaming
- **Investment approval:** Assumption Audit → Business Case Builder → Risk & Mitigation → Decision Memo
- **Post-approval delivery:** Operating Model Design → Transformation Roadmap → KPI Architect → Value Realization
- **Getting a decision through:** Narrative Builder → Stakeholder Alignment → Decision Memo

## Boundaries with other skills in this repo

- **mckinsey-consultant** — Use it when the user wants an end-to-end McKinsey-style report or PPT generated from a business question. Use this skill when the user needs one specific strategy method done well, or a short chain of them.
- **pricing** (marketing pack) — SaaS price pages, tiers and packaging copy → `pricing`. Diagnosing *why* pricing is failing (value vs. structure vs. discount discipline) → Pricing Strategy here.
- **competitor-profiling / competitors** — Profiles from URLs or comparison pages → those skills. Predicting rival moves and pre-committing responses → Competitive Intel here.
- Marketing execution (ads, CRO, emails, SEO, etc.) is out of scope; route to the marketing skills.

## Original pack

`references/ORIGINAL-README.md` is the pack's original README. Each `references/*.md` file keeps its original frontmatter, which describes when that method applies.
