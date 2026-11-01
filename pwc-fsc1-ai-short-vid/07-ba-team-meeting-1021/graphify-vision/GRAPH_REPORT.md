# Graph Report - corpus  (2026-11-01)

## Corpus Check
- Corpus is ~1,067 words - fits in a single context window. You may not need a graph.

## Summary
- 80 nodes · 120 edges · 8 communities
- Extraction: 77% EXTRACTED · 22% INFERRED · 1% AMBIGUOUS · INFERRED: 27 edges (avg confidence: 0.79)
- Token cost: 0 input · 78,769 output

## Community Hubs (Navigation)
- AI 信任與治理
- 現場共創與原型
- BlackRock FDBA 與規則即程式
- 領域脈絡與 AI 代理
- Builder 心態
- 為什麼現在要 Build
- 今日 BA 與數位分身
- UWOS 營運模式

## God Nodes (most connected - your core abstractions)
1. `Vision: the business analyst in 2030` - 13 edges
2. `BlackRock Forward Deployed Business Analyst (job posting)` - 12 edges
3. `Trust and governance designer` - 8 edges
4. `Four-trait ideal profile` - 7 edges
5. `FSC six core AI principles` - 7 edges
6. `The business analyst today (before)` - 7 edges
7. `Why Now Is the Time to Build` - 6 edges
8. `Insurance associations AI self-regulatory rules (2025, amended 2026)` - 6 edges
9. `McKinsey Underwriting Operating System (UWOS)` - 6 edges
10. `Skills the future BA needs` - 6 edges

## Surprising Connections (you probably didn't know these)
- `Operating model shaper` --semantically_similar_to--> `Target operating model shaping`  [INFERRED] [semantically similar]
  05-ba-vision-2030.md → 02-blackrock-forward-deployed-ba.md
- `Vision and ambition as the new bottleneck` --conceptually_related_to--> `Builder-translator BA`  [AMBIGUOUS]
  01-why-now-builders.md → 05-ba-vision-2030.md
- `Responsibility to guide AI risk` --conceptually_related_to--> `Trust and governance designer`  [INFERRED]
  01-why-now-builders.md → 05-ba-vision-2030.md
- `Skills the future BA needs` --semantically_similar_to--> `Four-trait ideal profile`  [INFERRED] [semantically similar]
  05-ba-vision-2030.md → 02-blackrock-forward-deployed-ba.md
- `Context architect (domain knowledge as a product)` --semantically_similar_to--> `Structured business context`  [INFERRED] [semantically similar]
  05-ba-vision-2030.md → 02-blackrock-forward-deployed-ba.md

## Hyperedges (group relationships)
- **BA before-to-after transformation pairs** — 04_ba_today_requirements_handoff, 04_ba_today_manual_documentation, 04_ba_today_knowledge_in_individual_heads, 04_ba_today_late_testing, 04_ba_today_scope_creep, 04_ba_today_ad_hoc_untrusted_ai, 05_ba_vision_2030_co_creation_forward_deployed, 05_ba_vision_2030_orchestrator_of_ai_agent_teams, 05_ba_vision_2030_context_library, 05_ba_vision_2030_rules_as_code, 05_ba_vision_2030_business_digital_twin, 05_ba_vision_2030_trust_and_governance_designer [EXTRACTED 1.00]
- **Human-governed AI control regime (regulation + BA governance role)** — 03_regulation_and_trends_human_in_the_loop, 03_regulation_and_trends_insurance_ai_self_regulatory_rules, 03_regulation_and_trends_machine_first_human_governed, 05_ba_vision_2030_trust_and_governance_designer, 05_ba_vision_2030_human_go_no_go_decision, 02_blackrock_forward_deployed_ba_critical_evaluation_of_ai_outputs [INFERRED 0.85]
- **Structured domain context powering AI agents** — 02_blackrock_forward_deployed_ba_structured_business_context, 05_ba_vision_2030_context_architect, 05_ba_vision_2030_context_library, 03_regulation_and_trends_context_as_key_differentiator, 03_regulation_and_trends_domain_specific_language_models, 05_ba_vision_2030_context_engineering [INFERRED 0.85]

## Communities (8 total, 0 thin omitted)

### Community 0 - "AI 信任與治理"
Cohesion: 0.15
Nodes (15): AI inventory, Fairness and human-centric values, Governance and accountability, Insurance associations AI self-regulatory rules (2025, amended 2026), Internal validation for AI in underwriting, claims, customer service, Privacy and customer rights, Regulation and technology trends shaping AI in insurance, FSC six core AI principles (+7 more)

### Community 1 - "現場共創與原型"
Cohesion: 0.20
Nodes (12): Functional prototypes via AI-enabled development, AI-native development platforms, AI governance literacy, Vision: the business analyst in 2030, How before maps to after, Change and adoption coach, Co-creation in the room (forward deployed), Context engineering (+4 more)

### Community 2 - "BlackRock FDBA 與規則即程式"
Cohesion: 0.21
Nodes (11): Blend of business analysis, product discovery, business architecture and prototyping, BlackRock, BlackRock Transformation Office (BTO), Critical evaluation of AI outputs with traceability and human judgment, BlackRock Forward Deployed Business Analyst (job posting), Decompose processes into explicit logic, controls and exception paths, Target operating model shaping, Testable solution hypotheses with measurable outcomes (+3 more)

### Community 3 - "領域脈絡與 AI 代理"
Cohesion: 0.24
Nodes (10): Structured business context, Domain-specific language models, Gartner Top Strategic Technology Trends 2026, Multiagent systems, Knowledge lives in individual heads, Manual minutes and documents, Acceptance criteria, evals and confidence thresholds, Context architect (domain knowledge as a product) (+2 more)

### Community 4 - "Builder 心態"
Cohesion: 0.32
Nodes (6): Vibe coders / individual builders, Bias to action of a hands-on builder, Enterprise perspective of a business architect, Four-trait ideal profile, Prioritization instincts of a product owner, Builder-translator BA

### Community 5 - "為什麼現在要 Build"
Cohesion: 0.25
Nodes (6): Alexandr Wang, Garry Tan, New business and creative models, This is a Once-in-a-Civilization Opportunity (talk), Why Now Is the Time to Build, Y Combinator

### Community 6 - "今日 BA 與數位分身"
Cohesion: 0.40
Nodes (6): Structured thinking of a business analyst, The business analyst today (before), Interview and hand off requirement documents, Scope keeps growing; trade-offs invisible, Static mockups; clients see system only after development, Business digital twin

### Community 7 - "UWOS 營運模式"
Cohesion: 0.47
Nodes (5): Humans shift from processing to oversight and exception handling, McKinsey Underwriting Operating System (UWOS), Three-lane routing (straight-through, assisted decision, expert judgment), UWOS four steps (lighthouse case, minimum data foundation, define controls, targeted workflows), Operating model shaper

## Ambiguous Edges - Review These
- `Vision and ambition as the new bottleneck` → `Builder-translator BA`  [AMBIGUOUS]
  05-ba-vision-2030.md · relation: conceptually_related_to

## Knowledge Gaps
- **20 isolated node(s):** `Alexandr Wang`, `Garry Tan`, `Y Combinator`, `New business and creative models`, `Prioritization instincts of a product owner` (+15 more)
  These have ≤1 connection - possible missing edges or undocumented components. (Counts symbols only; 21 node(s) total have ≤1 connection when file, concept and rationale nodes are included.)

## Suggested Questions
_Questions this graph is uniquely positioned to answer:_

- **Why does `Vision: the business analyst in 2030` connect `現場共創與原型` to `AI 信任與治理`, `BlackRock FDBA 與規則即程式`, `領域脈絡與 AI 代理`, `Builder 心態`, `為什麼現在要 Build`, `今日 BA 與數位分身`, `UWOS 營運模式`?**
  _High betweenness centrality (0.520) - this node is a cross-community bridge._
- **Are the 2 inferred relationships involving `Vision: the business analyst in 2030` (e.g. with `Why Now Is the Time to Build` and `BlackRock Forward Deployed Business Analyst (job posting)`) actually correct?**
  _`Vision: the business analyst in 2030` has 2 INFERRED edges - model-reasoned connections that need verification._
- **What connects `Alexandr Wang`, `Garry Tan`, `Y Combinator` to the rest of the system?**
  _20 weakly-connected nodes found - possible documentation gaps or missing edges._
- **Why does `Trust and governance designer` connect `AI 信任與治理` to `現場共創與原型`, `BlackRock FDBA 與規則即程式`, `領域脈絡與 AI 代理`, `為什麼現在要 Build`?**
  _High betweenness centrality (0.264) - this node is a cross-community bridge._
- **Are the 2 inferred relationships involving `Trust and governance designer` (e.g. with `Responsibility to guide AI risk` and `Critical evaluation of AI outputs with traceability and human judgment`) actually correct?**
  _`Trust and governance designer` has 2 INFERRED edges - model-reasoned connections that need verification._
- **Why does `BlackRock Forward Deployed Business Analyst (job posting)` connect `BlackRock FDBA 與規則即程式` to `現場共創與原型`, `領域脈絡與 AI 代理`, `Builder 心態`?**
  _High betweenness centrality (0.198) - this node is a cross-community bridge._

### Low-confidence Hints
_AMBIGUOUS edges — the extractor was unsure. Verify before acting on these._

- **What is the exact relationship between `Vision and ambition as the new bottleneck` and `Builder-translator BA`?**
  _Edge tagged AMBIGUOUS (relation: conceptually_related_to) - confidence is low._