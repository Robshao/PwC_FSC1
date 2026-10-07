# First-cut identities by problem type

An identity is MECE by construction, so start there. Put segmentation (channel, region, product, customer, device) as cuts across every branch, not as branches.

## General

| Problem | First cut | Next level |
|---|---|---|
| Profit down | Revenue − cost | Revenue = volume × price (× mix); cost = fixed + variable, or volume × unit cost |
| Revenue down (multi-product) | Price effect + volume effect + mix effect | Price-volume-mix bridge |
| Any ratio moved | Numerator vs denominator | Which one moved, then decompose it |
| Growth slowed | Volume vs value; base effect | Check comparison period before anything else |
| Cost up | Volume × unit price | Driver volumes; rate changes; fixed vs variable with a rule for semi-variable |
| Cycle time up | Sum of time at each step (waiting + working) | Steps with largest increase; rework loops; handoffs |
| Market share down | Our sales ÷ market size | Did the market grow while we did not, or did we shrink? |

## Retail and consumer goods

- Sell-in revenue = (consumer sell-out volume + change in retailer inventory) × net price per unit.
- Sell-out volume = category volume × our share; share = availability (distribution, out-of-stock) + demand generation (marketing) + shelf competitiveness (relative price, promotion, product, competitor launches).
- Net price = list price − promotions and trade spend, adjusted for mix.
- Classic traps: channel destocking after loading, quarter-end incentives on shipments, relative price gap after cost-driven increases.

## E-commerce

- Revenue = visitors × conversion rate × average order value.
- Conversion = funnel steps (view → add to cart → checkout → payment), cut by device.
- Average order value = items per order × price per item, including mix and discounts.
- Customer acquisition cost is a ratio: total acquisition spend ÷ new customers. Decompose both, and keep fulfilment costs (shipping, payment fees) out of acquisition cost.

## Subscription and SaaS

- Ending customers = starting customers + new − churned; revenue = customers × average revenue per customer.
- Average revenue per customer = package mix × price level − discounts.
- Keep counts and rates on different levels (churned customers on one level, churn rate below it).

## Banking

- Net interest income = average balance × net interest margin.
- Fee income = transactions × fee per transaction.
- Credit cost = exposure × probability of default × loss given default.

## Life insurance

- Profit (conceptually) = insurance service result + investment result (+ other). Confirm which profit line and which accounting basis the client means; under IFRS 17 (International Financial Reporting Standard 17) insurance revenue is not the same as premium received.
- New business premium = number of policies × average premium; policies = sum across channels (agents, bancassurance, brokers, online, telemarketing).
- Policy not continued = surrender + lapse (non-payment) + free-look cancellation; lapse = payment channel issues + customer affordability + customer unaware.
- New-business funnel: contact → needs analysis → proposal → application → underwriting (company-side conversion: decline, postpone, rated) → issuance.
- Acquisition cost per policy = acquisition spend ÷ new policies; spend by channel or by nature (commission, bank fees, marketing, medical exams), never both on one level.
- Observation windows matter: 13th and 25th month persistency.

## Non-life (property and casualty) insurance

- Underwriting result: combined ratio = loss ratio + expense ratio.
- Loss ratio = claim frequency × average claim severity ÷ average earned premium per policy. The denominator moves too (price competition, discounts).
- Separate current accident year from prior-year reserve development; separate gross from net of reinsurance; isolate catastrophes.
- Expense ratio = commission and acquisition cost + operating expenses (mostly fixed) over premium.
- Renewal rate: claims experience, renewal premium change (experience rating), renewal notification process, channel service, non-price competition, customer-side changes (sold the car). Claims experience and premium increase are confounded because both follow a claim.

## Operations and service

- Throughput = capacity × utilization × first-time-right rate.
- Backlog change = inflow − outflow.
- Complaint volume = transactions × complaint rate; split by cause code, and check whether cause codes exist before promising the analysis.
