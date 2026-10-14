# Phase 1 — Data Validation, Funnel Analysis & Preliminary Direction

Source: `data/rippling_nb_dataset.csv` (fictional dataset in the case). Scripts: `analysis/phase1_explore.py`, `analysis/phase1_extra.py`.
Legend: **[F]** observed fact from the data · **[A]** assumption · **[H]** hypothesis to test.

## 1. Dataset validation
| Check | Result |
|---|---|
| Shape | 60 rows = 12 months × 5 employee-size segments [F] |
| Period | Feb-2021 to Jan-2022, one record per segment per month [F] |
| Granularity | All rows are **monthly**. The PDF says "Month or Year", but there are no annual rows, so nothing needs reconciling [F] |
| Missing / duplicates | 0 nulls, 0 duplicate month×segment keys [F] |
| Funnel integrity | Leads ≥ Demos ≥ Qualified Opps ≥ Logos in all 60 rows. No zero-logo rows [F] |
| Formatting | ARR is stored as a currency string ("$43,406") and was converted to a number [F] |
| Anomaly | In Nov–Dec 2021, leads fell 47% vs Oct and every conversion rate also fell. Jan-22 then hit a record. Holiday seasonality is likely, but one year of data can't confirm it [F/H] |
| Funnel definition | Demos > Qualified Opps, so the demo happens *before* the AE accepts the opp (demo-then-qualify motion) [F] |
| Cohort caveat | Each rate divides two counts from the same month. These are not true cohort conversions (lag is unknown) [A] |

## 2. Funnel (12-month totals, aggregate numerator / aggregate denominator)
| | Leads | Demos | Qual Opps | Logos | ARR | L→D | D→Q | Q→W | L→W | ARR/logo | ARR/lead | ARR share |
|---|---|---|---|---|---|---|---|---|---|---|---|---|
| 1-10 | 1,180 | 437 | 351 | 118 | $0.35M | 37.0% | 80.3% | 33.6% | 10.0% | $2,987 | $299 | 3.5% |
| 11-50 | 2,529 | 936 | 745 | 253 | $1.02M | 37.0% | 79.6% | 34.0% | 10.0% | $4,014 | $402 | 10.1% |
| 51-100 | 3,246 | 1,163 | 926 | 312 | $2.91M | 35.8% | 79.6% | 33.7% | 9.6% | $9,313 | **$895** | 28.9% |
| 101-500 | 6,962 | 1,807 | 1,264 | 321 | $4.00M | 26.0% | 70.0% | 25.4% | 4.6% | $12,461 | $575 | 39.8% |
| 501+ | 3,235 | 608 | 396 | 78 | $1.77M | 18.8% | 65.1% | 19.7% | 2.4% | $22,725 | $548 | 17.6% |
| **Total** | **17,152** | **4,951** | **3,682** | **1,082** | **$10.05M** | **28.9%** | **74.4%** | **29.4%** | **6.3%** | **$9,285** | **$586** | 100% |

Growth from Feb-21 to Jan-22 (11 months apart; this is not true YoY): leads +66%, demos +31%, logos +33%, ARR +86%. Monthly new ARR went from $0.69M to $1.28M.
By segment (ARR): 1-10 −29%, 11-50 +22%, 51-100 +54%, 101-500 +85%, 501+ +429%.

Driver decomposition (LMDI, segment-level, Feb–Apr-21 vs Aug–Oct-21 so the holiday months don't distort it): ARR +$914K = **volume +$889K**, deal size +$268K, **conversion −$244K**.

## 3. Top five findings
1. **Growth comes from lead volume, and conversion is getting worse.** [F] From Feb to Oct, leads grew 63% but demos grew only 29%. Blended L→D fell from 35.5% to 28.1% (Feb→Jan). About 1.7 pts of that is mix shift toward larger companies. The other ~5.6 pts is a decline *within* segments. Post-demo rates (D→Q, Q→W) stayed flat. So the leak is at the **lead → demo** step, which is SDR territory.
2. **The leak looks like a capacity problem [H].** In 11-50, 51-100 and 101-500, L→D is strongly *negatively* correlated with monthly lead volume (r = −0.77, −0.91, −0.73, excluding Nov–Dec). The more leads came in, the smaller the share that got a demo, which is what a capacity ceiling looks like. If H1 L→D had held in the four non-holiday H2 months: ≈155 more demos and ≈$0.32M ARR (≈$1M annualized; a scenario, not a forecast). Alternative explanation: lead quality got worse as volume scaled.
3. **51-100 is the most efficient segment.** [F] Its conversion matches SMB (9.6% L→W) at 2.3× the deal size, giving the highest ARR per lead ($895, vs $575 for 101-500 and $548 for 501+). It also has the steepest volume-driven L→D erosion (42% → 34%). Speed-to-lead coverage here is the best-return investment.
4. **Growth is moving to larger companies, but complexity rises sharply above 100 employees.** [F] 101-500 and 501+ produce 57% of ARR. Their share of leads went from 51% to 62%, and their ARR per logo rose 32% and 44% (H1→H2 basis). L→W halves at the 100-employee line (9.6% → 4.6%) and halves again at 500 (2.4%). D→Q and Q→W both fall, so the issue is buyer complexity, not lead handling. That calls for SC coverage and better qualification, not just more SDRs. 501+ L→D did *not* fall as volume tripled, so there the bottleneck is win rate, not throughput.
5. **1-10 is shrinking and has the lowest value.** [F] Leads −24%, ARR −29%, $3.0K ARR per logo, 3.5% of ARR. It converts the same as 11-50, so it isn't broken, just small. It should move to a low-touch / digital-assist motion, not get more headcount.

Supporting fact: there is a Nov–Dec trough (−47% leads, L→W 3.3–4.1%). Staff to the non-holiday run-rate and use Q4 for hiring and ramp.

## 4. Proposed segmentation logic (to finalize in Phase 3)
Boundaries can only fall on the dataset's bucket edges (10 / 50 / 100 / 500).
- **Evidence for a break at 50:** ARR per logo jumps 2.3× ($4.0K → $9.3K) while conversion stays the same, so the AE work is the same but each deal is worth more.
- **Evidence for a break at 100:** conversion halves and D→Q / Q→W drop, i.e., sales complexity changes.
- **Evidence for a break at 500:** ARR per logo is 1.8× higher, conversion halves again, and it takes 41 leads and 7.8 demos per logo.

| Option | Boundaries | Verdict |
|---|---|---|
| A. Conventional 3-tier | SMB 1-100 / MM 101-500 / ENT 501+ | Simple, and conversion is uniform within each tier. But it mixes $3K and $9K deals in one AE queue, so low-value deals take AE time away from 51-100 |
| B. Conventional alt | SMB 1-50 / MM 51-500 / ENT 501+ | Splits on deal size, but puts two different conversion profiles in the same MM pod |
| **C. Recommended 4-tier** | **SMB Velocity 1-50 · Emerging MM 51-100 · Core MM 101-500 · Enterprise 501+** | Every boundary sits on an observed break in deal size or conversion. 1-10 is routed to a digital-assist sub-motion within Velocity |

## 5. Preliminary staffing assumptions (ALL illustrative, to be confirmed)
| Assumption | Velocity 1-50 | Emerging MM 51-100 | Core MM 101-500 | ENT 501+ |
|---|---|---|---|---|
| SDR leads worked / ramped month | 300 | 200 | 175 | 125 |
| AE demos held / ramped month | 45 | 30 | 20 | 12 |
| SC involvement (% of demos) | 5% | 15% | 50% | 100% |
| Ramp (months; 50% productivity during ramp) | SDR 2 / AE 3 | AE 3 | AE 4 | AE 6 |
| Target utilization | 85% | 85% | 85% | 85% |
| Fully loaded cost (SF; OTE × 1.25) | SDR $113K · AE $200K | AE $240K | AE $275K | AE $375K · SC $263K |
| Quota attainment (cross-check) | 80% | 80% | 80% | 80% |
| Working days / yr | 250 | 250 | 250 | 250 |

Demand basis (proposed): Jan-22 run-rate by segment × 12, with a toggle for growth based on the observed lead trend. Preliminary result is about 13 SDRs, 30 AEs and 11-13 SCs. In Enterprise, ARR per AE (~$0.4M) barely covers fully loaded AE + SC cost, which supports a test-before-scale approach there.

## 6. Material data limitations
- **No current headcount or rep-level data.** All capacity figures are assumption-driven. We can't measure actual productivity or prove the capacity hypothesis.
- Same-period ratios, not cohorts. No sales-cycle length or stage timestamps.
- One year only, so seasonality and YoY growth can't be confirmed. The window is Feb–Jan, not a calendar year.
- 101-500 is a very wide bucket (40% of leads). It probably hides a 101-250 vs 251-500 break.
- No lead source / channel, no outbound vs inbound split, no speed-to-lead or response-time data.
- No closed-lost reasons, competitors, discounting, product mix, retention/NRR, CAC or marketing spend.
- ARR is new-logo first-year only. No LTV view, which matters most for the cheap SMB logos.
