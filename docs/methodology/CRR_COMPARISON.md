# 🔍 Comparison with the CRR

This guide compares each rule of the [calculation contracts](CALCULATION_CONTRACTS.md)
`CALC-02` to `CALC-08`, which implement the Basel simplified standardised
approach (`SRC-02`, MAR40), with the simplified standardised approach of
Article 325(2) of Regulation (EU) No 575/2013 (the CRR) and the articles it
points to, so that each difference has its own implementation task and no rule
changes on a reading of the CRR alone. It follows the decisions on the scope
of the CRR regime ([#74](https://github.com/danielep71/K-FRTB-SSA/issues/74)).

It is a reading of the text, made on 2026-11-20. It validates nothing: no
rule of the CRR regime is implemented, tested or confirmed by a competent
authority, and the traceability statuses of the
[validation plan](VALIDATION_PLAN.md#traceability) are unchanged.

## How to read it

The CRR is cited by article and paragraph of `SRC-04`, the consolidated text of
2026-01-01. Where CRR3 or the delegated act amended the article, `SRC-05` or
`SRC-06` holds the authentic text ([sources](SOURCES.md#register)). The MAR40
paragraph of a rule is the one its contract cites.

| Result | Meaning |
| --- | --- |
| **Same** | The CRR says the same as the contract's rule, as far as the text read shows |
| **Different** | The CRR's rule is not the contract's, and the difference is stated |
| **CRR only** | The CRR has a rule that no contract has |
| **Not in the CRR** | A contract rule that the text read does not have |
| **Open** | The text read does not settle the point; the task named in *Follow-up* does |

Each **Different**, **CRR only** and **Open** row names the issue that takes
it. Where a rule rests on a technical standard that is not adopted or not read,
the row says so, and the rule is an assumption
([the standards](SOURCES.md#technical-standards-that-the-crr-leaves-to-the-eba)).

## Requirement and scaling, `CALC-08`

| Step | The contract's rule | The CRR | Result | Follow-up |
| --- | --- | --- | --- | --- |
| 1 | The requirement of a class is the sum of its charges and its option charge | Article 326: the requirement for position risk is the sum of the requirements for the general and specific risk of debt and equity instruments | **Same** for general and specific risk | |
| 1 | The option charge is added before scaling | Article 329(2) and Article 352(5): other risks of options are reflected in the own funds requirements, with no word on the multiplier | **Open** | [#83](https://github.com/danielep71/K-FRTB-SSA/issues/83) |
| 2 | Factors 1.30, 3.50, 1.90 and 1.20 | Article 325(2), points (a)(i), (a)(ii), (c) and (b) | **Same**: the words of the points were compared with `SRC-05` by a program in [#75](https://github.com/danielep71/K-FRTB-SSA/issues/75), and the factors with the parameters | |
| 2 | Securitisation positions are outside the data contract | Article 325(2)(d): the requirement for securitisation instruments of Article 337 is added without scaling | **CRR only**: outside the scope by decision | [#74](https://github.com/danielep71/K-FRTB-SSA/issues/74) |
| 3 | The requirement is the simple sum of the scaled requirements | Article 325(2): *the sum of the following own funds requirements* | **Same** | |
| 4 | Risk-weighted assets are the requirement multiplied by 12.5 | Article 92(6), point (b): the own funds requirements of Article 92(4), points (b) to (f), multiplied by 12.5 | **Same** | |
| | No transitional multiplier | Article 495s(2), of `SRC-06`: until 2029-12-31, a multiplier of 0.9 on the result of the simplified standardised approach; Article 495v does not apply to an institution that applies only Articles 326 to 361 | **CRR only** | [#77](https://github.com/danielep71/K-FRTB-SSA/issues/77) |
| | The workbook does not decide whether the approach may be used | Article 325a(1) and (2): at most 10% of total assets and EUR 500 million, assessed monthly; Article 94: 5% and EUR 50 million | **CRR only**: decided to be calculated in the workbook | [#84](https://github.com/danielep71/K-FRTB-SSA/issues/84) |
| | The requirement is that of the positions entered, with no trading book boundary | Article 325(1): all trading book positions and all non-trading book positions subject to foreign exchange or commodity risk | **CRR only** | [#85](https://github.com/danielep71/K-FRTB-SSA/issues/85) |

## Interest rate risk, specific risk, `CALC-02`

| Step | The contract's rule | The CRR | Result | Follow-up |
| --- | --- | --- | --- | --- |
| 1 | A position of category `NONE` carries no specific risk: the legs of swaps, forward rate agreements, forwards and futures, which the user enters | Articles 328 to 330 treat these derivatives as combinations of long and short positions. The text read does not say in terms that a leg carries no specific risk | **Open** | [#79](https://github.com/danielep71/K-FRTB-SSA/issues/79) |
| 2 | Positions in the identical issue are netted, and no other offsetting is made | Article 327(1): the excess of long over short positions in the same issue, and in identical futures, options and warrants, is the net position of the instrument | **Same** | |
| 2 | The own debt instruments of the institution are not mentioned | Article 327(1): holdings of own debt instruments are disregarded in calculating specific risk under Article 336; Article 327(2): no netting between a convertible and its underlying | **CRR only** | [#79](https://github.com/danielep71/K-FRTB-SSA/issues/79) |
| 3 | The rate is that of table 1 of MAR40.6, by category, rating grade and residual term | Article 336(1), table 1: 0% for debt securities that would receive a 0% risk weight under the standardised approach for credit risk; 0.25%, 1.00% and 1.60% by residual term (up to six months, up to 24 months, above) for those that would receive 20% or 50% and for the other qualifying items of paragraph 4; 8% for 100%; 12% for 150% | **Different**: the CRR keys the category by the risk weight of the credit risk standardised approach, where the contract keys it by category (government, qualifying, other) and rating grade | [#79](https://github.com/danielep71/K-FRTB-SSA/issues/79) |
| 3 | | Article 114(2), table 1: a central government's credit quality steps 1 to 6 weigh 0%, 20%, 50%, 100%, 100% and 150%. Article 122(1), table 1: a corporate's steps 1 to 6 weigh 20%, 50%, 75%, 100%, 150% and 150%, and an unrated one 100% | **Open**: a corporate at step 3 weighs 75%, which none of the four categories of Article 336(1), table 1, names; the text read does not say which applies | [#79](https://github.com/danielep71/K-FRTB-SSA/issues/79) |
| 3 | | Article 336(2): for an institution on the IRB approach, the issuer needs an internal rating with a probability of default at most that of the matching credit quality step | **CRR only** | [#79](https://github.com/danielep71/K-FRTB-SSA/issues/79) |
| 3 | | Article 336(3): half of the second category's requirement for the bonds that qualify for 10% under Article 129(4) to (6) | **CRR only** | [#79](https://github.com/danielep71/K-FRTB-SSA/issues/79) |
| 3 | The user classifies each position; the workbook does not apply MAR40.7 to MAR40.12 | Article 336(4): the other qualifying items depend on the institution's own assessment of liquidity and investment quality and a documented methodology notified to the competent authority | **Same** in that the workbook applies no criteria; the notification is the institution's | |
| 4 | The charge of an issue is the absolute net value multiplied by the rate | Article 336(1): the weighted net positions are summed regardless of their signs | **Same** | |
| | The cap at the maximum possible loss is not supported (MAR40.15) | Article 335: the institution may cap the requirement of a net position at the maximum possible default-risk loss | **Same**: optional in the CRR, and not supported | [#79](https://github.com/danielep71/K-FRTB-SSA/issues/79) |
| | Securitisation (MAR40.14) and the correlation trading portfolio (MAR40.22) are not supported | Articles 337 and 338, as amended by CRR3 | **Same**: not supported, by decision | [#74](https://github.com/danielep71/K-FRTB-SSA/issues/74) |
| | Credit derivatives are not supported (MAR40.16 to MAR40.21) | Articles 332, 346 and 347 | **Same**: not supported; the articles were read by title | |

## Interest rate risk, general risk, `CALC-03`

| Step | The contract's rule | The CRR | Result | Follow-up |
| --- | --- | --- | --- | --- |
| 1 | Each currency has its own ladder, and the requirements are summed | Article 334: net positions are classified by currency, and the requirement is calculated for each currency separately | **Same** | |
| 2 | A fixed-rate position is slotted by residual term to final maturity, a floating-rate one by the term to the next repricing | Article 339(2) | **Same** | |
| 3 | The bands, limits and risk weights of table 4 of MAR40.26, with a coupon threshold of 3% | Article 339(2) and (4), table 2 | **Same**: the 15 bands, their limits for the two coupon columns and their weights were compared with the parameters of `src/core/CORE_Config.bas` by a program | |
| 4 | The weighted position is the signed value multiplied by the weight | Article 339(2) | **Same** | |
| 5 | The vertical disallowance is 10% of the smaller side of a band | Article 339(9), point (a) | **Same** | |
| 6 | The horizontal disallowance is 40% of the matched amount in zone 1 and 30% in zones 2 and 3 | Article 339(9), points (b) to (d) | **Same** | |
| 7 | Zones 1 and 2, then 2 and 3, at 40%, then what is left of zones 1 and 3 at 100% | Article 339(9), point (e): 40% between zones 1 and 2 and between zones 2 and 3; point (f): **150%** of the matched weighted position between zones 1 and 3 | **Different**: the rate between zones 1 and 3 is 150%, where `HorizontalZone13` holds 1.00 | [#80](https://github.com/danielep71/K-FRTB-SSA/issues/80) |
| 7 | The order is an assumption of the workbook: MAR40.28 does not state it | Article 339(5) to (7): zones 1 and 2 first, then 2 and 3, then 1 and 3; paragraph 6 lets the institution reverse the first two | **Same** order as the workbook's, with an option the workbook lacks: the CRR settles the assumption for its regime | [#80](https://github.com/danielep71/K-FRTB-SSA/issues/80) |
| 8, 9 | The net position charge is the absolute value of what remains, and the requirement sums steps 5 to 8 | Article 339(9), point (g): 100% of the residual unmatched weighted positions | **Same** | |
| | The duration method is not supported (MAR40.29) | Article 340: an institution may use it instead of Article 339, on a consistent basis | **CRR only** as a choice of the institution | [#80](https://github.com/danielep71/K-FRTB-SSA/issues/80) |
| | A single ladder for currencies of insignificant business is not supported (MAR40.24 and MAR40.30) | Article 334 has each currency separately; the text read has no such exception | **Not in the CRR** | |
| | The user converts derivatives into legs; matched positions are left out of the input | Articles 328 to 330; Article 331(2): positions that meet its conditions on value, currency, rate and date may be treated as fully offsetting; Article 331(1): sensitivity models, with permission | **Same** in that the workbook converts nothing; **CRR only** for the models | |

## Equity risk, `CALC-04`

| Step | The contract's rule | The CRR | Result | Follow-up |
| --- | --- | --- | --- | --- |
| 1 | Positions in the identical equity or index in a market are netted | Article 327(1) and Article 341(1) | **Same** | |
| 2 | Specific risk is 8% of the absolute net position of each single name | Article 341(1) and Article 342: 8% of the overall gross position, the sum of the absolute values of the sum of the net long and of the net short positions | **Same** | |
| 3 | The charge for an index is 2% of its absolute net position, in place of specific risk | Article 344(3) and (4): a stock-index future may be broken down into its equities, with a notification to the competent authority, or treated as one equity, whose specific risk may be ignored where the future is exchange traded and represents a relevant, appropriately diversified index | **Different**: the CRR has no 2% charge. An index treated as one equity carries the general risk, and the specific risk unless the future is exchange traded and represents a relevant, appropriately diversified index, where the CRR lets it be ignored | [#81](https://github.com/danielep71/K-FRTB-SSA/issues/81) |
| 4 | General risk is 8% of the absolute net position of each market | Article 341(2) and Article 343: 8% of the overall net position, calculated for each market | **Same** in the rate and the calculation | |
| 4 | A market is the national market entered by the user | Article 341(3): the definition of *market* is left to technical standards of the EBA, `SRC-07`, whose Official Journal text was not read | **Open**: an assumption until `SRC-07` is verified | [#81](https://github.com/danielep71/K-FRTB-SSA/issues/81), [#91](https://github.com/danielep71/K-FRTB-SSA/issues/91) |
| 5 | The requirement is the sum of steps 2 to 4 | Article 326 | **Same** | |
| | The arbitrage strategies (MAR40.48 to MAR40.50) and depository receipts (MAR40.51) are not supported | Articles 345 to 347, on underwriting, were read by title | **Same**: not supported; not compared | |
| | No treatment of collective investment undertakings | Article 348(1), as amended by CRR3: 32% for position risk, and 40% with foreign exchange risk | **CRR only** | [#81](https://github.com/danielep71/K-FRTB-SSA/issues/81) |

## Commodities risk, `CALC-05`

| Step | The contract's rule | The CRR | Result | Follow-up |
| --- | --- | --- | --- | --- |
| 1 | A position is valued at the spot price in the reporting currency | Article 357(1) and Article 360(1) | **Same** | |
| 2 | Long and short positions in a commodity are netted, by the commodity's name; different commodities are not offset | Article 357(3): the net position in each commodity and in identical futures, options and warrants | **Same** for the commodity | |
| 2 | Netting between sub-categories is a national discretion of MAR40.67, and not supported | Article 357(4): positions in sub-categories that are deliverable against each other, and in similar commodities that are close substitutes with a correlation of at least 0.9 over at least one year, *shall be treated as positions in the same commodity* | **Different**: a rule of the CRR, not a discretion | [#82](https://github.com/danielep71/K-FRTB-SSA/issues/82) |
| 3 | The directional charge is 15% of the absolute net position | Article 360(1), point (a) | **Same** | |
| 4 | A further charge is 3% of the gross position | Article 360(1), point (b) | **Same** | |
| 5 | The requirement is the sum over the commodities | Article 360(2) | **Same** | |
| | The maturity ladder approach is not supported (MAR40.68 to MAR40.70) | Article 355: the institution chooses among Articles 359, 360 and 361. Article 359: a spread rate of 1.50% for each of seven bands, a carry rate of 0.6% and an outright rate of 15% | **CRR only** as a choice of the institution | [#82](https://github.com/danielep71/K-FRTB-SSA/issues/82) |
| | | Article 361: the extended maturity ladder, with the rates of its own table, for an institution with significant commodities business and an appropriately diversified portfolio, with a notification to the competent authority | **CRR only** | [#82](https://github.com/danielep71/K-FRTB-SSA/issues/82) |
| | | Article 356: ancillary agricultural commodities business | **CRR only** | [#82](https://github.com/danielep71/K-FRTB-SSA/issues/82) |
| | | Article 357(5): commodities transferred in a repurchase agreement or lent are included | **CRR only** | [#82](https://github.com/danielep71/K-FRTB-SSA/issues/82) |
| | A derivative is entered as a notional position | Article 358(1) and (2): futures, forwards and swaps as notional amounts with a maturity | **Same** in that the workbook converts nothing | |
| | Gold is FX risk and rejected as a commodity | Article 357(2) | **Same** | |

## Foreign exchange risk, `CALC-06`

| Step | The contract's rule | The CRR | Result | Follow-up |
| --- | --- | --- | --- | --- |
| 1 | The net position in each currency and in gold is the sum of its rows at the spot rate; the components are summed outside the workbook (MAR40.55 to MAR40.58) | Article 352(1): spot, forward, guarantees, the delta-based equivalent of options and the market value of other options; Article 352(3): net present value, if used consistently | **Same** | |
| 2 | The overall net open position is the greater of the sums of net long and net short positions, plus the net position in gold whatever its sign | Article 352(4): the net positions in each currency other than the reporting currency and the net long or short position in gold are converted and summed separately, the higher total being the overall net position. Article 351: the sum of the overall net foreign exchange position and the net gold position | **Open**: Articles 352(4) and 351 read together can be read with gold inside the totals or added to them. The text read does not settle it | [#78](https://github.com/danielep71/K-FRTB-SSA/issues/78) |
| 3 | The requirement is 8% of the overall net open position | Article 351: 8% of that sum | **Same** in the rate | |
| 3 | The exemption of MAR40.62 is not supported | Article 351: no requirement unless the sum exceeds 2% of total own funds | **CRR only** as a threshold of total own funds | [#78](https://github.com/danielep71/K-FRTB-SSA/issues/78) |
| | | Article 354: lower requirements for closely correlated currencies, 4% of the matched position, and 1.6% for the matched positions in the currencies of Member States participating in the second stage of the economic and monetary union; the list of currencies is left to technical standards | **CRR only** | [#78](https://github.com/danielep71/K-FRTB-SSA/issues/78) |
| | | Article 104c: with permission, the exclusion of a position that hedges a capital ratio; Article 325(1): no requirement for positions deducted from own funds | **CRR only** | [#78](https://github.com/danielep71/K-FRTB-SSA/issues/78) |
| | A row in the reporting currency carries no FX risk | Article 352(4): the currencies other than the reporting currency | **Same** | |

## Options, `CALC-07`

| Step | The contract's rule | The CRR | Result | Follow-up |
| --- | --- | --- | --- | --- |
| 1 | The user enters the charge of each class, by a method of MAR40.74 to MAR40.85 | Article 329(1): an option is treated as a position of its underlying multiplied by its delta, the delta being the exchange's or, with permission, the institution's own; Articles 329(2) and (3), 352(5) and (6), 358(3) and (4): other risks are reflected by methods that technical standards define, `SRC-08`, whose Official Journal text was not read | **Open**: the CRR gives no method of its own, and the standard is unverified, so the treatment is an assumption | [#83](https://github.com/danielep71/K-FRTB-SSA/issues/83), [#91](https://github.com/danielep71/K-FRTB-SSA/issues/91) |
| 2 | The charge is added before scaling | See `CALC-08`, step 1 | **Open** | [#83](https://github.com/danielep71/K-FRTB-SSA/issues/83) |

## What the CRR holds that no contract has

| Subject | The CRR | Follow-up |
| --- | --- | --- |
| Securitisation positions and the correlation trading portfolio | Articles 337 and 338 | Outside the scope by decision ([#74](https://github.com/danielep71/K-FRTB-SSA/issues/74)) |
| Underwriting | Articles 345 to 347, read by title | Outside the scope: no issue |
| Collective investment undertakings | Articles 348 to 350 and 353; the rates of Article 348(1) are in `CALC-04` above | [#81](https://github.com/danielep71/K-FRTB-SSA/issues/81) |
| Size conditions and the derogation for small trading book business | Articles 325a and 94 | [#84](https://github.com/danielep71/K-FRTB-SSA/issues/84) |
| The transitional multiplier and the derogation of Article 495u | `SRC-06` | [#77](https://github.com/danielep71/K-FRTB-SSA/issues/77), [#85](https://github.com/danielep71/K-FRTB-SSA/issues/85) |
| Non-trading book positions subject to foreign exchange or commodity risk | Articles 325(1) and 325(9) | [#85](https://github.com/danielep71/K-FRTB-SSA/issues/85) |

## What was not compared

The articles read in full, or in the parts a row cites, are those of the
*Verification record* of `SRC-04`. Read only in part: Articles 328, 329 and 332. Read by title
only: Articles 333, 345 to 350 and 353, and the table of Article 361. The credit risk articles of Title II, which Article 336 points to, were
read only for the tables cited above, and the credit quality steps that map a
rating to a step were not read. The comparison does not decide a case that the
text read leaves open: each **Open** row says what a decision needs.
