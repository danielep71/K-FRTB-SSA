# 📚 Sources

Every rule in the [calculation contracts](CALCULATION_CONTRACTS.md) cites an
entry of this register by its identifier and a locator, such as an article,
paragraph, table or template cell. A rule without a verified source is a
documented assumption, and its contract says so.

## Register

| ID | Title | Issuer | Version or date | Locator | Local copy | SHA-256 of the copy | Status | Verified by | Verified on |
| --- | --- | --- | --- | --- | --- | --- | --- | --- | --- |
| *Illustrative:* `SRC-01` | Example regulation | Example authority | 2026-01-01 | Official publication reference or URL | — | — | `unverified` | — | — |
| `SRC-02` | Basel Framework, chapter MAR40, *Simplified standardised approach* | Basel Committee on Banking Supervision, published by the Bank for International Settlements | In force from 2023-01-01; last update 2024-07-05; status *Current* when verified | <https://www.bis.org/committees/bcbs/basel-framework/standard/mar/40/inforce/2023-01-01/published/2024-07-05> | — | — | `verified` | Daniele Penza | 2026-11-16 |
| `SRC-03` | *Amendment to the capital accord to incorporate market risks*, updated November 2005, Part C.2, the worked example of the general market risk of interest rate instruments | Basel Committee on Banking Supervision, published by the Bank for International Settlements | Originally released 1996-01-04, modified 1997-09, updated 2005-11-14; published 2005-11-15; status *Superseded* by the Basel Framework | <https://www.bis.org/publ/bcbs119.pdf> (the publication page links the PDF of 63 pages, <https://www.bis.org/publications/200511-standards-amendment-capital-accord-incorporate-market-risks.pdf>) | — | — | `verified` | Daniele Penza | 2026-11-19 |
| `SRC-04` | Regulation (EU) No 575/2013 (the CRR), consolidated text: Part Three, Title IV, Articles 325 to 361, with Articles 94 and 104c | European Parliament and Council; consolidated text of the Publications Office of the European Union, on EUR-Lex | Consolidated version of 2026-01-01 (document `02013R0575-20260101`, version 020.001), which holds the amendments of Regulation (EU) 2024/1623 (M17), Delegated Regulation (EU) 2024/2795 (M18) and Delegated Regulation (EU) 2025/1496 (M21), and the corrigendum of 2025-04-16 to Regulation (EU) 2024/1623 (C8); the text states that it has no legal effect | <https://eur-lex.europa.eu/legal-content/EN/TXT/?uri=CELEX:02013R0575-20260101> | — | — | `verified` | Daniele Penza | 2026-11-20 |
| `SRC-05` | Regulation (EU) 2024/1623 of the European Parliament and of the Council of 31 May 2024 amending Regulation (EU) No 575/2013 as regards requirements for credit risk, credit valuation adjustment risk, operational risk, market risk and the output floor (CRR3) | European Parliament and Council, published in the Official Journal of the European Union | Official Journal L, 2024/1623, 2024-06-19 | <https://eur-lex.europa.eu/eli/reg/2024/1623/oj/eng> | — | — | `verified` | Daniele Penza | 2026-11-20 |
| `SRC-06` | Commission Delegated Regulation (EU) 2026/1221 of 4 June 2026 amending Regulation (EU) No 575/2013 as regards temporary targeted operational relief measures and targeted multipliers for the calculation of an institution's own funds requirements for market risk | European Commission, published in the Official Journal of the European Union | Official Journal L, 2026/1221, 2026-09-11; enters into force on the day following its publication and applies from 2027-01-01 | <https://data.europa.eu/eli/reg_del/2026/1221/oj> | — | — | `verified` | Daniele Penza | 2026-11-20 |
| `SRC-07` | Commission Delegated Regulation (EU) No 525/2014 of 12 March 2014, the regulatory technical standards on the definition of *market* for the overall net position in equity instruments (Article 341(3) of the CRR) | European Commission, published in the Official Journal of the European Union | 2014-03-12; its identification comes from the EBA's final draft and from secondary copies, not from the Official Journal | <https://eur-lex.europa.eu/eli/reg_del/2014/525/oj> | — | — | `unverified` | — | — |
| `SRC-08` | Commission Delegated Regulation (EU) No 528/2014 of 12 March 2014, the regulatory technical standards on the non-delta risk of options under the standardised approach to market risk (Articles 329(3), 352(6) and 358(4) of the CRR) | European Commission, published in the Official Journal of the European Union | 2014-03-12; its identification comes from the EBA's final draft and from secondary copies, not from the Official Journal | <https://eur-lex.europa.eu/eli/reg_del/2014/528/oj> | — | — | `unverified` | — | — |

`SRC-02` is the only source a rule of the Basel regime relies on. `SRC-03` is
registered for the printed values of one worked example, and no rule cites it:
its document is superseded by `SRC-02`, and the example is only a published
value that a case compares the maturity method with. The illustrative `SRC-01` stays with the illustrative test cases that cite it,
which the example suite of the test harness reads
([the project's suites](TEST_HARNESS.md#the-projects-suites)); no rule of
the workbook cites it.

`SRC-04` to `SRC-08` are the texts of the CRR regime of the milestone v0.3.0
([the decisions of #74](VALIDATION_PLAN.md#decisions)): the simplified
standardised approach of Article 325(2) of the CRR, with Chapters 2 to 4 of
Title IV of Part Three. No rule cites them yet: the
comparison of the rules with them is a task of its own
([#76](https://github.com/danielep71/K-FRTB-SSA/issues/76)), and a rule of the
CRR regime will cite them by article and paragraph. `SRC-04` is a consolidated
text, a documentation tool without legal effect: where CRR3 amended an
article, the rule is checked against the Official Journal text of the
amending act, `SRC-05` or `SRC-06`. Registering a source validates no rule: the
traceability statuses of the [validation plan](VALIDATION_PLAN.md#traceability)
stay as they are.

### Technical standards that the CRR leaves to the EBA

A rule of the CRR regime that depends on a technical standard cites the act,
or is an assumption where no act is registered. State on 2026-11-20:

| Article | Subject | Act | State |
| --- | --- | --- | --- |
| 341(3) | The definition of *market* for the overall net position in equity instruments | `SRC-07` | Identified, not read in the Official Journal: `unverified` |
| 329(3) and 352(6) | The other risks of options than delta risk | `SRC-08` | Identified, not read in the Official Journal: `unverified` |
| 94(10) and 325a(2) | The main risk driver of a position and whether it is long or short | None registered: the EBA's final report on the draft standards, of 2024-12-06, is known, and no adopted act was found | Assumption: the side of a position is the one entered |
| 325(9) | The calculation of the requirement for non-trading book positions subject to foreign exchange or commodity risk | None registered: no act adopted for the mandate that CRR3 renewed was found | Assumption |

- **ID:** `SRC-` followed by a number, never reused. An independent
  implementation used to compute reference values, such as another author's
  tool, is registered too, with an `IMPL-` ID
  ([test cases](TEST_CASES.md#reference-classes)).
- **Version or date:** the consolidated version, publication or application
  date that the workbook implements.
- **Locator:** where the issuer publishes it, precise enough to find the same
  version again.
- **Status:** `unverified`, `verified`, or `superseded by SRC-NN`.

## Local copies

A copy of a source, such as a regulation, a reporting template or the
issuer's own validation rules, is stored byte for byte under
`docs/methodology/sources/`. Git neither normalizes those files nor reports
a white-space error in them, and editors leave them unchanged. A `README.md`
in that directory, by convention, is ordinary documentation that Git and
editors treat as text; the register stays here.

- Name each copy `<ID>_<short-title>.<extension>`, for example
  `SRC-01_example-regulation.pdf`.
- Record its SHA-256 in the register (`sha256sum`, or `certutil -hashfile
  FILE SHA256` on Windows).
- Store a copy only if its terms allow redistribution: the copy goes with
  every clone of the repository.

## Terms of the sources

| ID | Terms, as the issuer states them | Read on | Local copy |
| --- | --- | --- | --- |
| `SRC-02` | The [terms and conditions](https://www.bis.org/about/terms-conditions) of the issuer's website allow a user to download, display, print, photocopy or redistribute its material for non-commercial purposes, and to reproduce a limited extract in other publications free of charge, with the issuer cited as the source: not more than 400 words of text or two tables or graphs, and in any case not more than 10% of the publication. Any other use needs its written permission | 2026-11-16 | None: see the [decisions](#decisions) |
| `SRC-03` | The same terms of the issuer's website as `SRC-02`: the document carries the notice *Brief excerpts may be reproduced or translated provided the source is stated* | 2026-11-19 | None: see the [decisions](#decisions) |
| `SRC-04` | The consolidated text states: *This text is meant purely as a documentation tool and has no legal effect. The Union's institutions do not assume any liability for its contents. The authentic versions of the relevant acts, including their preambles, are those published in the Official Journal of the European Union and available in EUR-Lex.* The terms under which EUR-Lex allows its content to be reused were not read: the legal notice returned no content to my fetch on 2026-11-20 | 2026-11-20 | None: see the [decisions](#decisions) |
| `SRC-05`, `SRC-06` | The Official Journal text is the authentic one. The terms under which EUR-Lex allows its content to be reused were not read, as for `SRC-04` | 2026-11-20 | None: see the [decisions](#decisions) |
| `SRC-07`, `SRC-08` | Not read: the Official Journal text of these acts was not retrieved | — | None |

The guides and the workbook cite `SRC-02` by paragraph and restate single
factors. They do not reproduce its tables: the values of its tables 1, 4 and
5 are parameters of the workbook, on the `Params` sheet, which the
[calculation contracts](CALCULATION_CONTRACTS.md#parameters) trace to their
paragraphs.

## Verifying an entry

1. Open the source from the issuer's own publication, not from a copy.
2. Confirm the version or date, and that it applies on the reporting dates
   the workbook supports.
3. If there is a local copy, recompute its SHA-256 and compare it with the
   register.
4. Check each locator that a calculation contract cites against the text.
5. Set the status to `verified`, with your name and the date.

When a source is replaced, keep its row, set it to `superseded by` the new
entry, and update every contract that cites it in the same change.

## Verification record

What was checked when an entry was set to `verified`:

| ID | Date | Read from | Checked |
| --- | --- | --- | --- |
| `SRC-02` | 2026-11-16 | The issuer's page at the locator, which states *Effective as of 01 Jan 2023*, *Last update 05 Jul 2024* and *Status: Current* | The paragraphs that the calculation contracts cite read in full, and the others of MAR40.1 to MAR40.86 for their subject; the FAQs not reviewed, except that of MAR40.2. Every locator of the [calculation contracts](CALCULATION_CONTRACTS.md) compared with its paragraph, and every value of the [parameters](CALCULATION_CONTRACTS.md#parameters) compared with the text: the four scaling factors of MAR40.2 with the symbol each belongs to, the multiplier of MAR40.1, table 1 of MAR40.6, the 15 time bands, zones and risk weights of table 4 of MAR40.26, the 10% of MAR40.27, table 5 of MAR40.28, the rates of MAR40.43, MAR40.47, MAR40.61, MAR40.72 and MAR40.73, and the example of table 9 of MAR40.61 |
| `SRC-03` | 2026-11-19 | The issuer's publication page, which states the dates and the status *Superseded* above, and the PDF it links, whose SHA-256 was `042d60b4a3630fa647abc7b2b1606fd4d6a30c24aa7f3c8198669f6f4b4cc40b` when read; the PDF is not kept in the repository | Part C.2 read in full: the positions of paragraph 1, table 10 and the charges of paragraphs 2 and 3, and recomputed by hand, $0.05 mn + $0.08 mn + $0.45 mn + $1.00 mn + $3.00 mn = $4.58 mn. The time bands, zones and risk weights of its table 1, and the disallowances of its table 2, compared with table 4 of MAR40.26, the 10% of MAR40.27 and table 5 of MAR40.28: the weights and the percentages are the same, and so are the zones of the column for a coupon of 3% or more. The other examples of Part C are not used: C.1 is the capital ratio, C.3 the maturity ladder approach and C.4 the delta-plus method, which the workbook does not calculate |
| `SRC-04` | 2026-11-20 | The consolidated document `02013R0575-20260101` of EUR-Lex, whose SHA-256 was `355380eead4964d2f290c843666c6e1b5819253bb1f7bed3d067e01d97720fc9` when read as HTML; the document is not kept in the repository. Its header states *EN, 01.01.2026, 020.001* and the notice quoted above | Articles 94, 104c, 325, 325a, 335, 336, 341 to 344, 351, 352 and 355 to 361 read in full, Article 361 with its Table 2, which the document marks with the corrigendum C2; Articles 329, 339 and 354 read in part, the table of Article 339 and the end of some paragraphs of Articles 329 and 354 not read; the titles of Articles 326 to 361 read. The lead of Article 325(2) and its points (a) to (d), the lead of Article 325a(1), point (f) of Article 325a(2) and Article 351 compared with the Official Journal text of `SRC-05` by a program, on text with its white space and the space before a punctuation mark normalized: the same words. Not checked: each locator that a calculation contract will cite is compared with the text when the contract cites it |
| `SRC-05` | 2026-11-20 | The Official Journal text at the locator, whose SHA-256 was `c3489f04077dcd622f890f1695f5a7b5584e16ff6c96663b82dfca8ae660a53e` when read as HTML; the document is not kept in the repository | The points that amend Articles 325, 325a, 332(3), 337, 338, 348, 351, 352(2) and 361 read in full: the simplified standardised approach of Article 325(2), with its factors 1.3, 3.5, 1.2 and 1.9 and the securitisation instruments of Article 337 not scaled, and the conditions of Article 325a(1) (10% of total assets, EUR 500 million, assessed monthly). Not read: the corrigendum of 2025-04-16, which `SRC-04` holds |
| `SRC-06` | 2026-11-20 | The Official Journal text at the locator, whose SHA-256 was `081d03d633d10be4b8e742f829f5c1ea8fb74be3c7bb5a12bbcac6d2510476b5` when read as HTML; the document is not kept in the repository | Articles 495s, 495u and 495v and Article 2 read in full: the multiplier of 0.9 of Article 495s(2) until 2029-12-31, Article 495u for an institution eligible for Article 94, Article 495v not applying to an institution that applies only Articles 326 to 361, and the entry into force on the day following publication, with application from 2027-01-01. Not verified: whether the scrutiny period ended without an objection. No page of the Parliament or the Council was read, and a search found no record of one; the act's publication in the Official Journal is the only evidence that it is in force |

The page shows the factors of MAR40.2 as images of their symbols. Read as
images, they give 1.30 for interest rate risk, 3.50 for equity risk, 1.90
for commodities risk and 1.20 for FX risk.

## Decisions

| Date | Decision | Reason | Issue |
| --- | --- | --- | --- |
| 2026-11-16 | The regulatory basis is the Basel Framework, chapter MAR40, and not a jurisdiction's implementation of it | The prototype's engine, its template and its README sheet were written against MAR40. A jurisdiction's rules, such as those of the European Union, differ in scope, dates and discretions, and would be a source of their own | [18](https://github.com/danielep71/K-FRTB-SSA/issues/18) |
| 2026-11-16 | The version is the one in force from 1 January 2023, as last updated on 5 July 2024 | It is the version the issuer marks as current. The workbook supports no earlier reporting basis | [18](https://github.com/danielep71/K-FRTB-SSA/issues/18) |
| 2026-11-16 | No copy of `SRC-02` is kept under `docs/methodology/sources/` | The issuer allows redistribution for non-commercial purposes only. A copy would go with every clone of a public repository, whose license sets no such limit on its users. The locator names the exact version instead | [18](https://github.com/danielep71/K-FRTB-SSA/issues/18) |
| 2026-11-16 | The project's sources are numbered from `SRC-02` | `SRC-01` is the illustrative entry, which the illustrative test cases still cite, and an ID is never reused | [18](https://github.com/danielep71/K-FRTB-SSA/issues/18) |
| 2026-11-19 | The 2005 update of the 1996 Amendment is registered as `SRC-03` for the printed values of its example C.2 only, although it is superseded | The Basel Framework prints no worked example of the maturity method, and this one applies the same ladder weights, disallowance percentages and zones as MAR40.26 to MAR40.28. A `published` value needs a registered source: the contracts keep citing `SRC-02` for every rule, and `SRC-03` supports no rule | [47](https://github.com/danielep71/K-FRTB-SSA/issues/47) |
| 2026-11-19 | No copy of `SRC-03` is kept under `docs/methodology/sources/` | As for `SRC-02`: the issuer's terms do not allow redistribution beyond brief excerpts, and a copy would go with every clone. The register gives the SHA-256 of the PDF that was read | [47](https://github.com/danielep71/K-FRTB-SSA/issues/47) |
| 2026-11-20 | The texts of the CRR regime are registered as `SRC-04` to `SRC-08`, sources of their own: none replaces `SRC-02`, which the Basel regime keeps citing | The decision of 2026-11-16 said that a jurisdiction's rules would be a source of their own. [#74](https://github.com/danielep71/K-FRTB-SSA/issues/74) added the CRR regime to the milestone v0.3.0, with the Basel regime unchanged | [75](https://github.com/danielep71/K-FRTB-SSA/issues/75) |
| 2026-11-20 | A rule of the CRR regime cites `SRC-04` for its locator, and is checked against the Official Journal text of `SRC-05` or `SRC-06` where CRR3 or the delegated act amended the article | The consolidated text has no legal effect, and the Official Journal text is the authentic one | [75](https://github.com/danielep71/K-FRTB-SSA/issues/75) |
| 2026-11-20 | No copy of `SRC-04` to `SRC-08` is kept under `docs/methodology/sources/` | A copy would go with every clone of a public repository, and I did not read the terms under which EUR-Lex allows its content to be reused: the legal notice returned no content on 2026-11-20. The register gives the SHA-256 of the HTML that was read. I reconsider this when I have read the terms | [75](https://github.com/danielep71/K-FRTB-SSA/issues/75) |
| 2026-11-20 | A technical standard of the CRR that no adopted act supports is an assumption of the rule that depends on it, and the table of the standards says which | `SRC-07` and `SRC-08` are identified but not read in the Official Journal, and no adopted act was found for Articles 94(10) and 325(9) | [75](https://github.com/danielep71/K-FRTB-SSA/issues/75) |
