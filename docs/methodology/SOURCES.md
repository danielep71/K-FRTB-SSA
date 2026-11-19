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

`SRC-02` is the only source a rule of the workbook relies on. `SRC-03` is
registered for the printed values of one worked example, and no rule cites it:
its document is superseded by `SRC-02`, and the example is only a published
value that a case compares the maturity method with. The illustrative `SRC-01` stays with the illustrative test cases that cite it,
which the example suite of the test harness reads
([the project's suites](TEST_HARNESS.md#the-projects-suites)); no rule of
the workbook cites it.

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
