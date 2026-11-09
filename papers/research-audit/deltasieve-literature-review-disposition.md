# DeltaSieve literature and evaluation review

## Intake and source identity

The DeltaSieve review became available on 2026-11-09 in Markdown, PDF and DOCX
forms. The Markdown input contains literature leads and evaluation proposals.
Its numbered references lack a resolvable bibliography. The PDF exposes no
research-source hyperlinks, and the DOCX contains no external relationships.
Primary-source checks therefore establish the citations used below.

| Input filename | SHA-256 |
| --- | --- |
| `deep-research-report-ds.md` | `6bc397552962b40cd8d2bd99e8bdccf423ffb4eee27c964e779d7e0f2aa198fc` |
| `DeltaSieve_ Related Work, Evaluationskritik und Empfehlungen.pdf` | `b091bfaee5f4ab485160fef55a2e6e23c34e4923a02aab5d856fcc61783ffe23` |
| `DeltaSieve_ Related Work, Evaluationskritik und Empfehlungen.docx` | `691c616e66edb8fcd98e3bcd884b7af88b751a244cf7223e858aa6365e3adf15` |

The three files are retained privately with the audit inputs. The published
draft uses checked primary literature and documented implementation evidence.
Recommendations in the input are investigation leads, with disposition below.
The SheetShrink review remains pending.

## Verified literature and manuscript changes

| Topic | Verified finding | Draft disposition |
| --- | --- | --- |
| SheetDiff | Chambers, Erwig and Luckey, VL/HCC 2010, pages 85–92, DOI `10.1109/VLHCC.2010.21`. The original paper presents inferred spreadsheet changes in one contextual view. | Added to related work. No claim that spreadsheet change inference itself is new. |
| RowColAlign | Harutyunyan, Borradaile, Chambers and Scaffidi, VL/HCC 2012, pages 7–14, DOI `10.1109/VLHCC.2012.6344471`. The indexed original abstract describes two-dimensional alignment and planted edit evaluation. | Added as a precedent for controlled spreadsheet mutations, with native visible-effect ground truth as an additional evaluation duty. |
| Ekstazi | ISSTA 2015, dynamic file dependencies and checksums. The input's ICSE 2014/call-graph description is incorrect. | Existing correct citation retained and enriched with DOI/pages. Mechanism clarified. |
| ReDeCheck | ISSTA 2017. It extracts responsive layout graphs using browser/DOM observations, then compares or analyzes those models. Screenshots support diagnosis. | Existing correct citation and model distinction retained. The input's FSE attribution and generic screenshot-diff characterization were not adopted. |
| Synkronizer | Edition documentation lists cell content, format and Excel-name comparison features. | Added as a vendor-documented comparator candidate, with edition and empirical coverage still requiring verification. |
| SpreadsheetsDiff | Vendor describes two-/three-way comparison, keys and workbook/VBA features. | Added as a vendor-described candidate. No untested speed or coverage ranking. |

Primary material for these checks is available at the following locations.

- [SheetDiff original paper](https://web.engr.oregonstate.edu/~erwig/papers/SheetDiff_VLHCC10.pdf)
- [SheetDiff publisher record](https://store.computer.org/csdl/proceedings-article/vlhcc/2010/05635197/12OmNzwpUg9)
- [RowColAlign indexed original manuscript](https://citeseerx.ist.psu.edu/document?doi=3b0e8127db450b34319f95a6be6a8090e01058e9&repid=rep1&type=pdf)
- [VL/HCC 2012 program](https://vlhcc2012.di.unisa.it/program/final-program/)
- [Author's publication record](https://anna.harutyunyan.net/wp-content/uploads/2017/08/anna-harutyunyan-cv-apr-2017.pdf)
- [Ekstazi original paper](https://users.ece.utexas.edu/~gligoric/papers/GligoricETAL15Ekstazi.pdf)
- [ReDeCheck original paper](https://www.gregorykapfhammer.com/download/research/papers/key/Walsh2017a-paper.pdf)
- [Synkronizer edition documentation](https://help11.synkronizer.com/intro_editions.htm)
- [SpreadsheetsDiff vendor description](https://spreadsheetsdiff.com/)

The RowColAlign full PDF fetch returned HTTP 403. Its indexed original abstract,
conference program and author record support the limited description above.
Detailed algorithm claims and reported accuracy values await full-text checking.

## Claims requiring correction or stronger evidence

| Input claim or proposal | Assessment |
| --- | --- |
| Only actually changed cells are rendered. | The planner includes dependencies and conservative fallback scopes. Selected regions can render equally. The existing method and new metrics preserve that distinction. |
| No existing system performs the same task. | The search supports relevant prior-work distinctions. It cannot establish an exhaustive absence claim or certify novelty. No universal priority claim was inserted. |
| Structural comparator descriptions establish all their omissions or full logical coverage. | A vendor feature list does not establish exhaustive coverage, cost complexity or missed rendering classes. Product/version-specific controls remain necessary. |
| DeltaSieve uses Excel or LibreOffice. | The inspected execution protocol uses native Excel with Office.js workers. No LibreOffice execution result supports that statement. |
| Hidden rows or layout-only changes are outside selection. | Geometry, row/column visibility and effective layout belong to the documented selection model. Mutation coverage remains an empirical obligation. |
| Dependency tracking establishes soundness. | Unsupported representations, native effects and incomplete reference evidence bound the claims. Existing fallbacks and explicit unsuccessful outcomes remain in the method. |
| The replay mechanism is a future VM-like computation recording system. | Current retained evidence and portable inspection have their own documented semantics. No verified basis was found for the proposed VM analogy or an unspecified REX citation. |
| Spreadsheet errors imply malformed workbooks. | Division-by-zero and reference-error values can be valid rendered calculation outcomes. They are separate fixtures from unreadable packages. |
| A selected equal image is a detector false positive. | It measures selection overhead. A reported pixel difference, its cause and a confirmed transformation regression have separate labels. |
| A fixed 30–50 cases provides a general confidence margin. | Precision depends on the event rate and sampling unit. Correlated mutations cannot be counted as independent workbooks. |

The input's examples of volatile functions also include MATCH without a verified
basis for treating it as volatile. The prepared controls use documented NOW and
RAND cases, with INDIRECT and OFFSET covered by explicit reference-policy checks.
Full native rebuilds can change observed values even when formula text is equal.
[Microsoft Excel recalculation](https://learn.microsoft.com/en-us/office/client-developer/excel/excel-recalculation)

## Evaluation changes applied

The draft now defines the workbook-pair unit, partial coverage, usable-reference
fraction, selection overhead and pixel-coordinate mapping requirements. The
mutation protocol spans content, effective formatting, conditional dependencies,
charts/objects, geometry/visibility, unsupported references, representation-only
controls and repeated identical native states. It records intended mutations
separately from the visible effects confirmed by the independent renderer.

For zero observed misses in independent Bernoulli trials with a common rate,
solving `(1-p)^n = 0.05` gives the one-sided 95% upper limit
`p = 1 - 0.05^(1/n)`. Values are 0.0950338529 at n=30 and 0.0581550791 at n=50.
These illustrative calculations are not study results. The draft cites the
NIST exact-binomial framework and states independence, clustering and sampling
limits. It makes no new population accuracy claim.
[NIST binomial confidence intervals](https://www.itl.nist.gov/div898/handbook/prc/section2/prc241.htm)

Cost reporting now includes all stated end-to-end phases, rendered pixels per
state, repeated work from overlapping captures and separate coverage-union size.
Matched selective/reference trials must use the same preparation, reuse and
environment policies. Timeouts retain censored status and completion counts.
The input's proportional-cost intuition needs measured analysis and native cost.

## Remaining evidence

- Complete and reconcile the actual mutation matrix and native reference sample.
- Freeze outcomes and cluster identities before computing population estimates.
- Run or locate matched cost observations for the claimed evaluation frame.
- Read the complete RowColAlign paper before expanding its algorithm discussion.
- Verify vendor capabilities under tested versions before extending the bounded
  comparator pilot or adding negative capability claims.
- Incorporate the SheetShrink review when available.

This checkpoint updates literature and evaluation prose. It does not add native
experiment results. The original-workbook Pandas prohibition remains in force.
