# SheetShrink literature review and scope assessment

## Input identity

The review received on 2026-11-09 defines SheetShrink as a data-normalization
and error-repair system. That problem statement differs materially from the
versioned cleanup contract. References and evaluation proposals were therefore
checked individually against the stored-package transformation studied here.
Embedded recommendations are treated as review material.

| Input filename | SHA-256 |
| --- | --- |
| `deep-research-report-sheet-shrink.md` | `d917e564bc6fa368959b48134942529c391872289f978005ee2bc112bdca5ac5` |
| `Executive Summary.pdf` | `c43fcdb6a504b83f223cc61d54ea82340e8b54f30b6f56f767b5302a5b52a0e1` |
| `Executive Summary (1).docx` | `4431711a309d9401d20ae58d7cea0de72899309cc2dc5f48706e4cc0bb23ef34` |

The Markdown was read in full. The five-page PDF and DOCX were checked for
content differences and source links. Text comparison found the same substantive
discussion, with citation markers and export formatting differing. Neither export
supplies research-source hyperlinks. The input's
reference list provides titles but lacks the promised DOI/link entries.
Private copies are retained with the audit inputs outside the paper artifact.

## Assessment criteria

The relevant comparison has two independent dimensions.

1. The transformation objective and preservation contract. SheetShrink removes
   admitted package redundancy and retains required workbook effects. Its
   production policy includes guarded hidden constant tails and preserves
   ordinary nonempty hidden tables. Formula repair, deduplication, date
   canonicalization and row reordering do not describe the evaluated method.
2. The information available when deciding. Cleanup of an XLSX input uses
   stored package contents before Excel loads the workbook. Native layout
   measurement and rendering belong to subsequent validation. Legacy XLS
   conversion is a separate upstream stage.

An Excel-hosted experiment can observe effective row geometry and compare
candidate edits. Fixing a measured row height through `ht` and `customHeight`
requires a trusted native measurement and a policy for changing automatic
sizing to fixed height. Writing these attributes is possible at package level,
but the current cleaner obtains no native measurements during its decisions.
Its font witnesses retain required effects using the package context. Stored
explicit heights remain available directly. These facts justify a distinction
in information and dependency. They do not establish a general computational
complexity ordering between every hosted and package-based method.
[Microsoft custom-height schema documentation](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.row.customheight)

Pre-opening execution is also insufficient on its own to establish novelty.
The closest package tools found during this check must remain in related work.

## Source-by-source disposition

| Source | Verified problem and setting | Relevance and action |
| --- | --- | --- |
| Gulwani, Automating String Processing in Spreadsheets using Input-Output Examples, 2011 | Example-driven synthesis of string transformations. The publication is POPL 2011, with an Excel add-in prototype. | Context for intended value transformation. The input taxonomy's PLDI attribution is incorrect. No separate citation added because the table-synthesis representative covers the distinction compactly. |
| Harris and Gulwani, Spreadsheet Table Transformations from Examples, 2011 | TableProg and ProgFromEx synthesize transformations from example input/output tables. William R. Harris and Sumit Gulwani are the authors. | Added one concise related-work citation for table synthesis. It does not supply a package-layout preservation result. |
| Singh and Gulwani, Learning Semantic String Transformations from Examples, 2012 | Synthesis combines semantic table lookup and syntactic string operations. | Valid literature lead for value transformation. Audited as context, represented by the shorter synthesis discussion instead of an additional near-duplicate citation. |
| Singh and Gulwani, Transforming Spreadsheet Data Types using Examples, 2016 | A probabilistic DSL handles ambiguous parsing/formatting using examples. POPL 2016. | Context for data-type normalization. Its repair objective differs from retaining an existing workbook's required effects. No separate main-text citation needed at this checkpoint. |
| Raman and Hellerstein, Potter's Wheel, 2001 | Interactive transformation, immediate record feedback and discrepancy detection. | Added as a representative interactive data-cleaning system. It is not classified as necessarily dependent on Excel. |
| Rekatsinas et al., HoloClean, 2017 | Probabilistic data repair combines constraints, external evidence and statistics. | Added a concise objective distinction. Its repair precision/recall values are not SheetShrink baselines or expected reduction metrics. |
| Hermans, Pinzger and van Deursen, formula smells, 2015 | Metrics identify formula smells, visualize them and suggest refactorings. Journal DOI `10.1007/s10664-013-9296-2`. | Added as adjacent spreadsheet refactoring. Maintainability and formula refactoring need their own behavioral obligations. Static execution alone does not solve native layout preservation. |
| Wrangler | The primary work is Kandel, Paepcke, Hellerstein and Heer, CHI 2011, on interactive transformation scripts. | The input's speculative SIGMOD 2012 label is corrected. Broader wrangling context is represented by Potter's Wheel to keep the section focused. |
| OpenRefine | Official history describes a community project renamed in 2012 after Google leadership. | The input's Apache attribution is unsupported by that history. Keep as optional wrangling context rather than a close workbook-preservation comparator. |
| Power Query, Remove Duplicates, Text to Columns, Flash Fill and validation UI | The report groups heterogeneous operations under normalization. | Assess their actual transformation goals separately. None of the grouping or feature names establishes a stored-package redundancy-removal contract. No categorical capability comparison was inserted. |
| Spreadsheet Compare, Synkronizer and DiffEngineX | Difference reporting and comparison roles are documented separately in the DeltaSieve audit. | Comparison evidence belongs to the validation discussion. These tools were not presented as automatically correcting or normalizing workbooks. |
| Unspecified Hellendoorn static-analysis lead | No paper title or resolvable citation is supplied in the report. | Deferred until an identifiable source can be verified. |

The checked primary literature is available at the following locations.

- [Gulwani, POPL 2011](https://www.microsoft.com/en-us/research/publication/automating-string-processing-spreadsheets-using-input-output-examples/)
- [Harris and Gulwani, TableProg manuscript](https://www.microsoft.com/en-us/research/wp-content/uploads/2012/08/pldi11-table-synthesis-tr.pdf)
- [Singh and Gulwani, semantic transformations](https://people.csail.mit.edu/rishabh/papers/semanticVLDB12.pdf)
- [Singh and Gulwani, data types](https://www.microsoft.com/en-us/research/wp-content/uploads/2016/06/popl16-1.pdf)
- [Potter's Wheel original paper](https://cs.brown.edu/courses/cs227/archives/2016/papers/PottersWheel.pdf)
- [HoloClean original paper](https://www.vldb.org/pvldb/vol10/p1190-rekatsinas.pdf)
- [Hermans et al., accepted manuscript](https://repository.tudelft.nl/file/File_872dd231-2eb8-4ab4-a527-14bac6bc842f)
- [Wrangler project publication record](https://idl.uw.edu/papers/wrangler)
- [OpenRefine project history](https://openrefine.org/openrefine_history)

## Closer cleanup comparisons added by the scope check

- [x] Microsoft Check Performance. The documentation describes unwanted-cell
  cleanup and explicitly warns about removal of pixel-art cells. Added as a
  native-host baseline with a documented policy boundary. No independent
  regression verdict follows from the warning alone.
  [Microsoft documentation](https://support.microsoft.com/en-us/excel/cleanup-cells-in-your-workbook)
- [x] Inquire Clean Excess Cell Formatting. Microsoft describes cleanup beyond
  the last nonblank cell and possible changes to conditional-formatting ranges.
  Added separately from Spreadsheet Compare, which has a comparison role.
  [Microsoft documentation](https://support.microsoft.com/en-us/excel/clean-excess-cell-formatting-on-a-worksheet-in-excel-for-windows)
- [x] XLStylesTool. Its developer describes an Open-XML-based excessive/corrupt
  style cleaner independent of Excel's object model. This is direct prior
  package-cleanup work. The developer's speed and exclusivity claims were not
  adopted as measurements. The inspected description does not establish the
  same cell-removal rules, compact witnesses or validation contract.
  [Developer description](https://sergeig888.wordpress.com/2011/03/21/net4-0-version-of-the-xlstylestool-is-now-available/)
- [x] NXPowerLite. Desktop 10 explicitly lists XLSX excess-format cleanup. The
  engine description discusses formatted blank cells in oversized worksheet
  parts. Added as a closer comparator than data repair. No measured superiority
  or native pixel-preservation outcome follows from its marketing example.
  [Desktop 10 features](https://neuxpower.com/blog/whats-new-in-nxpowerlite-desktop-10),
  [Engine mechanism description](https://neuxpower.com/blog/wecompress-file-compressor-engine-upgrade)
- [x] NXPowerLite dependency provenance. An older manual explicitly permits
  operation without installed Office. Keep that historical fact version-scoped.
  A current product's local/offline label alone does not establish its use or
  avoidance of native Office processes.
  [Historical manual, system requirements](https://cdn.neuxpower.com/releases/desktop/5.1.6/UsingNXPowerLite.pdf)
- [ ] Obtain matched candidate outputs and record versions/settings before
  claiming finer capability differences or comparative reduction and cost.

## Evaluation proposals

| Proposal in the input | Disposition |
| --- | --- |
| Infer and fix duplicates, date formats, row alignment or broken formulas. | Outside the studied transformation. Replaced with eligible redundancy, style masks, font witnesses, merges and dependency guards. |
| Treat each corpus workbook as a known clean reference. | A source workbook provides an observed baseline. Independent structural/native checks establish preservation evidence. |
| Compute repair precision and recall from injected semantic errors. | Does not measure the stated cleanup objective. Report reduction, retained witnesses, runtime, memory and preservation outcomes. Detection recall belongs to validation coverage under its own reference protocol. |
| Treat reordered columns as an identity-preserving control. | Column reordering can alter visible layout and formula/object behavior. Package entry order and compression are more appropriate representation controls. |
| Measure the set of suspicious regions as SheetShrink's error-localization denominator. | This confuses cleanup guards with DeltaSieve's capture selector. Their decisions and evaluation units remain explicit. |
| Stratify inputs, retain failures and quantify uncertainty. | Applicable with the actual corpus frame, source identities and supported cleanup configurations. |
| Include hidden content, volatile formulas, objects and complex inheritance. | Applicable as preservation stress classes, with declared policies and separate native recalculation controls. |
| Infer universal preservation from equal screenshots. | Finite selected captures establish bounded observations. Structural checks, independent references and explicit unsuccessful outcomes remain necessary. |

The performance protocol now records a cleaner's input model, native dependency,
access to measured layout and removal policy. First-opening comparisons include
the required source load for hosted cleaners. Package cleanup, opening, saving,
reopening and validation costs have explicit phases. Repeated-use measurements
report initial preparation separately. These definitions prevent a later native
cleanup stage from receiving free access to the source workbook in a first-open
comparison. No new performance result is introduced at this checkpoint.

## Manuscript actions and subsequent review

- [x] State the pre-native-loading operation in the abstract and introduction.
- [x] Describe the information available to the cleaner in the contract.
- [x] Explain the measured-height alternative and its different prerequisites.
- [x] Prioritize actual bloat/style cleanup before broader data transformation.
- [x] Cite representative synthesis, repair and formula-refactoring work with
  specific reasons for its relevance.
- [x] Keep parsing, loading and validation work in their respective roles.
- [x] Correct the evaluation objective and phase accounting.
- [ ] After the manuscript's methods and results are sufficiently settled,
  perform a fresh related-work search with that manuscript as the primary
  attachment. Ask for omitted close approaches, counterexamples and original
  sources, with extraction of required inputs, native dependencies, removed
  structures and preservation evidence for every proposed comparator.

The subsequent search should begin from the supplied manuscript's definitions.
It should assess the full combination of cleanup rules, information available
before native loading and retained effects. The current review does not establish
an exhaustive absence of prior art or a formal complexity claim.
