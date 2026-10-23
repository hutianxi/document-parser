# Changelog

Records the **externally visible changes** of each document-parser version: keys in the returned
structure, methods on `WordParser.Options`, public entry points, documented behaviour, and the measured
numbers each change produces on the three samples in this repository. The source javadoc and
[README.en.md](README.en.md) remain authoritative for implementation details;
the Chinese version is [CHANGELOG.md](CHANGELOG.md).

Newest first. The date after a version number is the **commit date of that tag**. When the working tree
holds unreleased changes, they are listed in a separate section at the top.

---

## 1.2.1 — 2026-10-24

### Build / dependencies

- Jackson 2.13.5 → 2.21.8, closing CVE-2026-83557 / 68497 / 19032 / 89425 / 89407 / 91776 and fixing
  CVE-2026-110879 / GHSA-m8xf-2cr7-77ph (published 2026-10-20; an incomplete fix for CVE-2026-54512 where
  generic arguments and array element types could still bypass the PolymorphicTypeValidator). `DatabindContext`
  in 2.21.8 adds array-element-type validation, and `ApprovedBaseTypeValidator` only exists in that release;
  2.21.7 and earlier do not contain that code. Jackson 2.x is still Java 8 bytecode (`ObjectMapper.class`
  major version 52). `mvn -B clean package` runs **242 tests with 0 failures and 1 skipped** under both JDK 8 and JDK 21;
  the JDK 8 main jar is **60777 bytes, 30 entries, 17 classes**, and the same source gives **61190 bytes** on
  JDK 21 with the same entry and class counts.
- The test dependency changed from `spring-boot-starter-test` to `org.junit.jupiter:junit-jupiter` (5.8.2,
  still managed by the parent pom), so spring-boot 2.7.18, Spring Framework 5.3.31, logback 1.2.12,
  snakeyaml 1.30, log4j-api 2.17.2, json-smart 2.4.11 / json-path 2.7.0, assertj 3.22.0, xmlunit 2.9.1
  and asm 9.3 no longer enter the test classpath. Public API, result structure and parsing behaviour are
  unchanged.

This version is published on Maven Central (2026-10-24; Portal deployment `11cdf8ca-0884-4626-9356-c82a6455144e`,
status PUBLISHED). The published `document-parser-1.2.1.jar` is **60777 bytes** and the `.pom` is **12136 bytes**,
byte-identical to the local release bundle; all four `.sha256` files were checked, and the main jar's is
`ce9af926820b55b18505d2efa51a02cbd92dbe73376fda7a46d3fec6298c0ca3` (the other three are in README.en.md §17).
1.2.0 is published on Central, but its POM still points at Jackson 2.21.7; use 1.2.1 to get 2.21.8.

## 1.2.0 — 2026-10-23

`pom.xml` is now `1.2.0` and `mvn -B clean package` produces `document-parser-1.2.0.jar`.
This release folds in everything externally visible since `v1.1.1`: the three new APIs (`source` / `name` /
`unmatchedHeaders`), the lookup-by-name entry points, borderless-table recovery in PDF, and the three defects
found before release. The output changes relative to 1.1.1 are under "Behaviour"; the test and build numbers
are under "Build artifacts".

### Added

- **Borderless PDF tables are now read**: besides the line-detection engine the flow engine is run a second
  time, and only those blocks that "have the same cell count in every row and overlap no ruled table" are
  taken. Measured on a borderless document built by a throwaway probe: **0 groups → 3 groups**; the ruled
  `sales-report.pdf` in this repository gained no group. Known cost: a few lines of loose text aligned
  with tab stops satisfy the same test and are therefore moved from `text` into `tables` as one table
  (`PdfBorderlessTableTest.tabAlignedParagraphsCurrentlyCountAsATable` pins this behaviour as an
  assertion). Test class `PdfBorderlessTableTest` (5 cases; one pins "prose above the table is not
  swallowed", see "Fixed").
- **Every table group carries its position, `source`**: `{ "page": <first page> }`, plus `"lastPage"` when
  the table spans a page. One key name and one field set across all three formats: for docx the page comes
  from `LayoutCollector`, for pdf it is the page number, for pptx it is the slide index — all 1-based.
  Measured: all 19 groups of the three samples have a page value, and the group that spans two pages is the
  collection ledger — one in docx and one in pdf, both `{"page": 1, "lastPage": 2}`. Written centrally by
  `Tables.markSource`; the grouping algorithm
  `Tables.group` does not touch it.
- **`Options.source(boolean)`** (default `true`): switch the page numbers off and the change applies to
  **all three formats at once**; each group is then left with `headers` and `rows` only, column names and
  content unchanged. Switching off also stops docx from building a `LayoutCollector` at all, which saves
  the whole-document layout pass — measured by a throwaway probe that ran off and on **alternately**, six
  times each, twice through, in the same JVM: **on 74–254 ms / off 11–35 ms**, and in both rounds the
  maximum of "off" stayed below the minimum of "on" in the same round. PDF/PPT page numbers are free
  anyway, so there the switch only decides whether the key is written (same rounds: pdf 681–960 ms vs
  696–984 ms, pptx 46–86 ms vs 44–71 ms — the ranges overlap). The probe was deleted after the run and is
  not in the repository.
- **`Options.name(String)` and the `name` key on table groups**: give a registered header group a name and,
  once it matches, that group gains `"name": "季度明细"` (`"Quarterly detail"` in English), so row loading
  locates the table by name
  instead of counting indexes into `tables`. The key is written at the moment a comparison against a
  registered group succeeds (inside `Tables.group`), which is why all three formats behave identically;
  unnamed groups, and tables whose headers did not come from a registered group, **do not carry the key**.
  Measured JSON lengths (same registration, before naming → after): docx **4740 → 4828** (4 of 7 groups
  named: 1, 2, 6, 7), pdf **6473 → 6561** (4 of 6: 1, 2, 5, 6), pptx **3378 → 3466** (4 of 6: 1, 2, 4, 5).
  Naming introduces no layout work, and no separate on/off timing was measured for it — those lengths,
  together with every other length quoted in the documentation (silent, registered-without-names,
  registered-with-names, BASE64 — four settings × three formats, **12** numbers in all), are now asserted
  by `CrossFormatTest.documentedJsonLengthsAreTheRealOnes`, so a regression turns the build red.
- **Retrieval by name**: `WordParser.tablesByName(result, name)` returns every group carrying that name (in
  `tables` order, returning the very objects held in the result);
  `WordParser.rowsByName(result, name)` flattens their data rows in the same order into one list that can
  be inserted row by row as it stands. No match, `null`, or a blank name all give an **empty list** — never
  an exception. Test class `TableByNameTest` (11 cases).
- **Top-level `unmatchedHeaders` switch** (`Options.unmatchedHeaders(boolean)`, default `false`): when on,
  the result gains a top-level key after `text` / `tables` / `images` that lists every registered header set
  which matched no group. Each item normally has only `headers` (the registration as written); if that
  registration also had a name, it carries a `name` too. It is the same rule as the stderr hint, just
  available to programs. Measured JSON lengths (named + switch on): docx **4857** (all five registrations
  matched, so the value is an empty list), pdf **6632**, pptx **3537** (one item,
  `["故障部件","工单数"]`). Tests `CrossFormatTest.unmatchedHeadersIsAnOptInTopLevelKey` and
  `unmatchedHeadersCarriesTheNameOfARegistrationThatMatchedNothing`.

### Fixed

The defects below were all found this round by going through the code method by method and assembling
minimal documents on the fly. **Each one was pinned by a failing test before the code was changed** — none
of them is a guess from reading. The "row merged across the full width" test is now the same rule in all
three formats.

- **A vertically merged data row in PPT was mistaken for a group row and its data was eaten**
  (`PptParser.parseTable`): the count used to be of *origin cells* (`getFirstRowIndex() == r` **and**
  `getFirstColumnIndex() == c`), so a vertical-merge continuation did not count as one of the row's own cells.
  In a two-column table whose column A is merged down two rows while column B has a value in each, the second
  row was left with a single origin cell and was therefore taken for a full-width group row: measured, `KEEP`
  blanked `value 2` (the result read `[{colA:first,colB:value1},{colA:first,colB:}]`) and `DROP` deleted the
  row outright, while the same layout in docx read `value 2` in both modes. The rule now excludes **horizontal**
  continuations only — the same rule as `WordParser.Grid.isSeparator` and `PdfParser.toGrid` — after which both
  modes match docx cell by cell. Test `PptParserTest.verticallyMergedDataRowIsNotAMistakenSeparator`.
- **A blank row merged across the full width was taken for a group row, splitting one table in two and eating
  the next data row as a header** (`WordParser.Grid.isSeparator` and `PptParser`; `PdfParser.toGrid` already had
  this rule, so it was the other two formats that were out of line): a spacer row with no text at all also
  satisfies "only one cell left in the row", so the row below it was consumed as a header by the
  "the row after a group row is a header" rule. Measured on a 4-row 2-column table (header `colA/colB`,
  `A B`, a blank merged row, `C D`), before the fix the parse gave **2 groups**: group 1 named `colA/colB`
  holding only `A B` plus that row of empty values, and group 2 named after the data `C/D` with **0 data rows**
  (the `C D` row was lost entirely); under `DROP` group 1 was left with a single row. Adding "**and the
  row is not entirely blank**" gives 1 group, column names still `colA` / `colB`, 3 data rows, and the blank row
  itself is kept as one row of empty values rather than silently dropped. Tests
  `WordEdgeCasesTest.blankFullWidthMergedRowIsNotASeparator` and
  `PptParserTest.blankFullWidthMergedRowIsNotASeparator` (both **failed first** — `expected: <1> but was: <2>` —
  and pass after the fix).
- **`Images.extension` followed the default `Locale`**: the no-argument `String.toLowerCase()` uses the
  platform default locale, so under Turkish (`tr_TR`) `"TIFF"` became `tıff` (dotless i) — a wrong extension,
  a wrong MIME subtype and a wrong file name. All three call sites (`Images.extension`, `Images.contentType`,
  `Results.normalize`) now use `Locale.ROOT`. Test `ImagesTest.extensionDoesNotFollowTheDefaultLocale` sets the
  default locale to `tr_TR` and asserts `tiff` / `png` / `jpg` / `image/tiff`.
- **When a borderless PDF table was supplemented, one prose line above it silently disappeared**
  (`PdfParser.flowTableFlow`): the flow engine's table bounding rectangle is about 25 pt taller than the
  actual cells (that gap is exactly where the evaluation watermark sits), so a prose line just above the
  table fell inside the bounding rectangle but outside every cell. The "already inside a table" test in
  `textFlows` therefore skipped it — it appeared in neither `tables` nor `text`. Registering the **cell
  rectangles** (`grid.bounds`, from `AbsorbedCell.getRectangle()`) for the prose check brings that line back
  into `text`; the ruled-table path (`tableFlow`) still uses the rectangles drawn by the lines, which cannot
  have this kind of inflation. Test `PdfBorderlessTableTest.proseAboveTheBorderlessTableStaysInText`.
- **The merged result of a multi-level header was not compared against the registrations again**
  (`Tables.group`): the docx ledger's two header rows are separate grid rows, so registering the **merged**
  row missed both when compared row by row (the upper row differs in 2 cells; the lower row has its trailing
  blanks trimmed and therefore a different column count). A registration meant to be shared across formats
  thus missed in docx. After the fix, consecutive header rows are merged and the merged result is compared
  once more, so docx hits as well — even at tolerance 0 (cells must be equal). With only the ledger
  registered, the three samples measure docx 13 rows, pdf 31, pptx 7, all with the columns
  `[客户名称, 合同金额, 已回款, 回款率, 备注]`. `TableByNameTest.aMultiLevelHeaderHitsByItsMergedRowInAllThreeFormats`
  pins tolerance 1 and 0 and all four registration spellings.
- **After merging one fragment of a cross-page table, the "last block on the page" flag was not updated,
  so an independent table on the next page could be merged in** (`PdfParser.mergeSplitTables`): merging
  replaced `page` and `bottomGap` with the later fragment's values but left `lastOnPage` at the first
  fragment's `true`. A two-page table cannot show this (the second comparison already reaches the last
  fragment), but with three or more pages, when the middle page's fragment is followed by an image or any
  other block, the next page's independent table is merged into the ledger as long as the column count and
  left edge line up. Fixture: page 1 = header + 29 rows, page 2 = 29 rows followed by an 8×8 PNG, page 3 =
  an independent two-column table `B1…B5`. Before the fix: **1 group** (`source={page=1,lastPage=3}`, with
  `B1…B5` folded into the ledger); after adding `previous.lastOnPage = flow.lastOnPage`: **2 groups** — the
  ledger with 58 rows, the independent table with `B2…B5` (4 rows) and `source={page:3}`. Test
  `PdfCrossPageTableTest.theChainAlsoStopsWhenANonTableFlowFollowsTheFragment` (removing the line reproduces
  1 group).
- **Multi-paragraph PPT text frames were flattened into one line** (`PptParser.addText`):
  `ITextFrame.getText()` separates paragraphs with `\r` (Windows uses `\r\n`), but the old code split on
  `\n` only, and `Results.clean` then turned each `\r` into a space, so `第一段\r第二段` was read as
  `第一段 第二段`. All three line endings are now split, and blank paragraphs take no line. Evaluation mode
  returns only the first paragraph per text frame, so an end-to-end test cannot see multiple paragraphs;
  the split was therefore extracted into package-private `textLines` and is pinned by the new
  `PptParserTextTest` (2 cases).
- **A table nested exactly 8 levels deep was falsely reported as "nesting exceeded"; deep leaf nodes were
  falsely reported too** (`WordParser.parseNestedTables` / `walk`): the 8th level is already the limit, but
  it has no child table, yet the old code printed "content deeper than this has been skipped" at the
  `parseNestedTables` entry as soon as `depth >= 8` — a normal 8-level document was reported as exceeding
  the limit. Likewise, `walk` checked depth before `isComposite`, so a leaf node at depth 16 (a bookmark-like
  node with no children to descend into) was reported as well. The warning is now emitted once and only when
  a 9th-level child table is actually skipped, and the child check comes before the depth check. Measured:
  8 levels → **8 groups, empty stderr**; 9 levels → **1 warning, first 8 groups unaffected**. Tests
  `WordEdgeCasesTest.exactlyEightNestedTablesAreAllReadWithoutAWarning` and
  `ninthNestedTableIsSkippedWithExactlyOneWarning` (the latter also asserts the warning appears exactly once).

### Documentation fixes

Errors found in the documentation in the same round (all corrected, not carried as known issues):

- "only the pdf ledger spans two pages" was wrong — `CrossFormatTest.sourceRecordsThePageEachBlockStartsOn`
  asserts the docx ledger is `{"page": 1, "lastPage": 2}` as well, and the arithmetic behind the 331 / 287 / 264
  deltas (`7 × 44 + 23`) confirms docx has exactly one `lastPage`.
- `-Dtest=PrintResultTest` prints **17** reports (three silent sample reports, docx "registered + DROP",
  three "registered + named", by-name matches, the four ledger spellings, the four public-entry-point
  reports, the header-naming and field-cache report, serialisation and image export, the option matrix,
  the unmatched list, the four image modes, and the error paths), not "one report per sample".
- "both docx header rows are two columns off the registered name" was imprecise: measured, the upper row
  differs in **2 cells** (beyond the default tolerance of 1) while the lower row has its trailing blanks
  trimmed and so differs in **column count**. Those two facts only explain why a row-by-row comparison
  missed; after the merge-then-rematch fix, registering the merged row hits in all three formats (tolerance 0
  included). The documentation now says so and the tests pin tolerance 1 / 0 and all four spellings.
- The 12 JSON lengths quoted in the documentation used to be prose only; they are assertions now
  (`CrossFormatTest.documentedJsonLengthsAreTheRealOnes`).
- **"the round is ~80 s slower than 1.1.1's, and that class went from 31.5 s to 76.5 s" is retracted**: it
  compared a round that was competing for CPU with an idle one. Running the class three times with that
  assertion and three times without it, alternating in the same JVM, gives 34.6 / 30.9 / 37.6 s against
  29.8 / 35.9 / 38.2 s — the difference disappears into round-to-round noise. The only cost that can be
  attributed precisely is the per-method figure in the surefire report
  (`documentedJsonLengthsAreTheRealOnes` **3.3–3.8 s on JDK 8 (two runs) / 3.2 s on JDK 21** in this round; earlier rounds
  measured around 3.5 / 5.6 s, from
  `target/surefire-reports/TEST-*.xml`). §11 and §13 of both READMEs are rewritten to measured numbers and
  now say plainly: read durations as an order of magnitude, not as a performance-regression signal.
- **The UTF-8 copies' character counts were undercounted**: Jackson's pretty-printer emits `\r\n` on Windows
  and the earlier tally skipped the `\r`. Re-measured with `wc -m`, the **18** copies total **191,974**
  characters (shortest 1,351, longest 106,026; the copies carrying a JSON block contain CRs and the count
  includes them).
- **The English README's example literals do not match the bundled samples**: §5.3 and §13 used translated
  column names, group names and keywords, while the three samples are Chinese documents. A note now spells out
  the real literals (`片区` / `销售额` / `季度明细` / keyword `销售` and so on), and the snippet under
  "Verifying one document from code" — which parses a bundled sample path — was corrected from
  `keywords("sales")` to `keywords("销售")`.
- Test counts, per-class durations, whole-round durations and jar sizes are all replaced with this round's
  measured values (below).

### Behaviour

- **The default output now carries `source`** (this is an output change relative to 1.1.1, not a
  zero-cost addition): the three samples' JSON in `META` mode grows by **331 / 287 / 264** characters
  (docx 7 groups, pdf 6, pptx 6 — one 44-character `"source" : { "page" : N }` per group; the ledgers in
  **both docx and pdf** span two pages and each carries one extra 23-character `lastPage`, which is how docx
  reaches `7 × 44 + 23 = 331`). The three deltas are reproducible: pull the `1.1.1` main jar from Maven
  Central, run the same samples in `META` mode and you get **4409 / 6186 / 3262**.
- **Key order within a group is fixed**: `headers` → `rows` → `name` (written only when registered and
  matched) → `source` (written only when `source(true)`). `name` comes after `rows` because
  `Tables.newBlock` creates those two keys first, which is what keeps the order identical on the
  new-group path and on the multi-level-header merge path.
- **One name can match several groups** (when one registration matches several groups cut out of the same
  physical table); those groups then carry the *same* `name` and `tablesByName` returns them all at once.
  When two consecutive header rows both match registered groups, the **first** name wins.
- **Cross-format differences on multi-level headers (re-measured after the patch)**: the grouping algorithm
  first compares grid rows one by one, then merges consecutive header rows and compares the merged result
  once more. Therefore registering the **merged** row (`客户名称 / 合同金额 / 已回款 / 回款率 / 备注`) hits in
  **docx, pdf and pptx**, at the default tolerance 1 and at tolerance 0; with only the ledger registered the
  three samples give 13 / 31 / 7 rows. Registering the **upper** row instead
  (`客户名称 / 回款情况 / 回款情况 / 回款率 / 备注`) is where the formats differ, all measured: docx 13 rows
  (both ledger header rows carry `w:tblHeader`, so both count as headers and are merged into one group);
  pptx 8 rows (only the upper row matches the registration and is taken as the header; the lower row is no
  longer a header and falls into the data — measured as row 0, exactly the 8 rows the silent mode reads);
  pdf misses (its grid contains the merged row only, not the upper one). The most robust cross-format
  registration is still to register both spellings under one name (`TestSupport.namedBothHeaderRows()`):
  measured, the three samples still give 30 / 45 / 21 rows with unchanged columns.
  `TableByNameTest` and `PrintResultTest.printLedgerNameRemedies` pin all of these spellings.
- **One class of noise removed from the "no row matched this header set" warning**: when one registration
  under a given `name` matched, another registration under the *same* name that missed is **no longer
  reported** (`Tables.warnUnmatchedHeaders`, shared by all three formats). The remedy above requires
  registering both spellings under one name, and one of them is bound to miss in any given document, so
  reporting it was pure noise. Registrations without a name are unaffected — `failure part`, which does not
  exist in the pdf, is still reported
  (`TableByNameTest.aRegistrationCoveredByASameNameMatchIsNoLongerReportedOnStderr`).
- **Top-level key order** is `text` → `tables` → `images` by default; only `unmatchedHeaders(true)` adds a
  fourth key, placed last. With the switch off the result is byte-identical to before (on docx,
  "registered but unnamed" is byte-identical to silent); with it on the other three keys are unchanged.
  An unmatched registration without a name has only `headers`; one that had a name also carries `name`.

### Tests

- **183 → 242 cases** (+59), test classes **10 → 15**. Main changes: new `PdfBorderlessTableTest` 5,
  `TableByNameTest` 11, `PdfMergedCellTableTest` 5 and `PptParserTextTest` 2 (tests package-private
  `textLines` directly, because evaluation mode cannot read multi-paragraph PPT text); `CrossFormatTest`
  9→15, `OptionsTest` 10→14, `PrintResultTest` 5→17, `TablesTest` 31→34, `PptParserTest` 17→19,
  `WordParserTest` 31→32, `ImagesTest` 24→25, `PdfCrossPageTableTest` 3→4, `WordEdgeCasesTest` 11→14;
  `PdfParserTest` 17 and `ResultsTest` 28 kept their counts but gained assertions.
  Full suite this round on both JDK 8 and JDK 21: **242 cases, 0 failures, 0 errors, 1 skipped** (the skip is
  the `-Ddoc` entry point).
- **New `CrossFormatTest.aMissingFileThrowsFromAllThreeParsers`**: all three parsers hand the path straight to
  Aspose without wrapping it. Measured, the message is `Can't find file: <path>.` in every case, but the types
  differ: docx `java.io.FileNotFoundException`, pdf `com.aspose.pdf.internal.ms.System.l5if`, pptx
  `com.aspose.slides.exceptions.FileNotFoundException`; a single
  `catch (java.io.FileNotFoundException)` only covers docx. `PrintResultTest.printErrorPaths` prints six error
  paths verbatim (the other three are FILE mode without `imageDir`, where all three parsers throw
  `IllegalStateException`).
- **The print-only class grew to 17 cases and 18 reports** (`PrintResultTest`; `printAnyDocument` is skipped by
  default), in four categories: (1) three silent sample reports, docx "registered + DROP", three
  "registered + named" (7 reports); (2) four public-entry-point reports — `printPublicApiOverloads`,
  `printOptionDefaultsAndSetters`, `printSerializationAndImageExport`, `printNamedMatchesAcrossFormats`;
  (3) five header-rule and option-matrix reports — `printLedgerNameRemedies`, `printOptionMatrix`,
  `printAllModeJson`, `printUnmatchedHeaders`, `printHeaderAndFieldRules`; (4) two reports for images and
  error paths — `printImageModes` and `printErrorPaths`. Every report keeps only the
  `assertHeadersMatchRowKeys` baseline assertion, because broken output has no reference value.
- **New `printPublicApiOverloads`**: the three samples × the four `parse` overloads (path / stream ×
  options / keywords), compared pairwise — the path and stream JSON are character-for-character identical
  (all `true`). Keyword input is not an alias of options: on the pdf, `片区/区域/销售额/回款率/备注` yields
  6 groups and 47 rows (silent: 45 — the two extra rows come from tables the keywords missed, whose header
  row is then treated as data); docx and pptx match silent mode.
- **New `printOptionDefaultsAndSetters`**: defaults are `isSilent=true`, `fullRow=KEEP`, `images=META`,
  `source=true`, `unmatchedHeaders=false`; `headerTolerance` has no public getter, so it is verified by
  registering `说明` instead of `备注` — the default 1 hits, 0 misses and prints the unmatched warning;
  `headerGroups(List)` and repeated `headers(...)` produce character-identical JSON; `name(...)` binds only
  to the most recent registration.
- **New `printHeaderAndFieldRules`**: duplicate headers get a suffix from the **physical column index**
  (`销售额-5`, `回款率-6`, not "the second duplicate"), and blank headers become `列N`; a docx built with
  data `1 + 2` but a `SUM(ABOVE)` cache of `300.00` reads back as 300.00, and an empty cache as an empty
  string — proof that the parser reads the cache and never executes field code. The three samples' total
  rows are printed in section 4 (docx 2096.30 / 1796.10, pdf two empty cells, ppt 1586 / 1384); the
  differences come from the source files themselves.
- **New `printSerializationAndImageExport`**: the three `toJson` outputs are character-identical; in
  `BASE64` mode docx / pdf / pptx produce 899370 / 58076 / 898156 characters (lengths only); every file
  written by `writeImages` has a SHA-1 equal to its `contentId`; calling it on a `META` result only warns,
  creates no directory and writes no file.
- `printNamedMatchesAcrossFormats` prints, name by name, which groups were hit, each hit's source page and
  every one of its rows, the row count `rowsByName` flattens them to, and whether the flattened list and the
  groups concatenated in order are the very same objects (compared by reference, printed as `true` for all
  nine checks across the three samples — nothing is copied); it ends with three names that hit nothing
  (unregistered / empty string / null all give an empty list). `printLedgerNameRemedies` prints the four
  registration spellings for the two-row-header ledger in the docx side by side — all four hit in docx, all
  give 13 rows and empty stderr: ① and ② via merge-then-rematch, ③ by matching the upper row directly,
  ④ by matching the upper row and then covering it with the same-name rule. The differences are in the other
  formats: pdf has only the merged row (③ misses), while pptx has both rows and registering only the upper
  one lets the lower row fall into the data (③ becomes 8 rows). `printAllModeJson` prints the **full JSON** of
  the same three samples × seven option sets: one summary line per mode (groups / rows, top-level keys, stderr
  line count) followed by the JSON, and it also writes 21 UTF-8 files to
  `target/print/json/<sample>-<mode>.json` — the other reports only give JSON lengths, so this is the one to
  read when you need to check fields one by one. `printOptionMatrix` is a three-samples ×
  seven-options table (groups / rows, JSON length, `name`, unmatched count, stderr line count, top-level
  keys); `printUnmatchedHeaders` compares the switch off and on, plus a second registration of
  `故障部件` alone named `部件统计`; `printImageModes` shows the size, fields and output path of META / BYTES /
  BASE64 / FILE, plus a "META + imageDir" pass proving it only warns and writes no file; `printErrorPaths`
  prints the six error paths above. Each of the 18 reports also keeps a UTF-8 copy under `target/print/`
  (`mvn clean` removes it); the copies total **191,974** characters this round.
- **Fixed four doclint heading errors**: the javadoc bodies of `WordParser.Options.name(String)` and
  `unmatchedHeaders(boolean)` each contain two sub-headings. Member javadoc starts one level above `<h3>`,
  so the **first heading in each comment must be `<h3>` and later sub-headings `<h4>`**: all-`<h3>` is
  reported as "unexpected heading", while all-`<h4>` makes the first heading "out of sequence", and
  `javadoc -Xdoclint:all,-missing` on JDK 8 rejects both. With the first heading `<h3>` and the rest
  `<h4>`, doclint measured **0 errors**; builds and tests are unaffected.
- To assemble minimal pptx files, `TestSupport` gained a set of builders (`pptx()` / `pTbl` / `pTr` / `pTc` /
  `pTcSpan` / `pTcHMergeContinue` / `pTcVerticalStart` / `pTcVerticalContinue`): the master, layout and theme
  are copied from the sample in this repository and only `[Content_Types].xml`, `ppt/slides/`, `ppt/media/` and
  the `presentation.xml` relationships are rewritten. The two PPT fixes above could only be reproduced because
  of it — "a vertical merge spanning two rows" and "a blank row merged across the full width" do not occur in
  the sample deck.

### Build artifacts

- Local `mvn -B clean package` (JDK 8): **BUILD SUCCESS**, **02:10**,
  **242 cases, 0 failures, 0 errors,
  1 skipped**; main jar **60236 bytes / 30 entries / 17 classes**, `major version: 52`;
  `javac -Xlint:all -Xpkginfo:always` **0** warnings (17 classes including `package-info.class`);
  `javadoc -Xdoclint:all,-missing` **0** errors. `mvn -B test` without `-DargLine` runs 242 tests with 0 failures and 1 skipped as well; that
  flag only affects the console and report encoding.
  The same `clean package` on JDK 21: **02:17**, jar
  **60650 bytes**, same entry and class counts, still `major version: 52` — packaging the same source with
  a different JDK changes the size by a few hundred bytes, so do not use jar size as a checksum.
  These are the release builds under `1.2.0`; a quieter round of the same source measured 01:25 (JDK 8) and
  01:32 (JDK 21), and the difference is machine load, so durations remain order-of-magnitude only.
  For comparison: the 1.1.1 main jar on Maven Central is 56275 bytes, and the difference is entirely the
  source changes listed in this section.
- **An earlier JDK 21 `clean package` ended in BUILD FAILURE for a reason outside the code, and it is
  recorded here as it happened**: the suite then had 214 cases and all passed, but
  `maven-jar-plugin:3.2.2:jar` failed with
  `Error assembling JAR: cn\net\tianxi\document\Tables.class not found`. In `target/classes`, `Tables.class`
  alone carried the timestamp 08:37 while every other class read 08:35 — 08:37 is exactly when I restored
  `Tables.java` with `cp`, so the IDE recompiled that one file in the middle of the build and collided with
  Maven's jar phase. Re-running the identical command gave BUILD SUCCESS.
  **It is an environment race, not a defect**; turn off the IDE's automatic build before reproducing a build.
- **"the round is ~80 s slower than 1.1.1's (`01:04`), entirely because
  `CrossFormatTest.documentedJsonLengthsAreTheRealOnes` parses 12 more times (31.5 s → 76.5 / 72.4 s)" is
  retracted**: it put a round that was competing for CPU (the same command measured 02:15–02:37 at the time)
  next to an idle one. In a quiet environment the previous round, `clean package` measured 01:25 (JDK 8) and
  01:40 (JDK 21); alternating three runs with that assertion against three
  without it in the same JVM gives
  34.6 / 30.9 / 37.6 s against 29.8 / 35.9 / 38.2 s. Every duration in the documentation is therefore stated
  as a range with "read it as an order of magnitude". The only precisely attributable figure is the
  per-method **3.3–3.8 s (JDK 8, two runs) / 3.2 s (JDK 21)** in the surefire report
  (`target/surefire-reports/TEST-*.xml`; earlier rounds measured around 3.5 / 5.6 s).

### Documentation and release records (after `v1.1.1`)

- Added the English [README.en.md](README.en.md) and the GitHub Pages project home
  [docs/index.html](docs/index.html) (Pages enabled and verified reachable).
- README gained three sections comparing this approach with OCR, its benefits, and its use cases (§1.1–1.3);
  **no OCR accuracy or timing figures are given** — no OCR engine has been run in this repository.
- The §1.1 timing ranges now use only the two measured JDK 8 rounds; the JDK 21 figures are listed
  separately.
- The attribution wording now includes the Gitee mirror address (GitHub primary + Gitee mirror side by
  side).
- Recorded the 1.1.1 Maven Central release result and its verification (Deployment PUBLISHED, a fresh empty
  repository pulled it through `dependency:get`, sha256 identical in three places).
- Release coordinates fixed as `cn.net.tianxi:document-parser`; on the first upload the namespace
  validation had not finished, so one artifact was published under `io.github.hutianxi`. Central artifacts
  cannot be deleted or modified, so that coordinate stays in the central repository permanently (see
  README §17).
- The working tree additionally holds an Excel-parsing selection report,
  `docs/excel-parser-selection.md` (side-by-side measurements of POI / EasyExcel / FastExcel /
  Aspose.Cells; not committed, not part of the jar).
- `pom.xml` carries three more **uncommitted** release-metadata edits that will ship with the next
  version's pom — the 1.1.1 pom on Central still has the old values: `<description>` now names both
  repository addresses, `<developers><email>` is `mailtianxi@qq.com`, and `<scm><connection>` points at the
  Gitee mirror (the anonymous read-only URL) while `<developerConnection>` and `<url>` stay on the GitHub
  primary. The version element itself is still `1.1.1`, not bumped.
- Added this file and its Chinese counterpart.
- That round brought both READMEs, both changelogs and the project home in line with 232 cases / 14 classes /
  60099 bytes (JDK 8) and 60500 (JDK 21), the named lengths 4828 / 6561 / 3466, and the 13 print reports,
  and documented merge-then-rematch, `unmatchedHeaders` and the prose-above-table fix in the matching
  sections.
- **Wrap-up round (working tree, uncommitted)**: `PrintResultTest` now has **16 cases / 17 reports** (added
  the four-overload comparison, `Options` defaults and setters, header naming and field caches, and
  serialisation plus image export), and [DESIGN.md](DESIGN.md) was added (module responsibilities, data flow,
  header detection, per-format adapters, boundaries and trade-offs). Both READMEs, CONTRIBUTING, both
  changelogs and the project home now carry **236 cases / 14 classes / 17 reports / 85,948 characters /
  01:32–01:42 (JDK 8) and 01:55 (JDK 21)**.
- **Print-report round (working tree, uncommitted)**: `PrintResultTest` gained `printAllModeJson` (the full
  JSON of three samples × 7 option sets, plus 21 `.json` files), taking it to **17 cases / 18 reports**. Both
  READMEs, CONTRIBUTING, DESIGN, both changelogs and the project home now carry **237 cases / 14 classes /
  18 reports / 191,974 characters / 01:25 (JDK 8) and 01:40 (JDK 21)**.
- **JSON-to-DB example (working tree, uncommitted)**: added [docs/json-to-db.md](docs/json-to-db.md)
  (header-signature routing, complete `ValueType` / `ColumnSpec` / `TableSpec` / `JsonToDb` implementation,
  a docx/pdf/pptx mapping table, Spring/MyBatis and same-JVM shortcuts, checklist and known limits). The
  example is outside the build and compiles on JDK 8 with Jackson 2.13.5. Both README trees and matching
  sections, plus the homepage documentation card, now link it.

### Pre-release wrap-up (2026-10-23)

This round did two things: fix the three defects named in the previous round's checklist, each with a
regression test, and clear project-unrelated session artefacts out of the working tree. That sets the
scope of `1.2.0` to "the new APIs (`source` / `name` / `unmatchedHeaders` / `tablesByName` /
`rowsByName`) plus these three fixes". At release time `pom.xml` was bumped to `1.2.0`; the rest of this
section is unchanged.

**The three fixes** (root causes are detailed under "Fixed" above; each started as a failing test, then the
code changed):

- **PDF cross-page merge**: `mergeSplitTables` must replace `page`, `bottomGap` and `lastOnPage` with the
  tail's values when a fragment is appended. With `lastOnPage` in place the new sample moves from **1 group
  to 2**; the negative control (removing just that line) drops back to 1. Test
  `PdfCrossPageTableTest.theChainAlsoStopsWhenANonTableFlowFollowsTheFragment`.
- **Multi-paragraph PPT text**: splitting was extracted into package-private `textLines`, which handles
  `\r\n`, `\r` and `\n` and skips blank paragraphs; two `PptParserTextTest` unit tests pin it directly,
  because evaluation mode hides multi-paragraph text from end-to-end samples.
- **Word nesting depth**: the warning is printed once and only when a real ninth-level child table is
  skipped, and `walk` checks for children before depth. Measured: 8 levels → **8 groups, empty stderr**;
  9 levels → one warning with the first 8 groups intact. Two `WordEdgeCasesTest` cases.

**Tests**: **237 → 242 cases (+5)**, test classes **14 → 15**. `WordEdgeCasesTest` 12→14,
`PdfCrossPageTableTest` 3→4, new `PptParserTextTest` 2; the other 12 classes are unchanged. A final
`mvn -B clean package` on JDK 8 and on JDK 21 both report **242 cases, 0 failures, 0 errors, 1 skipped**
(the skipped one is the `-Ddoc` entry point). Per-class counts come from
`target/surefire-reports/TEST-*.xml`.

**Docs and comments**: the per-class counts in DESIGN.md §11 now match this round's measurement (adding
`PptParserTextTest` 2, `WordEdgeCasesTest` 14, `PdfCrossPageTableTest` 4); the `ImagesTest` count cited
in `Images.java` was corrected to the actual 25; the nesting-depth wording in `WordParser` and `package-info` matches the new behaviour (skip from
level 9, warn only when content really is skipped); and a comment in `PdfParser` that pointed at an old
test name was renamed with it. New javadoc subsection headings follow the JDK 8 doclint rule (the first
heading in a member comment must be `<h3>`, later ones `<h4>`), and `javadoc -Xdoclint:all,-missing`
reports **0 errors**.

**Cleanup**: `.qoder-credits/` (Qoder session artefacts, unrelated to the project) has been deleted, and
`.gitignore` now lists `.qoder-credits/`, `.aoci/`, `AGENTS.md`, `aoci*.txt`, `.gitattributes` and
`.mcp.json*`. The last five stay local only at the author's request and do not enter the repository; all
of them show as `!!` in `git status --ignored`.

**Final gates**: under version `1.2.0`, JDK 8 `mvn -B clean package` is **BUILD SUCCESS** at **02:10**, main
jar **60,236 bytes / 30 entries / 17 classes**, `major version: 52`; the same command on JDK 21 gives
**242 tests, 1 skipped** at **02:17**, jar **60,650 bytes**, same entries and class count;
`javac -Xlint:all -Xpkginfo:always` reports **0 warnings** and `javadoc -Xdoclint:all,-missing`
**0 errors**.

**Release result**: commit `bfbdec9` and tag `v1.2.0` are pushed to GitHub and Gitee; Maven Central Portal
Deployment `923fda1a-a25c-46f3-b83d-a387966a5def` reached **PUBLISHED**, and the four artifacts on
`repo1.maven.org` — jar / sources / javadoc / pom — are **60,236 / 81,011 / 88,998 / 11,118** bytes with
sha256 values byte-for-byte identical to the local build. A brand-new empty local repository ran
`dependency:get` with **BUILD SUCCESS**, fetching a 60,236-byte jar with the same sha256. GitHub Release
`v1.2.0` was published on 2026-10-24 00:18 (+0800), and the `size` / `digest` of its three attachments match
Central. Full record in README §17.

---

## 1.1.1 — 2026-10-22

First publicly released version (tag `v1.1.1`, committed 2026-10-22 02:23:09 +0800).

- **One structure, three formats**: `WordParser` / `PdfParser` / `PptParser` have matching method
  signatures and return `Map<String, Object>` with the fixed keys `text` / `tables` / `images`; changing
  format means changing the class name.
- **Header grouping** (`Tables`, the only implementation, shared by all three formats): six-way priority
  (registered header set → keywords → the format's own header-row flag → the row right after a group row →
  the first row declared by the table style → the silent-mode first-row fallback), two-row header merging,
  three destinations for full-width merged rows (`FullRow.KEEP` / `FIRST` / `DROP`), merged cells always
  split and back-filled, column-name completion (empty header → `Column N`, duplicate name →
  `name-columnIndex`). Contract: `headers` equals the keys of every row, one by one.
- **PDF**: `TableAbsorber` reconstructs ruled tables geometrically, a long table split across pages is
  merged back into one (four continuation tests), and rows that a vertical merge chopped into pieces are
  clustered back by rectangle adjacency.
- **Images**: four output modes (`META` default / `BASE64` / `BYTES` / `FILE`), SHA-1 content fingerprint
  `contentId`, duplicated embedded images embedded once, linked images recorded by path without any
  network access, `writeImages` export.
- **Text cleaning** (`Results`, the only implementation): field codes dropped and cached field results
  kept, soft breaks and page breaks turned into spaces, control characters removed, Aspose
  evaluation-watermark lines filtered; a license is registered once per product per process.
- **Safe handling of untrusted documents**: only the file-name part is kept before writing, duplicates are
  suffixed inside the directory, links are not followed, recursion depth is capped.
- **Tests and artifact**: 183 test cases in 10 test classes (1 `-Ddoc` entry point skipped in normal runs),
  main jar 56275 bytes, Java 8 bytecode.
- **License and publishing**: Apache License 2.0 with a `NOTICE` attribution; the three Aspose artifacts
  are used under the Aspose EULA and their binaries are not redistributed; published to Maven Central as
  `cn.net.tianxi:document-parser:1.1.1` (sources jar 64964 bytes, javadoc jar 141571 bytes), GitHub Release
  `v1.1.1` carrying all three packages, GitHub primary and Gitee mirror kept in sync, `main` protected on
  both.
