# document-parser

[![License](https://img.shields.io/badge/License-Apache%202.0-blue.svg)](LICENSE)
[![Java](https://img.shields.io/badge/JDK-8%2B-orange.svg)](#3-tech-stack-and-requirements)
[![Maven Central](https://img.shields.io/badge/Maven%20Central-1.2.1-green.svg)](#17-build-and-publishing)
[![Tests](https://img.shields.io/badge/tests-241%20passed-brightgreen.svg)](#13-testing-and-quality)

[中文](README.md) ｜ **English**

> A document-parsing library: it turns **Word / PDF / PPT** into one and the same result shape —
> `text`, `tables`, `images`. All three formats return the identical structure, so switching format
> means switching a class name and nothing else.

| Project | Value |
| --- | --- |
| Maven coordinates | `cn.net.tianxi:document-parser:1.2.1` (currently on Central, published 2026-10-24; dependency upgrade over 1.2.0 only) |
| Java package | `cn.net.tianxi.document` |
| Language level | Java 8 (Aspose artifacts use the `jdk16` classifier) |
| Underlying technology | Aspose.Words / Aspose.PDF / Aspose.Slides 23.12; API and behaviour follow [docs.aspose.com](https://docs.aspose.com) |
| Artifact | Plain library jar — the released 1.2.1 artifact is **60777 bytes, 30 entries and 17 classes** on JDK 8 (`mvn -B clean package`; 61190 bytes for the same source on JDK 21), including `META-INF/LICENSE` and `META-INF/NOTICE`. 1.2.0 was 60236 bytes and 1.1.1 was 56275 bytes; release records are in [CHANGELOG.en.md](CHANGELOG.en.md) and §17. No Spring container, no CLI entry point |
| Tests | JUnit 5, 242 tests across 15 test classes: 241 pass, 1 `-Ddoc` entry point is skipped by default |
| Changelog | [CHANGELOG.en.md](CHANGELOG.en.md) (中文: [CHANGELOG.md](CHANGELOG.md)) — what changed in each version, with the measured numbers |
| License | [Apache License 2.0](LICENSE): free to use, modify, use commercially and redistribute, **provided the original author attribution is kept** (see [NOTICE](NOTICE)). Pull requests welcome (see [CONTRIBUTING.md](CONTRIBUTING.md)) |
| Repositories | GitHub [`hutianxi/document-parser`](https://github.com/hutianxi/document-parser) (primary) ｜ Gitee [`mailtianxi/document-parser`](https://gitee.com/mailtianxi/document-parser) (China mirror) |

The detailed design is documented in three layers: the overall design and implementation rationale
(module responsibilities, data flow, header detection, per-format adapters, boundaries and trade-offs)
lives in [DESIGN.md](DESIGN.md); package-level design (result shape, header-detection priority,
licensing and evaluation mode, security boundaries) lives in
[`package-info.java`](src/main/java/cn/net/tianxi/document/package-info.java); class and method detail
lives in the javadoc of each type and shows up on hover in your IDE. This README is the index over all
three layers plus all measured data.

---

## Contents

1. [Overview](#1-overview)
2. [Features](#2-features)
3. [Tech stack and requirements](#3-tech-stack-and-requirements)
4. [Project layout](#4-project-layout)
5. [Quick start](#5-quick-start)
6. [API](#6-api)
7. [Result structure](#7-result-structure)
8. [Options](#8-options)
9. [Capability boundaries: what is read, what is not](#9-capability-boundaries-what-is-read-what-is-not)
10. [Test samples](#10-test-samples)
11. [Measured baselines](#11-measured-baselines)
12. [Open-source license, Aspose licensing and evaluation mode](#12-open-source-license-aspose-licensing-and-evaluation-mode)
13. [Testing and quality](#13-testing-and-quality)
14. [FAQ](#14-faq)
15. [Known limitations and uncovered branches](#15-known-limitations-and-uncovered-branches)
16. [Maintainer guide](#16-maintainer-guide)
17. [Build and publishing](#17-build-and-publishing)
18. [Appendix: fix log (historical numbers)](#18-appendix-fix-log-historical-numbers)

---

## 1. Overview

A very common business requirement is "extract the text and tables from a customer-uploaded contract,
report or slide deck and load them into a database". The three formats have completely different
underlying models: Word has real table objects and merge flags, PDF has only drawn line segments and
scattered text, PPT is a per-slide shape tree. If each format gets its own parser and its own result
shape, the caller has to write three ingestion paths, and every rule change has to be made in three places.

This module's approach is: **keep the format differences inside the parsers and expose exactly one
result structure.**

- Three entry classes `WordParser` / `PdfParser` / `PptParser`, with one-to-one matching method
  signatures, returning a `Map<String, Object>` whose keys are always `text` / `tables` / `images`;
- Tables are normalised into "headers + row objects": `headers` equals the keys of every row
  item by item, so you can use `headers` directly as the column list when creating a table or mapping DTO fields;
- Merged cells are always split and back-filled, so the number of keys per row is constant;
- Header grouping, image collection and text cleaning each have exactly one implementation
  (`Tables` / `Images` / `Results`) shared by all three formats.

Things the design deliberately **does not** do: no OCR (a scan without a text layer simply cannot be
read), no layout re-flow, no `updateFields` to recompute fields, no persistence, no network calls.
It is a library of pure static methods — no bootstrap class, no CLI, no global mutable state
(apart from the Aspose license registration).

### 1.1 How this differs from OCR: it reads the text layer, it does not guess pixels

These are two entirely different technical routes, so let's state the difference before comparing them:

- **This project** follows "parse the format's object model": Aspose reads docx's
  `w:tbl`/`w:tc`/`w:sdt` nodes, PDF's text objects and line-segment coordinates, and pptx's shape tree.
  The character data is already stored in the file; the parser just pulls it out along the structure.
- **OCR** follows "image recognition": render the page to an image, classify glyph shapes pixel by pixel,
  then *infer* characters, lines and table rules. The text is inferred, not read.

| Dimension | document-parser (reads the text layer) | OCR (image recognition) |
| --- | --- | --- |
| Input precondition | Electronically native documents: docx saved by Word, pdf exported by a program, pptx saved by PowerPoint | Any page that can be turned into an image, **including scans and photos** |
| Where the text comes from | The character data inside the file, read verbatim | Inferred from pixels, so there is a misrecognition rate |
| Numbers and money amounts | Character-for-character identical to the original: the `合计` (total) row reads the cached field values `2096.30` / `1796.10`, asserted cell by cell by the suite (§11) | `0/O`, `1/l`, `8/B` and decimal points are easily confused glyph pairs; amounts usually need human review |
| Table structure | The **real** structure: merge flags (`gridSpan`/`vMerge`/`rowSpan`), two-row headers, tables nested inside a cell are all detected, split and back-filled | Reconstructed from rule detection; borderless tables, tables split across pages and merged cells easily lose rows or columns |
| Reading order | Document-tree order (paragraph → table → shape), matching the author's writing order | Inferred from layout coordinates; multi-column pages and floating text boxes easily cross over |
| Result stability | Repeated parsing of the same file yields **exactly the same result**, so it can be asserted in CI | Results change with engine version, resolution and binarisation parameters, which makes regression testing hard |
| Deployment shape | One jar plus three Aspose dependencies, all in-process, no network (linked images are recorded as metadata only, never downloaded) | Usually requires an extra recognition engine/model, or uploading documents to a third-party cloud service |
| Scans / image-only PDFs | **Cannot be read** (no text layer) — an explicit boundary documented in §9 | The only viable route |

> **Honest boundary**: this section gives no accuracy or timing figures for OCR. No OCR engine has ever
> been run in this repository, so no reproducible measurement exists to quote. To compare the two routes,
> run the same batch of files through both and compare. The right-hand column above describes the
> **inherent characteristics** of the OCR route, not the measured results of any product.

The per-file timings for this project *are* measured (three samples, evaluation mode, no `.lic`; a throwaway
probe test class ran two consecutive rounds of `mvn test` under JDK 8 and the table shows the range across
those two rounds — the probe was deleted after the run, as always, so it is not in the repository and
reproducing these numbers means writing one yourself):

| Sample | Size | First call (includes Aspose initialisation) | Subsequent single parse |
| --- | --- | --- | --- |
| docx | 675208 bytes | 2058–2396 ms | 83–227 ms |
| pdf (3 pages) | 184977 bytes | 4830–5363 ms | 671–1268 ms |
| pptx | 679289 bytes | 934–1006 ms | 61–75 ms |

Environment: JDK `1.8.0_191`, Maven 3.8.8, Windows 11. A second round under JDK 21.0.1 gave the same
order of magnitude (first call 3313 / 7360 / 2322 ms, subsequent 126–192 / 938–1655 / 86–99 ms). Three notes:

- **The slow first call is a fixed cost** — each Aspose product loads its classes and initialises its
  license once per process. A long-running service (Spring bean, thread pool) operates in the
  "subsequent" column.
- **The docx `source` key pays for layout construction**: building `LayoutCollector` itself costs 0 ms, but
  the first `getStartPageIndex` call lays out the whole document — a probe measured the split (load 16–24 ms,
  collector 0 ms, page lookups for 5 tables / 17 calls 98–131 ms). That is why the "subsequent" column moved
  from 23–30 ms in the previous release to 83–227 ms. If you do not need the page numbers, pass
  `Options.source(false)` and the Word side never creates a `LayoutCollector` at all: the same sample goes back
  to **11–35 ms**. That figure comes from a throwaway probe that **alternated** on and off inside one JVM
  (6 parses each, two rounds, deleted after the run) so the later mode could not pick up free JIT warm-up —
  74–254 ms on versus 11–35 ms off, and in both rounds the slowest "off" parse beat the fastest "on" parse.
  PDF and PPT already know their page number while parsing, so the switch changes nothing measurable there
  (same rounds: pdf 681–960 vs 696–984 ms, pptx 46–86 vs 44–71 ms); it is the key that disappears, and it
  disappears for all three formats at once (see §8).
- **The 700–1200 ms for PDF** goes into `TableAbsorber`'s geometric clustering, cross-page table merging and
  the **second, flow-engine pass over every page**. It correlates with page count and line-segment count, not
  with file size — the 184977-byte pdf is an order of magnitude slower than the 679289-byte pptx. The extra
  pass is ~1.3x the previous version: measured in one JVM on the 3-page sample, line engine alone 478 ms,
  flow engine alone 615 ms, both 922 ms.

### 1.2 What this buys you

- **One ingestion path for three formats**: all three return the same `text`/`tables`/`images` structure;
  switching format means switching a class name. Adding a new format means writing one more parser —
  the ingestion code does not change.
- **`headers` can be used as a column list directly**: for every table group, `headers` equals the keys
  of every row item by item, and merged cells are always split and back-filled, so the key count per row
  is constant — across 12 parses (3 samples × 4 option sets) the contract held with **0** mismatches
  (`CrossFormatTest.contractHoldsForEverySampleAndMode`). Creating tables, mapping DTOs and writing
  `INSERT` statements need no special handling.
- **CI-able**: the result is deterministic, so every number in README §11 is a test assertion —
  break it and the build goes red immediately.
- **No index counting when loading rows**: name the groups you intend to persist with
  `Options.name("Quarterly detail")` while registering their headers, and afterwards
  `WordParser.rowsByName(result, "Quarterly detail")` hands you exactly those rows; `source.page` tells you
  which page of the original they came from when something has to be checked by eye.
- **Stateless and concurrency-safe**: all three parsers are static methods with no global mutable state
  (apart from the Aspose license registration); a fully configured `Options` can be reused.
- **Nothing leaves the process, nothing touches the network**: purely local parsing; linked images are
  recorded as metadata only and never downloaded, and file names are reduced to their last path segment
  with collision suffixes before anything is written — usable for contracts and financial reports that
  must not leave the intranet.
- **Images handled on demand**: four output modes (metadata / base64 / write-to-disk-while-parsing /
  raw bytes) plus a SHA-1 content fingerprint `contentId`, so an image appearing in several documents
  is stored once.

### 1.3 Application scenarios

Scenarios it fits directly:

| Scenario | How to use it |
| --- | --- |
| Loading contracts / reports / ledgers into a database | One parse of the customer-uploaded docx yields body text + structured tables + attached images; `headers` maps straight onto database columns |
| RAG and knowledge-base chunking | Chunk `text` semantically, store `tables` separately with their structure intact (flattening a table into prose loses the column relationships), and write `images` to disk with `ImageMode.FILE` before uploading them to object storage |
| BI / report aggregation | Same-template Word / PDF reports submitted by branch units; register the headers once and column names become fixed, then batch-load into the warehouse |
| Finance and audit trails | Amounts are the values stored in the file (cached field results), not recognition output, so they are traceable |
| Cross-format consistency checks | Compare the docx / pdf / pptx versions of the same report cell by cell (this is exactly what `CrossFormatTest` does) |
| Archiving slide-deck content | Body text, tables and images extracted in one pass; speaker notes are not read (§9) |
| Intake pipeline with mixed scans and native files | Probe with this module first: if both `text` and `tables` come back empty there is no text layer, so fall back to OCR; otherwise use this module. **This triage rule is a recommendation — the repository contains no scanned sample, so it has not been measured** |

Scenarios that do not fit and need a different approach (each has a matching note in §9):

- **Scans, photos, image-only PDFs**: no text layer, OCR required;
- **Data inside charts, SmartArt text, content of embedded Excel/Word (OLE) objects**: Aspose's shape
  dispatch identifies the type, but this module does not read their content;
- **Headers, footers, footnotes, endnotes, comments**: not read on the Word path;
- **PPT speaker notes, masters, layout slides**: not read;
- **Anything that needs the layout preserved** (page position, fonts, coordinates, faithful HTML/PDF
  conversion): this module does not re-flow layout, it only gives content.

## 2. Features

- **One structure, three formats**: the result shapes of doc/docx, pdf and ppt/pptx are identical;
  `CrossFormatTest` verifies cell by cell that the core table of the same-source samples is identical
  across the three formats.
- **Header grouping**: however many header groups a physical table has, `tables` has that many elements,
  and each element carries only the data rows belonging to it — no cross-contamination. Two-row headers
  are supported (when the lower row leaves a column blank, the upper row's name is inherited).
- **Six kinds of header evidence** (priority top to bottom): a registered complete header set →
  keywords → the format's own title-row flag → the row immediately following a full-row merge →
  the first row declared by the table style → the first-row fallback in silent mode.
- **Merged cells split apart**: positions covered by horizontal merges (`gridSpan` / `colSpan` /
  rectangular column spans) and vertical merges (`vMerge` / `rowSpan` / rectangular row spans) are all
  back-filled with the value of the origin cell.
- **Group sub-headings are controllable**: rows merged across the full width (e.g. "FY2025 (closed)")
  can be kept in their own group, moved to the table's first group, or dropped, via `FullRow`.
- **PDF tables split across pages are merged back into one**: geometric clustering plus six continuation
  criteria (adjacent pages, same column count, aligned left edges, the earlier fragment is the last block on
  its page and the later one is the first block on its page, no header row in the later fragment; in silent
  mode an extra "one touches the page bottom, the next touches the page top"); a table spanning three or more
  pages keeps merging all the way down.
- **Borderless PDF tables are recovered by a second pass**: besides the line-segment engine, each page is also
  run through the text-flow engine, and only its "every row has the same cell count and it overlaps no ruled
  table" findings are taken (measured: a borderless sample went from 0 groups to 3; the fully ruled repository
  sample gained none).
- **Every table group carries its source location**: `source.page` (docx page / pdf page / pptx slide index,
  1-based), plus `lastPage` when the table spans a page break, so a suspicious row can be checked against the
  original document by flipping to that page. `Options.source(false)` removes the key for all three formats at
  once and spares docx the layout pass (§8).
- **Registered header groups can be named**: `Options.name("Quarterly detail")` binds to the group registered
  by the `headers(...)` call written just before it, and a group carries a `name` key only once that
  registration actually matched; `WordParser.tablesByName(result, name)` / `rowsByName(result, name)` fetch
  those groups (and their rows, flattened in `tables` order) by name. Unnamed groups, and groups whose header
  came from keywords, `w:tblHeader` or the first-row fallback, keep exactly the keys they had before.
- **Four image output modes**: metadata only (default), inline base64, write to disk while parsing,
  raw bytes; the SHA-1 content fingerprint `contentId` supports deduplication and caller-side caching.
- **Text cleaning**: field codes are dropped and cached field results kept, soft breaks and page breaks
  become spaces, control characters are removed, evaluation watermarks are filtered.
- **Safe handling of external documents**: only the last path segment is used before writing a file,
  name collisions inside a directory get a numeric suffix, linked images are neither read nor fetched,
  and recursion has a depth limit.
- **Stateless and concurrency-safe**: all three parsers are static methods; a configured `Options` can be
  reused (do not share it across threads while you are still configuring it).

## 3. Tech stack and requirements

| Dependency | Version | Purpose |
| --- | --- | --- |
| `com.aspose:aspose-words:jdk16` | 23.12 | doc/docx parsing |
| `com.aspose:aspose-pdf:jdk16` | 23.12 | pdf parsing (`TableAbsorber` / `ParagraphAbsorber` / `ImagePlacementAbsorber`) |
| `com.aspose:aspose-slides:jdk16` | 23.12 | ppt/pptx parsing |
| `com.fasterxml.jackson.core:jackson-databind` | 2.21.8 (from 1.2.1) | `toJson` serialisation |
| `org.junit.jupiter:junit-jupiter` (test scope) | managed by parent pom 2.7.18, JUnit 5.8.2 | unit tests |

Requirements:

- JDK 8 or newer (the Aspose artifacts use the `jdk16` classifier and work fine on Java 8): the full test
  suite was run under both JDK `1.8.0_191` and JDK `21.0.1`, and both **ran 242 tests with 0 failures and
  1 skipped**, and the class files produced by both are `major version: 52` (Java 8 bytecode) — because
  `java.version=1.8` makes maven-compiler-plugin use `-source/-target 1.8`, compiling with a newer JDK
  does **not** raise the bytecode version;
- Maven 3.8.8 (the version measured here), with access to `https://repository.aspose.com/repo/`;
- On headless servers add `-Djava.awt.headless=true` (Aspose uses AWT image classes internally);
- The parent pom is `spring-boot-starter-parent` **only to borrow its dependency version management**;
  this module has no Spring bootstrap class and does not configure `spring-boot-maven-plugin`.
  Tests only pull `junit-jupiter` and deliberately **do not use `spring-boot-starter-test`**, keeping
  components such as Spring 5.3.x and logback 1.2.x — which have no security-fix releases on those lines —
  off the test classpath.

## 4. Project layout

```
document-parser/
├── pom.xml                                  # plain library jar; parent pom for version management only; release profile for Central
├── README.md                                # Chinese
├── README.en.md                             # English (this file)
├── LICENSE                                  # full Apache License 2.0 text
├── NOTICE                                   # original author attribution + third-party components (Aspose commercial license)
├── CONTRIBUTING.md                          # contribution guide: environment, conventions, PR flow, license red lines
├── CHANGELOG.md / CHANGELOG.en.md            # release log: what changed in each version, with measured numbers
├── DESIGN.md                                # implementation rationale: module responsibilities, data flow, header detection, per-format adapters, boundaries
├── docs/index.html                          # GitHub Pages project homepage
├── docs/excel-parser-selection.md           # Excel parsing technology selection (measured POI / EasyExcel / FastExcel / Aspose.Cells comparison, Chinese; not shipped in the jar)
├── docs/json-to-db.md                       # JSON-to-database example: header-signature routing, type conversion, atomic per-table writes (Chinese; not shipped in the jar)
├── docs/feature-guide.md                    # feature guide: per-method and per-parameter reference, return fields, full sample data, acceptance steps (Chinese; not shipped in the jar)
├── .gitignore                               # target/, *.lic (private credentials), GPG signing output, IDEA project files, Office temp files ~$*
└── src
    ├── main/java/cn/net/tianxi/document
    │   ├── package-info.java                # package-level design notes (result shape, boundaries, licensing, security)
    │   ├── WordParser.java                  # entry point + Options/FullRow/ImageMode definitions + Word grid reading
    │   ├── PdfParser.java                   # entry point + PDF geometric reconstruction (clustering, cross-page merging)
    │   ├── PptParser.java                   # entry point + shape dispatch (table/text box/picture/group shape)
    │   ├── Tables.java                      # the header-grouping algorithm (the only copy, shared by all three formats)
    │   ├── Images.java                      # the image collection (the only copy: numbering, fingerprinting, dedup, four modes)
    │   └── Results.java                     # shared utilities: JSON, text cleaning, watermark filtering, licensing, file-name sanitising
    └── test
        ├── java/cn/net/tianxi/document       # 15 test classes, 242 tests (see "Testing and quality")
        │   ├── TestSupport.java             # sample paths, assertion helpers, parts for building minimal docx files on the fly
        │   ├── WordParserTest.java / PdfParserTest.java / PptParserTest.java   # end-to-end assertions on the three samples
        │   ├── PptParserTextTest.java       # text-frame line splitting: three line endings, blank paragraphs dropped (no Aspose)
        │   ├── CrossFormatTest.java         # cross-format consistency
        │   ├── TablesTest.java / ImagesTest.java / OptionsTest.java / ResultsTest.java  # unit tests of the shared components
        │   ├── WordEdgeCasesTest.java       # minimal docx built on the fly to cover edge structures
        │   ├── PdfBorderlessTableTest.java  # borderless PDF built on the fly: the flow engine as a supplement, plus its known false positive
        │   ├── PdfCrossPageTableTest.java   # three-page PDF built on the fly: the cross-page merge chain and where it must break
        │   ├── PdfMergedCellTableTest.java  # docx re-exported as PDF: merged cells, two-level headers, registration after the fact
        │   ├── TableByNameTest.java         # naming registered header groups: which groups carry `name`, and looking groups/rows up by name
        │   └── PrintResultTest.java         # asserts almost nothing, just prints parse results to the console (includes the -Ddoc entry)
        └── resources/{word,pdf,ppt}          # the three business samples (not packaged into the artifact jar)
```

`Tables` / `Images` / `Results` are all **package-private** and are not exposed outside the parsers;
the public surface is the three parser classes only.

## 5. Quick start

> This README describes **1.2.1** (published to Central on 2026-10-24). Compared with 1.2.0, 1.2.1 only raises Jackson to 2.21.8,
> fixing CVE-2026-110879 (an incomplete fix for CVE-2026-54512: generic arguments and array element types could
> still bypass the PolymorphicTypeValidator); the API, result shape and parsing behaviour are unchanged.
> Compared with 1.1.1, 1.2.0 adds the `source` key, the `name` key and the `tablesByName` / `rowsByName` entry
> points, the top-level `unmatchedHeaders` switch, and borderless-table recovery in PDF. The three main keys
> `text` / `tables` / `images`, the `headers`-vs-row-keys contract and the four image modes are unchanged from
> 1.1.1; the difference is that each group there has no `source` / `name` key and that borderless PDF tables are
> not recovered in 1.1.1. None of the new content exists while you stay on 1.1.1. See the 1.2.1 and 1.2.0
> sections of [CHANGELOG.en.md](CHANGELOG.en.md) and §17.

### 5.1 Add the dependency

```xml
<dependency>
    <groupId>cn.net.tianxi</groupId>
    <artifactId>document-parser</artifactId>
    <version>1.2.1</version>
</dependency>
```

> 1.2.1 was published to Central on 2026-10-24; if your mirror has not synced it yet, you can use `1.2.0`
> for now — the only difference is that it still ships Jackson 2.21.7.

The three Aspose artifacts come from `https://repository.aspose.com/repo/`, so consumers must declare
that repository in their own `pom.xml` (this project already declares it).

### 5.2 Minimal example

```java
Map<String, Object> result = WordParser.parse("D:/a.docx", "sales");   // PdfParser / PptParser: same name, same parameters
String json = WordParser.toJson(result);
```

The second parameter is an optional list of **header keywords**: a row is treated as a header row only if
every non-empty cell in it contains one of the keywords. If you pass no keywords and register no headers,
the parser runs in "silent mode", where the first row of each table may be used as a header as a fallback.

### 5.3 Common examples

The header and keyword literals below are written in English so they read naturally; **the three bundled
samples are Chinese documents**, so if you want to run these snippets against
`src/test/resources/word/sales-report.docx` (or the pdf / pptx of the same source), use the literals the
parser actually reads: `片区` district, `区域` region, `销售额` sales, `回款率` collection rate, `备注` remark,
`客户名称` customer name, `合同金额` contract amount, `已回款` collected, `季度明细` quarterly detail,
`回款台账` collection ledger, keyword `销售`. `PrintResultTest` prints the real headers of every group, which
is the quickest way to copy them.

```java
// Register complete headers (most stable across pages and formats) and drop group sub-heading rows
WordParser.Options options = WordParser.options()
        .headers("district", "region", "sales", "collection rate", "sales", "collection rate", "remark")  // one group; you may register several
        .name("Quarterly detail")              // binds to the headers(...) call written just before it
        .headers("customer", "contract amount", "collected", "collection rate", "remark")
        .name("Collection ledger")
        .headerTolerance(1)                          // allow 1 mismatching column (0 = exact per column), default 1
        .fullRow(WordParser.FullRow.DROP)            // drop rows merged across the full width
        .keywords("sales");                          // optional loose fallback

Map<String, Object> result = PdfParser.parse("D:/a.pdf", options);
```

```java
// Persisting: only the tables you load need a name; fetch them by name instead of counting tables indices
List<Map<String, Object>> tables = WordParser.tablesByName(result, "Quarterly detail");
List<String> columns = (List<String>) tables.get(0).get("headers");   // use this as the column list
for (Map<String, String> row : WordParser.rowsByName(result, "Quarterly detail")) {
    // the keys of each row equal that group's headers item by item; same-named groups are flattened in order
}
// a named group carries one extra key: { "headers": [...], "rows": [...], "name": "Quarterly detail", "source": {"page":1} }
```

```java
// Write images to disk while parsing and keep only the paths in the result
// (recommended for image-heavy documents: the bytes never stay in memory)
Map<String, Object> result = PptParser.parse("D:/a.pptx", WordParser.options()
        .images(WordParser.ImageMode.FILE)
        .imageDir(new File("D:/out/images")));       // FILE mode requires imageDir, otherwise IllegalStateException
```

```java
// Stream-based parsing (upload scenarios): the caller closes the stream, parse releases the document handle
try (InputStream in = upload.getInputStream()) {
    Map<String, Object> result = WordParser.parse(in, WordParser.options());
}
```

### 5.4 Run one test to see the actual parse result

If you would rather look at what a document parses into than write code, run the bundled
`PrintResultTest` (**asserts nothing, just prints**):

```
mvn -B test -Dtest=PrintResultTest                                              # all 18 reports
mvn -B test -Dtest=PrintResultTest#printAllSamplesWithRegisteredNames            # what the registered-and-named option set produces on all three samples
mvn -B test -Dtest=PrintResultTest#printAllModeJson                              # the full JSON of all three formats under 7 option sets, plus 21 .json files
mvn -B test -Dtest=PrintResultTest#printNamedMatchesAcrossFormats                # which groups each name hits and which rows you get back
mvn -B test -Dtest=PrintResultTest#printPublicApiOverloads                       # the four parse overloads side by side (path/stream × options/keywords)
mvn -B test -Dtest=PrintResultTest#printAnyDocument -Ddoc="D:/your-file.docx"    # your own document
```

The console shows the body text paragraph by paragraph, each table group's registered name (only when one
was given) and source page, its headers and every data row, the image metadata and the complete JSON,
and also writes a UTF-8 copy into `target/print/`.
Sample output and details are in [13. Testing and quality](#13-testing-and-quality).

## 6. API

The public methods of the three parsers correspond exactly (`XxxParser` = `WordParser` / `PdfParser` / `PptParser`):

| Method | Description |
| --- | --- |
| `static Map<String,Object> parse(String filePath, String... headerKeywords)` | Parse by path, header keywords only |
| `static Map<String,Object> parse(InputStream in, String... headerKeywords)` | Parse from a stream; the caller closes the stream |
| `static Map<String,Object> parse(String filePath, WordParser.Options options)` | Parse by path with full options |
| `static Map<String,Object> parse(InputStream in, WordParser.Options options)` | Parse from a stream with full options |
| `static String toJson(Object value)` | Serialise the result to indented JSON |
| `static void writeImages(Map<String,Object> result, File directory)` | Write the images of a `BASE64`-mode result to files |
| `static WordParser.Options options()` (`WordParser` only) | Create an options object, shared by all three formats |
| `static List<Map<String,Object>> tablesByName(Map<String,Object> result, String name)` (`WordParser` only) | Fetch the table groups whose `name` equals `name`, in `tables` order and as the very objects inside the result; an empty list when nothing matches |
| `static List<Map<String,String>> rowsByName(Map<String,Object> result, String name)` (`WordParser` only) | The data rows of those groups flattened in the same order, ready to insert row by row; an empty list when nothing matches |

Neither lookup parses anything or throws: a `null` result, a `null` name or a blank name all give an empty
list. The query string is trimmed at both ends; inner spaces and case are significant (a name is something
the caller chose, so the lookup does not guess the way header matching does).

Chained methods on `WordParser.Options`: `headers(String...)`, `name(String)`, `headerGroups(List<List<String>>)`,
`headerTolerance(int)`, `keywords(String...)`, `fullRow(FullRow)`, `images(ImageMode)`, `imageDir(File)`,
`source(boolean)`.
Two enums: `WordParser.FullRow` (`KEEP` / `FIRST` / `DROP`) and `WordParser.ImageMode`
(`BASE64` / `FILE` / `BYTES` / `META`).

`parse` / `toJson` / `writeImages` all `throws Exception`: they parse external documents, and corrupt
files, encrypted files and a format that does not match the extension all surface as exceptions for the
caller to handle (`options()` and the chained methods do not throw).

## 7. Result structure

Below is the **measured output** for `src/test/resources/word/sales-report.docx` (default
`ImageMode.META`, so images carry metadata only):

```json
{
  "text": "2025 East China sales performance report\nReport pages: 3\nArchive no.:\n…\nData as of 2025-12-31; the finance system's closed figures prevail.",
  "tables": [
    { "headers": ["district", "region", "sales", "collection rate", "sales-5", "collection rate-6", "remark"],
      "rows": [ { "district": "East China", "region": "Shanghai", "sales": "1280", "collection rate": "92%",
                  "sales-5": "1465", "collection rate-6": "95%", "remark": "target exceeded" } ],
      "source": { "page": 1 } }
  ],
  "images": [
    { "fileName": "image-1.png", "type": "PNG", "width": 1024, "height": 768,
      "size": 670939, "contentId": "487e915db2c53907e8c1551aeaffea235423210c" }
  ]
}
```

(The JSON above is an English rendering of the real output; the sample document itself is in Chinese,
so the actual keys and values are Chinese strings.)

### 7.1 `text`

The full body text, paragraphs joined with `\n`. **Text inside tables is not repeated here** (the single
exception is described in the result contract in the maintainer guide). All text cleaning happens in one
place, `Results.clean`:

- Word wraps a field in three control characters as "field code + field result"; the **code is dropped
  and only the cached result is kept** (otherwise a cell would show `PAGE`, and it might even be mistaken
  for a header). Paragraph 2 of the sample measures `Report pages: 3`;
- A field that was never computed disappears entirely. Paragraph 3 of the sample measures `Archive no.:`
  (the `DOCPROPERTY` was never evaluated, so its result is empty);
- Soft breaks (Shift+Enter) and page breaks become a **space** rather than being dropped; paragraph 4 of
  the sample measures `…figures are tax-inclusive sales. Data source: …`, and the space after the full
  stop is the trace of the soft break;
- Invisible control characters such as the cell terminator are dropped, and runs of spaces collapse to one.

### 7.2 `tables`

However many header groups a physical table has, `tables` has that many elements; each element's `rows`
contains only the rows belonging to the header above them.

- **Header-detection priority** (any hit counts as a header): a registered complete header set →
  keywords → the format's own title-row flag (Word's `w:tblHeader`) → **the row immediately following a
  full-row merge** → the first row declared by the table style (`tblLook firstRow` / `getFirstRow()`) →
  the first-row fallback in silent mode. Two easy traps: the row right after a group sub-heading is
  treated as the header of a new section (putting data rows directly after a group row creates an extra
  group); and once any header evidence is registered, the "first-row fallback" is switched off, so a
  table that matches nothing keeps its first row as a data row and its columns are named `Column N`
  (PDF has no title-row flag to rely on, so this is most visible there — see
  `PdfParserTest.registeringHeaderEvidenceTurnsOffTheFirstRowFallback`).
- **Column naming**: an empty header becomes `Column N`; columns that end up with the same name after a
  merge is flattened get a `name-columnNumber` suffix (in the sample, `sales` / `collection rate` each
  appear twice → `sales-5` / `collection rate-6`); when a registered header matches, the registered names
  are used, which is why the keys can be identical across the three formats.
- **Contract**: `headers` equals the keys of every row **item by item** (same order, same count), and the
  row objects are `LinkedHashMap`, so you can read values in column order. `CrossFormatTest` checked this
  across 3 samples × 4 option sets = 12 parses, with 0 mismatches.
- **Source location `source`**: each group carries one more key — `{ "page": startPage }`, plus `"lastPage"`
  when the table spans a page break (in the samples it is the collection ledger, one in docx and one in pdf,
  both `{"page": 1, "lastPage": 2}`). All three formats share the key name and the field: docx takes the page number
  from `LayoutCollector`, pdf takes the page number, pptx takes the slide index — all 1-based, so you can flip
  straight to that page to check the original. Groups cut out of one physical table all carry the same page
  range (groups are split by header, not by page). When a format cannot give a position the key is omitted
  entirely (measured: every group in all three samples has one, so that branch has never been taken).
  If you do not want the key at all, `Options.source(false)` switches it off for **all three formats together**:
  each group then carries just `headers` and `rows` with names and content otherwise unchanged, and the Word
  side skips `LayoutCollector` entirely (timings in §8; guarded by
  `CrossFormatTest.sourceSwitchOffDropsTheKeyFromAllThreeFormats`).
- **Registered name `name`**: written only when all three things hold — the group's headers were registered,
  the registration carried a name, and that registration actually matched in this document — and it is
  written after `rows`:

  ```json
  { "headers": ["district", "region", "sales", "collection rate", "sales-5", "collection rate-6", "remark"],
    "rows": [ … ],
    "name": "Quarterly detail",
    "source": { "page": 1 } }
  ```

  The default silent output therefore contains no `name` at all (the JSON at the top of §7 is that case —
  4740 characters in its real, Chinese form).
  With the sample registration that names three of five groups (`TestSupport.named()`), the measured JSON
  lengths go from 4740 → **4828** (docx, groups 1, 2, 6, 7 named), 6473 → **6561** (pdf, groups 1, 2, 5, 6) and
  3378 → **3466** (pptx, groups 1, 2, 4, 5). The docx ledger group is named too: after merging a multi-level
  header the grouping algorithm compares the *merged* row against the registrations once more, so registering
  that merged row hits (all five registrations hit on docx; stderr is empty — the pdf/pptx differences are in
  §8). Fetching is
  `WordParser.tablesByName(result, "Quarterly detail")` (groups in `tables` order) or
  `WordParser.rowsByName(result, "Quarterly detail")` (their rows flattened, one list to insert). `name` is
  not a primary key: the two header groups cut out of one physical table carry the *same* name (measured:
  docx groups 6 and 7). For database import, route on the headers themselves (for example
  `ObjectMapper.writeValueAsString(headers)`) rather than on `name`; a complete worked example — header
  signature routing, type conversion (empty strings, `86%`, thousands separators) and atomic per-table
  writes — is in [docs/json-to-db.md](docs/json-to-db.md) (Chinese).
- **Per-feature reference**: how to pass every parameter of every method, what each returned field means,
  and the complete parsed data of all three samples are collected in
  [docs/feature-guide.md](docs/feature-guide.md) (Chinese).
- **Tables nested inside a Word cell**: the child table does not go into the host cell's text (that cell
  stays empty); it appears as an **extra element** in `tables`, so neither side loses its row/column
  structure. Measured on the sample: the host table is `["ticket type","ticket count","handling notes"]`
  (the cell in question is `["hardware fault","42",""]`) and the child table is a separate group
  `["faulty part","ticket count"]` with 3 rows, whose `source` is the same page 2 as its host.

### 7.3 `images`

The metadata (`fileName`, `type`, `width`, `height`, `size`, `contentId`) is **always present in every
mode**; where the content goes is decided by `ImageMode`:

| Mode | Extra field in the result | Use case |
| --- | --- | --- |
| `META` (default) | none | You only need text and tables in the database and do not want to pay for image size and memory |
| `BASE64` | `"base64": "data:image/png;base64,…"` | Drop straight into a front-end `<img src>` |
| `BYTES` | `"bytes": byte[]` (Jackson encodes it as base64 without the `data:` prefix) | Java-side storage, feeding an OCR pipeline |
| `FILE` | `"path"`: absolute path under `imageDir` | Image-heavy documents that should be written to disk while parsing |

A few points:

- `contentId` is the **SHA-1 hex fingerprint of the image bytes**, used only to tell "are these the same
  bytes the same image" and as a caller-side cache key; it is not a storage location, not a URL, and
  carries no tamper-proofing duty. The docx and pptx samples share this value (the same PNG bytes) while
  the pdf differs (Word re-encoded it as JPEG on export) — parsing reports the bytes actually in the
  file and does not reconstruct anything.
- In inline modes (`BASE64`/`BYTES`), images whose content repeats are **not inlined twice**: the
  duplicate carries only an index pointing at the earlier entry (`"duplicateOf": 0`) and no content
  field, while its `contentId` is still its own fingerprint. `META` mode does not mark duplicates; in
  `FILE` mode each duplicate gets its own file, with a `-2` suffix on collision. Each of the three
  samples has exactly one image, so deduplication is verified by `ImagesTest`, which feeds bytes directly.
- For **linked images** (Word "Link to File", PPT "Link only") the bytes are not inside the document
  package, so only the `link` path is reported; this module never reads that file and never goes online,
  hence such entries have no `size`/`contentId`.
- `writeImages(result, dir)` only recognises the `base64` field in the result, so it is **effective for
  `BASE64` mode only**; calling it in the default (`META`) mode exports nothing and prints a hint on
  stderr about which mode to use. Likewise `imageDir(File)` only takes effect in `FILE` mode, and a
  mismatched combination also prints a hint.

Keep the size in mind (measured on the three samples, see "Measured baselines"): that docx has one image,
and the whole META-mode result is 4740 characters versus 899370 characters in BASE64 mode — the 670939-byte
PNG encodes to 894610 characters, which is 99.5% of it.

## 8. Options

`WordParser.Options` is a mutable object shared by all three formats; chained calls take effect immediately.

| Option | Default | Meaning and measured effect |
| --- | --- | --- |
| `headers(String...)` | not registered | Register one complete header group (call repeatedly to register several). On a match, the registered names are used. Measured on the samples: after registering 5 real header groups, the docx and pdf JSON is **byte-for-byte identical** to silent mode, while pptx goes from 22 rows to 21 (the lower header row is no longer treated as data). Trailing empty columns of the document row are **trimmed before comparison**, so a table whose last column has no title (the docx target table) matches when you register only the first 3 names, and the 4th column is still filled as `Column 4`; do not let your own registered group end with an empty name, or it will never match |
| `headerGroups(List<List<String>>)` | empty | Batch registration, equivalent to calling `headers(...)` several times. Note that it *is* registering group by group internally, so a `name(...)` written right after a batch binds to the **last group of that batch**; to name several groups use the `headers(...).name(...)` spelling |
| `name(String)` | no name | Names **the group registered by the `headers(...)` call written just before it**; on a match that group carries a `name` key (see §7). Practical recipe: run `PrintResultTest`, copy the `headers` of the groups you want to persist into `.headers(...).name("...")`, then read them back with `WordParser.rowsByName(result, "...")`. Four things to keep in mind: ① only groups that matched a registration get the key — anything detected via keywords, `w:tblHeader`, the group-row rule or the first-row fallback does not; ② one name can match several groups (they all come back from `tablesByName`, in `tables` order), so it is not a primary key; ③ matching follows `headerTolerance`, so a loosened tolerance can also make a name land on the wrong row — tighten it to `headerTolerance(0)`; ④ for two-row headers the three formats expose different grid rows, so one registration may match in pdf/pptx and miss in docx — register both spellings under **the same name** (`TestSupport.namedBothHeaderRows()`; measured: still 30/45/21 rows). Leading and trailing spaces in the name are trimmed, an empty or null name means "unnamed", calling it before any registration prints a hint on stderr and is ignored, and renaming a group is a contract change for whatever reads it by name |
| `headerTolerance(int)` | 1 | How many columns may mismatch during header comparison; 0 = exact per column. Whatever the value, at least one column must match exactly and the column count must be equal; negatives are clamped to 0 |
| `keywords(String...)` | none | Loose fallback: a row is a header row only if **every non-empty cell contains one of the keywords**. Null and blank entries are ignored |
| `fullRow(FullRow)` | `KEEP` | Where rows merged across the full width (group sub-headings) go: `KEEP` puts them in the nearest header group above, `FIRST` puts them in the physical table's first group, `DROP` discards them. Measured on docx: KEEP 7 groups / 30 rows, FIRST the same counts with only ownership changed, DROP 6 groups / 28 rows |
| `images(ImageMode)` | `META` | See 7.3 |
| `imageDir(File)` | null | Only effective in `FILE` mode; `FILE` without it throws `IllegalStateException` when the first image is reached, and setting it without `FILE` prints a hint on stderr |
| `source(boolean)` | `true` | Whether every group carries `source` (see §7). **All three formats switch together.** Off, the docx side never creates a `LayoutCollector`, so no layout pass is paid for: the same sample measured 74–254 ms on versus **11–35 ms off** with an on/off-alternating throwaway probe in one JVM. PDF and PPT get their page number for free, so the switch only decides whether the key is written (same rounds: pdf 681–960 vs 696–984 ms, pptx 46–86 vs 44–71 ms — overlapping ranges) |

**Silent mode** (`isSilent()`: neither headers nor keywords given) turns on the "first-row fallback" and
is good for a first look at an unknown document; register headers when you want stable ingestion.
`Options` itself is not synchronised, so do not share it across threads while configuring it; once
configured it can be used concurrently for many parses (`Images` is created per parse, options are
read-only), and when `imageDir` points at the same directory, colliding image names get a numeric suffix
via `Images.uniqueFile` instead of overwriting each other.

## 9. Capability boundaries: what is read, what is not

| Format | Read into the result | Not read (known out of scope) |
| --- | --- | --- |
| Word | Paragraphs, tables and images in each section's Body; text-box text; paragraphs and tables wrapped in block-level content controls (`w:sdt`); list items (bullets do not enter the text); the **cached result** of fields (`SUM(ABOVE)` and `PAGE` both read the cached value) | Headers, footers, footnotes, endnotes, comments; field codes are always dropped and no layout computation is done, so a field that was never computed is empty; an unfilled content control does not emit the placeholder text `Click here to enter text.` (4 spellings measured, see below) |
| PDF | Tables with drawn borders (geometric reconstruction), text-layer paragraphs, bitmaps; a long table split across pages is merged back into one (criteria: adjacent pages + same column count + aligned left edges + the earlier fragment is the last block on its page and the later one is the first block on its page + no header row found in the later fragment; in silent mode, where no header evidence is available, the geometric evidence "one fragment touches the page bottom, the next touches the page top" is required as well); rows cut into several blocks by a vertical merge are clustered back into the same table by rectangle adjacency; **besides the ruled tables, the text-flow engine is run as a second pass** that only accepts borderless candidates which are "every row has the same cell count and overlap no ruled table" (measured: a borderless sample went from 0 groups to 3, and `sales-report.pdf` gained none — see below) | Borderless tables are only recovered when **every row has the same number of cells**; ragged ones (a full-width title row on top, for instance) still land in `text`; the false positive that comes with the second pass: prose aligned with tab stops reads as a table (measured, see below); scans / image-only PDFs need OCR; when no header evidence is registered and the row cut by the page break is unusually tall (its bottom edge more than 1/5 of the page height above the page bottom), continuation is not recognised and the table splits into two groups, each with its own header |
| PPT | Shape text, tables and images on the body slides (group shapes expanded recursively, up to 32 levels; `VideoFrame` extends `PictureFrame`, so video frames take the picture path and what lands in `images` is their poster frame) | Speaker notes / masters / layout-slide text, text inside SmartArt graphics, chart data, content of embedded OLE objects (Excel sheets / Word documents), shape fill backgrounds, the video content itself |

Several boundaries confirmed by measurement:

- **Unfilled content controls emit no placeholder text**: for 4 spellings (`<w:p/>` content, empty
  `<w:sdtContent/>`, no `w:sdtPr`, with `w:placeholderText`) the measured `text` contains only the
  paragraph outside the control and `Click here to enter text.` never appears once — that hint lives in
  `w:sdtPr/w:placeholderText`, which is Word's presentation layer, not document content
  (`WordEdgeCasesTest.unfilledContentControlContributesNothing`).
- **Exporting to PDF loses field results**: this PDF was exported from an **extended** version of the
  report (30 customers; the source file is not in the repository) using **real Microsoft Word 2021**
  (the file's `Producer` says so), not converted by Aspose, so it carries no evaluation watermark. In that
  source docx the two total cells were cached `SUM(ABOVE)` results (a throwaway probe confirmed
  Aspose.Words read `2582.30` / `2226.60`), and Word recomputed the field on export, so in the PDF those
  two cells hold a single space → parsing faithfully reads empty strings
  (`["Total","","","86%",""]`, `PdfParserTest.exportedPdfLosesTheCachedSumFieldResult`). The docx sample
  in the repository has a 12-customer ledger whose total row reads the cached values
  `2096.30` / `1796.10` / `86%`. **Use the docx for money amounts.**
- **Borderless PDF tables: recovered, but only as a supplement**. A throwaway probe (deleted after the run;
  the same layout is now pinned by `PdfBorderlessTableTest`) compared the two `TableAbsorber` engines on three
  synthetic files: on a Word-exported page holding a borderless 5×3 table next to a ruled 5×3 table, the default
  (line-segment) engine found 1 table and the flow engine found 3; on a page of scattered text plus a borderless
  table the default found 0 and the flow engine found 3. Flipping the switch globally, however, corrupted the
  fully ruled repository sample — 8 blocks became 5 tables with merged rows shifted one column left, an empty
  phantom column in the target table and the nested table's row in the wrong place (2 of 35 and 12 of 38 cell
  rectangles reported as invalid). So the flow engine never replaces the line engine; it only contributes
  candidates that (a) overlap no ruled table and (b) have ≥2 rows, ≥2 columns and the same cell count in every
  row. On the repository sample those two filters reject all 5 flow candidates, so the result is unchanged;
  the price is that on a page mixing a ruled and a borderless table only the borderless one can be added.
  Known false positive, written into an assertion: four tab-aligned rows of prose satisfy both filters and move
  from `text` into `tables` (`PdfBorderlessTableTest.tabAlignedParagraphsCurrentlyCountAsATable`).
- **For PPT two-row headers, register the headers**: in silent mode the lower header row is treated as
  data — measured 8 rows with column names
  `["customer","collection status","collection status-3","collection rate","remark"]`; after registering
  headers it correctly merges into 7 rows.
- **PPT shape dispatch**: all eight classes `Table` / `GroupShape` / `Chart` / `OleObjectFrame` /
  `SmartArtShape` / `PictureFrame` / `VideoFrame` / `AutoShape` were run through `isAssignableFrom`;
  only `AutoShape` is an `IAutoShape`, the four branches do not overlap and their order does not affect
  the result; `Chart`, `OleObjectFrame` and `SmartArtShape` implement none of the interfaces, so those
  shapes are silently skipped — this is where "chart data / OLE content / SmartArt text cannot be read"
  comes from.

## 10. Test samples

The three samples are the same "2025 East China sales performance report", under
`src/test/resources/{word,pdf,ppt}/`, and are **not packaged into the artifact jar**:

| Sample | Size | How it was produced | Content |
| --- | --- | --- | --- |
| `word/sales-report.docx` | 675208 bytes | Hand-assembled OOXML zip (5 parts; `word/document.xml` is 33216 bytes plus `word/media/image1.png`) | 13 body paragraphs + 4 tables + one 1024×768 PNG dashboard screenshot |
| `pdf/sales-report.pdf` | 184977 bytes, 3 pages | Exported from the same-source extended report by **real Microsoft Word 2021** (the file's `Producer`), not converted by Aspose, hence no evaluation watermark | The ledger was extended to 30 customers so that the long table breaks between pages 1 and 2 |
| `ppt/sales-report.pptx` | 679289 bytes | Hand-assembled OOXML zip (16 parts, 3 slides); merge properties live on `<a:tc>`: the origin cell is `<a:tc gridSpan rowSpan>`, covered cells are `<a:tc hMerge="1">` / `vMerge` | 9 text boxes (3 of them inside one group shape) + 4 tables + the same PNG |

**Do not generate samples with Aspose itself** — without a license it truncates content during save, so
expected values cannot be pinned down. Every piece of text in the pptx sample is kept to 5 characters or
fewer precisely because in evaluation mode Aspose.Slides returns only the first paragraph of a text frame,
truncated to 5 characters; `PptParserTest.everyTextFrameInTheSampleStaysWithinTheEvaluationLimit` guards
this. If you put long text into a sample, every expected value has to be re-measured in a licensed
environment.

Each of the four tables stresses one structure (same source in all three formats):

1. **Quarterly sales detail**: single-row header, two pairs of same-named columns (`sales` /
   `collection rate` each appear twice), the `district` column vertically merged;
2. **Customer collection ledger**: two-row header with the lower row blank (`collection status` spans two
   columns and the column name is inherited from the upper row), and a `Total` row that is a
   `SUM(ABOVE)` field;
3. **After-sales ticket statistics**: in Word, the "hardware fault" cell contains a further two-column
   table (`faulty part` / `ticket count`), which becomes its own group; after exporting to PDF or
   building the PPT this nested table is gone, so registering it produces an "unmatched" hint on stderr;
4. **Annual targets by district**: rows merged across the full width (`FY2025 (closed)`,
   `FY2026 (budget issued)`), a fourth column with no title (filled as `Column 4`), and two header groups
   in one physical table.

## 11. Measured baselines

Environment: Maven 3.8.8, Windows 11, **no `.lic` at all**, so all three parsers run in evaluation mode;
the full suite was run under JDK `1.8.0_191` and JDK `21.0.1`, and both ran 242 tests with 0 failures and
1 skipped. Every number below is an assertion value in the current tests and can be reproduced with
`mvn -B clean test` (whole-round duration of the 1.2.1 dependency-upgrade round: `clean package` — compile, test
and jar — took **01:48** on JDK 8 and **02:11** on JDK 21; the 1.2.0 release round measured 02:10 / 02:17. It is
242 tests with 0 failures and 1 skipped, with or without `-DargLine="-Dfile.encoding=UTF-8"`, which only affects the Chinese encoding in the
console and the reports. The jar size depends on the JDK: **60777** bytes on JDK 8, **61190** on JDK 21;
the 1.2.0 release round produced 60236 / 60650).

**Do not treat the whole-round duration as a performance figure.** The 1.2.1 dependency-upgrade round measured
01:48 and 02:11, the 1.2.0 release round 02:10 and 02:17, while an idle round of the same code measured 01:25 and
01:32; even earlier the same command reached
**02:15–02:37** when another process was competing for CPU — nearly double. The only cost that can be attributed precisely is the per-method
timing in the surefire report: `CrossFormatTest.documentedJsonLengthsAreTheRealOnes`, which turns the
**12 JSON lengths** in this section (silent, registered-without-names, registered-with-names, BASE64 — four
settings × three formats) into assertions and therefore parses **12 more times**, measured **3.9 s on JDK 21 in
this round** (2.7–3.8 s in earlier JDK 8 runs; from `target/surefire-reports/TEST-*.xml`). An earlier revision claimed that this one assertion
made the round 80 s slower and the class go from 31.5 s to 76.5 s; that comparison put a contended round next
to an idle one and **does not reproduce** — running the class three times with the assertion and three times
without it, alternating, shows the difference is entirely round-to-round noise. The claim has been retracted;
the payoff is unchanged:
every length quoted in the documentation can now be falsified by `mvn test`. Per-file parse timings are in
[§1.1](#11-how-this-differs-from-ocr-it-reads-the-text-layer-it-does-not-guess-pixels).

| Sample | `text` | Header groups | Data rows | Images | META result length | BASE64 result length |
| --- | --- | --- | --- | --- | --- | --- |
| docx | 13 paragraphs | 7 | 30 | 1 PNG 1024×768, 670939 bytes | 4740 chars | 899370 chars |
| pdf | empty string (all body text is inside tables) | 6 | 45 | 1 JPEG 699×524, 38668 bytes | 6473 chars | 58076 chars |
| pptx | 9 lines (each ≤5 characters) | 6 | 22 | 1 PNG 1024×768, 670939 bytes | 3526 chars | 898156 chars |

(The lengths grew by 331 / 287 / 264 characters versus the previous release, all of it the new `source` key:
docx has 7 groups, pdf and pptx 6 each, and every group carries a `"source" : { "page" : N }` — 44 characters;
**the ledgers in both docx and pdf span a page break** and so carry one extra `lastPage` each, 23 more
characters, which gives `7 × 44 + 23 = 331` for docx. The three deltas are reproducible: pull the `1.1.1` main
jar from Maven Central, run the same samples in META mode and you get **4409 / 6186 / 3262**; subtract those from
the table above and you get 331 / 287 / 264.)

Group and row counts for the same samples under different options (also assertion values):

| Options | docx | pdf | pptx |
| --- | --- | --- | --- |
| Silent (default, `FullRow.KEEP`) | 7 groups / 30 rows | 6 groups / 45 rows | 6 groups / 22 rows |
| 5 real header groups registered | 7 groups / 30 rows, JSON byte-identical to silent | 6 groups / 45 rows, JSON byte-identical to silent | 6 groups / **21** rows (the lower header row is no longer data) |
| Registered headers + `DROP` | 6 groups / 28 rows | 5 groups / 43 rows | 5 groups / 19 rows |
| `FIRST` (nothing registered) | 7 groups / 30 rows, only the group rows move to the table's first group | 6 groups / 45 rows, same | 6 groups / 22 rows, same |
| 5 groups registered, 3 of them named (`TestSupport.named()`) | 7 groups / 30 rows, groups 1, 2, 6, 7 carry `name`, 4740 → **4828** chars | 6 groups / 45 rows, groups 1, 2, 5, 6 carry `name`, 6473 → **6561** chars | 6 groups / 21 rows, groups 1, 2, 4, 5 carry `name`, 3378 → **3466** chars |

The core table (quarterly sales detail) of the three samples is **identical cell by cell** across the three
formats and in both silent and registered modes: column names
`["district","region","sales","collection rate","sales-5","collection rate-6","remark"]` with 4 data rows
matching exactly (`CrossFormatTest.coreTableIsIdenticalAcrossThreeFormats`). The ledger row counts differ
between samples (docx 13 / pdf 31 / pptx 8) because the samples were trimmed to each format's readable
boundary when they were built — not because parsing lost rows.

### Behaviours covered by the samples (each guarded by a test)

| Behaviour | Measured result | Test |
| --- | --- | --- |
| Same-named columns get a column-number suffix | The second occurrence of `sales` / `collection rate` → `sales-5` / `collection rate-6`, with `headers` equal to the row keys item by item | `WordParserTest.sameNamedColumnsGetColumnNumberSuffix` |
| Vertical merges back-filled | The two rows of the `district` column have no content of their own and read `East China` / `South East China`; all 4 data rows are present and none is mistaken for a group row | `WordParserTest.verticallyMergedColumnIsFilledFromItsOriginRow`, `WordEdgeCasesTest.verticallyMergedDataRowIsNotAMistakenSeparator`, `PptParserTest.verticallyMergedDataRowIsNotAMistakenSeparator` |
| Two-row headers merged | Blank columns in the lower row inherit the upper name → `["customer","contract amount","collected","collection rate","remark"]` | `WordParserTest.twoLevelHeaderInheritsUpperLevelColumnNames`, `TablesTest.consecutiveHeaderRowsMergeIntoOneMultiLevelHeader` |
| Field results | The `Total` row reads Word's cached `SUM(ABOVE)` results `2096.30` / `1796.10` / `86%`; in the body the cached value `3` of the `PAGE` field is kept while the field code is dropped; a never-computed `DOCPROPERTY` disappears entirely, leaving `Archive no.:` | `WordParserTest.sumFieldResultIsReadFromTheCachedValue`, `WordEdgeCasesTest.calculatedFieldResultSurvivesInsideACell` |
| Soft breaks | Shift+Enter becomes a space: `…tax-inclusive sales. Data source: …` (the longest paragraph, 62 characters, is read in full — evaluation mode does not truncate Words' body text) | `WordParserTest.textJoinsSoftBreakLinesWithASpace` |
| Block-level content controls | A paragraph wrapped in a control enters `text` (`Prepared by: Sales Management Dept  Reviewed by: CFO`); a control wrapped in a `w:tc` is also seen through by `Grid.readRow`'s shallow paragraph search (`["alpha in control","beta"]`) | `WordParserTest.textReadsParagraphInsideBlockContentControl`, `WordEdgeCasesTest.paragraphInsideCellLevelContentControlIsRead` |
| Unfilled controls emit no placeholder | 4 spellings measured: `text` contains only the paragraph outside the control, and `Click here to enter text.` never appears | `WordEdgeCasesTest.unfilledContentControlContributesNothing` |
| Table nested in a cell | `tables` 1 → 2 groups: the host cell is empty (`["hardware fault","42",""]`) and the child table is its own group `["faulty part","ticket count"]` with 3 rows | `WordParserTest.nestedTableBecomesItsOwnGroupAndLeavesTheHostCellClean` |
| Three destinations for group rows | KEEP stays in its own group, FIRST moves to the table's first group (counts unchanged, only ownership), DROP discards the row entirely (docx 7→6 groups, 30→28 rows); **the row immediately after a group row is treated as a header** | Six cases in `TablesTest` (`separatorRow*` / `fullRow*` / `rowAfterASeparatorRowIsTreatedAsAHeader`) plus one per parser |
| A group row before any header | Starts a new group with columns filled as `Column 1..Column 4`, and the group text is not lost (docx group 4, pdf group 3, pptx group 2 are all of this kind) | `TablesTest.separatorRowBeforeAnyHeaderStartsANumberedGroup` plus one per parser |
| A blank row merged across the full width is not a group row | On a 4-row 2-column table assembled on the fly (header / `A B` / blank merged row / `C D`): 1 group, column names `colA` / `colB`, 3 data rows, and the blank row kept as one row of empty values. Before the fix the same file measured 2 groups, the second named after the data `C` / `D` with 0 data rows (the `C D` row was lost). One and the same rule in all three formats | `WordEdgeCasesTest.blankFullWidthMergedRowIsNotASeparator`, `PptParserTest.blankFullWidthMergedRowIsNotASeparator` (pdf always had it, see `PdfParser.toGrid`) |
| PDF table split across pages | The 31-row ledger merges back into one group: page 1 is drawn down to `Taizhou Pharma City` (row 20), page 2 continues from `Huzhou Building Materials` (row 21), and `Total` is last | `PdfParserTest.ledgerSplitAcrossTwoPagesIsMergedBackIntoOneGroup`, `CrossFormatTest.onlyPdfNeedsTheCrossPageMerge` |
| PDF vertical merges cut into blocks | The two rows of the core table and the "hardware fault 12" row of the ticket table were cut into independent small blocks by the absorber; after clustering they are still 4 rows | `PdfParserTest.coreTableRowsCutByVerticalMergeAreClusteredBack`, `PdfParserTest.ticketTableKeepsTheRowThatWasCutIntoItsOwnBlock` |
| Registering headers switches off the first-row fallback | PDF has no `tblHeader` flag to rely on: in silent mode the first row is a header, but once header evidence is registered, the first row of a non-matching table is a data row and columns are filled as `Column N` (the core table has 5 rows = header row + 4 data rows) | `PdfParserTest.registeringHeaderEvidenceTurnsOffTheFirstRowFallback`, `TablesTest.firstRowStaysADataRowOnceHeaderEvidenceIsRegistered` |
| Registered names override names read from the document | Default tolerance is 1 mismatching column: registering `remark` as `notes` still matches, and the result uses the registered `notes` (PDF cell text gains spaces from line wrapping, so comparison ignores spaces) | `WordParserTest.registeredHeadersOverrideTheNamesReadFromTheDocument`, `TablesTest.registeredHeadersOverrideTheNamesReadFromTheGrid` |
| Unmatched registered headers are reported | stderr shows `header set did not match any row` plus that group's column names; on PDF/PPT the reported group is `faulty part` (that nested table does not exist in those formats), while Word reports **nothing** (all five `named()` registrations hit and stderr is empty). **A registration covered by a same-name match is no longer reported** — once both spellings of a two-row header are registered under one `name`, the one that is bound to miss stops making noise; registrations without a name are unaffected | One `unmatchedHeaderGroupIsReportedOnStderr` per parser, `TableByNameTest.aRegistrationCoveredByASameNameMatchIsNoLongerReportedOnStderr` |
| Every group carries a `source` page | docx: core table page 1, ledger `{"page":1,"lastPage":2}`, the other 5 groups page 2 (the nested child table shares its host's page); pdf: first two groups page 1 (the ledger runs onto page 2), tickets page 2, last three page 3; pptx: `page` is the slide index 1／2／3. Groups cut from one physical table share the same page range | `CrossFormatTest.sourceRecordsThePageEachBlockStartsOn` |
| `source(false)` switches the key off for all three formats | Off, an unnamed group carries just `headers` and `rows` (a named group keeps `name`); the unnamed key set is asserted to be exactly those two, and column names and cell content come out identical to the default; docx therefore never builds a `LayoutCollector` (on/off alternating in one JVM: 11–35 ms off vs 74–254 ms on) | `CrossFormatTest.sourceSwitchOffDropsTheKeyFromAllThreeFormats`, `OptionsTest.sourceIsWrittenByDefaultAndCanBeSwitchedOff` |
| A registered name is written as `name` | Only the groups whose registration carried a name get the key, and the key order is fixed as `headers` → `rows` → `name` → `source` (with `source(false)` the key set is `[headers, rows, name]`). Groups cut from one physical table carry the **same** name (measured: docx groups 6 and 7), and `rowsByName` flattens their 5 rows in group order, returning the very row objects inside the result. `tablesByName` / `rowsByName` return an empty list for an unknown, blank or null name, and trim the query's outer spaces | `TableByNameTest` (11 tests), `TablesTest.aRegisteredNameIsWrittenOnlyOnTheGroupsThatMatched`, `OptionsTest.nameBindsToTheGroupRegisteredJustBeforeIt` |
| Which spelling of a multi-level header is registered (cross-format difference) | Registering the *merged* row (`customer / contract amount / collected / collection rate / remark`): all three formats hit, even at tolerance 0, with ledger rows docx 13 / pdf 31 / pptx 7. Registering the *upper* row: docx hits (13 rows — both header rows carry `w:tblHeader`, so they still merge), pptx hits but its lower row leaks into the data (7 → 8 rows), and pdf misses (its grid only offers the merged row; stderr reports it once). For one registration shared across formats, register both spellings under the same name: all three samples carry the `name` and the group/row counts stay unchanged (30 / 45 / 21) | `TableByNameTest.aMultiLevelHeaderHitsByItsMergedRowInAllThreeFormats`, `TableByNameTest.aRegistrationCoveredByASameNameMatchIsNoLongerReportedOnStderr` |
| Borderless PDF tables are recovered | Measured with a throwaway probe (deleted after the run; `PdfBorderlessTableTest` now pins the same layout): a Word-exported page with a borderless 5×3 table next to a ruled one yielded 1 table with the default engine and 3 with the flow engine; **switching globally** corrupted the fully ruled repository sample (8 blocks into 5 tables, merged rows shifted left, a phantom empty column, the nested row in the wrong place — 2 of 35 and 12 of 38 cell rectangles invalid), so only candidates that are uniform-row and overlap-free are accepted, which rejects all 5 flow candidates on the repository sample and leaves its result unchanged | `PdfBorderlessTableTest.borderlessTableIsRecoveredAsAThreeColumnGrid`, `ruledTableOnTheSamePageIsUnaffected`, `nothingElseBecomesATable` |
| Tab-aligned prose counts as a table | The known false positive of that second pass: 4 rows × 3 columns of body text laid out with tab stops satisfies both filters, so that text moves from `text` into `tables`. The assertion exists precisely to pin the behaviour — change it by changing the test | `PdfBorderlessTableTest.tabAlignedParagraphsCurrentlyCountAsATable` |
| The PPT two-row-header trade-off | Silent: 8 rows with column names `["customer","collection status","collection status-3","collection rate","remark"]`; with registered headers: 7 rows — this is the reason for "register headers for PPT" | `PptParserTest.silentModeCannotMergeTheTwoLevelHeader`, `registeredHeadersMergeTheTwoLevelHeader` |
| PPT group shapes | The three text boxes `data source` / `sales system` / `finance reconciliation` sit inside one group shape and are only reachable by recursing into it | `PptParserTest.recursesIntoGroupShapes` |
| Four image modes | META gives metadata only; BASE64 gives a data URI (decoding it back yields 670939 bytes, identical to the original image); BYTES gives the raw bytes; FILE writes to disk and gives an absolute path, adding exactly one file to the directory; FILE without `imageDir` throws `IllegalStateException`; META with `imageDir` prints a hint on stderr and adds no new file | The 7 image cases in `WordParserTest` + 25 in `ImagesTest` |
| Recursion depth limits | 12 nested tables → `table nesting exceeds` on stderr, the outer table still parses normally; 20 nested controls → `node nesting exceeds`, that whole branch is skipped | `WordEdgeCasesTest.tablesNestedBeyondTheDepthLimitAreSkippedWithAWarning`, `containersNestedBeyondTheDepthLimitAreSkippedWithAWarning` |
| Result contract | 3 samples × 4 option sets = 12 parses, with **0** places where `headers` disagrees with a row's keys, and image metadata complete | `CrossFormatTest.contractHoldsForEverySampleAndMode` and 3 more + `everyOptionCombinationKeepsTheContract` per parser |

## 12. Open-source license, Aspose licensing and evaluation mode

### 12.1 Open-source license and attribution (this project's code)

This project is open source under the **[Apache License 2.0](LICENSE)**, with attribution collected in
**[NOTICE](NOTICE)** (`Copyright 2026 hutianxi (cn.net.tianxi)`).

| You may | You must |
| --- | --- |
| Use it free of charge, including commercially | Ship `LICENSE` and `NOTICE` (or an equivalent readable copy) when you redistribute |
| Modify the source and create derivative works | Note "modified" in the files you changed (Apache 2.0 §4(b)) |
| Depend on it and **release your product closed-source** | Keep the copyright and attribution notices in the source files |
| Redistribute the original or a modified version | Not use the author's name to endorse your derivative work (§6) |
| Open an issue or a PR and become a contributor once merged | Your contribution is licensed under the same terms by default (§5); no separate agreement is needed |

The **attribution text** for consumers (put it in your own `NOTICE`, "About" page or README — copy it
verbatim; both repository URLs are included, GitHub being the primary repository and Gitee the China
mirror, so the source is reachable from whichever opens):

```
This product includes code from document-parser (cn.net.tianxi:document-parser),
Copyright 2026 hutianxi, Apache License 2.0,
original project https://github.com/hutianxi/document-parser
mirror           https://gitee.com/mailtianxi/document-parser
```

**Why Apache 2.0**: the requirement was "free to use + modifiable + merges welcome + the original
author's attribution must be kept".

| License | Free to use/modify | Closed-source derivatives | Explicit patent grant | NOTICE attribution mechanism | Verdict |
| --- | --- | --- | --- | --- | --- |
| MIT | yes | yes | no (copyright terms only) | no | Enough, but the patent risk is left unstated |
| BSD-3-Clause | yes | yes | no | no | Same, plus a non-endorsement clause |
| **Apache 2.0** | yes | yes | **yes** (§3) | **yes** (§4(d)) | **Chosen** |
| GPL / AGPL | yes | **no** (viral: derivatives must be open-sourced under the same license) | yes | yes | Conflicts with "let others use it closed-source" |

Apache 2.0 is the only mainstream permissive license that satisfies all four: permissive enough that
others may use it in closed-source commercial products, while `NOTICE` carries the attribution duty
downstream, plus an explicit patent grant (meaningful for enterprise document-processing code).

**How to contribute**: the process, code conventions, test requirements and license red lines are all in
[CONTRIBUTING.md](CONTRIBUTING.md). One-sentence version: fork → branch off `main` →
`mvn -B clean test` all green → open a PR (behaviour changes must include a before/after end-to-end JSON
comparison).

### 12.2 Aspose licensing and evaluation mode (third-party commercial dependency)

**This repository ships no Aspose binaries**: `pom.xml` only declares the dependencies, and the artifacts
are downloaded at build time from `https://repository.aspose.com/repo/`. Aspose.Words / Aspose.PDF /
Aspose.Slides are **commercial software of Aspose Pty Ltd, licensed under Aspose's own EULA and not
covered by this project's Apache 2.0** — which is also stated in the third-party section of
[NOTICE](NOTICE). Please buy your own Aspose license for production use; a `.lic` file is a private
credential, is blocked by `.gitignore`, and must never be committed in any PR.

Put `aspose-words.lic` / `aspose-pdf.lic` / `aspose-slides.lic` on the **classpath root** or in the
**process working directory**. Each product is registered at most once per process, and this happens
**before a document is loaded** (registering later has no effect on already-loaded documents). A failed
registration prints one line on stderr (earlier versions swallowed it in a silent catch, so an expired
lic showed up as "the output mysteriously lacks content").

Without a lic, parsing runs in evaluation mode, and **the three products' limits differ**:

| Product | Limitation | Visible on the current samples? |
| --- | --- | --- |
| Aspose.Slides | Text frames return only the **first paragraph**, truncated to 5 characters plus an appended notice; table cells go through the same text-frame path, so cells are truncated too | Yes — which is exactly why the pptx sample keeps every piece of text to ≤5 characters (guarded by a test), so the string read is the same with or without a lic |
| Aspose.Words | Injects watermark paragraphs into headers and footers | Yes: iterating only each section's Body plus the `Results.cleanContent` filter means none of the 13 body paragraphs contains a watermark sentence |
| Aspose.PDF | Text-layer paragraphs are not read completely | **No**: this PDF's body text is all inside tables and the space between pages holds only spacer paragraphs, so `text` is an empty string and this limit cannot be quantified on the sample. Tables and images are **unaffected** — the 31-row ledger, 45 data rows and 1 image are all read in full |

Traces left by evaluation mode are filtered uniformly, but **truncated content does not come back**: the
PPT text beyond those 5 characters and the PDF body paragraphs that are not fully read are simply
unavailable without a lic, so evaluation-mode output can be used as an example, not as a result.

Filtering detail: watermark sentences get split into several paragraphs by line breaking (especially in
PDF), so `Results.isEvaluationNotice` removes all whitespace and lowercases before comparing signature
fragments, which recognises even an isolated line like `Pty Ltd.`; `PptParser.stripEvaluationNotice`
removes the truncation notice. Known cost: body text that genuinely mentions something like
"Aspose Pty Ltd" is discarded as a watermark — with a lic in place this never triggers.

The samples themselves are unrelated to licensing: the docx and pptx are hand-assembled OOXML and the pdf
was exported by real Word, and none of the three contains evaluation watermarks; the pptx sidesteps
Slides' truncation via "all text ≤5 characters", and the docx's 13 body paragraphs (longest 62 characters)
prove Words does not truncate body text.

## 13. Testing and quality

`mvn -B clean test` (`JAVA_HOME` pointing at JDK 8) runs **242 tests in 15 test classes** (0 failures,
1 skipped — the skipped one is the `-Ddoc` entry point, see "See the parse result directly" below):

| Test class | Count | Coverage |
| --- | --- | --- |
| `WordParserTest` | 32 | docx sample end to end: 13 body paragraphs, 4 tables cell by cell, the three `FullRow` destinations, registered headers, four image modes, both entry points |
| `PdfParserTest` | 17 | pdf sample end to end: cross-page merging, vertical-merge clustering, lost field results, registered headers switching off the first-row fallback, JPEG re-encoding |
| `PptParserTest` | 19 | pptx sample end to end: shape order, group shapes, the two-row-header trade-off, the evaluation-mode 5-character self-check, plus minimal pptx files assembled on the fly to prove a vertical merge does not eat data and a blank full-width merged row is not taken for a separator |
| `PptParserTextTest` | 2 | Text-frame line splitting (package-private `textLines`): `\r\n`, `\r` and `\n` all split paragraphs, and blank paragraphs take no line. Evaluation mode returns only the first paragraph, so this test does not touch Aspose |
| `PdfCrossPageTableTest` | 4 | A three-page table PDF built on the fly: merged into one group, `source` recording first/last page, the chain stopping when another table follows on a page, the chain also stopping when an image follows the fragment, and the evaluation watermark alone breaking the merge |
| `PdfMergedCellTableTest` | 5 | A docx with merged cells assembled on the fly and exported to PDF: vertical and blank full-width merges are not group rows, docx and PDF agree cell by cell, both readings of a two-level header, and the watermark adding an image only |
| `CrossFormatTest` | 15 | Cross-format consistency: the core table identical cell by cell, the measured `source` page of every group pinned, `source(false)` dropping the key in all three formats, image fingerprints, one reused `Options`, the contract checked across all 12 parses, the 12 JSON lengths quoted in the documentation pinned one by one, the `unmatchedHeaders` top-level key, and the exception types of the three parsers for a missing file |
| `PdfBorderlessTableTest` | 5 | Borderless PDF tables: builds a docx in the test itself (borderless table + ruled table + tab-aligned rows), exports it to PDF, then checks the recovery, that the ruled table is untouched, that prose above the table is not swallowed by the table's bounding rectangle, and the one known false positive |
| `TablesTest` | 34 | The header-grouping algorithm (the only copy): the six-way priority, multi-level header merging, the three destinations for full-row merges, column-name completion, grid padding, and that a registered name is written only on the groups that matched |
| `ImagesTest` | 25 | The image collection (the only copy): numbering, type normalisation and sniffing, fingerprints, duplicates, four modes, external links, file-name sanitising, and that the extension does not follow the default `Locale` |
| `OptionsTest` | 14 | `Options` semantics: tolerance, keywords, silent mode, null handling, `source` on by default and switchable off, `unmatchedHeaders` off by default, and `name` binding to the most recent registration (one-at-a-time and bulk registration are equivalent) |
| `TableByNameTest` | 11 | Naming and name-based retrieval: which groups carry `name`, the key order `headers`/`rows`/`name`/`source`, the cross-format differences between registering the merged row and the upper row plus the dual-spelling remedy, that the default output carries no `name`, the order and null behaviour of `tablesByName`/`rowsByName`, and that a registration covered by a same-name match is no longer reported on stderr |
| `ResultsTest` | 28 | Text cleaning, evaluation-watermark detection, header normalisation, JSON serialisation, file-name sanitising, `writeImages`, license loading |
| `WordEdgeCasesTest` | 14 | Minimal docx files assembled on the fly to cover edge structures: content controls, nested tables, depth limits (8 levels warn not at all, 9 levels warn exactly once), vertical merges, fields, soft breaks, and that a blank full-width merged row is not a separator |
| `PrintResultTest` | 17 (1 skipped by default) | **Asserts nothing, only prints**, producing **18 reports** in four categories: (1) the full result of all three samples (silent / registered / named / DROP); (2) every public entry point (the four `parse` overloads side by side, `Options` defaults and setters, JSON serialisation and `writeImages`, retrieval by name); (3) header rules and the option matrix (four ledger registrations, 7 option sets with their scale and full JSON, the unmatched list, duplicate/blank headers and field caches); (4) the four image modes and the error paths. Each test also writes a UTF-8 copy under `target/print/` (see the next section) |

`TablesTest` / `ImagesTest` / `OptionsTest` / `ResultsTest` do not touch Aspose (they are fed grids, bytes
and minimal strings directly); in this round's single JDK 8 run they took **0.011 / 0.025 / 0.001 / 0.023 s**
(`PptParserTextTest` also only processes strings and took **0.000 s**). The end-to-end classes load Aspose and
the samples; the heaviest are `CrossFormatTest` **29.8 s** (parsing all
three formats repeatedly — `documentedJsonLengthsAreTheRealOnes` alone measures 2.7 s) and `PdfParserTest`
**14.2 s** (every page runs the table engine twice), followed by `PrintResultTest` **24.0 s** (18 reports parse
the samples repeatedly), `TableByNameTest` **7.3 s**, `PdfCrossPageTableTest` **4.0 s**, `WordParserTest`
**1.8 s**, `PptParserTest` **0.94 s**, `PdfBorderlessTableTest` **0.66 s** (it builds a docx and exports it to
PDF inside the test), `PdfMergedCellTableTest` **0.34 s** and `WordEdgeCasesTest` **0.32 s**.
The same classes in the JDK 21 round measured `CrossFormatTest` 38.3 (`documentedJsonLengthsAreTheRealOnes`
3.5 s), `PdfParserTest` 18.5, `PrintResultTest` 22.7, `TableByNameTest` 5.7 and `PdfCrossPageTableTest` 4.8 s.
A whole `mvn -B clean package` round (compile + test + jar) took **01:48** on JDK 8 and **02:11** on JDK 21
in the 1.2.1 dependency-upgrade round (the 1.2.0 release round measured 02:10 / 02:17). The range is dominated
by machine load (the same command measured 02:15–02:37 earlier while
another process was competing for CPU; an earlier round measured 01:32–01:42 on JDK 8 and 01:55 on JDK 21,
with `CrossFormatTest` about 10 s apart), so **read these durations as an order of
magnitude, not as a performance-regression signal** — see §11.

This round's build checks (JDK 8, 1.2.1): `javac -Xlint:all -Xpkginfo:always` with **0** warnings (17 classes,
including `package-info.class`), `javadoc -Xdoclint:all,-missing` with **0** errors,
`mvn -B clean package` BUILD SUCCESS with **242 tests**, 0 failures and 0 errors
(1 `-Ddoc` entry skipped; 242 tests with 0 failures and 1 skipped, with or without `-DargLine="-Dfile.encoding=UTF-8"`), main jar
**60777 bytes / 30 entries / 17 classes** and `major version: 52`. Do not treat the jar size as a checksum:
comment lines shift the line-number tables in the bytecode, and packaging with a different JDK changes the
number too — the same source gave **60777** bytes on JDK 8 and **61190** on JDK 21 in this round, while the
entry count, the class count and the bytecode version (`major version: 52`) matched in all three places.

### See the parse result directly (printed to the console)

Every other test class asserts measured values, which catches regressions but does not show you what the
output looks like. `PrintResultTest` is for humans: run one test and the console gives you **the body text
paragraph by paragraph, the registered name (only when you gave one) and source page of every table group,
the headers and all data rows, the image metadata and the complete JSON**, keeping only one baseline
assertion (the `headers`-vs-row-keys contract), because if the contract breaks the output has no reference
value either.

```
mvn -B test -Dtest=PrintResultTest                                    # all 18 reports (the commands below combined)

# 1. Full result of the three samples
mvn -B test -Dtest=PrintResultTest#printWordSample                    # docx only (silent)
mvn -B test -Dtest=PrintResultTest#printPdfSample                     # pdf only (silent)
mvn -B test -Dtest=PrintResultTest#printPptSample                     # pptx only (silent)
mvn -B test -Dtest=PrintResultTest#printAllSamplesWithRegisteredNames  # "register + name" for all three samples (3 reports)
mvn -B test -Dtest=PrintResultTest#printWordSampleWithRegisteredHeadersAndDrop  # docx registered headers + DROP

# 2. Every public entry point
mvn -B test -Dtest=PrintResultTest#printPublicApiOverloads             # four parse overloads: path/stream × options/keywords, compared on every sample
mvn -B test -Dtest=PrintResultTest#printOptionDefaultsAndSetters       # Options defaults, tolerance, bulk registration, name binding, keywords
mvn -B test -Dtest=PrintResultTest#printSerializationAndImageExport    # toJson + writeImages: file size and SHA-1 against contentId
mvn -B test -Dtest=PrintResultTest#printNamedMatchesAcrossFormats      # tablesByName / rowsByName: which groups and rows each name hits

# 3. Header rules and the option matrix
mvn -B test -Dtest=PrintResultTest#printLedgerNameRemedies             # the two-row-header ledger in the docx: four registration spellings side by side
mvn -B test -Dtest=PrintResultTest#printOptionMatrix                    # three samples × 7 option sets: groups / rows / JSON length / keys
mvn -B test -Dtest=PrintResultTest#printAllModeJson                     # the full JSON of the same 7 option sets, plus 21 .json files
mvn -B test -Dtest=PrintResultTest#printUnmatchedHeaders                # the unmatchedHeaders switch and the resulting list
mvn -B test -Dtest=PrintResultTest#printHeaderAndFieldRules             # duplicate headers with a column suffix, blank headers as 列N, SUM(ABOVE) cached only

# 4. Images and error paths
mvn -B test -Dtest=PrintResultTest#printImageModes                      # META / BYTES / BASE64 / FILE
mvn -B test -Dtest=PrintResultTest#printErrorPaths                      # exception types for missing files and FILE without imageDir
mvn -B test -Dtest=PrintResultTest#printAnyDocument -Ddoc="D:/your-file.docx"   # your own document (parser chosen by extension)
```

Clicking the run button of a single test in the IDE does the same. Output fragment (docx, silent mode;
**the fragments in this section are translated** — the parser prints the document's own text verbatim, and
the three samples are Chinese, so what you see on the console is Chinese):

```
==============================================================================
Document: D:\IdeaProjects\document-parser\src\test\resources\word\sales-report.docx (675208 bytes)
Options: silent — silent mode (no registered headers, no keywords, first-row fallback allowed)
         FullRow = KEEP, ImageMode = META
Scale: 13 body paragraphs, 7 table groups / 30 rows, 1 image
==============================================================================

---- text (13 paragraphs; text inside tables is not here) ----
[1] 2025 East China sales performance report
[2] Report pages: 3
[3] Archive no.:
…
---- tables (7 groups) ----

#1  7 columns × 4 rows, source {page=1}
headers = [district, region, sales, collection rate, sales-5, collection rate-6, remark]
  row1: district=East China | region=Shanghai | sales=1280 | collection rate=92% | sales-5=1465 | collection rate-6=95% | remark=target exceeded
…
---- images (1 image; META mode carries metadata only) ----
[1] {fileName=image-1.png, type=PNG, width=1024, height=768, size=670939, contentId=487e915d…}

---- JSON (4740 characters) ----
{ … }
```

Each test also writes a **UTF-8 copy** of the same output to `target/print/<sample>-<options>.txt`
(the absolute path is printed at the end), so you can open that file when the console encoding is wrong;
it lives under `target/` and `mvn clean` removes it. Measured character counts of the **18 copies** this
round — **191974 characters in total**, shortest 1351 and longest 106026 (`wc -m`, which includes the `\r\n`
that Jackson's pretty-printer uses on Windows): `docx-silent` 7950,
`docx-registered-DROP` 7640, `docx-registered-named` 8304, `pdf-silent` 10367, `pdf-registered-named` 10721,
`pptx-silent` 5949, `pptx-registered-named` 6072, `按名字取数-三份样例` 8568,
`docx-台账四种注册写法` 5511, `公共入口-四个重载` 1938, `选项默认值与setter` 1396,
`表头补名与域缓存` 2409, `序列化与图片导出` 1417, `选项矩阵-三份样例` 3210,
`全部模式-完整JSON` 106026 (the JSON bodies of three samples × 7 option sets, plus 21 per-mode `.json`
files under `target/print/json/`), `未匹配表头-三份样例` 1727, `docx-图片四种模式` 1418 and
`错误路径-三份样例` 1351.

**What the newer reports show**: `全部模式-完整JSON` prints the full JSON of the same three samples × 7 option
sets (one summary line per mode, then the JSON) and writes 21 per-mode `.json` files, so field-by-field
checks and diffs do not have to go through the console; `公共入口-四个重载` runs the four "path / stream × options / keywords"
combinations on every sample and prints whether the path and stream JSON are character-for-character
identical (all `true`); `选项默认值与setter` prints every `options()` default, the unmatched warning when
tolerance is tightened to 0, the equivalence of `headerGroups(List)` and repeated `headers(...)`, and that
`name(...)` binds only to the most recent registration; `表头补名与域缓存` reproduces `销售额-5` /
`回款率-6` and the populated total row from the source documents, and builds a docx whose data is 1+2 while
its `SUM(ABOVE)` cache says 300.00 to prove the parser reads the cache and never recomputes; and
`序列化与图片导出` shows that all three `toJson` implementations agree, that the SHA-1 of every file written
by `writeImages` equals its `contentId`, and that calling `writeImages` on a `META` result only warns and
writes nothing.

Once a name is registered, the title line of each group grows one more field — which is exactly what to
read when deciding which groups to name:

```
#1  7 columns × 4 rows, name = Quarterly detail, source {page=1}
#6  4 columns × 3 rows, name = Annual target, source {page=2}
#7  4 columns × 2 rows, name = Annual target, source {page=2}      ← two header groups of one physical table, one shared name
```

**Do not copy `Column N` into a registration value**: it is a positional name invented for a cell whose
title is empty in the document, not a word from the document (the group whose headers read
`[district, annual target, completion rate, Column 4]` is registered with the first three names only, and
the 4th column is still filled positionally as `Column 4`). Also note that `TestSupport.named()` leaves the
ledger group unnamed in docx — only groups 1, 6 and 7 carry a `name` there; the reason and the fix are in §8.

**To check "name → data" with your own eyes, run `printNamedMatchesAcrossFormats`**: for each name it prints
the groups it hit, every row of every hit and the row count `rowsByName` flattens them to, one section per
sample, and ends with three names that hit nothing. Measured fragment (docx):

```
---- name = "Annual target": 2 groups hit, rowsByName flattens to 5 rows ----
  group #6 of tables, source {page=2}, 4 columns × 3 rows
  headers = [district, annual target, completion rate, Column 4]
    row1: district=East China | annual target=2400.00 | completion rate=101% | Column 4=on target
    row2: district=South East China | annual target=1500.00 | completion rate=93% | Column 4=below target
    row3: district=FY2026 (budget issued) | annual target= | completion rate= | Column 4=
  group #7 of tables, source {page=2}, 4 columns × 2 rows
  …
  are rowsByName and the groups concatenated in order the very same objects: true (5 flattened / 5 concatenated)

---- names that hit nothing: an empty list, never null, never an exception ----
  rowsByName(result, "a name that was never registered") = [], 0 rows
  rowsByName(result, "")                                 = [], 0 rows
  rowsByName(result, null)                               = [], 0 rows
```

**Why the two-row-header ledger in the docx needs a different registration — run `printLedgerNameRemedies`
to see the comparison**: four spellings, one section each, listing the headers and names of all 7 groups, all
13 rows of the ledger group, the row count from `rowsByName` and stderr verbatim. The measured outcome is
plain to read — ① default tolerance 1, registering the merged row: `groups carrying a name: (none at all)`,
`rowsByName` **0 rows**, stderr reports the unmatched header set; ② the same registration with tolerance 2,
③ registering the real upper row of the document, ④ registering both spellings under the same name: all three
give `#2=Collection ledger`, `rowsByName` **13 rows** and **empty** stderr, the 13 rows are character-for-character
identical, and neither the 7 groups' headers nor their row counts move at all (still 7 groups / 30 rows) —
② does not drag any other group along by loosening the tolerance.

**BASE64 mode is not printed**: in the default `META` mode the three samples' JSON is 4740 / 6473 / 3526
characters, which is readable; switching to `BASE64`, that 670939-byte PNG inflates the JSON to 899370
characters, flooding the console with nothing useful. To look at image content use `ImageMode.FILE` +
`imageDir(File)`, and the output will carry the on-disk paths.

### Verifying one document from code

All three parser classes are pure static methods with no container dependency and no CLI entry point. To
inspect a single document from your own code:

```java
// the bundled sample is a Chinese document, so the keyword literal is 销售 ("sales")
Map<String, Object> result = WordParser.parse("src/test/resources/word/sales-report.docx",
        WordParser.options().keywords("销售").fullRow(WordParser.FullRow.DROP));
System.out.println(WordParser.toJson(result));
```

## 14. FAQ

**How do I see the parse result rather than just pass/fail assertions?**
Run `PrintResultTest`: `mvn -B test -Dtest=PrintResultTest` (the three samples), or
`mvn -B test -Dtest=PrintResultTest#printAnyDocument -Ddoc="D:/your-file.pdf"` for your own document.
It prints the body text, each table group's headers and data rows, the image metadata and the complete
JSON to the console, and leaves a UTF-8 copy in `target/print/`.
See [13. Testing and quality](#13-testing-and-quality).

**Why do my registered headers never match?**
Two common causes: first, **the column count must be equal** (tolerance relaxes "a few column names do not
match", not the count); second, the document row's **trailing empty columns are trimmed before
comparison**, so a registered group that itself ends with an empty name will never match
(`OptionsTest.nullTitlesBecomeEmptyNamesButStillCountAsColumns` guards this trade-off). The reverse usage
does work: the docx target table's last column has no title in the document, so registering the three
names `["district","annual target","completion rate"]` matches, and the 4th column is still filled by
position as `Column 4`. When nothing matches, stderr shows `header set did not match any row` plus that
group's column names — adjust the registered values accordingly. The one case that stays silent is a
registration **covered by another registration under the same `name`** that did match: when both spellings of
a multi-level header are registered under one name, the spelling that is bound to miss in this document no
longer makes noise.

**I gave a group a name, so why is there no `name` key in the result?**
`name` is only a label stuck on a header group that *already matched* a registration; it takes no part in
matching and does not widen it. Three cases all leave the key out: that group was registered without a
name; it was named but did not match in this document (stderr then prints `header set did not match any
row` plus that group's column names, unless another registration under the same name did match); or that
group's headers were decided by keywords, by `w:tblHeader`,
by the row after a full-width group row, or by the first-row fallback — there is no registered group to
compare with at all. Multi-level headers add a fourth, format-dependent case: "hits in pdf/pptx, misses in
docx"; registering both spellings under one name fixes it, see [8. Options](#8-options).
Retrieval always goes through `WordParser.tablesByName` / `rowsByName`, which return an empty list for a
name that matched nothing instead of throwing.

**Why is text inside tables missing from `text`?**
By convention table content appears only in `tables` and is not duplicated. The single exception is a PDF
paragraph from which no text rectangle can be obtained — there is no way to tell whether it sits inside a
table, and the choice was to let `text` have one extra line rather than let the body lose a paragraph.

**Why is the PDF sample's `text` an empty string?**
All of that PDF's body text falls inside table rectangles, and the space between pages holds only spacer
paragraphs, so by the convention above everything is skipped; the table data is in `tables`
(6 groups, 45 rows).

**Why does a table's first row become a data row with columns named `Column 1`, `Column 2` after I register headers?**
Registering any header evidence switches off the "first-row fallback". PDF/PPT have no Word-style
`w:tblHeader` flag, so in silent mode the first row is used as a header by fallback; once you register
something, a non-matching table's first row is treated as data. This is deliberate: your evidence must
not be overridden by a guess.

**Why is the PPT text only a few characters long?**
Evaluation mode: Aspose.Slides returns only the first paragraph of each text frame, truncated to 5
characters. Drop in an `aspose-slides.lic` to lift it.

**Why is the total row empty in the PDF but has numbers in the docx?**
The docx reads Word's cached `SUM(ABOVE)` field result; when Word exported the PDF it recomputed the
field and the cached value was not written into the PDF. Use the docx for money amounts.

**Why does `writeImages` write no files at all?**
It only recognises the `base64` field in the result, so it works for `ImageMode.BASE64` only. The default
is `META` (metadata only), and calling it then prints a hint on stderr about which mode to use; to write
files while parsing use `FILE` + `imageDir`.

**Why does `ImageMode.FILE` throw `IllegalStateException`?**
`FILE` mode requires `imageDir(File)` first, otherwise it throws when the first image is reached (the
message says you must specify an export directory first).

**Does parsing read files outside the document or access the network?**
No. Linked images only report their `link` path; every file name goes through `Results.safeName` before
writing, which keeps only the last path segment and adds a numeric suffix on collision inside the
directory, so even a name like `../../evil.png` can only land inside the target directory.

**Can I call it concurrently?**
Yes. All three parsers are stateless static methods; `Options` is mutable, so do not share it across
threads while configuring it — once configured it can be reused concurrently.

## 15. Known limitations and uncovered branches

These are defensive branches written from code paths; the repository has no matching sample, so they have
**no test coverage**, and verifying them requires building new edge samples:

- The **reverse** case of the silent-option PDF mis-merge where "two independent tables happen to have
  the same column count and aligned left edges" (the numbers from that time are in the appendix; they
  were produced with a throwaway probe from constructed coordinates);
- The `page <= 0` branch of `Tables.markSource` (when no position can be given, the `source` key is not
  written): across the 19 groups of the three samples every group yielded a page number in the measurement,
  so this branch has never been hit; layout exceptions such as tables inside text boxes or inside
  headers/footers wait for a sample before they can be verified;
- The out-of-range guard in `WordParser.Options.headerName(int)` and the two guards inside
  `tablesByName`/`rowsByName` for "`tables` is not a list" and "this element is not a map": `headerLabels`
  and the header groups are appended one-to-one, and the `tables` in a result is always a
  `List<Map<String,Object>>` this module put there itself, so a normal call cannot reach these lines — they
  exist only in case the two lists fall out of step later. `TableByNameTest` covers the inputs that really
  do occur: a null result, null/blank/unknown names, and leading/trailing spaces;
- The guard in `PdfParser.flowTableFlow` for "a flow table exposes no valid rectangle at all": all five
  flow tables of the probe had rectangles, so it never triggered;
- The empty-value guard for a PPT cell without a text frame;
- Linked images (Word "Link to File", PPT "Link only"): `ImagesTest` feeds a `link` path directly to
  verify the metadata shape, but none of the three samples contains a real linked image;
- The exact truncation amount of Aspose.PDF's evaluation mode on the **text layer**: the current sample's
  `text` is an empty string, so no number can be measured — verify this yourself once a lic is in place.

Large files and memory: `BASE64`/`BYTES` modes keep the image bytes in the result, so use `FILE` or `META`
for image-heavy documents; OOXML (docx/pptx) is a zip archive that is fully expanded into memory, so a
malformed file with an extreme compression ratio can blow up the process — put a size limit on the
upload/write step, and give parse calls a timeout and a dedicated thread pool.

## 16. Maintainer guide

Requirements for external contributors (environment, code conventions, PR flow, license red lines) are in
[CONTRIBUTING.md](CONTRIBUTING.md); this section is the module split from the maintainer's point of view.

| File | Responsibility |
| --- | --- |
| `WordParser` / `PdfParser` / `PptParser` | Only "read their own format into a `String[][]` grid + paragraphs + image bytes", plus their own `Options` entry points |
| `Tables` | The header-grouping algorithm (the only copy): detection priority, multi-level header merging, where full-row merges go, column-name completion (empty header → `Column N`, duplicate name → column-number suffix; `headers` and row keys share one source), and writing `name` on a registered hit (`markName`) |
| `Images` | The image collection (the only copy): file names, metadata, content fingerprints, duplicates, four output modes |
| `Results` | Shared utilities: JSON serialisation, text cleaning (field codes / soft breaks) and evaluation-watermark filtering, header normalisation, license loading, file-name sanitising before writing |

When changing behaviour, change it in one place according to this split — do not copy it into each parser.
The three formats must return the same structure; `Tables`/`Images`/`Results` are package-private and are
not exposed outside the parsers.

**Adding a new format** (e.g. excel): write an `XxxParser` that reads tables into a `String[][]`
(back-filling positions covered by merges with the origin cell's value), a `boolean[] separator`
(rows merged across the full width) and optionally a `boolean[] heading` (the format's own title-row
flag), then call `Tables.group(...)`; hand images to `new Images(options).add(...)`; finally assemble the
result under the three keys `text`/`tables`/`images`. Reuse `Options`, `toJson` and `writeImages` —
do not write another set.

**The result contract** (do not break it when changing things; `CrossFormatTest` checks it): a physical
table (including Word's nested tables, whose child table becomes a separate entry) yields as many
elements in `tables` as it has header groups; inside one group every row has the same key count and
`headers` equals each row's keys item by item (empty headers become `Column N`, duplicates become
`name-columnNumber`, and both rules share one implementation in `Tables.normalizeHeaders`); when a
registered header matches, the registered names are used (which is why the keys are the same across the
three formats); text inside tables is not repeated in `text`; image order = document-flow order;
`source.page` means the same thing in all three formats (page number / slide index, 1-based) and is
written uniformly by `Tables.markSource` — the grouping algorithm `Tables.group` never touches it.
`Options.source(false)` makes that key disappear for **all three formats together**, leaving an unnamed group with
just `headers` and `rows` (a named group keeps `name`); every other part of the contract is unchanged (each parser asks
`Options.writesSource()` first, and the Word side consequently never creates a `LayoutCollector`).
The key order inside a group is fixed: `headers` → `rows` → `name` (present only when the group was
registered, named, and matched) → `source` (present only under `source(true)`). `name` is written by
`Tables.group` at the moment a registered group matches (`markName`) — `markName` inside the grouping and
`markSource` outside it is deliberate: only the grouping knows which group matched, only the layout knows
the page. A new format has to do neither, it just hands `Options` to `Tables.group`; but do **not**
implement naming or page tracking a second time inside some parser, or the three formats diverge again.
**Registered names are part of the external contract too**: downstream code fetches rows with
`WordParser.tablesByName` / `rowsByName`, so renaming a group or changing what matches is as breaking as
changing a column name.

**After changing a sample**: every expected value was measured, so after changing a sample run the
corresponding test class, update the assertions from the new output, then sync the numbers in the two
"Measured baselines" tables; the sample numbers quoted in the javadoc (the three parser class docs,
`Results`, `Images`, `package-info`) must be updated too.

**Externally visible changes go in the changelog**: whenever a key in the returned structure, a method on
`Options`, a public entry point or a documented behaviour changes, add an entry to the unversioned section at the top
of both [CHANGELOG.md](CHANGELOG.md) and [CHANGELOG.en.md](CHANGELOG.en.md) with the measured numbers from
the three samples; only at release time is that section folded into a version entry and the `pom.xml`
version bumped. Pure internal refactoring and comment-only changes are not recorded.

## 17. Build and publishing

```
mvn -B clean test          # JAVA_HOME pointing at JDK 8; this round measured 242 tests (241 pass, 1 -Ddoc entry skipped)
mvn -B clean package       # artifact target/document-parser-1.2.1.jar (measured this round: 60777 bytes, 30 entries, 17 classes;
                           # classes live directly under cn/net/tianxi/document/ at the jar root;
                           # the prepare-package phase copies the root LICENSE and NOTICE into META-INF/,
                           # so consumers who only get the jar still see the attribution and third-party license notes)
```

The javadoc doclint check can also run the JDK's own `javadoc` directly (without Maven, for a quick check);
the `maven-javadoc-plugin` 3.4.1 in the `release` profile also produced the javadoc jar successfully this round:

```
mvn -B dependency:build-classpath -Dmdep.outputFile=target/cp.txt
javadoc -encoding UTF-8 -charset UTF-8 -Xdoclint:all,-missing \
        -cp "$(cat target/cp.txt)" -d target/site/apidocs $(find src/main/java -name '*.java')
```

That jar is what other projects reference as a `<dependency>`. The sample documents live under
`src/test/resources/{word,pdf,ppt}/` and are not packaged into the artifact jar.

If you ever build an executable fat jar, remember to add `<classifier>exec</classifier>` to
`spring-boot-maven-plugin`: without it, repackage replaces the main artifact in place with the fat jar
(log: `Replacing main artifact with repackaged archive`), all classes move into `BOOT-INF/classes/`, and
consumers referencing the dependency find not a single class.

### Publishing to Maven Central

The `release` profile in `pom.xml` exists for this: `maven-source-plugin` (sources jar) +
`maven-javadoc-plugin` (javadoc jar) + `maven-gpg-plugin` (signing) + `central-publishing-maven-plugin`
(bundle and upload to the Central Portal). Ordinary `mvn clean test/package` does not activate this
profile, so build speed and the local repository are unaffected.

**Why the coordinates are `cn.net.tianxi`**: Central requires the `groupId` to belong to a namespace you
have verified in the Portal. This project's `cn.net.tianxi` is a **domain-type namespace**
(Org: Tianxi, status Verified), verified by adding a Portal-provided TXT record to the DNS of
`tianxi.net.cn`; it naturally matches the Java package `cn.net.tianxi.document`, so no source change was
needed.

> Historical note: when 1.1.1 was first uploaded, `cn.net.tianxi` was not yet verified, so it was
> published once under the GitHub-type namespace as `io.github.hutianxi:document-parser:1.1.1`. Once a
> Central artifact is Published it **cannot be deleted or changed**, so those coordinates will stay in
> Central forever; future versions are published only under `cn.net.tianxi`, so please use the
> coordinates from section 5.1.

**Signing key**: RSA 4096, primary key id `41AF7E29CB53074D`, fingerprint
`5BBC2118BD5BC8805B65F69041AF7E29CB53074D`; the actual signature may be made by its subkey
`21AAF067632A8E6A` (fingerprint `CCE99252368722C3498F8B9021AAF067632A8E6A`), and `gpg --verify`
measurably reports `Good signature`. Both the primary key and the subkey have been pushed to
`keyserver.ubuntu.com` (retrieve with
`https://keyserver.ubuntu.com/pks/lookup?op=get&search=0x<fingerprint>`; both measurably return HTTP 200).
Central looks the public key up on a keyserver when verifying signatures, and rejects it if not found.
The private key on this machine has **no passphrase** (so the build can run fully automatically); to add
one: `gpg --edit-key 41AF7E29CB53074D` → `passwd`, then put the passphrase in settings.xml under
`<server><id>gpg</id><passphrase>`.

There are two publishing routes; 1.1.1, 1.2.0 and 1.2.1 all use the first (**it does not require putting the token on disk**):

**A. Build the bundle locally + upload through the Portal web UI**

```
mvn -B -P release deploy -DskipPublishing=true      # the version must not contain -SNAPSHOT
```

This produces `target/central-publishing/central-bundle.zip` (266857 bytes for 1.1.1, 236990 bytes for 1.2.1;
24 entries in both cases). The 1.1.1 zip held, under `cn/net/tianxi/document-parser/1.1.1/`, the
`document-parser-1.1.1.jar` (56275 bytes), `-sources.jar` (64964 bytes), `-javadoc.jar` (141571 bytes) and
`.pom`, each with an `.asc` signature and `md5/sha1/sha256/sha512` checksums.

Then go to Portal → **Publishing → Publish Component**: fill in the Deployment Name (e.g.
`document-parser 1.2.1 (cn.net.tianxi)`) and a description, choose this zip, and click Publish Component.
After submitting, a **VALIDATING** record appears in the Deployments list (the 1.1.1 one was
`960f618b-cb17-4cb1-9d40-96849df2a219`; 1.2.1 is `11cdf8ca-0884-4626-9356-c82a6455144e`); once validation
passes, a **Publish** button appears on the same row, and clicking it is what actually makes it public.

Two pitfalls (both hit in practice):

- The property that skips the upload is **`-DskipPublishing=true`**; writing
  `-Dcentral.skipPublishing=true` has no effect, and the plugin then tries to upload with placeholder
  credentials and fails with `Deployment failed while publishing`;
- Even with `skipPublishing=true`, the plugin still reads `<server><id>central</id>` from
  `settings.xml`, and without it fails immediately with
  `Unable to get publisher server properties for server id: central: NullPointerException`. So even when
  you only build the bundle you need a placeholder entry (pointing `-s` at a temporary settings file is
  enough — no need to pollute the global config).

**B. Direct upload from the command line** (suited to CI): put the Portal user token and the GPG
passphrase in `settings.xml`, then run `mvn -B -P release clean deploy` and the plugin uploads itself;
`autoPublish=false`, so you still click Publish in the web UI.

```xml
<servers>
  <server>
    <id>central</id>
    <username>token username generated in the Portal</username>
    <password>token password generated in the Portal</password>
  </server>
  <server>
    <id>gpg</id>
    <passphrase>key passphrase (this machine's key has none, leave empty)</passphrase>
  </server>
</servers>
```

Validation checks: presence of the sources and javadoc jars, completeness of the pom metadata
(`name`/`description`/`url`/`licenses`/`developers`/`scm` — all filled in for this project), whether the
signature verifies against a public keyserver, and whether the `groupId` belongs to a verified namespace.
If it fails, **Drop** it, fix and retry; once Published it cannot be withdrawn, and the same version
number cannot be uploaded twice — you must bump the version.

After Publish, syncing to `repo1.maven.org` takes tens of minutes to a few hours (the search index is
slower). Verification URL:

```
https://repo1.maven.org/maven2/cn/net/tianxi/document-parser/1.2.1/
```

1.1.1 and 1.2.1 both went through the whole flow and were **verified by measurement**:

- 1.1.1: Portal Deployment `960f618b-cb17-4cb1-9d40-96849df2a219` reached status PUBLISHED, and the
  jar / sources.jar / javadoc.jar byte counts listed in that directory match the local artifacts **of that
  build** (56275 / 64964 / 141571);
- 1.2.1: Portal Deployment `11cdf8ca-0884-4626-9356-c82a6455144e` reached status PUBLISHED on 2026-10-24;
  Central serves a 60777-byte jar, a 12136-byte `.pom`, an 81550-byte sources.jar and an 89001-byte
  javadoc.jar, matching the local release bundle. All four `.sha256` files were checked — jar
  `ce9af926820b55b18505d2efa51a02cbd92dbe73376fda7a46d3fec6298c0ca3`, sources
  `1d2fbc27afe3974036c4c6233b24577b94d7b4d014403af4caebb0b31e68420c`, javadoc
  `94ac6f87c12143e2d122009373db6afab727a5d5ea4a36c193e49440f6f94e8c`, pom
  `e6fda84c361fbfafc596490bc26134e25584a41f7c0cf39eec07e3c19dc78e53`.

For 1.1.1, a **brand-new empty local repository** was also used to run
`mvn dependency:get -Dartifact=cn.net.tianxi:document-parser:1.1.1 -Dtransitive=false`
(`-Dmaven.repo.local` pointed at a temporary directory so it could not hit an already-installed artifact
of the same name), giving BUILD SUCCESS, and the downloaded jar was
`https://repo.maven.apache.org/maven2/cn/net/tianxi/document-parser/1.1.1/document-parser-1.1.1.jar` with
sha256 `c2c47ba92823ee9464a9a80bb6125ee8b8aa34f562ee80d76f709df3ce233148` — identical in all three
places: Central's `.sha256`, the local artifact and the GitHub Release attachment. The temporary
repository was deleted afterwards.

> **Version / artifact consistency**: this README and the changelog now describe **1.2.1** (published
> 2026-10-24). The 1.2.1 on Central (main jar 60777 bytes) and 1.2.0 (60236 bytes) have exactly the same
> API; the only difference is Jackson 2.21.8 versus 2.21.7.

### 1.2.1 (2026-10-24)

`pom.xml` is now `1.2.1`; a local `mvn -B clean package` produces a **60777-byte** main jar on JDK 8
(**61190 bytes** for the same source on JDK 21) with 30 entries and 17 classes, and 242 tests run with 0 failures
(1 skipped). The sources and javadoc jars are produced by the `release` profile; the bundle built on
2026-10-24 02:32 holds JDK 8 sources and javadoc jars of **81550** and **89001** bytes (the javadoc footer
carries a generation timestamp, so that jar varies by a few bytes between runs). The main-jar sha256 is
`ce9af926820b55b18505d2efa51a02cbd92dbe73376fda7a46d3fec6298c0ca3` (jar entries carry build timestamps,
so a rebuild changes it — the Central record is authoritative after publishing).
A local `-P release deploy -DskipPublishing=true` run produced
`target/central-publishing/central-bundle.zip` (**236990 bytes, 24 files**, sha256
`fdcd57e91f009d3c3438c95a75dce74ccda23baa327e1f52eac30ae3ab422612`); all four `.asc` files pass
`gpg --verify` (the signature is made by subkey `21AAF067632A8E6A`). It was uploaded and published on 2026-10-24.

The only change is dependencies: Jackson 2.13.5 → 2.21.8, closing CVE-2026-83557 / 68497 / 19032 / 89425 /
89407 / 91776 and fixing CVE-2026-110879 / GHSA-m8xf-2cr7-77ph (published 2026-10-20, an incomplete fix for
CVE-2026-54512). `DatabindContext` in 2.21.8 adds array-element-type validation, and
`ApprovedBaseTypeValidator` only exists in this release; 2.21.7 and earlier do not contain that code.
Jackson 2.x is still Java 8 bytecode (`major version: 52`). The test dependency moved from
`spring-boot-starter-test` to `junit-jupiter`, keeping Spring 5.3.x / logback 1.2.x and similar off the test
classpath. The public API, result shape and parsing behaviour are unchanged.

Upload and verification record (2026-10-24): Portal Deployment `11cdf8ca-0884-4626-9356-c82a6455144e`,
status **PUBLISHED**; `repo1.maven.org` serves a **60777-byte** `document-parser-1.2.1.jar`, a **12136-byte**
`.pom`, an **81550-byte** sources.jar and an **89001-byte** javadoc.jar, matching the local release bundle.
All four `.sha256` files were checked one by one: jar
`ce9af926820b55b18505d2efa51a02cbd92dbe73376fda7a46d3fec6298c0ca3`, sources
`1d2fbc27afe3974036c4c6233b24577b94d7b4d014403af4caebb0b31e68420c`, javadoc
`94ac6f87c12143e2d122009373db6afab727a5d5ea4a36c193e49440f6f94e8c`, pom
`e6fda84c361fbfafc596490bc26134e25584a41f7c0cf39eec07e3c19dc78e53`.

### 1.2.0 (2026-10-23)

`pom.xml` is now `1.2.0`; a local `mvn -B clean package` produces a **60236-byte** main jar on JDK 8
(**60650 bytes** for the same source on JDK 21) with 30 entries and 17 classes, and 242 tests run with 0 failures (1 skipped).
The sources and javadoc jars are produced by the `release` profile at package time.

**Upload and verification record (measured 2026-10-24)**: Portal Deployment
`923fda1a-a25c-46f3-b83d-a387966a5def` is **PUBLISHED**, and every artifact on `repo1.maven.org` is
byte-for-byte identical to the local build:

| Artifact | Bytes | sha256 |
| --- | --- | --- |
| `document-parser-1.2.0.jar` | 60236 | `92db3a721654ea142aff88f6735cca5a557576febde5251b80b784aabf25a4b4` |
| `document-parser-1.2.0-sources.jar` | 81011 | `f564cd6626cb9df39d191d3bdb53f0dc6acee6932871f5dc51749690f9c39139` |
| `document-parser-1.2.0-javadoc.jar` | 88998 | `aea880dfe09ca25af534571c4f01b8521d00aeb5190c966896f807c6c831aaca` |
| `document-parser-1.2.0.pom` | 11118 | `fd7b4a0c976ff730c822e9d4d16e83d73c41e2ab1db2cd8403d7e0f1dfd5a2c5` |

The `central-bundle.zip` used for the upload is 235380 bytes and 24 files (245751 bytes uncompressed),
local sha256 `402f1d7e39f7790c0f6796548d1f6beba4cdcede6e5c6123a364dd97152fac82`; it is only the upload
container and is not served by Central on its own.

A **brand-new empty local repository** was then used for `mvn dependency:get
-Dartifact=cn.net.tianxi:document-parser:1.2.0 -Dtransitive=false` (`-Dmaven.repo.local` pointed at a
temporary directory): **BUILD SUCCESS**, with the jar fetched from
`https://repo.maven.apache.org/maven2/cn/net/tianxi/document-parser/1.2.0/document-parser-1.2.0.jar`,
60236 bytes and the same sha256 as the first row above. Run from this repository's directory, Maven first
looks for the artifact in the Aspose repository declared in the pom (it is not there) and then fetches it
from Central. The temporary repository was deleted afterwards.

The GitHub Release [`v1.2.0`](https://github.com/hutianxi/document-parser/releases/tag/v1.2.0) is published
as well (tag pointing at commit `bfbdec9`, marked Latest); the `size` and `digest` of its three attachments,
as returned by the GitHub API, match the table above item by item.

**One thing consumers must know**: the three Aspose artifacts are not on Maven Central, so anyone
depending on `document-parser` must declare `https://repository.aspose.com/repo/` in their own pom as
shown in section 5.1, otherwise dependency resolution fails.

After 1.1.1 was published to Central, a GitHub Release `v1.1.1` was created with the jar, sources.jar and
javadoc.jar attached, so people not using Maven can download them directly:
[v1.1.1](https://github.com/hutianxi/document-parser/releases/tag/v1.1.1) is published and all three
attachments are in Assets (the sha256 shown by GitHub matches the local artifacts). The 1.2.0 release has
also been created: [v1.2.0](https://github.com/hutianxi/document-parser/releases/tag/v1.2.0) (published
2026-10-24 00:18 +0800, Latest) carries the same three artifacts, and their `size` / `digest` match the
table above.

## 18. Appendix: fix log (historical numbers)

The numbers below were measured with throwaway probes while fixing bugs (the probes were deleted after
use and are not in the repository). **The constructed samples they belonged to have since been replaced
by the three current business samples, so re-running will not reproduce the same numbers**; they are kept
to explain why the current code looks the way it does.

- **Text cleaning** (constructed document: a paragraph containing a `PAGE` field and a soft break, plus a
  two-cell row whose first cell holds only a field; `<XX>` is the hex of a control character):
  `before <13>PAGE<14><15> after<0D>` was `before PAGE after` before the fix → `before after` after;
  `line1<0B>line2<0D>` was `line1line2` before → `line1 line2` after;
  `<13>PAGE<14><15><07>normal cell<07>` produced headers `["PAGE","normal cell"]` before →
  `["Column 1","normal cell"]` after; a whole evaluation-watermark paragraph used to enter `text`
  verbatim and is now an empty string.
- **A data row with a vertical merge mistaken for a full-row merge** (a two-column table where `Column A`
  is vertically merged across two rows and `Column B` has a value per row): the old `Grid.isSeparator`
  criterion excluded both horizontally and vertically continuing cells, so a row "left with only one
  origin cell" was taken for a group sub-heading. Before the fix, `KEEP` produced
  `[{Column A:"first",Column B:"value1"},{Column A:"first",Column B:""}]` (`value2` wiped to an empty
  string) and `DROP` left only one row; after switching to horizontal continuation only, all three modes
  give `[{Column A:"first",Column B:"value1"},{Column A:"first",Column B:"value2"}]`. At the time, the old
  and new `isSeparator` were compiled into two directories and the same probe was run against each
  (5 samples × 3 `FullRow` modes = 15 runs): group counts, row counts and `tables` JSON lengths matched
  item by item, with only that vertically merged table differing.
- **Cross-page mis-merge under silent PDF options** (`hasHeaderCandidate` is only meaningful once header
  evidence is registered; with neither registered, no row can count as a header, so of the four criteria
  only geometry remains, and two independent tables on adjacent pages with the same column count and
  aligned left edges get merged into one): using the real coordinates of the old `sample.pdf` (page height
  841.9), two fragments were constructed and `continues` was called directly — the previous page's table
  drawn to mid-page (bottom edge 548.25, 0.65 of the page height) and the next page's starting at the top
  (top edge 769.65, 0.086 of the page height from the top). Before the fix "merges = true"; after the fix
  "merges = false" under silent options; for the same pair of positions, both versions gave
  "merges = true" when header evidence was registered, showing the new geometric condition only kicks in
  as a fallback when there is no header evidence. The end-to-end `tables` JSON was compared too: before
  the fix the two merged into **1 group** (page 2's header row stuffed as data under page 1's header),
  after the fix **2 groups**, each with its own header.
- **A long table spanning three or more pages only merged its first two** (after `mergeSplitTables` merged
  one fragment, `previous` still pointed at the first fragment, while `continues` requires adjacent page
  numbers, so page 3 could not match it): three page fragments were constructed by reflection using real
  coordinates (each fragment drawn from page top to page bottom). Under silent options, before the fix
  **2 fragments** remained (pages 1+2 merged into 2 rows, page 3 alone with 1 row whose first row was also
  taken as a header by fallback); after the fix **1 fragment with 3 rows**. The fix is to replace
  `previous`'s `page` and `bottomGap` with the later fragment's after each merge (the left edge is not
  replaced — it is always compared against the first fragment, so a 2pt-per-page tolerance cannot
  accumulate into drift).
- **Column-name alignment** (`headers` is used as a column list for ingestion, so it must equal every
  row's keys item by item): columns whose header is blank, and columns that end up with the same name
  after a merge is flattened, used to be named only inside the row objects (`Column 3`, `Total-2`) while
  `headers` still held `""` and two identical `Total`s; `Tables.normalizeHeaders` now completes both by
  the same rule, so a constructed header `Total | Total | Name` yields
  `headers=["Total","Total-2","Name"]` with row keys matching exactly. At the time, 41 header groups and
  105 data rows were run together with the constructed samples, with 0 mismatches.
- **`ObjectMapper` created on every call**: `toJson` used to build a new `ObjectMapper` per call; it now
  reuses one static instance (`ObjectMapper` is thread-safe).
- **The way runs of spaces were collapsed**: it used to be `while (contains("  ")) + replace`, which
  re-copied the whole text a dozen times for a long run of spaces; measured on a cell with 3 million
  spaces, the old approach took 95 ms and a single regex pass takes 10 ms.
- **Repeated license registration**: an Aspose license is process-level global state, and `parse` used to
  re-read the lic file on every call; when batch-parsing hundreds of documents all that IO was wasted.
  Now each product is registered once, and a failed registration prints one line.
