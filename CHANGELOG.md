# 更新日志

记录 document-parser 每个版本的**对外可见变化**：返回结构里的键、`WordParser.Options` 的方法、公开入口、
行为口径，以及每项改动在仓库三份样例上的实测数字。实现细节仍以源码 javadoc 与 [README.md](README.md) 为准；
英文见 [CHANGELOG.en.md](CHANGELOG.en.md)。

最新在最上。版本号后的日期是**该 tag 的提交日期**。工作树里有尚未发版的改动时，会在最上方单独列出。

---

## 1.2.0 — 2026-10-23

`pom.xml` 已升到 `1.2.0`，`mvn -B clean package` 的产物是 `document-parser-1.2.0.jar`。
本版收进 `v1.1.1` 之后累积的全部对外可见变化：`source` / `name` / `unmatchedHeaders` 三个新 API、
按名字取组的入口、PDF 无边框表补读，以及发版前查出的三个缺陷修复。
相对 1.1.1 的输出变化见「行为口径」，测试与构建数字见「构建产物」。

### 新增

- **PDF 无边框表补读**：线段引擎之外再跑一遍流式引擎，只收「每行单元格数相同、且不与任何有线表重叠」的那几张。
  实测临时探针造的无边框样例 **0 组 → 3 组**，仓库里有线的 `sales-report.pdf` 一组没多。
  已知代价：用制表位排齐的几行散文本满足同样判据，会被当成一张表从 `text` 挪进 `tables`
  （`PdfBorderlessTableTest.tabAlignedParagraphsCurrentlyCountAsATable` 把这个行为写死成断言）。
  测试类 `PdfBorderlessTableTest`（5 条；其中一条钉住「表上方正文不被吞掉」，见「修复」）。
- **每组表带来源位置 `source`**：`{ "page": 起始页 }`，跨页的表另带 `"lastPage"`。
  三种格式同一个键名同一套字段：docx 是 `LayoutCollector` 算出的页码、pdf 是页码、pptx 是幻灯片序号，都从 1 起。
  实测三份样例 19 组每组都有位置值，跨页的是那份回款台账，docx 与 pdf 各一份（都是 `{"page": 1, "lastPage": 2}`）。
  由 `Tables.markSource` 统一写，分组算法 `Tables.group` 本身不碰它。
- **`Options.source(boolean)`**（默认 `true`）：不要页码就关掉，**三种格式一起生效**，关掉后每组只剩
  `headers` 与 `rows`，列名与内容一字不变。docx 关掉后连 `LayoutCollector` 都不建，省下整篇版面计算 ——
  一次性探针在同一 JVM 里把开／关**交替**各跑 6 次、连跑两轮：**开 74–254ms／关 11–35ms**，
  两轮里「关」的最大值都低于同一轮「开」的最小值。PDF／PPT 的页码本来就免费，开关只决定写不写这个键
  （同轮实测 pdf 681–960ms 对 696–984ms、pptx 46–86ms 对 44–71ms，区间重叠）。探针跑完即删，仓库里没有它。
- **`Options.name(String)` 与每组表的 `name` 键**：注册表头时给这组起个名字，命中之后该组多写一个
  `"name": "季度明细"`，入库时按名字定位，不必再去数 `tables` 下标。
  键写在与注册组比对成功的那一刻（`Tables.group` 内部），所以三种格式行为一致；
  没起名的组、以及没命中注册组的表**都不写这个键**。
  实测字符数（同一份注册，起名前 → 起名后）：docx **4740 → 4828**（7 组里 4 组带 name：第 1、2、6、7 组）、
  pdf **6473 → 6561**（6 组里 4 组：第 1、2、5、6 组）、pptx **3378 → 3466**（6 组里 4 组：第 1、2、4、5 组）。
  起名不引入任何版面计算，本轮也没有为它单测开／关耗时 —— 这几个长度与文档里其余的长度数字（静默、
  注册不起名、注册起名、BASE64 四组 × 三种格式，共 **12 个**）现在都是
  `CrossFormatTest.documentedJsonLengthsAreTheRealOnes` 的断言值，改坏了测试立刻红。
- **按 name 取组取行**：`WordParser.tablesByName(result, name)` 返回所有 `name` 相同的组（按 `tables` 原顺序，
  返回的是结果里的同一批对象）；`WordParser.rowsByName(result, name)` 把这几组的数据行按同序摊平成一个列表，
  直接可以逐行入库。没命中、传 null、传空白都返回**空列表**，不抛异常。测试类 `TableByNameTest`（11 条）。
- **顶层 `unmatchedHeaders` 开关**（`Options.unmatchedHeaders(boolean)`，默认 `false`）：打开后结果多一个
  顶层键，排在 `text`／`tables`／`images` 之后，列出「注册了却一组都没命中」的表头集合。每项默认只有
  `headers`（注册时的原文）；如果这组注册起过名字，再带一个 `name`。它和标准错误上那条提示是同一份规则，
  只是给了程序读的入口。实测三份样例的 JSON 长度（起名 + 打开）：docx **4857**（五组全命中，值是空列表）、
  pdf **6632**、pptx **3537**（各一项 `["故障部件","工单数"]`）。测试
  `CrossFormatTest.unmatchedHeadersIsAnOptInTopLevelKey`、
  `unmatchedHeadersCarriesTheNameOfARegistrationThatMatchedNothing`。

### 修复

下面几条都是本轮「逐方法核对 + 现场拼最小文档」查出来的真实缺陷，**每条都是先有一条失败的测试、再改代码**，
不是读代码读出来的猜测。三种格式的「跨整行合并行」判据也在这轮统一成同一条。

- **PPT 纵向合并的数据行被误判成分组行，数据被吃掉**（`PptParser.parseTable`）：原先数的是「合并起始格」
  （`getFirstRowIndex() == r` 且 `getFirstColumnIndex() == c`），纵向合并的延续格因此不算这一行自己的格子。
  于是「A 列纵向合并跨两行 + B 列每行各有值」的两列表里，第二行只剩一个起始格、被当成跨整行的分组小标题：
  实测 `KEEP` 把 `值二` 抹成空串（读到 `[{列A:第一,列B:值一},{列A:第一,列B:}]`）、`DROP` 整行删掉，
  而同一份版面的 docx 两种模式都读到 `值二`。判据改成只排除**横向**合并延续格（与 `WordParser.Grid.isSeparator`、
  `PdfParser.toGrid` 同一条）之后，两种模式都与 docx 逐格相同。
  测试 `PptParserTest.verticallyMergedDataRowIsNotAMistakenSeparator`。
- **全空的跨整行合并行被当成分组行，一张表被切成两组、下一行数据被当成表头**（`WordParser.Grid.isSeparator`
  与 `PptParser`；`PdfParser.toGrid` 本来就有这条判据，是这两种格式没对齐）：排版里常见的空白合并行
  也满足「整行只剩一个格」，于是它后面的那一行按「分组行之后的行是表头」被吃成表头 ——
  实测 4 行 2 列的表（`列A/列B` 表头、`甲乙`、空白合并行、`丙丁`）改前读出 **2 组**：
  第 1 组列名 `列A/列B`、只有 `甲乙` 与那行空值，第 2 组列名成了数据 `丙/丁`、**0 行数据**（`丙丁` 整行丢失）；
  `DROP` 下第 1 组只剩 1 行。判据补上「**且这一行不是全空**」之后：1 组、列名仍是 `列A/列B`、3 行数据
  （空白行本身留成一行空值，不凭空丢内容）。
  测试 `WordEdgeCasesTest.blankFullWidthMergedRowIsNotASeparator`、`PptParserTest.blankFullWidthMergedRowIsNotASeparator`
  （两条都是**先失败**（`expected: <1> but was: <2>`）**后通过**）。
- **`Images.extension` 跟着默认 `Locale` 走**：`String.toLowerCase()` 无参版本用平台默认区域，
  在土耳其语环境（`tr_TR`）下 `"TIFF"` 会变成 `tıff`（无点的 i），扩展名与 MIME 子类型一起错，
  写出来的文件名也跟着错。三处（`Images.extension`、`Images.contentType`、`Results.normalize`）统一改成
  `Locale.ROOT`。测试 `ImagesTest.extensionDoesNotFollowTheDefaultLocale`：把默认 Locale 设成 `tr_TR`
  再断言 `tiff`／`png`／`jpg`／`image/tiff`。
- **PDF 补读无边框表时，表格上方一行正文会无声消失**（`PdfParser.flowTableFlow`）：流式引擎给的整表外接矩形
  比实际格子高约 25pt（那一段正好是评估水印占的位置），表格上方那行正文落在外接矩形内、却在每个格子之外，
  `textFlows` 的「已经在表格里」判据就把它跳过了 —— 既没进 `tables` 也没进 `text`。改成把**格子矩形**
  （`grid.bounds`，来自 `AbsorbedCell.getRectangle()`）登记给正文判重之后，这行正文回到 `text`；
  有线表那一支（`tableFlow`）照旧用线段围出的矩形，那是画出来的边框，量不出这种虚高。
  测试 `PdfBorderlessTableTest.proseAboveTheBorderlessTableStaysInText`。
- **多级表头合并完之后没有再与注册集比一次**（`Tables.group`）：docx 台账的两行表头在网格里是分开的，
  注册「合并之后那一行」逐行比对时两行都命不中（上层差 2 格、下层末尾空列被裁掉后列数不同），
  于是一份跨格式通用的注册在 docx 上落空。改成连续表头行合并完再拿合并结果比一次后，docx 也命中；
  容忍度收到 0（要求逐格相等）仍命中。只注册台账这一组时三份样例实测：docx 13 行、pdf 31 行、pptx 7 行，
  列名都是 `[客户名称, 合同金额, 已回款, 回款率, 备注]`。测试
  `TableByNameTest.aMultiLevelHeaderHitsByItsMergedRowInAllThreeFormats` 把容忍度 1／0 与四种注册写法都钉住。
- **PDF 跨页长表并完一截后「本页最后一块」的标记没跟着换，下一页的独立表会被错并进来**（`PdfParser.mergeSplitTables`）：
  并完只把 `page`／`bottomGap` 换成后一截的，`lastOnPage` 还停在首截的 `true`。两页的长表看不出问题
  （第二次比较之前已经到最后一截），三页以上、且中间那页的片段后面还跟着图片或别的块时，
  下一页的独立表只要列数与左边界对得上就会被并进台账。现造样例：第 1 页表头 + 29 行、第 2 页 29 行、
  片段后跟一张 8×8 PNG、第 3 页独立两列表 `B1…B5`。改前读出 **1 组**（`source={page=1,lastPage=3}`，
  独立表的 `B1…B5` 全并进台账）；补上 `previous.lastOnPage = flow.lastOnPage` 之后 **2 组**：台账 58 行、
  独立表 `B2…B5` 4 行、`source={page:3}`。测试
  `PdfCrossPageTableTest.theChainAlsoStopsWhenANonTableFlowFollowsTheFragment`（去掉这一行即复现 1 组）。
- **PPT 多段文本框被压成一行**（`PptParser.addText`）：`ITextFrame.getText()` 用 `\r` 分隔段落
  （Windows 换行是 `\r\n`），旧代码只按 `\n` 拆，随后 `Results.clean` 又把 `\r` 变成空格，
  于是 `第一段\r第二段` 读成 `第一段 第二段`。改成 `\r\n`／`\r`／`\n` 三种都拆，空段不占行。
  评估模式一个文本框只给第一段，端到端测试照不到多段，所以拆行逻辑抽成包级 `textLines`，
  由新增的 `PptParserTextTest`（2 条单测）直接钉住。
- **Word 恰好 8 层的表格被误报「嵌套超过」，深层叶子节点也会误报**（`WordParser.parseNestedTables` / `walk`）：
  第 8 层的表已经到上限，但它下面根本没有子表，旧代码仍在 `parseNestedTables` 入口按 `depth >= 8`
  直接打「更里面的内容已跳过」—— 正常的 8 层文档被误报成超限；`walk` 先判深度再判 `isComposite`，
  深度到 16 的叶子节点（书签类，没有子节点可钻）同样被报。改成提示只在真的遇到第 9 层子表时打一次、
  先判有无子节点再判深度。实测 8 层：**8 组、标准错误为空**；9 层：**1 条提示、前 8 组照常**。测试
  `WordEdgeCasesTest.exactlyEightNestedTablesAreAllReadWithoutAWarning`、
  `ninthNestedTableIsSkippedWithExactlyOneWarning`（后者的提示还断言只出现一次）。

### 文档修正

同轮查出的文档错误（都已改，不是「已知问题」）：

- 「跨页的只有 pdf 那份台账」是错的 —— `CrossFormatTest.sourceRecordsThePageEachBlockStartsOn` 断言 docx 的台账
  同样是 `{"page": 1, "lastPage": 2}`；331／287／264 那三个差值的算法（`7 × 44 + 23`）也印证 docx 有一个 `lastPage`。
- `-Dtest=PrintResultTest` 打的是 **17** 份报告（三份样例静默各一份、docx「注册 + DROP」一份、三份「注册 + 起名」、
  按名字取数、台账四种写法、公共入口四个重载、选项默认值与 setter、表头补名与域缓存、序列化与图片导出、
  选项矩阵、未匹配清单、图片四模式、错误路径），原文写成「三份样例各打一份」。
- 「docx 台账两行与注册名都差两列」的说法不准确：实测上层那行是差 **2 格**（超默认容忍度 1），
  下层那行是**末尾空列被裁掉后列数不同** —— 这两条只解释「逐行比为什么命不中」。合并后再比一次之后，
  照合并后那行注册也能命中（容忍度 0 也行），所以文档现在写的是三种格式都命中，
  并把容忍度 1／0 与四种写法都写进测试当证据。
- 文档里的 12 个 JSON 长度原先只是文字，现在全部是断言（`CrossFormatTest.documentedJsonLengthsAreTheRealOnes`）。
- **「整轮比 1.1.1 那轮慢约 80 秒、单类 31.5 秒 → 76.5 秒」这条不成立，已撤回**：那是把有别的进程抢 CPU 的一轮
  与空闲的一轮并列比出来的。同一 JVM 里把这条断言开／关**交替**各跑三轮：开 34.6／30.9／37.6 秒、
  关 29.8／35.9／38.2 秒，差值完全被轮次噪声吃掉。能精确归因的只有 surefire 报告里的单条耗时
  （`documentedJsonLengthsAreTheRealOnes` **3.3 ～ 3.8 秒（JDK 8，两轮）／3.2 秒（JDK 21）**，取自
  `target/surefire-reports/TEST-*.xml`；上几轮这两个数在 3.5／5.6 附近波动）。README §11 与 §13 的耗时口径按实测重写，并写明
  「耗时只看数量级、别当性能回归的判据」。
- **UTF-8 副本的字符数原先少算了**：Jackson 美化输出在 Windows 上用 `\r\n`，早前那次统计没把 `\r` 算进去。
  按 `wc -m` 重测 **18 份**副本，合计 **191974** 个字符（最短 1351、最长 106026；带 JSON 的副本含 CR，统计已计入）。
- **英文 README 的示例字面量直接跑仓库样例命不中**：§5.3 与 §13 的列名、名字、关键字都是译名，
  而三份样例是中文文档。已补一节说明并给出真实字面量映射（`片区`／`销售额`／`季度明细`／关键字 `销售` 等），
  「在代码里验证一个文档」那段的 `keywords("sales")` 改成 `keywords("销售")`（那一处解析的正是仓库样例路径）。
- 测试条数、每类耗时、整轮耗时、jar 字节数全部回填本轮实测值（见下）。

### 行为口径

- **默认输出多了 `source` 键**（这是相对 1.1.1 的输出变化，不是零成本兼容）：三份样例 META 模式的 JSON
  各多 **331／287／264** 个字符（docx 7 组、pdf 6 组、pptx 6 组，每组多一段 44 字符的 `"source" : { "page" : N }`；
  docx 与 pdf 的台账都跨页、各多一个 23 字符的 `lastPage`，`7 × 44 + 23 = 331` 正好对上）。
  这三个差值可复现：从 Maven Central 拉 `1.1.1` 的主 jar 跑同一份样例，META 长度是 **4409／6186／3262**。
- **每组的键顺序固定**为 `headers` → `rows` → `name`（只在注册且命中时写）→ `source`（只在 `source(true)` 时写）。
  `name` 排在 `rows` 之后：`Tables.newBlock` 先建这两个键，新组与多级表头合并两条路径的键顺序才一致。
- **一个名字可以命中多组**（同一份注册在一张物理表里被切出多组时都命中），此时这几组写同一个 `name`，
  `tablesByName` 会一次返回多组；连续两行表头都命中注册组时按**第一个**名字为准。
- **多级表头的跨格式差异（补丁后重测）**：分组算法先逐行比对网格，连续表头行合并完再拿合并结果比一次。
  所以照「合并之后那一行」注册（`客户名称／合同金额／已回款／回款率／备注`）时，**docx／pdf／pptx 都命中**，
  默认容忍度 1 与容忍度 0 都命中；只注册台账这一组时三份样例分别 13／31／7 行。
  换成照「上层那一行」注册（`客户名称／回款情况／回款情况／回款率／备注`）才有格式差异，三条都是实测：
  docx 13 行（台账两行表头在 XML 里都带 `w:tblHeader`，两行都算表头、照样合并成一组）；
  pptx 8 行（只有上层那行命中注册、被当成表头，下层那行不再算表头就落进了数据，实测它是第 0 行，
  与静默模式读到的 8 行一模一样）；pdf 命不中（它的网格里给的就是合并后那一行，没有上层那一行）。
  跨格式共用一份注册时最稳的写法仍是上下两种写法都注册、起同一个名字
  （`TestSupport.namedBothHeaderRows()`）：实测三份样例仍是 30／45／21 行、列名一字不变。
  `TableByNameTest` 与 `PrintResultTest.printLedgerNameRemedies` 把这些写法都钉住。
- **「表头集合未匹配到任何行」的提示少了一类噪声**：同一个 `name` 之下只要有一份注册命中了，
  另一份命不中就**不再提示**（`Tables.warnUnmatchedHeaders`，三种格式共用）。上面那条解法要求把上下两种写法
  都注册到同一个名字，其中一种在任何一份文档里都注定命不中，再提示就是噪声。没起名字的注册不受影响 ——
  pdf 里不存在的 `故障部件` 照样提示（`TableByNameTest.aRegistrationCoveredByASameNameMatchIsNoLongerReportedOnStderr`）。
- **顶层键顺序**默认是 `text` → `tables` → `images`；打开 `unmatchedHeaders(true)` 才多第四个键，且排在最后。
  关闭时结果与改动前逐字节相同（docx 上「注册五组但不起名」的 JSON 与静默模式逐字节相同）；
  打开后其余三个键的内容一字不变。没起名的未命中项只有 `headers` 一个键，起过名字的再带 `name`。

### 测试

- **183 → 242 条**（+59），测试类 **10 → 15**。主要变化：新增 `PdfBorderlessTableTest` 5、
  `TableByNameTest` 11、`PdfMergedCellTableTest` 5、`PptParserTextTest` 2（直接测包级 `textLines`，
  评估模式读不到 PPT 的多段文本）；`CrossFormatTest` 9→15、`OptionsTest` 10→14、`PrintResultTest` 5→17、
  `TablesTest` 31→34、`PptParserTest` 17→19、`WordParserTest` 31→32、`ImagesTest` 24→25、
  `PdfCrossPageTableTest` 3→4、`WordEdgeCasesTest` 11→14；`PdfParserTest` 17、`ResultsTest` 28
  条数未变但都补了断言。
  本轮 JDK 8 与 JDK 21 各跑完整套件都是 **242 条、0 失败、0 错误、1 跳过**（跳过的是 `-Ddoc` 入口）。
- **新增 `CrossFormatTest.aMissingFileThrowsFromAllThreeParsers`**：三个解析器对不存在的文件都把路径直接交给
  Aspose，没有自己包一层。实测消息都是 `Can't find file: <路径>.`，但类型各不相同：docx
  `java.io.FileNotFoundException`、pdf `com.aspose.pdf.internal.ms.System.l5if`、pptx
  `com.aspose.slides.exceptions.FileNotFoundException`；想统一 `catch (java.io.FileNotFoundException)` 只有
  docx 兜得住。`PrintResultTest.printErrorPaths` 把六条错误路径原样打印（另外三条是 FILE 模式缺 `imageDir`，
  三个解析器统一抛 `IllegalStateException`）。
- **打印类扩展到 17 条用例、18 份报告**（`PrintResultTest`，默认跳过 `printAnyDocument`），按用途分四类：
  一、三份样例静默各一份；docx「注册 + DROP」一份；三份「注册 + 起名」三份（共 7 份）；
  二、公共接口逐个跑四份：`printPublicApiOverloads`、`printOptionDefaultsAndSetters`、
  `printSerializationAndImageExport`、`printNamedMatchesAcrossFormats`；
  三、表头规则与选项矩阵五份：`printLedgerNameRemedies`、`printOptionMatrix`、`printAllModeJson`、
  `printUnmatchedHeaders`、`printHeaderAndFieldRules`；四、图片与错误路径两份：`printImageModes`、`printErrorPaths`。
  所有报告只保留 `assertHeadersMatchRowKeys` 一条底线断言，因为契约破了输出也没有参考价值。
- **新增 `printPublicApiOverloads`**：三份样例 × 四个 `parse` 重载（路径 / 流 × 选项 / 关键字）逐一对拍，
  路径与流的 JSON 逐字符相同（三份样例全打 `true`）；关键字入参不是选项的别名 —— pdf 上用
  `片区/区域/销售额/回款率/备注` 时是 6 组 47 行（静默 45 行，多出的两行是没被关键字命中的表把表头当成了数据），
  docx／pptx 与静默相同。
- **新增 `printOptionDefaultsAndSetters`**：默认 `isSilent=true`、`fullRow=KEEP`、`images=META`、`source=true`、
  `unmatchedHeaders=false`；`headerTolerance` 没有公开 getter，用「备注写成说明」验证默认 1 命中、收到 0 落空并提示；
  `headerGroups(List)` 与逐组 `headers(...)` 的 JSON 逐字符相同；`name(...)` 只绑最近一次注册。
- **新增 `printHeaderAndFieldRules`**：重名列按**物理列号**补后缀（`销售额-5`、`回款率-6`，不是「第几个重名」），
  空表头补 `列N`；现造「数据 1 + 2、`SUM(ABOVE)` 缓存写 300.00」的 docx，读出来就是 300.00，缓存为空则给空串，
  证明解析器只读域缓存、不执行域代码。三份样例的合计行分别打在报告第 4 节
  （docx 2096.30／1796.10、pdf 两格为空、ppt 1586／1384），差异来自源文件本身。
- **新增 `printSerializationAndImageExport`**：三个类的 `toJson` 输出逐字符相同；`BASE64` 模式下
  docx 899370、pdf 58076、pptx 898156 个字符（只打长度）；`writeImages` 写出文件的 SHA-1 与 `contentId` 一致；
  `META` 结果调用时只提示、不创建目录、不写文件。
- `printNamedMatchesAcrossFormats` 按名字逐个打出命中哪几组、每组来源页与全部行、`rowsByName` 拼平后的行数，
  以及「拼平结果与按组接起来是不是同一批对象」（按引用比，三份样例 9 处全打 `true`，证明不复制），
  末尾再打三个取不到的名字（未注册／空串／null 都给空列表）。
  `printLedgerNameRemedies` 把 docx 那张两行表头台账的四种注册写法逐个打出来对比 —— 四种在 docx 上都命中、
  都是 13 行、标准错误都是空的：①② 靠多级表头合并完再比一次，③ 直接命中上层那一行，④ 先命中上层那一行、
  再由同名规则盖住。差异在别的格式上：pdf 只有合并后那一行（③ 命不中），pptx 上下两行都在但只注册上层时
  下层会漏进数据（③ 变 8 行）。`printAllModeJson` 把同一批三份样例 × 七种选项的**完整 JSON** 原样打出来，
  每个模式先给一行摘要（组数／行数、顶层键、标准错误行数），再把 21 份 JSON 各写一份 UTF-8 文件到
  `target/print/json/<样例名>-<模式别名>.json` —— 前面的报告只给 JSON 长度，要逐字段核对时看这一份。
  `printOptionMatrix` 是三份样例 × 七种选项的横向对照（组数／行数、JSON 长度、
  `name`、未命中组数、标准错误行数、顶层键）；`printUnmatchedHeaders` 直接对照开关前后与「只注册 `故障部件`、
  起名 `部件统计`」两份注册；`printImageModes` 打 META／BYTES／BASE64／FILE 四种模式的体积、字段与落盘路径，
  并另跑一遍「META + imageDir」确认只提示、不新增文件；`printErrorPaths` 打上面那六条错误路径。
  18 份报告各留一份 UTF-8 副本到 `target/print/`（`mvn clean` 会带走），合计 **191974** 个字符。
- **修正 4 处 doclint 标题错误**：`WordParser.Options.name(String)` 与 `unmatchedHeaders(boolean)` 的 javadoc
  正文里各有两个小节标题。成员 javadoc 的隐含起始层级比 `<h3>` 高一级，所以**每个注释里的第一个标题必须是
  `<h3>`、后面的小节再用 `<h4>`**：全写 `<h3>` 报「使用了意外的标题」，全写 `<h4>` 则第一个标题报
  「标题层级不连续」，`javadoc -Xdoclint:all,-missing`（JDK 8）两种写法都会报错。按「首标题 `<h3>`、其后再
  `<h4>`」改完实测 **0 错误**，构建与测试不受影响。
- 为了拼最小 pptx，`TestSupport` 里新增一套现造 pptx 的工具（`pptx()`／`pTbl`／`pTr`／`pTc`／`pTcSpan`／
  `pTcHMergeContinue`／`pTcVerticalStart`／`pTcVerticalContinue`），母版、版式、主题从仓库样例里拷，
  只重写 `[Content_Types].xml`、`ppt/slides/`、`ppt/media/` 与 `presentation.xml` 的关系 ——
  上面那两条 PPT 修复就是靠它才造得出「纵向合并跨两行」「跨整行的空白合并行」这两种版面。

### 构建产物

- 本机 `mvn -B clean package`（JDK 8）：**BUILD SUCCESS**，**02:10**，
  **242 条、0 失败、0 错误、1 跳过**；
  主 jar **60236 字节 / 30 个条目 / 17 个 class**、`major version: 52`；`javac -Xlint:all -Xpkginfo:always`
  **0** 告警（17 个 class，含 `package-info.class`）、`javadoc -Xdoclint:all,-missing` **0** 错误；
  `mvn -B test` 不加 `-DargLine` 也是 **242/242**，那一行只影响控制台与报告里的中文编码。
  JDK 21 跑同一条 `clean package`：**02:17**、jar **60650 字节**、条目与 class 数同上、字节码仍是
  `major version: 52` —— 同一份源码换个 JDK 打包，字节数会差几百个，jar 大小别当校验和用。
  这两轮是 `1.2.0` 版本号下的发版构建；同一批源码在更安静的一轮里是 01:25（JDK 8）与 01:32（JDK 21），
  差值来自机器负载，口径仍是「只看数量级」。
  对照：Maven Central 上的 1.1.1 主 jar 56275 字节，差值全部来自本节列出的源码改动。
- **JDK 21 早先有一次 `clean package` 是 BUILD FAILURE，原因不在代码，需要如实记下**：当时测试数还是
  214 条，全部通过，却在 `maven-jar-plugin:3.2.2:jar` 阶段报
  `Error assembling JAR: cn\net\tianxi\document\Tables.class not found`。
  查 `target/classes` 时发现只有 `Tables.class` 的时间戳是 08:37，其余全是 08:35 ——
  那正是我用 `cp` 还原 `Tables.java` 的时刻，IDE 在构建中途把这单个文件重新编译了一次，与 Maven 的 jar 阶段撞车。
  原样重跑同一条命令即 BUILD SUCCESS。**结论是环境竞态、不是缺陷**，但复现构建时请先关掉 IDE 的自动编译。
- **「整轮比 1.1.1 那轮（`01:04`）慢约 80 秒、单类 31.5 秒 → 76.5 / 72.4 秒」这条已撤回**：
  那是拿抢 CPU 的一轮（同一条命令当时测到 02:15 ～ 02:37）与空闲的一轮并列比出来的。
  安静环境下上一轮 `clean package` 实测 01:25（JDK 8）与 01:40（JDK 21），同一 JVM 里把
  `documentedJsonLengthsAreTheRealOnes` 开／关**交替**各跑三轮为 34.6／30.9／37.6 秒 对
  29.8／35.9／38.2 秒，差值被轮次噪声完全吃掉。能精确归因的只有 surefire 报告里的单条耗时
  **3.3 ～ 3.8 秒（JDK 8，两轮）／3.2 秒（JDK 21）**（`target/surefire-reports/TEST-*.xml`）。
  所以文档里所有耗时都改成「区间 + 只看数量级」的口径。

### 文档与发布记录（`v1.1.1` 之后）

- 新增英文版 [README.en.md](README.en.md) 与 GitHub Pages 项目主页 [docs/index.html](docs/index.html)（已开启并验证可访问）。
- README 增加与 OCR 的路线对比、这样做的好处、应用场景三节（§1.1–1.3）；**不给 OCR 的准确率或耗时数字** ——
  仓库里没跑过任何 OCR 引擎。
- §1.1 的耗时区间改为只用 JDK 8 的两轮实测值，JDK 21 的数字单独列出。
- 署名写法补上 Gitee 镜像地址（GitHub 主仓库 + Gitee 镜像并写）。
- 记录 1.1.1 在 Maven Central 的发布结果与实测验证（Deployment PUBLISHED、全新空仓库 `dependency:get` 拉通、
  sha256 三处一致）。
- 发布坐标定为 `cn.net.tianxi:document-parser`；首次上传时命名空间还没验证完，先用 `io.github.hutianxi`
  发过一次，Central 的构件不能删改，那个坐标会永久留在中央仓库（详见 README §17）。
- 工作树里另有一份 Excel 解析选型报告 `docs/excel-parser-selection.md`（POI / EasyExcel / FastExcel / Aspose.Cells
  的横向实测，未提交、不进 jar）。
- `pom.xml` 还有三处**未提交**的发布元信息改动，会随下一个版本的 pom 一起发布，Central 上 1.1.1 的 pom 里是旧值：
  `<description>` 补上两个仓库地址、`<developers><email>` 改为 `mailtianxi@qq.com`、
  `<scm><connection>` 改为 Gitee 镜像（匿名只读地址），`<developerConnection>` 与 `<url>` 仍指 GitHub 主仓库。
  版本号本身仍是 `1.1.1`，没升。
- 新增本文件与英文对照版。
- 当时把两版 README、两版本文件与主页统一回填到 232 条 / 14 类 / 60099（JDK 8）与 60500（JDK 21）、
  起名后 4828／6561／3466、13 份打印报告，并把「合并后再比一次」「`unmatchedHeaders`」「表上方正文」写进对应章节。
- **收尾轮（工作树，未提交）**：`PrintResultTest` 补到 **16 条用例 / 17 份报告**（新增公共入口四重载对拍、
  `Options` 默认值与 setter、表头补名与域缓存、序列化与图片导出），新增 [DESIGN.md](DESIGN.md)
  （模块职责、数据流、表头判定、三格式适配、边界与取舍），两版 README、CONTRIBUTING、两版本文件与主页
  回填到 **236 条 / 14 类 / 17 份报告 / 85948 字符 / 01:32 ～ 01:42（JDK 8）与 01:55（JDK 21）**。
- **打印补全（工作树，未提交）**：`PrintResultTest` 新增 `printAllModeJson`（三份样例 × 7 种选项的完整 JSON，
  另存 21 份 `.json`），用例 **16 → 17**、报告 **17 → 18**；两版 README、CONTRIBUTING、DESIGN、
  两版本文件与主页回填到 **237 条 / 14 类 / 18 份报告 / 191974 字符 / 01:25（JDK 8）与 01:40（JDK 21）**。
- **入库示例（工作树，未提交）**：新增 [docs/json-to-db.md](docs/json-to-db.md)（表头签名路由、
  `ValueType` / `ColumnSpec` / `TableSpec` / `JsonToDb` 完整实现、docx/pdf/pptx 映射表、
  Spring/MyBatis 与同 JVM 直达、检查单与已知边界）；示例代码不进构建，在 JDK 8 + Jackson 2.13.5
  下编译通过。两版 README 的目录树与对应段落、主页的文档卡已加入口。

### 发版前收尾（2026-10-23）

本轮只做两件事：修掉上一轮清单里点名的三个缺陷并各补回归测试；把与项目无关的会话产物清出工作树。
`1.2.0` 的范围因此定为「新 API（`source` / `name` / `unmatchedHeaders` / `tablesByName` / `rowsByName`）
+ 这三个修复」。发版时 `pom.xml` 升到 `1.2.0`，本节其余内容不变。

**三个修复**（详细成因见上文「修复」，每条都是先写一条会失败的测试、再改代码）：

- **PDF 跨页合并**：`mergeSplitTables` 并完一截要把 `page`／`bottomGap`／`lastOnPage` 一起换成后一截的；
  新样例补上 `lastOnPage` 之后从 **1 组变 2 组**，负对照（只删这一行）立刻退回 1 组。测试
  `PdfCrossPageTableTest.theChainAlsoStopsWhenANonTableFlowFollowsTheFragment`。
- **PPT 多段文本**：拆行抽成包级 `textLines`，`\r\n`／`\r`／`\n` 三种都拆、空段不占行；
  `PptParserTextTest` 2 条单测直接钉住（端到端样例在评估模式下照不到多段）。
- **Word 嵌套深度**：提示只在真的遇到第 9 层子表时打一次，`walk` 先判有无子节点再判深度；
  8 层实测 **8 组、标准错误为空**，9 层 1 条提示、前 8 组照常。测试 `WordEdgeCasesTest` 两条。

**测试**：**237 → 242 条（+5）**，测试类 **14 → 15**。`WordEdgeCasesTest` 12→14、
`PdfCrossPageTableTest` 3→4、新增 `PptParserTextTest` 2；其余 12 个类条数不变。
最终 `mvn -B clean package` 在 JDK 8 与 JDK 21 上各跑一遍，都是 **242 条、0 失败、0 错误、1 跳过**
（跳过的是 `-Ddoc` 入口）。逐类条数取自 `target/surefire-reports/TEST-*.xml`。

**文档与注释**：`DESIGN.md` §11 的分类计数改成本轮实测（补上 `PptParserTextTest` 2、
`WordEdgeCasesTest` 14、`PdfCrossPageTableTest` 4）；`Images.java` 里引用的 `ImagesTest` 条数改回实际的 25 条；
`WordParser` 与 `package-info` 的「嵌套深度」说法改成与新行为一致（第 9 层起跳过、提示只在真要跳内容时出现）；
`PdfParser` 里指向旧测试名的注释同步改名。新加的 javadoc 小节标题按 JDK 8 doclint 的规则排
（成员注释里第一个标题必须是 `<h3>`、其后用 `<h4>`），`javadoc -Xdoclint:all,-missing` **0 报错**。

**清理**：`.qoder-credits/`（Qoder 会话产物，与项目无关）已删除；`.gitignore` 新增
`.qoder-credits/`、`.aoci/`、`AGENTS.md`、`aoci*.txt`、`.gitattributes` 与 `.mcp.json*`。
后五类按作者要求只留本地、不入库，`git status --ignored` 里都是 `!!`。

**最终校验**：`1.2.0` 版本号下，JDK 8 `mvn -B clean package` **BUILD SUCCESS**、**02:10**，主 jar
**60236 字节 / 30 个条目 / 17 个 class**、`major version: 52`；JDK 21 同一条命令 **242/242**、**02:17**、
jar **60650 字节**，条目与 class 数同上；`javac -Xlint:all -Xpkginfo:always` **0** 告警、
`javadoc -Xdoclint:all,-missing` **0** 错误。

发版准备到此完成：版本号已升、两套 JDK 都跑过完整构建；**提交、打 `v1.2.0` tag 与上传 Maven Central
仍未执行**，发布记录待上传成功后补进 README §17。

---

## 1.1.1 — 2026-10-22

首个公开发布版本（tag `v1.1.1`，提交于 2026-10-22 02:23:09 +0800）。

- **一套结构三种格式**：`WordParser` / `PdfParser` / `PptParser` 方法签名一一对应，
  返回 `Map<String, Object>`，键固定是 `text` / `tables` / `images`，换格式只换类名。
- **表头分组**（`Tables`，唯一一份，三种格式共用）：六种判定依据（注册表头集合 → 关键字 → 格式自带标题行标记 →
  紧跟分组行之后的行 → 表样式声明的首行 → 静默模式首行兜底）、两行表头合并、跨整行合并行的三种去向
  （`FullRow.KEEP` / `FIRST` / `DROP`）、合并单元格一律拆开回填、列名补全（空表头 → `列N`，同名列 → `名-列号`）。
  契约：`headers` 与每一行的键逐个相等。
- **PDF**：`TableAbsorber` 几何还原有线表、跨页长表并回一张（四条续接判据）、纵向合并被切成几块的行按矩形相邻聚回。
- **图片**：四种输出模式（`META` 默认 / `BASE64` / `BYTES` / `FILE`）、SHA-1 内容指纹 `contentId`、
  内嵌模式重复图片不重复内嵌、外链图片只记路径不联网、`writeImages` 导出。
- **文本清洗**（`Results`，唯一一份）：域代码丢弃、域缓存结果保留，软换行与分页符换成空格，
  控制字符清除，Aspose 评估水印过滤；许可证每个产品一个进程只注册一次。
- **对外部文档的安全处理**：写文件前只取文件名部分、目录内重名加序号、外链不读取不联网、递归深度有上限。
- **测试与产物**：183 条测试 / 10 个测试类（1 条 `-Ddoc` 入口平时跳过），主 jar 56275 字节，Java 8 字节码。
- **许可与发布**：Apache License 2.0 + `NOTICE` 署名，Aspose 三件套按 Aspose EULA、仓库不随包分发其二进制；
  发布到 Maven Central `cn.net.tianxi:document-parser:1.1.1`（sources 64964 字节、javadoc 141571 字节），
  GitHub Release `v1.1.1` 附三个包，GitHub 主仓库与 Gitee 镜像同步，两边 `main` 均设保护分支。
