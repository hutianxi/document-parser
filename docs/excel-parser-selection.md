# Excel 解析选型报告与 `ExcelParser` 实施计划

> 结论先行：**主干用 Apache POI 5.5.1，xlsx 走事件流（SAX）读、`<mergeCells>` 自己在同一次遍历里取，usermodel 只作为 `.xls` 与小表的兜底；Aspose.Cells 不进主干，留作可选后端；EasyExcel / FastExcel 的**默认无模型读法**拿不到合并区（实测塌成 1 组 / 6 行），本次原写"拿不到合并区坐标、不采用"——这条否决理由已于 2026-10-22 复核文档时被推翻：两个库都有 `extraRead(CellExtraTypeEnum.MERGE)` 通道，能给出 `firstRowIndex/firstColumnIndex/lastRowIndex/lastColumnIndex`（见第 9 节第 12 条）。所以 Excel 后端最终选 POI 事件流还是 Fesod，要靠那次复验决定，本报告第 1 节矩阵与第 8 节计划里涉及 EasyExcel 的判断在复验前都只算"待复核"。**
>
> 本文所有数字都来自 2026-10-22 在本机跑的一次性探针（`ExcelProbeTest`，跑完即删，未进仓库），JDK 1.8.0_191，Windows，16 核。凡是没实测过的，都在第 9 节明确列出，不猜数字。

## 0. 这次调研怎么做的

探针用代码现造两份 xlsx，再用六条路线分别读同一批文件，把每一格的取值、组数、行数、耗时全部打到 `target/excel-probe/report.txt`：

`structured.xlsx`（7 470 字节，7 张表，专门覆盖项目里已经踩过的结构）

| 表 | 里面放了什么 |
| --- | --- |
| 季度销售明细 | 纵向合并 `A2:A3`、`A4:A5`；冻结首行 `createFreezePane(0, 1)`；两组同名列（销售额 ×2、回款率 ×2） |
| 客户回款台账 | 两行表头：`C1:D1` 横向合并 + `A1:A2`/`B1:B2`/`E1:E2` 纵向合并；`B6=SUM(B3:B5)`、`C6=SUM(C3:C5)` 公式格 |
| 分片区年度目标 | 跨整行合并 `A1:D1`、`A5:D5`（分组小标题），同一张物理表里两组表头 |
| 格式与取值 | 中文日期格式、ISO 日期、`#,##0.00`、`0.00%`、`000` 补零；一格文本一格数字混排 |
| 两张表同页 | `A1:C1` 合并标题 + 第 4 行整行不存在（不是空单元格，是没有 Row）+ 表头重复出现两次 |
| 空表 | 一张行数为 0 的表 |
| 隐藏表 | `state=HIDDEN`，里面有一行 |

`big.xlsx`（2 452 171 字节，SXSSF 生成，100 000 行 × 7 列，700 000 个非空格）用于耗时与内存。

六条路线：① POI usermodel（DOM）；② POI 事件流（`XSSFReader` + `XSSFSheetXMLHandler`）；③ EasyExcel 4.0.3 无模型 Map；④ FastExcel 1.3.0 无模型 Map；⑤ Aspose.Cells 23.12（无 lic，评估模式）；⑥ 手写解 OOXML（`java.util.zip` + SAX，未实现，只在第 7 节评估）。

关键约束是项目自己的：`Tables.group(String[][] rows, boolean[] separator, boolean[] heading, Options, Set<Integer> matched, boolean firstRowStyled)` 要求调用方先交出一张**合并区已回填**的网格，外加「哪些行是跨整行合并」的 `separator[]`。所以选型的第一判据不是速度，而是**这条后端能不能把这两样东西给出来**。

## 1. 评估矩阵

| 维度 | POI 5.5.1 | EasyExcel 4.0.3 | FastExcel 1.3.0 | Aspose.Cells 23.12 | 手写 OOXML |
| --- | --- | --- | --- | --- | --- |
| 合并区坐标 | ✅ `getMergedRegions()` 实测给出 `A2:A3 A4:A5` | ⚠️ 默认无模型读法给不出；反射查 `XlsxReadSheetHolder` 无 merge 方法。**但文档另有 `extraRead(CellExtraTypeEnum.MERGE)` 通道可拿到坐标，本次探针没启用，待复验（第 9.12 条）** | ⚠️ 反射查 `XlsxReadSheetHolder` 同样无 merge 方法；但 1.3.0 源码含 `MergeCellTagHandler`/`MergeCellsRecordHandler`，与 EasyExcel 同走 `extra()` 通道，探针未启用，待复验 | ✅ `getMergeCells()` 返回 4 个 `CellArea` | ⚠️ 要自己解析 `<mergeCells>` |
| 事件流里的合并区 | ⚠️ `XSSFSheetXMLHandler` **不回调** `mergeCell`；实测挂 `XMLFilter` 数到 2/4/2/0/1 个元素，监听器一个也收不到 | — | — | — | ✅ 本来就是自己解析 |
| 分组行（`separator`）线索 | ✅ 由合并区自己算 | ⚠️ 默认读法算不出；若 `extra()` 通道能拿到坐标即可自算（待复验） | ⚠️ 同左 | ✅ 同 POI 算法可自算 | ⚠️ 同手写解析 |
| 表头线索 | ✅ `getPaneInformation()` 实测 `isFreezePane=true hSplitTopRow=1` | ⚠️ 探针未测 | ⚠️ 探针未测 | ✅ `getFreezedPanes()=[1,0,1,0]`、`paneState=0` | ⚠️ 自己读 `<pane>` |
| 隐藏表识别 | ✅ `visibility=HIDDEN` | ⚠️ 能读到内容（实测 1 行），可见性入口探针未测 | ⚠️ 同左 | ✅ `visible=false` | ⚠️ 自己读 `workbook.xml` |
| 空表 | ✅ `lastRowNum=-1` | ✅ 回调 0 次（表名保持初始值） | ✅ 同左 | ✅ `maxDataRow=-1` | —（未实现） |
| 显示值（带格式） | ⚠️ 中文日期给成 `2025"年"12"月"31"日"`，**引号漏出来** | ✅ `2025年12月31日` | ✅ 同左 | ✅ `2025年12月31日` | ⚠️ 要自己实现格式引擎 |
| 原始值 | ✅ `46022.0` / `0.15` | ⚠️ Map 模式只有字符串 | ⚠️ 同左 | ✅ `getStringValueWithoutFormat()`=`46022`、`getValue()`=`2025-12-31T00:00:00` | —（未实现） |
| 公式格 | ⚠️ `DataFormatter.formatCellValue(cell)` 返回**公式文本** `SUM(B3:B5)`；带 evaluator 或 `getNumericCellValue()` 才是缓存值 596.8 | ✅ `596.8`（缓存值） | ✅ `596.8` | ✅ `596.8` | ⚠️ 自己取 `<v>` |
| 100 000 行耗时 | 事件流拼网格 1.97–2.14 s；DOM 3.46–5.92 s | 1.07–1.20 s | 1.12–1.87 s | 打开 0.60–0.68 s + 遍历 0.47–0.61 s | 预期与事件流同档（未实现，见第 9 节） |
| −Xmx256 MB 能否读完 | ❌ DOM **OutOfMemoryError**（6.0 s 崩）；✅ 事件流读完，峰值 129.3 MB | ✅ 读完，峰值 156.5 MB | ✅ 读完，峰值 154.5 MB | ⚠️ 没崩但极慢：120 s 只走完 5 620 / 99 999 行 | ⚠️ 未测 |
| 自带 jar 体积 | 18.6 MB（含全部传递依赖） | 0.96 MB + 依赖 POI ⇒ 约 19.6 MB | 1.44 MB + ehcache/jaxb + POI ⇒ 约 22 MB+ | 8.35 MB（**零传递依赖**） | 约 0 MB |
| 字节码级别 | major 52（Java 8） | major 52 | major 52 | major 50（Java 6） | — |
| 许可 | Apache-2.0 | Apache-2.0 | Apache-2.0 | 商业订阅，**与 Words/PDF/Slides 分产品线，需另购** | 自有代码 |
| 上游活跃度（Maven Central 元数据） | 最新 5.5.1，元数据更新 2025-11-30 | 最新 4.0.3，元数据更新 **2024-09-11（约两年没发版）** | 最新 1.3.0，更新 2025-08-25（EasyExcel 的分家续作） | 已到 26.10，更新 2026-10-16 | — |
| 与现有代码风格 | 项目里没有任何 POI 代码，需新增一套 | 监听器回调模型，与 `Tables.group` 的「整张网格」模型不搭 | 同左 | 与 Aspose.Words/PDF/Slides 同族（`License.setLicense`、`Results` 的水印过滤已就位） | 全自己维护 |

一句话读法：**能直接交出 `Tables.group` 要的合并区坐标的只有 POI usermodel 和 Aspose.Cells 两条，这两条给出的分组结果逐字符相同（第 2.2 节实测）；EasyExcel 与 FastExcel 在默认读法上是空白，读同样的文件会塌成"全是 `列N`、组数不对"（它们是否有可用的合并区通道见第 9.12 条，待复验）；POI 事件流与手写 OOXML 需要自己在同一次遍历里收 `<mergeCells>`，做得到但要多写代码。** 速度差距（1 秒级）不是决策点，256 MB 档谁能读完才是（第 4.3 节）。

## 2. 结构能力实测

### 2.1 合并区：三条路线给的东西完全不同

同一张「客户回款台账」，文件里 4 个合并区：

```
POI  getMergedRegions()  → A1:A2  B1:B2  C1:D1  E1:E2
Aspose getMergeCells()    → r0-0c2-3  r0-1c0-0  r0-1c1-1  r0-1c4-4   （类型 com.aspose.cells.CellArea）
EasyExcel / FastExcel     → ReadSheetHolder 上无参且名字含 merge 的方法 = 无
POI 事件流                → SAX 原始事件里 mergeCell 元素出现 4 次，但 XSSFSheetXMLHandler 一个都不回调给监听器
```

EasyExcel 的 Map 里，被合并覆盖的格子是**有键、值为 null**（实测 `行0 {0=客户名称, 1=合同金额, 2=回款情况, 3=null, 4=备注}`、`行1 {0=null, 1=null, 2=已回款, ...}`）。看起来像"信息还在"，但没有任何一处能告诉你那个 null 是"被合并盖住了"还是"本来就是空的"——两者在输出里必须走不同分支（前者回填，后者留空）。

### 2.2 分组结果一致性：POI 与 Aspose 给出逐字符相同的结果

把三份网格（POI usermodel、Aspose、EasyExcel 原样）分别喂进同一个 `Tables.group`，用同一套注册表头跑：

| 表 | POI usermodel（回填） | Aspose.Cells（回填） | EasyExcel 原样网格 |
| --- | --- | --- | --- |
| 季度销售明细 | 1 组 / 4 行，`[片区, 区域, 销售额, 回款率, 销售额-5, 回款率-6, 备注]` | 1 组 / 4 行，同上 | 未测（探针只把台账的 EasyExcel 网格喂进分组算法；此表 EasyExcel 给出的 `行2 {0=null,…}` 说明合并覆盖格同样是 null） |
| 客户回款台账 | 2 组 / 5 行，命中注册下标 `[1]`，第二组 `[客户名称, 合同金额, 已回款, 回款率, 备注]` / 4 行 | 2 组 / 5 行，同上，逐字一致 | **1 组 / 6 行**，列名退化成 `[列1…列5]`，注册表头一个都没命中（下标 `[]`） |
| 分片区年度目标 | 3 组 / 5 行，分组行下标 `0,4`，命中 `[2, 3]`，`DROP` 后 2 组 / 3 行 | 3 组 / 5 行，同上 | 未测（此表主要考合并区，EasyExcel 拿不到坐标即等于"不回填"档） |
| 两张表同页 | 3 组 / 4 行，命中 `[5]`（真空行不当分组行时 2 组 / 5 行） | 3 组 / 4 行，同上 | 未测组数；实测行下标 `[0,1,2,4,5]`——文件里第 4 行整行不存在，EasyExcel 直接不回调，`separator[]` 与行数得由调用方自己补 |
| 空表 | 0 组 | 0 组 | 0 回调 → 0 组 |
| 隐藏表 | 读到 1 行 | 读到 1 行（`visible=false` 可判断） | 读到 1 行（无可见性入口） |

台账这一行是本次最硬的证据：**同一份文件、同一个分组算法，只差「合并区有没有回填」，结果从 2 组 / 5 行 + 正确列名，塌成 1 组 / 6 行 + 全是 `列N`。** 换句话说，选后端不是在选速度，是在选"能不能拿到合并区"。

另外补了一条反向对照：POI 网格如果**不回填**（只保留左上格），台账退化成 `1 组 / 6 行`、`[列1…列5]`，与 EasyExcel 的表现一模一样。所以这不是 POI 的独家优势，是"合并区坐标"的独家优势——本次探针用的是 EasyExcel/FastExcel 的默认无模型读法，拿不到坐标就只能塌；它们另有 `extra()` 通道，本次没测，第 9.12 条复验之前不要把这句当成"永远补不出来"。探针试过按"相邻格值相同"猜合并，实测无效（`手工回填 + 注册表头` 仍是 1 组 / 6 行，因为空字符串彼此相同、非空值又不相邻）。

### 2.3 显示值与原始值

| 单元格（文件里的格式） | POI `DataFormatter` | POI 原始 | EasyExcel / FastExcel | Aspose `getStringValue` | Aspose `getStringValueWithoutFormat` |
| --- | --- | --- | --- | --- | --- |
| 日期 `2025"年"12"月"31"日` | `2025"年"12"月"31"日"` ❌ 引号漏出 | `46022.0` | `2025年12月31日` ✅ | `2025年12月31日` ✅ | `46022` |
| 日期 `yyyy-mm-dd` | `2026-01-05` | `46027.0` | `2026-01-05` | `2026-01-05` | `46027` |
| 数字 `#,##0.00` | `1,280.50` | `1280.5` | `1,280.50` | `1,280.50` | `1280.50` |
| 数字 `0.00%` | `15.00%` | `0.15` | `15.00%` | `15.00%` | `0.15` |
| 数字 `000` | `007` | `7.0` | `007` | `007` | `7` |
| 公式 `SUM(B3:B5)` | **`SUM(B3:B5)`** ❌ 公式文本 | `596.8`（缓存值） | `596.8` | `596.8` | `596.80` |

两个坑都有实测支撑：

1. `DataFormatter.formatCellValue(cell)` 对公式格返回的是**公式字符串**，必须走 `formatCellValue(cell, evaluator)` 或直接 `getNumericCellValue()`（POI 默认不重算，读的是文件里的缓存结果，实测 `getCachedFormulaResultType=NUMERIC`、缓存值 596.8，显式 `evaluate()` 也是 596.8）。
2. `getStringCellValue()` 对数字格抛 `IllegalStateException`（实测三处：日期格、百分比格、公式格），`getDateCellValue()` 对纯数字 `1280.5` 会给出 `Fri Jul 03 12:00:00 CST 1903` 这种荒谬日期。所有取值都要按 `getCellType()` 分派，不能一路 `getStringCellValue()` 走到底。

POI 的中文日期引号问题（`2025"年"12"月"31"日"`）源于 `DataFormatter` 按 Excel 规则把格式里的字面量加引号；Aspose 和 EasyExcel 都没有这个行为。若要沿用 POI，需要在 `Results.cleanContent` 那条清洗链上加一条"成对的引号包单个汉字/拉丁字去引号"，属于可解但要写代码+样例的部分（本次未实现，见第 9 节）。

## 3. 事件流的一条硬坑（必须先知道再选型）

第一轮探针跑出来是「0 行 0 格」——`XSSFSheetXMLHandler` 收到事件但一个回调都没发。原因不在数据，在 SAX 工厂配置：

```java
SAXParserFactory factory = SAXParserFactory.newInstance();
factory.setNamespaceAware(true);   // 少了这一行，整张表被静默跳过
```

实测对照（同一份文件、同一段代码，只切这个开关）：`namespaceAware=false` → `endRow` 次数 **0**、`cell` 回调数 **0**；`namespaceAware=true` → `endRow` 5 次、`cell` 33 次（第 0 张表实际就是 5 行 33 格）。`XSSFSheetXMLHandler.startElement` 按 `spreadsheetml` 命名空间过滤事件，非命名空间感知时 `uri` 是空串，于是全部提前返回。

这类坑的可怕之处在于**它不报错，只给空结果**。写进选型结论：事件流路线必须在测试里断言行数与格数，不能只看"没抛异常"。

## 4. 性能与内存

### 4.1 耗时（100 000 行 × 7 列，2 452 171 字节，JDK 1.8.0_191）

同一条路线跑了两轮独立进程，区间如实列出（POI DOM 两轮差 1.6 秒，不平均，因为平均会藏住 GC 抖动）：

| 路线 | 第 1 轮 | 第 2 轮 | 备注 |
| --- | --- | --- | --- |
| POI usermodel + `DataFormatter` | 5 919 ms | 3 567 ms | 全表 700 000 格 |
| POI usermodel 只取字符串 | 3 463 ms | 3 579 ms | 同上 |
| POI 事件流，带 formatter（只统计） | 2 417 ms | 2 302 ms | 行 100 000 / 格 700 000 |
| POI 事件流，不带 formatter | 1 460 ms | 1 416 ms | 同上 |
| **POI 事件流 → 拼 `String[][]` 网格** | 2 141 ms | 1 973 ms | `ExcelParser` 真正要做的事 |
| EasyExcel 无模型 Map | 1 201 ms | 1 074 ms | 回调 100 000 行 |
| FastExcel 无模型 Map | 1 865 ms | 1 118 ms | 回调 100 000 行 |
| Aspose.Cells 打开工作簿 | 676 ms | 599 ms | `maxDataRow=99999 maxDataColumn=6` 正确 |
| Aspose.Cells 逐格全表遍历 | 470 ms | 606 ms | **含打开**（打开分别 401 / 536 ms） |

怎么读这张表：

- 快慢差距存在，但**都不是瓶颈**：最慢的 DOM 也就 3.5–5.9 秒读 70 万格，业务上完全可用。选型不该被 1 秒级的差距牵着走。
- Aspose 的耗时几乎全在"打开"，之后逐格很便宜（打开 401 ms / 遍历 ~69 ms）。这说明它在打开时就把整张表物化了——**这也正是它在小堆下翻脸的原因，见 4.3 节 E 组**。
- **`DataFormatter` 才是 POI 侧最贵的一步**：同一份文件、同一套事件流，带 formatter 2.302–2.417 s，不带 1.416–1.460 s，差 0.886–0.957 s（相当于在不带 formatter 的基础上 **+63%~66%**；formatter 自身占带 formatter 总耗时的 38%~40%）。所以「显示值」这件事在 POI 上是有账单的。
- **小表没有单独计时**：探针只给 100 000 行这一档打了耗时，`structured.xlsx`（7 张表、最大 7 行）的输出里没有耗时列。按上面 70 万格 / 1.4–5.9 秒这一档推算，几十行的表是毫秒级，**但这是推算，不是实测**；真要写进 README 的基线，得另跑一组。

### 4.2 堆占用（同 JVM 顺序跑，只作量级参考）

| 路线 | 读完瞬间已用堆（第 1 轮 / 第 2 轮，与 4.1 同两轮） |
| --- | --- |
| POI usermodel + formatter | 748.4 / 832.6 MB |
| POI usermodel 只取字符串 | 840.9 / 696.0 MB |
| POI 事件流拼网格 | 995.3 / 805.1 MB |
| EasyExcel | 778.3 / 1 033.7 MB |
| FastExcel | 1 242.0 / 1 022.5 MB |

这一组数**不能横向比较**：六条路线在同一个 JVM 里顺序跑，后跑的继承前一轮的残留对象（探针还把每轮的行都缓存在 `List` 里以证明"读完了"）。`maxMemory=3 580.5 MB` 下都不 OOM。真正有意义的对照是下面按独立 JVM + 固定上限重跑的一组。

### 4.3 内存上限分组（`-Xmx256m`，每条一个独立 JVM，峰值由 5 ms 采样线程得到）

| 组 | 结果 |
| --- | --- |
| A POI usermodel + `DataFormatter` | ❌ **`OutOfMemoryError`**：6 007 ms 时崩，采样峰值 199.8 MB（上限 228.0 MB），10 万行没读完 |
| B POI 事件流 → 拼 `String[][]` 网格 | ✅ 读完，2 085 ms，峰值 129.3 MB |
| C EasyExcel Map（缓存全部 10 万行） | ✅ 读完，2 120 ms，峰值 156.5 MB |
| D FastExcel Map（同上） | ✅ 读完，1 921 ms，峰值 154.5 MB |
| E Aspose.Cells 逐格全表 | ⚠️ **120 秒上限内只走到第 5 620 行 / 共 99 999 行**（非空格 39 347），采样峰值 216.6 MB（上限 255.5 MB）—— 没崩，但在疯狂 full GC |
| F POI 事件流只统计（不缓存行） | ✅ 读完，1 728 ms，峰值 **85.1 MB**（上限 228.0 MB），行 100 000 / 格 700 000 |

这组才是决策依据，三件事一次说清：

- **同一个文件、同一个 256 MB 上限，POI 的 DOM 路线直接崩，事件流路线用不到 130 MB 读完全表。** 所以 `ExcelParser` 的大表路径必须是事件流；usermodel 只能留给 `.xls`（POI 的 HSSF 没有等价的事件流方案，或者说要另写一套）和"表很小、需要随机访问合并区"的场景。
- **Aspose.Cells 在 4.1 表里那 0.47 秒是有前提的**：堆给够。同一份文件放进 256 MB，120 秒只走完 5.6% 的行；而 POI 事件流在同一档位 1.7 秒读完、峰值 85 MB。它"打开即物化整表"的特征在大表 + 小堆下会翻脸。
- **不缓存行的代价只有 85 MB**（F 组），缓存成网格要 129 MB（B 组）——两者相差 44 MB，即 70 万个字符串本身。

注意 B 组（129.3 MB）里含 70 万个字符串组成的网格本身——这本来就是"把整张表交给 `Tables.group`"的必要代价。真要再省，就得给 `ExcelParser` 加"按区块流式分组"，那是算法层的改造，不是后端选型能解决的（见第 8 节的分阶段计划）。

E 组的读数方式要说清楚：探针在 120 秒处主动跳出循环、把"走到第几行"报出来，所以它是**正常返回**而不是异常；这一档拿不到 Aspose 的完整耗时，只能给出"5 620 行 / 120 秒"这个下限式的描述。

## 5. 依赖体积、字节码与版本冲突

各 jar 实际体积（本地仓库 `D:/JavaDevelopTools/apache-maven-3.8.8/repo`）：

```
POI 路线合计 18.6 MB
  poi 2.87 + poi-ooxml 1.96 + poi-ooxml-lite 5.72 + xmlbeans 5.3.0 2.11
  + commons-compress 1.28.0 1.07 + commons-io 2.21.0 0.56 + commons-lang3 0.56
  + commons-codec 1.15 0.34 + commons-math3 3.6.1 2.11 + commons-collections4 4.5.0 0.86
  + curvesapi 1.08 0.11 + SparseBitSet 1.3 0.02 + log4j-api 2.17.2 0.29
EasyExcel 自带 0.96 MB（easyexcel 0.00 是个壳 + easyexcel-core 0.52 + easyexcel-support 0.44），依赖 POI ⇒ 总量约 19.6 MB
FastExcel 自带 0.96 MB + commons-csv 0.05 + ehcache 3.10.8 1.73（还拖 javax.cache 与 jaxb-runtime 2.3.9）+ 依赖 POI ⇒ 约 22 MB+
Aspose.Cells 8.35 MB，零传递依赖
```

对照现有产物：`aspose-words-jdk16` 16.7 MB、`aspose-pdf-jdk16` 77.6 MB、`aspose-slides-jdk16` 29.2 MB。**POI 的 18.6 MB 或 Aspose.Cells 的 8.35 MB 相对于当前依赖体积都不算新增负担。**

字节码 major version（`javap -v` 抽每个 jar 的第一个 class 实测）：`poi`/`poi-ooxml`/`poi-ooxml-lite`/`xmlbeans`/`commons-compress`/`commons-io`/`commons-collections4`/`easyexcel-core`/`easyexcel-support`/`fastexcel`/`fastexcel-support` = **52（Java 8）**；`curvesapi` = 51（Java 7）；`commons-math3` = 49（Java 5）；`aspose-cells` = **50（Java 6）**。全部 ≤ 52，与 `java.version=1.8` 的编译目标兼容，**不需要像 Words/PDF/Slides 那样挑 `jdk16` classifier**（`aspose-cells` 在 Central/Aspose 仓库上也没有 classifier 变体，只有 plain jar）。

版本冲突（`mvn dependency:tree -Dverbose` 实测，父 POM 是 `spring-boot-starter-parent:2.7.18`）：

```
commons-codec   POI 5.5.1 想要 1.20.0  → 被父 POM 管成 1.15   （version managed from 1.20.0）
log4j-api       POI 5.5.1 想要 2.24.3  → 被父 POM 管成 2.17.2 （version managed from 2.24.3）
commons-lang3   POI 5.5.1 想要 3.18.0  → 被父 POM 管成 3.12.0 （version managed from 3.18.0）
commons-io      poi-ooxml-lite 要 2.20.0 → 收敛到 2.21.0
EasyExcel 4.0.3 自带 poi / poi-ooxml 5.2.5 → omitted for conflict with 5.5.1
FastExcel 1.3.0 自带 poi / poi-ooxml 5.4.1 → omitted for conflict with 5.5.1
```

这三项降级是本次唯一的真实风险点，但**已经实测过**：探针跑的就是"降级后的 POI 5.5.1"（commons-codec 1.15 / log4j-api 2.17.2 / commons-lang3 3.12.0），五条实现路线（POI usermodel、POI 事件流、EasyExcel、FastExcel、Aspose.Cells）读 7 张表与 10 万行全部跑通，没有 `NoSuchMethodError`、没有 `IncompatibleClassChangeError`。若用户侧 Spring Boot 版本更低或更高，仍应在集成测试里重复这一条。

顺带一条架构事实：EasyExcel/FastExcel 都把 POI 作为自己的传递依赖（5.2.5 / 5.4.1），也就是说**引入它们并不能绕开 POI，只是把 POI 换成它们挑的版本**。这与"少一个依赖"的直觉相反。

## 6. 许可与维护

POI、EasyExcel、FastExcel 都是 Apache-2.0，与本项目 LICENSE 兼容，可以直接进 `pom.xml` 的 compile 作用域并随 jar 发布。

Aspose.Cells 是商业订阅产品，且**与项目现有的 Words/PDF/Slides 分属不同产品线**——现有 Aspose 授权不必然覆盖 Cells，需要单独确认/购买。仓库里已经有 `*.lic` 的 `.gitignore` 红线与 `Results.applyLicense` + 评估水印过滤这套基础设施，接 Cells 的工程成本很低，但商务成本要用户自己拍板。

本次 Aspose.Cells 全程在**无 lic 的评估模式**下测（`License.isLicenseSet()=false`）。实测痕迹：抽查 `structured.xlsx` 前 20 行共 103 个非空格，含 Aspose/评估水印字样的 **0 个**；`big.xlsx` 的 100 000 行也全部读到、末行 `R99999C0` 内容正确。Aspose 官方文档对评估模式的具体限制条款本次**没有去核对**，所以不能拿"没看到水印"当"没有功能限制"用（见第 9 节）。

维护状态用 Maven Central 元数据里的时间戳说话（不引用博客或 README 的口头说法）：

| 组件 | 仓库里的最新版本 | 元数据 `lastUpdated` | 距今 |
| --- | --- | --- | --- |
| `org.apache.poi:poi-ooxml` | 5.5.1 | 2025-11-30 | 约 11 个月 |
| `com.alibaba:easyexcel` | 4.0.3 | **2024-09-11** | **约 25 个月没有新版本** |
| `cn.idev.excel:fastexcel` | 1.3.0 | 2025-08-25 | 约 2 个月 |
| `com.aspose:aspose-cells` | 26.10 | 2026-10-16 | 6 天 |

EasyExcel 已经两年不出版本，官方站点首页声明项目已进入维护模式；它的社区续作 FastExcel（同班人马改组后的 `cn.idev.excel`）随后捐给了 Apache，更名 **Apache Fesod**（版本线 2.0.0 / 2.0.1 / 2.0.2 / 2.1.0-incubating，仓库最近 push 2026-10-22，Apache-2.0，孵化期，newcomer 坐标要在引入前核实）。顺带说明：外部常引用的 `github.com/meistangtao/fastexcel` 只是上游的一个个人 fork（3 star、建库当天最后一次 push、0 open issue），不代表这个项目本身，看能力要看 `apache/fesod`。这解释了为什么 FastExcel 与 EasyExcel 的 API 几乎一字不差、实测结果也一致（见第 2 节）。**如果将来一定要用这条线，应该选 Apache Fesod，而不是 EasyExcel，更不该盯那个个人 fork。** 至于本报告原先"两者都过不了合并区这一关"那句，2026-10-22 复核后降级为待复验，见第 9.12 条。

## 7. 关于「手写解 OOXML」这一条

`xlsx` 就是一个 zip 包，理论上是零依赖的正路。探针里为它验过两个前提：

- `<mergeCells>` 在 sheet XML 里就是普通元素，实测挂一个 `XMLFilter` 就能数到（第 0/1/2/3/4 张表分别是 2、4、2、0、1 个）——POI 事件流拿不到的东西，自己解析能拿到。
- 事件流本身可行（POI 的 `XSSFReader` 就是这条路，实测 100 000 行 1.4–2.1 秒、256 MB 档峰值 129 MB）。

但要看清"自己解析"具体要重造什么：共享字符串表、样式表与格式串（`numFmt`）、Excel 的日期序列号与 1900 闰年 bug、公式缓存值、富文本 run 拼接、`inlineStr`、`xlsm` 与 `xlsb` 的差异。**第 4.1 节那个"带 formatter 比不带慢 0.89–0.96 秒"就是这部分工作的市场价格**——`DataFormatter` 是 POI 用总耗时里 38%~40% 的份额换来的一个成熟实现，而这块正是手写路线要自己啃的。

结论：**不做。** 本项目已经是"重依赖"库（Aspose 三个产品 120 MB+），零依赖带来的收益不成立，成本却是一次性重建半个 POI。真要有"不许引入任何依赖"的场景，正确的切法是把 `Tables` 那一层单独抽成一个不依赖任何解析器的模块，而不是重写解析器。

## 8. `ExcelParser` 实施计划

### 8.0 先定契约

对外只加一个类，与现有三个解析器同形：

```java
ExcelParser.parse(String filePath)                       // 关键字简写，同 PptParser
ExcelParser.parse(String filePath, WordParser.Options)    // 复用现有 Options（headers/keywords/fullRow）
ExcelParser.parse(InputStream inputStream, ...)           // 见下面的限制
```

返回结构在现有 `{text, tables, images}` 上做两处扩展，且必须是**加键而不是改键**（`CrossFormatTest` 的断言只读每组的 `headers` 与 `rows`，见其 `TestSupport.headers/rows/values` 三个助手）：

1. `tables` 仍是一个平铺的分组列表，每组在 `Tables.group` 的 `{headers, rows}` 之外**追加一个 `sheet` 键**。Excel 的 sheet 对应 PPT 的 slide，但 PPT 的页位置在 `text` 里有上下文、Excel 的 sheet 名丢了就找不回来；`WordParser`/`PptParser` 的结果里没有这个键，属于 Excel 独有的追加项。
2. `text` 放「按 sheet 分节的渲染文本」：每张表一行 `## sheet名`，其后是各分组的表头行与数据行（用现有 `Results` 的表格文本渲染）。**不新造一种文本格式**，否则 `text` 与 `tables` 的对齐关系要重新解释一遍。

两条限制要在 JavaDoc 里写明：

- **`.xlsx` 的事件流路线需要 `File`**（`OPCPackage.open(File)`）。`InputStream` 入口只能先落临时文件，或退回 usermodel；这一条决定 `parse(InputStream)` 的实现策略，也是 POI 相对 Aspose 的真实短板（Aspose 的 `Workbook` 有多种构造入口，本次实测只走了文件路径 `new Workbook(String)`，`InputStream` 入口未测）。
- **Excel 没有「一张 sheet = 一张表」的保证**。空白行、重复表头、同一页两张表都是常态，这正是第 2.2 节那些退化结果出现的原因，也是必须复用 `Tables.group` 的理由（行为与 docx/pptx 对齐，跨格式一致性测试才有意义）。

### 8.1 分四步落地

**第 1 步：usermodel 版打通全链路（一次提交，可独立合并）**

`XSSFWorkbook` / `WorkbookFactory.create()` 读 7 张表，产出 `{headers, rows, sheet}`。这一步只需要三件事，全部有实测依据：

```java
// ① 合并区回填 + 分组行：POI 侧与 Aspose 侧算法一致，实测五张表逐格对齐
CellRangeAddress region 起点铺满覆盖格 → values[][]
separator[r] = columnCount > 1 && (columnCount - 被覆盖格数) == 1;
// ② 表头弱线索：只有冻结首行可用
heading[0] = pane != null && pane.isFreezePane()
        && pane.getHorizontalSplitTopRow() == 1 && pane.getVerticalSplitLeftColumn() == 0;
// ③ 取值分派（不能用 getStringCellValue 一杆子到底）
switch (cell.getCellType()) {
    case FORMULA: formatter.formatCellValue(cell, evaluator);  // 只传 cell 会拿到公式文本
    ...
}
```

`heading`/`separator` 这两条的语义与 `PptParser.parseTable` 完全同构，可以直接对照实现；差异只在 Excel **不提供** Word 的 `w:tblHeader`，表头标记只能靠冻结窗格这一条弱线索，所以跨多组表头时仍然要靠 `Options.headers(...)` 注册。

测试：把探针报告第 2 节（POI usermodel 产出网格后喂 `Tables.group`）那五张表的结果原样转成断言（`季度销售明细` 1 组 / 4 行 + 同名列后缀；`客户回款台账` 2 组 / 5 行、命中下标 `[1]`；`分片区年度目标` 3 组 / 5 行、`DROP` 后 2 组 / 3 行；`两张表同页` 注册表头 3 组 / 4 行；空表 0 组；隐藏表按 `Options` 决定收不收）。

样例文件要**按 `sales-report.docx` 的同源数据重造**再放进 `src/test/resources/excel/`：探针那份 `structured.xlsx` 的明细行数值与现有三份样例不同（它是为覆盖结构写的，不是为跨格式对齐写的），直接搬过来会让 `CrossFormatTest` 的对齐断言变成假的。正式样例造好后，`CrossFormatTest` 的分发链（现在是 if 链：WORD / PDF / 其余走 PPT）加一条 Excel 分支，核心表的列名与 4 行数据必须与另外三种格式逐格相同。

**第 2 步：加事件流路线，把大表接管过来**

`XSSFReader` + `XSSFSheetXMLHandler` + `namespaceAware=true`（第 3 节的坑，测试里必须断言行数与格数，不能只看没抛异常）；在同一次遍历上挂一个 `XMLFilter` 收 `<mergeCells>` 的 `ref`，回填后喂同一个 `Tables.group`。触发条件建议按规模切换而不是按用户配置，但**判据不能靠"先 usermodel 打开看一眼"**——4.3 节 A 组已经证明 10 万行的表在 256 MB 下 DOM 打不开。可选的判据有三个：sheet XML 开头的 `<dimension ref="A1:G100000"/>`（事件流里顺手就能拿到，本次探针没测，实现时要验）、`File.length()`、或者干脆全程只走事件流。`sheet.getLastRowNum()`（实测能给出 4 / 5 / 6 / −1）只在 usermodel 分支里有意义，事件流里没有 `Sheet` 对象可问。阈值本身要跑一次实测才写死，探针的 100 000 行 / 2.0 s / 129 MB 是这条线的成本上界。

这一步同时给 `.xls` 留位置：HSSF 侧的事件流（`org.apache.poi.hssf.eventusermodel` 那一套，`HSSFEventFactory` / `HSSFRequest`）本次**没有测**，第 1 步先用 usermodel 覆盖 `.xls`，第 2 步只对 `.xlsx` 换事件流。

**第 3 步：格式与取值的补齐**

- 中文日期里漏出的引号（`2025"年"12"月"31"日"`）：在 `Results.cleanContent` 链上加一条规则，或改用「`getDateCellValue()` 命中日期格式就按 `yyyy-MM-dd` 输出」的策略。**两种做法都要先有样例再写代码**，本次只有 `格式与取值` 这一张表 4 行数据，不足以定规则。
- 数字要原始值还是显示值：`Tables.group` 收到的必须是字符串，所以只能二选一。建议**默认显示值**（与 Word/PPT 里读到的"人看到的文字"一致），并预留一个 Options 开关取原始值——原始值路线在探针里已跑通（`46022.0`、`0.15`、`7`）。
- 公式：只读缓存值，**不重算**（重算要 `FormulaEvaluator`，对全表有性能与副作用风险；实测缓存值与显式重算同为 596.8，但那是我们自己造的表）。

**第 4 步：能力边界与文档**

`README` 的第 9 节（能力边界）与第 15 节（已知限制）各加一段：不支持 `.xlsb`、不读图表、不读数据透视表、条件格式不感知、隐藏 sheet/行列的默认策略、以及"Aspose.Cells 后端未进主干"这件事本身。`package-info.java` 补 Excel 的返回结构。

### 8.2 关于 Aspose.Cells 后端

如果商务侧确认 Cells 可用（或愿意单独买），第 1 步的实现可以整体换成 Aspose 版本，工作量更小：`getMergeCells()` 直接给 `CellArea`、`getStringValue()` 直接给干净的显示值、`getFreezedPanes()` 给冻结信息，第 2.2 节的分组结果与 POI 版逐字一致（这点是本次实测里最省心的发现——算法不用重调）。代价是 8.35 MB jar、许可另购、以及 `new Workbook(String)` 走的是文件路径（`InputStream` 支持未测）。

架构上建议：**不要为两种后端做抽象层。** 项目现在是"一个格式一个解析器类、静态方法、无接口"，加一层 `ExcelBackend` 抽象只会让 `Tables` 那条共用链更难读。真要留后路，把网格产出这一步写成包内私有方法，将来换实现时改一个方法体。

## 9. 本次没有测到的（别当成结论）

明确列出，供下一轮补：

1. **`.xls`（BIFF8）完全没测。** 探针只造 `.xlsx`。HSSF 的 usermodel 与事件流行为、`getMergedRegionAt` 的语义差异都要单独验。
2. **`.xlsm`（带宏）、`.xlsb`（二进制）没测。** `OPCPackage` 能开，但 sheet 部件名与关系类型不同，`XSSFReader.getSheetsData()` 的迭代顺序是否仍与 sheet 顺序一致，未验证。
3. **Aspose.Cells 在 `-Xmx256m` 档拿不到完整耗时**：探针在 120 秒处主动中断（走到第 5 620 行），只能说"这一档它不可用"，不能换算成"它需要多少堆"。512 MB / 1 GB 档的表现也没测——如果将来真要接 Cells 后端，这一条必须先补。另外 4.2 表里它那一行没有独立测（同 JVM 顺序跑的残留不可信）。
4. **Aspose.Cells 评估模式的官方限制条款没查文档**，"实测未见水印"≠"没有功能限制"，也没有与有 lic 的环境做过对照。
5. **EasyExcel/FastExcel 的可见性入口没穷尽。** 探针只反射查了 `ReadSheetHolder` 上"无参且名字含 merge"的方法（结果：无）。隐藏 sheet / 隐藏行列的感知方式、以及它们是否暴露 sheet 索引与顺序，没有逐一查。**补充（2026-10-22 复核）**：合并区这条其实有第二个入口，见第 12 条。
6. **`ExcelParser` 的图片提取没测。** `Images` 那条共用链在 POI 侧对应 `XSSFDrawing`/`getAllPictures`，本次探针完全没碰图片与嵌入对象。
7. **表头识别的弱线索只验了冻结窗格。** 「前两行是表头」「A 列是序号」这类经验规则一概没测，也没测 Excel 里常见的合并单元格伪表头。
8. **真实业务文件没测。** 探针用的是代码生成的表；现实中带筛选（`autoFilter`）、带隐藏行、带外部链接、带共享公式（`<f t="shared" si="..."/>`——POI 与 Aspose 各自怎么把它呈现成单元格文本，本次未测，探针里的 `SUM(B3:B5)` 是普通公式）的文件都要再过一轮。
9. **`DataFormatter` 中文日期引号问题的两种修法都没实现**，只确认了问题存在（第 2.3 节）以及 EasyExcel/Aspose 不受影响。
10. **第 5 节的依赖降级兼容性只在 JDK 8 + Spring Boot 2.7.18 父 POM 这一组配置下验过**，JDK 21 侧本轮未跑探针（历史基线数字另见 README 第 11 节）。
11. **手写 OOXML 路线（第 7 节）没有实现，也就没有耗时数字。** 那一节的判断建立在两个实测点之上——"POI 事件流就是这条路的实测成本"和"formatter 占带 formatter 总耗时的 38%~40%"，不是基准测试结论。
12. **EasyExcel / FastExcel 的 `extra()` 合并区通道没实测，本报告原先的否决理由靠不住，必须复验。** 复核文档与源码后确认：EasyExcel 读侧有 `.extraRead(CellExtraTypeEnum.MERGE)`（默认不读），命中后经 `ReadListener.extra(CellExtra, AnalysisContext)` 给出 `firstRowIndex/firstColumnIndex/lastRowIndex/lastColumnIndex`，官方原文还写明"由于是流式读取，没法在读取到单元格数据的时候直接读取到额外信息，所以只能最后通知哪些单元格有哪些额外信息"（[快速读数据](https://easyexcel.opensource.alibaba.com/docs/current/quickstart/read)）；FastExcel 1.3.0 的对应实现是 `analysis/v07/handlers/MergeCellTagHandler` 与 `v03/handlers/MergeCellsRecordHandler`，与 EasyExcel 同源（[read/extra.md](https://github.com/apache/fesod/blob/1.3.0/website/docs/read/extra.md)）。本次探针只测了默认无模型读法，所以第 1 节矩阵、第 2.2 节结论、第 8.1 节"自己挂 `XMLFilter` 收 `<mergeCells>`"那一步，都只在"不开 extra"的前提下成立。**复验方法**（约 30 分钟，探针跑完即删）：同一份 `structured.xlsx` 开 `extraRead(MERGE)`，把 `extra()` 坐标回填进 `String[][]` 后喂 `Tables.group`，与第 2.2 节 POI usermodel 的结果逐格比；同时在 `-Xmx256m` 独立 JVM 下量峰值内存——extra 是"读完行才给坐标"，意味着网格必须整张留在内存里，这一步会不会把它挤出可用区间，才是它能否进主干的真正判据。若逐格一致且内存过关，第 8.1 节第 2 步的手写解析应改为复用这条通道。

## 10. 探针怎么复现

探针类 `ExcelProbeTest` 与临时依赖是一次性的，跑完已从仓库删除（`pom.xml` 恢复、`target/excel-probe/` 不入库）。要重跑，按本文第 0 节的六条路线重新加 test 作用域依赖即可：`poi-ooxml 5.5.1`、`easyexcel 4.0.3`、`fastexcel 1.3.0`、`aspose-cells ${aspose.version}`（Aspose 仓库的 plain jar，无 classifier）。

采样约定（沿用 README 第 11 节的做法）：每条耗时独立进程跑两轮、只报区间不报平均；内存上限分组每条一个独立 JVM。