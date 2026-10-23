# 设计说明（DESIGN）

本文说明 **为什么这样实现、各部分怎么协作、哪些行为是契约、哪些是已知取舍**。
每个方法的参数、返回值与边界仍以源码 javadoc 为准；怎么调用看 [README.md](README.md)；
对外可见的变化看 [CHANGELOG.md](CHANGELOG.md)。三份样例的实测结果可以直接跑
`mvn -B test -Dtest=PrintResultTest` 看（18 份报告，见第 11 节）。

---

## 1. 目标与边界

**目标**：把 Word / PDF / PPT 读成同一种结构 —— `Map` 里的 `text`、`tables`、`images`，
换格式只换类名；表格要能按「表头 → 行」直接入库，列名与行键永远对齐。

**明确不做**：

- 不做 OCR。扫描件、图片型 PDF 没有文字层，读不出来是预期行为。
- 不做版式还原、渲染、截图比对；`source` 只记页码／幻灯片序号，不记坐标。
- 不执行公式。Word 的 `SUM(ABOVE)` 只读文档里缓存的域结果，不重算（见 §7）。
- 不做文档编辑、不写回原文件；`writeImages` 只导出图片。
- 不引入 Spring、启动类、`main` 或命令行解析，保持纯静态工具类。

底层依赖 Aspose.Words / Aspose.PDF / Aspose.Slides（商业库，不随本仓库分发）。
许可证文件放在 classpath 根目录或进程工作目录；各产品每个进程只注册一次，且必须发生在
加载文档之前 —— 晚注册对已经构造出来的 `Document` 无效。

---

## 2. 模块与职责

| 类 | 职责 | 不负责 |
| --- | --- | --- |
| `WordParser` | doc/docx 的入口；遍历 Body 节点、还原网格、取页码 | 表头分组规则、图片编码、JSON |
| `PdfParser` | pdf 的入口；两遍表格识别、跨页并表、按坐标排序内容项 | 同上 |
| `PptParser` | ppt/pptx 的入口；按形状顺序读表格／文本／图片 | 同上 |
| `Tables` | **唯一一份**表头判定、分组、多级表头合并、补列名、写 `name`／`source` | 不碰 Aspose，不知道页码怎么来 |
| `Images` | 图片元信息、四种输出模式、`contentId`、重名文件、安全文件名 | 不解析文档 |
| `Results` | 文本清洗、域处理、JSON 序列化、`writeImages`、许可证查找 | 不解析文档 |
| `package-info` | 包级使用说明与返回结构示例 | — |

依赖方向是单行道：三个解析器 → `Tables` / `Images` / `Results`，`Tables` 只依赖
`WordParser.Options` 与 `Results.normalize`；`Images` 只依赖 `Options`。解析器之间互不调用。
`toJson`、`writeImages` 在三个解析器上各有一个同签名入口；`tablesByName`、`rowsByName`
只放在 `WordParser`，因为三种格式的返回结构相同，一份取值入口就够。

---

## 3. 公共 API 与调用链

每个解析器都有四个 `parse` 重载：

```java
parse(String path, String... headerKeywords)
parse(InputStream in, String... headerKeywords)
parse(String path, Options options)
parse(InputStream in, Options options)
```

- 路径与流两种入参在**同一文件、同一选项**下产出逐字符相同的 JSON（`PrintResultTest.printPublicApiOverloads`
  在三份样例上实测全为 `true`）。
- 流由**调用方关闭**；解析器只读，不接管流的生命周期。
- 关键字入参不是选项的别名：它关掉「首行兜底」。pdf 上用
  `片区/区域/销售额/回款率/备注` 这组关键字时读到 6 组 47 行，静默模式是 6 组 45 行 ——
  多出的两行是没被关键字命中的表把表头当成了数据。需要稳定结果时，优先注册完整表头。

一次路径解析的调用链：

```
open(path)                       ← 先注册许可证，再 new Document
  → parse(document, options)
      → read / walk              ← 遍历模型，段落进 text，表格成网格，图片进 Images
      → Tables.group(...)        ← 表头判定、分组、合并、补列名、写 name
      → Tables.markSource(...)   ← 各格式补 page / lastPage（source(true) 时）
      → Tables.warnUnmatchedHeaders(...)
      → LinkedHashMap{text, tables, images[, unmatchedHeaders]}
      → close(document)（PDF 必须关，Windows 上否则删不掉原文件）
```

`Options` 在解析过程中只读，可以配好一次反复用于多次解析；图片编号每次都从 1 开始。

---

## 4. 结果模型与契约

```json
{
  "text": "……",
  "tables": [
    {
      "headers": ["片区", "区域", "销售额", "回款率", "销售额-5", "回款率-6", "备注"],
      "rows": [
        {"片区": "华东大区", "区域": "上海", "销售额": "1280", "回款率": "92%",
         "销售额-5": "1465", "回款率-6": "95%", "备注": "超额完成"}
      ],
      "name": "季度明细",
      "source": {"page": 1}
    }
  ],
  "images": [
    {"fileName": "image-1.png", "type": "PNG", "width": 1024, "height": 768,
     "size": 670939, "contentId": "487e915db2c53907e8c1551aeaffea235423210c"}
  ]
}
```

**硬契约**（`CrossFormatTest` / `TestSupport.assertHeadersMatchRowKeys` 守着）：

1. `headers` 与每一行的键**逐个相等**，顺序也一致（行是 `LinkedHashMap`）。
2. 空表头补 `列N`，同名列补 `名-列号`；`headers` 与行键共用同一个 `Tables.columnName`，
   所以补名规则不可能只改一边。
3. 长短不一的行补齐到整表最大列数；单元格值不为 `null`。
4. 表格里的文字不在 `text` 里重复；图片顺序 = 文档流顺序。
5. 每组的键顺序固定 `headers` → `rows` → `name`（只在注册且命中时）→ `source`
   （只在 `source(true)` 时）。顶层默认 `text` → `tables` → `images`；
   打开 `unmatchedHeaders(true)` 才多第四个键，且排在最后。
6. `name` 是调用方在注册时起的名字，不是主键：一个名字可以命中多组
   （同一张物理表被两段表头切开时），取数用 `rowsByName` 一次拼平。
7. `source` 默认开：`{page: 起始页}`，跨页的表多一个 `lastPage`。`source(false)` 时
   三种格式一起不写这个键，docx 侧连 `LayoutCollector` 都不建。

**为什么用 `Map` + `LinkedHashMap`**：不绑定任何 JSON 库、字段顺序稳定、Jackson 可直接序列化；
代价是没有编译期类型检查，所以用第 11 节的契约测试兜底。

### 4.1 重名列为什么是 `销售额-5`

样例源表头第一行确实是 `片区, 区域, 销售额, 回款率, 销售额, 回款率, 备注`，两个名字各出现两次。
JSON 对象的键不能重复，必须消歧。后缀取**物理表的绝对列号（1 起）**：

- `销售额-5` 表示「第 5 列」，回原文核对时按列号去数即可；
- 若改成「第几个重名」（`销售额-2`），与 `列N` 的语义不一致，也失去了位置信息；
- 三份样例上规则一致（docx/pdf/pptx 都是 `销售额-5`、`回款率-6`），
  pptx 台账上层还有 `回款情况-3`。

这条规则由 `PrintResultTest.printHeaderAndFieldRules` 用真实样例 + 现造 docx 双重复现。

---

## 5. 表头识别与分组

### 5.1 输入

`Tables.group(String[][] rows, boolean[] separator, boolean[] heading, Options options,
Set<Integer> matched, boolean firstRowStyled)`：

- `rows`：已把合并单元格**拆开回填**后的网格，每行长度一致；
- `separator[r]`：第 r 行是不是「跨整行合并」的分组小标题；
- `heading[r]`：格式自带的标题行标记（格式没有这种标记时传 `null`）；
- `firstRowStyled`：表样式是否声明首行为标题（Word 的 `tblLook firstRow`）；
- `matched`：出参，记录实际命中的注册组下标，供未匹配提示使用。

### 5.2 六种表头依据（任一命中即可）

按优先级：

1. **注册的完整表头集合**（`Options.headers(...)` / `headerGroups(...)`）；
2. **关键字**（`Options.keywords(...)`）：一行的所有非空单元格都含关键字之一；
3. **格式自带的标题行标记**：docx 的 `w:tblHeader` / `tblLook firstRow`、pptx 的 `firstRow`；
4. **紧跟跨整行合并行之后的那一行**；
5. **表样式声明的首行**；
6. **静默模式的首行兜底**：只有既没注册表头、也没给关键字时才允许。

注册表头命中后，输出列名直接用**注册名**，不用文档里读出来的那一行 —— 这样 PDF 单元格
因换行多出的空格不会让三种格式的键不一致。

### 5.3 匹配口径

`Options.matchHeader` 的规则：

- 忽略大小写与所有空白（`Results.normalize`）；
- 忽略行尾空列；列数必须相同；
- 不一致的列数 ≤ `min(headerTolerance, 列数 - 1)`，默认 `headerTolerance = 1`；
- 至少要求一列完全相等，避免容忍度过宽把数据行当成表头。

`headerTolerance(0)` 要求逐列相等。实测把核心表的「备注」注册成「说明」时，默认 1 命中、
收到 0 落空并提示未匹配 —— 收紧不是无代价的。

### 5.4 多级表头

连续的多个表头行先合并：低层缺名时沿用上层同列的名字。合并完还会拿**合并结果**再与注册集
比一次，所以照「合并之后那一行」注册（`客户名称/合同金额/已回款/回款率/备注`）在
docx／pdf／pptx 上都命中，容忍度 0 也命中。反过来只注册上层那一行有格式差异：

- docx：照样 13 行（两行表头都带 `w:tblHeader`，会连着合并）；
- pptx：下层那行不再算表头，漏进数据，7 行变 8 行；
- pdf：网格里只有合并后那一行，命不中（容忍度调到 2 才命中）。

跨格式共用一份注册时，把上下两种写法都注册、起**同一个名字**最稳。同名之下只要有一份命中，
另一份命不中就不再提示未匹配（否则这条推荐写法会天天刷噪声）。

### 5.5 跨整行合并行（分组小标题）

判据：整行只剩一个**不是横向合并延续格**的格子，**且这一行不是全空**。

- 全空合并行只是排版间隔；把它当成分组行，紧跟它的数据行会被误当成表头，
  一张表会被切成两组、下一行数据丢失（`WordEdgeCasesTest` / `PptParserTest` 各有一条先失败后通过的用例）。
- 纵向合并的延续格仍算这一行自己的格子；只排除横向延续格，与三种格式的 `toGrid` / `Grid.isSeparator` 同一条。
- 分组行的去向由 `FullRow` 决定：`KEEP` 留在当前组（默认）、`FIRST` 归到同一张物理表的第一组、
  `DROP` 丢弃。无论哪种，它都标志着「上一组的表头区结束」，紧跟其后的行不会再被并进上一层表头。

### 5.6 分组、补名与 `name`

- 遇到新的表头行就开新组；数据行挂在上面最近的一组表头下。
- 空表头补 `列N`；与前面列重名补 `名-（列号）`。`rowMap` 用同一个函数，保证键与 `headers` 对齐。
- `name` 写在 `Tables.group` **内部**、命中注册表头的那一刻；没起名的组、靠关键字／格式标记／
  首行兜底判出来的组都没有这个键。`source` 由解析器在分组**外部**补，因为只有各格式知道页码。
- 连续两行表头各自命中不同注册时，以**第一个**名为准，不覆盖；代价是第二个名字在这份文档上落不了地。
- `tablesByName` 返回结果里的原对象（不是副本），`rowsByName` 把同名各组的行按组顺序拼平。
  没命中、传 `null`、传空白都返回空列表，不抛异常。

### 5.7 为什么默认不给所有组自动起名

自动名字只能来自表头内容或序号，前者会随文档改字而变，后者会随分组顺序而变；两者都会让下游
「按名字取数」的代码悄悄错位。`name` 的定位是**调用方声明自己要哪张表**，所以保持显式：
注册了就写，没注册就没有。工单类型那张表在 `TestSupport.named()` 里只注册、没起名，
结果里没有 `name` 是这条契约的正常表现；要按名字取就补一个 `.name("...")`。

---

## 6. 三种格式的适配

### 6.1 Word（doc/docx）

- 逐节遍历 Body 的节点树：段落进 `text`，表格进 `tables`，图片进 `images`；
  块级内容控件（`w:sdt`）与文本框里的内容会往下钻，不会因为多一层包装整块丢掉。
- 表格用 `Grid` 展开 `gridSpan` / `vMerge`：合并覆盖的位置回填起始格的值；
  单元格内嵌套表不摊进母表格子文本，而是作为额外一组排在母表分组之后。
- 页码用 `LayoutCollector` 计算；`source(false)` 时不创建它 —— 这是 docx 侧最贵的一步。
- 页眉页脚、脚注尾注、批注不读（版式附属内容，评估水印也注入在那里）。

### 6.2 PDF

PDF 的内容流顺序不等于人读顺序，所以分两阶段：

1. **每页收集内容项**，都带上页码与纵向位置：
   - 有线表：`TableAbsorber` 吸出来的表格块先用并查集按「矩形相接（2pt 容差）」聚成一张物理表；
   - 无边框表：再开一遍流式引擎，只收「每行格数相同、≥2 行 ≥2 列、且不与任何有线表矩形相交」的表；
   - 正文：段落中心点落在表格矩形内的跳过，避免表格文字在 `text` 里重复；
   - 图片：带位置收集。
2. **按「页码升序、同页自上而下」排序**，再交给 `Tables.group`；跨页长表在排序之后、分组之前并回一张。

跨页续接的五个条件：相邻页、列数相同且左边对齐（2pt）、上一截是上一页最后一段内容且下一截是
下一页第一段内容、下一截里找不到任何能当表头的行。静默模式没有表头依据，再要求
「上一截贴页底、下一截贴页首」（各占页高 20% 以内）作为几何证据（第五条之后加的一条）。
三页以上的长表靠「并完一截就把尾巴换成后一截」继续接 —— 尾巴的页码、离页底距离和
「是不是本页最后一块」都要换成后一截的，否则中间页后面跟了图片时下一页的独立表会被错并。

已知代价：

- 扫描件／图片型 PDF 没有文字层，本模块不做 OCR；
- 用制表位排齐的几行散文本满足无边框表判据，会被当成表从 `text` 挪进 `tables`
  （`PdfBorderlessTableTest.tabAlignedParagraphsCurrentlyCountAsATable` 把这个误报钉住）；
- 静默模式下，上一页最后一行很高（下边界离页底超过页高的 1/5）时判不出续接，会拆成两组；
- 中间那页的表格片段后面还有别的内容时，续接链条会断 —— 这是「不把两张无关表并错」换来的安全性。

### 6.3 PPT（ppt/pptx）

- 幻灯片里 `getShapes()` 的顺序就是绘制顺序，天然是人读顺序，**不需要按坐标重排**；
  表格与图片的编号因此直接按文档流走。
- 四类形状：`ITable`、`IPictureFrame`、`IGroupShape`（递归展开，深度上限 32）、`IAutoShape`（文本）。
  最宽的 `IAutoShape` 放在最后兜底，避免 Aspose 以后调整继承关系时把表格文字当普通文本读走。
- 表格按 `rowSpan` / `colSpan` 把起始格的值铺满覆盖区域；跨整行合并行的判据只看横向延续格。
- 一张表只在一页幻灯片上，没有「跨页表」；`source` 的起止页写同一个序号。
- 不读：备注页／母版／版式页、`Chart`、`OLE` 对象、`SmartArt` 内文字、形状填充底图。

---

## 7. 文本与域

三种格式的文本都过 `Results.cleanContent`：

- Word 用 `\u0013 域代码 \u0014 域结果 \u0015` 存域；丢掉域代码、只留域结果，
  否则单元格里会多出 `PAGE`，还可能被当成表头。
- **不调用 `updateFields`**：域结果没算过就是空的。`SUM(ABOVE)` 读缓存值，不重算 ——
  这是文档解析器而不是表格计算引擎的取舍。`PrintResultTest.printHeaderAndFieldRules` 用
  「1 + 2 但缓存写 300.00」的现造文档证明读出来就是 300.00，缓存为空则给空串。
- 软换行、分页符换成空格而不是丢弃（丢了会把 Shift+Enter 分的两行粘成一句）；
  单元格结束符等不可见控制字符丢弃；连续空格压成一个。
- Aspose 评估模式的水印文字统一过滤；pptx 的截断提示单独剥离。被截断的内容回不来，
  所以评估模式输出只能当样例看。

---

## 8. 图片

- 元信息在任何模式下都给：`fileName`、`type`、`width`、`height`、`size`、`contentId`（SHA-1）。
- `ImageMode`：
  - `META`（默认）：只要元信息；
  - `BYTES`：原始 `byte[]`；
  - `BASE64`：`data:image/...;base64,...`，前端可直接塞 `<img src>`；
  - `FILE`：边解析边落盘，结果里只留路径，必须配 `imageDir`。
- 内嵌模式下内容相同的图片用 `contentId` 标 `duplicateOf`，不重复内嵌；
  外链图片只记路径，不联网、不下载。
- `writeImages(result, dir)` 只处理带 `base64` 的结果；用 `FILE` 解析时不需要再调它。
  写文件前只取文件名部分（防路径穿越），目录内重名自动加序号。
- 顺序 = 文档流顺序；`BASE64` 会把 JSON 撑得很大（docx 样例 4740 → 899370 字符），
  需要正文就选 `FILE` / `BYTES`。

---

## 9. 错误、日志与资源

- **不包装底层异常**：文件不存在时三种格式抛的类型不同（docx `java.io.FileNotFoundException`、
  pdf Aspose 内部异常、pptx `FileNotFoundException`），消息都是 `Can't find file: <路径>.`。
  调用方要统一捕获，只能按格式分别处理；`PrintResultTest.printErrorPaths` 原样打印这六条路径。
- `FILE` 模式没配 `imageDir`：三个解析器统一抛 `IllegalStateException`。
- 可恢复的配置问题走标准错误，不中断解析：未匹配表头、`name` 没有归属、`META + imageDir`、
  `writeImages` 没有 base64 内容、组合形状嵌套超限。
- PDF 的 `Document` 必须在 `finally` 里关闭（持有文件句柄）；流由调用方关闭。
- 许可证注册用 `synchronized` 去重，一个产品每进程只注册一次；失败会提示一行，不静默吞掉。

---

## 10. 线程安全、性能与版本

- 没有跨文档的可变全局状态；每次解析创建自己的模型、`Images` 与结果 `Map`。
  `Options` 在解析中只读，可复用。许可证注册是进程级全局状态，由静态同步方法保护。
  **未做**并发压力测试，多线程同时解析不在当前验证范围内。
- 性能上唯一能精确归因的开关是 `source(false)`：docx 不建 `LayoutCollector`，
  同一 JVM 开／关交替实测 74–254ms 对 11–35ms（两轮里「关」的最大值都低于同轮「开」的最小值）。
  PDF／PPT 的页码本来就免费，开关只决定写不写键。
- 其他耗时只看数量级，别当性能回归判据（机器负载会把差值吃掉）。
- `pom.xml` 已是 `1.2.1`（未发布）：1.2.0 收进的新 API 与三个修复已在 Central 上，1.2.1 只把 Jackson 升到
  `2.21.8` 修 CVE-2026-110879，API、返回结构与解析行为不变。
  返回结构、`Options` 方法、公开入口、命中口径的变化都算契约变更，必须进 CHANGELOG 并升版本。

---

## 11. 测试策略

分三层：

1. **单元测试**（不经过 Aspose，直接喂网格／文本／字节）：
   `TablesTest` 34、`ResultsTest` 28、`ImagesTest` 25、`OptionsTest` 14、`PptParserTextTest` 2。
2. **端到端断言**（三份样例 + 现造最小文档）：
   `CrossFormatTest` 15、`WordParserTest` 32、`PptParserTest` 19、`PdfParserTest` 17、
   `WordEdgeCasesTest` 14、`TableByNameTest` 11、`PdfBorderlessTableTest` 5、
   `PdfMergedCellTableTest` 5、`PdfCrossPageTableTest` 4。
3. **打印给人看的报告**：`PrintResultTest` 17 条用例 / 18 份报告，只保留一条契约断言
   （`assertHeadersMatchRowKeys`），其余把真实结果原样打出来，供发布前人工核对：
   - 三份样例的完整结果（7 份）；
   - 公共接口逐个跑（4 份）：四个 `parse` 重载、`Options` 默认值与 setter、
     `toJson`／`writeImages`、`tablesByName`／`rowsByName`；
   - 表头规则与选项矩阵（5 份，其中 `printAllModeJson` 把三份样例 × 7 种选项的完整 JSON 打出来）；
   - 图片与错误路径（2 份）。

   运行 `mvn -B test -Dtest=PrintResultTest` 会同时在 `target/print/` 留 18 份 UTF-8 副本
   （实测合计 191974 个字符，最短 1351、最长 106026），`mvn clean` 一并清掉。

当前基线：**242 条测试 / 15 个测试类 / 0 失败 / 1 跳过**（跳过的是 `-Ddoc` 入口），
JDK 8 与 JDK 21 都按同一份源码编译（字节码 `major version: 52`）。
新增格式的做法：写一个 `XxxParser`，把表格读成 `String[][]` + `boolean[] separator` +
可选 `boolean[] heading`，交给 `Tables.group(...)`；图片交给 `new Images(options).add(...)`；
`Options`、`toJson`、`writeImages` 复用，不要在解析器里各抄一份。

---

## 12. 已知限制与取舍（汇总）

| 限制 | 原因 / 取舍 |
| --- | --- |
| 扫描件读不出文字 | 不做 OCR，只读文字层 |
| 无边框 PDF 表可能误报 | 流式引擎的判据；制表位排齐的散文本会被当成表（已用测试钉住） |
| 静默模式跨页判断可能拆表 | 只靠「贴页底／贴页首」的几何证据，宁拆不并错 |
| 多级表头只注册上层时有格式差异 | 各格式给出的网格行不同；推荐上下两种写法同名注册 |
| 同名列必须加后缀 | JSON 键不能重复；用绝对列号保留位置 |
| `name` 不是主键 | 一个名字可命中多组，取数用 `rowsByName` |
| 域结果不重算 | 不执行 `updateFields`，只读缓存 |
| 三份样例的合计值不同 | 源文件存的内容不同：docx 有 SUM 缓存，pdf 导出丢缓存，ppt 是字面量 |
| 评估模式内容被截断 | Aspose 限制；过滤痕迹但无法还原被截断内容 |
| PPT 图表／SmartArt／OLE 不读 | 不在当前形状分支内，已知范围外 |

---

## 13. 改代码前请先确认

1. 结果契约（`headers` = 行键、键顺序、空/重名列补名）不能破，`CrossFormatTest` 会红。
2. 表头判定、图片、文本清洗分别只改 `Tables`、`Images`、`Results`，不要在解析器里复制规则。
3. javadoc 里的数字必须来自实测；改行为或样例后重跑对应报告，把数字重新写进去。
4. 对外可见的变化同时补 [CHANGELOG.md](CHANGELOG.md) 与 [CHANGELOG.en.md](CHANGELOG.en.md)。
5. 发布前跑 `mvn -B clean test`、`PrintResultTest` 全量报告与 doclint；不要提交 `target/`、
   `*.lic`、真实客户文档。
