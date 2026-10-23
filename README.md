# document-parser

[![License](https://img.shields.io/badge/License-Apache%202.0-blue.svg)](LICENSE)
[![Java](https://img.shields.io/badge/JDK-8%2B-orange.svg)](#3-技术栈与环境要求)
[![Maven Central](https://img.shields.io/badge/Maven%20Central-1.2.0-green.svg)](#17-构建)
[![Tests](https://img.shields.io/badge/tests-241%20passed-brightgreen.svg)](#13-测试与质量)

[中文](README.md) ｜ **English** → [README.en.md](README.en.md)

> 文档解析模块：把 **Word / PDF / PPT** 统一解析成 `text`、`tables`、`images` 三个集合，
> 三种格式返回同一套结构，换格式只换类名。

| 项目信息 | 值 |
| --- | --- |
| Maven 坐标 | `cn.net.tianxi:document-parser:1.2.1`（Central 当前版本，2026-10-24 发布；相对 1.2.0 只升级 Jackson 依赖） |
| Java 包 | `cn.net.tianxi.document` |
| 语言级别 | Java 8（Aspose 用 `jdk16` classifier 的构件） |
| 底层技术 | Aspose.Words / Aspose.PDF / Aspose.Slides 23.12，API 与行为以 [docs.aspose.com](https://docs.aspose.com) 为准 |
| 产物 | 普通库 jar（1.2.1 发版产物 **60777 字节、30 个条目、17 个 class**，JDK 8 下 `mvn -B clean package`；JDK 21 打同一份源码为 61190 字节，含 `META-INF/LICENSE` 与 `META-INF/NOTICE`；1.2.0 是 60236 字节，1.1.1 是 56275 字节；发布记录见 [CHANGELOG.md](CHANGELOG.md) 与 §17），无 Spring 容器、无命令行入口 |
| 测试 | JUnit 5，242 条测试、15 个测试类：241 条通过，1 条是 `-Ddoc` 入口、平时跳过 |
| 更新日志 | [CHANGELOG.md](CHANGELOG.md)（英文 [CHANGELOG.en.md](CHANGELOG.en.md)）—— 每个版本对外可见的变化与实测数字 |
| 开源许可 | [Apache License 2.0](LICENSE)：可免费使用、修改、商用、再分发，**须保留原作者署名**（见 [NOTICE](NOTICE)），欢迎提 PR（见 [CONTRIBUTING.md](CONTRIBUTING.md)） |
| 代码仓库 | GitHub [`hutianxi/document-parser`](https://github.com/hutianxi/document-parser)｜Gitee [`mailtianxi/document-parser`](https://gitee.com/mailtianxi/document-parser) |

更细的说明分三层：整体设计与实现思路（模块职责、数据流、表头判定、各格式适配、边界与取舍）见
[DESIGN.md](DESIGN.md)；包级设计（返回结构、表头判定优先级、许可证与评估模式、安全边界）写在
[`package-info.java`](src/main/java/cn/net/tianxi/document/package-info.java)，
类与方法的细节写在各自的 javadoc 里，IDE 里悬停即可看到。本 README 是这三层的索引 + 实测数据。

---

## 目录

1. [项目简介](#1-项目简介)
2. [特性](#2-特性)
3. [技术栈与环境要求](#3-技术栈与环境要求)
4. [目录结构](#4-目录结构)
5. [快速开始](#5-快速开始)
6. [API 一览](#6-api-一览)
7. [返回结构详解](#7-返回结构详解)
8. [选项详解](#8-选项详解)
9. [能力边界：读到什么、读不到什么](#9-能力边界读到什么读不到什么)
10. [测试样例](#10-测试样例)
11. [实测基线](#11-实测基线)
12. [开源许可、Aspose 许可证与评估模式](#12-开源许可aspose-许可证与评估模式)
13. [测试与质量](#13-测试与质量)
14. [常见问题](#14-常见问题)
15. [已知限制与未覆盖分支](#15-已知限制与未覆盖分支)
16. [维护指南](#16-维护指南)
17. [构建](#17-构建)
18. [附录：修复记录（历史数字）](#18-附录修复记录历史数字)

---

## 1. 项目简介

业务上经常要「把客户上传的合同、报告、汇报 PPT 里的文字和表格抽出来入库」。这三种格式的底层模型完全不同：
Word 有真正的表格对象和合并标记，PDF 只有画出来的线段和散落的文字，PPT 是一页页形状。
如果每种格式各写一套解析、各返回一种结构，上层就要写三份入库逻辑，而且改一处规则要改三个地方。

本模块的做法是：**把「格式差异」关在解析器里，对外只暴露一种结果结构**。

- 三个入口类 `WordParser` / `PdfParser` / `PptParser`，方法签名一一对应，返回 `Map<String, Object>`，
  键固定是 `text` / `tables` / `images`；
- 表格统一还原成「表头 + 行对象」：`headers` 逐个等于每一行的键，所以可以直接拿 `headers` 当列清单去建表或映射字段；
- 合并单元格一律拆开回填，每行的键数量因此恒定；
- 表头分组、图片收集、文本清洗各只有一份实现（`Tables` / `Images` / `Results`），三种格式共用。

设计上刻意**不做**的事：不做 OCR（扫描件没有文字层就读不出来）、不做版面重排、不调 `updateFields` 重算域、
不做入库和网络请求。它是一个纯静态方法库，没有启动类、没有命令行、没有全局状态（除了 Aspose 的许可证注册）。

### 1.1 与 OCR 的区别：读文档里的文字层，不猜像素

这是两条完全不同的技术路线，先说清楚再谈优劣：

- **本项目**走的是「解析格式对象模型」：Aspose 直接读 docx 的 `w:tbl`/`w:tc`/`w:sdt` 节点、PDF 的文字对象与线段坐标、
  pptx 的形状树。文件里本来就存着字符数据，解析器只是把它按结构取出来。
- **OCR** 走的是「图像识别」：把页面渲染成图，逐像素判断字形，再猜字、猜行、猜表格框线。文字是**推断出来的**，不是读出来的。

| 维度 | document-parser（读文字层） | OCR（图像识别） |
| --- | --- | --- |
| 输入前提 | 电子原生文档：Word 存的 docx、程序导出的 pdf、PowerPoint 存的 pptx | 任何能变成图片的页面，**包括扫描件、拍照件** |
| 文字从哪来 | 文件内的字符数据，原样读出 | 从像素推断，存在误识率 |
| 数字与金额 | 与原文逐字符一致：`合计` 行读到的就是域缓存值 `2096.30`/`1796.10`，这些值由测试逐格断言（§11） | `0/O`、`1/l`、`8/B`、小数点这类字形易混，金额通常要人工复核 |
| 表格结构 | 拿到**真结构**：合并单元格标记（`gridSpan`/`vMerge`/`rowSpan`）、两行表头、单元格内嵌套子表都能识别并拆开回填 | 靠框线检测还原；无边框表、跨页断表、合并格容易错行错列 |
| 阅读顺序 | 按文档树顺序（段落 → 表格 → 形状），与作者的写作顺序一致 | 按版面坐标推断，多栏、浮动文本框容易串行 |
| 结果稳定性 | 同一份文件重复解析结果**完全相同**，可以写成断言进 CI | 换引擎版本、换分辨率、换二值化参数结果会变，难做回归 |
| 部署形态 | 一个 jar + 三个 Aspose 依赖，进程内完成，不联网（外链图片只记元信息、不下载） | 一般要额外装识别引擎/模型，或把文档上传到第三方云服务 |
| 扫描件 / 图片型 PDF | **读不出来**（没有文字层），这是 §9 写明的边界 | 唯一可行的路线 |

> **诚实边界**：本节不给 OCR 的准确率或耗时数字。仓库里没有跑过任何 OCR 引擎，编不出可复现的实测值；
> 要横向对比，请拿同一批文件在两边各跑一遍再比。上表右列写的是 OCR 这条路线的**固有特征**，不是某个产品的实测结果。

本项目的单文件耗时是实测的（三份样例，评估模式、无 `.lic`；用一次性探针测试类在 JDK 8 下连跑两轮 `mvn test`，取两轮区间 —— 探针按约定跑完就删，现在仓库里没有它，要复现得自己写一个）：

| 样例 | 大小 | 首次调用（含 Aspose 初始化） | 之后单次解析 |
| --- | --- | --- | --- |
| docx | 675208 字节 | 2058–2396 ms | 83–227 ms |
| pdf（3 页） | 184977 字节 | 4830–5363 ms | 671–1268 ms |
| pptx | 679289 字节 | 934–1006 ms | 61–75 ms |

环境：JDK `1.8.0_191`、Maven 3.8.8、Windows 11。换 JDK 21.0.1 再测一轮，量级相同
（首次 3313／7360／2322 ms，之后 126–192／938–1655／86–99 ms）。三点说明：

- **首次调用慢是固定开销**，来自每个 Aspose 产品第一次加载类和初始化许可证，同一进程内只发生一次；
  常驻服务（Spring Bean、线程池）里跑的是「之后」那一列。
- **docx 的 `source` 键要付版面构建**：`LayoutCollector` 的构造本身 0ms，第一次 `getStartPageIndex` 才把整篇版面
  算出来 —— 探针拆开实测（加载 16–24ms、建 collector 0ms、5 张表 17 次取页码 98–131ms），这就是「之后」那一列
  从上一版的 23–30ms 涨到 83–227ms 的原因。不要页码就 `Options.source(false)`：Word 侧连 collector 都不建，
  同一份样例实测回到 **11–35ms**（一次性探针在同一 JVM 里把开／关**交替**各跑 6 次、连跑两轮，避免后跑的那一种
  白捡 JIT 预热：开 74–254ms 对 关 11–35ms，两轮里「关」的最大值都低于同一轮「开」的最小值；探针跑完就删，
  仓库里没有它）。PDF／PPT 的页码是解析过程中本来就有的，开／关实测区间重叠（pdf 681–960ms 对 696–984ms、
  pptx 46–86ms 对 44–71ms），不额外花钱，但开关对三种格式一起生效（关掉后三种格式都不写 `source`）。
- **PDF 那 700–1200 ms** 花在 `TableAbsorber` 的几何聚类、跨页并表和**每页第二遍流式引擎**上，与页数、线段数量
  相关，不是文件体积决定的（184977 字节的 pdf 比 679289 字节的 pptx 慢一个数量级）。第二遍是上一版的 1.3 倍左右：
  同一 JVM 里拆开实测 3 页 —— 只跑线段 478ms、只跑流式 615ms、两遍 922ms。

### 1.2 这样做的好处

- **上层只写一份逻辑**：三种格式返回同一套 `text`/`tables`/`images` 结构，换格式只换类名；
  加一种新格式只需再写一个解析器，入库代码不用动。
- **`headers` 可以直接当列清单**：每组表的 `headers` 逐个等于每一行的键，合并单元格一律拆开回填，
  所以每行键数量恒定 —— 12 次解析（3 份样例 × 4 种选项）契约 **0** 处不一致
  （`CrossFormatTest.contractHoldsForEverySampleAndMode`）。建表、映射 DTO、写 `INSERT` 都不用特殊处理。
- **可以进 CI 当回归**：结果是确定的，所以 README §11 里的每个数字都是测试断言值，改坏了立刻红。
- **入库定位不用数下标**：注册表头时顺手 `Options.name("季度明细")`，之后 `WordParser.rowsByName(result, "季度明细")`
  直接拿到那几张表的数据行；`source.page` 又给出它原来在第几页，核对时翻得到页。
- **无状态、可并发**：三个解析器都是静态方法，没有全局可变状态（除 Aspose 的许可证注册），配好的 `Options` 可复用。
- **不出进程、不联网**：纯本地解析，外链图片只记元信息不下载，写文件前只取文件名部分并处理重名 ——
  合同、财报这类不允许出内网的文档可以直接用。
- **图片按需处理**：四种输出模式（元信息 / base64 / 边解析边落盘 / 原始字节）+ SHA-1 内容指纹 `contentId`，
  同一张图在多份文档里出现只存一次。

### 1.3 应用场景

适合直接上的场景：

| 场景 | 怎么用 |
| --- | --- |
| 合同 / 报告 / 台账入库 | 客户上传的 docx 一次解析成正文 + 结构化表 + 附件图，`headers` 直接映射数据库列 |
| RAG 与知识库切片 | `text` 做语义切片，`tables` 保留结构单独存（表格拍平成文字会丢列关系），`images` 用 `ImageMode.FILE` 落盘再传对象存储 |
| BI / 报表汇聚 | 下属单位交上来的同模板 Word / PDF 报表，注册表头后列名固定，直接批量入仓 |
| 财务与审计留痕 | 金额取的是文件里的原值（域缓存结果），不是识别值，可追溯 |
| 多格式一致性核对 | 同一份报告的 docx / pdf / pptx 三版逐格对比（`CrossFormatTest` 就是这么干的） |
| PPT 汇报内容归档 | 正片文字、表格、图片一次抽出，备注页读不到（§9） |
| 扫描件 + 电子件混合的收件流程 | 用本模块先探一遍：`text` 与 `tables` 都空说明没有文字层，降级走 OCR，其余走本模块。**这条分流判据是建议，仓库里没有扫描件样例，未做实测** |

不适合、需要换方案的场景（都在 §9 有对应说明）：

- **扫描件、拍照件、图片型 PDF**：没有文字层，必须 OCR；
- **图表（Chart）里的数据、SmartArt 文字、嵌入的 Excel/Word（OLE 对象）内容**：Aspose 的形状分发识别得出类型，但本模块不读它们的内容；
- **页眉页脚、脚注尾注、批注**：Word 那一支不读；
- **PPT 备注页、母版、版式页文字**：不读；
- **需要保留版面**（分页位置、字体、坐标、原样转 HTML/PDF）：本模块不做版面重排，只给内容。

## 2. 特性

- **一套结构三种格式**：doc/docx、pdf、ppt/pptx 的返回结构完全一致，`CrossFormatTest` 逐格核对同源样例的核心表在三种格式下相同。
- **表头分组**：一张物理表里有几组表头，`tables` 里就有几个元素，每组只装属于它的数据行，不串组；支持两行表头（上层跨列、下层留空时列名沿用上层）。
- **六种表头判定依据**（优先级从上到下）：注册的完整表头集合 → 关键字 → 格式自带的标题行标记 → 紧跟分组行之后的行 → 表样式声明的首行 → 静默模式下的第一行兜底。
- **合并单元格拆开**：横向（`gridSpan` / `colSpan` / 矩形跨列）、纵向（`vMerge` / `rowSpan` / 矩形跨行）覆盖的位置都回填起始格的值。
- **分组小标题可控**：跨整行合并的行（例如「2025 年度（已结账）」）按 `FullRow` 保留在原组、统一归到首组，或直接丢弃。
- **PDF 跨页长表并回一张**：几何聚类 + 六条续接判据（相邻页、列数相同、左边界对齐、前一截是上页最后一块且后一截是本页第一块、后一截找不出表头行；静默模式再加「一截贴页底、一截贴页首」），跨三页以上也是一路并下去。
- **PDF 无边框表补读**：线段引擎之外再跑一遍流式引擎，只收「每行格数相同、不与有线表重叠」的那几张（实测无边框样例 0 组 → 3 组，全有线的仓库样例一组没多）。
- **每组表带来源位置**：`source.page`（docx 页码 / pdf 页码 / pptx 幻灯片序号，1 起），跨页的表另带 `lastPage`，回原文核对直接翻到那一页；不要这个键就 `Options.source(false)`，三种格式一起关掉，docx 因此省掉整篇版面计算。
- **注册时给表起名**：`Options.name("季度明细")` 绑定写在它前面最近一次 `headers(...)` 的那一组，命中之后该组多一个 `name` 键；`WordParser.tablesByName(result, name)` / `rowsByName(result, name)` 按名字直接取出那几组和它们的数据行 —— 一张文档里只有要入库的那几张表需要名字，不注册就不写这个键。
- **图片四种输出模式**：只要元信息（默认）、base64 内嵌、边解析边落盘、原始字节；SHA-1 内容指纹 `contentId` 支持去重与调用方缓存。
- **文本清洗**：域代码丢掉、域缓存结果保留，软换行/分页符换成空格，控制字符清除，评估水印过滤。
- **对外部文档的安全处理**：写文件前只取文件名部分，目录内重名加序号，外链图片不读取、不联网，递归深度有上限。
- **无状态、可并发**：三个解析器都是静态方法，配好的 `Options` 可以复用（配置期间不要跨线程共享）。

## 3. 技术栈与环境要求

| 依赖 | 版本 | 用途 |
| --- | --- | --- |
| `com.aspose:aspose-words:jdk16` | 23.12 | doc/docx 解析 |
| `com.aspose:aspose-pdf:jdk16` | 23.12 | pdf 解析（`TableAbsorber` / `ParagraphAbsorber` / `ImagePlacementAbsorber`） |
| `com.aspose:aspose-slides:jdk16` | 23.12 | ppt/pptx 解析 |
| `com.fasterxml.jackson.core:jackson-databind` | 2.21.8（1.2.1 起） | `toJson` 序列化 |
| `org.junit.jupiter:junit-jupiter`（test 作用域） | 由父 pom 2.7.18 管理，JUnit 5.8.2 | 单元测试 |

环境要求：

- JDK 8 及以上（Aspose 构件是 `jdk16` classifier，Java 8 可直接用）：本机在 JDK `1.8.0_191` 与 JDK `21.0.1` 下
  各跑一遍完整测试都是 **242 条、0 失败 0 错误、1 条跳过**，两边产出的 class 文件都是 `major version: 52`（Java 8 字节码）——
  `java.version=1.8` 会让 maven-compiler-plugin 走 `-source/-target 1.8`，所以用高版本 JDK 编译**不会**抬高字节码版本；
- Maven 3.8.8（本机实测版本），需要能访问 `https://repository.aspose.com/repo/`；
- 无头服务器加 `-Djava.awt.headless=true`（Aspose 内部会用到 AWT 的图像类）；
- 父 pom 用 `spring-boot-starter-parent` **只为借它的依赖版本管理**，本模块没有 Spring 启动类，也没配 `spring-boot-maven-plugin`；
  测试只引 `junit-jupiter`、**不引 `spring-boot-starter-test`**，避免把 Spring 5.3.x、logback 1.2.x 等在该版本线上
  已没有安全修复版本的组件带进测试类路径。

## 4. 目录结构

```
document-parser/
├── pom.xml                                  # 普通库 jar；父 pom 只借版本管理；release profile 供发 Central 用
├── README.md
├── README.en.md                             # 英文版（给审核员与非中文使用者看）
├── LICENSE                                  # Apache License 2.0 全文
├── NOTICE                                   # 原作者署名 + 第三方组件（Aspose 商业授权）说明
├── CONTRIBUTING.md                          # 贡献指南：环境、约定、PR 流程、许可红线
├── CHANGELOG.md                             # 更新日志：每个版本对外可见的变化与实测数字（另有英文版）
├── DESIGN.md                                # 实现思路：模块职责、数据流、表头判定、各格式适配、边界与取舍
├── docs/index.html                          # GitHub Pages 项目主页
├── docs/excel-parser-selection.md           # Excel 解析选型报告（POI/EasyExcel/FastExcel/Aspose.Cells 实测对比），不进 jar
├── docs/json-to-db.md                       # JSON 入库示例：按表头签名路由、类型转换、整表原子写，不进 jar
├── .gitignore                               # target/、*.lic（私有凭证）、GPG 签名产物、IDEA 工程文件、Office 临时文件 ~$*
└── src
    ├── main/java/cn/net/tianxi/document
    │   ├── package-info.java                # 包级设计说明（返回结构、边界、许可证、安全）
    │   ├── WordParser.java                  # 入口 + Options/FullRow/ImageMode 定义 + Word 网格读取
    │   ├── PdfParser.java                   # 入口 + PDF 几何还原（聚类、跨页并表）
    │   ├── PptParser.java                   # 入口 + 形状分发（表格/文本框/图片/组合形状）
    │   ├── Tables.java                      # 表头分组算法（唯一一份，三种格式共用）
    │   ├── Images.java                      # 图片集合（唯一一份：编号、指纹、去重、四种模式）
    │   └── Results.java                     # 共用工具：JSON、文本清洗、水印过滤、许可证、文件名消毒
    └── test
        ├── java/cn/net/tianxi/document       # 15 个测试类、242 个测试（见「测试与质量」）
        │   ├── TestSupport.java             # 样例路径、断言助手、现场拼最小 docx/pptx 的零件
        │   ├── WordParserTest.java / PdfParserTest.java / PptParserTest.java   # 三份样例的端到端断言
        │   ├── PptParserTextTest.java       # 文本框拆行：三种换行、空段不占行（不走 Aspose）
        │   ├── CrossFormatTest.java         # 三种格式横向一致性
        │   ├── TablesTest.java / ImagesTest.java / OptionsTest.java / ResultsTest.java  # 共用组件的单元测试
        │   ├── WordEdgeCasesTest.java       # 现场拼最小 docx 覆盖边角结构
        │   ├── PdfBorderlessTableTest.java  # 现造无边框 PDF，验证流式引擎补读与已知误报
        │   ├── PdfCrossPageTableTest.java   # 现造三页长表 PDF，验证跨页并表与断链
        │   ├── PdfMergedCellTableTest.java  # 现造含合并单元格的 docx 再导出 PDF，验证合并行判据
        │   ├── TableByNameTest.java         # 注册时起名：name 键写在哪些组、按 name 取组取行
        │   └── PrintResultTest.java         # 不断言、只把解析结果打到控制台看（含 -Ddoc 入口）
        └── resources/{word,pdf,ppt}          # 三份业务样例（不打进产物 jar）
```

`Tables` / `Images` / `Results` 都是**包内可见**，解析器之外不暴露；对外只有三个解析器类。

## 5. 快速开始

> 本 README 描述的是 **1.2.1**（2026-10-24 已发布到 Central）。相对 1.2.0，1.2.1 只把 Jackson 升到 2.21.8，修复
> CVE-2026-110879（CVE-2026-54512 的不完整修复：泛型参数与数组元素类型仍可绕过 PolymorphicTypeValidator），
> API、返回结构与解析行为不变。相对 1.1.1，1.2.0 新增了 `source` 键、`name` 键与 `tablesByName` / `rowsByName`
> 入口、顶层 `unmatchedHeaders` 开关，并补读了 PDF 无边框表；三个主键 `text` / `tables` / `images` 的名字、
> `headers` 与行键的契约、四种图片模式在 1.1.1 里都一样，差别只是每组没有 `source`／`name` 这两个键，
> 以及 PDF 里无边框的表在 1.1.1 里读不出来。仍在使用 1.1.1 时，这些新内容都不可用；
> 完整变化见 [CHANGELOG.md](CHANGELOG.md) 的 1.2.1 与 1.2.0 两节、§17。

### 5.1 引入依赖

```xml
<dependency>
    <groupId>cn.net.tianxi</groupId>
    <artifactId>document-parser</artifactId>
    <version>1.2.1</version>
</dependency>
```

> 1.2.1 已于 2026-10-24 发布到 Central；若你用的镜像站还没同步，可以暂时把版本号换成 `1.2.0`，
> 代价是 Jackson 仍是 2.21.7。

Aspose 的三个构件来自 `https://repository.aspose.com/repo/`，引用方需要在自己的 `pom.xml` 里声明这个仓库
（本工程已经声明）。

### 5.2 最小示例

```java
Map<String, Object> result = WordParser.parse("D:/a.docx", "销售");   // PdfParser / PptParser 同名同参
String json = WordParser.toJson(result);
```

第二个参数是可选的**表头关键字**：一行里每个非空单元格都含关键字之一，才认它是表头行。
不给关键字、也不注册表头时是「静默模式」，允许拿每张表的第一行兜底当表头。

### 5.3 常用示例

```java
// 注册完整表头（跨页/跨格式时最稳），并丢掉分组小标题行
WordParser.Options options = WordParser.options()
        .headers("片区", "区域", "销售额", "回款率", "销售额", "回款率", "备注")   // 一组表头，可注册多组
        .name("季度明细")                             // 绑定写在它前面最近一次 headers(...) 的那一组
        .headers("客户名称", "合同金额", "已回款", "回款率", "备注")
        .name("回款台账")
        .headerTolerance(1)                          // 允许 1 列不一致（0 = 逐列相等），默认 1
        .fullRow(WordParser.FullRow.DROP)            // 跨整行合并的分组行直接丢弃
        .keywords("销售");                           // 可选，宽松兜底

Map<String, Object> result = PdfParser.parse("D:/a.pdf", options);
```

```java
// 入库：只有要进库的那几张表需要名字，按 name 取组取行，不必去数 tables 的下标
List<Map<String, Object>> tables = WordParser.tablesByName(result, "季度明细");
List<String> columns = (List<String>) tables.get(0).get("headers");   // 建表 / 映射字段照这份列清单
for (Map<String, String> row : WordParser.rowsByName(result, "季度明细")) {
    // row 的键逐个等于那组的 headers；同名分组按 tables 里的顺序摊平成一批行
}
// 起名的组在结果里多一个键：{ "headers": [...], "rows": [...], "name": "季度明细", "source": {"page":1} }
```

```java
// 图片边解析边落盘，结果里只留路径（图片多的文档推荐，不把字节留在内存里）
Map<String, Object> result = PptParser.parse("D:/a.pptx", WordParser.options()
        .images(WordParser.ImageMode.FILE)
        .imageDir(new File("D:/out/images")));       // FILE 模式必须先给 imageDir，否则抛 IllegalStateException
```

```java
// 流式解析（上传场景）：流由调用方关闭，文档句柄由 parse 负责释放
try (InputStream in = upload.getInputStream()) {
    Map<String, Object> result = WordParser.parse(in, WordParser.options());
}
```

### 5.4 先跑一条测试，直接看解析结果

不想写代码就想看某个文档被解析成了什么样，跑仓库自带的 `PrintResultTest`（**不断言、只打印**）：

```
mvn -B test -Dtest=PrintResultTest                                              # 全部 18 份报告
mvn -B test -Dtest=PrintResultTest#printAllSamplesWithRegisteredNames            # 看「注册 + 起名」后每组长什么样
mvn -B test -Dtest=PrintResultTest#printAllModeJson                              # 三格式 × 7 种选项的完整 JSON，另存 21 份 .json
mvn -B test -Dtest=PrintResultTest#printNamedMatchesAcrossFormats                # 看每个名字命中哪几组、拿到哪些行
mvn -B test -Dtest=PrintResultTest#printPublicApiOverloads                       # 四个 parse 重载逐一对拍（路径 / 流、选项 / 关键字）
mvn -B test -Dtest=PrintResultTest#printAnyDocument -Ddoc="D:/你的文档.docx"      # 看你自己的文档
```

控制台会给出正文逐段、每组表的注册名（起过名字才有）与来源页、表头、全部数据行、图片元信息和完整 JSON，
同时写一份 UTF-8 副本到 `target/print/`。输出样例与细节见 [13. 测试与质量](#13-测试与质量)。

## 6. API 一览

三个解析器的公开方法完全对应（`XxxParser` = `WordParser` / `PdfParser` / `PptParser`）：

| 方法 | 说明 |
| --- | --- |
| `static Map<String,Object> parse(String filePath, String... headerKeywords)` | 按路径解析，只给表头关键字 |
| `static Map<String,Object> parse(InputStream in, String... headerKeywords)` | 按流解析；流由调用方关闭 |
| `static Map<String,Object> parse(String filePath, WordParser.Options options)` | 按路径 + 完整选项解析 |
| `static Map<String,Object> parse(InputStream in, WordParser.Options options)` | 按流 + 完整选项解析 |
| `static String toJson(Object value)` | 把结果序列化成缩进 JSON |
| `static void writeImages(Map<String,Object> result, File directory)` | 把 `BASE64` 模式结果里的图片落成文件 |
| `static WordParser.Options options()`（仅 `WordParser`） | 新建一份选项，三种格式共用 |
| `static List<Map<String,Object>> tablesByName(Map<String,Object> result, String name)`（仅 `WordParser`） | 按注册时起的 `name` 取那几组表，返回的是结果里的同一批对象、按 `tables` 原顺序；没命中给空列表 |
| `static List<Map<String,String>> rowsByName(Map<String,Object> result, String name)`（仅 `WordParser`） | 把上面那几组的数据行按同序摊平成一个列表，逐行入库；没命中给空列表 |

两个取值入口都不解析、不抛异常（传 null 结果或 null/空白名字都是空列表），名字两端空格会被剪掉、中间的不剪。

`WordParser.Options` 的链式方法：`headers(String...)`、`name(String)`、`headerGroups(List<List<String>>)`、
`headerTolerance(int)`、`keywords(String...)`、`fullRow(FullRow)`、`images(ImageMode)`、`imageDir(File)`、`source(boolean)`。
两个枚举：`WordParser.FullRow`（`KEEP` / `FIRST` / `DROP`）、`WordParser.ImageMode`（`BASE64` / `FILE` / `BYTES` / `META`）。

`parse` / `toJson` / `writeImages` 都 `throws Exception`：解析的是外部文档，损坏文件、加密文件、
格式与扩展名不符都会抛出来，由调用方决定怎么兜（`options()` 与各链式方法不抛）。

## 7. 返回结构详解

下面是 `src/test/resources/word/sales-report.docx` 的**实测输出**（默认 `ImageMode.META`，图片只有元信息）：

```json
{
  "text": "2025 年度华东区销售业绩报告\n报告页次：3 页\n归档编号：\n…\n以上数据截至 2025 年 12 月 31 日，最终以财务系统结账数为准。",
  "tables": [
    { "headers": ["片区", "区域", "销售额", "回款率", "销售额-5", "回款率-6", "备注"],
      "rows": [ { "片区": "华东大区", "区域": "上海", "销售额": "1280", "回款率": "92%",
                  "销售额-5": "1465", "回款率-6": "95%", "备注": "超额完成" } ],
      "source": { "page": 1 } }
  ],
  "images": [
    { "fileName": "image-1.png", "type": "PNG", "width": 1024, "height": 768,
      "size": 670939, "contentId": "487e915db2c53907e8c1551aeaffea235423210c" }
  ]
}
```

### 7.1 `text`

全文文本，段落之间用 `\n` 拼接，**表格里的文字不在这里重复一遍**（唯一例外见「维护指南」里的结果契约）。
文本清洗集中在 `Results.clean` 一处：

- Word 用三段控制字符把域包成「域代码 + 域结果」，这里**丢掉代码只留缓存结果**
  （否则单元格里会冒出 `PAGE`，还可能被当成表头）。样例第 2 段实测 `报告页次：3 页`；
- 没算过的域整块消失。样例第 3 段实测 `归档编号：`（`DOCPROPERTY` 从没算过，域结果为空）；
- 软换行（Shift+Enter）、分页符换成**空格**而不是丢弃，样例第 4 段实测 `…统计口径为含税销售额。 数据来源：…`，
  句号后面那个空格就是软换行留下的痕迹；
- 单元格结束符等不可见控制字符丢弃，连续空格压成一个。

### 7.2 `tables`

一张物理表里有几组表头，`tables` 里就有几个元素；每个元素的 `rows` 只包含它上面那个表头的行。

- **表头判定优先级**（任一命中即算表头）：注册的完整表头集合 → 关键字 → 格式自带的标题行标记（Word 的 `w:tblHeader`）
  → **紧跟在跨整行合并行之后的行** → 表样式声明的首行（`tblLook firstRow` / `getFirstRow()`）→ 静默模式下的第一行兜底。
  两条容易踩的：分组小标题后面那一行会被当成新一段的表头（分组行之后直接放数据行会多出一组）；
  一旦注册了任何表头依据，「第一行兜底」就关掉，没命中的表首行按数据行处理、列名补成 `列N`
  （PDF 没有标题行标记可依赖，这条在它身上最明显，见 `PdfParserTest.registeringHeaderEvidenceTurnsOffTheFirstRowFallback`）。
- **列名规则**：空表头补 `列N`；合并摊平后同名的列补 `名-列号`（样例里 `销售额`/`回款率` 各出现两次 →
  `销售额-5`/`回款率-6`）；命中注册表头时用注册的原名，所以三种格式的键可以一致。
- **契约**：`headers` 与每一行的键**逐个相等**（顺序与数量都一致），行对象是 `LinkedHashMap`，
  按列顺序取值即可。`CrossFormatTest` 在三份样例 × 四种选项组合共 12 次解析上查过，不一致 0 处。
- **来源位置 `source`**：每组多一个 `source` 键 —— `{ "page": 起始页 }`，跨页的表再加 `"lastPage"`（实测样例里跨页的
  是那份回款台账，docx 与 pdf 各一份，都是 `{"page": 1, "lastPage": 2}`）。三种格式共用同一个键名与字段：docx 取 `LayoutCollector`
  算出的页码、pdf 取页码、pptx 取幻灯片序号，都从 1 起，回原文核对时照着这个数字翻页即可。
  同一张物理表切出的几组表头写的是同一段页码（按表头切组，不按页切）；某格式给不出位置时整键不写（实测三份样例
  每组都有，这条分支还没被触发过）。不要这个键就 `Options.source(false)`：三种格式一起生效，关掉后每组只剩
  `headers` 与 `rows`，列名与内容一字不变，docx 因此不建 `LayoutCollector`（耗时见 §8，一致性由
  `CrossFormatTest.sourceSwitchOffDropsTheKeyFromAllThreeFormats` 守着）。
- **注册名 `name`**：只在「注册了这组表头 + 注册时起了名 + 文档里确实命中」三者都成立时写，写在 `rows` 之后：

  ```json
  { "headers": ["片区", "区域", "销售额", "回款率", "销售额-5", "回款率-6", "备注"],
    "rows": [ … ],
    "name": "季度明细",
    "source": { "page": 1 } }
  ```

  所以默认的静默输出里一个 `name` 都没有（上面那份 4740 字符的 JSON 就是），`options()` 里起了名字的样例
  （`TestSupport.named()`：注册 5 组、3 组有名）实测 docx **4740 → 4828**、pdf **6473 → 6561**、
  pptx **3378 → 3466** 个字符，带名字的组分别是第 1、2、6、7 组／第 1、2、5、6 组／第 1、2、4、5 组。
  docx 的台账组也带上了名字：分组算法在多级表头合并完之后会拿合并结果再与注册比一次，所以注册
  「合并后那一行」同样命中（五种注册在 docx 上全部命中，标准错误为空；pdf/pptx 的差异见 §8 的注意事项）。
  取数用 `WordParser.tablesByName(result, "季度明细")`（同名多组按 `tables` 原顺序返回）或
  `WordParser.rowsByName(result, "季度明细")`（把同名各组的行拼平成一批，直接逐行入库），没命中给空列表。
  `name` 不是主键：同一张物理表切出的两个表头组会带同一个名字（docx 的第 6、7 组实测就是这种）。
  入库时如果要一个稳定的路由键，用表头本身（例如 `ObjectMapper.writeValueAsString(headers)`），不要依赖 `name`；
  按表头签名映射到目标表、类型转换（空串、`86%`、千分位）与整表原子写入的完整示例见
  [docs/json-to-db.md](docs/json-to-db.md)。
- **Word 单元格内嵌套表**：子表不进母表格子的文本（母表那一格留空），而是作为**额外元素**排在 `tables` 里，
  两边的行列结构都不丢。样例实测母表 `["工单类型","工单数","处理说明"]`（那一格是 `["硬件故障","42",""]`），
  子表另成一组 `["故障部件","工单数"]` 3 行，`source` 与母表同为第 2 页。

### 7.3 `images`

元信息（`fileName`、`type`、`width`、`height`、`size`、`contentId`）**任何模式都给**；内容去向由 `ImageMode` 决定：

| 模式 | 结果里多出来的字段 | 适用场景 |
| --- | --- | --- |
| `META`（默认） | 无 | 只要文字和表格入库，不为图片付体积和内存 |
| `BASE64` | `"base64": "data:image/png;base64,…"` | 前端直接塞进 `<img src>` |
| `BYTES` | `"bytes": byte[]`（Jackson 编成不带 `data:` 前缀的 base64） | Java 侧转存、送 OCR |
| `FILE` | `"path"`：`imageDir` 下的绝对路径 | 图片多、要边解析边落盘 |

几个要点：

- `contentId` 是图片字节的 **SHA-1 十六进制指纹**，只用来认「同一份字节是不是同一张图」和给调用方当缓存键；
  它不是存储位置、不是 URL，也不承担防篡改职责。docx 与 pptx 样例的这个值相同（同一份 PNG 字节），pdf 不同
  （Word 导出时把它重编成了 JPEG）—— 解析照实报告文件里的字节，不做还原。
- 内嵌模式（`BASE64`/`BYTES`）下内容重复的图片**不重复内嵌**，只留一个指向前一条的下标（`"duplicateOf": 0`，
  且不带内容字段，`contentId` 仍是它自己的指纹）；`META` 模式不做去重标记；`FILE` 模式下每张重复图片各自落一个文件，
  重名加 `-2` 后缀。三份样例各只有一张图，去重这条由 `ImagesTest` 直接喂字节验证。
- **外链图片**（Word「链接到文件」、PPT「仅链接」）的字节不在文档包里，只给 `link` 路径，本模块不会去读那个文件、
  也不联网，所以这类条目没有 `size`/`contentId`。
- `writeImages(result, dir)` 只认结果里的 `base64` 字段，因此**只对 `BASE64` 模式有效**；默认（`META`）下调用它
  一张也导不出，会在标准错误提示该换哪种模式。同理 `imageDir(File)` 只在 `FILE` 模式下生效，配错组合同时也会提示。

体积要心里有数（三份样例实测，见「实测基线」）：那份 docx 只有一张图，META 模式整串结果 4740 个字符，
BASE64 模式 899370 个字符 —— 那张 670939 字节的 PNG 编码后就是 894610 个字符，占了 99.5%。

## 8. 选项详解

`WordParser.Options` 是三种格式共用的可变对象，链式调用即改即生效。

| 选项 | 默认 | 含义与实测效果 |
| --- | --- | --- |
| `headers(String...)` | 不注册 | 注册一组完整表头（可多次调用注册多组）。命中后列名用注册的原名。样例实测：注册 5 组真表头后，docx 与 pdf 的 JSON 与静默模式**逐字节相同**，pptx 从 22 行变 21 行（下层表头行不再被当成数据）。比对时会把文档那一行**末尾的空列剪掉**，所以末列没有标题的表（docx 目标表）注册前 3 个列名就能命中、第 4 列仍补成 `列4`；注册组自己不要以空列名结尾，否则永远命不中 |
| `headerGroups(List<List<String>>)` | 空 | 批量注册，与多次 `headers(...)` 等价。注意它内部就是按顺序逐组注册，所以紧跟其后的 `name(...)` 绑的是**这一批的最后一组**；要给多组起名改用 `headers(...).name(...)` 的逐组写法 |
| `name(String)` | 不起名 | 给<b>写在它前面最近一次 `headers(...)` 注册的那一组</b>起名，命中之后该组多一个 `name` 键（见 §7）。取值技巧是跑 `PrintResultTest` 抄那几组的 `headers` 再顺手起名；入库时用 `WordParser.rowsByName(result, 名字)` 取行。注意四条：① 只有命中注册组的表头组才写这个键，靠关键字／`w:tblHeader`／首行兜底判出来的组没有；② 一个名字可以命中多组（同名各组按 `tables` 顺序都在 `tablesByName` 里），它不是主键；③ 命中判定跟着 `headerTolerance` 走，容差放宽了名字也可能跟着「误命中」；④ 多级表头三种格式给网格的行不一样 —— 实测注册「合并后那一行」三种格式都命中（合并完还会拿结果再比一次，容忍度收到 0 也命中）；改注册「上层那一行」则 docx 13 行、pptx 8 行（下层表头行落进数据）、pdf 命不中。跨格式共用一份注册时把上下两种写法各注册一次、起<b>同一个名字</b>（`TestSupport.namedBothHeaderRows()`，实测三份样例仍是 30／45／21 行）。名字首尾空格会被剪掉，空串／null 等于不起名；一组都没注册过就调用会在标准错误提示并忽略。改名等于改契约，下游按名字取数的代码会断 |
| `headerTolerance(int)` | 1 | 表头比对允许几列不一致；0 = 逐列相等。无论取值多少，都至少要求一列完全相等且列数相同；负数会被夹到 0 |
| `keywords(String...)` | 不给 | 宽松兜底：**行内每个非空单元格都含关键字之一**才算表头行。null 与空白项被忽略 |
| `fullRow(FullRow)` | `KEEP` | 跨整行合并行（分组小标题）的去向：`KEEP` 归到它上面最近的表头组、`FIRST` 统一归到该物理表的第一组、`DROP` 丢弃。样例实测 docx：KEEP 7 组／30 行、FIRST 同数量只改归属、DROP 6 组／28 行 |
| `images(ImageMode)` | `META` | 见 7.3 |
| `imageDir(File)` | null | 只在 `FILE` 模式下生效；`FILE` 没配它会在解析到第一张图时抛 `IllegalStateException`，配了却不是 `FILE` 会在标准错误提示 |
| `source(boolean)` | `true` | 每组表带不带 `source`（见 §7），**三种格式一起生效**。关掉之后 docx 不再创建 `LayoutCollector`，整篇版面就不用算了：同一份样例开／关交替实测 **74–254ms（开）／11–35ms（关）**；PDF／PPT 的页码本来就免费，开关只决定写不写这个键（同轮实测 pdf 681–960ms 对 696–984ms、pptx 46–86ms 对 44–71ms，区间重叠） |

**静默模式**（`isSilent()`：既没注册表头、也没给关键字）会打开「第一行兜底」，适合先跑一遍看看文档长什么样；
要稳定入库建议注册表头。选项本身没做同步，配置期间别跨线程共享；配好之后可以并发用于多次解析
（`Images` 每次解析新建，选项只读），`imageDir` 指向同一目录时重名图片由 `Images.uniqueFile` 加序号，不会互相覆盖。

## 9. 能力边界：读到什么、读不到什么

| 格式 | 读进结果 | 读不到（已知范围外） |
| --- | --- | --- |
| Word | 各节 Body 的段落、表格、图片；文本框文字；块级内容控件（`w:sdt`）包住的段落和表格；列表项（项目符号不进文本）；域**缓存的结果**（`SUM(ABOVE)`、`PAGE` 都读缓存值） | 页眉页脚、脚注尾注、批注；域代码一律丢弃、不做版面计算，所以没算过的域是空的；没填写的控件不会给出占位文字 `Click here to enter text.`（4 种写法实测，见下） |
| PDF | 有线段边框的表格（几何还原）、文字层段落、位图；跨页断开的长表并回一张（判据：相邻页 + 列数相同 + 左边界对齐 + 前一截是上页最后一块、后一截是本页第一块 + 后一截找不出表头行；静默模式没有表头依据可用，再要求「一截贴页底、一截贴页首」的几何证据）；纵向合并处被吸收器切成几块的行按矩形相邻聚回同一张表；**有线表之外再补一遍流式引擎**，只收「每行单元格数相同、且不与任何有线表重叠」的无边框表（实测无边框样例 0 组 → 3 组，`sales-report.pdf` 一组没多，见下） | 无边框表格只在**每行单元格数相同**时才补读，行数不齐的（首行是大标题那种）仍在 `text` 里；补读带来的误报：用制表位排齐的几行散文本会被当成一张表（实测，见下）；扫描件／图片型 PDF 需要 OCR；没注册表头依据、又被分页切断的那一行特别高（下边界离页底超过页高的 1/5）时认不出续接，一张表拆成两组、各带一个表头 |
| PPT | 正片各页的形状文字、表格、图片（组合形状递归展开，上限 32 层；`VideoFrame` 继承 `PictureFrame`，视频帧走图片那一支，收进 `images` 的是它的封面图） | 备注页／母版／版式页文字、SmartArt 图形内的文字、图表数据、嵌入的 OLE 对象（Excel 表／Word 文档）里的内容、形状填充底图、视频内容本身 |

几个实测确认过的边界：

- **没填写的内容控件不给占位文字**：4 种写法（`<w:p/>` 内容、空 `<w:sdtContent/>`、没有 `w:sdtPr`、带 `w:placeholderText`）
  实测 `text` 都只有控件外那一段，`Click here to enter text.` 一次都没出现 —— 那句提示存在 `w:sdtPr/w:placeholderText` 里，
  属于 Word 的显示层，不是文档内容（`WordEdgeCasesTest.unfilledContentControlContributesNothing`）。
- **导出成 PDF 会丢域结果**：这份 PDF 是用**加长版**报告（30 家客户，源文件不在仓库里）导出的，
  那份源 docx 的合计行两格是 `SUM(ABOVE)` 的缓存结果（当时用探针核对过 Aspose.Words 读到 `2582.30`/`2226.60`），
  Word 导出 PDF 时重算了这个域，PDF 里那两格只剩一个空格 → 解析如实读出空串
  （`["合计","","","86%",""]`，`PdfParserTest.exportedPdfLosesTheCachedSumFieldResult`）。
  仓库里的 docx 样例台账只有 12 家客户，合计行读的是缓存值 `2096.30`/`1796.10`/`86%`。**要金额请用 docx 那一份。**
- **PPT 两行表头建议注册表头**：静默模式下层表头行被当成数据，实测 8 行、列名
  `["客户名称","回款情况","回款情况-3","回款率","备注"]`；注册表头后正确合并成 7 行。
- **PPT 形状分发**：`Table`／`GroupShape`／`Chart`／`OleObjectFrame`／`SmartArtShape`／`PictureFrame`／`VideoFrame`／
  `AutoShape` 八个类逐个跑过 `isAssignableFrom`，只有 `AutoShape` 是 `IAutoShape`，四条分支互不重叠、顺序不影响结果；
  `Chart`、`OleObjectFrame`、`SmartArtShape` 一个接口都不沾，整形状被静默跳过 —— 这就是「图表数据／OLE 内容／
  SmartArt 文字读不到」的由来。

## 10. 测试样例

三份样例是同一份「2025 年度华东区销售业绩报告」，放在 `src/test/resources/{word,pdf,ppt}/` 下，**不打进产物 jar**：

| 样例 | 大小 | 怎么来的 | 内容 |
| --- | --- | --- | --- |
| `word/sales-report.docx` | 675208 字节 | 手工拼的 OOXML 压缩包（5 个部件，`word/document.xml` 33216 字节 + `word/media/image1.png`） | 13 段正文 + 4 张表 + 1 张 1024×768 的 PNG 看板截图 |
| `pdf/sales-report.pdf` | 184977 字节，3 页 | 用**真 Microsoft Word 2021** 把同源的加长版报告导出（文件里的 `Producer` 就是它），不是 Aspose 转的，所以没有评估水印 | 台账加长到 30 家客户，为的是让长表在第 1／2 页之间断开 |
| `ppt/sales-report.pptx` | 679289 字节 | 手工拼的 OOXML 压缩包（16 个部件，3 张幻灯片），合并单元格属性写在 `<a:tc>` 上：起点 `<a:tc gridSpan rowSpan>`、被覆盖格 `<a:tc hMerge="1">`/`vMerge` | 9 个文本框（其中 3 个在一个组合形状里）+ 4 张表 + 同一张 PNG |

**别用 Aspose 自己生成样例** —— 没有 lic 时它保存阶段就会把内容截断，期望值就写不死了。
pptx 样例里所有文字都压在 5 个字符以内，正是因为评估模式下 Aspose.Slides 只给文本框的第一段且截到 5 个字符；
`PptParserTest.everyTextFrameInTheSampleStaysWithinTheEvaluationLimit` 守着这条，
真要往样例里塞长文本，所有期望值都得在有 lic 的环境重测。

四张表各自压一种结构（三种格式同源）：

1. **季度销售明细**：单行表头、两组同名列（`销售额`/`回款率` 各出现两次）、`片区` 列纵向合并；
2. **客户回款台账**：两行表头且下一层留空（`回款情况` 跨两列，列名沿用上一层）、`合计` 行是 `SUM(ABOVE)` 域；
3. **售后工单统计**：Word 里「硬件故障」那格再套一张两列小表（`故障部件`/`工单数`），子表另成一组；
   导出成 PDF／做成 PPT 后这张嵌套小表不在了，所以注册它会在标准错误上收到「未匹配」提示；
4. **分片区年度目标**：跨整行合并的分组行（`2025 年度（已结账）`、`2026 年度（预算下达）`）、
   第四列没有标题（补成 `列4`）、同一张表里两组表头。

## 11. 实测基线

环境：Maven 3.8.8、Windows 11、**没有任何 `.lic`**，三个解析器都跑在评估模式；
JDK 用 `1.8.0_191` 与 `21.0.1` 各跑了一遍完整测试，两边都是 242 条、0 失败 0 错误、1 条跳过。
下面的数字全部是当前测试里的断言值，`mvn -B clean test` 就能复现。整轮耗时（1.2.1 依赖升级轮：
JDK 8 `clean package` 编译 + 测试 + 打包 **01:48**；JDK 21 同一条命令 **02:11**；加不加 `-DargLine`
那一行都是 242 条、0 失败、1 条跳过，它只影响控制台与报告里的中文编码。产物字节数 JDK 8 **60777**、JDK 21 **61190**。
作为对照，1.2.0 发版轮是 02:10 / 02:17、60236 / 60650 字节）。

**整轮耗时别当性能指标用**：1.2.1 依赖升级轮是 01:48（JDK 8）、02:11（JDK 21），1.2.0 发版轮是
02:10 与 02:17，上一轮空闲时是 01:25 与 01:32；早先机器被别的进程抢占时，同一条命令测到过
**02:15 ～ 02:37**，差出近一倍。能精确归因的只有 surefire 报告里的
每方法计时 —— `CrossFormatTest.documentedJsonLengthsAreTheRealOnes`（把下面这 **12 个 JSON 长度**逐个写成断言，
为此多解析 12 次）单条实测 **3.9 秒（1.2.0 那轮 JDK 21；1.2.1 那轮是 3.5 秒）**，早前 JDK 8 几轮在 **2.7 ～ 3.8 秒**之间
（`target/surefire-reports/TEST-*.xml`）。
早先文档写过「这条断言让整轮慢了 80 秒、单类 31.5 秒 → 76.5 秒」，那是拿一轮被抢占的耗时去比另一轮空闲耗时得出的，
**复现不出来**：把这条断言开／关交替各跑三轮，整类耗时差值完全被轮次噪声吃掉。那句已按实测撤回，好处不变：
文档里每一个长度数字都能被 `mvn test` 打脸，改坏了立刻红。
单份文件的解析耗时见 [§1.1](#11-与-ocr-的区别读文档里的文字层不猜像素)。

| 样例 | `text` | 表头组 | 数据行 | 图片 | META 结果长度 | BASE64 结果长度 |
| --- | --- | --- | --- | --- | --- | --- |
| docx | 13 段 | 7 | 30 | 1 张 PNG 1024×768、670939 字节 | 4740 字符 | 899370 字符 |
| pdf | 空串（正文全在表格里） | 6 | 45 | 1 张 JPEG 699×524、38668 字节 | 6473 字符 | 58076 字符 |
| pptx | 9 行（每行 ≤5 字符） | 6 | 22 | 1 张 PNG 1024×768、670939 字节 | 3526 字符 | 898156 字符 |

（长度比上一版各多 331／287／264 个字符，全部来自新加的 `source` 键：docx 7 组、pdf 6 组、pptx 6 组，每组多一段
`"source" : { "page" : N }`；**docx 与 pdf 的台账都跨页**、各多一个 `lastPage`，所以 docx 比 pptx 多的那一截
正好是 `7 × 44 + 23 = 331`。这三个差值是可复现的：从 Maven Central 拉下 `1.1.1` 的主 jar，对同一份样例跑一遍
META 模式得到 **4409／6186／3262**，与上表相减就是 331／287／264。）

同一份样例换选项后的组数／行数（也都是断言值）：

| 选项 | docx | pdf | pptx |
| --- | --- | --- | --- |
| 静默（默认，`FullRow.KEEP`） | 7 组 / 30 行 | 6 组 / 45 行 | 6 组 / 22 行 |
| 注册 5 组真表头 | 7 组 / 30 行，JSON 与静默逐字节相同 | 6 组 / 45 行，JSON 与静默逐字节相同 | 6 组 / **21** 行（下层表头行不再当数据） |
| 注册表头 + `DROP` | 6 组 / 28 行 | 5 组 / 43 行 | 5 组 / 19 行 |
| `FIRST`（不注册） | 7 组 / 30 行，只是分组行挪到该表第一组 | 6 组 / 45 行，同 | 6 组 / 22 行，同 |
| 注册 5 组 + 其中 3 组起名（`TestSupport.named()`） | 7 组 / 30 行，第 1、2、6、7 组带 `name`，4740 → **4828** 字符 | 6 组 / 45 行，第 1、2、5、6 组带 `name`，6473 → **6561** 字符 | 6 组 / 21 行，第 1、2、4、5 组带 `name`，3378 → **3466** 字符 |

三份样例的核心表（季度销售明细）在三种格式、静默与注册两种模式下**逐格相同**：
列名 `["片区","区域","销售额","回款率","销售额-5","回款率-6","备注"]`，4 行数据一字不差
（`CrossFormatTest.coreTableIsIdenticalAcrossThreeFormats`）。台账行数三份不同（docx 13 / pdf 31 / pptx 8）
是造样例时按各格式的可读边界裁的，不是解析丢了行。

### 样例覆盖到的行为（每条都有测试守着）

| 行为 | 实测结果 | 测试 |
| --- | --- | --- |
| 同名列补列号后缀 | `销售额`/`回款率` 第二次出现 → `销售额-5`/`回款率-6`，`headers` 与行键逐个相等 | `WordParserTest.sameNamedColumnsGetColumnNumberSuffix` |
| 纵向合并回填 | `片区` 列的江苏／安徽两行自己没有内容，读出 `华东大区`/`华东南区`；4 行数据都在，没被当成分组行 | `WordParserTest.verticallyMergedColumnIsFilledFromItsOriginRow`、`WordEdgeCasesTest.verticallyMergedDataRowIsNotAMistakenSeparator`、`PptParserTest.verticallyMergedDataRowIsNotAMistakenSeparator` |
| 两行表头合并 | 下层空着的列沿用上层名字 → `["客户名称","合同金额","已回款","回款率","备注"]` | `WordParserTest.twoLevelHeaderInheritsUpperLevelColumnNames`、`TablesTest.consecutiveHeaderRowsMergeIntoOneMultiLevelHeader` |
| 域结果 | `合计` 行读的是 Word 缓存的 `SUM(ABOVE)` 结果 `2096.30`/`1796.10`/`86%`；正文里 `PAGE` 域的缓存值 `3` 留着、域代码丢掉；没算过的 `DOCPROPERTY` 整块消失，只剩 `归档编号：` | `WordParserTest.sumFieldResultIsReadFromTheCachedValue`、`WordEdgeCasesTest.calculatedFieldResultSurvivesInsideACell` |
| 软换行 | Shift+Enter 换成空格：`…含税销售额。 数据来源：…`（最长那段 62 个字符完整读出，评估模式没截 Words 的正文） | `WordParserTest.textJoinsSoftBreakLinesWithASpace` |
| 块级内容控件 | 控件里包的段落进 `text`（`编制部门：销售管理部 审核：财务总监`）；套在 `w:tc` 里的控件，`Grid.readRow` 的浅层段落查找也看穿得到（`["控件里的甲","乙"]`） | `WordParserTest.textReadsParagraphInsideBlockContentControl`、`WordEdgeCasesTest.paragraphInsideCellLevelContentControlIsRead` |
| 没填写的控件不给占位文字 | 4 种写法实测：`text` 都只有控件外那一段，`Click here to enter text.` 一次都没出现 | `WordEdgeCasesTest.unfilledContentControlContributesNothing` |
| 单元格内嵌套表 | `tables` 1 → 2 组：母表那一格是空的（`["硬件故障","42",""]`），子表另成一组 `["故障部件","工单数"]` 3 行 | `WordParserTest.nestedTableBecomesItsOwnGroupAndLeavesTheHostCellClean` |
| 分组行三种去向 | KEEP 留在自己那组、FIRST 归到该表第一组（组数行数不变，只改归属）、DROP 整行丢掉（docx 7→6 组、30→28 行）；分组行**后面紧跟的那一行按表头处理** | `TablesTest` 的 `separatorRow*`／`fullRow*`／`rowAfterASeparatorRowIsTreatedAsAHeader` 六条 + 三个解析器各一条 |
| 表头之前就碰到分组行 | 另起一组、列名补成 `列1..列4`，分组文字不丢（docx 组 4、pdf 组 3、pptx 组 2 都是这种） | `TablesTest.separatorRowBeforeAnyHeaderStartsANumberedGroup` + 三个解析器各一条 |
| 全空的跨整行合并行不是分组行 | 现场拼的 4 行 2 列表（表头／`甲乙`／空白合并行／`丙丁`）：1 组、列名 `列A`/`列B`、3 行数据，空白行留成一行空值。改前实测是 2 组、第 2 组列名成了数据 `丙`/`丁` 且 0 行数据（`丙丁` 整行丢失）。三种格式同一条判据 | `WordEdgeCasesTest.blankFullWidthMergedRowIsNotASeparator`、`PptParserTest.blankFullWidthMergedRowIsNotASeparator`（pdf 一直是这条，见 `PdfParser.toGrid`） |
| PDF 跨页长表 | 31 行台账并回一组：第 1 页画到 `泰州医药城`（第 20 行），第 2 页从 `湖州建材贸易`（第 21 行）接着画，`合计` 在最后 | `PdfParserTest.ledgerSplitAcrossTwoPagesIsMergedBackIntoOneGroup`、`CrossFormatTest.onlyPdfNeedsTheCrossPageMerge` |
| PDF 纵向合并被切块 | 核心表的江苏／安徽两行、工单表的「硬件故障 12」那行都被吸收器切成独立小块，聚回来后仍是 4 行 | `PdfParserTest.coreTableRowsCutByVerticalMergeAreClusteredBack`、`PdfParserTest.ticketTableKeepsTheRowThatWasCutIntoItsOwnBlock` |
| 注册表头关掉首行兜底 | PDF 没有 `tblHeader` 标记可依赖：静默时首行当表头，注册了表头依据之后没命中的表首行按数据行处理、列名补成 `列N`（核心表 5 行 = 表头行 + 4 行数据） | `PdfParserTest.registeringHeaderEvidenceTurnsOffTheFirstRowFallback`、`TablesTest.firstRowStaysADataRowOnceHeaderEvidenceIsRegistered` |
| 注册名覆盖文档里读到的列名 | 默认容忍 1 列不一致：把 `备注` 注册成 `说明` 仍命中，结果里的列名用注册的 `说明`（PDF 的单元格文字会因换行多出空格，比对时忽略空格） | `WordParserTest.registeredHeadersOverrideTheNamesReadFromTheDocument`、`TablesTest.registeredHeadersOverrideTheNamesReadFromTheGrid` |
| 未匹配的注册表头会提示 | 标准错误上出现 `表头集合未匹配到任何行` + 那组列名；PDF／PPT 上提示的是 `故障部件`（嵌套小表在这两种格式里不存在），Word 上**不提示**（`named()` 的五组注册全部命中，标准错误为空）。**同名之下已有一份注册命中时不再提示** —— 把上下两种写法注册成同一个 `name` 之后，注定命不中的那一份不再刷噪声；没起名字的注册不受这条影响 | 三个解析器各一条 `unmatchedHeaderGroupIsReportedOnStderr`、`TableByNameTest.aRegistrationCoveredByASameNameMatchIsNoLongerReportedOnStderr` |
| 每组带来源页 `source` | docx：核心表第 1 页、台账 `{"page":1,"lastPage":2}`、其余 5 组第 2 页（嵌套子表与母表同页）；pdf：前两组第 1 页（台账跨到第 2 页）、工单第 2 页、后三组第 3 页；pptx：`page` 是幻灯片序号 1／2／3。同一张物理表切出的几组写同一段页码 | `CrossFormatTest.sourceRecordsThePageEachBlockStartsOn` |
| `source(false)` 三种格式一起关 | 关掉之后每组只剩 `headers` 与 `rows` 两个键，列名与内容一字不变；docx 因此不建 `LayoutCollector`（同一 JVM 里开／关交替实测 11–35ms 对 74–254ms） | `CrossFormatTest.sourceSwitchOffDropsTheKeyFromAllThreeFormats`、`OptionsTest.sourceIsWrittenByDefaultAndCanBeSwitchedOff` |
| 注册时起名写 `name` 键 | 只有「命中注册组且注册时起了名」的那几组多一个 `name`，键顺序固定 `headers` → `rows` → `name` → `source`；没起名的组与靠关键字／兜底判出来的组一个键都不多（关掉 `source` 后是 `[headers, rows, name]`）。docx 第 6、7 组带的是<b>同一个</b> `name`（同一张物理表切出的两组表头），`rowsByName` 把两组 5 行按组顺序拼平、返回的还是结果里的同一批对象。`tablesByName`／`rowsByName` 对未知名、空白、null 一律给空列表 | `TableByNameTest`（11 条）、`TablesTest.aRegisteredNameIsWrittenOnlyOnTheGroupsThatMatched`、`OptionsTest.nameBindsToTheGroupRegisteredJustBeforeIt` |
| 多级表头按哪一种写法注册（跨格式差异） | 注册「合并后那一行」（`客户名称/合同金额/已回款/回款率/备注`）：三种格式都命中，容忍度收到 0 也命中，台账行数 docx 13／pdf 31／pptx 7。注册「上层那一行」：docx 命中（13 行，上下两行都带 `w:tblHeader` 标记、照样合并）、pptx 命中但下层那行落进数据（7 → 8 行）、pdf 命不中（它的网格里只有合并后那一行，标准错误提示一条）。跨格式共用一份注册时把上下两种写法都注册、起同一个名字最稳：三份样例都带上 `name`，组数与行数不变（30／45／21） | `TableByNameTest.aMultiLevelHeaderHitsByItsMergedRowInAllThreeFormats`、`TableByNameTest.aRegistrationCoveredByASameNameMatchIsNoLongerReportedOnStderr` |
| PDF 无边框表补读 | 临时探针实测（跑完已删，同一份版面现在由 `PdfBorderlessTableTest` 守住）：Word 导出的「无边框 5×3 表 + 有线 5×3 表」默认只读到 1 张、加流式引擎读到 3 张；但**整体切换**会把全有线的仓库样例从 8 块读成 5 张并读坏（合并行左移、目标表多一列空列、嵌套行插错行次，35 格里 2 格、38 格里 12 格矩形无效），所以只收「每行格数相同 + 不与有线表重叠」的，仓库样例 5 张流式表全被挡掉、结果不变 | `PdfBorderlessTableTest.borderlessTableIsRecoveredAsAThreeColumnGrid`、`ruledTableOnTheSamePageIsUnaffected`、`nothingElseBecomesATable` |
| PDF 制表位散文本会被当成表 | 上面那条补读的已知误报：4 行×3 列、每行用制表位排齐的正文满足判据，于是从 `text` 挪进 `tables`（这条断言就是为了把这个行为写死，将来要改先动测试） | `PdfBorderlessTableTest.tabAlignedParagraphsCurrentlyCountAsATable` |
| PPT 两行表头的取舍 | 静默 8 行、列名 `["客户名称","回款情况","回款情况-3","回款率","备注"]`；注册表头后 7 行 —— 这就是「PPT 建议注册表头」的理由 | `PptParserTest.silentModeCannotMergeTheTwoLevelHeader`、`registeredHeadersMergeTheTwoLevelHeader` |
| PPT 组合形状 | `数据来源／销售系统／财务对账` 三个文本框在一个组合形状里，递归展开才读得到 | `PptParserTest.recursesIntoGroupShapes` |
| 图片四种模式 | META 只有元信息；BASE64 给 data URI（解码回来 670939 字节，与原图一致）；BYTES 给原始字节；FILE 落盘并给绝对路径、目录里只多一个文件；FILE 没配 `imageDir` 抛 `IllegalStateException`；META 配了 `imageDir` 在标准错误提示且不新增任何文件 | `WordParserTest` 的图片那 7 条 + `ImagesTest` 25 条 |
| 递归深度上限 | 12 层嵌套表 → 标准错误 `表格嵌套超过`，外层照常解析；20 层嵌套控件 → `节点嵌套超过`，那一支整个跳过 | `WordEdgeCasesTest.tablesNestedBeyondTheDepthLimitAreSkippedWithAWarning`、`containersNestedBeyondTheDepthLimitAreSkippedWithAWarning` |
| 结果契约 | 三份样例 × 四种选项组合共 12 次解析，`headers` 与每行键不一致 **0** 处，图片元信息齐全 | `CrossFormatTest.contractHoldsForEverySampleAndMode` 等 4 条 + 各解析器的 `everyOptionCombinationKeepsTheContract` |

## 12. 开源许可、Aspose 许可证与评估模式

### 12.1 开源许可与署名（本项目的代码）

本项目按 **[Apache License 2.0](LICENSE)** 开源，署名信息集中在 **[NOTICE](NOTICE)** 里
（`Copyright 2026 hutianxi (cn.net.tianxi)`）。

| 你可以 | 你需要 |
| --- | --- |
| 免费使用，含商业用途 | 分发时带上 `LICENSE` 与 `NOTICE`（或等价的可读副本） |
| 修改源码、做成衍生作品 | 在被改过的文件里注明「已修改」（Apache 2.0 第 4(b) 条） |
| 把它作为依赖打进自己的项目并**闭源发布** | 保留源码文件里的版权与署名信息 |
| 再分发原版或修改版 | 不要用作者名义为衍生作品背书（第 6 条） |
| 提 issue、提 PR，合并后成为贡献者 | 提交的贡献默认按同一许可授权（第 5 条），无需另签协议 |

给使用方的**署名写法**（放你自己的 `NOTICE`、「关于」页或 README 里，照抄即可。两个仓库地址都带上：
GitHub 是主仓库，Gitee 是国内镜像，哪边能打开都能找到源码）：

```
本产品包含 document-parser 的代码（cn.net.tianxi:document-parser），
Copyright 2026 hutianxi，Apache License 2.0，
原项目地址 https://github.com/hutianxi/document-parser
镜像地址   https://gitee.com/mailtianxi/document-parser
```

**为什么是 Apache 2.0**：需求是「免费用 + 可改 + 欢迎合并 + 必须保留原作者信息」。

| 许可 | 免费用/可改 | 可闭源衍生 | 明确专利授权 | 有 NOTICE 署名传递机制 | 结论 |
| --- | --- | --- | --- | --- | --- |
| MIT | 是 | 是 | 否（只有版权条款） | 否 | 够用，但专利风险没写清 |
| BSD-3-Clause | 是 | 是 | 否 | 否 | 同上，另带一条不背书条款 |
| **Apache 2.0** | 是 | 是 | **是**（第 3 条） | **是**（第 4(d) 条） | **采用** |
| GPL / AGPL | 是 | **否**（会传染，衍生作品必须同许可开源） | 是 | 是 | 与「允许别人闭源使用」冲突 |

Apache 2.0 是唯一同时满足这四条的主流宽松许可：宽松到别人可以闭源商用，又靠 `NOTICE` 把署名义务传递下去，
还额外给了专利授权（对做企业级文档处理的代码有意义）。

**贡献入口**：流程、代码约定、测试要求、许可红线都写在 [CONTRIBUTING.md](CONTRIBUTING.md)。
一句话版本：Fork → 从 `main` 切分支 → `mvn -B clean test` 全绿 → 提 PR（行为变更要附修改前后的端到端 JSON 对比）。

### 12.2 Aspose 许可证与评估模式（第三方商业依赖）

**本仓库不随包分发任何 Aspose 二进制**：`pom.xml` 只声明依赖，构件在构建时从
`https://repository.aspose.com/repo/` 下载。Aspose.Words / Aspose.PDF / Aspose.Slides 是 Aspose Pty Ltd 的
**商业软件，按 Aspose 自己的 EULA 授权，不在本项目的 Apache 2.0 覆盖范围内**——这一点也写进了 [NOTICE](NOTICE)
的第三方组件段落。生产使用请自行购买 Aspose 许可；`.lic` 文件是私有凭证，`.gitignore` 已拦住，任何 PR 都不要提交它。

`aspose-words.lic` / `aspose-pdf.lic` / `aspose-slides.lic` 放 **classpath 根目录**或**进程工作目录**即可，
每个产品在一个进程内只注册一次，且发生在**加载文档之前**（晚注册对已加载的文档无效）。
注册失败会在标准错误提示一行（早期版本静默 catch 掉，lic 过期时现象是「输出莫名其妙少了内容」）。

没有 lic 时按评估模式解析，**三个产品的限制并不相同**：

| 产品 | 限制 | 当前样例上看得见吗 |
| --- | --- | --- |
| Aspose.Slides | 读文本框只给**第一段**、且截到 5 个字符再追加提示；表格单元格走的也是文本框那一条，所以单元格同样被截 | 看得见，也正因为这样 pptx 样例把所有文字压到 ≤5 个字符（有测试守着），有没有 lic 读到的都是同一串 |
| Aspose.Words | 往页眉页脚注入水印段落 | 看得见：只遍历各节 Body、加上 `Results.cleanContent` 过滤，13 段正文里没有一句水印 |
| Aspose.PDF | 文字层段落读不全 | **看不见**：这份 PDF 的正文全在表格里，页面之间只有撑开用的空段落，`text` 是空串，所以这条限制在样例上量不出数字。表格与图片**不受影响** —— 31 行台账、45 行数据、1 张图片全部完整读出 |

评估模式留下的痕迹统一过滤，但**被截断的内容回不来**：PPT 那 5 个字符之外的文字、PDF 读不全的正文段落，
没有 lic 就拿不到，所以评估模式的输出只能当样例看，不能当结果用。

过滤细节：水印句子会被排版换行切成好几段（PDF 尤其明显），所以 `Results.isEvaluationNotice` 去掉所有空格、
转小写之后比对特征片段，被切成 `Pty Ltd.` 这样的孤行也认得出来；`PptParser.stripEvaluationNotice` 负责去掉截断提示。
已知代价：正文里真写到「Aspose Pty Ltd」这类字样时会被当成水印丢掉，接了 lic 就不触发。

样例本身与 lic 无关：docx／pptx 是手工拼的 OOXML，pdf 是真 Word 导出的，三份都不含评估水印；
pptx 靠「所有文字 ≤5 个字符」把 Slides 的截断绕开，docx 的 13 段正文（最长 62 字符）验证了 Words 不截正文。

## 13. 测试与质量

`mvn -B clean test`（`JAVA_HOME` 指向 JDK 8）跑完 **242 个测试、15 个测试类**（0 失败、1 跳过 ——
跳过的那条是 `-Ddoc` 入口，见下面「直接看解析结果」）：

| 测试类 | 个数 | 覆盖 |
| --- | --- | --- |
| `WordParserTest` | 32 | docx 样例端到端：正文 13 段、4 张表逐格、`FullRow` 三种去向、注册表头、图片四种模式、两个入口 |
| `PdfParserTest` | 17 | pdf 样例端到端：跨页并表、纵向合并聚块、丢域结果、注册表头关掉首行兜底、JPEG 重编 |
| `PptParserTest` | 19 | pptx 样例端到端：形状顺序、组合形状、两行表头的取舍、评估模式 5 字符自检、现场拼最小 pptx 验纵向合并不吃数据与全空合并行不当分组行 |
| `PptParserTextTest` | 2 | 文本框拆行（包级 `textLines`）：`\r\n`／`\r`／`\n` 三种换行都拆成独立段落，空段不占行。评估模式只给第一段，这条不走 Aspose |
| `PdfCrossPageTableTest` | 4 | 现场造三页长表 PDF：跨页并成一组、`source` 记首末页、中间页有别的表时链条断开、片段后面跟一张图时下一页的独立表也不会被错并、评估水印本身就能断链 |
| `PdfMergedCellTableTest` | 5 | 现场拼 docx 再导出 PDF：纵向合并与全空跨行合并不当分组行、docx 与 PDF 逐格一致、两行表头两种读法、水印只多一张图 |
| `CrossFormatTest` | 15 | 三种格式横向一致性：核心表逐格相同、`source` 页码逐组钉住、`source(false)` 三种格式一起关、图片指纹、一个 `Options` 复用、契约 12 次全查、文档里的 12 个 JSON 长度逐个钉住、`unmatchedHeaders` 顶层键、缺文件时三解析器的异常类型差异 |
| `PdfBorderlessTableTest` | 5 | PDF 无边框表：测试里用 Aspose.Words 现造一份「无边框表 + 有线表 + 制表位段落」的 docx 再导出 PDF，验证补读、有线表不受影响、表上方正文不被外接矩形吞掉、以及那条已知误报 |
| `TablesTest` | 34 | 表头分组算法（唯一一份）：六条判定优先级、多级表头合并、分组行三种去向、列名补全、网格补齐、注册名只写在命中的组上 |
| `ImagesTest` | 25 | 图片集合（唯一一份）：编号、类型归一化与嗅探、指纹、重复图片、四种模式、外链、文件名消毒、扩展名不跟着默认 `Locale` 走 |
| `OptionsTest` | 14 | `Options` 的判定语义：容忍度、关键字、静默模式、空值处理、`source` 默认开可关、`unmatchedHeaders` 默认关、`name` 绑最近一次注册（逐组与批量两种写法等价） |
| `TableByNameTest` | 11 | 起名与按名字取数：`name` 写在哪些组、键顺序 `headers`/`rows`/`name`/`source`、三种格式同一份注册的命中差异与上下两种写法都注册的解法、默认输出不带 `name`、`tablesByName`/`rowsByName` 的顺序与空值行为、同名之下已命中就不再报「未匹配」 |
| `ResultsTest` | 28 | 文本清洗、评估水印识别、表头归一化、JSON 序列化、文件名消毒、`writeImages`、许可证加载 |
| `WordEdgeCasesTest` | 14 | 现场拼最小 docx 覆盖边角结构：内容控件、嵌套表、深度上限（8 层不告警、9 层只提示一次）、纵向合并、域、软换行、全空合并行不当分组行 |
| `PrintResultTest` | 17（1 条默认跳过） | **不断言、只打印**，分四类共 **18 份报告**：一、三份样例的完整结果（静默 / 注册 / 起名 / DROP）；二、公共接口逐个跑（4 个 `parse` 重载对拍、`Options` 默认值与 setter、JSON 序列化与 `writeImages`、按名字取数）；三、表头规则与选项矩阵（台账四种注册写法、7 种选项的规模与完整 JSON、未匹配清单、重名 / 空表头 / 域缓存）；四、图片四种模式与错误路径。每条都写一份 UTF-8 副本到 `target/print/`（详见下一节） |

`TablesTest`／`ImagesTest`／`OptionsTest`／`ResultsTest` 不经过 Aspose（直接喂网格、字节和最小字符串），
本轮 JDK 8 单轮实测各 **0.011／0.025／0.001／0.023 秒**（`PptParserTextTest` 也只测字符串，**0.000 秒**）；
端到端那几类要加载 Aspose 与样例，最重的是
`CrossFormatTest` **29.8 秒**（三种格式反复解析，其中 `documentedJsonLengthsAreTheRealOnes` 单条 2.7 秒）
和 `PdfParserTest` **14.2 秒**（PDF 每页跑两遍表格引擎），`PrintResultTest` **24.0 秒**（18 份报告要把样例反复解析）、
`TableByNameTest` **7.3 秒**、`PdfCrossPageTableTest` **4.0 秒**、`WordParserTest` **1.8 秒**、
`PptParserTest` **0.94 秒**、`PdfBorderlessTableTest` **0.66 秒**（要在测试里现造 docx 再导出 PDF）、
`PdfMergedCellTableTest` **0.34 秒**、`WordEdgeCasesTest` **0.32 秒**。
JDK 21 同一轮里这几个数分别是 `CrossFormatTest` 38.3（`documentedJsonLengthsAreTheRealOnes` 3.5 秒）、
`PdfParserTest` 18.5、`PrintResultTest` 22.7、`TableByNameTest` 5.7、`PdfCrossPageTableTest` 4.8 秒。
整轮 `mvn -B clean package`（编译 + 测试 + 打包）1.2.1 依赖升级轮 JDK 8 **01:48**、JDK 21 **02:11**
（1.2.0 发版轮为 02:10 / 02:17）。
区间受机器负载影响很大（早先有别的进程抢 CPU 时，同一条命令测到 02:15 ～ 02:37；
之前同一条命令还测到过 JDK 8 01:32 ～ 01:42、JDK 21 01:55，`CrossFormatTest` 两轮差约 10 秒），
所以**耗时只看数量级、别当性能回归的判据**，详见 §11。

本轮构建核对（JDK 8，1.2.1）：`javac -Xlint:all -Xpkginfo:always` **0** 告警（17 个 class，含 `package-info.class`），
`javadoc -Xdoclint:all,-missing` **0** 错误，`mvn -B clean package` BUILD SUCCESS、**242 个测试 0 失败 0 错误**
（1 条 `-Ddoc` 入口跳过），主 jar **60777 字节 / 30 个条目 / 17 个 class**、`major version: 52`。
jar 字节数别当校验和用：只动 javadoc 注释也会让字节码里的行号表跟着挪，换 JDK 打包数字也变 ——
同一份源码本轮 JDK 8 打出 60777、JDK 21 打出 61190，条目数、class 数与字节码版本（`major version: 52`）三处都一致。

### 直接看解析结果（打印到控制台）

其余测试类都在断言实测值，能挡回归但看不出输出长什么样。`PrintResultTest` 就是给人看的：
跑一条测试，控制台直接给出**正文逐段、每组表的注册名（起过名字才有）与来源页、表头、全部数据行、图片元信息、完整 JSON**，
只保留一条底线断言（`headers` 与每行键的契约），因为契约破了输出也没参考价值。

```
mvn -B test -Dtest=PrintResultTest                                    # 全部 18 份报告（下面各条合起来）

# 一、三份样例的完整结果
mvn -B test -Dtest=PrintResultTest#printWordSample                    # 只看 docx（静默）
mvn -B test -Dtest=PrintResultTest#printPdfSample                     # 只看 pdf（静默）
mvn -B test -Dtest=PrintResultTest#printPptSample                     # 只看 pptx（静默）
mvn -B test -Dtest=PrintResultTest#printWordSampleWithRegisteredHeadersAndDrop  # docx 注册表头 + DROP
mvn -B test -Dtest=PrintResultTest#printAllSamplesWithRegisteredNames  # 三份样例的「注册 + 起名」完整报告（3 份）

# 二、公共接口方法逐个跑
mvn -B test -Dtest=PrintResultTest#printPublicApiOverloads             # 4 个 parse 重载：路径/流 × 选项/关键字，逐份对拍
mvn -B test -Dtest=PrintResultTest#printOptionDefaultsAndSetters       # Options 默认值、容忍度、批量注册、name 绑定、关键字
mvn -B test -Dtest=PrintResultTest#printSerializationAndImageExport    # toJson + writeImages：落盘大小与 SHA-1 对 contentId
mvn -B test -Dtest=PrintResultTest#printNamedMatchesAcrossFormats      # tablesByName / rowsByName：每个名字命中哪几组、哪些行

# 三、表头规则与选项矩阵
mvn -B test -Dtest=PrintResultTest#printLedgerNameRemedies             # docx 两行表头那张表：四种注册写法逐个对比
mvn -B test -Dtest=PrintResultTest#printOptionMatrix                    # 三份样例 × 7 种选项：组数 / 行数 / JSON 长度 / 键
mvn -B test -Dtest=PrintResultTest#printAllModeJson                     # 同一批 7 种选项的完整 JSON，另存 21 份 .json
mvn -B test -Dtest=PrintResultTest#printUnmatchedHeaders                # unmatchedHeaders 的开关与清单
mvn -B test -Dtest=PrintResultTest#printHeaderAndFieldRules             # 重名表头 -列号、空表头 列N、SUM(ABOVE) 只读缓存

# 四、图片与错误路径
mvn -B test -Dtest=PrintResultTest#printImageModes                      # META / BYTES / BASE64 / FILE 四种图片模式
mvn -B test -Dtest=PrintResultTest#printErrorPaths                      # 缺文件、FILE 少 imageDir 的异常类型
mvn -B test -Dtest=PrintResultTest#printAnyDocument -Ddoc="D:/你的文档.docx"   # 看自己的文档（按扩展名自动选解析器）
```

IDE 里点某一条测试的运行按钮效果一样。输出片段（docx，静默模式）：

```
==============================================================================
文档：D:\IdeaProjects\document-parser\src\test\resources\word\sales-report.docx（675208 字节）
选项：静默 —— 静默模式（未注册表头、未给关键字，允许首行兜底当表头）
       FullRow = KEEP，ImageMode = META
规模：正文 13 段，表格 7 组 / 30 行，图片 1 张
==============================================================================

---- text（13 段，表格里的文字不在这里）----
[1] 2025 年度华东区销售业绩报告
[2] 报告页次：3 页
[3] 归档编号：
…
---- tables（7 组）----

#1  7 列 × 4 行，来源 {page=1}
headers = [片区, 区域, 销售额, 回款率, 销售额-5, 回款率-6, 备注]
  行1: 片区=华东大区 | 区域=上海 | 销售额=1280 | 回款率=92% | 销售额-5=1465 | 回款率-6=95% | 备注=超额完成
…
---- images（1 张，META 模式只有元信息）----
[1] {fileName=image-1.png, type=PNG, width=1024, height=768, size=670939, contentId=487e915d…}

---- JSON（4740 个字符）----
{ … }
```

每条测试还会把同一份输出写一份 **UTF-8 副本**到 `target/print/`（末尾打印绝对路径），
控制台编码不合适时直接开这个文件；它在 `target/` 下，`mvn clean` 会带走。本轮实测 **18 份**副本共 **191974** 个字符
（`wc -m`，含 Jackson 美化输出在 Windows 上用的 `\r\n`；最短 1351、最长 106026）：
`docx-静默` 7950、`docx-注册表头-DROP` 7640、`docx-注册表头-起名` 8304、`pdf-静默` 10367、
`pdf-注册表头-起名` 10721、`pptx-静默` 5949、`pptx-注册表头-起名` 6072、`按名字取数-三份样例` 8568、
`docx-台账四种注册写法` 5511、`公共入口-四个重载` 1938、`选项默认值与setter` 1396、
`表头补名与域缓存` 2409、`序列化与图片导出` 1417、`选项矩阵-三份样例` 3210、
`全部模式-完整JSON` 106026（三份样例 × 7 种选项的 JSON 正文，另把 21 份单模式 JSON 写到 `target/print/json/`）、
`未匹配表头-三份样例` 1727、`docx-图片四种模式` 1418、`错误路径-三份样例` 1351。

**几份新报告看什么**：`全部模式-完整JSON` 把同一批三份样例 × 7 种选项的完整 JSON 原样打出来
（每个模式先一行摘要，再整份 JSON），并另存 21 份单模式 `.json`，要逐字段核对或做 diff 不必再从控制台里翻；
`公共入口-四个重载` 把三份样例的「路径 / 流 × 选项 / 关键字」四组调用逐一对拍，
路径与流的 JSON 逐字符相同（全 `true`）；`选项默认值与setter` 打出 `options()` 的每个默认值、容忍度收紧到 0
时的未命中提示、`headerGroups(List)` 与逐组 `headers(...)` 等价、`name(...)` 只绑最近一次注册；
`表头补名与域缓存` 用源文档复现「`销售额-5` / `回款率-6`」与「合计格有值」，并现造一份
「数据 1+2、`SUM(ABOVE)` 缓存写 300.00」的 docx，证明只读缓存不重算；`序列化与图片导出` 验证三个类的
`toJson` 输出相同、`writeImages` 落盘文件的 SHA-1 与 `contentId` 一致、`META` 结果调用时只提示不写文件。

起过名字之后每组的标题行多一节名字，照着它把要入库的那几组注册 + 起名最省事：

```
#1  7 列 × 4 行，name = 季度明细，来源 {page=1}
#6  4 列 × 3 行，name = 年度目标，来源 {page=2}
#7  4 列 × 2 行，name = 年度目标，来源 {page=2}      ← 同一张物理表切出的两组表头，带同一个 name
```

**抄列名时别把 `列N` 抄进注册值**：那是原文那一格为空时补出来的序号名，不是文档里的字
（`[片区, 年度目标, 完成率, 列4]` 这一组注册时只写前三个，第 4 列仍按位置补成 `列4`）。
另外 `TestSupport.named()` 在 docx 上第 1、2、6、7 组都带 `name`（台账那组靠「多级表头合并后再比一次」命中）；
三种格式的完整分布见 §8 的注意事项。

**想直接核对「名字 → 数据」对不对，跑 `printNamedMatchesAcrossFormats`**：它按名字把命中的组、
每组的全部行、`rowsByName` 拼平之后的行数一次打完，三份样例各一段，末尾还打三个取不到的名字。实测片段：

```
---- name = 「年度目标」：命中 2 组，rowsByName 拼平 5 行 ----
  tables 里的第 6 组，来源 {page=2}，4 列 × 3 行
  headers = [片区, 年度目标, 完成率, 列4]
    行1: 片区=华东大区 | 年度目标=2400.00 | 完成率=101% | 列4=达标
    行2: 片区=华东南区 | 年度目标=1500.00 | 完成率=93% | 列4=未达标
    行3: 片区=2026 年度（预算下达） | 年度目标= | 完成率= | 列4=
  tables 里的第 7 组，来源 {page=2}，4 列 × 2 行
  …
  rowsByName 与上面各组按组顺序接起来是同一批对象吗：true（拼平 5 行 / 接起来 5 行）

---- 取不到的名字：给空列表，不给 null、不抛异常 ----
  rowsByName(result, "没注册过的名字") = []，行数 0
  rowsByName(result, "")             = []，行数 0
  rowsByName(result, null)          = []，行数 0
```

**docx 那张两行表头的台账为什么要换种写法注册，跑 `printLedgerNameRemedies` 看对比**：四种写法各打一段，
每段列出 7 组的表头与名字、台账那组的全部 13 行、`rowsByName` 拿到的行数和标准错误原文。实测结论一眼可见 ——
① 默认容忍度 1、照合并后那一行注册：`带 name 的组：（一个都没有）`、`rowsByName` **0 行**、
标准错误提示未匹配；② 同一份注册把容忍度调到 2、③ 照文档里真实的上层那一行注册、④ 上下两种写法都注册起同一个名字：
三种都是 `#2=回款台账`、`rowsByName` **13 行**、标准错误**空**，而且 13 行数据逐字相同、7 组的表头与行数一字未变
（都是 7 组 / 30 行），② 也没有因为放宽容差把别的组带跑偏。

**不打 BASE64 模式**：默认 `META` 下三份样例的 JSON 是 4740／6473／3526 个字符，读得完；
换 `BASE64` 那张 670939 字节的 PNG 会把 JSON 撑到 899370 个字符，刷屏且看不出东西。
要看图片内容请用 `ImageMode.FILE` + `imageDir(File)`，输出里会带落盘路径。

### 在代码里验证一个文档

三个解析类都是纯静态方法，不依赖任何容器，也没有命令行入口。要在自己的代码里单看某个文档：

```java
Map<String, Object> result = WordParser.parse("src/test/resources/word/sales-report.docx",
        WordParser.options().keywords("销售").fullRow(WordParser.FullRow.DROP));
System.out.println(WordParser.toJson(result));
```

## 14. 常见问题

**怎么直接看到解析结果，而不是只看断言过不过？**
跑 `PrintResultTest`：`mvn -B test -Dtest=PrintResultTest`（三份样例），或
`mvn -B test -Dtest=PrintResultTest#printAnyDocument -Ddoc="D:/你的文档.pdf"` 看自己的文档。
它把正文、每组表的表头与数据行、图片元信息、完整 JSON 全打到控制台，并在 `target/print/` 留一份 UTF-8 副本。
详见 [13. 测试与质量](#13-测试与质量)。

**为什么注册了表头却总也命不中？**
两个常见原因：一是**列数必须相同**（容忍度只放宽「某几列名字对不上」，不放宽列数）；
二是比对前会把文档那一行**末尾的空列剪掉**，所以注册组自己以空列名结尾时永远命不中
（`OptionsTest.nullTitlesBecomeEmptyNamesButStillCountAsColumns` 守着这条取舍）。
反过来的用法是有效的：docx 目标表末列在文档里没有标题，注册 `["片区","年度目标","完成率"]` 三个名字就能命中，
结果里第 4 列仍按位置补成 `列4`。命不中时标准错误上会提示 `表头集合未匹配到任何行` + 那组列名，照着改注册值即可；
唯一不提示的情况是**同一个 `name` 之下已经有另一份注册命中了**（多级表头把上下两种写法注册到同一个名字时，
注定命不中的那一份不再刷噪声）。

**起了名字为什么结果里没有 `name` 键？**
`name` 只是贴在「已经命中这组注册表头」的那个表头组上的标签，不参与也不放宽命中判定。三种情况都不会写这个键：
这一组注册时没起名字；起了名字但这份文档里没命中（标准错误会提示 `表头集合未匹配到任何行` + 那组列名，
但同名之下已有另一份注册命中时不提示）；
这一组的表头是靠关键字、`w:tblHeader`、分组行后那条或首行兜底判出来的（没有注册组可对照）。
多级表头照「上层那一行」注册时，docx 与 pptx 能命中（pptx 下层那行会落进数据）、pdf 命不中；注册「合并后那一行」则三种格式都命中。跨格式共用一份注册时把上下两种写法注册成同一个名字最稳，见 [8. 选项详解](#8-选项详解)。
取数一律走 `WordParser.tablesByName` / `rowsByName`，它们对没命中的名字给空列表而不是抛异常。

**为什么 `text` 里没有表格里的字？**
按约定表格内容只出现在 `tables`，不重复一遍。唯一例外是 PDF 里取不到任何文字矩形的段落 ——
没法判断它在不在表格里，宁可让 `text` 多一句也不让正文少一段。

**为什么 PDF 样例的 `text` 是空串？**
那份 PDF 的正文全落在表格矩形内，页面之间只有撑开用的空段落，所以按上面的约定被全部跳过；
表格数据在 `tables` 里（6 组 45 行）。

**为什么注册表头之后，某张表的第一行变成数据行、列名成了 `列1`、`列2`？**
注册任何表头依据都会关掉「第一行兜底」。PDF／PPT 没有 Word 的 `w:tblHeader` 标记，
静默时靠首行兜底认表头，注册后没命中的表就按数据行处理了 —— 这是刻意行为，目的是不让兜底猜测盖过你给的依据。

**为什么 PPT 读出来的文字都只有几个字？**
评估模式：Aspose.Slides 只给每个文本框的第一段、并截到 5 个字符。放一个 `aspose-slides.lic` 即可。

**为什么合计行在 PDF 里是空的、在 docx 里有数？**
docx 读的是 Word 缓存的 `SUM(ABOVE)` 域结果；Word 导出 PDF 时重算了这个域，缓存值没写进 PDF。
要金额请用 docx。

**为什么 `writeImages` 一个文件都没写？**
它只认结果里的 `base64` 字段，因此只对 `ImageMode.BASE64` 有效。默认是 `META`（只有元信息），
此时调用它会在标准错误提示该换哪种模式；要边解析边落盘请用 `FILE` + `imageDir`。

**为什么 `ImageMode.FILE` 抛 `IllegalStateException`？**
`FILE` 模式必须先 `imageDir(File)`，否则解析到第一张图时抛出（提示里写着要先指定导出目录）。

**解析会不会读到文档外的文件或联网？**
不会。外链图片只回报 `link` 路径；写文件前文件名一律过 `Results.safeName`，只取文件名部分，
目录内重名加序号，所以 `../../evil.png` 这种名字也只能落在目标目录里。

**能并发调用吗？**
能。三个解析器都是无状态静态方法；`Options` 是可变对象，配置期间不要跨线程共享，配好之后可以并发复用。

## 15. 已知限制与未覆盖分支

按代码路径改的防御分支，仓库里没有对应样例，也就**没有测试覆盖**，谈验证得再造边角样例：

- PDF 静默选项下「两张独立表恰好等列数、左边界对齐」的误并**反向**用例（当时的数字见附录，是靠探针按坐标构造的）；
- `Tables.markSource` 的 `page <= 0` 分支（给不出位置就不写 `source` 键）：三份样例 19 组实测每组都拿得到页码，
  这条分支一次没触发过；文本框里的表、页眉页脚里的表这类版面例外要等有了样例再验证；
- `WordParser.Options.headerName(int)` 的下标越界保护与 `tablesByName`／`rowsByName` 里「`tables` 不是列表」
  「条目不是 Map」两道守卫：`headerLabels` 与 `headerGroups` 是一一对应地加进去的、结果里的 `tables` 一律是本模块
  自己放进去的 `List<Map<String,Object>>`，正常调用走不到这几条，写它们只是防以后两边不同步；
  `TableByNameTest` 覆盖的是 null 结果、null／空白／未知名与首尾空格这些真会发生的入参；
- `PdfParser.flowTableFlow` 里「流式表拿不到任何有效矩形」的守卫分支：探针那 5 张流式表全有矩形，没触发过；
- PPT 单元格没有文本框时的空值保护；
- 外链图片（Word「链接到文件」、PPT「仅链接」）：`ImagesTest` 直接喂 `link` 路径验证了元信息形状，
  但三份样例里都没有真的外链图片；
- Aspose.PDF 评估模式对**文字层**的具体截断量：当前样例的 `text` 是空串，量不出数字，接入 lic 后请自行核对。

大文件与内存：`BASE64`/`BYTES` 模式会把图片字节都留在结果里，图片多的文档请用 `FILE` 或 `META`；
OOXML（docx/pptx）是压缩包，解压后全量进内存，压缩比极高的畸形文件可能直接把进程撑爆，
所以上传/落盘环节要有大小上限，解析调用最好带超时与独立线程池。

## 16. 维护指南

对外部贡献者的要求（环境、代码约定、PR 流程、许可红线）写在 [CONTRIBUTING.md](CONTRIBUTING.md)，
下面这一节是维护者视角的模块分工。

| 文件 | 职责 |
| --- | --- |
| `WordParser` / `PdfParser` / `PptParser` | 只负责「把自家格式读成 `String[][]` 网格 + 段落 + 图片字节」，以及各自的 `Options` 入口 |
| `Tables` | 表头分组算法（唯一一份）：判定优先级、多级表头合并、跨整行合并行的去向、列名补全（空表头→`列N`、同名→加列号后缀，`headers` 与行键同源）、注册命中时写 `name`（`markName`） |
| `Images` | 图片集合（唯一一份）：文件名、元信息、内容指纹、重复图片、四种输出模式 |
| `Results` | 共用工具：JSON 序列化、文本清洗（域代码／软换行）与评估水印过滤、表头归一化、许可证加载、写文件前的文件名消毒 |

改行为时按这个分工改一处即可，不要在各解析器里抄一份。三种格式的返回结构必须一致，`Tables`/`Images`/`Results`
是包内可见的，解析器之外不暴露。

**加一种新格式**（例如 excel）：写一个 `XxxParser`，把表格读成 `String[][]`（合并覆盖的位置回填起始格的值）、
`boolean[] separator`（跨整行合并的行）、可选 `boolean[] heading`（格式自带的标题行标记），然后调
`Tables.group(...)`；图片交给 `new Images(options).add(...)`；最后按 `text/tables/images` 三个键组装结果。
`Options`、`toJson`、`writeImages` 直接复用，不要另写一套。

**结果契约**（改动时别破坏，`CrossFormatTest` 会查）：一张物理表（含 Word 的嵌套表，子表另成一条）有几组表头，
`tables` 就有几个元素；同一组里每行的键数量一致，且 `headers` 逐个等于每行的键（空表头补 `列N`、
同名列补 `名-列号`，两处规则由 `Tables.normalizeHeaders` 共用一份）；命中注册表头时列名用注册的原名
（三种格式的键因此相同）；表格里的文字不在 `text` 里重复出现；图片顺序 = 文档流顺序；
每组的 `source.page` 在三种格式下含义相同（页码／幻灯片序号，1 起），由 `Tables.markSource` 统一写，
分组算法 `Tables.group` 本身不碰它；`Options.source(false)` 时这个键**三种格式一起消失**，每组只剩
`headers` 与 `rows`，其余契约一字不变（各解析器先问 `Options.writesSource()`，Word 侧因此连
`LayoutCollector` 都不创建）。
每组的键顺序固定是 `headers` → `rows` → `name`（注册且命中并起过名字才有）→ `source`（`source(true)` 才有）。
`name` 由 `Tables.group` 在命中注册组的那一刻写（`markName`）—— 与 `markSource` 一个在内一个在外是刻意的：
命中的是第几组只有分组算法知道，页码只有版面知道。加新格式时这两件事都不用管，把 `Options` 交给 `Tables.group` 就行；
但也**不要**在某个解析器里另写一份起名或页码逻辑，那样三种格式就又不一致了。

**改样例之后**：所有期望值都是实测出来的，改了样例就把对应测试类跑一遍、按新输出改断言，
再把「实测基线」两张表里的数字同步过来；javadoc 里引用的样例数字（三个解析器类说明、`Results`、`Images`、
`package-info`）也要一起改。

**注册名同样是对外契约**：下游用 `WordParser.tablesByName` / `rowsByName` 按名字取数，改名或改命中口径都会
让别人的代码断掉 —— 与改列名同级别，属于要在更新日志里写清楚的变化。

**对外可见的变化记进更新日志**：返回结构里的键、`Options` 的方法、公开入口、行为口径只要有改动，就往
[CHANGELOG.md](CHANGELOG.md) 与 [CHANGELOG.en.md](CHANGELOG.en.md) 最上方的未发布一节各补一条，带上三份样例上的
实测数字；发布时才把这一节收进正式版本条目，并同时升 `pom.xml` 的版本号。纯内部重构与只改注释不写。

## 17. 构建

```
mvn -B clean test          # JAVA_HOME 指向 JDK 8；本轮实测 242 条测试（241 通过、1 条 -Ddoc 入口跳过）
mvn -B clean package       # 产物 target/document-parser-1.2.1.jar（本轮实测 60777 字节、30 个条目、17 个 class，
                           # 类直接放在 jar 根目录的 cn/net/tianxi/document/ 下；
                           # prepare-package 阶段会把根目录的 LICENSE 与 NOTICE 复制进 META-INF/，
                           # 所以只拿到 jar 的使用方也看得到署名与第三方授权说明）
```

javadoc 的 doclint 核对可以用 JDK 自带的 `javadoc` 直接跑（不经过 Maven，便于快速复核）；
`release` profile 里的 `maven-javadoc-plugin` 3.4.1 在本轮也已成功产出 javadoc jar：

```
mvn -B dependency:build-classpath -Dmdep.outputFile=target/cp.txt
javadoc -encoding UTF-8 -charset UTF-8 -Xdoclint:all,-missing \
        -cp "$(cat target/cp.txt)" -d target/site/apidocs $(find src/main/java -name '*.java')
```

给别的工程当 `<dependency>` 引的就是那个 jar。样例文档在 `src/test/resources/{word,pdf,ppt}/` 下，不会被打进产物 jar。

将来真要打成可执行 fat jar，记得给 `spring-boot-maven-plugin` 加 `<classifier>exec</classifier>`：少了这一行，
repackage 会把主构件原地换成 fat jar（日志 `Replacing main artifact with repackaged archive`），
类全跑进 `BOOT-INF/classes/`，别人引依赖时一个类都找不到。

### 发布到 Maven Central

`pom.xml` 里的 `release` profile 就是为此准备的：`maven-source-plugin`（源码包）+ `maven-javadoc-plugin`
（javadoc 包）+ `maven-gpg-plugin`（签名）+ `central-publishing-maven-plugin`（打包并上传到 Central Portal）。
平时 `mvn clean test/package` 不带这个 profile，所以构建速度和本地仓库都不受影响。

**坐标为什么是 `cn.net.tianxi`**：Central 要求 `groupId` 属于你在 Portal 上已验证的命名空间。本项目的
`cn.net.tianxi` 是**域名型命名空间**（Org: Tianxi，状态 Verified），验证方式是在 `tianxi.net.cn` 的 DNS 上
加一条 Portal 给出的 TXT 记录；它与 Java 包名 `cn.net.tianxi.document` 天然一致，源码不用动。

> 历史说明：1.1.1 首次上传时 `cn.net.tianxi` 还没验证完，先用 GitHub 型命名空间发过一次
> `io.github.hutianxi:document-parser:1.1.1`。Central 的构件一旦 Publish 就**不能删也不能改**，所以那个坐标
> 会一直留在中央仓库里；本项目后续版本只发 `cn.net.tianxi`，请用上面 5.1 节的坐标。

**签名密钥**：RSA 4096，主密钥 key id `41AF7E29CB53074D`、指纹 `5BBC2118BD5BC8805B65F69041AF7E29CB53074D`，
实际签名可能落在它的子密钥 `21AAF067632A8E6A`（指纹 `CCE99252368722C3498F8B9021AAF067632A8E6A`）上，
`gpg --verify` 实测 `Good signature`。主密钥与子密钥都已推到 `keyserver.ubuntu.com`
（用 `https://keyserver.ubuntu.com/pks/lookup?op=get&search=0x<指纹>` 取回，两者实测都是 HTTP 200）。
Central 校验签名时会去公钥服务器查这把公钥，查不到就打回。
本机这把私钥**没有口令**（为了让构建能全自动跑）；要加口令：`gpg --edit-key 41AF7E29CB53074D` → `passwd`，
之后构建时把口令配在 settings.xml 的 `<server><id>gpg</id><passphrase>` 里。

发布有两条路，1.1.1、1.2.0 与 1.2.1 都走第一条（**不需要把 Token 落到磁盘上**）：

**A. 本地打 bundle + Portal 网页上传**

```
mvn -B -P release deploy -DskipPublishing=true      # 版本号不能带 -SNAPSHOT
```

产出 `target/central-publishing/central-bundle.zip`（1.1.1 那次实测 266857 字节、1.2.1 这次 236990 字节，都是 24 个条目）：
1.1.1 的 zip 里是 `cn/net/tianxi/document-parser/1.1.1/` 下的 `document-parser-1.1.1.jar`（56275 字节）、
`-sources.jar`（64964 字节）、`-javadoc.jar`（141571 字节）、`.pom`，各带一个 `.asc` 签名和
`md5/sha1/sha256/sha512` 四个校验和。

然后到 Portal → **Publishing → Publish Component**：填 Deployment Name（如 `document-parser 1.2.1 (cn.net.tianxi)`）
与描述，Choose File 选这个 zip，点 Publish Component。提交后 Deployments 列表里出现一条 **VALIDATING** 记录
（1.1.1 那次是 `960f618b-cb17-4cb1-9d40-96849df2a219`；1.2.1 是 `11cdf8ca-0884-4626-9356-c82a6455144e`），校验通过后同一行会出现
**Publish** 按钮，点它才真正公开。

两个坑（都实测踩过）：

- 跳过上传的属性名是 **`-DskipPublishing=true`**，写成 `-Dcentral.skipPublishing=true` 不生效，插件会拿占位凭证真去上传并报
  `Deployment failed while publishing`；
- 即使 `skipPublishing=true`，插件仍会去 `settings.xml` 读 `<server><id>central</id>`，读不到直接
  `Unable to get publisher server properties for server id: central: NullPointerException`。
  所以只打 bundle 也要给一段占位配置（`-s` 指到一个临时 settings 文件即可，不必污染全局配置）。

**B. 命令行直传**（适合 CI）：`settings.xml` 里放 Portal 的 User Token 与 GPG 口令，然后
`mvn -B -P release clean deploy`，插件自己上传；`autoPublish=false`，同样要去网页点 Publish。

```xml
<servers>
  <server>
    <id>central</id>
    <username>Portal 生成的 Token 用户名</username>
    <password>Portal 生成的 Token 密码</password>
  </server>
  <server>
    <id>gpg</id>
    <passphrase>密钥口令（本机这把无口令，留空）</passphrase>
  </server>
</servers>
```

校验会查这些项：有没有 sources 包和 javadoc 包、pom 元数据全不全（`name`/`description`/`url`/`licenses`/
`developers`/`scm`，本项目都已补齐）、签名能不能在公钥服务器上验通、`groupId` 是否属于已验证命名空间。
没过就 **Drop** 掉改完重来；一旦 Publish 就撤不回，同一版本号也不能再传第二次，只能升版本号。

Publish 之后同步到 `repo1.maven.org` 需要几十分钟到几小时（搜索索引更慢），验证地址：

```
https://repo1.maven.org/maven2/cn/net/tianxi/document-parser/1.2.1/
```

1.1.1 与 1.2.1 都已经走完全流程并**实测验证过**：

- 1.1.1：Portal Deployment `960f618b-cb17-4cb1-9d40-96849df2a219` 状态 PUBLISHED，目录里的
  jar / sources.jar / javadoc.jar 字节数与**当时**的本机产物一致（56275 / 64964 / 141571）；
- 1.2.1：Portal Deployment `11cdf8ca-0884-4626-9356-c82a6455144e` 状态 PUBLISHED（2026-10-24），
  Central 上的 jar 为 60777 字节、`.pom` 为 12136 字节、sources.jar 为 81550 字节、javadoc.jar 为 89001 字节，
  与本机发布包一致；四份 `.sha256` 也逐个核对过 —— jar `ce9af926820b55b18505d2efa51a02cbd92dbe73376fda7a46d3fec6298c0ca3`、
  sources `1d2fbc27afe3974036c4c6233b24577b94d7b4d014403af4caebb0b31e68420c`、
  javadoc `94ac6f87c12143e2d122009373db6afab727a5d5ea4a36c193e49440f6f94e8c`、
  pom `e6fda84c361fbfafc596490bc26134e25584a41f7c0cf39eec07e3c19dc78e53`。

1.1.1 还用过一个**全新的空本地仓库**跑 `mvn dependency:get -Dartifact=cn.net.tianxi:document-parser:1.1.1 -Dtransitive=false`
（`-Dmaven.repo.local` 指到临时目录，避免命中本机已装的同名构件），BUILD SUCCESS，下到的 jar 是
`https://repo.maven.apache.org/maven2/cn/net/tianxi/document-parser/1.1.1/document-parser-1.1.1.jar`，
sha256 `c2c47ba92823ee9464a9a80bb6125ee8b8aa34f562ee80d76f709df3ce233148`，与 Central 上的 `.sha256`、
本地产物、GitHub Release 附件三处完全相同。临时仓库用完已删。

> **版本号与产物的一致性**：README 与 CHANGELOG 现在描述的是 **1.2.1**（2026-10-24 已发布）。Central 上的
> 1.2.1（主 jar 60777 字节）与 1.2.0（60236 字节）API 完全相同，差别只在 Jackson 2.21.8 与 2.21.7。

### 1.2.1（2026-10-24）

`pom.xml` 已升到 `1.2.1`；本机 JDK 8 `mvn -B clean package` 打出主 jar **60777 字节**
（JDK 21 打同一份源码是 **61190 字节**）、30 个条目、17 个 class，242 条测试、1 条跳过（241 通过）。
`sources.jar` 与 `javadoc.jar` 由 `release` profile 生成；2026-10-24 02:32 打出的 bundle 里，
JDK 8 产物是 **81550** 与 **89001** 字节（javadoc 页脚带生成时间，这个 jar 每轮会差几个字节）。
主 jar sha256 `ce9af926820b55b18505d2efa51a02cbd92dbe73376fda7a46d3fec6298c0ca3`
（jar 条目带构建时间戳，重打即变，发布时以 Central 上的记录为准）。
本地已用 `-P release deploy -DskipPublishing=true` 打出 `target/central-publishing/central-bundle.zip`
（**236990 字节、24 个文件**，sha256 `fdcd57e91f009d3c3438c95a75dce74ccda23baa327e1f52eac30ae3ab422612`），
四个 `.asc` 均 `gpg --verify` 通过（实际签名落在子密钥 `21AAF067632A8E6A` 上）；已上传并在 2026-10-24 发布。

改动只有依赖：Jackson 2.13.5 → 2.21.8，补齐 CVE-2026-83557 / 68497 / 19032 / 89425 / 89407 / 91776，
并修复 CVE-2026-110879 / GHSA-m8xf-2cr7-77ph（2026-10-20 发布，CVE-2026-54512 的不完整修复）。
2.21.8 的 `DatabindContext` 里新增了数组元素类型校验，`ApprovedBaseTypeValidator` 也只在这个版本出现；
2.21.7 及更早版本没有这些代码。Jackson 2.x 仍是 Java 8 字节码（`major version: 52`）。
测试依赖从 `spring-boot-starter-test` 换成 `junit-jupiter`，不再把 Spring 5.3.x / logback 1.2.x 等带进测试类路径。
对外 API、返回结构与解析行为不变。

上传与验证记录（2026-10-24）：Portal Deployment `11cdf8ca-0884-4626-9356-c82a6455144e`，状态 **PUBLISHED**；
`repo1.maven.org` 上 `document-parser-1.2.1.jar` 为 **60777 字节**、`.pom` 为 **12136 字节**、
`sources.jar` 为 **81550 字节**、`javadoc.jar` 为 **89001 字节**，与本机发布包一致。四份 `.sha256` 逐个核对：
jar `ce9af926820b55b18505d2efa51a02cbd92dbe73376fda7a46d3fec6298c0ca3`、
sources `1d2fbc27afe3974036c4c6233b24577b94d7b4d014403af4caebb0b31e68420c`、
javadoc `94ac6f87c12143e2d122009373db6afab727a5d5ea4a36c193e49440f6f94e8c`、
pom `e6fda84c361fbfafc596490bc26134e25584a41f7c0cf39eec07e3c19dc78e53`。

### 1.2.0（2026-10-23）

`pom.xml` 已升到 `1.2.0`；本机 `mvn -B clean package` 在 JDK 8 上打出主 jar **60236 字节**
（JDK 21 打同一份源码是 **60650 字节**）、30 个条目、17 个 class，242 条测试、1 条跳过（241 通过）。
sources.jar 与 javadoc.jar 由 `release` profile 在打包时生成。

**上传与验证记录（2026-10-24 实测）**：Portal Deployment `923fda1a-a25c-46f3-b83d-a387966a5def` 状态
**PUBLISHED**；`repo1.maven.org` 上的产物与本机构建逐字节一致：

| 产物 | 字节数 | sha256 |
| --- | --- | --- |
| `document-parser-1.2.0.jar` | 60236 | `92db3a721654ea142aff88f6735cca5a557576febde5251b80b784aabf25a4b4` |
| `document-parser-1.2.0-sources.jar` | 81011 | `f564cd6626cb9df39d191d3bdb53f0dc6acee6932871f5dc51749690f9c39139` |
| `document-parser-1.2.0-javadoc.jar` | 88998 | `aea880dfe09ca25af534571c4f01b8521d00aeb5190c966896f807c6c831aaca` |
| `document-parser-1.2.0.pom` | 11118 | `fd7b4a0c976ff730c822e9d4d16e83d73c41e2ab1db2cd8403d7e0f1dfd5a2c5` |

上传用的 `central-bundle.zip` 是 235380 字节、24 个文件（解压后 245751 字节），本机 sha256
`402f1d7e39f7790c0f6796548d1f6beba4cdcede6e5c6123a364dd97152fac82`；它只是上传容器，Central 上不单独提供。

再用一个**全新的空本地仓库**跑 `mvn dependency:get -Dartifact=cn.net.tianxi:document-parser:1.2.0
-Dtransitive=false`（`-Dmaven.repo.local` 指到临时目录），结果 **BUILD SUCCESS**，jar 从
`https://repo.maven.apache.org/maven2/cn/net/tianxi/document-parser/1.2.0/document-parser-1.2.0.jar` 取回，
60236 字节、sha256 与上表第一行相同。在本仓库目录下跑时，Maven 会先按 pom 里声明的 Aspose 仓库找一遍这个构件
（那里没有），随后从 Central 取到。临时仓库用完已删。

GitHub Release [`v1.2.0`](https://github.com/hutianxi/document-parser/releases/tag/v1.2.0) 也已发布（tag 指向
提交 `bfbdec9`，标记为 Latest），三个附件经 GitHub API 查到的 `size` 与 `digest` 和上表逐项相同。

**使用方要知道的一件事**：Aspose 的三个构件不在 Maven Central 上，所以别人引 `document-parser` 时，
必须像 5.1 节写的那样在自己的 pom 里声明 `https://repository.aspose.com/repo/`，否则依赖解析会失败。

1.1.1 发布后已在 GitHub 上建过 Release `v1.1.1`，jar、sources.jar、javadoc.jar 三个附件都在 Assets 里
（GitHub 侧显示的 sha256 与本地产物一致）：[v1.1.1](https://github.com/hutianxi/document-parser/releases/tag/v1.1.1)。
1.2.0 的 Release 也已建好：[v1.2.0](https://github.com/hutianxi/document-parser/releases/tag/v1.2.0)（2026-10-24
00:18 +0800 发布，Latest），三个附件的 `size` 与 `digest` 同上表。

## 18. 附录：修复记录（历史数字）

下面几组数字是修 bug 时用一次性探针测的（探针跑完就删，不在仓库里），**它们对应的构造样例已经被现在这三份
业务样例替换，所以重跑不出同样的数字**，留着是为了说明现在这段代码为什么长这样。

- **文本清洗**（构造文档：段落里插一个 `PAGE` 域、一个软换行，另有一行两格的表、第一格只放域；`<XX>` 是控制字符的十六进制）：
  `前文 <13>PAGE<14><15> 后文<0D>` 修复前 → `前文 PAGE 后文`，修复后 → `前文 后文`；
  `第一行<0B>第二行<0D>` 修复前 → `第一行第二行`，修复后 → `第一行 第二行`；
  `<13>PAGE<14><15><07>普通格<07>` 修复前表头 → `["PAGE","普通格"]`，修复后 → `["列1","普通格"]`；
  评估水印整段修复前原样进 `text`，修复后是空串。
- **纵向合并的数据行被误判成跨整行合并**（两列表，`列A` 跨两行纵向合并、`列B` 每行各有值）：
  `Grid.isSeparator` 旧判据把横向延续格和纵向延续格一起排除，那一行「只剩一个起始格」就被当成分组小标题。
  修复前 `KEEP` 得到 `[{列A:"第一",列B:"值一"},{列A:"第一",列B:""}]`（`值二` 被抹成空串）、`DROP` 只剩一行；
  改后只按横向延续格判定，三种模式都是 `[{列A:"第一",列B:"值一"},{列A:"第一",列B:"值二"}]`。
  当时把新旧两份 `isSeparator` 编到两个目录、同一份探针各跑一遍（5 个样例 × 3 种 `FullRow` 共 15 次），
  组数、行数、`tables` JSON 长度逐项相同，只有那张纵向合并表不同。
- **PDF 静默选项下的跨页误并**（`hasHeaderCandidate` 要注册过表头依据才有意义；两者都没注册时任何行都算不上表头，
  判据里只剩几何条件，两页恰好等列数、左边界对齐的两张独立表格会被并成一张）：按旧 `sample.pdf` 的真实坐标
  （页高 841.9）造两个片段直接调 `continues` —— 上一页那张画到页面中部（下边界 548.25，占页高 0.65）、
  下一页那张顶着页首（上边界 769.65，离页首占页高 0.086），修复前「会并=true」，修复后静默选项「会并=false」；
  同一对位置在注册过表头依据时两版都是「会并=true」，说明新加的几何条件只在没有表头依据时兜底。
  端到端也对比过 `tables` JSON：修复前两张并成 **1 组**（页 2 的表头行被当数据塞进页 1 的表头下），
  修复后 **2 组**、各带自己的表头。
- **跨三页以上的长表只并了头两页**（`mergeSplitTables` 并完一截之后 `previous` 仍指向第一截，
  而 `continues` 要求页码相邻，第三页跟它比就不上了）：用反射按真实坐标造三页片段（每页那一截都从页首画到页底），
  静默选项下修复前剩 **2 截**（第 1+2 页并成 2 行、第 3 页单独 1 行，它的首行还会被兜底当成表头），
  修复后剩 **1 截 3 行**。修法是并完之后把 `previous` 的 `page`、`bottomGap` 换成后一截的
  （左边界不换，始终跟第一截比，免得每页 2pt 的容差累积成漂移）。
- **列名对齐**（入库要拿 `headers` 当列清单，所以它必须逐个等于每一行的键）：表头是空格的列、合并摊平后同名的列，
  过去只有行对象里补了名（`列3`、`合计-2`），`headers` 里仍是 `""` 和两个一模一样的 `合计`；
  改由 `Tables.normalizeHeaders` 按同一条规则补全，构造表头 `合计 | 合计 | 名称` 出来
  `headers=["合计","合计-2","名称"]`，行键一字不差。当时连构造样例一起跑过 41 个表头组、105 行数据，不一致 0 处。
- **`ObjectMapper` 每次新建**：`toJson` 原来每次调用都新建一个 `ObjectMapper`，现在复用一份静态实例
  （`ObjectMapper` 线程安全）。
- **连续空格压成一个的写法**：原来用 `while (contains("  ")) + replace`，一串长空格就要把整个文本重抄十几遍，
  实测 300 万空格的单元格：旧写法 95ms，改成一次正则 10ms。
- **许可证重复注册**：Aspose 的许可证是进程级全局状态，原来每调一次 `parse` 就重读一遍 lic 文件；
  批量解析几百个文档时这些 IO 全是白费的，现在一个产品只注册一次，且注册失败会提示一行。
