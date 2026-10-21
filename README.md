# document-parser

[![License](https://img.shields.io/badge/License-Apache%202.0-blue.svg)](LICENSE)
[![Java](https://img.shields.io/badge/JDK-8%2B-orange.svg)](#3-技术栈与环境要求)
[![Maven Central](https://img.shields.io/badge/Maven%20Central-1.1.1-green.svg)](#17-构建)
[![Tests](https://img.shields.io/badge/tests-183%20passed-brightgreen.svg)](#13-测试与质量)

> 文档解析模块：把 **Word / PDF / PPT** 统一解析成 `text`、`tables`、`images` 三个集合，
> 三种格式返回同一套结构，换格式只换类名。

| 项目信息 | 值 |
| --- | --- |
| Maven 坐标 | `cn.net.tianxi:document-parser:1.1.1` |
| Java 包 | `cn.net.tianxi.document` |
| 语言级别 | Java 8（Aspose 用 `jdk16` classifier 的构件） |
| 底层技术 | Aspose.Words / Aspose.PDF / Aspose.Slides 23.12，API 与行为以 [docs.aspose.com](https://docs.aspose.com) 为准 |
| 产物 | 普通库 jar（`document-parser-1.1.1.jar`，56275 字节、30 个条目、17 个 class，含 `META-INF/LICENSE` 与 `META-INF/NOTICE`），无 Spring 容器、无命令行入口 |
| 测试 | JUnit 5，183 个测试全过（其中 1 条是 `-Ddoc` 入口，平时跳过） |
| 开源许可 | [Apache License 2.0](LICENSE)：可免费使用、修改、商用、再分发，**须保留原作者署名**（见 [NOTICE](NOTICE)），欢迎提 PR（见 [CONTRIBUTING.md](CONTRIBUTING.md)） |
| 代码仓库 | GitHub [`hutianxi/document-parser`](https://github.com/hutianxi/document-parser)｜Gitee [`mailtianxi/document-parser`](https://gitee.com/mailtianxi/document-parser) |

更细的说明分两层：包级设计（返回结构、表头判定优先级、许可证与评估模式、安全边界）写在
[`package-info.java`](src/main/java/cn/net/tianxi/document/package-info.java)，
类与方法的细节写在各自的 javadoc 里，IDE 里悬停即可看到。本 README 是这两层的索引 + 实测数据。

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
| 数字与金额 | 与原文逐字符一致：`合计` 行读到的就是域缓存值 `2096.30`/`1796.10`，183 个测试逐格断言（§11） | `0/O`、`1/l`、`8/B`、小数点这类字形易混，金额通常要人工复核 |
| 表格结构 | 拿到**真结构**：合并单元格标记（`gridSpan`/`vMerge`/`rowSpan`）、两行表头、单元格内嵌套子表都能识别并拆开回填 | 靠框线检测还原；无边框表、跨页断表、合并格容易错行错列 |
| 阅读顺序 | 按文档树顺序（段落 → 表格 → 形状），与作者的写作顺序一致 | 按版面坐标推断，多栏、浮动文本框容易串行 |
| 结果稳定性 | 同一份文件重复解析结果**完全相同**，可以写成断言进 CI | 换引擎版本、换分辨率、换二值化参数结果会变，难做回归 |
| 部署形态 | 一个 jar + 三个 Aspose 依赖，进程内完成，不联网（外链图片只记元信息、不下载） | 一般要额外装识别引擎/模型，或把文档上传到第三方云服务 |
| 扫描件 / 图片型 PDF | **读不出来**（没有文字层），这是 §9 写明的边界 | 唯一可行的路线 |

> **诚实边界**：本节不给 OCR 的准确率或耗时数字。仓库里没有跑过任何 OCR 引擎，编不出可复现的实测值；
> 要横向对比，请拿同一批文件在两边各跑一遍再比。上表右列写的是 OCR 这条路线的**固有特征**，不是某个产品的实测结果。

本项目的单文件耗时是实测的（三份样例，评估模式、无 `.lic`，两轮 `mvn test` 各测一次取区间）：

| 样例 | 大小 | 首次调用（含 Aspose 初始化） | 之后单次解析 |
| --- | --- | --- | --- |
| docx | 675208 字节 | 1460–1591 ms | 23–33 ms |
| pdf（3 页） | 184977 字节 | 3248–3415 ms | 573–706 ms |
| pptx | 679289 字节 | 1138–1344 ms | 69–88 ms |

环境：JDK `1.8.0_191`、Maven 3.8.8、Windows 11。两点说明：

- **首次调用慢是固定开销**，来自每个 Aspose 产品第一次加载类和初始化许可证，同一进程内只发生一次；
  常驻服务（Spring Bean、线程池）里跑的是「之后」那一列。
- **PDF 那 600 ms 上下**花在 `TableAbsorber` 的几何聚类和跨页并表上，与页数、线段数量相关，不是文件体积决定的
  （184977 字节的 pdf 比 679289 字节的 pptx 慢一个数量级）。

### 1.2 这样做的好处

- **上层只写一份逻辑**：三种格式返回同一套 `text`/`tables`/`images` 结构，换格式只换类名；
  加一种新格式只需再写一个解析器，入库代码不用动。
- **`headers` 可以直接当列清单**：每组表的 `headers` 逐个等于每一行的键，合并单元格一律拆开回填，
  所以每行键数量恒定 —— 12 次解析（3 份样例 × 4 种选项）契约 **0** 处不一致
  （`CrossFormatTest.contractHoldsForEverySampleAndMode`）。建表、映射 DTO、写 `INSERT` 都不用特殊处理。
- **可以进 CI 当回归**：结果是确定的，所以 README §11 里的每个数字都是测试断言值，改坏了立刻红。
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
- **PDF 跨页长表并回一张**：几何聚类 + 四条续接判据，跨三页以上也是一路并下去。
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
| `com.fasterxml.jackson.core:jackson-databind` | 2.13.5 | `toJson` 序列化 |
| `spring-boot-starter-test`（test 作用域） | 由父 pom 2.7.18 管理，JUnit 5.8.2 | 单元测试 |

环境要求：

- JDK 8 及以上（Aspose 构件是 `jdk16` classifier，Java 8 可直接用）：本机在 JDK `1.8.0_191` 与 JDK `21.0.1` 下
  各跑一遍完整测试都是 **183 个测试全过、1 条跳过**，两边产出的 class 文件都是 `major version: 52`（Java 8 字节码）——
  `java.version=1.8` 会让 maven-compiler-plugin 走 `-source/-target 1.8`，所以用高版本 JDK 编译**不会**抬高字节码版本；
- Maven 3.8.8（本机实测版本），需要能访问 `https://repository.aspose.com/repo/`；
- 无头服务器加 `-Djava.awt.headless=true`（Aspose 内部会用到 AWT 的图像类）；
- 父 pom 用 `spring-boot-starter-parent` **只为借它的依赖版本管理**，本模块没有 Spring 启动类，也没配 `spring-boot-maven-plugin`。

## 4. 目录结构

```
document-parser/
├── pom.xml                                  # 普通库 jar；父 pom 只借版本管理；release profile 供发 Central 用
├── README.md
├── LICENSE                                  # Apache License 2.0 全文
├── NOTICE                                   # 原作者署名 + 第三方组件（Aspose 商业授权）说明
├── CONTRIBUTING.md                          # 贡献指南：环境、约定、PR 流程、许可红线
├── .gitignore                               # target/、*.lic（私有凭证）、GPG 签名产物、IDEA 工程文件
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
        ├── java/cn/net/tianxi/document       # 10 个测试类、183 个测试（见「测试与质量」）
        │   ├── TestSupport.java             # 样例路径、断言助手、现场拼最小 docx 的零件
        │   ├── WordParserTest.java / PdfParserTest.java / PptParserTest.java   # 三份样例的端到端断言
        │   ├── CrossFormatTest.java         # 三种格式横向一致性
        │   ├── TablesTest.java / ImagesTest.java / OptionsTest.java / ResultsTest.java  # 共用组件的单元测试
        │   ├── WordEdgeCasesTest.java       # 现场拼最小 docx 覆盖边角结构
        │   └── PrintResultTest.java         # 不断言、只把解析结果打到控制台看（含 -Ddoc 入口）
        └── resources/{word,pdf,ppt}          # 三份业务样例（不打进产物 jar）
```

`Tables` / `Images` / `Results` 都是**包内可见**，解析器之外不暴露；对外只有三个解析器类。

## 5. 快速开始

### 5.1 引入依赖

```xml
<dependency>
    <groupId>cn.net.tianxi</groupId>
    <artifactId>document-parser</artifactId>
    <version>1.1.1</version>
</dependency>
```

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
        .headers("客户名称", "合同金额", "已回款", "回款率", "备注")
        .headerTolerance(1)                          // 允许 1 列不一致（0 = 逐列相等），默认 1
        .fullRow(WordParser.FullRow.DROP)            // 跨整行合并的分组行直接丢弃
        .keywords("销售");                           // 可选，宽松兜底

Map<String, Object> result = PdfParser.parse("D:/a.pdf", options);
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
mvn -B test -Dtest=PrintResultTest                                              # 三份样例各打一份
mvn -B test -Dtest=PrintResultTest#printAnyDocument -Ddoc="D:/你的文档.docx"      # 看你自己的文档
```

控制台会给出正文逐段、每组表的表头与全部数据行、图片元信息和完整 JSON，同时写一份 UTF-8 副本到
`target/print/`。输出样例与细节见 [13. 测试与质量](#13-测试与质量)。

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

`WordParser.Options` 的链式方法：`headers(String...)`、`headerGroups(List<List<String>>)`、
`headerTolerance(int)`、`keywords(String...)`、`fullRow(FullRow)`、`images(ImageMode)`、`imageDir(File)`。
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
                  "销售额-5": "1465", "回款率-6": "95%", "备注": "超额完成" } ] }
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
- **Word 单元格内嵌套表**：子表不进母表格子的文本（母表那一格留空），而是作为**额外元素**排在 `tables` 里，
  两边的行列结构都不丢。样例实测母表 `["工单类型","工单数","处理说明"]`（那一格是 `["硬件故障","42",""]`），
  子表另成一组 `["故障部件","工单数"]` 3 行。

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

体积要心里有数（三份样例实测，见「实测基线」）：那份 docx 只有一张图，META 模式整串结果 4409 个字符，
BASE64 模式 899039 个字符 —— 那张 670939 字节的 PNG 编码后就是 894610 个字符，占了 99.5%。

## 8. 选项详解

`WordParser.Options` 是三种格式共用的可变对象，链式调用即改即生效。

| 选项 | 默认 | 含义与实测效果 |
| --- | --- | --- |
| `headers(String...)` | 不注册 | 注册一组完整表头（可多次调用注册多组）。命中后列名用注册的原名。样例实测：注册 5 组真表头后，docx 与 pdf 的 JSON 与静默模式**逐字节相同**，pptx 从 22 行变 21 行（下层表头行不再被当成数据）。比对时会把文档那一行**末尾的空列剪掉**，所以末列没有标题的表（docx 目标表）注册前 3 个列名就能命中、第 4 列仍补成 `列4`；注册组自己不要以空列名结尾，否则永远命不中 |
| `headerGroups(List<List<String>>)` | 空 | 批量注册，与多次 `headers(...)` 等价 |
| `headerTolerance(int)` | 1 | 表头比对允许几列不一致；0 = 逐列相等。无论取值多少，都至少要求一列完全相等且列数相同；负数会被夹到 0 |
| `keywords(String...)` | 不给 | 宽松兜底：**行内每个非空单元格都含关键字之一**才算表头行。null 与空白项被忽略 |
| `fullRow(FullRow)` | `KEEP` | 跨整行合并行（分组小标题）的去向：`KEEP` 归到它上面最近的表头组、`FIRST` 统一归到该物理表的第一组、`DROP` 丢弃。样例实测 docx：KEEP 7 组／30 行、FIRST 同数量只改归属、DROP 6 组／28 行 |
| `images(ImageMode)` | `META` | 见 7.3 |
| `imageDir(File)` | null | 只在 `FILE` 模式下生效；`FILE` 没配它会在解析到第一张图时抛 `IllegalStateException`，配了却不是 `FILE` 会在标准错误提示 |

**静默模式**（`isSilent()`：既没注册表头、也没给关键字）会打开「第一行兜底」，适合先跑一遍看看文档长什么样；
要稳定入库建议注册表头。选项本身没做同步，配置期间别跨线程共享；配好之后可以并发用于多次解析
（`Images` 每次解析新建，选项只读），`imageDir` 指向同一目录时重名图片由 `Images.uniqueFile` 加序号，不会互相覆盖。

## 9. 能力边界：读到什么、读不到什么

| 格式 | 读进结果 | 读不到（已知范围外） |
| --- | --- | --- |
| Word | 各节 Body 的段落、表格、图片；文本框文字；块级内容控件（`w:sdt`）包住的段落和表格；列表项（项目符号不进文本）；域**缓存的结果**（`SUM(ABOVE)`、`PAGE` 都读缓存值） | 页眉页脚、脚注尾注、批注；域代码一律丢弃、不做版面计算，所以没算过的域是空的；没填写的控件不会给出占位文字 `Click here to enter text.`（4 种写法实测，见下） |
| PDF | 有线段边框的表格（几何还原）、文字层段落、位图；跨页断开的长表并回一张（判据：相邻页 + 列数相同 + 左边界对齐 + 后一截找不出表头行；没注册表头依据时最后这条换成「一截贴页底、一截贴页首」的几何证据）；纵向合并处被吸收器切成几块的行按矩形相邻聚回同一张表 | **没有边框线的表格**（只有底色或空白对齐）会被当成普通段落进 `text`；扫描件／图片型 PDF 需要 OCR；没注册表头依据、又被分页切断的那一行特别高（下边界离页底超过页高的 1/5）时认不出续接，一张表拆成两组、各带一个表头 |
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
JDK 用 `1.8.0_191` 与 `21.0.1` 各跑了一遍完整测试，两边都是 183 个测试全过、1 条跳过。
下面的数字全部是当前测试里的断言值，`mvn -B clean test` 就能复现（整轮耗时 JDK 8 实测 30.8 秒、JDK 21 实测 35.2 秒；
加不加 `-DargLine="-Dfile.encoding=UTF-8"` 都是 183/183，那一行只影响控制台与报告里的中文编码）。
单份文件的解析耗时见 [§1.1](#11-与-ocr-的区别读文档里的文字层不猜像素)。

| 样例 | `text` | 表头组 | 数据行 | 图片 | META 结果长度 | BASE64 结果长度 |
| --- | --- | --- | --- | --- | --- | --- |
| docx | 13 段 | 7 | 30 | 1 张 PNG 1024×768、670939 字节 | 4409 字符 | 899039 字符 |
| pdf | 空串（正文全在表格里） | 6 | 45 | 1 张 JPEG 699×524、38668 字节 | 6186 字符 | 57789 字符 |
| pptx | 9 行（每行 ≤5 字符） | 6 | 22 | 1 张 PNG 1024×768、670939 字节 | 3262 字符 | 897892 字符 |

同一份样例换选项后的组数／行数（也都是断言值）：

| 选项 | docx | pdf | pptx |
| --- | --- | --- | --- |
| 静默（默认，`FullRow.KEEP`） | 7 组 / 30 行 | 6 组 / 45 行 | 6 组 / 22 行 |
| 注册 5 组真表头 | 7 组 / 30 行，JSON 与静默逐字节相同 | 6 组 / 45 行，JSON 与静默逐字节相同 | 6 组 / **21** 行（下层表头行不再当数据） |
| 注册表头 + `DROP` | 6 组 / 28 行 | 5 组 / 43 行 | 5 组 / 19 行 |
| `FIRST`（不注册） | 7 组 / 30 行，只是分组行挪到该表第一组 | 6 组 / 45 行，同 | 6 组 / 22 行，同 |

三份样例的核心表（季度销售明细）在三种格式、静默与注册两种模式下**逐格相同**：
列名 `["片区","区域","销售额","回款率","销售额-5","回款率-6","备注"]`，4 行数据一字不差
（`CrossFormatTest.coreTableIsIdenticalAcrossThreeFormats`）。台账行数三份不同（docx 13 / pdf 31 / pptx 8）
是造样例时按各格式的可读边界裁的，不是解析丢了行。

### 样例覆盖到的行为（每条都有测试守着）

| 行为 | 实测结果 | 测试 |
| --- | --- | --- |
| 同名列补列号后缀 | `销售额`/`回款率` 第二次出现 → `销售额-5`/`回款率-6`，`headers` 与行键逐个相等 | `WordParserTest.sameNamedColumnsGetColumnNumberSuffix` |
| 纵向合并回填 | `片区` 列的江苏／安徽两行自己没有内容，读出 `华东大区`/`华东南区`；4 行数据都在，没被当成分组行 | `WordParserTest.verticallyMergedColumnIsFilledFromItsOriginRow`、`WordEdgeCasesTest.verticallyMergedDataRowIsNotAMistakenSeparator` |
| 两行表头合并 | 下层空着的列沿用上层名字 → `["客户名称","合同金额","已回款","回款率","备注"]` | `WordParserTest.twoLevelHeaderInheritsUpperLevelColumnNames`、`TablesTest.consecutiveHeaderRowsMergeIntoOneMultiLevelHeader` |
| 域结果 | `合计` 行读的是 Word 缓存的 `SUM(ABOVE)` 结果 `2096.30`/`1796.10`/`86%`；正文里 `PAGE` 域的缓存值 `3` 留着、域代码丢掉；没算过的 `DOCPROPERTY` 整块消失，只剩 `归档编号：` | `WordParserTest.sumFieldResultIsReadFromTheCachedValue`、`WordEdgeCasesTest.calculatedFieldResultSurvivesInsideACell` |
| 软换行 | Shift+Enter 换成空格：`…含税销售额。 数据来源：…`（最长那段 62 个字符完整读出，评估模式没截 Words 的正文） | `WordParserTest.textJoinsSoftBreakLinesWithASpace` |
| 块级内容控件 | 控件里包的段落进 `text`（`编制部门：销售管理部 审核：财务总监`）；套在 `w:tc` 里的控件，`Grid.readRow` 的浅层段落查找也看穿得到（`["控件里的甲","乙"]`） | `WordParserTest.textReadsParagraphInsideBlockContentControl`、`WordEdgeCasesTest.paragraphInsideCellLevelContentControlIsRead` |
| 没填写的控件不给占位文字 | 4 种写法实测：`text` 都只有控件外那一段，`Click here to enter text.` 一次都没出现 | `WordEdgeCasesTest.unfilledContentControlContributesNothing` |
| 单元格内嵌套表 | `tables` 1 → 2 组：母表那一格是空的（`["硬件故障","42",""]`），子表另成一组 `["故障部件","工单数"]` 3 行 | `WordParserTest.nestedTableBecomesItsOwnGroupAndLeavesTheHostCellClean` |
| 分组行三种去向 | KEEP 留在自己那组、FIRST 归到该表第一组（组数行数不变，只改归属）、DROP 整行丢掉（docx 7→6 组、30→28 行）；分组行**后面紧跟的那一行按表头处理** | `TablesTest` 的 `separatorRow*`／`fullRow*`／`rowAfterASeparatorRowIsTreatedAsAHeader` 六条 + 三个解析器各一条 |
| 表头之前就碰到分组行 | 另起一组、列名补成 `列1..列4`，分组文字不丢（docx 组 4、pdf 组 3、pptx 组 2 都是这种） | `TablesTest.separatorRowBeforeAnyHeaderStartsANumberedGroup` + 三个解析器各一条 |
| PDF 跨页长表 | 31 行台账并回一组：第 1 页画到 `泰州医药城`（第 20 行），第 2 页从 `湖州建材贸易`（第 21 行）接着画，`合计` 在最后 | `PdfParserTest.ledgerSplitAcrossTwoPagesIsMergedBackIntoOneGroup`、`CrossFormatTest.onlyPdfNeedsTheCrossPageMerge` |
| PDF 纵向合并被切块 | 核心表的江苏／安徽两行、工单表的「硬件故障 12」那行都被吸收器切成独立小块，聚回来后仍是 4 行 | `PdfParserTest.coreTableRowsCutByVerticalMergeAreClusteredBack`、`PdfParserTest.ticketTableKeepsTheRowThatWasCutIntoItsOwnBlock` |
| 注册表头关掉首行兜底 | PDF 没有 `tblHeader` 标记可依赖：静默时首行当表头，注册了表头依据之后没命中的表首行按数据行处理、列名补成 `列N`（核心表 5 行 = 表头行 + 4 行数据） | `PdfParserTest.registeringHeaderEvidenceTurnsOffTheFirstRowFallback`、`TablesTest.firstRowStaysADataRowOnceHeaderEvidenceIsRegistered` |
| 注册名覆盖文档里读到的列名 | 默认容忍 1 列不一致：把 `备注` 注册成 `说明` 仍命中，结果里的列名用注册的 `说明`（PDF 的单元格文字会因换行多出空格，比对时忽略空格） | `WordParserTest.registeredHeadersOverrideTheNamesReadFromTheDocument`、`TablesTest.registeredHeadersOverrideTheNamesReadFromTheGrid` |
| 未匹配的注册表头会提示 | 标准错误上出现 `表头集合未匹配到任何行` + 那组列名；PDF／PPT 上提示的是 `故障部件`（嵌套小表在这两种格式里不存在），Word 上提示的是 `客户名称`（两行表头摊平后列数与注册的一行 5 列不同） | 三个解析器各一条 `unmatchedHeaderGroupIsReportedOnStderr` |
| PPT 两行表头的取舍 | 静默 8 行、列名 `["客户名称","回款情况","回款情况-3","回款率","备注"]`；注册表头后 7 行 —— 这就是「PPT 建议注册表头」的理由 | `PptParserTest.silentModeCannotMergeTheTwoLevelHeader`、`registeredHeadersMergeTheTwoLevelHeader` |
| PPT 组合形状 | `数据来源／销售系统／财务对账` 三个文本框在一个组合形状里，递归展开才读得到 | `PptParserTest.recursesIntoGroupShapes` |
| 图片四种模式 | META 只有元信息；BASE64 给 data URI（解码回来 670939 字节，与原图一致）；BYTES 给原始字节；FILE 落盘并给绝对路径、目录里只多一个文件；FILE 没配 `imageDir` 抛 `IllegalStateException`；META 配了 `imageDir` 在标准错误提示且一个文件也不写 | `WordParserTest` 的图片那 7 条 + `ImagesTest` 24 条 |
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

`mvn -B clean test`（`JAVA_HOME` 指向 JDK 8）跑完 **183 个测试、10 个测试类**（0 失败、1 跳过 ——
跳过的那条是 `-Ddoc` 入口，见下面「直接看解析结果」）：

| 测试类 | 个数 | 覆盖 |
| --- | --- | --- |
| `WordParserTest` | 31 | docx 样例端到端：正文 13 段、4 张表逐格、`FullRow` 三种去向、注册表头、图片四种模式、两个入口 |
| `PdfParserTest` | 17 | pdf 样例端到端：跨页并表、纵向合并聚块、丢域结果、注册表头关掉首行兜底、JPEG 重编 |
| `PptParserTest` | 17 | pptx 样例端到端：形状顺序、组合形状、两行表头的取舍、评估模式 5 字符自检 |
| `CrossFormatTest` | 9 | 三种格式横向一致性：核心表逐格相同、图片指纹、一个 `Options` 复用、契约 12 次全查 |
| `TablesTest` | 31 | 表头分组算法（唯一一份）：六条判定优先级、多级表头合并、分组行三种去向、列名补全、网格补齐 |
| `ImagesTest` | 24 | 图片集合（唯一一份）：编号、类型归一化与嗅探、指纹、重复图片、四种模式、外链、文件名消毒 |
| `OptionsTest` | 10 | `Options` 的判定语义：容忍度、关键字、静默模式、空值处理 |
| `ResultsTest` | 28 | 文本清洗、评估水印识别、表头归一化、JSON 序列化、文件名消毒、`writeImages`、许可证加载 |
| `WordEdgeCasesTest` | 11 | 现场拼最小 docx 覆盖边角结构：内容控件、嵌套表、深度上限、纵向合并、域、软换行 |
| `PrintResultTest` | 5（1 条默认跳过） | **不断言、只打印**：把三份样例的解析结果原样输出到控制台，看「这个文档到底解析成了什么样」 |

`TablesTest`／`ImagesTest`／`OptionsTest`／`ResultsTest` 不经过 Aspose（直接喂网格、字节和最小字符串），
实测各 0.05 秒以内（31/24/10/28 条分别 0.013、0.043、0.000、0.035 秒）；端到端那几类要加载 Aspose 与样例，
最重的是 `CrossFormatTest`（19.6 秒，三种格式反复解析）和 `PdfParserTest`（10.2 秒），整轮 `mvn clean test` 38 秒。

本轮构建核对：`javac -Xlint:all` **0** 告警，`javadoc -Xdoclint:all,-missing` **0** 错误，
`mvn -B clean package` BUILD SUCCESS、183 个测试 0 失败 0 错误（加不加 `-DargLine="-Dfile.encoding=UTF-8"` 都是 183/183）。

### 直接看解析结果（打印到控制台）

其余测试类都在断言实测值，能挡回归但看不出输出长什么样。`PrintResultTest` 就是给人看的：
跑一条测试，控制台直接给出**正文逐段、每组表的表头与全部数据行、图片元信息、完整 JSON**，
只保留一条底线断言（`headers` 与每行键的契约），因为契约破了输出也没参考价值。

```
mvn -B test -Dtest=PrintResultTest                                    # 三份样例 + 一组选项对比，4 条各打一份
mvn -B test -Dtest=PrintResultTest#printWordSample                    # 只看 docx
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

#1  7 列 × 4 行
headers = [片区, 区域, 销售额, 回款率, 销售额-5, 回款率-6, 备注]
  行1: 片区=华东大区 | 区域=上海 | 销售额=1280 | 回款率=92% | 销售额-5=1465 | 回款率-6=95% | 备注=超额完成
…
---- images（1 张，META 模式只有元信息）----
[1] {fileName=image-1.png, type=PNG, width=1024, height=768, size=670939, contentId=487e915d…}

---- JSON（4409 个字符）----
{ … }
```

每条测试还会把同一份输出写一份 **UTF-8 副本**到 `target/print/<样例名>-<选项>.txt`（末尾打印绝对路径），
控制台编码不合适时直接开这个文件；它在 `target/` 下，`mvn clean` 会带走。本轮实测四份副本的字符数：
docx 静默 7055、docx 注册表头+DROP 6842、pdf 静默 9460、pptx 静默 5215。

**不打 BASE64 模式**：默认 `META` 下三份样例的 JSON 是 4409／6186／3262 个字符，读得完；
换 `BASE64` 那张 670939 字节的 PNG 会把 JSON 撑到 899039 个字符，刷屏且看不出东西。
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
结果里第 4 列仍按位置补成 `列4`。命不中时标准错误上会提示 `表头集合未匹配到任何行` + 那组列名，照着改注册值即可。

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
| `Tables` | 表头分组算法（唯一一份）：判定优先级、多级表头合并、跨整行合并行的去向、列名补全（空表头→`列N`、同名→加列号后缀，`headers` 与行键同源） |
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
（三种格式的键因此相同）；表格里的文字不在 `text` 里重复出现；图片顺序 = 文档流顺序。

**改样例之后**：所有期望值都是实测出来的，改了样例就把对应测试类跑一遍、按新输出改断言，
再把「实测基线」两张表里的数字同步过来；javadoc 里引用的样例数字（三个解析器类说明、`Results`、`Images`、
`package-info`）也要一起改。

## 17. 构建

```
mvn -B clean test          # JAVA_HOME 指向 JDK 8；本轮实测 183 个测试全过（1 条 -Ddoc 入口跳过）
mvn -B clean package       # 产物 target/document-parser-1.1.1.jar（56275 字节、30 个条目、17 个 class，
                           # 类直接放在 jar 根目录的 cn/net/tianxi/document/ 下；
                           # prepare-package 阶段会把根目录的 LICENSE 与 NOTICE 复制进 META-INF/，
                           # 所以只拿到 jar 的使用方也看得到署名与第三方授权说明）
```

javadoc 的 doclint 核对用 JDK 自带的 `javadoc` 直接跑（本机离线仓库里没有 `maven-javadoc-plugin`，
`mvn javadoc:javadoc` 实测报 `No plugin found for prefix 'javadoc'`）：

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

发布有两条路，本项目 1.1.1 走的是第一条（**不需要把 Token 落到磁盘上**）：

**A. 本地打 bundle + Portal 网页上传**

```
mvn -B -P release deploy -DskipPublishing=true      # 版本号不能带 -SNAPSHOT
```

产出 `target/central-publishing/central-bundle.zip`（实测 266857 字节、24 个条目）：
`cn/net/tianxi/document-parser/1.1.1/` 下的 `document-parser-1.1.1.jar`（56275 字节）、`-sources.jar`（64964 字节）、
`-javadoc.jar`（141571 字节）、`.pom`，各带一个 `.asc` 签名和 `md5/sha1/sha256/sha512` 四个校验和。

然后到 Portal → **Publishing → Publish Component**：填 Deployment Name（如 `document-parser 1.1.1 (cn.net.tianxi)`）
与描述，Choose File 选这个 zip，点 Publish Component。提交后 Deployments 列表里出现一条 **VALIDATING** 记录
（本次的 Deployment ID 是 `960f618b-cb17-4cb1-9d40-96849df2a219`），校验通过后同一行会出现 **Publish** 按钮，
点它才真正公开。

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
https://repo1.maven.org/maven2/cn/net/tianxi/document-parser/1.1.1/
```

1.1.1 已经走完全流程并**实测验证过**：Portal Deployment `960f618b-cb17-4cb1-9d40-96849df2a219` 状态 PUBLISHED，
上面这个目录列出的 jar / sources.jar / javadoc.jar 字节数与本机产物一致（56275 / 64964 / 141571）；
再用一个**全新的空本地仓库**跑 `mvn dependency:get -Dartifact=cn.net.tianxi:document-parser:1.1.1 -Dtransitive=false`
（`-Dmaven.repo.local` 指到临时目录，避免命中本机已装的同名构件），BUILD SUCCESS，下到的 jar 是
`https://repo.maven.apache.org/maven2/cn/net/tianxi/document-parser/1.1.1/document-parser-1.1.1.jar`，
sha256 `c2c47ba92823ee9464a9a80bb6125ee8b8aa34f562ee80d76f709df3ce233148`，与 Central 上的 `.sha256`、
本地产物、GitHub Release 附件三处完全相同。临时仓库用完已删。

**使用方要知道的一件事**：Aspose 的三个构件不在 Maven Central 上，所以别人引 `document-parser` 时，
必须像 5.1 节写的那样在自己的 pom 里声明 `https://repository.aspose.com/repo/`，否则依赖解析会失败。

发完 Central 之后，顺手在 GitHub 上建 Release `v1.1.1`（tag 已经推上去了），把 jar、sources.jar、javadoc.jar
附上，方便不用 Maven 的人直接下载：[v1.1.1](https://github.com/hutianxi/document-parser/releases/tag/v1.1.1)
已发布，三个附件都在 Assets 里（GitHub 侧显示的 sha256 与本地产物一致）。

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
  四条判据只剩几何，两页恰好等列数、左边界对齐的两张独立表格会被并成一张）：按旧 `sample.pdf` 的真实坐标
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
