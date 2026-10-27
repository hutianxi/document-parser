# document-parser 功能说明书：接口、参数、返回字段与样例数据

> 适用版本：**1.2.1**（`cn.net.tianxi:document-parser:1.2.1`，2026-10-24 已发布到 Maven Central）
>
> 本文按“接口 → 参数 → 返回字段 → 真实数据”的顺序写，目标是不看源码也能把每个入口调对、
> 知道每个参数传什么、知道结果里每个键是什么意思。所有数字都来自本仓库
> `PrintResultTest` 在 2026-10-26 生成的 `target/print/` 报告与 242 条测试
> （15 个测试类，241 通过、1 跳过），不是估算值；关键结论在对应小节后标了报告文件，可自行复现。
>
> 与其它文档的分工：README 讲“是什么、怎么快速用”，DESIGN.md 讲“为什么这样实现”，
> 本文讲“每个接口每个参数具体怎么用、结果里每个字段长什么样”。

## 目录

- [1. 读法与功能地图](#1-读法与功能地图)
- [2. 公共 API 总表](#2-公共-api-总表)
- [3. 解析入口 `parse`：四个重载逐参数](#3-解析入口-parse四个重载逐参数)
- [4. `options()` 与 `Options` 默认值](#4-options-与-options-默认值)
- [5. 配置与取值方法逐个说明](#5-配置与取值方法逐个说明)
- [6. 两个枚举：`FullRow` 与 `ImageMode`](#6-两个枚举fullrow-与-imagemode)
- [7. 返回数据字典](#7-返回数据字典)
- [8. 三种格式分别能解析出什么](#8-三种格式分别能解析出什么)
- [9. 三份样例的完整数据](#9-三份样例的完整数据)
- [10. 场景配方（可直接编译运行）](#10-场景配方可直接编译运行)
- [11. 异常与标准错误](#11-异常与标准错误)
- [12. 自己跑一遍验收](#12-自己跑一遍验收)
- [13. 已知限制与不承诺的行为](#13-已知限制与不承诺的行为)
- [14. 速查卡](#14-速查卡)

## 1. 读法与功能地图

### 1.1 一句话

把 `.docx` / `.pdf` / `.pptx` 里**已经存在的文字层**读成一份结构化 Java 结果：
`text`（正文）+ `tables`（表头和行）+ `images`（图片元信息或内容），不猜像素、不做 OCR。

### 1.2 能力地图

| 能力 | 对外接口 | 结果里的位置 | 详见 |
| --- | --- | --- | --- |
| 三种格式解析（路径/流 × 关键字/选项） | 三个入口类的 `parse` | 整个返回值 | 第 3 节 |
| 正文抽取与清洗 | 自动 | `text` | 7.1、8.1 |
| 表格分组、合并单元格、跨页并表、PDF 无边框补读 | 自动 | `tables` | 7.2、8.2、8.4、8.5 |
| 表头注册、容忍度、关键字 | `headers` / `headerGroups` / `headerTolerance` / `keywords` | 影响 `tables` | 5.1–5.5 |
| 起名与按名取数 | `name` / `tablesByName` / `rowsByName` | 组里的 `name` | 5.3、5.13、5.14、7.2 |
| 来源页码/幻灯片号 | `source(boolean)` | 组里的 `source` | 5.9、7.4 |
| 图片四模式、指纹、去重、落盘 | `images` / `imageDir` / `writeImages` | `images` | 5.7、5.8、5.12、6.2、7.5 |
| 未匹配表头清单 | `unmatchedHeaders(boolean)` | 顶层 `unmatchedHeaders` | 5.10、7.6 |
| 跨整行合并行的去向 | `fullRow` | 影响分组与行归属 | 6.1 |
| JSON 序列化 | `toJson` | 字符串 | 5.11、7.1 |
| 异常与提示 | 调用方 `catch` / 标准错误 | — | 第 11 节 |

**检查结论**：上表每一项都有本次跑出来的报告或测试证据；并发说明（4.3）是使用约定，
本次没有做压力测试，因此不声称“已验证高并发”。

### 1.3 数字口径

- 字符数一律指 `toJson` 返回的 Java `String.length()`（UTF-16 代码单元）；本文数字在 Windows 上测得，
  换行是平台默认（`\r\n`），同一份结果在 Linux（`\n`）上会少掉每个 `\r`：docx 4740 → 4521、pdf 6473 → 6161、
  pptx 3526 → 3353。不是 UTF-8 字节数；样例文字都在 BMP 内，字符数与 `wc -m` 的计数一致；
- “组 / 行”指 `tables` 里的元素个数与所有元素 `rows` 的行数合计；
- 三份样例是同一份业务报告按格式分别制作的，**台账行数 13 / 31 / 7 是样例本身裁剪不同**，
  不是解析器丢行；表头合并方式与合计行才是真正的格式差异。

## 2. 公共 API 总表

三个入口类的包名都是 `cn.net.tianxi.document`。`Options` 只有一份，定义在 `WordParser` 里，
三种格式共用 `WordParser.Options`。

| 方法 | `WordParser` | `PdfParser` | `PptParser` | 说明 |
| --- | --- | --- | --- | --- |
| `parse(String, String...)` | ✓ | ✓ | ✓ | 路径 + 表头关键字 |
| `parse(InputStream, String...)` | ✓ | ✓ | ✓ | 流 + 表头关键字 |
| `parse(String, Options)` | ✓ | ✓ | ✓ | 路径 + 完整选项 |
| `parse(InputStream, Options)` | ✓ | ✓ | ✓ | 流 + 完整选项 |
| `options()` | ✓ | ✗ | ✗ | 新建选项；PDF/PPT 用 `WordParser.options()` |
| `toJson(Object)` | ✓ | ✓ | ✓ | 序列化成缩进 JSON |
| `writeImages(Map, File)` | ✓ | ✓ | ✓ | 把 BASE64 结果里的图片落盘 |
| `tablesByName(Map, String)` | ✓ | ✗ | ✗ | 按名字取组；三种格式的结果都能传进来 |
| `rowsByName(Map, String)` | ✓ | ✗ | ✗ | 按名字取行；同上 |

`PdfParser` / `PptParser` 没有 `options()`、`tablesByName`、`rowsByName`，这不是功能缺失：
选项类型和结果结构是共用的，直接调用 `WordParser.options()` 与 `WordParser.tablesByName(...)` 即可。

逐参数说明：`parse` 的四个重载见 3.1–3.4，`options()` 与十个配置方法见 4–5.10，
`toJson` / `writeImages` / `tablesByName` / `rowsByName` 见 5.11–5.14。

枚举也定义在 `WordParser` 里：

```java
WordParser.FullRow.KEEP / FIRST / DROP
WordParser.ImageMode.META / BYTES / BASE64 / FILE
```

## 3. 解析入口 `parse`：四个重载逐参数

### 3.1 四个签名

```java
public static Map<String, Object> parse(String filePath, String... headerKeywords) throws Exception
public static Map<String, Object> parse(InputStream inputStream, String... headerKeywords) throws Exception
public static Map<String, Object> parse(String filePath, WordParser.Options options) throws Exception
public static Map<String, Object> parse(InputStream inputStream, WordParser.Options options) throws Exception
```

三种格式的签名一一对应：`Options` 是 `WordParser.Options`，三个入口类共用同一份类型。

```java
WordParser.parse(...);
PdfParser.parse(...);
PptParser.parse(...);
```

### 3.2 参数 `filePath`

| 项 | 内容 |
| --- | --- |
| 类型 | `String` |
| 必填 | 是 |
| 传什么 | `.doc` / `.docx`（WordParser）、`.pdf`（PdfParser）、`.ppt` / `.pptx`（PptParser）的文件路径 |
| 路径规则 | 绝对路径直接用；相对路径按**进程当前工作目录**解析（仓库测试都从仓库根目录运行） |
| 返回值 | 不是文件内容，而是第 7 节那份 `Map` |
| 异常 | 文件不存在时三种格式抛出的类型不同，见 11.1 |

示例：

```java
Map<String, Object> result = WordParser.parse("D:/docs/销售报告.docx");
Map<String, Object> same = WordParser.parse("src/test/resources/word/sales-report.docx");
```

### 3.3 参数 `inputStream`

| 项 | 内容 |
| --- | --- |
| 类型 | `java.io.InputStream` |
| 必填 | 是 |
| 传什么 | 任意可读流：`FileInputStream`、`ByteArrayInputStream`、上传接口拿到的流等 |
| 谁关闭 | **调用方**。三个解析器都不关闭传入的流；PDF 只关闭自己内部的 `Document` |
| 位置要求 | 从流的当前位置开始读，通常应是流的开头 |
| 结果 | 与同一份文件用路径入参**逐字符相同**（三份样例实测都是 `true`） |

示例：

```java
try (InputStream in = new FileInputStream("销售报告.docx")) {
    Map<String, Object> result = WordParser.parse(in);
    System.out.println(result.get("tables"));
}
```

### 3.4 参数 `headerKeywords`（可变参数）

| 项 | 内容 |
| --- | --- |
| 类型 | `String...`，可传 0 个、1 个或多个，也可直接传 `String[]` |
| 必填 | 否 |
| 默认 | 不传 = 静默模式，允许“首行兜底当表头” |
| 判定规则 | 一行里**每个非空单元格**都包含任意一个关键字，这一行才算表头 |
| 空值 | `null` 与纯空白项被忽略 |
| 副作用 | 一旦传了关键字，就关闭“静默模式的首行兜底” |
| 注意 | 这是宽松匹配，不是精确匹配；要稳定入库请用 `headers(...)` 注册完整表头 |

示例：

```java
// 0 个关键字：静默模式
Map<String, Object> a = WordParser.parse("报告.docx");

// 2 个关键字：每个非空表头格都要含“销售”或“回款”
Map<String, Object> b = WordParser.parse("报告.docx", "销售", "回款");

// 传数组与逐个传等价
String[] keys = {"销售", "回款"};
Map<String, Object> c = WordParser.parse("报告.docx", keys);
```

为什么样例里传 5 个关键字，pdf 反而多出 2 行：传关键字会关闭“首行兜底”。pdf 的台账和售后工单
两张表，首行不满足“每个非空格都含关键字”，表头行就落进了数据，列名也退化成 `列1…列N`
（实测 45 → 47 行，JSON 6473 → 6445 字符）。所以关键字适合“先探路”，不适合当入库契约。

### 3.5 参数 `options`

| 项 | 内容 |
| --- | --- |
| 类型 | `WordParser.Options` |
| 必填 | 是（该重载没有默认值） |
| 怎么来 | 用 `WordParser.options()` 新建，再链式配置 |
| 可否复用 | 可以。配好之后多次解析都读同一份选项，解析过程不修改它 |
| 可否为 `null` | 没有约定，按非 `null` 使用 |
| 与关键字的关系 | `options` 里可以同时有 `headers(...)` 与 `keywords(...)`；注册表头优先判定 |

示例：

```java
WordParser.Options options = WordParser.options()
        .headers("客户名称", "合同金额", "已回款", "回款率", "备注")
        .name("回款台账")
        .images(WordParser.ImageMode.META);

Map<String, Object> docx = WordParser.parse("报告.docx", options);
Map<String, Object> pdf = PdfParser.parse("报告.pdf", options);
Map<String, Object> ppt = PptParser.parse("报告.pptx", options);
```

### 3.6 返回值与异常

返回值固定是 `Map<String, Object>`，键的顺序就是插入顺序：

```text
text → tables → images → [unmatchedHeaders]
```

- `text`：`String`，永远存在，可能是空串；
- `tables`：`List<Map<String, Object>>`，永远存在，没有表时为空列表；
- `images`：`List<Map<String, Object>>`，永远存在，没有图时为空列表；
- `unmatchedHeaders`：只有 `unmatchedHeaders(true)` 时才存在。

四个 `parse` 都声明 `throws Exception`。解析失败时异常原样抛出，不做包装；
需要区分“文件不存在 / 选项错误 / 格式不匹配”时看 11.1 与 11.2。

### 3.7 路径、流、关键字、选项的实测对比

同一份文件、同一份选项，路径与流的 JSON 逐字符相同。三份样例实测：

| 格式 | 路径 + Options | 流 + Options | 路径 + 5 关键字 | 流 + 5 关键字 | 零关键字 |
| --- | --- | --- | --- | --- | --- |
| docx | 7 组 / 30 行，4828 字符 | 4828 | 7 组 / 30 行，4740 | 4740 | 4740 |
| pdf | 6 组 / 45 行，6561 字符 | 6561 | 6 组 / 47 行，6445 | 6445 | 6473 |
| pptx | 6 组 / 21 行，3466 字符 | 3466 | 6 组 / 22 行，3526 | 3526 | 3526 |

表中的 Options 是“注册 5 组表头、其中 3 组起名”的那一份（5.3）。只注册不起名时，
docx 与 pdf 的 JSON 与静默结果逐字节相同；pptx 仍是 21 行（注册后下层表头行被正确合并）。

> 证据：`target/print/公共入口-四个重载.txt`、`选项矩阵-三份样例.txt`。

### 3.8 `PdfParser` / `PptParser` 的差异

| 项 | docx | pdf | pptx |
| --- | --- | --- | --- |
| 输入扩展名 | `.doc` / `.docx` | `.pdf` | `.ppt` / `.pptx` |
| 内部引擎 | Aspose.Words | Aspose.PDF | Aspose.Slides |
| 路径/流结果 | 一致 | 一致 | 一致 |
| 缺文件异常 | `java.io.FileNotFoundException` | Aspose 内部类 | `com.aspose.slides.exceptions.FileNotFoundException` |
| 额外方法 | `options` / `tablesByName` / `rowsByName` | 无 | 无 |

## 4. `options()` 与 `Options` 默认值

### 4.1 新建

```java
WordParser.Options options = WordParser.options();
```

`options()` 每次返回一个新对象，不抛异常。虽然 `Options` 类本身可以 `new`，
推荐统一用 `options()`，以后默认值变化时调用方不用改。

### 4.2 默认值（实测）

| 配置 | 默认值 | 含义 |
| --- | --- | --- |
| 注册表头 | 0 组 | 不注册、不给关键字 = 静默模式 |
| `headerTolerance` | `1` | 允许 1 列不一致 |
| 表头关键字 | 空 | 不做关键字匹配 |
| `fullRow` | `KEEP` | 跨整行合并行保留并归到当前组 |
| `images` | `META` | 只要元信息 |
| `imageDir` | `null` | 不指定导出目录 |
| `source` | `true` | 每组带 `page` / `lastPage` |
| `unmatchedHeaders` | `false` | 不写顶层键 |

### 4.3 生命周期与线程

- 可变对象，链式方法“即改即生效”，配置阶段不要跨线程共享；
- 配好之后可以并发用于多次解析：解析过程只读选项，图片编号每次都从 1 开始；
- 同一次解析内部不共享 `Options` 的临时状态；
- `imageDir` 指向同一目录时，重名文件由 `Images.uniqueFile` 加序号，不会互相覆盖，
  但“先查后建”有并发窗口，极端并发请每次解析用不同目录。

## 5. 配置与取值方法逐个说明

`5.1`–`5.10` 是 `Options` 的链式方法，都返回 `this`，可以串起来；`5.11`–`5.14` 是三个入口类共用的静态方法。
每一项都给：签名、参数、作用、实测、坑、示例。

### 5.1 `headers(String... titles)`

**签名**

```java
public Options headers(String... titles)
```

**参数**

| 参数 | 类型 | 必填 | 含义 |
| --- | --- | --- | --- |
| `titles` | `String...` | 是 | 一组的完整列名，顺序必须与文档列顺序一致 |

**作用**

注册一组“完整表头”。文档里某一行与它匹配时：

1. 这一行被判为表头；
2. 结果里的 `headers` 使用**注册时写的原名**（不是从文档读出来的那一行）；
3. 如果这一组还起了名字，组里多一个 `name` 键。

可重复调用注册多组；匹配时按注册顺序取第一个命中的。

**怎么算匹配**

- 列数必须相同（文档行末尾的空列会先被裁掉，注册组本身不裁）；
- 允许 `headerTolerance` 列不一致（默认 1），但至少要有 1 列完全相等；
- 比较时忽略大小写与一切空格，所以 `客户 名称`、`客户　名称`、`客户名称` 等价。

**实测**

注册样例的 5 组真表头后：

| 格式 | 静默 | 注册（不起名） |
| --- | --- | --- |
| docx | 7 组 / 30 行，4740 字符 | 7 组 / 30 行，4740 字符（逐字节相同） |
| pdf | 6 组 / 45 行，6473 字符 | 6 组 / 45 行，6473 字符（逐字节相同） |
| pptx | 6 组 / 22 行，3526 字符 | 6 组 / 21 行，3378 字符 |

pptx 少 1 行，是因为注册后两行表头被合并，下层“客户名称/合同金额/…”那一行不再被当数据。

**坑**

- 不要把**行尾空列**写进注册：文档行比对前会裁掉末尾空列，注册组不会，列数就对不上了；
- 注册名本身重名或为空时，输出仍按统一规则补成 `名-列号` / `列N`；
- 写错一个字通常不会报错，只是“没命中”；打开 `unmatchedHeaders(true)` 才能在结果里看到。

**示例**

```java
WordParser.Options options = WordParser.options()
        .headers("片区", "区域", "销售额", "回款率", "销售额", "回款率", "备注")
        .headers("客户名称", "合同金额", "已回款", "回款率", "备注")
        .headers("工单类型", "工单数", "处理说明");
```

> 证据：`选项默认值与setter.txt`、`选项矩阵-三份样例.txt`、`TablesTest`。

### 5.2 `headerGroups(List<List<String>> groups)`

**签名**

```java
public Options headerGroups(List<List<String>> groups)
```

**参数**

| 参数 | 类型 | 必填 | 含义 |
| --- | --- | --- | --- |
| `groups` | `List<List<String>>` | 是 | 多组完整表头，外层每个元素是一组，内层顺序即列顺序 |

**作用**

批量注册，等价于按顺序对每个内层列表调用一次 `headers(...)`。返回 `this`。

**实测**

同一批表头，`headerGroups` 与逐组 `headers` 在 docx 上都是 7 组 / 30 行、4740 字符，
JSON 逐字符相同（`true`）。

**坑**

- `name(...)` 绑定的是**这一批的最后一组**；要给多组分别起名，改回逐组写法；
- Java 8 没有 `List.of`，用 `Arrays.asList`：

```java
WordParser.Options options = WordParser.options().headerGroups(Arrays.asList(
        Arrays.asList("片区", "区域", "销售额", "回款率", "销售额", "回款率", "备注"),
        Arrays.asList("客户名称", "合同金额", "已回款", "回款率", "备注")
));
```

> 证据：`选项默认值与setter.txt`。

### 5.3 `name(String label)`

**签名**

```java
public Options name(String label)
```

**参数**

| 参数 | 类型 | 必填 | 含义 |
| --- | --- | --- | --- |
| `label` | `String` | 是 | 组名；首尾空格会被去掉；`null` 或空串等于不给这一组起名 |

**作用**

给**写在它前面最近一次 `headers(...)` 注册的那一组**起名。只有“注册 + 起名 + 文档里命中”
三者都成立，结果的这一组才多一个 `name` 键，键顺序是：

```text
headers → rows → name → source
```

**实测**

注册 5 组、其中 3 组起名后：

| 格式 | 带 `name` 的组 | `季度明细` | `回款台账` | `年度目标` |
| --- | --- | --- | --- | --- |
| docx | #1 #2 #6 #7 | 4 行 | 13 行 | 5 行（3 + 2） |
| pdf | #1 #2 #5 #6 | 4 行 | 31 行 | 5 行（3 + 2） |
| pptx | #1 #2 #4 #5 | 4 行 | 7 行 | 5 行（3 + 2） |

JSON 从“注册不起名”到“注册并起名”整串多 88 个字符：docx 4740 → 4828、
pdf 6473 → 6561、pptx 3378 → 3466（4 个带 `name` 的组，每个约 22 个字符）。

**坑**

- `name` 不是主键：同一张物理表切出的两个表头组可以同名（docx #6/#7、pdf #5/#6、pptx #4/#5）；
- 只靠关键字、格式标题行标记、首行兜底判出来的组**没有** `name`；
- 一组都没注册就调用 `name(...)`：不抛异常，标准错误提示后忽略；
- 名字比较是 `equals`：查询串会 `trim`，但中间空格与大小写都算差异；
- 改名等于改契约，下游按名字取数的代码会断。

**示例**

```java
WordParser.Options options = WordParser.options()
        .headers("片区", "区域", "销售额", "回款率", "销售额", "回款率", "备注")
        .name("季度明细")
        .headers("客户名称", "合同金额", "已回款", "回款率", "备注")
        .name("回款台账");
```

> 证据：`sales-report.*-注册表头-起名.txt`、`按名字取数-三份样例.txt`、`TableByNameTest`。

### 5.4 `headerTolerance(int columns)`

**签名**

```java
public Options headerTolerance(int columns)
```

**参数**

| 参数 | 类型 | 必填 | 含义 |
| --- | --- | --- | --- |
| `columns` | `int` | 是 | 允许几列不一致；`0` = 逐列完全相等；负数按 `0` 处理 |

**作用**

放宽或收紧表头匹配。无论取值多大，都仍然要求：

1. 列数相同（文档行末尾空列先裁掉）；
2. 至少 1 列完全相等。

第二条是为了防止“容忍度一宽，数据行也被当表头”。

**实测**

把核心表的“备注”故意注册成“说明”（差 1 列）：

| 取值 | 结果 | 标准错误 |
| --- | --- | --- |
| `1`（默认） | 命中，`rowsByName("季度明细")` 4 行 | 空 |
| `0` | 不命中，没有 `name`，0 行 | `[WordParser] 表头集合未匹配到任何行：[...]` |

**坑**

- 放宽后名字也可能“误命中”相近的表头；起名与容忍度要一起考虑；
- 容忍度放宽的是“格子的值”，**从不放宽列数**。

> 证据：`选项默认值与setter.txt`、`docx-台账四种注册写法.txt`、`TablesTest`。

### 5.5 `keywords(String... values)`

**签名**

```java
public Options keywords(String... values)
```

**参数**

| 参数 | 类型 | 必填 | 含义 |
| --- | --- | --- | --- |
| `values` | `String...` | 是 | 关键字；`null` 与纯空白项被忽略 |

**作用**

宽松兜底：一行的**每个非空单元格**都含任意一个关键字，这一行才判为表头。
传了关键字就关闭“静默模式的首行兜底”，但不会关闭已注册表头与格式自带标题行标记。

**实测**

| 场景 | 结果 |
| --- | --- |
| docx，`keywords("销售")` | 7 组 / 30 行，4740 字符，与静默 JSON 逐字符相同 |
| pdf，5 个关键字 | 6 组 / 47 行，6445 字符（首行兜底被关闭，2 张表的表头行落进数据，列名退化为 `列1…`） |
| pptx，5 个关键字 | 6 组 / 22 行，3526 字符（与静默相同） |

**坑**

- 关键字是“包含”而不是“等于”；
- 空单元格不参与判断；
- 一整行都是空时不可能是表头；
- 入库请用完整表头注册，关键字只适合探路。

> 证据：`公共入口-四个重载.txt`、`选项默认值与setter.txt`。

### 5.6 `fullRow(FullRow mode)`

**签名**

```java
public Options fullRow(WordParser.FullRow mode)
```

**参数**

| 参数 | 类型 | 必填 | 含义 |
| --- | --- | --- | --- |
| `mode` | `FullRow` | 是 | `KEEP` / `FIRST` / `DROP`；传 `null` 等同 `KEEP` |

**作用**

决定“跨整行合并的行”（例如 `2025 年度（已结账）` 这种分组小标题）的去向。详见 6.1。

**实测**

| 格式 | `KEEP` | `FIRST` | `DROP` |
| --- | --- | --- | --- |
| docx | 7 组 / 30 行，4740 字符 | 7 组 / 30 行，4737 字符 | 6 组 / 28 行，4436 字符 |
| pdf | 6 组 / 45 行，6473 字符 | 6 组 / 45 行，6470 字符 | 5 组 / 43 行，6169 字符 |
| pptx | 6 组 / 21 行，3378 字符 | 6 组 / 21 行，3375 字符 | 5 组 / 19 行，3089 字符 |

**坑**

- 分组行**后面紧跟的那一行按表头处理**，所以“分组标题 + 直接放数据”会多切一组；
- `FIRST` 不改组数与行数，只改归属；`DROP` 会同时减少组数与行数。

> 证据：`选项矩阵-三份样例.txt`、`TablesTest`。

### 5.7 `images(ImageMode mode)`

**签名**

```java
public Options images(WordParser.ImageMode mode)
```

**参数**

| 参数 | 类型 | 必填 | 含义 |
| --- | --- | --- | --- |
| `mode` | `ImageMode` | 是 | `META` / `BYTES` / `BASE64` / `FILE`；传 `null` 等同 `META` |

**作用**

决定图片内容放哪里。任何模式都给元信息：`fileName`、`type`、`width`、`height`、`size`、`contentId`。
详见 6.2 与 7.5。

**实测（docx，一张 1024×768 PNG，670939 字节）**

| 模式 | 多出的字段 | 整串 JSON 字符数 |
| --- | --- | --- |
| `META` | 无 | 4740 |
| `BYTES` | `bytes`（`byte[]`） | 899347 |
| `BASE64` | `base64`（带 `data:image/png;base64,` 前缀） | 899370 |
| `FILE` | `path`（绝对路径） | 4832 |

**坑**

- `BASE64` / `BYTES` 会把整篇图片字节留在结果里，图片多时优先 `FILE` 或 `META`；
- 默认 `META` 是为“只要文字表格”的场景省体积，不是功能缺失。

> 证据：`docx-图片四种模式.txt`、`序列化与图片导出.txt`。

### 5.8 `imageDir(File directory)`

**签名**

```java
public Options imageDir(File directory)
```

**参数**

| 参数 | 类型 | 必填 | 含义 |
| --- | --- | --- | --- |
| `directory` | `File` | 是 | 导出目录；不存在时在有图片要写的那一刻创建（含父目录），没有图片就不会创建 |

**作用**

只在 `ImageMode.FILE` 下生效：每张图写进这个目录，结果里留 `fileName` 与 `path`（绝对路径）。

**实测**

- `FILE` + 目录：写出 `image-1.png`，670939 字节，SHA-1 与 `contentId` 相同；
- `FILE` 但没给 `imageDir`：解析到第一张图时抛
  `IllegalStateException: ImageMode.FILE 需要先调用 imageDir(File) 指定导出目录`；
- `META` 却给了 `imageDir`：不写文件，标准错误提示
  `[WordParser] imageDir 只在 ImageMode.FILE 下生效，当前是 META，... 不会写入图片`
  （前缀固定是 `[WordParser]`：三种格式共用 `Images`，用 PDF/PPT 入口时也是这个前缀）。

**坑**

- 相对路径按进程工作目录解析，结果里的 `path` 始终是绝对路径；
- 重名文件会加 `-2` 后缀，结果里的 `fileName` 是实际落盘名；
- 同一目录并发写有“先查后建”窗口，见 4.3。

> 证据：`docx-图片四种模式.txt`、`错误路径-三份样例.txt`、`ImagesTest`。

### 5.9 `source(boolean enabled)`

**签名**

```java
public Options source(boolean enabled)
```

**参数**

| 参数 | 类型 | 必填 | 含义 |
| --- | --- | --- | --- |
| `enabled` | `boolean` | 是 | `true` 每组带 `source`；`false` 整键不写 |

**作用**

控制组里的来源位置。`source = {page: 起始页}`，跨页再加 `lastPage`；
pptx 的 `page` 就是幻灯片序号。三种格式一起生效。

**实测**

| 格式 | 开（默认） | 关 |
| --- | --- | --- |
| docx（起名） | 4828 字符 | 4497 字符 |
| pdf（起名） | 6561 字符 | 6274 字符 |
| pptx（起名） | 3466 字符 | 3202 字符 |

关掉后每组少 `source` 一个键（起过名的组仍有 `name`），列名与内容一字不变。docx 因此不创建 `LayoutCollector`，
同一份样例交替实测：开 74–254ms，关 11–35ms。

**坑**

- 某格式给不出位置（`page <= 0`）时整键不写，不要假设永远有 `source`；
- 同一张物理表切出的多组写同一段页码（组按表头切，不按页切）。

> 证据：`选项矩阵-三份样例.txt`、`CrossFormatTest`。

### 5.10 `unmatchedHeaders(boolean enabled)`

**签名**

```java
public Options unmatchedHeaders(boolean enabled)
```

**参数**

| 参数 | 类型 | 必填 | 含义 |
| --- | --- | --- | --- |
| `enabled` | `boolean` | 是 | `true` 才写顶层键 `unmatchedHeaders` |

**作用**

把“注册了却一组都没命中”的表头集合放进结果，供程序读取。默认关闭；
关闭时输出与不调用该方法完全相同（逐字节）。

**形状**

```json
[
  { "headers": ["故障部件", "工单数"], "name": "部件统计" }
]
```

每一项：

- `headers`：注册时写的**原文**（不是归一化后的小写去空格形式）；
- `name`：注册时起过名字才有；
- 全部命中时是空列表 `[]`，不是 `null`。

**实测**

| 格式 | 清单 | 标准错误 |
| --- | --- | --- |
| docx | `[]`（文档确实有该表） | 空 |
| pdf | 1 组 | `[PdfParser] 表头集合未匹配到任何行：[故障部件, 工单数]` |
| pptx | 1 组 | `[PptParser] 表头集合未匹配到任何行：[故障部件, 工单数]` |

打开后 JSON 字符数：docx 4828 → 4857、pdf 6561 → 6632、pptx 3466 → 3537。

**坑**

- 同一个 `name` 之下已有一份注册命中时，另一份命不中**既不提示也不进清单**；
- 没起名字的注册不受这条抑制，照样报；
- 未命中不会让解析失败，结果照给，只是那一组数据可能并进了上一组。

> 证据：`未匹配表头-三份样例.txt`、`CrossFormatTest`。

### 5.11 `toJson(Object value)`

**签名**

```java
public static String toJson(Object value) throws Exception   // WordParser / PdfParser / PptParser 各一份
```

**参数**

| 参数 | 类型 | 必填 | 含义 |
| --- | --- | --- | --- |
| `value` | `Object` | 是 | 一般就是 `parse` 的返回值；任何 Jackson 能序列化的对象都行；传 `null` 输出 4 个字符 `null` |

**作用**

把结果整串序列化成缩进 JSON。三种格式共用同一份实现（内部的 `Results.toJson`），
所以同一个对象用三个入口得到的字符串逐字符相同（`序列化与图片导出.txt` 里三份样例都是 `true`）。

输出格式（逐条实测）：

- 2 空格缩进，键与值之间是 `"key" : value`（冒号前后各一个空格）；
- 中文不转义，原样输出；
- `byte[]`（`ImageMode.BYTES`）编成**不带** `data:` 前缀的 base64；
- 末尾**没有**换行，最后一个字符是 `}`（docx 4740 个字符，第 4740 个是 `}`）；
- 换行是平台默认分隔符：Windows 是 `\r\n`（docx 的 4740 个字符里有 219 个 `\r`），
  Linux/macOS 是 `\n`，同一份结果在 Linux 上会是 4521 个字符；
- 键顺序就是插入顺序：顶层 `text` → `tables` → `images` →（可选）`unmatchedHeaders`；
  组内 `headers` → `rows` →（可选）`name` →（可选）`source`。

**实测**

三份样例的整串字符数：

| 样例 | `META`（默认） | `BASE64` |
| --- | --- | --- |
| docx | 4740 | 899370 |
| pdf | 6473 | 58076 |
| pptx | 3526 | 898156 |

`toJson(null)` 得到 4 个字符的 `null`，不是空串。

**坑**

- 字符数不是字节数：这里统计的是 Java `String.length()`，写成 UTF-8 文件后体积会变大；
- 换行随平台变，拿 JSON 做签名或摘要前先统一换行；本文所有“逐字符相同”都指同一平台内的比较；
- 不带末尾换行，直接拿去拼日志时注意自己补；
- `BASE64` / `BYTES` 的结果动辄近 900KB，别把整串塞进日志或数据库宽表；
- 签名带 `throws Exception`，序列化失败会把异常抛给调用方。

**示例**

```java
import cn.net.tianxi.document.WordParser;

import java.io.File;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.util.Map;

public class ToJsonDemo {
    public static void main(String[] args) throws Exception {
        Map<String, Object> result = WordParser.parse("src/test/resources/word/sales-report.docx");
        String json = WordParser.toJson(result);
        Files.write(new File("sales-report.json").toPath(), json.getBytes(StandardCharsets.UTF_8));
        System.out.println("字符数=" + json.length() + "，末尾字符=" + json.charAt(json.length() - 1));
    }
}
```

实测输出：`字符数=4740，末尾字符=}`。

> 证据：`序列化与图片导出.txt`、`全部模式-完整JSON.txt`、`ResultsTest`。

### 5.12 `writeImages(Map<String, Object> result, File directory)`

**签名**

```java
public static void writeImages(Map<String, Object> result, File directory) throws Exception   // 三个类各一份
```

**参数**

| 参数 | 类型 | 必填 | 含义 |
| --- | --- | --- | --- |
| `result` | `Map<String, Object>` | 是 | `parse` 的返回值，不能为 `null`；里面没有 `images` 列表时静默返回 |
| `directory` | `File` | 是 | 输出目录；只在真的写第一张图那一刻创建（含父目录）。传 `null` 不报错，会写进进程当前工作目录 |

**作用**

只处理带 `base64` 字段的条目（也就是 `ImageMode.BASE64` 的结果）：解码 → 用 `fileName` 消毒出文件名 →
重名加 `-2`、`-3` 后缀 → 写文件。用 `ImageMode.FILE` 解析时图片已经落过盘，不需要再调它。

**实测**（`序列化与图片导出.txt`，`BASE64` 模式）

| 样例 | 写出文件 | 大小 | 校验 |
| --- | --- | --- | --- |
| docx | `image-1.png` | 670939 字节 | SHA-1 与 `contentId` 相同 |
| pdf | `image-1.jpg` | 38668 字节 | SHA-1 与 `contentId` 相同 |
| pptx | `image-1.png` | 670939 字节 | SHA-1 与 `contentId` 相同 |

对默认 `META` 结果调用：目录**不存在、0 个文件**，标准错误打印
`[writeImages] 结果里的图片没有 base64 内容，一张都没导出 —— 解析时用 WordParser.ImageMode.BASE64（默认是 META，只给元信息），或者改用 ImageMode.FILE 配合 imageDir(File) 边解析边落盘`。

**坑**

- 只能导 `BASE64`：`META` 没内容（会提示），`BYTES` 只有 `byte[]` 也导不出，`FILE` 已经写过；
- `result` 传 `null` 直接 `NullPointerException`（实测）；
- `directory` 传 `null` 会把图片写进工作目录且不提示（实测），生产代码里始终显式给目录；
- 文件名来自结果数据，`../../evil.png` 会被削成目录内的 `evil.png`（`ResultsTest` 有专门用例）；
- 重名不覆盖：第二次写同名文件得到 `image-1-2.png`；
- 不清理目录里的旧文件，也没有“本次导出了几张”的返回值；
- 同一目录并发写有“先查后建”窗口，见 4.3。

**示例**

```java
import cn.net.tianxi.document.WordParser;

import java.io.File;
import java.util.List;
import java.util.Map;

public class WriteImagesDemo {
    public static void main(String[] args) throws Exception {
        Map<String, Object> result = WordParser.parse(
                "src/test/resources/word/sales-report.docx",
                WordParser.options().images(WordParser.ImageMode.BASE64));
        WordParser.writeImages(result, new File("out/images"));
        System.out.println(((List<?>) result.get("images")).size() + " 张图片已导出到 out/images");
    }
}
```

实测输出：`1 张图片已导出到 out/images`，目录下是 `image-1.png`（670939 字节）。

> 证据：`序列化与图片导出.txt`、`docx-图片四种模式.txt`、`ResultsTest`。

### 5.13 `tablesByName(Map<String, Object> result, String name)`

**签名**

```java
public static List<Map<String, Object>> tablesByName(Map<String, Object> result, String name)
```

**入口说明**：只有 `WordParser` 上有这个方法；`PdfParser.parse` / `PptParser.parse` 的结果照样传得进来，
因为三种格式的返回结构一样。

**参数**

| 参数 | 类型 | 必填 | 含义 |
| --- | --- | --- | --- |
| `result` | `Map<String, Object>` | 是 | 任一解析入口的返回值；`null` 返回空列表，不抛异常 |
| `name` | `String` | 是 | `options().name("...")` 注册时写的名字；只去首尾空格，其余按 `equals` 精确比较（区分大小写、中间空格不去）；`null` 或全空白返回空列表 |

**作用**

在 `tables` 里找所有 `name` 等于查询串的组，按原顺序返回。返回的是结果里的**原对象**，不是副本。
组数可能大于 1：同一张物理表被两段表头切成两组时，两组可以同名。

**实测**（docx，注册 5 组、其中 3 组起名）

| 查询 | 命中组数 | 说明 |
| --- | --- | --- |
| `季度明细` | 1 | 第 1 组，4 行 |
| `回款台账` | 1 | 第 2 组，13 行 |
| `年度目标` | 2 | 第 6、7 组，分别 3 行与 2 行 |
| `不存在的名字` | 0 | 空列表，不是 `null` |
| `""` / `null` / `"   "` | 0 | 不抛异常 |
| `"  季度明细  "` | 1 | 首尾空格被去掉 |
| `"季度 明细"` | 0 | 中间空格算差异 |

pdf 与 pptx 的 `年度目标` 也都命中 2 组（pdf 是第 5、6 组，pptx 是第 4、5 组）。

**坑**

- 名字不是主键，别假设唯一；要行数据用 5.14；
- 返回原对象，改它等于改 `result`；
- 只认 `Options.name(...)` 起的名，表头文字本身不是名字；
- 没起名的组永远查不到（包括“注册了但没起名”的组）；
- 查询串不做表头那套归一化：表头比对忽略大小写与空格，名字比对不忽略。

**示例**

```java
import cn.net.tianxi.document.WordParser;

import java.util.List;
import java.util.Map;

public class TablesByNameDemo {
    public static void main(String[] args) throws Exception {
        WordParser.Options options = WordParser.options()
                .headers("片区", "年度目标", "完成率").name("年度目标");
        Map<String, Object> result = WordParser.parse("src/test/resources/word/sales-report.docx", options);

        List<Map<String, Object>> groups = WordParser.tablesByName(result, "年度目标");
        System.out.println("命中 " + groups.size() + " 组");
        for (Map<String, Object> group : groups) {
            System.out.println(group.get("headers") + " -> " + ((List<?>) group.get("rows")).size() + " 行");
        }
    }
}
```

实测输出：

```text
命中 2 组
[片区, 年度目标, 完成率, 列4] -> 3 行
[片区, 年度目标, 完成率, 列4] -> 2 行
```

三份样例用同一份注册跑出来组数与行数都一样（pptx 的数值没有 `.00`，见 9.3）。

> 证据：`按名字取数-三份样例.txt`、`TableByNameTest`。

### 5.14 `rowsByName(Map<String, Object> result, String name)`

**签名**

```java
public static List<Map<String, String>> rowsByName(Map<String, Object> result, String name)
```

**参数**

| 参数 | 类型 | 必填 | 含义 |
| --- | --- | --- | --- |
| `result` | `Map<String, Object>` | 是 | 同 5.13：任一解析入口的返回值，`null` 返回空列表 |
| `name` | `String` | 是 | 同 5.13：注册名，去首尾空格后 `equals`；`null` / 全空白返回空列表 |

**作用**

先按 5.13 找到所有同名组，再把各组的 `rows` 按组顺序首尾相接成一个列表。行还是原来的行对象，
键就是**那一组**的列名；同名各组的列名不一致时，行与行的键可能不同。

**实测**

| 查询 | docx | pdf | pptx |
| --- | --- | --- | --- |
| `年度目标` | 5 行（3 + 2） | 5 行（3 + 2） | 5 行（3 + 2） |
| `季度明细` | 4 行 | 4 行 | 4 行 |
| `回款台账` | 13 行 | 31 行 | 7 行 |

只注册台账（合并后那一行）时同样是 docx 13、pdf 31、pptx 7，容忍度 `0` 与 `1` 的结果一样。
取不到的名字返回空列表。

docx 的 `年度目标` 五行依次是：

```text
{片区=华东大区, 年度目标=2400.00, 完成率=101%, 列4=达标}
{片区=华东南区, 年度目标=1500.00, 完成率=93%, 列4=未达标}
{片区=2026 年度（预算下达）, 年度目标=, 完成率=, 列4=}
{片区=华东大区, 年度目标=2800.00, 完成率=—, 列4=}
{片区=华东南区, 年度目标=1750.00, 完成率=—, 列4=}
```

**坑**

- 只拼行，不带组边界：想知道“哪几行来自哪一组”用 5.13 自己遍历；
- 不按列合并、不去重、不排序，顺序就是组顺序加组内行顺序；
- 同名各组列名不同时，落库前先校验 `headers`，否则同一个字段名在两张表里可能含义不同；
- 空单元格是 `""`，不是 `null`；正常结果里不会缺键，需要判断“有没有这一列”用 `containsKey`；
- 返回原对象，改行等于改结果。

**示例**

```java
import cn.net.tianxi.document.WordParser;

import java.util.List;
import java.util.Map;

public class RowsByNameDemo {
    public static void main(String[] args) throws Exception {
        WordParser.Options options = WordParser.options()
                .headers("客户名称", "合同金额", "已回款", "回款率", "备注").name("回款台账");
        Map<String, Object> result = WordParser.parse("src/test/resources/word/sales-report.docx", options);

        List<Map<String, String>> rows = WordParser.rowsByName(result, "回款台账");
        System.out.println("共 " + rows.size() + " 行");
        System.out.println("第 1 行：" + rows.get(0));
        System.out.println("最后一行：" + rows.get(rows.size() - 1));
    }
}
```

实测输出：

```text
共 13 行
第 1 行：{客户名称=上海浦东城建集团, 合同金额=486.00, 已回款=486.00, 回款率=100%, 备注=已结清}
最后一行：{客户名称=合计, 合同金额=2096.30, 已回款=1796.10, 回款率=86%, 备注=}
```

> 证据：`按名字取数-三份样例.txt`、`TableByNameTest`。

## 6. 两个枚举：`FullRow` 与 `ImageMode`

### 6.1 `FullRow`

定义“跨整行合并行”（整行只剩一个起始格）的去向：

| 常量 | 行为 | docx 实测 |
| --- | --- | --- |
| `KEEP`（默认） | 保留，归到它上面最近的表头组，作为该组一行数据，其余列为空串 | 7 组 / 30 行 |
| `FIRST` | 保留，统一归到**同一张物理表**的第一组 | 7 组 / 30 行，4737 字符 |
| `DROP` | 丢弃 | 6 组 / 28 行，4436 字符 |

`FIRST` 的实测细节：`2026 年度（预算下达）` 从目标表那组挪到了
`["列1","列2","列3","列4"]` 那组，那组因此有 2 行；组数与总行数不变，只是归属变了。

`DROP` 会同时减少组数：`2025 年度（已结账）` 那组整组消失。

**注意**：分组行后面紧跟的那一行按表头处理，所以分组行之后直接放数据会多切一组。

### 6.2 `ImageMode`

| 常量 | 结果里放什么 | 适用场景 | docx JSON |
| --- | --- | --- | --- |
| `META`（默认） | 只有元信息 | 只要文字表格 | 4740 字符 |
| `BYTES` | `bytes`：原始 `byte[]`；Jackson 序列化成**不带 `data:` 前缀**的 base64 字符串 | Java 侧转存、送 OCR | 899347 字符 |
| `BASE64` | `base64`：`data:image/png;base64,...` 完整 data URI | 前端直接塞 `<img src>` | 899370 字符 |
| `FILE` | `path`：`imageDir` 下的绝对路径；`fileName` 是实际落盘名 | 图片多，边解析边落盘 | 4832 字符 |

任何模式都给的元信息：`fileName`、`type`、`width`、`height`、`size`、`contentId`。
外链图片例外：没有 `size` / `contentId`，只有 `link`。

**去重**：只有 `BASE64` / `BYTES` 做内容去重。重复图片不再重复内嵌，只留
`"duplicateOf": 0`（指向前一条的 0 基下标），且不带内容字段；`contentId` 仍是它自己的指纹。
`META` 不做去重标记；`FILE` 每张图各落一个文件。

> 证据：`docx-图片四种模式.txt`、`ImagesTest`。

## 7. 返回数据字典

### 7.1 顶层

| 键 | 类型 | 何时出现 | 说明 |
| --- | --- | --- | --- |
| `text` | `String` | 永远 | 段落用 `\n` 拼接；表格里的文字不在里面重复 |
| `tables` | `List<Map<String,Object>>` | 永远 | 没有表时为空列表 |
| `images` | `List<Map<String,Object>>` | 永远 | 没有图时为空列表 |
| `unmatchedHeaders` | `List<Map<String,Object>>` | `unmatchedHeaders(true)` | 全部命中时为空列表 |

键顺序固定为 `text` → `tables` → `images` → `unmatchedHeaders`（打开时）。

### 7.2 `tables` 的元素

| 键 | 类型 | 何时出现 | 说明 |
| --- | --- | --- | --- |
| `headers` | `List<String>` | 永远 | 列名；空表头补 `列N`，重名补 `名-绝对列号` |
| `rows` | `List<Map<String,String>>` | 永远 | 每行是 `LinkedHashMap`，键顺序与 `headers` 逐个相等 |
| `name` | `String` | 注册 + 起名 + 命中 | 组名；可以多个组同名 |
| `source` | `Map<String,Integer>` | 默认；`source(false)` 时无 | `page` 起始页/幻灯片号；跨页加 `lastPage` |

组内键顺序固定：`headers` → `rows` → `name` → `source`。
没有 `name` 时是 `headers` → `rows` → `source`；
关掉 `source` 时是 `headers` → `rows` → `name`（起过名才有 `name`）。

一个组的 JSON 形状（docx 年度目标第二段，2 行）：

```json
{
  "headers" : [ "片区", "年度目标", "完成率", "列4" ],
  "rows" : [ {
    "片区" : "华东大区",
    "年度目标" : "2800.00",
    "完成率" : "—",
    "列4" : ""
  }, {
    "片区" : "华东南区",
    "年度目标" : "1750.00",
    "完成率" : "—",
    "列4" : ""
  } ],
  "name" : "年度目标",
  "source" : {
    "page" : 2
  }
}
```

### 7.3 行

| 项 | 内容 |
| --- | --- |
| 类型 | `Map<String,String>`，实现是 `LinkedHashMap` |
| 键 | 与所属组的 `headers` 逐个相等，顺序相同 |
| 值 | 单元格清洗后的文本，**永远是字符串**，不做数字/日期/百分比转换 |
| 空单元格 | 空串 `""`，不是 `null` |
| 缺格 | 补空串，保证每行键数与 `headers` 一致 |
| 合并单元格 | 横向合并每格都回填合并值；纵向合并延续格回填首格的值 |
| 数字与符号 | 原样保留：`486.00`、`92%`、`—`、`2096.30` 都不会被改写 |
| 域 | 只读域缓存结果；缓存为空给空串，不重算 |

### 7.4 `source`

| 字段 | 类型 | 何时出现 | 含义 |
| --- | --- | --- | --- |
| `page` | `int` | 默认 | 起始页或幻灯片序号，1 起 |
| `lastPage` | `int` | 跨页且 `lastPage > page` | 结束页 |

- docx：`LayoutCollector` 给出的页码；`source(false)` 时连 collector 都不建；
- pdf：页码，1 起；跨页并表时记首末页；
- pptx：幻灯片序号，每张表单页；
- 某格式给不出位置（`page <= 0`）时整键不写。

### 7.5 `images` 的元素

| 键 | 类型 | 何时出现 | 说明与样例 |
| --- | --- | --- | --- |
| `fileName` | `String` | 永远 | `image-1.png`，每次解析从 1 开始；FILE 模式是实际落盘名 |
| `type` | `String` | 永远 | 大写类型：`PNG`、`JPEG`、`GIF`、`BMP`、`TIFF`、`EMF`、`WMF`、`UNKNOWN` |
| `width` | `int` | 永远 | 像素宽；读不出时为 `0` |
| `height` | `int` | 永远 | 像素高；读不出时为 `0` |
| `size` | `int` | 内嵌图片 | 原始字节数，`670939` |
| `contentId` | `String` | 内嵌图片 | 字节的 SHA-1 十六进制，40 字符 |
| `base64` | `String` | `BASE64` | `data:image/png;base64,...` |
| `bytes` | `byte[]` | `BYTES` | 原始字节；JSON 里是不带前缀的 base64 字符串 |
| `path` | `String` | `FILE` | 落盘后的绝对路径 |
| `duplicateOf` | `Integer` | 内嵌模式的重复图 | 第一条相同图片的 0 基下标 |
| `link` | `String` | 外链图片 | 原始链接路径；不读文件、不联网 |

字段顺序：`fileName` → `type` → `width` → `height` → `size` → `contentId` →
内容字段（`base64` / `bytes` / `path`）或 `duplicateOf`。
外链是 `fileName` → `type` → `width` → `height` → `link`。

### 7.6 `unmatchedHeaders` 的元素

| 键 | 类型 | 何时出现 | 说明 |
| --- | --- | --- | --- |
| `headers` | `List<String>` | 永远 | 注册时写的原文列名 |
| `name` | `String` | 注册时起过名 | 组名 |

### 7.7 字段出现矩阵

| 字段 | META | BYTES | BASE64 | FILE | 外链 |
| --- | --- | --- | --- | --- | --- |
| `fileName` | ✓ | ✓ | ✓ | ✓ | ✓ |
| `type` | ✓ | ✓ | ✓ | ✓ | ✓ |
| `width` / `height` | ✓ | ✓ | ✓ | ✓ | ✓ |
| `size` | ✓ | ✓ | ✓ | ✓ | ✗ |
| `contentId` | ✓ | ✓ | ✓ | ✓ | ✗ |
| `base64` | ✗ | ✗ | ✓ | ✗ | ✗ |
| `bytes` | ✗ | ✓ | ✗ | ✗ | ✗ |
| `path` | ✗ | ✗ | ✗ | ✓ | ✗ |
| `duplicateOf` | ✗ | 重复时 | 重复时 | ✗ | ✗ |
| `link` | ✗ | ✗ | ✗ | ✗ | ✓ |

## 8. 三种格式分别能解析出什么

### 8.1 通用

三种格式都输出同一套 `text` / `tables` / `images` 结构；
表头分组算法、列名补全、`name`、`source`、`unmatchedHeaders`、图片四模式完全共用。
差别只在“能读到哪些内容”和“表格怎么还原”。

### 8.2 docx

| 能读到 | 说明 |
| --- | --- |
| Body 段落 | 13 段；域结果、软换行、内容控件都覆盖 |
| 表格 | 7 组 / 30 行；含嵌套子表、双表头组、纵向合并、空白整行 |
| 图片 | 段落里与单元格里的都收；外链只给 `link` |
| 页码 | `LayoutCollector`，跨页组带 `lastPage` |
| 读不到 | 页眉页脚、脚注尾注、批注、未计算域的域代码 |

嵌套表：子表不摊进母表格子，母表那一格留空，子表作为**额外元素**排在 `tables` 里。
嵌套上限 8 层，容器（内容控件等）上限 16 层；超限只跳过该分支并在标准错误提示
（表格层级每张表最多提示一次，容器层级每个被跳过的分支各提示一次）。

### 8.3 pdf

| 能读到 | 说明 |
| --- | --- |
| 表格 | 6 组 / 45 行；靠线段几何还原，跨页长表并成一组 |
| 无边框表 | 第二遍文字流引擎补读；每行单元格数要一致，且不与有线表重叠 |
| 正文段落 | 从 `MarkupSection` 取；样例正文全在表格里，所以 `text` 是 0 段 |
| 图片 | JPEG 38668 字节，699×524 |
| 页码 | `page`；跨页组带 `lastPage` |
| 读不到 | 扫描件/图片型 PDF（无文字层）、无边框且每行格数不一致的表 |

跨页并表条件：相邻页 + 列数相同 + 左边界对齐 + 前一截是上页最后一块 +
后一截是本页第一块 + 后一截找不出表头行。中间页有别的表或图片时链条断开。

已知误报：用制表位排齐的几行散文本可能被当成一张表。

### 8.4 pptx

| 能读到 | 说明 |
| --- | --- |
| 正片文本框 | 9 段；每个形状的段落按行拆 |
| 表格 | 6 组 / 22 行；合并单元格回填、组合形状递归展开（上限 32 层） |
| 图片 | 与 docx 同一份 PNG；视频取封面图 |
| 页码 | 幻灯片序号 |
| 读不到 | 备注页、母版、版式页、SmartArt 文字、图表数据、OLE 对象内部内容 |

两行表头是 pptx 最容易踩的点：静默模式下没有把上下两行合并，下层表头行会落进数据，
台账变成 8 行；注册合并后的表头（或上下两种写法各注册一次）后变成 7 行。

### 8.5 三格式差异总表

| 维度 | docx | pdf | pptx |
| --- | --- | --- | --- |
| 默认组数 / 行数 | 7 / 30 | 6 / 45 | 6 / 22 |
| 默认 JSON 字符 | 4740 | 6473 | 3526 |
| 正文档数 | 13 | 0 | 9 |
| 图片 | 1024×768 PNG | 699×524 JPEG | 同一份 PNG |
| 回款台账行数 | 13 | 31 | 7（注册后）/ 8（静默） |
| 台账合计金额 | `2096.30` / `1796.10`（域缓存） | 两格空（导出丢域） | `1586` / `1384`（普通文本） |
| 两行表头静默 | 已合并 | 已摊平 | **未合并**，下层落进数据 |
| 注册后两行表头 | 合并 | 合并 | 合并 |
| 跨页/跨片 | 台账 `{page=1,lastPage=2}` | 台账 `{page=1,lastPage=2}` | 每张表单页 |
| 缺文件异常 | `java.io.FileNotFoundException` | Aspose 内部类 | `com.aspose.slides.exceptions.FileNotFoundException` |
| 评估模式影响 | 正文照读；页眉水印不进结果 | 文字层照读（扫描件本来就没有文字层） | 每段超过 5 字符会被截断（本例文字都 ≤5 字符，未触发） |

## 9. 三份样例的完整数据

三份样例都在 `src/test/resources/` 下：

| 格式 | 文件 | 大小 |
| --- | --- | --- |
| docx | `word/sales-report.docx` | 675208 字节 |
| pdf | `pdf/sales-report.pdf` | 184977 字节 |
| pptx | `ppt/sales-report.pptx` | 679289 字节 |

下面默认选项（`KEEP` / `META` / `source=true` / `unmatchedHeaders=false`）。

### 9.1 docx 正文（13 段）

```text
[1] 2025 年度华东区销售业绩报告
[2] 报告页次：3 页
[3] 归档编号：
[4] 本报告覆盖华东大区与华东南区两个销售片区，统计口径为含税销售额。 数据来源：销售管理系统与财务系统对账结果，金额单位：万元。
[5] 编制部门：销售管理部 审核：财务总监
[6] 一、季度销售明细
[7] 注：回款率 = 已回款金额 ÷ 含税销售额，四舍五入保留整数。
[8] 二、客户回款台账
[9] 三、售后工单统计
[10] 四、分片区年度目标完成情况
[11] 五、系统看板截图
[12] 附图 1：销售管理系统业绩看板截图
[13] 以上数据截至 2025 年 12 月 31 日，最终以财务系统结账数为准。
```

第 2 段证明域结果保留（`PAGE` 的缓存值 3），第 3 段证明未计算域整块消失，
第 4 段句号后的空格是 Shift+Enter 软换行留下的。

### 9.2 docx 的 7 组表（30 行）

**#1 季度明细** 7 列 × 4 行，`{page=1}`

```text
headers = [片区, 区域, 销售额, 回款率, 销售额-5, 回款率-6, 备注]
1: 华东大区 | 上海 | 1280 | 92% | 1465 | 95% | 超额完成
2: 华东大区 | 江苏 | 986 | 88% | 1120 | 90% |
3: 华东南区 | 浙江 | 760 | 85% | 890 | 87% | 回款偏长
4: 华东南区 | 安徽 | 540 | 80% | 616 | 83% |
```

**#2 客户回款台账** 5 列 × 13 行，`{page=1, lastPage=2}`

```text
headers = [客户名称, 合同金额, 已回款, 回款率, 备注]
1: 上海浦东城建集团 | 486.00 | 486.00 | 100% | 已结清
2: 江苏恒通机械制造 | 320.50 | 280.00 | 87% |
3: 杭州云栖科技 | 275.80 | 200.00 | 73% | 逾期
4: 苏州精密仪器 | 198.00 | 198.00 | 100% | 已结清
5: 宁波海港物流 | 165.40 | 120.00 | 73% |
6: 合肥新能源 | 142.00 | 100.00 | 70% | 逾期
7: 南京软件谷 | 128.60 | 128.60 | 100% | 已结清
8: 温州轻工贸易 | 96.00 | 60.00 | 63% |
9: 无锡传感科技 | 88.50 | 88.50 | 100% | 已结清
10: 芜湖装备制造 | 76.20 | 40.00 | 52% | 逾期
11: 嘉兴纺织 | 65.00 | 65.00 | 100% | 已结清
12: 徐州工程机械 | 54.30 | 30.00 | 55% |
13: 合计 | 2096.30 | 1796.10 | 86% |
```

**#3 售后工单统计** 3 列 × 4 行，`{page=2}`

```text
headers = [工单类型, 工单数, 处理说明]
1: 硬件故障 | 42 |
2: 硬件故障 | 12 | 更换配件后复测通过
3: 软件异常 | 27 | 远程处理，平均响应 2 小时
4: 现场支持 | 15 | 派工上门，覆盖江浙沪皖
```

**#4 故障部件（嵌套子表）** 2 列 × 3 行，`{page=2}`

```text
headers = [故障部件, 工单数]
1: 读卡器 | 18
2: 打印机 | 15
3: 显示屏 | 9
```

**#5 分组标题被首行兜底成一组** 4 列 × 1 行，`{page=2}`

```text
headers = [列1, 列2, 列3, 列4]
1: 2025 年度（已结账） | | |
```

**#6 年度目标（第一段）** 4 列 × 3 行，`{page=2}`

```text
headers = [片区, 年度目标, 完成率, 列4]
1: 华东大区 | 2400.00 | 101% | 达标
2: 华东南区 | 1500.00 | 93% | 未达标
3: 2026 年度（预算下达） | | |
```

**#7 年度目标（第二段）** 4 列 × 2 行，`{page=2}`

```text
headers = [片区, 年度目标, 完成率, 列4]
1: 华东大区 | 2800.00 | — |
2: 华东南区 | 1750.00 | — |
```

`年度目标` 起名后 #6、#7 都带同一个 `name`；`rowsByName` 拼平是 5 行。

### 9.3 docx 的完整 JSON（4740 字符）

```json
{
  "text" : "2025 年度华东区销售业绩报告\n报告页次：3 页\n归档编号：\n本报告覆盖华东大区与华东南区两个销售片区，统计口径为含税销售额。 数据来源：销售管理系统与财务系统对账结果，金额单位：万元。\n编制部门：销售管理部 审核：财务总监\n一、季度销售明细\n注：回款率 = 已回款金额 ÷ 含税销售额，四舍五入保留整数。\n二、客户回款台账\n三、售后工单统计\n四、分片区年度目标完成情况\n五、系统看板截图\n附图 1：销售管理系统业绩看板截图\n以上数据截至 2025 年 12 月 31 日，最终以财务系统结账数为准。",
  "tables" : [ {
    "headers" : [ "片区", "区域", "销售额", "回款率", "销售额-5", "回款率-6", "备注" ],
    "rows" : [ {
      "片区" : "华东大区",
      "区域" : "上海",
      "销售额" : "1280",
      "回款率" : "92%",
      "销售额-5" : "1465",
      "回款率-6" : "95%",
      "备注" : "超额完成"
    }, {
      "片区" : "华东大区",
      "区域" : "江苏",
      "销售额" : "986",
      "回款率" : "88%",
      "销售额-5" : "1120",
      "回款率-6" : "90%",
      "备注" : ""
    }, {
      "片区" : "华东南区",
      "区域" : "浙江",
      "销售额" : "760",
      "回款率" : "85%",
      "销售额-5" : "890",
      "回款率-6" : "87%",
      "备注" : "回款偏长"
    }, {
      "片区" : "华东南区",
      "区域" : "安徽",
      "销售额" : "540",
      "回款率" : "80%",
      "销售额-5" : "616",
      "回款率-6" : "83%",
      "备注" : ""
    } ],
    "source" : {
      "page" : 1
    }
  }, {
    "headers" : [ "客户名称", "合同金额", "已回款", "回款率", "备注" ],
    "rows" : [ {
      "客户名称" : "上海浦东城建集团",
      "合同金额" : "486.00",
      "已回款" : "486.00",
      "回款率" : "100%",
      "备注" : "已结清"
    }, {
      "客户名称" : "江苏恒通机械制造",
      "合同金额" : "320.50",
      "已回款" : "280.00",
      "回款率" : "87%",
      "备注" : ""
    }, {
      "客户名称" : "杭州云栖科技",
      "合同金额" : "275.80",
      "已回款" : "200.00",
      "回款率" : "73%",
      "备注" : "逾期"
    }, {
      "客户名称" : "苏州精密仪器",
      "合同金额" : "198.00",
      "已回款" : "198.00",
      "回款率" : "100%",
      "备注" : "已结清"
    }, {
      "客户名称" : "宁波海港物流",
      "合同金额" : "165.40",
      "已回款" : "120.00",
      "回款率" : "73%",
      "备注" : ""
    }, {
      "客户名称" : "合肥新能源",
      "合同金额" : "142.00",
      "已回款" : "100.00",
      "回款率" : "70%",
      "备注" : "逾期"
    }, {
      "客户名称" : "南京软件谷",
      "合同金额" : "128.60",
      "已回款" : "128.60",
      "回款率" : "100%",
      "备注" : "已结清"
    }, {
      "客户名称" : "温州轻工贸易",
      "合同金额" : "96.00",
      "已回款" : "60.00",
      "回款率" : "63%",
      "备注" : ""
    }, {
      "客户名称" : "无锡传感科技",
      "合同金额" : "88.50",
      "已回款" : "88.50",
      "回款率" : "100%",
      "备注" : "已结清"
    }, {
      "客户名称" : "芜湖装备制造",
      "合同金额" : "76.20",
      "已回款" : "40.00",
      "回款率" : "52%",
      "备注" : "逾期"
    }, {
      "客户名称" : "嘉兴纺织",
      "合同金额" : "65.00",
      "已回款" : "65.00",
      "回款率" : "100%",
      "备注" : "已结清"
    }, {
      "客户名称" : "徐州工程机械",
      "合同金额" : "54.30",
      "已回款" : "30.00",
      "回款率" : "55%",
      "备注" : ""
    }, {
      "客户名称" : "合计",
      "合同金额" : "2096.30",
      "已回款" : "1796.10",
      "回款率" : "86%",
      "备注" : ""
    } ],
    "source" : {
      "page" : 1,
      "lastPage" : 2
    }
  }, {
    "headers" : [ "工单类型", "工单数", "处理说明" ],
    "rows" : [ {
      "工单类型" : "硬件故障",
      "工单数" : "42",
      "处理说明" : ""
    }, {
      "工单类型" : "硬件故障",
      "工单数" : "12",
      "处理说明" : "更换配件后复测通过"
    }, {
      "工单类型" : "软件异常",
      "工单数" : "27",
      "处理说明" : "远程处理，平均响应 2 小时"
    }, {
      "工单类型" : "现场支持",
      "工单数" : "15",
      "处理说明" : "派工上门，覆盖江浙沪皖"
    } ],
    "source" : {
      "page" : 2
    }
  }, {
    "headers" : [ "故障部件", "工单数" ],
    "rows" : [ {
      "故障部件" : "读卡器",
      "工单数" : "18"
    }, {
      "故障部件" : "打印机",
      "工单数" : "15"
    }, {
      "故障部件" : "显示屏",
      "工单数" : "9"
    } ],
    "source" : {
      "page" : 2
    }
  }, {
    "headers" : [ "列1", "列2", "列3", "列4" ],
    "rows" : [ {
      "列1" : "2025 年度（已结账）",
      "列2" : "",
      "列3" : "",
      "列4" : ""
    } ],
    "source" : {
      "page" : 2
    }
  }, {
    "headers" : [ "片区", "年度目标", "完成率", "列4" ],
    "rows" : [ {
      "片区" : "华东大区",
      "年度目标" : "2400.00",
      "完成率" : "101%",
      "列4" : "达标"
    }, {
      "片区" : "华东南区",
      "年度目标" : "1500.00",
      "完成率" : "93%",
      "列4" : "未达标"
    }, {
      "片区" : "2026 年度（预算下达）",
      "年度目标" : "",
      "完成率" : "",
      "列4" : ""
    } ],
    "source" : {
      "page" : 2
    }
  }, {
    "headers" : [ "片区", "年度目标", "完成率", "列4" ],
    "rows" : [ {
      "片区" : "华东大区",
      "年度目标" : "2800.00",
      "完成率" : "—",
      "列4" : ""
    }, {
      "片区" : "华东南区",
      "年度目标" : "1750.00",
      "完成率" : "—",
      "列4" : ""
    } ],
    "source" : {
      "page" : 2
    }
  } ],
  "images" : [ {
    "fileName" : "image-1.png",
    "type" : "PNG",
    "width" : 1024,
    "height" : 768,
    "size" : 670939,
    "contentId" : "487e915db2c53907e8c1551aeaffea235423210c"
  } ]
}
```

### 9.4 pdf 的 6 组表（45 行）

`text` 是空串（0 段）；图片是 `image-1.jpg`，JPEG，699×524，38668 字节。

**#1 季度明细** 7 列 × 4 行，`{page=1}`

```text
headers = [片区, 区域, 销售额, 回款率, 销售额-5, 回款率-6, 备注]
1: 华东大区 | 上海 | 1280 | 92% | 1465 | 95% | 超额完成
2: 华东大区 | 江苏 | 986 | 88% | 1120 | 90% |
3: 华东南区 | 浙江 | 760 | 85% | 890 | 87% | 回款偏长
4: 华东南区 | 安徽 | 540 | 80% | 616 | 83% |
```

**#2 客户回款台账** 5 列 × 31 行，`{page=1, lastPage=2}`

```text
headers = [客户名称, 合同金额, 已回款, 回款率, 备注]
1: 上海浦东城建集团 | 486.00 | 486.00 | 100% | 已结清
2: 江苏恒通机械制造 | 320.50 | 280.00 | 87% |
3: 杭州云栖科技 | 275.80 | 200.00 | 73% | 逾期
4: 苏州精密仪器 | 198.00 | 198.00 | 100% | 已结清
5: 宁波海港物流 | 165.40 | 120.00 | 73% |
6: 合肥新能源 | 142.00 | 100.00 | 70% | 逾期
7: 南京软件谷 | 128.60 | 128.60 | 100% | 已结清
8: 温州轻工贸易 | 96.00 | 60.00 | 63% |
9: 无锡传感科技 | 88.50 | 88.50 | 100% | 已结清
10: 芜湖装备制造 | 76.20 | 40.00 | 52% | 逾期
11: 嘉兴纺织 | 65.00 | 65.00 | 100% | 已结清
12: 徐州工程机械 | 54.30 | 30.00 | 55% |
13: 常州轨道交通 | 48.20 | 48.20 | 100% | 已结清
14: 南通船舶重工 | 45.60 | 40.00 | 88% |
15: 绍兴黄酒集团 | 42.30 | 42.30 | 100% | 已结清
16: 台州汽配制造 | 39.80 | 30.00 | 75% | 逾期
17: 盐城风电设备 | 36.50 | 36.50 | 100% | 已结清
18: 扬州日化股份 | 33.20 | 25.00 | 75% |
19: 镇江铝业加工 | 30.60 | 30.60 | 100% | 已结清
20: 泰州医药城 | 28.40 | 20.00 | 70% | 逾期
21: 湖州建材贸易 | 26.10 | 26.10 | 100% | 已结清
22: 金华五金城 | 24.80 | 18.00 | 73% |
23: 连云港港口集团 | 22.50 | 22.50 | 100% | 已结清
24: 淮安电子科技 | 20.30 | 15.00 | 74% | 逾期
25: 宿迁板材加工 | 18.60 | 18.60 | 100% | 已结清
26: 丽水山货贸易 | 16.40 | 12.00 | 73% |
27: 衢州化工股份 | 15.20 | 15.20 | 100% | 已结清
28: 舟山渔业集团 | 13.80 | 10.00 | 72% | 逾期
29: 宣城机电制造 | 12.50 | 12.50 | 100% | 已结清
30: 池州矿业开发 | 11.20 | 8.00 | 71% |
31: 合计 | | | 86% |
```

**#3 售后工单统计** 3 列 × 4 行，`{page=2}`

```text
headers = [工单类型, 工单数, 处理说明]
1: 硬件故障 | 42 |
2: 硬件故障 | 12 | 更换配件后复测通过
3: 软件异常 | 27 | 远程处理，平均响应 2 小时
4: 现场支持 | 15 | 派工上门，覆盖江浙沪皖
```

**#4 分组标题** 4 列 × 1 行，`{page=3}`

```text
headers = [列1, 列2, 列3, 列4]
1: 2025 年度（已结账） | | |
```

**#5 年度目标（第一段）** 4 列 × 3 行，`{page=3}`

```text
headers = [片区, 年度目标, 完成率, 列4]
1: 华东大区 | 2400.00 | 101% | 达标
2: 华东南区 | 1500.00 | 93% | 未达标
3: 2026 年度（预算下达） | | |
```

**#6 年度目标（第二段）** 4 列 × 2 行，`{page=3}`

```text
headers = [片区, 年度目标, 完成率, 列4]
1: 华东大区 | 2800.00 | — |
2: 华东南区 | 1750.00 | — |
```

合计行两个金额为空，是 Word 导出 PDF 时 `SUM(ABOVE)` 的缓存结果丢失，
解析如实读出空串；要金额请用 docx。

### 9.5 pptx 的正文与 6 组表（22 行）

正文 9 段：

```text
[1] 2025年
[2] 销售业绩
[3] 华东大区
[4] 回款率
[5] 86%
[6] 达标
[7] 数据来源
[8] 销售系统
[9] 财务对账
```

**#1 季度明细** 7 列 × 4 行，`{page=1}`

```text
headers = [片区, 区域, 销售额, 回款率, 销售额-5, 回款率-6, 备注]
1: 华东大区 | 上海 | 1280 | 92% | 1465 | 95% | 超额完成
2: 华东大区 | 江苏 | 986 | 88% | 1120 | 90% |
3: 华东南区 | 浙江 | 760 | 85% | 890 | 87% | 回款偏长
4: 华东南区 | 安徽 | 540 | 80% | 616 | 83% |
```

**#2 客户回款台账（静默模式）** 5 列 × 8 行，`{page=2}`

```text
headers = [客户名称, 回款情况, 回款情况-3, 回款率, 备注]
1: 客户名称 | 合同金额 | 已回款 | 回款率 | 备注        ← 下层表头被当成数据
2: 上海城建 | 486 | 486 | 100% | 已结清
3: 江苏恒通 | 320 | 280 | 87% |
4: 杭州云栖 | 275 | 200 | 73% | 逾期
5: 苏州精密 | 198 | 198 | 100% | 已结清
6: 宁波海港 | 165 | 120 | 73% |
7: 合肥新能源 | 142 | 100 | 70% | 逾期
8: 合计 | 1586 | 1384 | 87% |
```

注册 `客户名称/合同金额/已回款/回款率/备注`（或上下两种写法各注册一次）后，
`headers` 变成这 5 个名字，第 1 行不再出现，共 7 行。

**#3 分组标题** 4 列 × 1 行，`{page=3}`

```text
headers = [列1, 列2, 列3, 列4]
1: 2025年 | | |
```

**#4 年度目标（第一段）** 4 列 × 3 行，`{page=3}`

```text
headers = [片区, 年度目标, 完成率, 列4]
1: 华东大区 | 2400 | 101% | 达标
2: 华东南区 | 1500 | 93% | 未达标
3: 2026年 | | |
```

**#5 年度目标（第二段）** 4 列 × 2 行，`{page=3}`

```text
headers = [片区, 年度目标, 完成率, 列4]
1: 华东大区 | 2800 | — |
2: 华东南区 | 1750 | — |
```

**#6 售后工单统计** 3 列 × 4 行，`{page=3}`

```text
headers = [工单类型, 工单数, 处理说明]
1: 硬件故障 | 42 |
2: 硬件故障 | 12 | 更换配件
3: 软件异常 | 27 | 远程处理
4: 现场支持 | 15 | 派工上门
```

pptx 样例把所有文字都压到 5 个字符以内（`更换配件`、`远程处理`、`派工上门`），
所以有没有许可证读到的都是同一串；评估模式只会截断超过 5 个字符的文本，且截掉的部分不可恢复。

### 9.6 图片四模式的真实字段（docx）

META（默认，4740 字符）：

```json
{
  "fileName" : "image-1.png",
  "type" : "PNG",
  "width" : 1024,
  "height" : 768,
  "size" : 670939,
  "contentId" : "487e915db2c53907e8c1551aeaffea235423210c"
}
```

BASE64（899370 字符，下面只截取前缀，实际字符串 894610 个字符）：

```json
{
  "fileName" : "image-1.png",
  "type" : "PNG",
  "width" : 1024,
  "height" : 768,
  "size" : 670939,
  "contentId" : "487e915db2c53907e8c1551aeaffea235423210c",
  "base64" : "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAABAAAAAMACAYAAAC..."
}
```

BYTES（899347 字符）：字段名是 `bytes`，JSON 里同样是一段 base64，但**没有**
`data:image/png;base64,` 前缀。

FILE（4832 字符）：

```json
{
  "fileName" : "image-1.png",
  "type" : "PNG",
  "width" : 1024,
  "height" : 768,
  "size" : 670939,
  "contentId" : "487e915db2c53907e8c1551aeaffea235423210c",
  "path" : "D:\\IdeaProjects\\document-parser\\target\\print\\images-docx\\image-1.png"
}
```

`path` 是本机实测值；你的环境会是 `imageDir` 的绝对路径。

pdf 的图片是 JPEG：38668 字节、699×524、
`contentId = 80c8e65be48c09a0f173845e4dfc072bc541fb68`；
BASE64 结果整串 58076 字符。pptx 与 docx 是同一份 PNG 字节，
BASE64 结果 898156 字符。

### 9.7 `unmatchedHeaders` 的真实 JSON

pdf 注册 `[故障部件, 工单数]` 并起名 `部件统计`，但 pdf 样例没有这张表：

```json
[ {
  "headers" : [ "故障部件", "工单数" ],
  "name" : "部件统计"
} ]
```

不起名时只有 `headers`：

```json
[ {
  "headers" : [ "故障部件", "工单数" ]
} ]
```

docx 有这张表（#4），所以清单是 `[]`。

## 10. 场景配方（可直接编译运行）

以下代码按 Java 8 语法写，`javac -source 8 -target 8` 可编译。
运行时把样例路径换成你自己的文件。

### 10.1 最小读取

```java
import cn.net.tianxi.document.WordParser;

import java.util.List;
import java.util.Map;

public class ReadOne {
    public static void main(String[] args) throws Exception {
        Map<String, Object> result = WordParser.parse("src/test/resources/word/sales-report.docx");

        String text = (String) result.get("text");
        List<Map<String, Object>> tables = asTables(result.get("tables"));

        System.out.println("正文段数 = " + text.split("\n", -1).length);
        System.out.println("表格组数 = " + tables.size());
        System.out.println("第一组列名 = " + tables.get(0).get("headers"));
        System.out.println("第一组行数 = " + asRows(tables.get(0).get("rows")).size());
    }

    @SuppressWarnings("unchecked")
    private static List<Map<String, Object>> asTables(Object value) {
        return (List<Map<String, Object>>) value;
    }

    @SuppressWarnings("unchecked")
    private static List<Map<String, String>> asRows(Object value) {
        return (List<Map<String, String>>) value;
    }
}
```

输出（默认选项）：

```text
正文段数 = 13
表格组数 = 7
第一组列名 = [片区, 区域, 销售额, 回款率, 销售额-5, 回款率-6, 备注]
第一组行数 = 4
```

### 10.2 按名字取台账，并做类型转换

```java
import cn.net.tianxi.document.WordParser;

import java.math.BigDecimal;
import java.util.List;
import java.util.Map;

public class ReadLedger {
    public static void main(String[] args) throws Exception {
        WordParser.Options options = WordParser.options()
                .headers("客户名称", "合同金额", "已回款", "回款率", "备注")
                .name("回款台账")
                .unmatchedHeaders(true);

        Map<String, Object> result = WordParser.parse(
                "src/test/resources/word/sales-report.docx", options);

        List<Map<String, String>> rows = WordParser.rowsByName(result, "回款台账");
        if (rows.isEmpty()) {
            System.out.println("没找到台账，未匹配的注册 = " + result.get("unmatchedHeaders"));
            return;
        }

        for (Map<String, String> row : rows) {
            String customer = row.get("客户名称");
            if ("合计".equals(customer)) {
                continue;                      // 合计行单独处理
            }
            BigDecimal contract = money(row.get("合同金额"));
            BigDecimal paid = money(row.get("已回款"));
            Integer rate = percent(row.get("回款率"));
            System.out.println(customer + " | " + contract + " | " + paid + " | " + rate + "%");
        }
    }

    private static BigDecimal money(String value) {
        String text = value == null ? "" : value.replace(",", "").trim();
        return text.length() == 0 ? null : new BigDecimal(text);
    }

    private static Integer percent(String value) {
        String text = value == null ? "" : value.trim();
        if (text.length() == 0) {
            return null;
        }
        if (text.endsWith("%")) {
            text = text.substring(0, text.length() - 1);
        }
        return Integer.valueOf(text);
    }
}
```

docx 实测输出前两行：

```text
上海浦东城建集团 | 486.00 | 486.00 | 100%
江苏恒通机械制造 | 320.50 | 280.00 | 87%
```

注意：解析器不做类型转换，`486.00` 是字符串；`money` / `percent` 是调用方自己的转换。
pdf 的合计金额是空串，`money` 会返回 `null`，调用方要处理。

### 10.3 同一份注册跨三种格式

```java
import cn.net.tianxi.document.PdfParser;
import cn.net.tianxi.document.PptParser;
import cn.net.tianxi.document.WordParser;

import java.util.List;
import java.util.Map;

public class CrossFormat {
    public static void main(String[] args) throws Exception {
        WordParser.Options options = WordParser.options()
                // 两行表头合并之后的那一行
                .headers("客户名称", "合同金额", "已回款", "回款率", "备注")
                .name("回款台账")
                // 文档里真实的上层那一行（跨列格写两遍），同一个名字
                .headers("客户名称", "回款情况", "回款情况", "回款率", "备注")
                .name("回款台账");

        print("docx", WordParser.parse("src/test/resources/word/sales-report.docx", options));
        print("pdf", PdfParser.parse("src/test/resources/pdf/sales-report.pdf", options));
        print("pptx", PptParser.parse("src/test/resources/ppt/sales-report.pptx", options));
    }

    private static void print(String format, Map<String, Object> result) {
        List<Map<String, String>> rows = WordParser.rowsByName(result, "回款台账");
        System.out.println(format + " 台账 " + rows.size() + " 行");
        if (!rows.isEmpty()) {
            System.out.println("  第一行 = " + rows.get(0));
        }
    }
}
```

实测：

```text
docx 台账 13 行
pdf 台账 31 行
pptx 台账 7 行
```

三份样例客户数本来就不同；组数与总行数分别是 7/30、6/45、6/21，没有多出组。

### 10.4 图片导出：FILE 与 BASE64 两种

```java
import cn.net.tianxi.document.WordParser;

import java.io.File;
import java.util.Map;

public class ExportImages {
    public static void main(String[] args) throws Exception {
        // 方式一：解析时直接落盘
        Map<String, Object> fileResult = WordParser.parse(
                "src/test/resources/word/sales-report.docx",
                WordParser.options()
                        .images(WordParser.ImageMode.FILE)
                        .imageDir(new File("out/images-file")));
        System.out.println(fileResult.get("images"));

        // 方式二：先拿 BASE64，再统一导出
        Map<String, Object> base64Result = WordParser.parse(
                "src/test/resources/word/sales-report.docx",
                WordParser.options().images(WordParser.ImageMode.BASE64));
        WordParser.writeImages(base64Result, new File("out/images-base64"));
        System.out.println("已导出到 out/images-base64");
    }
}
```

`FILE` 的结果里图片是：

```text
{fileName=image-1.png, type=PNG, width=1024, height=768, size=670939,
 contentId=487e915db2c53907e8c1551aeaffea235423210c, path=...\image-1.png}
```

`writeImages` 只认 `base64` 字段；默认 `META` 结果调用它不会写文件，会在标准错误提示。

### 10.5 先看陌生文档长什么样

```java
import cn.net.tianxi.document.WordParser;

import java.util.List;
import java.util.Map;

public class Explore {
    public static void main(String[] args) throws Exception {
        Map<String, Object> result = WordParser.parse(args[0]);
        List<Map<String, Object>> tables = asTables(result.get("tables"));
        for (int i = 0; i < tables.size(); i++) {
            Map<String, Object> table = tables.get(i);
            System.out.println("#" + (i + 1) + " " + table.get("headers")
                    + " rows=" + asRows(table.get("rows")).size()
                    + " source=" + table.get("source"));
        }
    }

    @SuppressWarnings("unchecked")
    private static List<Map<String, Object>> asTables(Object value) {
        return (List<Map<String, Object>>) value;
    }

    @SuppressWarnings("unchecked")
    private static List<Map<String, String>> asRows(Object value) {
        return (List<Map<String, String>>) value;
    }
}
```

先跑一遍拿到真实 `headers`，再把它们写进 `headers(...)` 注册并起名。不要凭肉眼猜表头：
合并单元格、两行表头、末尾空列都会改变最终列名。

### 10.6 判断文档里有没有某张表

```java
import cn.net.tianxi.document.PdfParser;
import cn.net.tianxi.document.WordParser;

import java.util.Map;

public class Exists {
    public static void main(String[] args) throws Exception {
        WordParser.Options options = WordParser.options()
                .headers("故障部件", "工单数")
                .name("部件统计")
                .unmatchedHeaders(true);

        Map<String, Object> result = PdfParser.parse(
                "src/test/resources/pdf/sales-report.pdf", options);

        boolean present = !WordParser.tablesByName(result, "部件统计").isEmpty();
        System.out.println("present = " + present);
        System.out.println("unmatchedHeaders = " + result.get("unmatchedHeaders"));
    }
}
```

pdf 实测（标准错误那一行会先于标准输出打印）：

```text
[PdfParser] 表头集合未匹配到任何行：[故障部件, 工单数]
present = false
unmatchedHeaders = [ {headers=[故障部件, 工单数], name=部件统计} ]
```

用 `tablesByName` 判断比看 `unmatchedHeaders` 更直接；后者是给告警和日志用的。

### 10.7 从上传流读取

```java
import cn.net.tianxi.document.PdfParser;
import cn.net.tianxi.document.WordParser;

import java.io.FileInputStream;
import java.io.InputStream;
import java.util.Map;

public class ReadStream {
    public static void main(String[] args) throws Exception {
        try (InputStream in = new FileInputStream("src/test/resources/pdf/sales-report.pdf")) {
            Map<String, Object> result = PdfParser.parse(in, WordParser.options()
                    .headers("客户名称", "合同金额", "已回款", "回款率", "备注")
                    .name("回款台账"));
            System.out.println(WordParser.rowsByName(result, "回款台账").size());
        }
    }
}
```

流由调用方关闭；解析器不会关它。上传接口拿到 `MultipartFile` 时用
`multipartFile.getInputStream()` 传进去即可。

## 11. 异常与标准错误

### 11.1 缺文件：三种格式的异常类型不一样（1.2.1 实测）

| 入口 | 异常类型 | 消息 |
| --- | --- | --- |
| `WordParser.parse("缺失.docx")` | `java.io.FileNotFoundException` | `Can't find file: ...` |
| `PdfParser.parse("缺失.pdf")` | `com.aspose.pdf.internal.ms.System.l5if` | 内部异常包着 `Can't find file: ...`，消息含内部类名与堆栈片段 |
| `PptParser.parse("缺失.pptx")` | `com.aspose.slides.exceptions.FileNotFoundException` | `Can't find file: ...` |

PDF 抛的是 Aspose 内部混淆类名，**不要按具体类型 catch**，按 `Exception` 兜底。

### 11.2 选项组合错误

`ImageMode.FILE` 但没给 `imageDir`：解析到第一张图时抛

```text
java.lang.IllegalStateException: ImageMode.FILE 需要先调用 imageDir(File) 指定导出目录
```

三种格式一致。

### 11.3 标准错误提示（不抛异常）

| 场景 | 提示开头 |
| --- | --- |
| 注册表头未命中任何行 | `[WordParser]` / `[PdfParser]` / `[PptParser]` + `表头集合未匹配到任何行：[...]` |
| 没有注册就调用 `name(...)` | `[name] 还没有注册任何表头集合，这一次 ... 没有归属，已忽略` |
| `META` 模式却给了 `imageDir` | `[WordParser] imageDir 只在 ImageMode.FILE 下生效，当前是 META，... 不会写入图片`（三种入口共用 `Images`，前缀固定 `[WordParser]`） |
| `META` 结果调用 `writeImages` | `[writeImages] 结果里的图片没有 base64 内容，一张都没导出 —— ...` |
| 表格嵌套超过 8 层 | 每张表最多提示一次，然后跳过更深的子表 |
| 容器嵌套超过 16 层 | 每个被跳过的分支各提示一次 |

这些提示不是异常，结果照给。生产环境建议打开 `unmatchedHeaders(true)` 把其中最关键的一条
（注册表头没命中）变成程序可读的字段。

## 12. 自己跑一遍验收

### 12.1 全量测试

```bash
mvn -B test
```

期望：`Tests run: 242, Failures: 0, Errors: 0, Skipped: 1`，`BUILD SUCCESS`，
15 个测试类。跳过的那条是需要 `-Ddoc` 的打印用例。

### 12.2 打印真实结果（18 份报告）

```bash
mvn -B test -Dtest=PrintResultTest
```

报告写到 `target/print/`，本次跑出 18 份：

| 报告 | 看什么 |
| --- | --- |
| `sales-report.docx-静默.txt` / `.pdf-静默.txt` / `.pptx-静默.txt` | 默认 7/30、6/45、6/22；全部组与行 |
| `sales-report.*-注册表头-起名.txt` | `name` 写在哪些组、各组行数 |
| `sales-report.docx-注册表头-DROP.txt` | `fullRow` 的去向 |
| `公共入口-四个重载.txt` | 路径/流、关键字/选项的 12 组对比 |
| `选项默认值与setter.txt` | `Options` 默认值与 `name` / 容忍度行为 |
| `选项矩阵-三份样例.txt` | KEEP/FIRST/DROP、source、unmatchedHeaders 的字符数 |
| `按名字取数-三份样例.txt` | 13 / 31 / 7 行、5 行年度目标、空名字返回 `[]` |
| `未匹配表头-三份样例.txt` | docx 空清单、pdf/pptx 1 组 |
| `表头补名与域缓存.txt` | `销售额-5`、`列4`、`SUM` 缓存 300.00 |
| `docx-台账四种注册写法.txt` | 四种写法都命中 13 行 |
| `docx-图片四种模式.txt` | 4740 / 899347 / 899370 / 4832 |
| `序列化与图片导出.txt` | 两个 `toJson` 逐字符相同、落盘 SHA-1 |
| `全部模式-完整JSON.txt` | 各选项组合的完整 JSON |
| `错误路径-三份样例.txt` | 三种异常类型与 FILE 缺目录 |

### 12.3 检查点

1. 静默模式的报告里“标准错误”应为空；注册了未命中的表头（pdf/pptx 各一条）或
   `META + imageDir` 时会多出提示，属于预期；
2. 每组的 `headers` 与每行键必须逐个相等（`PrintResultTest` 用 `assertHeadersMatchRowKeys` 守着）；
3. 同一份文件跑两次，报告 md5 应一致（本次实测一致）。

报告文件名是中文，Windows 控制台若显示乱码，是编码问题，不是解析问题；文件本身是 UTF-8。

## 13. 已知限制与不承诺的行为

- **不做 OCR**：扫描件、图片型 PDF 读不出文字；需要 OCR 请先过 OCR 引擎。
- **Word 不读页眉页脚、脚注尾注、批注**；PPT 不读备注页、母版、版式页、SmartArt 文字、
  图表数据、OLE 对象内部内容；PDF 不读无边框且每行单元格数不一致的表。
- **不重算域**：只读 `fldChar separate` 与 `end` 之间的缓存结果；缓存为空就给空串。
  导出成 PDF 后域结果可能丢失（见 9.4 的合计行）。
- **不保证 `name` 唯一**：同一张物理表切出的多个组可以同名。
- **容忍度会误命中**：放宽 `headerTolerance` 后，名字也可能绑定到相近的表头。
- **`keywords` 是宽松匹配**：传了关键字就关闭首行兜底；pdf 上实测两张表的表头行落进数据
  （45 → 47 行），入库请用完整表头注册。
- **评估模式会截断 PPT 文本**（每段 5 字符）且不可恢复；生产环境请配置许可证。
- **列名后缀是绝对列号**：`销售额-5` 的 5 是物理第 5 列，不是“第 5 个重复”。
- **`source` 可能缺失**：某格式给不出位置时整键不写；三份样例都有，但不要假设永远有。
- **所有表格值都是字符串**：不做数字、日期、百分比转换；空单元格是空串。
- **不提供自动格式识别**：调用方按扩展名/内容自己选 `WordParser` / `PdfParser` / `PptParser`；
  扩展名与实际内容不符时由 Aspose 抛异常。
- **性能数字不是承诺**：5.9 的耗时是特定机器上的区间，只用于比较 `source` 开关的相对影响。

## 14. 速查卡

### 14.1 方法签名

```java
// 三个解析器都有
Map<String, Object> parse(String filePath, String... headerKeywords) throws Exception
Map<String, Object> parse(InputStream inputStream, String... headerKeywords) throws Exception
Map<String, Object> parse(String filePath, WordParser.Options options) throws Exception
Map<String, Object> parse(InputStream inputStream, WordParser.Options options) throws Exception
String toJson(Object value) throws Exception
void writeImages(Map<String, Object> result, File directory) throws Exception

// 只有 WordParser 有；PDF/PPT 的结果也能传给这两个方法
WordParser.Options options()
List<Map<String, Object>> tablesByName(Map<String, Object> result, String name)
List<Map<String, String>> rowsByName(Map<String, Object> result, String name)
```

### 14.2 Options 一览

```java
WordParser.options()
        .headers("列1", "列2")              // 注册一组完整表头
        .headerGroups(groups)               // 批量注册
        .name("组名")                        // 给最近注册的一组起名
        .headerTolerance(1)                 // 允许 1 列不一致
        .keywords("销售", "回款")            // 关键字兜底
        .fullRow(WordParser.FullRow.KEEP)   // 分组行去向
        .images(WordParser.ImageMode.META)  // 图片输出方式
        .imageDir(new File("out/images"))   // FILE 模式目录
        .source(true)                       // 带 page / lastPage
        .unmatchedHeaders(true);            // 顶层未匹配清单
```

### 14.3 结果键

```text
顶层：text, tables, images, [unmatchedHeaders]
组：  headers, rows, [name], [source]
行：  键 = headers，值 = 字符串
图片：fileName, type, width, height, [size, contentId],
      [base64 | bytes | path | duplicateOf | link]
未匹配：headers, [name]
```

### 14.4 默认值

```text
fullRow = KEEP
headerTolerance = 1
images = META
imageDir = null
source = true
unmatchedHeaders = false
```

---

如果发现本文与实测不一致，以 `target/print/` 的报告与测试断言为准，并请提 issue：
https://github.com/hutianxi/document-parser/issues
