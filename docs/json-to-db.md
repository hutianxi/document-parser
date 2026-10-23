# 从解析结果 JSON 到数据库：映射与批量入库示例

本文回答一个具体问题：拿到 `toJson` 的输出后，怎样把 `tables` 里的一张张表**稳定地**落到数据库，
而不是靠表在数组里的位置硬编码。

示例代码是独立片段，**没有加入本库的构建与测试**；数据库方言、驱动、连接池、事务边界按你的工程调整。
本文也不讨论图片入库、正文分段入库和 OCR——那些不在 `tables` 的契约里。

---

## 1. 先看清 JSON 的实际形状

默认选项（`META`、`source(true)`）下，一份结果的骨架是：

```json
{
  "text" : "……",
  "tables" : [ {
    "headers" : [ "客户名称", "合同金额", "已回款", "回款率", "备注" ],
    "rows" : [ {
      "客户名称" : "上海浦东城建集团",
      "合同金额" : "486.00",
      "已回款" : "486.00",
      "回款率" : "100%",
      "备注" : "已结清"
    } ],
    "name" : "回款台账",
    "source" : { "page" : 1, "lastPage" : 2 }
  }, {
    "headers" : [ "工单类型", "工单数", "处理说明" ],
    "rows" : [ {
      "工单类型" : "硬件故障",
      "工单数" : "42",
      "处理说明" : ""
    } ],
    "source" : { "page" : 2 }
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

这张骨架来自仓库里的 docx 样例（这里只摘了台账与工单两组），真实副本见
`target/print/sales-report.docx-注册表头-起名.txt`
（把 `name` 去掉就是静默输出的形状；`text` 用省略号代替，其余字段取自实测）。
静默与起名两版报告都由 `mvn -B test -Dtest=PrintResultTest` 生成。需要先明确的四件事：

1. **每一行的键与 `headers` 完全一致**，顺序也一致。重名列已经被解析器补成 `销售额-5`、
   `回款率-6`（绝对列号），空列名补成 `列N`，所以每行不会缺列。补名只追加一次列号，
   极端构造下仍可能撞名，见 §10 的最后一条。
2. **`name` 是可选的**。只有「注册了这组表头 + 注册时起了名 + 文档里确实命中」三者都成立时才有；
   上面的工单表没有注册，所以没有 `name`。
3. **`source` 也是可选的**。调用 `source(false)` 后每组表都没有这个键。
4. **`unmatchedHeaders` 是可选的顶层键**，只在 `unmatchedHeaders(true)` 时出现，且只反映生产端的注册情况，
   不能替代消费端的表路由。

`name` 还有一个容易忽略的性质：**它不是主键**。同一张物理表切出的两个表头组可以带同一个名字——
docx 命名轮的第 6、7 组都叫 `年度目标`（3 行 + 2 行），见
`target/print/sales-report.docx-注册表头-起名.txt`。

---

## 2. 为什么不要按 `name` 建索引

如果消费端写成下面这样，问题不会在 JSON 解析时报出来，而是在数据上悄悄发生：

```java
Map<String, JsonNode> byName = new HashMap<>();
for (JsonNode table : root.get("tables")) {
    byName.put(table.get("name").asText(), table);   // 未命名组：NullPointerException
}
```

用 Jackson 2.13.5 对「1 个命名组 + 2 个未命名组」实测，几种常见读法的结果如下：

| 读法 | 结果 |
| --- | --- |
| `JsonNode.has("name")` | `false`，不抛异常 |
| `JsonNode.get("name")` | `null`，不抛异常 |
| `JsonNode.path("name").asText("<absent>")` | `<absent>`，不抛异常 |
| 反序列化成 `Map` | `containsKey("name") == false`，不抛异常 |
| 反序列化成含 `String name` 的 POJO | `name == null`，不抛异常 |
| `get("name").asText()` | `NullPointerException` |
| `@JsonCreator` 上 `required = true` | `MismatchedInputException`：缺少 `name` |
| 按 `name` 放进 `HashMap` | 3 张表只剩 2 个键，未命名组都落在 `null` 上互相覆盖 |
| 按 `name` 放进 `ConcurrentHashMap` | `NullPointerException` |
| 按数组下标遍历 | 3 张表全部保留 |

结论：**JSON 本身是合法的，异常来自消费端的假设**。`name` 适合当日志标签、可读标识，
不适合当 `Map` 的键，更不适合当数据库主键。即便所有组都起了名，同名组仍会互相覆盖。

---

## 3. 推荐的路由键：表头签名

`headers` 是每张表一定有的字段，而且顺序稳定。把 `headers` 序列化成规范字符串，作为「表结构签名」：

```java
String signature(List<String> headers) throws JsonProcessingException {
    return new ObjectMapper().writeValueAsString(headers);
}
```

这样得到的是一段规范 JSON（例如 `["客户名称","合同金额","已回款","回款率","备注"]`），
不会因为表头里恰好含分隔符而碰撞。不要用 `String.join("|", headers)`：表头本身可能包含 `|`，
`["a","b|c"]` 与 `["a|b","c"]` 会得到同一个键。

签名是**顺序敏感**的：同一组列名换一个顺序，会被当成另一种表。这是有意的——入库的列顺序、
类型转换规则都跟着表结构走；如果源文档可能重排同一组列，按列名绑定（命名参数）即可，
但签名仍建议保留顺序，避免两种语义混在一个映射里。

把签名映射到「目标表 + 列映射」的注册表：

```java
Map<String, TableSpec> specs = new ConcurrentHashMap<>();
specs.put(signature(spec.sourceHeaders), spec);
```

每张表只查一次；命中后所有行共用同一份映射。未命中的表进死信，不要猜、不要静默丢。

---

## 4. 一份可落地的实现

下面是一个不依赖 Spring、只用 JDK + Jackson + JDBC 的最小实现。Jackson 已经在解析库的依赖里；
JDBC 驱动由你的工程提供。

### 4.1 列类型与转换

表格行里的值**全部是字符串**：金额是 `"2096.30"`，比率是 `"86%"`，空单元格是 `""`。
类型转换必须由消费端按列配置，解析器不推断类型。

```java
import java.math.BigDecimal;

enum ValueType {
    TEXT {
        @Override Object convert(String raw) {
            return raw;                       // 空串保留为空串
        }
    },
    DECIMAL {
        @Override Object convert(String raw) {
            String s = raw.trim();
            return s.isEmpty() ? null : new BigDecimal(s);   // 空串 -> NULL
        }
    },
    PERCENT {
        @Override Object convert(String raw) {
            String s = raw.trim();
            if (s.isEmpty()) {
                return null;
            }
            if (!s.endsWith("%")) {
                throw new IllegalArgumentException("缺少 % 号: " + raw);
            }
            return new BigDecimal(s.substring(0, s.length() - 1).trim()).movePointLeft(2);
        }
    },
    INTEGER {
        @Override Object convert(String raw) {
            String s = raw.trim();
            return s.isEmpty() ? null : new BigDecimal(s).intValueExact();
        }
    };

    abstract Object convert(String raw);
}
```

几个刻意的选择：

- `TEXT` 把空串保留为 `""`，`DECIMAL` / `PERCENT` / `INTEGER` 把空串转成 `null`。
  PDF 导出的台账合计行里，`合同金额` 和 `已回款` 实测就是空串；直接 `new BigDecimal("")`
  会抛 `NumberFormatException`。
- `PERCENT` 把 `"86%"` 转成 `0.86`。如果数据库按 0–100 存，去掉 `movePointLeft(2)` 即可，
  但要在整个工程里统一，不能两种口径混用。
- 不自动去掉千分位逗号。`"1,234.50"` 会直接报错；如果源文档确实这么写，在 `DECIMAL` 里显式
  `replace(",", "")`，并同时补测试——不要默默容错。
- 数字、日期、编码类字段（邮编、工号）不要靠「看起来像数字」推断，全部按列显式配置。

### 4.2 目标表与列映射

```java
import java.util.ArrayList;
import java.util.Collections;
import java.util.HashSet;
import java.util.List;
import java.util.Set;

final class ColumnSpec {
    final String source;      // JSON 里的列名
    final String target;      // 数据库列名（白名单）
    final ValueType type;

    ColumnSpec(String source, String target, ValueType type) {
        this.source = source;
        this.target = identifier(target);
        this.type = type;
    }

    static String identifier(String value) {
        if (!value.matches("[A-Za-z_][A-Za-z0-9_]*")) {
            throw new IllegalArgumentException("不是安全标识符: " + value);
        }
        return value;
    }
}

final class TableSpec {
    final List<String> sourceHeaders;
    final String table;
    final List<ColumnSpec> columns;
    final String sql;

    TableSpec(String table, List<String> sourceHeaders, List<ColumnSpec> columns) {
        if (sourceHeaders.isEmpty() || columns.isEmpty()) {
            throw new IllegalArgumentException("表头与列映射不能为空");
        }
        this.table = ColumnSpec.identifier(table);
        this.sourceHeaders = Collections.unmodifiableList(new ArrayList<String>(sourceHeaders));
        this.columns = Collections.unmodifiableList(new ArrayList<ColumnSpec>(columns));

        Set<String> seen = new HashSet<String>();
        StringBuilder names = new StringBuilder();
        StringBuilder marks = new StringBuilder();
        for (ColumnSpec column : this.columns) {
            if (!seen.add(column.target)) {
                throw new IllegalArgumentException("目标列重复: " + column.target);
            }
            if (names.length() > 0) {
                names.append(", ");
                marks.append(", ");
            }
            names.append(column.target);
            marks.append('?');
        }
        this.sql = "INSERT INTO " + this.table + " (" + names + ") VALUES (" + marks + ")";
    }
}
```

`table` 和 `target` 都会走标识符白名单。**不要**把 `headers` 直接拼进 SQL：
表头来自文档，`备注; DROP TABLE ...` 这种内容一旦进了 SQL 就是注入面。目标表名和目标列名只能来自
你自己的配置，不能来自解析结果。

### 4.3 导入器

```java
import com.fasterxml.jackson.core.JsonProcessingException;
import com.fasterxml.jackson.core.type.TypeReference;
import com.fasterxml.jackson.databind.JsonNode;
import com.fasterxml.jackson.databind.ObjectMapper;

import java.io.IOException;
import java.sql.Connection;
import java.sql.PreparedStatement;
import java.sql.SQLException;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;
import java.util.concurrent.ConcurrentHashMap;
import java.util.function.Consumer;

final class JsonToDb {
    private final ObjectMapper mapper = new ObjectMapper();
    private final Map<String, TableSpec> specs = new ConcurrentHashMap<String, TableSpec>();
    private final Consumer<String> deadLetter;
    private final int batchSize;

    JsonToDb(Consumer<String> deadLetter, int batchSize) {
        this.deadLetter = deadLetter;
        this.batchSize = batchSize;
    }

    void register(TableSpec spec) throws JsonProcessingException {
        specs.put(signature(spec.sourceHeaders), spec);
    }

    void importJson(String json, Connection connection) throws IOException, SQLException {
        JsonNode root = mapper.readTree(json);
        JsonNode tables = root.path("tables");
        if (!tables.isArray()) {
            deadLetter.accept("结果里没有 tables 数组");
            return;
        }
        for (int i = 0; i < tables.size(); i++) {
            JsonNode table = tables.get(i);
            List<String> headers;
            try {
                headers = mapper.convertValue(table.get("headers"),
                        new TypeReference<List<String>>() { });
            } catch (RuntimeException e) {
                deadLetter.accept(describe("headers 不是字符串数组: " + e.getMessage(), i, table, null));
                continue;
            }

            TableSpec spec = specs.get(signature(headers));
            if (spec == null) {
                deadLetter.accept(describe("未注册的表头", i, table, headers));
                continue;
            }

            List<Object[]> batch = readRows(spec, table, i);
            if (batch != null) {
                write(connection, spec, batch);
            }
        }
    }

    private String signature(List<String> headers) throws JsonProcessingException {
        return mapper.writeValueAsString(headers);
    }

    private List<Object[]> readRows(TableSpec spec, JsonNode table, int tableIndex) {
        JsonNode rows = table.path("rows");
        if (!rows.isArray()) {
            deadLetter.accept(describe("rows 不是数组", tableIndex, table, null));
            return null;
        }
        List<Object[]> batch = new ArrayList<Object[]>(rows.size());
        for (int r = 0; r < rows.size(); r++) {
            JsonNode row = rows.get(r);
            Object[] values = new Object[spec.columns.size()];
            for (int c = 0; c < spec.columns.size(); c++) {
                ColumnSpec column = spec.columns.get(c);
                String raw = row.path(column.source).asText("");
                try {
                    values[c] = column.type.convert(raw);
                } catch (RuntimeException e) {
                    // 整张表不入库，避免部分提交
                    deadLetter.accept(describe("第 " + (r + 1) + " 行 " + column.source
                            + " 转换失败: " + e.getMessage(), tableIndex, table, null));
                    return null;
                }
            }
            batch.add(values);
        }
        return batch;
    }

    private void write(Connection connection, TableSpec spec, List<Object[]> batch)
            throws SQLException {
        try (PreparedStatement ps = connection.prepareStatement(spec.sql)) {
            int pending = 0;
            for (Object[] values : batch) {
                for (int i = 0; i < values.length; i++) {
                    ps.setObject(i + 1, values[i]);
                }
                ps.addBatch();
                if (++pending == batchSize) {
                    ps.executeBatch();
                    pending = 0;
                }
            }
            if (pending > 0) {
                ps.executeBatch();
            }
        }
    }

    private String describe(String reason, int index, JsonNode table, List<String> headers) {
        return "table#" + index
                + " name=" + table.path("name").asText("<unnamed>")
                + " page=" + table.path("source").path("page").asText("?")
                + " headers=" + (headers != null ? headers : table.path("headers"))
                + " reason=" + reason;
    }
}
```

这个实现里有三处是刻意写死的：

- **转换阶段一张表要么整表入库，要么整表进死信**。任何一行转换失败就 `return null`，不写入半张表；
  `write` 中途的 `SQLException` 不伪装成死信，交给外层事务回滚。
  对财务、台账类数据，部分提交比失败更难排查。
- **表头只用来查映射，不用来拼 SQL**。SQL 在 `TableSpec` 构造时由白名单列名生成一次。
- **每张表只做一次签名查找**，行内不再比较表头；每行只做按列转换和参数绑定。

事务边界由调用方控制：

```java
connection.setAutoCommit(false);
try {
    importer.importJson(json, connection);
    connection.commit();
} catch (IOException | RuntimeException | SQLException e) {
    connection.rollback();
    throw e;
}
```

如果业务允许「坏表跳过、好表提交」，就把 `deadLetter` 换成持久化的死信表，并接受
`importJson` 里已写入的表；两种策略不要混着用。

---

## 5. 注册三份样例的映射

下面这张表来自 `target/print/` 下三份样例的静默报告（`sales-report.{docx,pdf,pptx}-静默.txt`）
与起名报告（`-注册表头-起名.txt`），行数是实测值：

| 文档里的表 | `headers` | 建议目标表 | 行数 |
| --- | --- | --- | --- |
| 季度销售明细 | `片区, 区域, 销售额, 回款率, 销售额-5, 回款率-6, 备注` | `sales_detail` | 4 |
| 客户回款台账（三格式，注册合并后表头） | `客户名称, 合同金额, 已回款, 回款率, 备注` | `payment_ledger` | docx 13 / pdf 31 / pptx 7 |
| 同一台账的 pptx 静默表头（仅上层） | `客户名称, 回款情况, 回款情况-3, 回款率, 备注` | 不建议直接入库 | pptx 8（其中 1 行是伪表头） |
| 售后工单统计 | `工单类型, 工单数, 处理说明` | `work_order` | 4 |
| 故障部件子表（仅 docx） | `故障部件, 工单数` | `work_order_part` | 3 |
| 分片区年度目标 | `片区, 年度目标, 完成率, 列4` | `annual_target` | 3 + 2（两组同名表头） |

注册代码：

```java
importer.register(new TableSpec("payment_ledger",
        Arrays.asList("客户名称", "合同金额", "已回款", "回款率", "备注"),
        Arrays.asList(
                new ColumnSpec("客户名称", "customer_name", ValueType.TEXT),
                new ColumnSpec("合同金额", "contract_amount", ValueType.DECIMAL),
                new ColumnSpec("已回款", "received_amount", ValueType.DECIMAL),
                new ColumnSpec("回款率", "received_rate", ValueType.PERCENT),
                new ColumnSpec("备注", "remark", ValueType.TEXT))));

importer.register(new TableSpec("work_order",
        Arrays.asList("工单类型", "工单数", "处理说明"),
        Arrays.asList(
                new ColumnSpec("工单类型", "ticket_type", ValueType.TEXT),
                new ColumnSpec("工单数", "ticket_count", ValueType.INTEGER),
                new ColumnSpec("处理说明", "resolution", ValueType.TEXT))));
```

几点说明：

- **三份台账用合并后的 5 列注册一次即可**：docx / pdf / pptx 都命中，行数分别是 13 / 31 / 7。
  pptx 静默输出里的 `回款情况` 两列只是上层表头；只按那组签名注册时，下层表头会漏成 1 行数据
  （`客户名称=客户名称`、`回款情况=合同金额` …），行数变成 8，不建议直接入库。
- **年度目标的两组表头完全一样**，注册一次即可；两组都会命中并各插各的行。如果它们逻辑上要落到
  不同目标表，就不能只靠表头签名，需要再加 `source.page` 或出现序号做二级键。
- **`故障部件` 子表只在 docx 里有**；pdf/pptx 没有这张表，所以只注册它在 pdf/pptx 上不会命中，
  也不需要为它准备空表。

---

## 6. 命名参数与 MyBatis

上面的例子用 `PreparedStatement` 的 `?` 占位，所以有一层按列绑定的循环。如果你用命名参数，
这层循环可以交给框架。

Spring `NamedParameterJdbcTemplate`：

```java
String sql = "INSERT INTO payment_ledger "
        + "(customer_name, contract_amount, received_amount, received_rate, remark) "
        + "VALUES (:customer_name, :contract_amount, :received_amount, :received_rate, :remark)";

List<Map<String, Object>> rows = new ArrayList<Map<String, Object>>();
for (JsonNode row : table.get("rows")) {
    Map<String, Object> params = new LinkedHashMap<String, Object>();
    params.put("customer_name", row.path("客户名称").asText(""));
    params.put("contract_amount", ValueType.DECIMAL.convert(row.path("合同金额").asText("")));
    params.put("received_amount", ValueType.DECIMAL.convert(row.path("已回款").asText("")));
    params.put("received_rate", ValueType.PERCENT.convert(row.path("回款率").asText("")));
    params.put("remark", row.path("备注").asText(""));
    rows.add(params);
}
namedJdbc.batchUpdate(sql, rows.toArray(new Map[0]));
```

MyBatis 的 `<foreach>` 也能把行列表直接展开；**列名和表名只能用 `<include>` 或白名单拼进去，
不要用 `${headers}`**。`#{}` 只用于值。

如果目标列名恰好与 JSON 表头一致，且不需要类型转换，可以整行 Map 直接批量绑定；
但只要有一个列需要改名或转类型，就老老实实显式构造参数——比在 SQL 里做字符串替换安全得多。

---

## 7. 同一个 JVM 里，可以不过 JSON

解析和入库在同一个进程时，JSON 是一次不必要的往返。直接拿解析结果：

```java
Map<String, Object> result = WordParser.parse(path, options);
List<Map<String, Object>> tables = WordParser.tablesByName(result, "回款台账");
for (Map<String, Object> table : tables) {
    List<Map<String, String>> rows =
            (List<Map<String, String>>) (List<?>) table.get("rows");
    // 行仍然是「列名 -> 字符串」，直接按上面的 TableSpec 绑定
}
```

- `tablesByName` 返回同名各组的**同一批对象**，按 `tables` 原顺序；
- `rowsByName` 把同名各组的行拼平，适合「同名表就是要一起入库」的场景；
- 两条入口对未注册、空串、`null` 的名字都给空列表，不抛异常；
- 这样省掉一次 `toJson` 和一次 `readTree`，也避免把大图片的 BASE64 写进字符串。

只有跨服务、落文件、进消息队列时，才需要 JSON。

---

## 8. 生产端怎么配合

消费端的表头签名越稳定，入库越省事。生产端建议：

1. **明确要入库的表，注册完整表头并起名**：

   ```java
   WordParser.options()
           .headers("客户名称", "合同金额", "已回款", "回款率", "备注").name("回款台账")
           .headers("工单类型", "工单数", "处理说明").name("工单统计");
   ```

2. **开发/预发阶段打开 `unmatchedHeaders(true)`**，把注册了但没命中的表头打到顶层
   `unmatchedHeaders` 里；它能在发布前抓住「表头写错一个列名」「跨格式列名不同」这类问题。
   注意它只是生产端诊断，不能替代消费端的死信。
3. **跨格式共用一份注册时，把上下两种表头写法各注册一次、起同一个名字**。多级表头的
   上下两行在不同格式里给的行不一样，只注册上层可能在 pdf 上命不中、在 pptx 上把下层表头
   漏进数据（实测 docx 13 行、pptx 8 行、pdf 命不中）。
4. **不要在运行时关掉 `source` 又把页码当日志字段**。`source(false)` 会省掉 docx 的版面计算，
   但结果里就没有 `page` 了；入库日志需要页码时保持默认。

---

## 9. 检查单

上线前逐条确认：

- [ ] 没有用 `name` 当 `Map` 的键或数据库主键；`name` 只用于日志和人工识别。
- [ ] 路由键是 `headers` 的规范序列化，不是 `String.join("|", headers)`。
- [ ] 目标表名、目标列名来自白名单配置，不来自文档表头。
- [ ] 空串有明确策略：文本保留 `""`，数值/比率转 `NULL`，不直接 `new BigDecimal("")`。
- [ ] 金额、比率、整数的转换规则在整个工程里只有一份，不在每个导入点各写一遍。
- [ ] 一张表要么整表入库，要么整表进死信；不出现「前 10 行进了、第 11 行失败」。
- [ ] 未命中的表头、转换失败的行都带 `table#序号`、`name`（可能为 `<unnamed>`）、
      `source.page`（可能为 `?`）进死信表。
- [ ] 用 docx / pdf / pptx 三份样例各跑一遍，并额外造一份「未命名组 + 重名列」的文档验证路由；
      pptx 台账应为 7 行，若出现 `客户名称=客户名称` 那一行，说明只注册了上层表头。
- [ ] 如果同一份文档里可能有两组完全相同的表头，明确它们是「同一目标表的两批行」还是
      「不同目标表」，后者需要页码/序号做二级键。

---

## 10. 已知边界

- **解析器不推断类型**。`"00123"` 是编号还是数字，只能由你按列决定；不要写「看起来像数字就当数字」的通用转换。
- **空串不等于 0**。PDF 导出的合计行可能没有缓存值，空串代表「文档里没有这个值」，不是「值等于 0」。
- **`source` 是表级页码，不是行级页码**。跨页合并后只有 `page` 和 `lastPage`；需要行级定位时，
  要在业务侧自己加行号。
- **没有 upsert / 去重**。解析器只负责读出文档里的行，重复导入、幂等键、唯一约束由数据库和业务层负责。
- **`name` 可能重复、可能缺失、可能受 `headerTolerance` 影响误命中**。把它当标签，不要当契约主键。
- **表头签名是顺序敏感的**。源文档列顺序变化会被当成新表；需要容忍重排时，用命名参数按列名绑定，
  并把签名改成「排序后的列名集合」，但要在配置里明确这是另一种匹配口径。
- **表头补名只追加一次列号**。如果文档的列名本身长得就像补名结果（例如同表出现 `销售额`、`销售额-5`、
  `销售额-5-5`，又有重名的 `销售额`），极端构造下补名后仍可能重复，此时行对象的键会少于列数。
  入库前检查 `new HashSet<String>(headers).size() == headers.size()`；不相等就进死信，不要按位置硬映射。
  正常的业务表头不会触发，但这是「`headers` 与行键逐项一致」的唯一例外。
