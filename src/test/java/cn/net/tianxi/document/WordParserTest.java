package cn.net.tianxi.document;

import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

import java.io.File;
import java.io.FileInputStream;
import java.io.InputStream;
import java.util.Arrays;
import java.util.List;
import java.util.Map;

import static org.junit.jupiter.api.Assertions.*;

/**
 * {@link WordParser} 的功能测试，样例是 {@code src/test/resources/word/sales-report.docx}
 * （2025 年度华东区销售业绩报告：13 段正文 + 4 张表 + 1 张看板截图）。
 *
 * <p>期望值全部来自这份样例的实测输出，改样例就要重跑一遍再改断言。样例里的四张表分别覆盖：
 * <ol>
 *   <li><b>季度销售明细</b>：单行表头、两组同名列（{@code 销售额}/{@code 回款率} 各出现两次）、{@code 片区} 列纵向合并；</li>
 *   <li><b>客户回款台账</b>：两行表头且下一层留空（列名沿用上一层）、{@code 合计} 行是 {@code SUM(ABOVE)} 域、
 *       13 行数据；这张表导出成 PDF 后会跨页断开，PDF 那边验证并回一张；</li>
 *   <li><b>售后工单统计</b>：单元格里再套一张两列小表（{@code 故障部件}/{@code 工单数}），子表另成一组；</li>
 *   <li><b>分片区年度目标</b>：跨整行合并的分组行（{@code 2025 年度（已结账）}、{@code 2026 年度（预算下达）}）、
 *       第四列没有标题（补成 {@code 列4}）、同一张表里两组表头。</li>
 * </ol>
 */
class WordParserTest {

    /**
     * 静默模式（不注册表头、不给关键字）下的实测结果：7 个表头组、30 行数据、1 张图片。
     */
    private static final int SILENT_GROUPS = 7;
    private static final int SILENT_ROWS = 30;

    /**
     * 正文实测是这 13 段：域结果留着、域代码丢掉，软换行换成空格，内容控件里的段落也读到了。
     */
    private static final String TEXT = String.join("\n",
            "2025 年度华东区销售业绩报告",
            "报告页次：3 页",
            "归档编号：",
            "本报告覆盖华东大区与华东南区两个销售片区，统计口径为含税销售额。"
                    + " 数据来源：销售管理系统与财务系统对账结果，金额单位：万元。",
            "编制部门：销售管理部 审核：财务总监",
            "一、季度销售明细",
            "注：回款率 = 已回款金额 ÷ 含税销售额，四舍五入保留整数。",
            "二、客户回款台账",
            "三、售后工单统计",
            "四、分片区年度目标完成情况",
            "五、系统看板截图",
            "附图 1：销售管理系统业绩看板截图",
            "以上数据截至 2025 年 12 月 31 日，最终以财务系统结账数为准。");

    private static Map<String, Object> silent() throws Exception {
        return WordParser.parse(TestSupport.WORD, WordParser.options());
    }

    /* ---------------- 整体形状与契约 ---------------- */

    @Test
    void readsWholeReport() throws Exception {
        Map<String, Object> result = silent();

        assertEquals(SILENT_GROUPS, TestSupport.tables(result).size(), "表头组数");
        assertEquals(SILENT_ROWS, TestSupport.totalRows(result), "数据行数");
        assertEquals(1, TestSupport.images(result).size(), "图片数");
        TestSupport.assertHeadersMatchRowKeys(result);
    }

    @Test
    void everyOptionCombinationKeepsTheContract() throws Exception {
        // 契约（headers 逐个等于每行的键）在任何选项组合下都不能破，否则调用方没法拿 headers 当列清单
        WordParser.Options[] all = {
                WordParser.options(),
                WordParser.options().keywords("销售"),
                TestSupport.registered(),
                TestSupport.registered().fullRow(WordParser.FullRow.FIRST),
                TestSupport.registered().fullRow(WordParser.FullRow.DROP),
                WordParser.options().fullRow(WordParser.FullRow.DROP),
                WordParser.options().images(WordParser.ImageMode.BASE64),
        };
        for (WordParser.Options options : all) {
            TestSupport.assertHeadersMatchRowKeys(WordParser.parse(TestSupport.WORD, options));
        }
    }

    /* ---------------- 正文：域、软换行、内容控件 ---------------- */

    @Test
    void textIsTheThirteenBodyParagraphs() throws Exception {
        assertEquals(TEXT, TestSupport.text(silent()));
    }

    @Test
    void textKeepsCalculatedFieldResultAndDropsUncalculatedField() {
        // 第 2 段是 PAGE 域，缓存的域结果 3 留着、域代码 PAGE 丢掉；第 3 段是 DOCPROPERTY 域且没算过，
        // 整块消失，只剩域前面的「归档编号：」—— 那是排版指令不是内容，见 Results.clean
        List<String> lines = Arrays.asList(TEXT.split("\n"));
        assertEquals("报告页次：3 页", lines.get(1));
        assertEquals("归档编号：", lines.get(2));
        assertFalse(lines.get(1).contains("PAGE"), "域代码不该进正文");
        assertFalse(lines.get(2).contains("DOCPROPERTY"), "域代码不该进正文");
    }

    @Test
    void textJoinsSoftBreakLinesWithASpace() {
        // 第 4 段在 Word 里是 Shift+Enter 分的两行，粘成一句就看不出原本换行了，所以换成空格
        assertTrue(TEXT.split("\n")[3].contains("含税销售额。 数据来源："), "软换行应该换成空格：" + TEXT);
    }

    @Test
    void textReadsParagraphInsideBlockContentControl() {
        // 「编制部门…审核…」那一段包在块级内容控件（w:sdt）里：只看 Body 直接子节点会整段丢掉
        assertTrue(TEXT.contains("编制部门：销售管理部 审核：财务总监"), "内容控件里的段落没读到：" + TEXT);
    }

    @Test
    void textDoesNotRepeatTableCells() throws Exception {
        // 表格里的文字不进 text，否则同一段内容在结果里出现两遍
        // （挑的是只出现在表里的串：正文第 4 段本来就写着「华东大区」，用它断言会误判）
        String text = TestSupport.text(silent());
        assertFalse(text.contains("上海浦东城建集团"), "台账表的客户名不该进 text");
        assertFalse(text.contains("2400.00"), "明细表的金额不该进 text");
        assertFalse(text.contains("超额完成"), "明细表备注列不该进 text");
        assertFalse(text.contains("更换配件后复测通过"), "工单表说明列不该进 text");
    }

    /* ---------------- 表 1：同名列 + 纵向合并 ---------------- */

    @Test
    void sameNamedColumnsGetColumnNumberSuffix() throws Exception {
        // 表头里「销售额」「回款率」各出现两次（上半年／全年），摊平后必须靠后缀区分，否则一行数据里两列会互相覆盖
        assertEquals(Arrays.asList("片区", "区域", "销售额", "回款率", "销售额-5", "回款率-6", "备注"),
                TestSupport.headers(TestSupport.table(silent(), 0)));
    }

    @Test
    void verticallyMergedColumnIsFilledFromItsOriginRow() throws Exception {
        // 片区列纵向合并：江苏那行、安徽那行自己没有内容，要回填起始格的值
        Map<String, Object> result = silent();
        assertEquals(Arrays.asList("华东大区", "上海", "1280", "92%", "1465", "95%", "超额完成"),
                TestSupport.values(result, 0, 0));
        assertEquals(Arrays.asList("华东大区", "江苏", "986", "88%", "1120", "90%", ""),
                TestSupport.values(result, 0, 1));
        assertEquals(Arrays.asList("华东南区", "浙江", "760", "85%", "890", "87%", "回款偏长"),
                TestSupport.values(result, 0, 2));
        assertEquals(Arrays.asList("华东南区", "安徽", "540", "80%", "616", "83%", ""),
                TestSupport.values(result, 0, 3));
        assertEquals(4, TestSupport.rows(TestSupport.table(result, 0)).size(), "纵向合并不该把行判成分组行");
    }

    /* ---------------- 表 2：两行表头 + 域结果 ---------------- */

    @Test
    void twoLevelHeaderInheritsUpperLevelColumnNames() throws Exception {
        // 上层是「客户名称 | 回款情况(跨两列) | 回款率 | 备注」，下层只填了「合同金额 | 已回款」，
        // 空着的列沿用上层名字 —— 两行连续表头且中间没有数据行时才这样合并
        assertEquals(Arrays.asList("客户名称", "合同金额", "已回款", "回款率", "备注"),
                TestSupport.headers(TestSupport.table(silent(), 1)));
    }

    @Test
    void sumFieldResultIsReadFromTheCachedValue() throws Exception {
        // 合计行的金额是 SUM(ABOVE) 域：Word 把上次算出的结果缓存在文档里，解析读的就是这个缓存值
        Map<String, Object> result = silent();
        List<Map<String, String>> rows = TestSupport.rows(TestSupport.table(result, 1));
        assertEquals(13, rows.size(), "12 行客户 + 1 行合计");
        assertEquals(Arrays.asList("上海浦东城建集团", "486.00", "486.00", "100%", "已结清"),
                TestSupport.values(rows.get(0)));
        assertEquals(Arrays.asList("合计", "2096.30", "1796.10", "86%", ""),
                TestSupport.values(rows.get(rows.size() - 1)));
    }

    /* ---------------- 表 3：单元格内嵌套表 ---------------- */

    @Test
    void nestedTableBecomesItsOwnGroupAndLeavesTheHostCellClean() throws Exception {
        // 工单表的「硬件故障」那格里套了一张两列小表：母表那一格不摊平子表文字，子表另成一组排在母表之后
        Map<String, Object> result = silent();
        Map<String, Object> outer = TestSupport.table(result, 2);
        Map<String, Object> inner = TestSupport.table(result, 3);

        assertEquals(Arrays.asList("工单类型", "工单数", "处理说明"), TestSupport.headers(outer));
        assertEquals(Arrays.asList("硬件故障", "42", ""), TestSupport.values(result, 2, 0));
        assertEquals(Arrays.asList("软件异常", "27", "远程处理，平均响应 2 小时"), TestSupport.values(result, 2, 2));
        assertEquals(Arrays.asList("现场支持", "15", "派工上门，覆盖江浙沪皖"), TestSupport.values(result, 2, 3));

        assertEquals(Arrays.asList("故障部件", "工单数"), TestSupport.headers(inner));
        assertEquals(Arrays.asList("读卡器", "18"), TestSupport.values(result, 3, 0));
        assertEquals(Arrays.asList("打印机", "15"), TestSupport.values(result, 3, 1));
        assertEquals(Arrays.asList("显示屏", "9"), TestSupport.values(result, 3, 2));
    }

    /* ---------------- 表 4：跨整行合并的分组行 ---------------- */

    @Test
    void separatorRowBeforeAnyHeaderStartsANumberedGroup() throws Exception {
        // 目标表第一行就是跨整行合并的「2025 年度（已结账）」，此时还没有任何表头：另起一组、列名用序号兜底，
        // 否则这行分组文字会无声丢掉
        Map<String, Object> result = silent();
        assertEquals(Arrays.asList("列1", "列2", "列3", "列4"), TestSupport.headers(TestSupport.table(result, 4)));
        assertEquals(Arrays.asList("2025 年度（已结账）", "", "", ""), TestSupport.values(result, 4, 0));
    }

    @Test
    void oneTableCanCarryTwoHeaderGroups() throws Exception {
        // 同一张物理表里「2026 年度（预算下达）」又是一行分组合并行，它结束了上一组、开出新的一组
        Map<String, Object> result = silent();
        assertEquals(Arrays.asList("片区", "年度目标", "完成率", "列4"),
                TestSupport.headers(TestSupport.table(result, 5)));
        assertEquals(Arrays.asList("华东大区", "2400.00", "101%", "达标"), TestSupport.values(result, 5, 0));
        assertEquals(Arrays.asList("华东南区", "1500.00", "93%", "未达标"), TestSupport.values(result, 5, 1));
        assertEquals(Arrays.asList("2026 年度（预算下达）", "", "", ""), TestSupport.values(result, 5, 2));

        assertEquals(TestSupport.headers(TestSupport.table(result, 5)),
                TestSupport.headers(TestSupport.table(result, 6)));
        assertEquals(Arrays.asList("华东大区", "2800.00", "—", ""), TestSupport.values(result, 6, 0));
        assertEquals(Arrays.asList("华东南区", "1750.00", "—", ""), TestSupport.values(result, 6, 1));
    }

    @Test
    void untitledLastColumnIsNamedByItsPosition() throws Exception {
        // 目标表第四列（达标/未达标）在文档里没有标题，补成「列4」；headers 与行键因此仍然对齐
        List<String> headers = TestSupport.headers(TestSupport.table(silent(), 5));
        assertEquals("列4", headers.get(3));
    }

    /* ---------------- FullRow 三种去向 ---------------- */

    @Test
    void dropRemovesEverySeparatorRow() throws Exception {
        // DROP：两行分组小标题都没了 —— 组数 7→6（「2025 年度（已结账）」那组整组消失），行数 30→28
        Map<String, Object> result =
                WordParser.parse(TestSupport.WORD, TestSupport.registered().fullRow(WordParser.FullRow.DROP));

        assertEquals(6, TestSupport.tables(result).size(), "组数");
        assertEquals(28, TestSupport.totalRows(result), "行数");
        String json = WordParser.toJson(result);
        assertFalse(json.contains("2025 年度（已结账）"), "分组行应该被丢掉");
        assertFalse(json.contains("2026 年度（预算下达）"), "分组行应该被丢掉");
        TestSupport.assertHeadersMatchRowKeys(result);
    }

    @Test
    void firstRoutesEverySeparatorRowOfATableToThatTablesFirstGroup() throws Exception {
        // FIRST：行数与组数都和默认相同（分组行都留着），但目标表里第二行分组小标题「2026 年度（预算下达）」
        // 从它自己那组挪到了这张表的第一组 —— 也就是「2025 年度（已结账）」开出来的那组序号列
        Map<String, Object> keep = silent();
        Map<String, Object> first =
                WordParser.parse(TestSupport.WORD, WordParser.options().fullRow(WordParser.FullRow.FIRST));

        assertEquals(SILENT_ROWS, TestSupport.totalRows(first), "行数不变");
        assertEquals(SILENT_GROUPS, TestSupport.tables(first).size(), "组数不变");
        assertEquals(3, TestSupport.rows(TestSupport.table(keep, 5)).size(), "KEEP：分组行留在自己那组");
        assertEquals(2, TestSupport.rows(TestSupport.table(first, 5)).size(), "FIRST：分组行被挪走");

        assertEquals(Arrays.asList("列1", "列2", "列3", "列4"), TestSupport.headers(TestSupport.table(first, 4)));
        assertEquals(Arrays.asList("2025 年度（已结账）", "", "", ""), TestSupport.values(first, 4, 0));
        assertEquals(Arrays.asList("2026 年度（预算下达）", "", "", ""), TestSupport.values(first, 4, 1));
        TestSupport.assertHeadersMatchRowKeys(first);
    }

    /* ---------------- 注册表头 / 关键字 ---------------- */

    @Test
    void registeringTheTrueHeadersChangesNothing() throws Exception {
        // 样例的表头本来就写得规范：注册真表头后结果与静默模式逐字节相同。
        // 注册的价值在别处 —— 表头有错字／PDF 单元格换行时，用注册名才能保证三种格式的键一致（见 PdfParserTest）
        assertEquals(WordParser.toJson(silent()), WordParser.toJson(WordParser.parse(TestSupport.WORD,
                TestSupport.registered())));
    }

    @Test
    void keywordsOptionIsTheSameAsVarargsEntry() throws Exception {
        // parse(path, keywords...) 就是 options().keywords(...)，两条入口结果一致
        assertEquals(WordParser.toJson(silent()),
                WordParser.toJson(WordParser.parse(TestSupport.WORD, "销售")));
    }

    @Test
    void registeredHeadersOverrideTheNamesReadFromTheDocument() throws Exception {
        // 注册名少一列时命不中（列数必须相同）；这里用「错一个字」验证容忍度：默认容忍 1 列不一致，仍然命中，
        // 且结果里的列名用注册的原名
        Map<String, Object> result = WordParser.parse(TestSupport.WORD, WordParser.options()
                .headers("片区", "区域", "销售额", "回款率", "销售额", "回款率", "说明"));

        assertEquals(Arrays.asList("片区", "区域", "销售额", "回款率", "销售额-5", "回款率-6", "说明"),
                TestSupport.headers(TestSupport.table(result, 0)));
    }

    @Test
    void unmatchedHeaderGroupIsReportedOnStderr() throws Exception {
        // 注册了却一行都没匹配上，说明表头文本与文档不一致；这种情况会让数据行整段并进上一组，所以要提示
        String stderr = TestSupport.captureStderr(() ->
                WordParser.parse(TestSupport.WORD, WordParser.options().headers("不存在的列", "也不存在的列")));

        assertTrue(stderr.contains("表头集合未匹配到任何行"), stderr);
        assertTrue(stderr.contains("不存在的列"), stderr);
    }

    @Test
    void unmatchedHeadersListKeepsTheRegisteredSpelling(@TempDir File directory) throws Exception {
        // 清单与提示里写的都是注册时的<b>原文</b>，不是比对用的归一化形式（去掉全部空白、转小写）——
        // 归一化之后调用方拿它去改自己的注册代码会对不上号，这条线索就白给了。
        // 要现造一份 docx 才看得出差别：三份样例的表头全是中文，归一化前后长得一样
        File file = TestSupport.docx(directory, "ascii-headers.docx", TestSupport.tbl(2,
                TestSupport.tr(TestSupport.tc("AREA"), TestSupport.tc("TARGET")),
                TestSupport.tr(TestSupport.tc("East"), TestSupport.tc("100"))));

        Map<String, Object> result = WordParser.parse(file.getPath(), WordParser.options()
                .headers("area", " target ")                       // 归一化后与文档那一行相等 → 命中
                .headers("Total  Amount", "COUNT").name("总额")     // 一行都没命中 → 进清单
                .unmatchedHeaders(true));

        @SuppressWarnings("unchecked")
        List<Map<String, Object>> unmatched = (List<Map<String, Object>>) result.get("unmatchedHeaders");

        assertEquals(1, unmatched.size(), "两组注册里只有第二组没命中");
        assertEquals(Arrays.asList("Total  Amount", "COUNT"), unmatched.get(0).get("headers"),
                "原文照抄：中间那两个空格与大写都留着，只 trim 掉首尾");
        assertEquals("总额", unmatched.get(0).get("name"), "起了名字的注册把名字一起带上");
        assertEquals(Arrays.asList("area", "target"), TestSupport.headers(TestSupport.table(result, 0)),
                "命中那组的列名换成注册时写的名字（trim 过），文档里是大写也一样");
    }

    @Test
    void aTwoLevelHeaderMatchesTheMergedRowAtAnyTolerance() throws Exception {
        // 台账在 Word 里是两行表头（上层的「回款情况」跨两列，两行都带 w:tblHeader），注册的却是合并之后的那一行。
        // 网格里单独的两行都对不上：上层那行与注册名差 2 格、超出默认容忍度 1；下层那行只有中间两格有字
        // （纵向合并的延续格没有文字），末尾空列在比对前被裁掉，只剩 3 格、列数也不同。
        // 命中靠的是分组算法在合并完这两行之后拿<b>合并结果</b>又比了一次，所以容忍度收到 0 也照样命中 ——
        // 卡的不是容差，而是「有没有拿合并结果比过」。列名与分组都不受注册写法影响。
        final String stderr = TestSupport.captureStderr(() ->
                WordParser.parse(TestSupport.WORD, TestSupport.registered()));

        assertEquals("", stderr, "五组注册在 docx 上全部命中，不该留下未匹配提示");
        assertFalse(stderr.contains("片区"), "核心表与目标表都该命中：" + stderr);

        Map<String, Object> result = WordParser.parse(TestSupport.WORD, TestSupport.registered());
        assertEquals(Arrays.asList("客户名称", "合同金额", "已回款", "回款率", "备注"),
                TestSupport.headers(TestSupport.table(result, 1)));

        // 上面那句「差 2 格、超出容忍度 1」是可验证的，不是推测：容忍度调到 2 时上层那行本身就能命中；
        // 照着文档里真实的上层那一行注册，默认容忍度下也命中。三种写法都落在第 1 组，组数与列名不变
        Map<String, Object> tolerant = WordParser.parse(TestSupport.WORD, WordParser.options()
                .headers("客户名称", "合同金额", "已回款", "回款率", "备注").name("回款台账")
                .headerTolerance(2));
        assertEquals("回款台账", TestSupport.table(tolerant, 1).get("name"), "容忍度 2 能盖住那 2 格差异");
        assertEquals(Arrays.asList("客户名称", "合同金额", "已回款", "回款率", "备注"),
                TestSupport.headers(TestSupport.table(tolerant, 1)));

        Map<String, Object> strict = WordParser.parse(TestSupport.WORD, WordParser.options()
                .headers("客户名称", "合同金额", "已回款", "回款率", "备注").name("回款台账")
                .headerTolerance(0));
        assertEquals("回款台账", TestSupport.table(strict, 1).get("name"), "容忍度 0 靠合并结果命中");

        Map<String, Object> upperRow = WordParser.parse(TestSupport.WORD, WordParser.options()
                .headers("客户名称", "回款情况", "回款情况", "回款率", "备注").name("回款台账"));
        assertEquals("回款台账", TestSupport.table(upperRow, 1).get("name"), "注册上层那一行则默认容忍度即可命中");
        assertEquals(Arrays.asList("客户名称", "合同金额", "已回款", "回款率", "备注"),
                TestSupport.headers(TestSupport.table(upperRow, 1)), "合并之后列名与注册合并写法时一致");
        assertEquals(7, TestSupport.tables(upperRow).size(), "换一种注册写法不改变分组");
    }

    /* ---------------- 图片 ---------------- */

    @Test
    void metaModeGivesOnlyMetadata() throws Exception {
        Map<String, Object> image = TestSupport.image(silent(), 0);

        assertEquals("image-1.png", image.get("fileName"));
        assertEquals("PNG", image.get("type"));
        assertEquals(1024, image.get("width"));
        assertEquals(768, image.get("height"));
        assertEquals(670939, image.get("size"));
        assertEquals("487e915db2c53907e8c1551aeaffea235423210c", image.get("contentId"));
        assertFalse(image.containsKey("base64"), "默认不带图片内容");
        assertFalse(image.containsKey("bytes"), "默认不带图片内容");
        assertFalse(image.containsKey("path"), "默认不落盘");
        TestSupport.assertImageMeta(image);
    }

    @Test
    void base64ModeCarriesADataUriAndGrowsTheResult() throws Exception {
        String meta = WordParser.toJson(silent());
        Map<String, Object> result =
                WordParser.parse(TestSupport.WORD, WordParser.options().images(WordParser.ImageMode.BASE64));
        String base64 = WordParser.toJson(result);
        String payload = (String) TestSupport.image(result, 0).get("base64");

        assertTrue(payload.startsWith("data:image/png;base64,"), payload.substring(0, 40));
        // 实测：META 4740 字符 → BASE64 899370 字符，那张 670939 字节的图编码后约 89 万字符。
        // 这就是「只关心文字和表格时别选 BASE64」的理由
        assertTrue(base64.length() > meta.length() * 100, "meta=" + meta.length() + " base64=" + base64.length());
        assertEquals(670939, java.util.Base64.getDecoder()
                .decode(payload.substring(payload.indexOf(',') + 1)).length, "解码回来应该与原图字节数相同");
    }

    @Test
    void bytesModeCarriesTheRawBytes() throws Exception {
        Map<String, Object> result =
                WordParser.parse(TestSupport.WORD, WordParser.options().images(WordParser.ImageMode.BYTES));
        Object bytes = TestSupport.image(result, 0).get("bytes");

        assertTrue(bytes instanceof byte[], "BYTES 模式应该给原始字节");
        assertEquals(670939, ((byte[]) bytes).length);
    }

    @Test
    void fileModeWritesTheImageAndKeepsThePath(@TempDir File directory) throws Exception {
        Map<String, Object> result = WordParser.parse(TestSupport.WORD,
                WordParser.options().images(WordParser.ImageMode.FILE).imageDir(directory));
        Map<String, Object> image = TestSupport.image(result, 0);

        File written = new File(directory, String.valueOf(image.get("fileName")));
        assertTrue(written.isFile(), "图片应该落盘：" + written);
        assertEquals(670939, written.length());
        assertEquals(written.getAbsolutePath(), image.get("path"));
        assertFalse(image.containsKey("base64"), "FILE 模式结果里不留内容");
        assertFalse(image.containsKey("bytes"), "FILE 模式结果里不留内容");
        assertEquals(1, directory.listFiles().length, "只有一张图，不该多出文件");
    }

    @Test
    void fileModeWithoutDirectoryFails() {
        // 静默失败最难受：解析跑完了、图片一张也没落地，所以直接抛
        IllegalStateException thrown = assertThrows(IllegalStateException.class, () ->
                WordParser.parse(TestSupport.WORD, WordParser.options().images(WordParser.ImageMode.FILE)));
        assertTrue(thrown.getMessage().contains("imageDir"), thrown.getMessage());
    }

    @Test
    void imageDirWithoutFileModeWarns(@TempDir File directory) throws Exception {
        // imageDir 只在 FILE 模式下生效，配错组合要提示，否则调用方会以为图已经导出去了
        String stderr = TestSupport.captureStderr(() ->
                WordParser.parse(TestSupport.WORD, WordParser.options().imageDir(directory)));

        assertTrue(stderr.contains("imageDir"), stderr);
        assertTrue(stderr.contains(String.valueOf(directory)), stderr);
        assertEquals(0, directory.listFiles().length, "META 模式不该写文件");
    }

    @Test
    void writeImagesExportsTheBase64Payload(@TempDir File directory) throws Exception {
        Map<String, Object> result =
                WordParser.parse(TestSupport.WORD, WordParser.options().images(WordParser.ImageMode.BASE64));

        WordParser.writeImages(result, directory);

        File written = new File(directory, "image-1.png");
        assertTrue(written.isFile(), "图片应该导出到目录里");
        assertEquals(670939, written.length(), "导出的字节数应与原图一致");
    }

    /* ---------------- 入口 ---------------- */

    @Test
    void parsesFromInputStream() throws Exception {
        InputStream stream = new FileInputStream(TestSupport.WORD);
        try {
            assertEquals(WordParser.toJson(silent()), WordParser.toJson(WordParser.parse(stream,
                    WordParser.options())));
        } finally {
            stream.close();
        }
    }

    @Test
    void missingFileReportsThePath() {
        // Aspose.Words 抛的是 JDK 的 FileNotFoundException，消息是「Can't find file: <路径>.」，路径原样带上。
        // pdf 与 pptx 抛的是各自库里的类型，三个解析器不统一，见 CrossFormatTest.aMissingFileThrowsFromAllThreeParsers
        String path = "src/test/resources/word/not-here.docx";

        java.io.FileNotFoundException thrown = assertThrows(java.io.FileNotFoundException.class, () ->
                WordParser.parse(path, WordParser.options()));

        assertTrue(thrown.getMessage().contains(path), thrown.getMessage());
    }
}
