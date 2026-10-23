package cn.net.tianxi.document;

import org.junit.jupiter.api.Test;

import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;
import java.util.Map;

import static org.junit.jupiter.api.Assertions.*;

/**
 * 注册时起的组名（{@code name} 键）与按名字取组取行的两个入口
 * （{@link WordParser#tablesByName(Map, String)}、{@link WordParser#rowsByName(Map, String)}）。
 *
 * <p>这一条链是三种格式共用的：{@code name} 写在 {@link Tables#group} 命中注册表头的那一刻，
 * 取值入口只看结果里的键，所以 {@link PdfParser} 与 {@link PptParser} 的结果一样能按名字取。
 * 分组算法本身（谁能命中注册集、多级表头怎么合）在 {@code TablesTest} 里直接喂网格测过，这里只测端到端。
 *
 * <h2>样例上的实测分布（{@code TestSupport.named()}：五组注册、其中三组起了名字）</h2>
 * <table border="1">
 *   <caption>三份样例里带 {@code name} 的组与 JSON 长度（组下标从 0 起）</caption>
 *   <tr><th>样例</th><th>带 name 的组</th><th>行数</th><th>静默</th><th>注册不起名</th><th>注册并起名</th></tr>
 *   <tr><td>sales-report.docx</td><td>[0] 季度明细、[1] 回款台账、[5][6] 年度目标 —— 共 4 组</td><td>30</td>
 *       <td>7 组 / 4740 字符</td><td>7 组 / 4740 字符</td><td>7 组 / 4828 字符</td></tr>
 *   <tr><td>sales-report.pdf</td><td>[0] 季度明细、[1] 回款台账、[4][5] 年度目标 —— 共 4 组</td><td>45</td>
 *       <td>6 组 / 6473 字符</td><td>6 组 / 6473 字符</td><td>6 组 / 6561 字符</td></tr>
 *   <tr><td>sales-report.pptx</td><td>[0] 季度明细、[1] 回款台账、[3][4] 年度目标 —— 共 4 组</td><td>21</td>
 *       <td>6 组 / 3526 字符</td><td>6 组 / 3378 字符</td><td>6 组 / 3466 字符</td></tr>
 * </table>
 *
 * <p>{@code 年度目标} 是三份样例都命中、而且每份都命中<b>两组</b>的那一组：同一张物理表被 2025 与 2026
 * 两段表头切成两个组。它专门用来验证「一个名字对应多组」不是意外而是常态。
 *
 * <p>{@code 回款台账} 是<b>两行表头</b>的那一张：上层「{@code 客户名称｜回款情况}（跨两列）{@code ｜回款率｜备注}」，
 * 下层空两格填「{@code 合同金额｜已回款}」。注册的是<b>合并之后</b>的五个列名，三种格式现在都命中 ——
 * 分组算法除了拿网格里的每一行单独比，多级表头合并完还会拿<b>合并结果</b>再比一次。
 * 剩下的格式差异在「照上层那一行注册」：docx 与 pptx 的网格里有那一行、pdf 里没有（pdf 给的就是合并后那一行），
 * 而 pptx 只有上层那行被当成表头，下层那行落进数据，行数从 7 变成 8。
 * 详见 {@link #aMultiLevelHeaderHitsByItsMergedRowInAllThreeFormats()}。
 */
class TableByNameTest {

    /* ---------------- name 键写在哪几组 ---------------- */

    @Test
    void onlyTheGroupsWhoseHeaderMatchedANamedRegistrationCarryTheKey() throws Exception {
        Map<String, Object> word = WordParser.parse(TestSupport.WORD, TestSupport.named());

        assertEquals("季度明细", TestSupport.table(word, 0).get("name"), "单行表头，注册名与文档那一行逐列相等");
        assertEquals("回款台账", TestSupport.table(word, 1).get("name"),
                "两行表头：注册的是合并之后那一行，合并完再比一次就命中了");
        assertNoName(TestSupport.table(word, 2), "工单表注册了但没起名");
        assertNoName(TestSupport.table(word, 3), "嵌套小表连注册都没有");
        assertNoName(TestSupport.table(word, 4), "分组小标题自己开的那一组，列名是补出来的 列1..列4，不可能命中注册集");
        assertEquals("年度目标", TestSupport.table(word, 5).get("name"));
        assertEquals("年度目标", TestSupport.table(word, 6).get("name"), "同一张物理表的两个表头组带同一个名字");
    }

    @Test
    void keyOrderIsHeadersRowsNameSourceAndUnnamedGroupsStayUntouched() throws Exception {
        Map<String, Object> word = WordParser.parse(TestSupport.WORD, TestSupport.named());

        for (int i : new int[]{0, 1}) {
            assertEquals(Arrays.asList("headers", "rows", "name", "source"),
                    new ArrayList<Object>(TestSupport.table(word, i).keySet()),
                    "第 " + (i + 1) + " 组起过名字：name 在 rows 之后、source 之前");
        }
        for (int i = 2; i <= 4; i++) {
            assertEquals(Arrays.asList("headers", "rows", "source"),
                    new ArrayList<Object>(TestSupport.table(word, i).keySet()), "第 " + (i + 1) + " 组没起名，键集合不变");
        }
    }

    @Test
    void nameWorksTheSameWayOnAllThreeFormats() throws Exception {
        // 名字写在分组算法里，三种格式共用同一份，所以命中位置只差在「第几组」上，键与值完全一致
        Map<String, Object> word = WordParser.parse(TestSupport.WORD, TestSupport.named());
        Map<String, Object> pdf = PdfParser.parse(TestSupport.PDF, TestSupport.named());
        Map<String, Object> ppt = PptParser.parse(TestSupport.PPT, TestSupport.named());

        assertEquals("季度明细", TestSupport.table(word, 0).get("name"));
        assertEquals("季度明细", TestSupport.table(pdf, 0).get("name"));
        assertEquals("季度明细", TestSupport.table(ppt, 0).get("name"));

        assertEquals(2, WordParser.tablesByName(word, "年度目标").size(), "docx：2025 与 2026 两组");
        assertEquals(2, WordParser.tablesByName(pdf, "年度目标").size(), "pdf：跨页并回一张表后仍然是两组");
        assertEquals(2, WordParser.tablesByName(ppt, "年度目标").size(), "pptx：两段表头各一组");

        assertEquals(5, WordParser.rowsByName(word, "年度目标").size(), "3 行 + 2 行");
        assertEquals(5, WordParser.rowsByName(pdf, "年度目标").size());
        assertEquals(5, WordParser.rowsByName(ppt, "年度目标").size());
    }

    @Test
    void aMultiLevelHeaderHitsByItsMergedRowInAllThreeFormats() throws Exception {
        // 台账是两行表头：上层「客户名称｜回款情况（跨两列）｜回款率｜备注」，下层空两格填「合同金额｜已回款」。
        // 分组算法先拿网格里的每一行单独比，多级表头合并完还会拿合并结果再比一次，
        // 所以注册「合并之后那一行」三种格式都命中；容忍度收到 0（要求逐格相等）也一样命中。
        List<String> merged = Arrays.asList("客户名称", "合同金额", "已回款", "回款率", "备注");
        for (int tolerance : new int[]{1, 0}) {
            WordParser.Options only = WordParser.options()
                    .headers(merged.toArray(new String[merged.size()])).name("回款台账").headerTolerance(tolerance);
            String why = "只注册台账、容忍度 " + tolerance;

            assertEquals(13, WordParser.rowsByName(WordParser.parse(TestSupport.WORD, only), "回款台账").size(),
                    why + "：docx 12 家客户 + 合计");
            assertEquals(31, WordParser.rowsByName(PdfParser.parse(TestSupport.PDF, only), "回款台账").size(),
                    why + "：pdf 是加长版样例，30 家 + 合计");
            assertEquals(7, WordParser.rowsByName(PptParser.parse(TestSupport.PPT, only), "回款台账").size(),
                    why + "：pptx 6 家 + 合计");
        }

        // 五组一起注册时结论不变，而且台账在三份样例里都是第 2 组
        Map<String, Object> word = WordParser.parse(TestSupport.WORD, TestSupport.named());
        Map<String, Object> pdf = PdfParser.parse(TestSupport.PDF, TestSupport.named());
        Map<String, Object> ppt = PptParser.parse(TestSupport.PPT, TestSupport.named());
        for (Map<String, Object> result : Arrays.asList(word, pdf, ppt)) {
            assertEquals(merged, TestSupport.headers(TestSupport.table(result, 1)), "列名三种格式一致");
            assertEquals("回款台账", TestSupport.table(result, 1).get("name"), "都带上了名字");
        }

        // 换成注册「上层那一行」就有格式差异了，三条都是实测出来的：
        //   docx 13 行 —— 台账两行表头在 XML 里都带 w:tblHeader，两行都算表头、照样合并成一组；
        //   pptx 8 行 —— 它的网格里上下两行都在，但只有上层那行命中注册、被当成表头，
        //                下层那行不再是表头就落进了数据（实测它是第 0 行，值恰好是 客户名称/合同金额/已回款/回款率/备注），
        //                与静默模式读到的 8 行一模一样 —— 这就是「多级表头只注册了上层」的代价；
        //   pdf  命不中 —— 它的网格里给的就是合并后那一行，没有上层那一行，标准错误上会提示一条。
        WordParser.Options upper = WordParser.options()
                .headers("客户名称", "回款情况", "回款情况", "回款率", "备注").name("回款台账");
        assertEquals(13, WordParser.rowsByName(WordParser.parse(TestSupport.WORD, upper), "回款台账").size(),
                "docx 命中上层那一行，下层那行靠 w:tblHeader 合并进来");
        assertEquals(8, WordParser.rowsByName(PptParser.parse(TestSupport.PPT, upper), "回款台账").size(),
                "pptx 也命中上层那一行，但下层那行没被合并、当成数据留在第 0 行");
        assertTrue(WordParser.rowsByName(PdfParser.parse(TestSupport.PDF, upper), "回款台账").isEmpty(),
                "pdf 命不中：它的网格里没有上层那一行");

        // 跨格式共用一份注册时最稳的写法仍然是上下两种都注册、起同一个名字，组数与行数都不变
        WordParser.Options both = TestSupport.namedBothHeaderRows();

        Map<String, Object> wordBoth = WordParser.parse(TestSupport.WORD, both);
        Map<String, Object> pdfBoth = PdfParser.parse(TestSupport.PDF, both);
        Map<String, Object> pptBoth = PptParser.parse(TestSupport.PPT, both);

        assertEquals("回款台账", TestSupport.table(wordBoth, 1).get("name"), "docx 靠上层那一行命中");
        assertEquals("回款台账", TestSupport.table(pdfBoth, 1).get("name"), "pdf 靠合并后那一行命中");
        assertEquals("回款台账", TestSupport.table(pptBoth, 1).get("name"));
        assertEquals(30, TestSupport.totalRows(wordBoth), "多注册一组不改变行数：docx 仍是 7 组 30 行");
        assertEquals(45, TestSupport.totalRows(pdfBoth), "pdf 仍是 6 组 45 行");
        assertEquals(21, TestSupport.totalRows(pptBoth), "pptx 仍是注册模式的 6 组 21 行");
        assertEquals(TestSupport.headers(TestSupport.table(word, 1)),
                TestSupport.headers(TestSupport.table(wordBoth, 1)), "列名不受多注册这一组的影响");
        assertEquals(WordParser.tablesByName(word, "年度目标").size(),
                WordParser.tablesByName(wordBoth, "年度目标").size(), "其它名字的命中也不受影响");
    }

    @Test
    void theDefaultResultHasNoNameKeyAtAll() throws Exception {
        // 红线：不起名（含只注册不起名）时，结果与改动前逐字节相同，多出来的键只有 source
        for (String sample : TestSupport.ALL_SAMPLES) {
            Map<String, Object> result = parse(sample, WordParser.options());
            for (Map<String, Object> table : TestSupport.tables(result)) {
                assertEquals(Arrays.asList("headers", "rows", "source"),
                        new ArrayList<Object>(table.keySet()), sample + " 的默认键集合");
            }
        }

        assertEquals(WordParser.toJson(WordParser.parse(TestSupport.WORD, WordParser.options())),
                WordParser.toJson(WordParser.parse(TestSupport.WORD, TestSupport.registered())),
                "注册五组表头但不起名，JSON 与静默模式逐字节相同");
    }

    @Test
    void nameSurvivesTheSourceSwitchAndKeepsTheOrder() throws Exception {
        Map<String, Object> word = WordParser.parse(TestSupport.WORD, TestSupport.named().source(false));

        assertEquals(Arrays.asList("headers", "rows", "name"),
                new ArrayList<Object>(TestSupport.table(word, 0).keySet()), "关掉 source 之后起名那组多一个 name");
        assertEquals(Arrays.asList("headers", "rows"),
                new ArrayList<Object>(TestSupport.table(word, 2).keySet()), "没起名的组仍然只有两个键");
        assertEquals("季度明细", WordParser.tablesByName(word, "季度明细").get(0).get("name"), "取数入口不受开关影响");
    }

    /* ---------------- 未匹配提示与 name 的关系 ---------------- */

    @Test
    void aRegistrationCoveredByASameNameMatchIsNoLongerReportedOnStderr() throws Exception {
        // 文档给的解法（多级表头把上下两种写法都注册、起同一个 name）不该反过来留下一条提示：
        // 同名之下已有一份命中，另一份命不中就不算「文档里没有这张表」，再提示就是噪声
        String word = TestSupport.captureStderr(() ->
                WordParser.parse(TestSupport.WORD, TestSupport.namedBothHeaderRows()));
        assertFalse(word.contains("客户名称"), "同名注册已经命中，不该再提示：" + word);
        assertFalse(word.contains("合同金额"), word);

        // 没有起名字的注册不受这条影响：pdf 里没有 docx 那张嵌套小表，照样提示
        String pdf = TestSupport.captureStderr(() ->
                PdfParser.parse(TestSupport.PDF, TestSupport.namedBothHeaderRows()));
        assertTrue(pdf.contains("故障部件"), "没起名字的未匹配注册照样提示：" + pdf);

        // 只注册合并后那一行（named()）时，docx 现在也命中了（多级表头合并完会再比一次），
        // 于是五组全命中、标准错误上一个字都不打 —— 实测就是空字符串
        String wordNamed = TestSupport.captureStderr(() ->
                WordParser.parse(TestSupport.WORD, TestSupport.named()));
        assertEquals("", wordNamed, "docx 上 named() 的五组注册全部命中，不该有提示");
    }

    /* ---------------- 取值入口 ---------------- */

    @Test
    void tablesByNameReturnsTheSameObjectsInTheTablesOrder() throws Exception {
        Map<String, Object> word = WordParser.parse(TestSupport.WORD, TestSupport.named());
        List<Map<String, Object>> hits = WordParser.tablesByName(word, "年度目标");

        assertEquals(2, hits.size());
        assertSame(TestSupport.table(word, 5), hits.get(0), "返回的是结果里的原对象，不是副本");
        assertSame(TestSupport.table(word, 6), hits.get(1));
        assertEquals(TestSupport.headers(TestSupport.table(word, 5)), TestSupport.headers(hits.get(0)),
                "顺序按 tables 里的先后");
    }

    @Test
    void rowsByNameFlattensEveryNamedGroupInTheSameOrder() throws Exception {
        Map<String, Object> word = WordParser.parse(TestSupport.WORD, TestSupport.named());
        List<Map<String, String>> rows = WordParser.rowsByName(word, "年度目标");

        List<Map<String, String>> expected = new ArrayList<Map<String, String>>();
        expected.addAll(TestSupport.rows(TestSupport.table(word, 5)));
        expected.addAll(TestSupport.rows(TestSupport.table(word, 6)));
        assertEquals(expected, rows, "同名各组的行按组顺序拼平");

        assertEquals(Arrays.asList("华东大区", "2400.00", "101%", "达标"), TestSupport.values(rows.get(0)),
                "第 1 行来自第 6 组");
        assertEquals(Arrays.asList("华东大区", "2800.00", "—", ""), TestSupport.values(rows.get(3)),
                "第 4 行是第 7 组的第一行；两组列名一致，所以键顺序也一致");
        assertSame(TestSupport.rows(TestSupport.table(word, 6)).get(0), rows.get(3), "行对象也是原对象");
    }

    @Test
    void aSingleGroupTableGivesOneHitAndItsOwnRows() throws Exception {
        Map<String, Object> word = WordParser.parse(TestSupport.WORD, TestSupport.named());

        assertEquals(1, WordParser.tablesByName(word, "季度明细").size());
        assertEquals(4, WordParser.rowsByName(word, "季度明细").size(), "4 行数据，跨页与不跨页的都在这里");
        assertEquals(Arrays.asList("片区", "区域", "销售额", "回款率", "销售额-5", "回款率-6", "备注"),
                TestSupport.headers(TestSupport.table(word, 0)), "列名仍是补好重名后缀的那一份");
    }

    @Test
    void unknownBlankOrNullNamesGiveEmptyListsInsteadOfFailing() throws Exception {
        Map<String, Object> word = WordParser.parse(TestSupport.WORD, TestSupport.named());

        assertTrue(WordParser.tablesByName(word, "不存在的名字").isEmpty(), "没命中返回空集合，不是 null");
        assertTrue(WordParser.rowsByName(word, "不存在的名字").isEmpty());
        assertTrue(WordParser.tablesByName(word, null).isEmpty(), "name 传 null 不抛异常");
        assertTrue(WordParser.tablesByName(word, "   ").isEmpty(), "全空格等于没起名，也就命中不了");
        assertTrue(WordParser.tablesByName(null, "季度明细").isEmpty(), "result 传 null 不抛异常");
        assertTrue(WordParser.rowsByName(null, "季度明细").isEmpty());

        // 比对是 equals：查询串只去首尾空格，中间空格与大小写都算差别
        // （与表头比对相反 —— 表头要猜文档里的手打差异，名字是调用方自己给的，不该再猜）
        assertEquals(1, WordParser.tablesByName(word, "  季度明细  ").size(), "首尾空格会被去掉");
        assertTrue(WordParser.tablesByName(word, "季度 明细").isEmpty(), "中间空格不去");
        assertEquals(1, WordParser.tablesByName(word, "季度明细 ").size(), "尾随空格同样被去掉，等价于原名");
    }

    /* ---------------- 工具 ---------------- */

    private static Map<String, Object> parse(String sample, WordParser.Options options) throws Exception {
        if (sample.endsWith(".pdf")) {
            return PdfParser.parse(sample, options);
        }
        if (sample.endsWith(".pptx")) {
            return PptParser.parse(sample, options);
        }
        return WordParser.parse(sample, options);
    }

    /**
     * 这一组不该有 {@code name} 键（注册了但没起名，或者根本没命中带名字的注册集）。
     */
    private static void assertNoName(Map<String, Object> table, String why) {
        assertFalse(table.containsKey("name"), why + " —— 这一组不该有 name 键");
    }
}
