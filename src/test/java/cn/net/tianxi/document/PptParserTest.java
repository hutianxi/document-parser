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
 * {@link PptParser} 的功能测试，样例是 {@code src/test/resources/ppt/sales-report.pptx}
 * （同一份业绩报告的 3 页幻灯片，手工拼的 OOXML：合并单元格属性写在 {@code <a:tc>} 上）。
 *
 * <p>三页分别是：
 * <ol>
 *   <li>标题文本框 + 核心表（{@code 片区} 列跨行合并）+ 三个 KPI 文本框；</li>
 *   <li>台账表（两行表头：{@code 回款情况} 跨两列、{@code 客户名称}/{@code 回款率}/{@code 备注} 跨两行）
 *       + 一个<b>组合形状</b>（里面三个文本框）+ 看板截图；</li>
 *   <li>年度目标表（跨整列合并的分组行）+ 工单表。</li>
 * </ol>
 *
 * <h3>评估模式对这份样例的影响</h3>
 * 没有 lic 时 Aspose.Slides 读文本框只给<b>第一段</b>、且截到 5 个字符再追加提示（表格单元格走的也是文本框
 * 那一条，{@code PptParser} 对它同样调 {@code stripEvaluationNotice}），提示会被去掉。
 * 样例因此把每处文字都写成一段、不超过 5 个字符，有没有 lic 读到的都是同一串
 * （{@link #everyTextFrameInTheSampleStaysWithinTheEvaluationLimit} 守着这条）。
 * 代价是它验证不了「长文本能不能完整读出来」—— 那一条由 docx 样例覆盖（它的段落有 60 多个字符，读出来是完整的）。
 */
class PptParserTest {

    /**
     * 静默模式实测：6 个表头组、22 行数据、1 张图片。
     */
    private static final int SILENT_GROUPS = 6;
    private static final int SILENT_ROWS = 22;

    /**
     * 9 个文本框的文字，按形状在页面里出现的顺序。
     */
    private static final String TEXT = String.join("\n",
            "2025年", "销售业绩", "华东大区", "回款率", "86%", "达标", "数据来源", "销售系统", "财务对账");

    private static Map<String, Object> silent() throws Exception {
        return PptParser.parse(TestSupport.PPT, WordParser.options());
    }

    /* ---------------- 整体形状与契约 ---------------- */

    @Test
    void readsWholeDeck() throws Exception {
        Map<String, Object> result = silent();

        assertEquals(SILENT_GROUPS, TestSupport.tables(result).size(), "表头组数");
        assertEquals(SILENT_ROWS, TestSupport.totalRows(result), "数据行数");
        assertEquals(1, TestSupport.images(result).size(), "图片数");
        TestSupport.assertHeadersMatchRowKeys(result);
    }

    @Test
    void everyOptionCombinationKeepsTheContract() throws Exception {
        WordParser.Options[] all = {
                WordParser.options(),
                TestSupport.registered(),
                TestSupport.registered().fullRow(WordParser.FullRow.FIRST),
                TestSupport.registered().fullRow(WordParser.FullRow.DROP),
                WordParser.options().keywords("片区"),
        };
        for (WordParser.Options options : all) {
            TestSupport.assertHeadersMatchRowKeys(PptParser.parse(TestSupport.PPT, options));
        }
    }

    /* ---------------- 形状文字 ---------------- */

    @Test
    void readsShapeTextInPageOrder() throws Exception {
        assertEquals(TEXT, TestSupport.text(silent()));
    }

    @Test
    void recursesIntoGroupShapes() throws Exception {
        // 「数据来源／销售系统／财务对账」三个文本框在一个组合形状里：不递归展开就整组丢掉
        String text = TestSupport.text(silent());
        assertTrue(text.contains("数据来源"), text);
        assertTrue(text.contains("销售系统"), text);
        assertTrue(text.contains("财务对账"), text);
    }

    @Test
    void tableTextDoesNotLeakIntoText() throws Exception {
        String text = TestSupport.text(silent());
        assertFalse(text.contains("上海城建"), "表格内容不该进 text：" + text);
        assertFalse(text.contains("工单类型"), "表格内容不该进 text：" + text);
    }

    /* ---------------- 表 ---------------- */

    @Test
    void coreTableFillsTheRowSpannedFirstColumn() throws Exception {
        Map<String, Object> result = silent();
        assertEquals(Arrays.asList("片区", "区域", "销售额", "回款率", "销售额-5", "回款率-6", "备注"),
                TestSupport.headers(TestSupport.table(result, 0)));
        assertEquals(Arrays.asList("华东大区", "上海", "1280", "92%", "1465", "95%", "超额完成"),
                TestSupport.values(result, 0, 0));
        assertEquals(Arrays.asList("华东大区", "江苏", "986", "88%", "1120", "90%", ""),
                TestSupport.values(result, 0, 1), "跨行合并的片区列要回填起始格的值");
        assertEquals(Arrays.asList("华东南区", "安徽", "540", "80%", "616", "83%", ""),
                TestSupport.values(result, 0, 3));
    }

    @Test
    void silentModeCannotMergeTheTwoLevelHeader() throws Exception {
        // 已知取舍：pptx 的表格没有「标题行」标记可用（本样例只声明了首行是标题），
        // 静默模式下第二行表头被当成数据行，列名因此是上层的「回款情况」加去重后缀。
        // 注册表头就能正确合并（见下一条），这也是「PPT 建议注册表头」的理由
        Map<String, Object> result = silent();
        Map<String, Object> ledger = TestSupport.table(result, 1);

        assertEquals(Arrays.asList("客户名称", "回款情况", "回款情况-3", "回款率", "备注"),
                TestSupport.headers(ledger));
        assertEquals(8, TestSupport.rows(ledger).size(), "下层表头行被当成了一行数据");
        assertEquals(Arrays.asList("客户名称", "合同金额", "已回款", "回款率", "备注"),
                TestSupport.values(result, 1, 0));
        assertEquals(Arrays.asList("上海城建", "486", "486", "100%", "已结清"), TestSupport.values(result, 1, 1));
    }

    @Test
    void registeredHeadersMergeTheTwoLevelHeader() throws Exception {
        // 下层表头行摊平后正好是「客户名称/合同金额/已回款/回款率/备注」（跨行合并的列回填了上层的名字），
        // 命中注册表头 → 与上一行连续表头合并，7 行数据不再多出一行表头
        Map<String, Object> result = PptParser.parse(TestSupport.PPT, TestSupport.registered());
        Map<String, Object> ledger = TestSupport.table(result, 1);

        assertEquals(Arrays.asList("客户名称", "合同金额", "已回款", "回款率", "备注"), TestSupport.headers(ledger));
        assertEquals(7, TestSupport.rows(ledger).size());
        assertEquals(Arrays.asList("上海城建", "486", "486", "100%", "已结清"), TestSupport.values(result, 1, 0));
        assertEquals(Arrays.asList("合计", "1586", "1384", "87%", ""),
                TestSupport.values(result, 1, 6));
        assertEquals(21, TestSupport.totalRows(result), "比静默模式少一行（那行是下层表头）");
    }

    @Test
    void everyTextFrameInTheSampleStaysWithinTheEvaluationLimit() throws Exception {
        // 样例自检：评估模式下 Aspose.Slides 读文本框（表格单元格也是文本框）只给第一段、截到 5 个字符，
        // 所以这份 pptx 里所有文字都写成了 ≤5 个字符的一段 —— 这样有 lic / 没 lic 读到的才是同一串。
        // 这条断言是给改样例的人看的：真要塞长文本，下面所有期望值都得按有 lic 的环境重测
        Map<String, Object> result = silent();
        for (String line : TestSupport.text(result).split("\n")) {
            assertTrue(line.length() <= 5, "文本框超过 5 个字符，评估模式会截断：" + line);
        }
        for (Map<String, Object> table : TestSupport.tables(result)) {
            for (Map<String, String> row : TestSupport.rows(table)) {
                for (String cell : row.values()) {
                    assertTrue(cell.length() <= 5, "单元格超过 5 个字符，评估模式会截断：" + cell);
                }
            }
        }
    }

    @Test
    void separatorRowBeforeAnyHeaderStartsANumberedGroup() throws Exception {
        Map<String, Object> result = silent();
        assertEquals(Arrays.asList("列1", "列2", "列3", "列4"), TestSupport.headers(TestSupport.table(result, 2)));
        assertEquals(Arrays.asList("2025年", "", "", ""), TestSupport.values(result, 2, 0));
    }

    @Test
    void goalAndTicketTables() throws Exception {
        Map<String, Object> result = silent();
        assertEquals(Arrays.asList("片区", "年度目标", "完成率", "列4"),
                TestSupport.headers(TestSupport.table(result, 3)));
        assertEquals(Arrays.asList("华东大区", "2400", "101%", "达标"), TestSupport.values(result, 3, 0));
        assertEquals(Arrays.asList("华东南区", "1500", "93%", "未达标"), TestSupport.values(result, 3, 1));
        assertEquals(Arrays.asList("2026年", "", "", ""), TestSupport.values(result, 3, 2));
        assertEquals(Arrays.asList("华东大区", "2800", "—", ""), TestSupport.values(result, 4, 0));
        assertEquals(Arrays.asList("华东南区", "1750", "—", ""), TestSupport.values(result, 4, 1));

        assertEquals(Arrays.asList("工单类型", "工单数", "处理说明"),
                TestSupport.headers(TestSupport.table(result, 5)));
        List<Map<String, String>> tickets = TestSupport.rows(TestSupport.table(result, 5));
        assertEquals(4, tickets.size());
        assertEquals(Arrays.asList("硬件故障", "42", ""), TestSupport.values(tickets.get(0)));
        assertEquals(Arrays.asList("现场支持", "15", "派工上门"), TestSupport.values(tickets.get(3)));
    }

    @Test
    void verticallyMergedDataRowIsNotAMistakenSeparator(@TempDir File directory) throws Exception {
        // 回归：origins 曾经只数「合并起始格」，纵向合并的延续格因此不算，这一行只剩一个格子、被当成分组小标题，
        // KEEP 下第二行的值被抹成空串、DROP 下整行被删。判据改成只按横向延续格判定，与 docx（WordParser.Grid）
        // 和 pdf（PdfParser.toGrid）一致 —— 同样这份表格三种格式读出来应当逐格相同
        File file = TestSupport.pptx(directory, "vmerge.pptx", TestSupport.pTbl(2,
                TestSupport.pTr(TestSupport.pTc("列A"), TestSupport.pTc("列B")),
                TestSupport.pTr(TestSupport.pTcVerticalStart("第一", 2), TestSupport.pTc("值一")),
                TestSupport.pTr(TestSupport.pTcVerticalContinue(), TestSupport.pTc("值二"))));

        Map<String, Object> keep = PptParser.parse(file.getPath(), WordParser.options());
        Map<String, Object> drop = PptParser.parse(file.getPath(),
                WordParser.options().fullRow(WordParser.FullRow.DROP));

        assertEquals(Arrays.asList("列A", "列B"), TestSupport.headers(TestSupport.table(keep, 0)));
        assertEquals(2, TestSupport.rows(TestSupport.table(keep, 0)).size(), "两行数据都还在");
        assertEquals(Arrays.asList("第一", "值一"), TestSupport.values(keep, 0, 0));
        assertEquals(Arrays.asList("第一", "值二"), TestSupport.values(keep, 0, 1), "延续格回填起始格的值");
        assertEquals(2, TestSupport.rows(TestSupport.table(drop, 0)).size(), "DROP 也不该删掉数据行");
        assertEquals(Arrays.asList("第一", "值二"), TestSupport.values(drop, 0, 1));
    }

    @Test
    void blankFullWidthMergedRowIsNotASeparator(@TempDir File directory) throws Exception {
        // 与 docx／pdf 同一条判据：整行合并但一格内容都没有的空白行不是分组行，
        // 否则紧跟它的那一行会被当成表头（一行数据没了、表还被切成两组）
        File file = TestSupport.pptx(directory, "blank-full-row.pptx", TestSupport.pTbl(2,
                TestSupport.pTr(TestSupport.pTc("列A"), TestSupport.pTc("列B")),
                TestSupport.pTr(TestSupport.pTc("甲"), TestSupport.pTc("乙")),
                TestSupport.pTr(TestSupport.pTcSpan("", 2), TestSupport.pTcHMergeContinue()),
                TestSupport.pTr(TestSupport.pTc("丙"), TestSupport.pTc("丁"))));

        Map<String, Object> result = PptParser.parse(file.getPath(), WordParser.options());

        assertEquals(1, TestSupport.tables(result).size(), "空白行不该把表切成两组");
        assertEquals(Arrays.asList("列A", "列B"), TestSupport.headers(TestSupport.table(result, 0)), "丙丁不该被当成表头");
        assertEquals(3, TestSupport.rows(TestSupport.table(result, 0)).size());
        assertEquals(Arrays.asList("", ""), TestSupport.values(result, 0, 1), "空白行留成一行空值");
        assertEquals(Arrays.asList("丙", "丁"), TestSupport.values(result, 0, 2));
    }

    /* ---------------- FullRow 三种去向 ---------------- */

    @Test
    void dropRemovesEverySeparatorRow() throws Exception {
        Map<String, Object> result =
                PptParser.parse(TestSupport.PPT, TestSupport.registered().fullRow(WordParser.FullRow.DROP));

        assertEquals(5, TestSupport.tables(result).size(), "组数");
        assertEquals(19, TestSupport.totalRows(result), "行数");
        // 只看 tables：正文里也有「2025年」那个文本框，整串 JSON 里搜不到区别
        String tables = PptParser.toJson(TestSupport.tables(result));
        assertFalse(tables.contains("2025年"), "分组行应该被丢掉：" + tables);
        assertFalse(tables.contains("2026年"), "分组行应该被丢掉：" + tables);
        TestSupport.assertHeadersMatchRowKeys(result);
    }

    @Test
    void firstRoutesEverySeparatorRowOfATableToThatTablesFirstGroup() throws Exception {
        Map<String, Object> first =
                PptParser.parse(TestSupport.PPT, WordParser.options().fullRow(WordParser.FullRow.FIRST));

        assertEquals(SILENT_GROUPS, TestSupport.tables(first).size(), "组数不变");
        assertEquals(SILENT_ROWS, TestSupport.totalRows(first), "行数不变");
        assertEquals(2, TestSupport.rows(TestSupport.table(first, 2)).size(), "两行分组小标题都归到序号列那组");
        assertEquals(Arrays.asList("2026年", "", "", ""), TestSupport.values(first, 2, 1));
        assertEquals(2, TestSupport.rows(TestSupport.table(first, 3)).size(), "目标表那组只剩数据行");
    }

    /* ---------------- 图片与提示 ---------------- */

    @Test
    void imageIsTheSamePictureAsInTheWordSample() throws Exception {
        Map<String, Object> image = TestSupport.image(silent(), 0);

        assertEquals("image-1.png", image.get("fileName"));
        assertEquals("PNG", image.get("type"));
        assertEquals(1024, image.get("width"));
        assertEquals(768, image.get("height"));
        assertEquals(670939, image.get("size"));
        assertEquals("487e915db2c53907e8c1551aeaffea235423210c", image.get("contentId"));
        TestSupport.assertImageMeta(image);
    }

    @Test
    void unmatchedHeaderGroupIsReportedOnStderr() throws Exception {
        String stderr = TestSupport.captureStderr(() ->
                PptParser.parse(TestSupport.PPT, TestSupport.registered()));

        assertTrue(stderr.contains("故障部件"), "幻灯片里没有那张嵌套小表，该组应当被提示：" + stderr);
        assertFalse(stderr.contains("客户名称"), "台账表头该命中：" + stderr);
    }

    /* ---------------- 入口 ---------------- */

    @Test
    void varargsEntryWithoutKeywordsIsSilentMode() throws Exception {
        assertEquals(PptParser.toJson(silent()), PptParser.toJson(PptParser.parse(TestSupport.PPT)));
    }

    @Test
    void parsesFromInputStream() throws Exception {
        InputStream stream = new FileInputStream(TestSupport.PPT);
        try {
            assertEquals(PptParser.toJson(silent()),
                    PptParser.toJson(PptParser.parse(stream, WordParser.options())));
        } finally {
            stream.close();
        }
    }
}
