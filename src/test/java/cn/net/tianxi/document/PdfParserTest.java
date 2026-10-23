package cn.net.tianxi.document;

import org.junit.jupiter.api.Test;

import java.io.FileInputStream;
import java.io.InputStream;
import java.util.Arrays;
import java.util.List;
import java.util.Map;

import static org.junit.jupiter.api.Assertions.*;

/**
 * {@link PdfParser} 的功能测试，样例是 {@code src/test/resources/pdf/sales-report.pdf}
 * （同一份业绩报告用真 Microsoft Word 导出的 3 页 PDF，不是 Aspose 转的，所以文件里没有评估水印）。
 *
 * <p>PDF 没有文档模型，表格全靠线段几何还原（{@code TableAbsorber}），因此这份样例覆盖的是几条与 Word 不同的路径：
 * <ul>
 *   <li><b>跨页长表并回一张</b>：31 行的台账在第 1 页画到「泰州医药城」、第 2 页从「湖州建材贸易」接着画，
 *       靠「相邻页 + 列数相同 + 左边界对齐 + 后一截找不出表头行」并成一组；</li>
 *   <li><b>纵向合并把一张表切成几块</b>：核心表的 {@code 片区} 列纵向合并，合并处没有横线，
 *       吸收器把「江苏」「安徽」那两行切成了独立的小块，再按矩形相邻聚回同一张表；</li>
 *   <li><b>两行表头在 PDF 里已经摊平</b>：Word 里「回款情况」跨两列的两行表头，导出后吸收器读到的是一行 5 格
 *       （{@code 客户名称/合同金额/已回款/回款率/备注}），所以注册的表头在 PDF 里能直接命中；</li>
 *   <li><b>正文全在表格里</b>：页面之间只用空段落撑开（空段落不进 {@code text}），所以 {@code text} 是空串。</li>
 * </ul>
 */
class PdfParserTest {

    /**
     * 静默模式实测：6 个表头组、45 行数据、1 张图片。
     */
    private static final int SILENT_GROUPS = 6;
    private static final int SILENT_ROWS = 45;

    private static Map<String, Object> silent() throws Exception {
        return PdfParser.parse(TestSupport.PDF, WordParser.options());
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
        WordParser.Options[] all = {
                WordParser.options(),
                TestSupport.registered(),
                TestSupport.registered().fullRow(WordParser.FullRow.FIRST),
                TestSupport.registered().fullRow(WordParser.FullRow.DROP),
                WordParser.options().fullRow(WordParser.FullRow.DROP),
                WordParser.options().keywords("客户"),
        };
        for (WordParser.Options options : all) {
            TestSupport.assertHeadersMatchRowKeys(PdfParser.parse(TestSupport.PDF, options));
        }
    }

    @Test
    void textIsEmptyBecauseEveryParagraphLivesInATable() throws Exception {
        // 样例的正文全部落在表格里，页面之间只有撑开用的空段落；空段落与评估水印都不进 text。
        // 这是 PDF 的正常输出而不是「什么都没读到」：表格与图片都读全了（见下面各条）
        assertEquals("", TestSupport.text(silent()));
    }

    /* ---------------- 跨页长表 ---------------- */

    @Test
    void ledgerSplitAcrossTwoPagesIsMergedBackIntoOneGroup() throws Exception {
        Map<String, Object> result = silent();
        Map<String, Object> ledger = TestSupport.table(result, 1);
        List<Map<String, String>> rows = TestSupport.rows(ledger);

        assertEquals(Arrays.asList("客户名称", "合同金额", "已回款", "回款率", "备注"), TestSupport.headers(ledger));
        assertEquals(31, rows.size(), "30 行客户 + 1 行合计，跨页断开后要并回一张");
        // 第 1 页的第一行与第 2 页的最后一行在同一组里，说明真的并起来了（没并的话会变成两组、各带一个表头）
        assertEquals("上海浦东城建集团", rows.get(0).get("客户名称"));
        assertEquals("泰州医药城", rows.get(19).get("客户名称"), "第 1 页画到这一行为止");
        assertEquals("湖州建材贸易", rows.get(20).get("客户名称"), "第 2 页从这一行接着画");
        assertEquals("池州矿业开发", rows.get(29).get("客户名称"));
        assertEquals(Arrays.asList("合计", "", "", "86%", ""), TestSupport.values(rows.get(30)));
    }

    @Test
    void exportedPdfLosesTheCachedSumFieldResult() throws Exception {
        // 导出这份 PDF 用的加长版源 docx（30 家客户，源文件不在仓库里）合计行两格是 SUM(ABOVE) 域的缓存结果
        // （探针核对过 Aspose.Words 读到 2582.30/2226.60），Word 导出 PDF 时重算了这个域，
        // 导出的 PDF 里这两格只剩一个空格 —— 解析如实读出空串。
        // 仓库里的 docx 样例台账是 12 家客户，合计读的是缓存值 2096.30/1796.10（见 WordParserTest）。
        // 结论：域结果在 docx 里是缓存的、读得到；转成 PDF 就不一定还在，要金额请用 docx 那一份
        assertEquals(Arrays.asList("合计", "", "", "86%", ""), TestSupport.values(silent(), 1, 30));
    }

    /* ---------------- 纵向合并被切成几块 ---------------- */

    @Test
    void coreTableRowsCutByVerticalMergeAreClusteredBack() throws Exception {
        // 吸收器把「江苏」「安徽」两行切成了独立小块（合并处没有横线），按矩形相邻聚回来后 4 行都在同一组，
        // 片区列的值也从合并起始行回填了 —— 与 Word 读到的完全一致（见 CrossFormatTest）
        Map<String, Object> result = silent();
        assertEquals(Arrays.asList("片区", "区域", "销售额", "回款率", "销售额-5", "回款率-6", "备注"),
                TestSupport.headers(TestSupport.table(result, 0)));
        assertEquals(Arrays.asList("华东大区", "江苏", "986", "88%", "1120", "90%", ""),
                TestSupport.values(result, 0, 1));
        assertEquals(Arrays.asList("华东南区", "安徽", "540", "80%", "616", "83%", ""),
                TestSupport.values(result, 0, 3));
        assertEquals(4, TestSupport.rows(TestSupport.table(result, 0)).size());
    }

    /* ---------------- 其余三张表 ---------------- */

    @Test
    void ticketTableKeepsTheRowThatWasCutIntoItsOwnBlock() throws Exception {
        // 「硬件故障 | 12 | 更换配件后复测通过」那一行在 PDF 里也被切成了独立小块（首列跨行合并），
        // 聚回来后工单表仍是完整的 4 行
        Map<String, Object> result = silent();
        assertEquals(Arrays.asList("工单类型", "工单数", "处理说明"), TestSupport.headers(TestSupport.table(result, 2)));
        assertEquals(Arrays.asList("硬件故障", "42", ""), TestSupport.values(result, 2, 0));
        assertEquals(Arrays.asList("硬件故障", "12", "更换配件后复测通过"), TestSupport.values(result, 2, 1));
        assertEquals(Arrays.asList("软件异常", "27", "远程处理，平均响应 2 小时"), TestSupport.values(result, 2, 2));
        assertEquals(Arrays.asList("现场支持", "15", "派工上门，覆盖江浙沪皖"), TestSupport.values(result, 2, 3));
    }

    @Test
    void separatorRowBeforeAnyHeaderStartsANumberedGroup() throws Exception {
        Map<String, Object> result = silent();
        assertEquals(Arrays.asList("列1", "列2", "列3", "列4"), TestSupport.headers(TestSupport.table(result, 3)));
        assertEquals(Arrays.asList("2025 年度（已结账）", "", "", ""), TestSupport.values(result, 3, 0));
    }

    @Test
    void goalTableSplitsIntoTwoGroups() throws Exception {
        Map<String, Object> result = silent();
        assertEquals(Arrays.asList("片区", "年度目标", "完成率", "列4"),
                TestSupport.headers(TestSupport.table(result, 4)));
        assertEquals(Arrays.asList("华东大区", "2400.00", "101%", "达标"), TestSupport.values(result, 4, 0));
        assertEquals(Arrays.asList("华东南区", "1500.00", "93%", "未达标"), TestSupport.values(result, 4, 1));
        assertEquals(Arrays.asList("2026 年度（预算下达）", "", "", ""), TestSupport.values(result, 4, 2));
        assertEquals(Arrays.asList("华东大区", "2800.00", "—", ""), TestSupport.values(result, 5, 0));
        assertEquals(Arrays.asList("华东南区", "1750.00", "—", ""), TestSupport.values(result, 5, 1));
    }

    /* ---------------- FullRow 三种去向 ---------------- */

    @Test
    void dropRemovesEverySeparatorRow() throws Exception {
        Map<String, Object> result =
                PdfParser.parse(TestSupport.PDF, TestSupport.registered().fullRow(WordParser.FullRow.DROP));

        assertEquals(5, TestSupport.tables(result).size(), "组数：「2025 年度（已结账）」那组整组消失");
        assertEquals(43, TestSupport.totalRows(result), "行数：两行分组小标题都没了");
        String json = PdfParser.toJson(result);
        assertFalse(json.contains("2025 年度（已结账）"), json);
        assertFalse(json.contains("2026 年度（预算下达）"), json);
        TestSupport.assertHeadersMatchRowKeys(result);
    }

    @Test
    void firstRoutesEverySeparatorRowOfATableToThatTablesFirstGroup() throws Exception {
        Map<String, Object> first =
                PdfParser.parse(TestSupport.PDF, WordParser.options().fullRow(WordParser.FullRow.FIRST));

        assertEquals(SILENT_GROUPS, TestSupport.tables(first).size(), "组数不变");
        assertEquals(SILENT_ROWS, TestSupport.totalRows(first), "行数不变");
        assertEquals(2, TestSupport.rows(TestSupport.table(first, 3)).size(), "两行分组小标题都归到序号列那组");
        assertEquals(Arrays.asList("2026 年度（预算下达）", "", "", ""), TestSupport.values(first, 3, 1));
        assertEquals(2, TestSupport.rows(TestSupport.table(first, 4)).size(), "目标表那组只剩数据行");
    }

    /* ---------------- 注册表头 ---------------- */

    @Test
    void registeringTheTrueHeadersChangesNothingHere() throws Exception {
        // PDF 里两行表头已经摊平成一行 5 格，注册的台账表头因此能直接命中；命不中的只有「故障部件/工单数」那组
        // （嵌套小表在导出成 PDF 时已经不在了）。结果与静默模式逐字节相同
        assertEquals(PdfParser.toJson(silent()),
                PdfParser.toJson(PdfParser.parse(TestSupport.PDF, TestSupport.registered())));
    }

    @Test
    void registeringHeaderEvidenceTurnsOffTheFirstRowFallback() throws Exception {
        // 静默模式（既没注册表头、也没给关键字）允许「首行兜底」当表头；一旦调用方给了表头依据，兜底就关掉，
        // 没命中的表首行按数据行处理、列名补成序号。这是有意的：给了依据就以依据为准，
        // 免得把一行数据当表头（PDF 没有 Word 的 tblHeader 标记可依赖）
        Map<String, Object> result = PdfParser.parse(TestSupport.PDF,
                WordParser.options().headers("客户名称", "合同金额", "已回款", "回款率", "备注"));
        Map<String, Object> core = TestSupport.table(result, 0);

        assertTrue(TestSupport.headers(core).get(0).startsWith("列"), "核心表的表头行没命中注册表头：" + core);
        assertEquals("片区", TestSupport.values(result, 0, 0).get(0), "表头行本身按数据行留在结果里");
        assertEquals(5, TestSupport.rows(core).size(), "表头行 + 4 行数据");
        // 台账表头命中了注册的那组，列名用注册的原名
        assertEquals(Arrays.asList("客户名称", "合同金额", "已回款", "回款率", "备注"),
                TestSupport.headers(TestSupport.table(result, 1)));
    }

    @Test
    void unmatchedHeaderGroupIsReportedOnStderr() throws Exception {
        String stderr = TestSupport.captureStderr(() ->
                PdfParser.parse(TestSupport.PDF, TestSupport.registered()));

        assertTrue(stderr.contains("故障部件"), "嵌套表在 PDF 里不存在，该组应当被提示：" + stderr);
        assertFalse(stderr.contains("客户名称"), "台账表头在 PDF 里能命中：" + stderr);
    }

    /* ---------------- 图片 ---------------- */

    @Test
    void imageWasReEncodedToJpegByTheWordExport() throws Exception {
        // 同一张看板截图，docx/pptx 里是 1024×768 的 PNG（670939 字节），Word 导出 PDF 时重编成 699×524 的 JPEG。
        // 解析照实报告文件里的字节，不做还原：contentId 因此与另两种格式不同
        Map<String, Object> image = TestSupport.image(silent(), 0);

        assertEquals("image-1.jpg", image.get("fileName"));
        assertEquals("JPEG", image.get("type"));
        assertEquals(699, image.get("width"));
        assertEquals(524, image.get("height"));
        assertEquals(38668, image.get("size"));
        assertEquals("80c8e65be48c09a0f173845e4dfc072bc541fb68", image.get("contentId"));
        TestSupport.assertImageMeta(image);
    }

    /* ---------------- 入口 ---------------- */

    @Test
    void varargsEntryWithoutKeywordsIsSilentMode() throws Exception {
        assertEquals(PdfParser.toJson(silent()), PdfParser.toJson(PdfParser.parse(TestSupport.PDF)));
    }

    @Test
    void parsesFromInputStream() throws Exception {
        InputStream stream = new FileInputStream(TestSupport.PDF);
        try {
            assertEquals(PdfParser.toJson(silent()),
                    PdfParser.toJson(PdfParser.parse(stream, WordParser.options())));
        } finally {
            stream.close();
        }
    }
}
