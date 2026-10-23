package cn.net.tianxi.document;

import org.junit.jupiter.api.BeforeAll;
import org.junit.jupiter.api.Test;

import java.io.File;
import java.util.Arrays;
import java.util.List;
import java.util.Map;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertTrue;

/**
 * 现造 PDF 上的合并单元格与两级表头：{@code PdfParser#toGrid} 的「跨整行合并行」判据
 * （{@code origins == 1 && !blank}）和 {@link Tables} 的多级表头合并，与 {@link WordParser} 逐格对齐。
 *
 * <p>样例是测试现场生成的（{@code target/generated-test-resources/mergedcell/}，不入库）：先用
 * {@link TestSupport#docx} 拼一份最小 docx（三张表，都带 {@code w:tblBorders} 六条单线），
 * 再用 Aspose.Words 导出成 PDF。之所以自己拼 OOXML 而不用 {@code DocumentBuilder} 排：实测 builder 的
 * {@code setVerticalMerge}/{@code setHorizontalMerge} 存出来的 docx 里根本没有 {@code vMerge}/
 * {@code gridSpan}（合并没生效，横向合并那行还把网格撑成了 6 列），拿它当样例测不出合并这条路径。
 *
 * <p>导出的是「Aspose.Words 转的 PDF」，所以带评估水印（每页一张 920×500 的图，见
 * {@link PdfCrossPageTableTest}）。单页上这张图不影响表格抽取，{@code tables} 与不带图的解析结果逐字节相同 ——
 * 本类因此不需要先删图；只有跨页合并的样例才必须删。
 */
class PdfMergedCellTableTest {

    private static final File DIR = new File("target/generated-test-resources/mergedcell");

    private static Map<String, Object> word;
    private static Map<String, Object> pdf;
    private static Map<String, Object> pdfRegistered;

    /** {@link TestSupport#tbl} 的表格是不带框线的（{@code w:tblStyle} 指了个不存在的样式），这里补上实线边框。 */
    private static final String BORDERS = "<w:tblBorders>"
            + "<w:top w:val=\"single\" w:sz=\"6\" w:color=\"000000\"/>"
            + "<w:left w:val=\"single\" w:sz=\"6\" w:color=\"000000\"/>"
            + "<w:bottom w:val=\"single\" w:sz=\"6\" w:color=\"000000\"/>"
            + "<w:right w:val=\"single\" w:sz=\"6\" w:color=\"000000\"/>"
            + "<w:insideH w:val=\"single\" w:sz=\"6\" w:color=\"000000\"/>"
            + "<w:insideV w:val=\"single\" w:sz=\"6\" w:color=\"000000\"/>"
            + "</w:tblBorders>";

    private static String tbl(int columns, String... rows) {
        return TestSupport.tbl(columns, rows).replace("</w:tblPr>", BORDERS + "</w:tblPr>");
    }

    @BeforeAll
    static void generateAndParse() throws Exception {
        if (!DIR.isDirectory() && !DIR.mkdirs()) {
            throw new IllegalStateException("建不出目录：" + DIR.getAbsolutePath());
        }
        File docx = TestSupport.docx(DIR, "merged.docx",
                // V：列 0 的第 2、3 行纵向合并，第 3 行其余两格是自己的数据
                tbl(3,
                        TestSupport.tr(TestSupport.tc("V0C0"), TestSupport.tc("V0C1"), TestSupport.tc("V0C2")),
                        TestSupport.tr(TestSupport.tcVerticalStart("V1C0"), TestSupport.tc("V1C1"),
                                TestSupport.tc("V1C2")),
                        TestSupport.tr(TestSupport.tcVerticalContinue(), TestSupport.tc("V2C1"),
                                TestSupport.tc("V2C2")),
                        TestSupport.tr(TestSupport.tc("V3C0"), TestSupport.tc("V3C1"), TestSupport.tc("V3C2"))),
                TestSupport.emptyP(),
                // S：第 3 行是跨整行合并、而且一个字都没写
                tbl(3,
                        TestSupport.tr(TestSupport.tc("S0C0"), TestSupport.tc("S0C1"), TestSupport.tc("S0C2")),
                        TestSupport.tr(TestSupport.tc("S1C0"), TestSupport.tc("S1C1"), TestSupport.tc("S1C2")),
                        TestSupport.tr(TestSupport.tcSpan("", 3)),
                        TestSupport.tr(TestSupport.tc("S3C0"), TestSupport.tc("S3C1"), TestSupport.tc("S3C2"))),
                TestSupport.emptyP(),
                // T：两级表头，上层第二格横向跨两列
                tbl(3,
                        TestSupport.tr(TestSupport.tc("T0C0"), TestSupport.tcSpan("T0C1", 2)),
                        TestSupport.tr(TestSupport.tc("T1C0"), TestSupport.tc("T1C1"), TestSupport.tc("T1C2")),
                        TestSupport.tr(TestSupport.tc("T2C0"), TestSupport.tc("T2C1"), TestSupport.tc("T2C2"))));

        File pdfFile = new File(DIR, "merged.pdf");
        new com.aspose.words.Document(docx.getAbsolutePath())
                .save(pdfFile.getAbsolutePath(), com.aspose.words.SaveFormat.PDF);

        word = WordParser.parse(docx.getPath(), WordParser.options());
        pdf = PdfParser.parse(pdfFile.getAbsolutePath(), WordParser.options());
        pdfRegistered = PdfParser.parse(pdfFile.getAbsolutePath(), WordParser.options()
                .headers("T1C0", "T1C1", "T1C2").name("下层")
                .headers("T0C0", "T0C1", "T0C1").name("上层"));
    }

    /**
     * 纵向合并的<b>延续行</b>不是「跨整行的分组行」：合并只发生在列 0，这一行其余两格是自己的数据，
     * 所以它照常进 {@code rows}，列 0 回填起始格的值。docx 与 PDF 逐格相同。
     */
    @Test
    void verticallyMergedDataRowIsNotASeparatorRow() {
        for (Map<String, Object> result : Arrays.asList(word, pdf)) {
            assertEquals(3, TestSupport.tables(result).size(), "三张表就是三组：" + TestSupport.tables(result));
            Map<String, Object> group = TestSupport.table(result, 0);
            assertEquals(Arrays.asList("V0C0", "V0C1", "V0C2"), TestSupport.headers(group), "列名");
            assertEquals(3, TestSupport.rows(group).size(), "延续行没被当成分组行切走");
            assertEquals(Arrays.asList("V1C0", "V2C1", "V2C2"), TestSupport.values(TestSupport.rows(group).get(1)),
                    "列 0 回填起始格，其余两格是自己的值");
        }
    }

    /**
     * 跨整行合并、而且<b>整行空白</b>的那一行不算分组行（{@code !blank} 这半条判据），
     * 它作为一行全空的数据行留在 {@code rows} 里 —— 表没有被切成两组，也没有少一行。
     */
    @Test
    void blankFullWidthMergedRowStaysAnOrdinaryDataRow() {
        for (Map<String, Object> result : Arrays.asList(word, pdf)) {
            Map<String, Object> group = TestSupport.table(result, 1);
            assertEquals(Arrays.asList("S0C0", "S0C1", "S0C2"), TestSupport.headers(group), "列名");
            assertEquals(3, TestSupport.rows(group).size(), "空白合并行没有结束表头区、也没有另起一组");
            assertEquals(Arrays.asList("", "", ""), TestSupport.values(TestSupport.rows(group).get(1)), "空白行原样留着");
            assertEquals(Arrays.asList("S3C0", "S3C1", "S3C2"), TestSupport.values(TestSupport.rows(group).get(2)), "行序没乱");
        }
    }

    /** 两份文件读出来的表格部分逐字节相同（只有 docx 那边没有水印图，见 {@link #watermarkAddsAnImageOnly()}）。 */
    @Test
    void theTwoFormatsAgreeCellForCell() throws Exception {
        assertEquals(PdfParser.toJson(TestSupport.tables(word)), PdfParser.toJson(TestSupport.tables(pdf)),
                "同一份内容 docx 与 PDF 的 tables 应当逐字节相同");
    }

    /**
     * 两级表头在现造 PDF 上的两种读法。静默模式：上层那行当列名（横向合并摊平成
     * {@code T0C1}/{@code T0C1-3}），下层那行落进数据。注册模式：上下两行都命中（PDF 的网格里上层那行
     * 就是 {@code T0C0/T0C1/T0C1} 三格，逐字等于注册名），合并后列名取<b>下层</b>那行，
     * {@code name} 取<b>先命中</b>的那份注册（{@code 上层}）—— 与 {@link Tables} 里
     * 「不去抢第一个名字为准」那条注释说的是同一件事。
     */
    @Test
    void twoLevelHeaderReadsBothWaysAndNamesByFirstMatch() {
        Map<String, Object> silent = TestSupport.table(pdf, 2);
        assertEquals(Arrays.asList("T0C0", "T0C1", "T0C1-3"), TestSupport.headers(silent), "静默模式取上层那行");
        assertEquals(2, TestSupport.rows(silent).size(), "下层那行在静默模式下是数据");
        assertEquals(Arrays.asList("T1C0", "T1C1", "T1C2"), TestSupport.values(TestSupport.rows(silent).get(0)));

        Map<String, Object> registered = TestSupport.table(pdfRegistered, 2);
        assertEquals(Arrays.asList("T1C0", "T1C1", "T1C2"), TestSupport.headers(registered), "合并后取下层列名");
        assertEquals(1, TestSupport.rows(registered).size(), "两行都算表头，只剩 T2 一行数据");
        assertEquals("上层", registered.get("name"), "先命中的那份注册赢");
        assertEquals(3, TestSupport.tables(pdfRegistered).size(), "注册不改变组数");
        assertFalse(TestSupport.table(pdfRegistered, 0).containsKey("name"), "没命中注册的那两组不该带 name");
        assertFalse(TestSupport.table(pdfRegistered, 1).containsKey("name"), "同上");
    }

    /** 评估水印只在 {@code images} 里多一项，表格抽取完全不受影响（单页上它排在表格之后也没关系）。 */
    @Test
    void watermarkAddsAnImageOnly() {
        assertEquals(0, TestSupport.images(word).size(), "docx 那边没有水印图");
        List<Map<String, Object>> images = TestSupport.images(pdf);
        assertEquals(1, images.size(), "PDF 这边一张水印图");
        assertEquals(920, images.get(0).get("width"), "水印宽");
        assertEquals(500, images.get(0).get("height"), "水印高");
        assertTrue(String.valueOf(pdf.get("text")).isEmpty(), "水印文字不进 text：" + pdf.get("text"));
    }
}
