package cn.net.tianxi.document;

import org.junit.jupiter.api.BeforeAll;
import org.junit.jupiter.api.Test;

import java.io.File;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;
import java.util.Map;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertTrue;

/**
 * 无边框表格：{@link PdfParser} 补读流式引擎之后的行为。
 *
 * <p>样例是测试现场生成的（{@code target/generated-test-resources/}，不入库）：先用 Aspose.Words 排一份
 * docx（一张<b>清掉全部边框</b>的 5 行×3 列表 + 一张同样的<b>有线</b>表做对照 + 4 行用制表位对齐的三列正文），
 * 再导出成 PDF。走「Word 导出 PDF」这条链路是因为真实业务里的无边框表绝大多数就是这么来的，
 * PDF 文件里只剩散落的文字、一条框线都没有。
 *
 * <p>评估模式的痕迹（有 lic 时都没有）：Aspose.Words 导出 PDF 时在<b>每一页</b>盖一张 920×500 的水印
 * <b>图</b>（实测本页 {@code images} 就 1 项，文件名 {@code image-1.jpg}），另外正文段落会被截断 ——
 * 那句 {@code PROSE 一段普通正文…} 只留下开头的 {@code PROSE}，中文部分整个没了。所以本类对 {@code tables}
 * 的断言全部与 lic 无关，对 {@code text} 只断言「{@code PROSE} 在不在」这一件事（两种情况都在）。
 * 水印图还会打断跨页合并，见 {@link PdfCrossPageTableTest}。
 */
class PdfBorderlessTableTest {

    private static final File DIR = new File("target/generated-test-resources/borderless");
    private static final List<String> NBR = Arrays.asList("NBR0C0", "NBR0C1", "NBR0C2");

    private static Map<String, Object> result;

    @BeforeAll
    static void generateAndParse() throws Exception {
        if (!DIR.isDirectory() && !DIR.mkdirs()) {
            throw new IllegalStateException("建不出目录：" + DIR.getAbsolutePath());
        }
        File pdf = new File(DIR, "borderless.pdf");
        com.aspose.words.Document doc = new com.aspose.words.Document();
        com.aspose.words.DocumentBuilder b = new com.aspose.words.DocumentBuilder(doc);
        b.getFont().setSize(9);
        b.writeln("PROSE 一段普通正文，没有对齐，不该被当成表");

        com.aspose.words.Table borderless = b.startTable();
        for (int r = 0; r < 5; r++) {
            for (int c = 0; c < 3; c++) {
                b.insertCell();
                b.write("NBR" + r + "C" + c);
            }
            b.endRow();
        }
        borderless.clearBorders();
        b.endTable();
        b.writeln();

        com.aspose.words.Table ruled = b.startTable();
        for (int r = 0; r < 5; r++) {
            for (int c = 0; c < 3; c++) {
                b.insertCell();
                b.write("BDR" + r + "C" + c);
            }
            b.endRow();
        }
        int[] types = {com.aspose.words.BorderType.TOP, com.aspose.words.BorderType.BOTTOM,
                com.aspose.words.BorderType.LEFT, com.aspose.words.BorderType.RIGHT,
                com.aspose.words.BorderType.HORIZONTAL, com.aspose.words.BorderType.VERTICAL};
        for (int type : types) {
            ruled.setBorder(type, com.aspose.words.LineStyle.SINGLE, 0.5, java.awt.Color.BLACK, true);
        }
        b.endTable();
        b.writeln();

        b.getParagraphFormat().getTabStops().add(new com.aspose.words.TabStop(90.0));
        b.getParagraphFormat().getTabStops().add(new com.aspose.words.TabStop(180.0));
        b.getParagraphFormat().getTabStops().add(new com.aspose.words.TabStop(270.0));
        for (int r = 0; r < 4; r++) {
            b.writeln("TABR" + r + "C0\tTABR" + r + "C1\tTABR" + r + "C2");
        }

        doc.save(new File(DIR, "borderless.docx").getAbsolutePath());
        doc.save(pdf.getAbsolutePath(), com.aspose.words.SaveFormat.PDF);
        result = PdfParser.parse(pdf.getAbsolutePath(), WordParser.options());
    }

    /**
     * 无边框表补得回来，而且列数对：5 行×3 列 = 首行当表头 + 4 行数据，逐格与写入时相同。
     */
    @Test
    void borderlessTableIsRecoveredAsAThreeColumnGrid() {
        Map<String, Object> group = groupContaining("NBR0C0");
        assertEquals(NBR, TestSupport.headers(group), "无边框表的列名：聚类会把 3 列读成 5 列，见 flowGrid");
        List<List<String>> rows = values(group);
        assertEquals(4, rows.size(), "5 行里首行当了表头");
        for (int r = 0; r < rows.size(); r++) {
            assertEquals(Arrays.asList("NBR" + (r + 1) + "C0", "NBR" + (r + 1) + "C1", "NBR" + (r + 1) + "C2"),
                    rows.get(r), "第 " + r + " 行");
        }
    }

    /**
     * 同一页上的有线表照旧走线段引擎，位置和内容都不受补读影响。
     */
    @Test
    void ruledTableOnTheSamePageIsUnaffected() {
        Map<String, Object> group = groupContaining("BDR0C0");
        assertEquals(Arrays.asList("BDR0C0", "BDR0C1", "BDR0C2"), TestSupport.headers(group));
        assertEquals(4, values(group).size());
        assertEquals("BDR4C2", values(group).get(3).get(2), "最后一格");
    }

    /**
     * 已知误报，写死在这儿：制表位对齐的三列正文满足「不与有线表重叠 + 每行格数一致」，会被当成一张表，
     * 这段文字因此从 {@code text} 挪进 {@code tables}。这是为拿到无边框表付的代价（见
     * {@code PdfParser#flowOnlyTables}），哪天要改这条行为，先改这个断言。
     */
    @Test
    void tabAlignedParagraphsCurrentlyCountAsATable() {
        Map<String, Object> group = groupContaining("TABR0C0");
        assertEquals(Arrays.asList("TABR0C0", "TABR0C1", "TABR0C2"), TestSupport.headers(group));
        assertEquals(3, values(group).size(), "4 行里首行当了表头");
        assertTrue(String.valueOf(result.get("text")).indexOf("TABR0C0") < 0,
                "进了表格就不该在 text 里重复一遍");
    }

    /**
     * 三块内容之外不多不少：整页正好 3 组，说明补读没有在别处凭空造表。
     */
    @Test
    void nothingElseBecomesATable() {
        List<Map<String, Object>> groups = TestSupport.tables(result);
        assertEquals(3, groups.size(), "组内容：" + groups);
    }

    /**
     * 表格上方那行正文要留在 {@code text} 里。补读时「已经在表格里、正文不用再收一遍」用的是格子矩形而不是
     * 流式引擎给的整表外接矩形 —— 外接矩形比格子本身高出约 25pt，实测 {@code PROSE} 那一行正好落在
     * 这段空隙里，用外接矩形就会被当成表格内容吞掉，{@code text} 变成空串。这条断言钉住这个修正。
     */
    @Test
    void proseAboveTheBorderlessTableStaysInText() {
        assertTrue(String.valueOf(result.get("text")).contains("PROSE"),
                "相邻正文被补读的表格区域吞掉了：" + result.get("text"));
    }

    /* ---------------- 取数 ---------------- */

    /**
     * 第一格里有这段文字的那一组；找不到就说明这一张表没读出来。
     */
    private static Map<String, Object> groupContaining(String marker) {
        for (Map<String, Object> group : TestSupport.tables(result)) {
            if (!TestSupport.rows(group).isEmpty()
                    && marker.equals(TestSupport.values(TestSupport.rows(group).get(0)).get(0))) {
                return group;
            }
            if (TestSupport.headers(group).contains(marker)) {
                return group;
            }
        }
        throw new AssertionError("没有哪一组的第一个非空格是 " + marker + "；实际组名："
                + TestSupport.tables(result));
    }

    /**
     * 每组转成「行 × 列」的值列表，逐格比较用（列名按 {@code headers} 的顺序取）。
     */
    private static List<List<String>> values(Map<String, Object> group) {
        List<String> headers = TestSupport.headers(group);
        List<List<String>> rows = new ArrayList<List<String>>();
        for (Map<String, String> row : TestSupport.rows(group)) {
            List<String> cells = new ArrayList<String>();
            for (String header : headers) {
                cells.add(row.get(header));
            }
            rows.add(cells);
        }
        return rows;
    }
}
