package cn.net.tianxi.document;

import org.junit.jupiter.api.BeforeAll;
import org.junit.jupiter.api.Test;

import java.io.File;
import java.util.Arrays;
import java.util.List;
import java.util.Map;

import static org.junit.jupiter.api.Assertions.*;

/**
 * 跨三页的长表：{@code PdfParser#mergeSplitTables}「并完一截就把 {@code previous} 的页码与下边界换成后一截的」
 * 这条机制。仓库里的 PDF 样例（{@link TestSupport#PDF}）台账只跨两页，钉不住它 —— 两页的话 {@code previous}
 * 不换成后一截也照样能并上，只有三页以上才看得出「第二次比较用的是换过的那一截」。
 *
 * <p>三组样例都是测试现场生成的（{@code target/generated-test-resources/crosspage/}，不入库）：
 * <ul>
 *   <li><b>clean</b>：Aspose.Words 排一张 90 行两列的长表导出 PDF，占满三页。导出会盖上评估水印（每页一张图），
 *       再用 Aspose.PDF 把每页的图片删掉重存一次 —— 水印图排在表格片段之后，会让
 *       {@code Flow.lastOnPage=false}，跨页合并直接失效（见 {@link #theEvaluationWatermarkAloneIsEnoughToBreakTheMerge()}）。</li>
 *   <li><b>interrupted</b>：把 {@link TestSupport#PDF} 的第 1、2、2 页拼成一份 3 页 PDF，
 *       即「长表跨页时中间那一页后面还有别的内容（另一张表）」这种真实排版。</li>
 *   <li><b>marker-interrupted</b>：现造一份两页台账（第 2 页片段后面跟一张小图），第 3 页再放一张
 *       同列数、左对齐的独立小表。小图让片段的 {@code lastOnPage=false}；合并时必须把
 *       这个值也接到 {@code previous} 上，否则第 3 页的独立表会被错并进台账。</li>
 * </ul>
 */
class PdfCrossPageTableTest {

    private static final File DIR = new File("target/generated-test-resources/crosspage");

    /**
     * 数据行数，加表头共 90 行；行高写死 20pt，刚好排满三页。
     */
    private static final int DATA_ROWS = 89;

    private static Map<String, Object> cleanSilent;
    private static Map<String, Object> cleanRegistered;
    private static Map<String, Object> watermarkedSilent;
    private static Map<String, Object> interruptedSilent;
    private static Map<String, Object> markerInterruptedSilent;

    @BeforeAll
    static void generateAndParse() throws Exception {
        if (!DIR.isDirectory() && !DIR.mkdirs()) {
            throw new IllegalStateException("建不出目录：" + DIR.getAbsolutePath());
        }
        File watermarked = new File(DIR, "crosspage-watermarked.pdf");
        buildLongTableDoc().save(watermarked.getAbsolutePath(), com.aspose.words.SaveFormat.PDF);

        // 逐页删掉图片（评估水印）再存一份干净的：Aspose.PDF 评估模式存盘不加水印
        File clean = new File(DIR, "crosspage.pdf");
        com.aspose.pdf.Document scrub = new com.aspose.pdf.Document(watermarked.getAbsolutePath());
        for (int page = 1; page <= scrub.getPages().size(); page++) {
            scrub.getPages().get_Item(page).getResources().getImages().delete();
        }
        scrub.save(clean.getAbsolutePath());
        scrub.close();

        // 第 1、2、2 页拼三页：第 2 页的台账片段后面还跟着工单表，重复一次就是「中间那页有别的内容」
        File interrupted = new File(DIR, "interrupted.pdf");
        com.aspose.pdf.Document out = new com.aspose.pdf.Document();
        int[] pages = {1, 2, 2};
        for (int index = 0; index < pages.length; index++) {
            com.aspose.pdf.Document source = new com.aspose.pdf.Document(TestSupport.PDF);
            out.getPages().add(source.getPages().get_Item(pages[index]));
        }
        out.save(interrupted.getAbsolutePath());
        out.close();

        // 两页台账，第 2 页片段后盖一张小图，再拼一张第 3 页的独立小表
        File ledgerPdf = new File(DIR, "ledger-two-pages.pdf");
        buildTwoPageLedger().save(ledgerPdf.getAbsolutePath(), com.aspose.words.SaveFormat.PDF);
        File ledgerClean = new File(DIR, "ledger-two-pages-clean.pdf");
        scrubImages(ledgerPdf, ledgerClean);

        File independentPdf = new File(DIR, "independent.pdf");
        buildIndependentTable().save(independentPdf.getAbsolutePath(), com.aspose.words.SaveFormat.PDF);
        File independentClean = new File(DIR, "independent-clean.pdf");
        scrubImages(independentPdf, independentClean);

        File markerInterrupted = new File(DIR, "marker-interrupted.pdf");
        com.aspose.pdf.Document ledger = new com.aspose.pdf.Document(ledgerClean.getAbsolutePath());
        com.aspose.pdf.Document independent = new com.aspose.pdf.Document(independentClean.getAbsolutePath());
        com.aspose.pdf.Document composed = new com.aspose.pdf.Document();
        for (int page = 1; page <= ledger.getPages().size(); page++) {
            composed.getPages().add(ledger.getPages().get_Item(page));
        }
        composed.getPages().add(independent.getPages().get_Item(1));
        // 第 2 页表格片段后面再放一张小图：片段因此不是该页最后一块内容
        File marker = new File(DIR, "marker.png");
        java.awt.image.BufferedImage pixels = new java.awt.image.BufferedImage(8, 8,
                java.awt.image.BufferedImage.TYPE_INT_RGB);
        javax.imageio.ImageIO.write(pixels, "png", marker);
        com.aspose.pdf.ImageStamp stamp = new com.aspose.pdf.ImageStamp(marker.getAbsolutePath());
        stamp.setXIndent(10);
        stamp.setYIndent(10);
        stamp.setWidth(8);
        stamp.setHeight(8);
        composed.getPages().get_Item(2).addStamp(stamp);
        composed.save(markerInterrupted.getAbsolutePath());
        composed.close();
        ledger.close();
        independent.close();

        cleanSilent = PdfParser.parse(clean.getAbsolutePath(), WordParser.options());
        cleanRegistered = PdfParser.parse(clean.getAbsolutePath(),
                WordParser.options().headers("客户名称", "金额"));
        watermarkedSilent = PdfParser.parse(watermarked.getAbsolutePath(), WordParser.options());
        interruptedSilent = PdfParser.parse(interrupted.getAbsolutePath(), WordParser.options());
        markerInterruptedSilent = PdfParser.parse(markerInterrupted.getAbsolutePath(), WordParser.options());
    }

    /**
     * 90 行两列表：表头「客户名称｜金额」，数据 {@code C1001/1.00 … C1089/89.00}，六条边框 0.5pt 黑线。
     */
    private static com.aspose.words.Document buildLongTableDoc() throws Exception {
        com.aspose.words.Document doc = new com.aspose.words.Document();
        com.aspose.words.DocumentBuilder builder = new com.aspose.words.DocumentBuilder(doc);
        builder.getFont().setSize(9);
        builder.getRowFormat().setHeightRule(com.aspose.words.HeightRule.EXACTLY);
        builder.getRowFormat().setHeight(20.0);

        com.aspose.words.Table table = builder.startTable();
        builder.insertCell();
        builder.write("客户名称");
        builder.insertCell();
        builder.write("金额");
        builder.endRow();
        for (int r = 1; r <= DATA_ROWS; r++) {
            builder.insertCell();
            builder.write("C" + (1000 + r));
            builder.insertCell();
            builder.write(r + ".00");
            builder.endRow();
        }
        setBorders(table);
        builder.endTable();

        doc.save(new File(DIR, "crosspage.docx").getAbsolutePath());
        return doc;
    }

    /**
     * 两页台账：第 1 页 30 行（表头 + 29 行数据），第 2 页 29 行数据；小图由测试在拼装时盖在第 2 页。
     * 行高 20pt、页高 792pt、上下边距 72pt：第 1 页表格下边界离页底约 120pt（0.15 页高），
     * 第 2 页片段离页底约 140pt（0.18 页高），都落在跨页合并的 0.2 页高判据内。
     */
    private static com.aspose.words.Document buildTwoPageLedger() throws Exception {
        com.aspose.words.Document doc = new com.aspose.words.Document();
        com.aspose.words.DocumentBuilder builder = new com.aspose.words.DocumentBuilder(doc);
        builder.getFont().setSize(9);
        builder.getRowFormat().setHeightRule(com.aspose.words.HeightRule.EXACTLY);
        builder.getRowFormat().setHeight(20.0);

        writeTable(builder, 1, 29, true);
        builder.insertBreak(com.aspose.words.BreakType.PAGE_BREAK);
        writeTable(builder, 30, 58, false);

        doc.save(new File(DIR, "ledger-two-pages.docx").getAbsolutePath());
        return doc;
    }

    /**
     * 第 3 页的独立小表：两列 5 行、没有表头，文本是 {@code B1/1.00 … B5/5.00}。
     */
    private static com.aspose.words.Document buildIndependentTable() throws Exception {
        com.aspose.words.Document doc = new com.aspose.words.Document();
        com.aspose.words.DocumentBuilder builder = new com.aspose.words.DocumentBuilder(doc);
        builder.getFont().setSize(9);
        builder.getRowFormat().setHeightRule(com.aspose.words.HeightRule.EXACTLY);
        builder.getRowFormat().setHeight(20.0);

        com.aspose.words.Table table = builder.startTable();
        for (int r = 1; r <= 5; r++) {
            builder.insertCell();
            builder.write("B" + r);
            builder.insertCell();
            builder.write(r + ".00");
            builder.endRow();
        }
        setBorders(table);
        builder.endTable();

        doc.save(new File(DIR, "independent.docx").getAbsolutePath());
        return doc;
    }

    /**
     * 一张两列表（可选表头），数据 {@code C1001/1.00 …}。
     */
    private static void writeTable(com.aspose.words.DocumentBuilder builder, int from, int to, boolean header)
            throws Exception {
        com.aspose.words.Table table = builder.startTable();
        if (header) {
            builder.insertCell();
            builder.write("客户名称");
            builder.insertCell();
            builder.write("金额");
            builder.endRow();
        }
        for (int r = from; r <= to; r++) {
            builder.insertCell();
            builder.write("C" + (1000 + r));
            builder.insertCell();
            builder.write(r + ".00");
            builder.endRow();
        }
        setBorders(table);
        builder.endTable();
    }

    private static void setBorders(com.aspose.words.Table table) throws Exception {
        int[] types = {com.aspose.words.BorderType.TOP, com.aspose.words.BorderType.BOTTOM,
                com.aspose.words.BorderType.LEFT, com.aspose.words.BorderType.RIGHT,
                com.aspose.words.BorderType.HORIZONTAL, com.aspose.words.BorderType.VERTICAL};
        for (int type : types) {
            table.setBorder(type, com.aspose.words.LineStyle.SINGLE, 0.5, java.awt.Color.BLACK, true);
        }
    }

    /**
     * 删掉每页的图片（生成 PDF 时盖上的评估水印），另存一份；水印图排在表格片段之后会打断跨页合并。
     */
    private static void scrubImages(File source, File target) {
        com.aspose.pdf.Document document = new com.aspose.pdf.Document(source.getAbsolutePath());
        try {
            for (int page = 1; page <= document.getPages().size(); page++) {
                document.getPages().get_Item(page).getResources().getImages().delete();
            }
            document.save(target.getAbsolutePath());
        } finally {
            document.close();
        }
    }

    /**
     * 三页的长表并成一组：89 行数据、{@code source={page:1, lastPage:3}}，注册表头与静默模式读出同一份 JSON。
     */
    @Test
    void aThreePageTableIsMergedIntoOneGroupInBothModes() throws Exception {
        List<Map<String, Object>> both = Arrays.asList(cleanSilent, cleanRegistered);
        for (Map<String, Object> result : both) {
            assertEquals(1, TestSupport.tables(result).size(), "三页只该有一组：" + TestSupport.tables(result));
            assertEquals(DATA_ROWS, TestSupport.totalRows(result), "数据行数（表头那一行不算）");

            Map<String, Object> ledger = TestSupport.table(result, 0);
            assertEquals(Arrays.asList("客户名称", "金额"), TestSupport.headers(ledger), "列名");
            assertEquals("{page=1, lastPage=3}", String.valueOf(TestSupport.source(ledger)),
                    "并完之后 source 记的是首截的起始页与末截的结束页");
            assertEquals(DATA_ROWS, TestSupport.rows(ledger).size(), "行数");
            assertEquals(Arrays.asList("C1001", "1.00"),
                    TestSupport.values(TestSupport.rows(ledger).get(0)), "第一行在第 1 页");
            assertEquals(Arrays.asList("C1089", "89.00"),
                    TestSupport.values(TestSupport.rows(ledger).get(DATA_ROWS - 1)), "最后一行在第 3 页");
            TestSupport.assertHeadersMatchRowKeys(result);
        }
        assertEquals(PdfParser.toJson(cleanSilent), PdfParser.toJson(cleanRegistered),
                "注册表头不改变这份样例的分组结果（首行本来就是列名）");
    }

    /**
     * 中间那一页在表格片段后面还有别的内容时，链条就断在这儿：第 1+2 页并成一张，第 3 页另起一张，
     * 而且静默模式下第 3 页那截的首行（{@code 湖州建材贸易}）被当成了列名。
     *
     * <p>原因是 {@code mergeSplitTables} 只认「前一截是那一页的最后一块内容」：第 2 页片段后面跟着工单表，
     * 工单表那截就成了新的 {@code previous}，第 3 页的台账片段（5 列）再去比列数不同的工单表（3 列），并不上。
     */
    @Test
    void theChainStopsAtAPageThatHasSomethingElseAfterTheFragment() {
        List<Map<String, Object>> tables = TestSupport.tables(interruptedSilent);
        assertEquals(5, tables.size(), "组内容：" + tables);

        Map<String, Object> merged = tables.get(1);
        assertEquals("{page=1, lastPage=2}", String.valueOf(TestSupport.source(merged)), "前两页并成一张");
        assertEquals(31, TestSupport.rows(merged).size(), "台账 30 家客户一行不少");

        Map<String, Object> rest = tables.get(3);
        assertEquals("{page=3}", String.valueOf(TestSupport.source(rest)), "第 3 页那截没并上，也没有 lastPage");
        assertEquals(Arrays.asList("湖州建材贸易", "26.10", "26.10-3", "100%", "已结清"),
                TestSupport.headers(rest), "首行被当成列名，重名列按补名规则加了列号");
        assertEquals(10, TestSupport.rows(rest).size(), "第 2 页台账片段剩下 10 家客户");
    }

    /**
     * 中间页的片段后面跟的不是另一张表（这里是一张小图）时，链条同样要断：第 1、2 页并成一张，
     * 第 3 页那张同列数、左对齐的独立小表另成一组。
     *
     * <p>回归点：{@code mergeSplitTables} 把第 2 页片段并进 {@code previous} 后，必须同时把
     * {@code lastOnPage} 也从 {@code true} 换成片段的 {@code false}。漏换的话，第 3 页的独立表
     * 仍然看到「第一截是页底最后一块」，会被错并进台账（末行变成 {@code B5/5.00}）。
     */
    @Test
    void theChainAlsoStopsWhenANonTableFlowFollowsTheFragment() {
        List<Map<String, Object>> tables = TestSupport.tables(markerInterruptedSilent);
        assertEquals(2, tables.size(), "第 3 页的独立表不该并进台账：" + tables);

        Map<String, Object> ledger = tables.get(0);
        assertEquals("{page=1, lastPage=2}", String.valueOf(TestSupport.source(ledger)),
                "第 1、2 页的台账并成一张");
        assertEquals(58, TestSupport.rows(ledger).size(), "台账 58 行数据一行不少");
        assertEquals(Arrays.asList("C1058", "58.00"),
                TestSupport.values(TestSupport.rows(ledger).get(57)), "最后一行还是台账的");

        Map<String, Object> independent = tables.get(1);
        assertEquals("{page=3}", String.valueOf(TestSupport.source(independent)),
                "第 3 页的独立表没并上，也没有 lastPage");
        assertEquals(Arrays.asList("B1", "1.00"), TestSupport.headers(independent),
                "静默模式下首行被当成列名");
        assertEquals(4, TestSupport.rows(independent).size(), "B2–B5 四行数据");
        assertEquals(Arrays.asList("B5", "5.00"),
                TestSupport.values(TestSupport.rows(independent).get(3)));

        assertEquals(1, TestSupport.images(markerInterruptedSilent).size(), "第 2 页那张小图照常进 images");
        TestSupport.assertHeadersMatchRowKeys(markerInterruptedSilent);
    }

    /**
     * 评估水印一张图就足以打断合并：同样的三页长表，带着 Aspose.Words 的水印图读出来是 3 组、少了 2 行
     * （两截各自的首行当了列名）。
     *
     * <p>水印图排在每页表格片段<b>之后</b>，于是那一段的 {@code lastOnPage} 成了 {@code false}，
     * {@code continues} 判不出「后一截是前一页的延续」。这是评估模式下现场生成 PDF 样例的硬限制，
     * 也是本类要先用 Aspose.PDF 把图删掉再解析的原因。
     */
    @Test
    void theEvaluationWatermarkAloneIsEnoughToBreakTheMerge() {
        assertEquals(3, TestSupport.tables(watermarkedSilent).size(),
                "组内容：" + TestSupport.tables(watermarkedSilent));
        assertEquals(DATA_ROWS - 2, TestSupport.totalRows(watermarkedSilent),
                "三截里两截的首行被当成列名，各少一行");
        assertFalse(TestSupport.source(TestSupport.table(watermarkedSilent, 0)).containsKey("lastPage"),
                "第一截没并上任何东西：" + TestSupport.source(TestSupport.table(watermarkedSilent, 0)));
        assertEquals(Arrays.asList("C1031", "31.00"), TestSupport.headers(TestSupport.table(watermarkedSilent, 1)),
                "第二截的首行当了列名");

        List<Map<String, Object>> images = TestSupport.images(watermarkedSilent);
        assertEquals(3, images.size(), "每页一张水印图");
        for (Map<String, Object> image : images) {
            TestSupport.assertImageMeta(image);
            assertEquals(920, image.get("width"), "水印宽");
            assertEquals(500, image.get("height"), "水印高");
            assertEquals(images.get(0).get("contentId"), image.get("contentId"), "三页是同一张图");
        }
        assertTrue(TestSupport.images(cleanSilent).isEmpty(), "删掉图片之后 clean 样例一张图都没有");
    }
}
