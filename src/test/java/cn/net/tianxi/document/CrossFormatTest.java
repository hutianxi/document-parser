package cn.net.tianxi.document;

import org.junit.jupiter.api.Test;

import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;
import java.util.Map;

import static org.junit.jupiter.api.Assertions.*;

/**
 * 三种格式的横向一致性测试：同一份「2025 年度华东区销售业绩报告」存成 docx / pdf / pptx，
 * 解析出来的核心表必须逐格相同，其余差异要能一句话说清是哪一格式的特性造成的。
 *
 * <p>三份样例共用同一张核心表（季度销售明细）、同一套表头、同一张看板截图；数据量按各格式的可读边界裁剪过
 * （PPT 的单元格受评估模式 5 个字符限制、PDF 需要一张长到跨页的表），所以台账行数三份不同，见
 * {@link #ledgerRowCountDiffersBecauseTheSamplesWereSizedPerFormat}。
 */
class CrossFormatTest {

    /**
     * 核心表的列名：三种格式、静默与注册两种模式都必须给出这一份。
     */
    private static final List<String> CORE_HEADERS =
            Arrays.asList("片区", "区域", "销售额", "回款率", "销售额-5", "回款率-6", "备注");

    private static final List<List<String>> CORE_ROWS = Arrays.asList(
            Arrays.asList("华东大区", "上海", "1280", "92%", "1465", "95%", "超额完成"),
            Arrays.asList("华东大区", "江苏", "986", "88%", "1120", "90%", ""),
            Arrays.asList("华东南区", "浙江", "760", "85%", "890", "87%", "回款偏长"),
            Arrays.asList("华东南区", "安徽", "540", "80%", "616", "83%", ""));

    private static Map<String, Object> parse(String sample, WordParser.Options options) throws Exception {
        if (sample.equals(TestSupport.WORD)) {
            return WordParser.parse(sample, options);
        }
        if (sample.equals(TestSupport.PDF)) {
            return PdfParser.parse(sample, options);
        }
        return PptParser.parse(sample, options);
    }

    /* ---------------- 核心表逐格相同 ---------------- */

    @Test
    void coreTableIsIdenticalAcrossThreeFormats() throws Exception {
        for (String sample : TestSupport.ALL_SAMPLES) {
            Map<String, Object> result = parse(sample, WordParser.options());
            Map<String, Object> core = TestSupport.table(result, 0);

            assertEquals(CORE_HEADERS, TestSupport.headers(core), sample + " 的列名");
            List<Map<String, String>> rows = TestSupport.rows(core);
            assertEquals(CORE_ROWS.size(), rows.size(), sample + " 的行数");
            for (int r = 0; r < rows.size(); r++) {
                assertEquals(CORE_ROWS.get(r), TestSupport.values(rows.get(r)), sample + " 第 " + r + " 行");
            }
        }
    }

    @Test
    void coreTableIsIdenticalWhenHeadersAreRegistered() throws Exception {
        // 注册表头时列名用注册的原名，同名列仍按统一规则补后缀 —— 三种格式的键因此一模一样。
        // PDF 单元格换行、PPT 评估截断都不会让键跑偏，这正是注册表头的用途
        for (String sample : TestSupport.ALL_SAMPLES) {
            Map<String, Object> result = parse(sample, TestSupport.registered());
            assertEquals(CORE_HEADERS, TestSupport.headers(TestSupport.table(result, 0)), sample);
            assertEquals(CORE_ROWS.get(0), TestSupport.values(result, 0, 0), sample);
        }
    }

    /* ---------------- Word 与 PDF：同一份内容的两种载体 ---------------- */

    @Test
    void wordAndPdfAgreeOnEverySharedBlock() throws Exception {
        Map<String, Object> word = WordParser.parse(TestSupport.WORD, WordParser.options());
        Map<String, Object> pdf = PdfParser.parse(TestSupport.PDF, WordParser.options());

        // 比内容时把 source 摘掉：页码取决于各文件自己的分页（docx 台账 12 家、pdf 30 家，同一张目标表
        // 在 docx 落在第 2 页、pdf 落在第 3 页），页码本身由 {@link #sourceRecordsThePageEachBlockStartsOn} 钉住
        // 工单表：word 组2 / pdf 组2
        assertEquals(WordParser.toJson(TestSupport.withoutSource(TestSupport.table(word, 2))),
                WordParser.toJson(TestSupport.withoutSource(TestSupport.table(pdf, 2))), "工单表");
        // 目标表前面那行分组小标题另成一组：word 组4 / pdf 组3
        assertEquals(WordParser.toJson(TestSupport.withoutSource(TestSupport.table(word, 4))),
                WordParser.toJson(TestSupport.withoutSource(TestSupport.table(pdf, 3))), "分组行那组");
        // 目标表两组：word 组5、组6 / pdf 组4、组5
        assertEquals(WordParser.toJson(TestSupport.withoutSource(TestSupport.table(word, 5))),
                WordParser.toJson(TestSupport.withoutSource(TestSupport.table(pdf, 4))), "2025 年度目标");
        assertEquals(WordParser.toJson(TestSupport.withoutSource(TestSupport.table(word, 6))),
                WordParser.toJson(TestSupport.withoutSource(TestSupport.table(pdf, 5))), "2026 年度目标");
    }

    @Test
    void sourceRecordsThePageEachBlockStartsOn() throws Exception {
        // 三份样例的 source 值都是跑出来的实测值（docx 7 组、pdf 6 组、pptx 6 组，逐组打印过）
        Map<String, Object> word = WordParser.parse(TestSupport.WORD, WordParser.options());
        Map<String, Object> pdf = PdfParser.parse(TestSupport.PDF, WordParser.options());
        Map<String, Object> ppt = PptParser.parse(TestSupport.PPT, WordParser.options());

        // 跨页的台账：pdf 与 docx 都是从第 1 页跨到第 2 页，起始页与结束页都给
        assertEquals("{page=1, lastPage=2}", String.valueOf(TestSupport.source(TestSupport.table(pdf, 1))));
        assertEquals("{page=1, lastPage=2}", String.valueOf(TestSupport.source(TestSupport.table(word, 1))));
        // 不跨页只给 page
        assertEquals("{page=1}", String.valueOf(TestSupport.source(TestSupport.table(word, 0))));
        assertEquals("{page=3}", String.valueOf(TestSupport.source(TestSupport.table(pdf, 5))));
        // 同一张物理表切出的两组年度目标，来源是同一页 —— 组按表头切，不按页切
        assertEquals(TestSupport.source(TestSupport.table(pdf, 4)),
                TestSupport.source(TestSupport.table(pdf, 5)), "pdf 两组年度目标同页");
        // 单元格里的嵌套小表跟着母表所在的页（docx 第 2 页）
        assertEquals("{page=2}", String.valueOf(TestSupport.source(TestSupport.table(word, 3))));
        // pptx 用幻灯片序号，同样 1 起
        assertEquals("{page=1}", String.valueOf(TestSupport.source(TestSupport.table(ppt, 0))));
        assertEquals("{page=3}", String.valueOf(TestSupport.source(TestSupport.table(ppt, 5))));
    }

    @Test
    void sourceSwitchOffDropsTheKeyFromAllThreeFormats() throws Exception {
        // 开关对三种格式一起生效：关掉之后每组只剩 headers 与 rows 两个键，其余内容一字不变
        // （docx 因此连 LayoutCollector 都不建，同一 JVM 里开／关交替实测 11–35ms 对 74–254ms，见 README §1.1）
        for (String sample : TestSupport.ALL_SAMPLES) {
            Map<String, Object> result = parse(sample, WordParser.options().source(false));

            for (Map<String, Object> table : TestSupport.tables(result)) {
                assertNull(table.get("source"), sample + " 关了开关还留 source");
                assertEquals(Arrays.asList("headers", "rows"), new ArrayList<Object>(table.keySet()), sample + " 剩下的键");
            }
            assertEquals(CORE_HEADERS, TestSupport.headers(TestSupport.table(result, 0)), sample + " 的列名不受影响");
            assertEquals(CORE_ROWS.get(0), TestSupport.values(result, 0, 0), sample + " 的内容不受影响");
        }
    }

    @Test
    void ledgerRowCountDiffersBecauseTheSamplesWereSizedPerFormat() throws Exception {
        // 台账三份行数不同，是造样例时按各格式的边界裁的，不是解析丢了行：
        // docx 12 家客户 + 合计；pdf 30 家 + 合计（要长到跨页才能验证并回一张）；
        // ppt 6 家 + 合计（客户名要压在 5 个字符内，静默模式下还多一行下层表头）
        Map<String, Object> word = WordParser.parse(TestSupport.WORD, WordParser.options());
        Map<String, Object> pdf = PdfParser.parse(TestSupport.PDF, WordParser.options());
        Map<String, Object> ppt = PptParser.parse(TestSupport.PPT, WordParser.options());

        assertEquals(13, TestSupport.rows(TestSupport.table(word, 1)).size(), "docx");
        assertEquals(31, TestSupport.rows(TestSupport.table(pdf, 1)).size(), "pdf");
        assertEquals(8, TestSupport.rows(TestSupport.table(ppt, 1)).size(), "ppt 静默模式");
        // 三份的第一行客户是同一家（名字长短不同），合计行都在最后
        assertEquals("上海浦东城建集团", TestSupport.rows(TestSupport.table(word, 1)).get(0).get("客户名称"));
        assertEquals("上海浦东城建集团", TestSupport.rows(TestSupport.table(pdf, 1)).get(0).get("客户名称"));
        assertEquals("上海城建", TestSupport.rows(TestSupport.table(ppt, 1)).get(1).get("客户名称"));
    }

    @Test
    void onlyPdfNeedsTheCrossPageMerge() throws Exception {
        // 同一张台账：docx 里是一张完整的表，pdf 里被分页切成两截要靠几何判据并回来。
        // 判据失效的代价是「一张表拆成两组、各带一个表头」，所以这里断言组数而不是只看行数
        Map<String, Object> pdf = PdfParser.parse(TestSupport.PDF, WordParser.options());
        int ledgerGroups = 0;
        for (Map<String, Object> table : TestSupport.tables(pdf)) {
            if (TestSupport.headers(table).contains("客户名称")) {
                ledgerGroups++;
            }
        }
        assertEquals(1, ledgerGroups, "跨页的台账应该并回一组");
    }

    /* ---------------- 图片 ---------------- */

    @Test
    void imageBytesDependOnTheContainerNotOnTheParser() throws Exception {
        Map<String, Object> word = TestSupport.image(WordParser.parse(TestSupport.WORD, WordParser.options()), 0);
        Map<String, Object> ppt = TestSupport.image(PptParser.parse(TestSupport.PPT, WordParser.options()), 0);
        Map<String, Object> pdf = TestSupport.image(PdfParser.parse(TestSupport.PDF, WordParser.options()), 0);

        // docx 与 pptx 里存的是同一张 PNG，字节相同 → 指纹相同
        assertEquals(word.get("contentId"), ppt.get("contentId"), "docx 与 pptx 的图片字节相同");
        assertEquals(word.get("size"), ppt.get("size"));
        // Word 导出 PDF 时把它重编成 JPEG，字节变了 → 指纹、扩展名、尺寸都跟着变，解析照实报告
        assertNotEquals(word.get("contentId"), pdf.get("contentId"), "pdf 里是重编过的 JPEG");
        assertEquals("PNG", word.get("type"));
        assertEquals("JPEG", pdf.get("type"));
        assertEquals("image-1.png", word.get("fileName"));
        assertEquals("image-1.jpg", pdf.get("fileName"), "扩展名跟着实际类型走");
    }

    /* ---------------- 文档里的实测基线 ---------------- */

    /**
     * README §11「实测基线」、docs/index.html 与 CHANGELOG 里写的 JSON 字符数，逐个钉在这里。
     *
     * <p>为什么要单独一条断言：那些数字在文档里被当成「体积要心里有数」的依据，改了样例或改了序列化
     * 就会漂移，而漂移本身不违反任何契约 —— 没有这条断言的话，文档里的数字会慢慢变成没人核对过的传说。
     */
    @Test
    void documentedJsonLengthsAreTheRealOnes() throws Exception {
        assertEquals(4740, length(TestSupport.WORD, WordParser.options()), "docx 静默");
        assertEquals(6473, length(TestSupport.PDF, WordParser.options()), "pdf 静默");
        assertEquals(3526, length(TestSupport.PPT, WordParser.options()), "pptx 静默");

        // 注册五组但不起名：docx 与 pdf 的 JSON 与静默模式逐字节相同（那条断言在 TableByNameTest），
        // pptx 因为两级表头被正确合并、少了一行数据而变短
        assertEquals(4740, length(TestSupport.WORD, TestSupport.registered()), "docx 注册不起名");
        assertEquals(6473, length(TestSupport.PDF, TestSupport.registered()), "pdf 注册不起名");
        assertEquals(3378, length(TestSupport.PPT, TestSupport.registered()), "pptx 注册不起名");

        // 起名：命中的组多一个 name 键，其余一字不变
        assertEquals(4828, length(TestSupport.WORD, TestSupport.named()), "docx 注册并起名");
        assertEquals(6561, length(TestSupport.PDF, TestSupport.named()), "pdf 注册并起名");
        assertEquals(3466, length(TestSupport.PPT, TestSupport.named()), "pptx 注册并起名");

        // BASE64 模式：那张 670939 字节的 PNG 编码后 894610 个字符，docx 的体积涨了 190 倍
        assertEquals(899370, length(TestSupport.WORD, WordParser.options().images(WordParser.ImageMode.BASE64)),
                "docx BASE64");
        assertEquals(58076, length(TestSupport.PDF, WordParser.options().images(WordParser.ImageMode.BASE64)),
                "pdf BASE64（里面是 38668 字节的 JPEG）");
        assertEquals(898156, length(TestSupport.PPT, WordParser.options().images(WordParser.ImageMode.BASE64)),
                "pptx BASE64（与 docx 同一张图，差在没有跨页台账那些行）");
    }

    private static int length(String sample, WordParser.Options options) throws Exception {
        return WordParser.toJson(parse(sample, options)).length();
    }

    /* ---------------- 选项与契约 ---------------- */

    @Test
    void oneOptionsInstanceServesAllThreeParsers() throws Exception {
        // Options 是三种格式共用的（PdfParser/PptParser 接收的就是 WordParser.Options），
        // 且配置好之后可以重复用于多次解析：解析过程不改它，图片编号每次都从 1 开始
        WordParser.Options options = TestSupport.registered().fullRow(WordParser.FullRow.KEEP);

        Map<String, Object> word = WordParser.parse(TestSupport.WORD, options);
        Map<String, Object> pdf = PdfParser.parse(TestSupport.PDF, options);
        Map<String, Object> ppt = PptParser.parse(TestSupport.PPT, options);
        Map<String, Object> wordAgain = WordParser.parse(TestSupport.WORD, options);

        assertEquals(WordParser.toJson(word), WordParser.toJson(wordAgain), "同一个 Options 重复解析结果要一致");
        assertEquals("image-1.png", TestSupport.image(word, 0).get("fileName"));
        assertEquals("image-1.png", TestSupport.image(ppt, 0).get("fileName"));
        assertEquals("image-1.jpg", TestSupport.image(pdf, 0).get("fileName"));
        assertEquals(CORE_HEADERS, TestSupport.headers(TestSupport.table(word, 0)));
        assertEquals(CORE_HEADERS, TestSupport.headers(TestSupport.table(pdf, 0)));
        assertEquals(CORE_HEADERS, TestSupport.headers(TestSupport.table(ppt, 0)));
    }

    @Test
    void contractHoldsForEverySampleAndMode() throws Exception {
        // 结果契约（headers 逐个等于每行的键、图片元信息齐全）在三份样例 × 四种选项组合下都成立
        WordParser.Options[] all = {
                WordParser.options(),
                WordParser.options().fullRow(WordParser.FullRow.FIRST),
                WordParser.options().fullRow(WordParser.FullRow.DROP),
                TestSupport.registered().fullRow(WordParser.FullRow.DROP),
        };
        for (String sample : TestSupport.ALL_SAMPLES) {
            for (WordParser.Options options : all) {
                Map<String, Object> result = parse(sample, options);
                TestSupport.assertHeadersMatchRowKeys(result);
                assertEquals(1, TestSupport.images(result).size(), sample + " 的图片数");
                TestSupport.assertImageMeta(TestSupport.image(result, 0));
                assertTrue(TestSupport.tables(result).size() >= 5, sample + " 的表头组数");
                assertTrue(TestSupport.totalRows(result) >= 19, sample + " 的数据行数");
            }
        }
    }

    @Test
    void dropNeverKeepsMoreThanKeep() throws Exception {
        // 丢掉分组行只会让组和行变少，不会变多；FIRST 只改归属、不改数量
        for (String sample : TestSupport.ALL_SAMPLES) {
            int keepRows = TestSupport.totalRows(parse(sample, WordParser.options()));
            int keepGroups = TestSupport.tables(parse(sample, WordParser.options())).size();
            int firstRows = TestSupport.totalRows(parse(sample,
                    WordParser.options().fullRow(WordParser.FullRow.FIRST)));
            int dropRows = TestSupport.totalRows(parse(sample,
                    WordParser.options().fullRow(WordParser.FullRow.DROP)));
            int dropGroups = TestSupport.tables(parse(sample,
                    WordParser.options().fullRow(WordParser.FullRow.DROP))).size();

            assertEquals(keepRows, firstRows, sample + "：FIRST 只改归属，行数不变");
            assertEquals(keepGroups, TestSupport.tables(parse(sample,
                    WordParser.options().fullRow(WordParser.FullRow.FIRST))).size(), sample + "：FIRST 组数不变");
            assertTrue(dropRows <= keepRows, sample + "：DROP 行数 " + dropRows + " > KEEP " + keepRows);
            assertTrue(dropGroups <= keepGroups, sample + "：DROP 组数变多了");
        }
    }

    @Test
    void unmatchedHeadersIsAnOptInTopLevelKey() throws Exception {
        // 顶层键默认只有三个：多一个未知键会让按 POJO 反序列化的下游直接抛 UnrecognizedPropertyException
        // （Jackson 默认 FAIL_ON_UNKNOWN_PROPERTIES=true），所以这份清单要显式打开。
        // 打开之后它排在最后，其余三个键的内容一字不变 —— 它就是标准错误那条提示的同一份数据，只是给了程序读的入口
        for (String sample : TestSupport.ALL_SAMPLES) {
            Map<String, Object> off = parse(sample, TestSupport.named());
            Map<String, Object> on = parse(sample, TestSupport.named().unmatchedHeaders(true));

            assertEquals(Arrays.asList("text", "tables", "images"), new ArrayList<Object>(off.keySet()),
                    sample + "：默认没有 unmatchedHeaders 这个顶层键");
            assertEquals(Arrays.asList("text", "tables", "images", "unmatchedHeaders"),
                    new ArrayList<Object>(on.keySet()), sample + "：打开后多一个顶层键，排在最后");
            assertEquals(off.get("text"), on.get("text"), sample + "：text 不受影响");
            assertEquals(off.get("tables"), on.get("tables"), sample + "：tables 不受影响");
            assertEquals(off.get("images"), on.get("images"), sample + "：images 不受影响");

            List<Map<String, Object>> unmatched = unmatched(on);
            if (sample.equals(TestSupport.WORD)) {
                assertTrue(unmatched.isEmpty(), "docx 上 named() 的五组注册全部命中，清单是空列表而不是 null");
            } else {
                // 那张嵌套小表只有 docx 有，pdf／pptx 里没有对应的那一行；它注册了但没起名
                assertEquals(1, unmatched.size(), sample + " 只有一组注册没命中");
                assertEquals(Arrays.asList("故障部件", "工单数"), unmatched.get(0).get("headers"));
                assertEquals(Arrays.asList("headers"), new ArrayList<Object>(unmatched.get(0).keySet()),
                        "没起名字的注册就只有 headers 一个键");
            }
        }
    }

    @Test
    void unmatchedHeadersCarriesTheNameOfARegistrationThatMatchedNothing() throws Exception {
        // 起了名字却没命中的注册照样进清单、并且带上 name —— 这正是「按名字取数取到空列表」时最想知道的一条。
        // pdf 的台账在网格里只给了合并后那一行，照上层那一行注册（默认容忍度 1）命不中（容忍度调到 2 才命中）
        WordParser.Options options = WordParser.options()
                .headers("客户名称", "回款情况", "回款情况", "回款率", "备注").name("回款台账")
                .unmatchedHeaders(true);

        Map<String, Object> pdf = PdfParser.parse(TestSupport.PDF, options);
        List<Map<String, Object>> unmatched = unmatched(pdf);

        assertEquals(1, unmatched.size());
        assertEquals(Arrays.asList("headers", "name"), new ArrayList<Object>(unmatched.get(0).keySet()),
                "键序固定：headers 在前、name 在后");
        assertEquals(Arrays.asList("客户名称", "回款情况", "回款情况", "回款率", "备注"),
                unmatched.get(0).get("headers"), "写的是注册时的原文，不是比对用的归一化形式");
        assertEquals("回款台账", unmatched.get(0).get("name"));
        assertTrue(WordParser.tablesByName(pdf, "回款台账").isEmpty(), "取数入口给的是空列表，不是 null");
    }

    @SuppressWarnings("unchecked")
    private static List<Map<String, Object>> unmatched(Map<String, Object> result) {
        return (List<Map<String, Object>>) result.get("unmatchedHeaders");
    }

    /* ---------------- 文件不存在 ---------------- */

    @Test
    void aMissingFileThrowsFromAllThreeParsers() throws Exception {
        // 三个解析器的 parse(String, Options) 都是把路径直接交给 Aspose 打开、没有自己包一层，
        // 所以「文件不存在」时抛什么由各自的库决定。实测三份的消息都是「Can't find file: <路径>.」，
        // 但异常类型各不相同 —— 想统一 catch java.io.FileNotFoundException 的话只有 docx 兜得住
        String wordPath = "src/test/resources/word/not-here.docx";
        String pdfPath = "src/test/resources/pdf/not-here.pdf";
        String pptPath = "src/test/resources/ppt/not-here.pptx";

        Exception word = assertThrows(Exception.class, () -> WordParser.parse(wordPath, WordParser.options()));
        Exception pdf = assertThrows(Exception.class, () -> PdfParser.parse(pdfPath, WordParser.options()));
        Exception ppt = assertThrows(Exception.class, () -> PptParser.parse(pptPath, WordParser.options()));

        assertTrue(word.getMessage().contains(wordPath), "docx：" + word.getMessage());
        assertTrue(pdf.getMessage().contains(pdfPath), "pdf：" + pdf.getMessage());
        assertTrue(ppt.getMessage().contains(pptPath), "pptx：" + ppt.getMessage());

        assertTrue(word instanceof java.io.FileNotFoundException,
                "docx 抛的就是 JDK 的 FileNotFoundException");
        assertFalse(pdf instanceof java.io.FileNotFoundException,
                "pdf 抛的是 Aspose.PDF 内部包装过的类型：" + pdf.getClass().getName());
        assertFalse(ppt instanceof java.io.FileNotFoundException,
                "pptx 抛的是 com.aspose.slides.exceptions.FileNotFoundException，与 JDK 那个同名不同类："
                        + ppt.getClass().getName());
    }
}
