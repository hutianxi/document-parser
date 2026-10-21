package cn.net.tianxi.document;

import org.junit.jupiter.api.Test;

import java.util.Arrays;
import java.util.List;
import java.util.Map;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertNotEquals;
import static org.junit.jupiter.api.Assertions.assertTrue;

/**
 * 三种格式的横向一致性测试：同一份「2025 年度华东区销售业绩报告」存成 docx / pdf / pptx，
 * 解析出来的核心表必须逐格相同，其余差异要能一句话说清是哪一格式的特性造成的。
 *
 * <p>三份样例共用同一张核心表（季度销售明细）、同一套表头、同一张看板截图；数据量按各格式的可读边界裁剪过
 * （PPT 的单元格受评估模式 5 个字符限制、PDF 需要一张长到跨页的表），所以台账行数三份不同，见
 * {@link #ledgerRowCountDiffersBecauseTheSamplesWereSizedPerFormat}。
 */
class CrossFormatTest {

    /** 核心表的列名：三种格式、静默与注册两种模式都必须给出这一份。 */
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

        // 工单表：word 组2 / pdf 组2
        assertEquals(WordParser.toJson(TestSupport.table(word, 2)),
                PdfParser.toJson(TestSupport.table(pdf, 2)), "工单表");
        // 目标表前面那行分组小标题另成一组：word 组4 / pdf 组3
        assertEquals(WordParser.toJson(TestSupport.table(word, 4)),
                PdfParser.toJson(TestSupport.table(pdf, 3)), "分组行那组");
        // 目标表两组：word 组5、组6 / pdf 组4、组5
        assertEquals(WordParser.toJson(TestSupport.table(word, 5)),
                PdfParser.toJson(TestSupport.table(pdf, 4)), "2025 年度目标");
        assertEquals(WordParser.toJson(TestSupport.table(word, 6)),
                PdfParser.toJson(TestSupport.table(pdf, 5)), "2026 年度目标");
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
}
