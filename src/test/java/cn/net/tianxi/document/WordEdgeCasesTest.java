package cn.net.tianxi.document;

import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

import java.io.File;
import java.util.List;
import java.util.Map;

import static java.util.Arrays.asList;
import static org.junit.jupiter.api.Assertions.*;

/**
 * Word 的边角结构测试：用 {@link TestSupport} 现场拼最小 docx（4 个部件的 OOXML 包），
 * 覆盖样例文件里不值得专门存一份的结构，以及两处「畸形文件不能把栈撑爆」的防御分支。
 *
 * <p>不用 Aspose 自己保存这些文件 —— 没有 lic 时它保存阶段就会把文本截断，期望值就写不死了。
 */
class WordEdgeCasesTest {

    /* ---------------- 内容控件 ---------------- */

    @Test
    void paragraphInsideCellLevelContentControlIsRead(@TempDir File directory) throws Exception {
        // w:sdt 套在 w:tc 里（Word 表单和模板常见）：Grid.readRow 只做浅层段落查找，
        // 实测 Aspose 的浅层查找会看穿控件，内容不丢
        File file = TestSupport.docx(directory, "sdt-in-cell.docx", TestSupport.tbl(2,
                TestSupport.tr(TestSupport.tc("列1"), TestSupport.tc("列2")),
                TestSupport.tr(TestSupport.tcInner(TestSupport.sdt(TestSupport.p("控件里的甲"))),
                        TestSupport.tc("乙"))));

        Map<String, Object> result = WordParser.parse(file.getPath(), WordParser.options());

        assertEquals(1, TestSupport.tables(result).size());
        assertEquals(asList("控件里的甲", "乙"), TestSupport.values(result, 0, 0));
    }

    @Test
    void unfilledContentControlContributesNothing(@TempDir File directory) throws Exception {
        // 实测：没填写的「纯文本」控件不会把占位文字（Click here to enter text.）交出来 ——
        // 那句提示存在 w:sdtPr/w:placeholderText 里，是 Word 显示层的，不属于文档内容，所以 text 里看不到它
        File file = TestSupport.docx(directory, "empty-sdt.docx", TestSupport.p("正文"), TestSupport.emptyTextSdt());

        String text = TestSupport.text(WordParser.parse(file.getPath(), WordParser.options()));

        assertEquals("正文", text, "空控件连空行都不留");
    }

    @Test
    void contentControlWrappingAParagraphContributesThatParagraph(@TempDir File directory) throws Exception {
        // 控件里真写了内容才读得到：它自己算一个段落，与控件外的正文各占一行
        File file = TestSupport.docx(directory, "sdt.docx", TestSupport.p("正文"),
                TestSupport.sdt(TestSupport.p("控件里的段落")));

        String text = TestSupport.text(WordParser.parse(file.getPath(), WordParser.options()));

        assertEquals("正文\n控件里的段落", text);
    }

    /* ---------------- 嵌套表 ---------------- */

    @Test
    void nestedTableBecomesItsOwnGroupAndLeavesTheHostCellEmpty(@TempDir File directory) throws Exception {
        // Cell.getText() 会把子表文字摊进母格（母格变成一长串），只取母格自己的直接段落 + 子表另成一组，
        // 两边的行列结构都保得住
        String inner = TestSupport.tbl(1,
                TestSupport.tr(TestSupport.tc("嵌套表内容")),
                TestSupport.tr(TestSupport.tc("子表数据")));
        File file = TestSupport.docx(directory, "nested.docx", TestSupport.tbl(3,
                TestSupport.tr(TestSupport.tc("外A1"), TestSupport.tc("外A2"), TestSupport.tc("外A3")),
                TestSupport.tr(TestSupport.tc("甲"),
                        TestSupport.tcInner(inner + TestSupport.emptyP()),
                        TestSupport.tc("丙"))));

        Map<String, Object> result = WordParser.parse(file.getPath(), WordParser.options());
        List<Map<String, Object>> tables = TestSupport.tables(result);

        assertEquals(2, tables.size(), "母表一组、子表一组");
        assertEquals(asList("外A1", "外A2", "外A3"), TestSupport.headers(tables.get(0)));
        assertEquals(asList("甲", "", "丙"), TestSupport.values(result, 0, 0), "母表那一格不摊平子表文字");
        assertEquals(asList("嵌套表内容"), TestSupport.headers(tables.get(1)));
        assertEquals(asList("子表数据"), TestSupport.values(result, 1, 0));
    }

    @Test
    void tablesNestedBeyondTheDepthLimitAreSkippedWithAWarning(@TempDir File directory) throws Exception {
        // 畸形文件里表格套表格套不完：超过上限就跳过并提示，不能让栈爆掉
        File file = TestSupport.docx(directory, "deep-table.docx", nestedTables(12));

        final Map<String, Object>[] result = holder();
        String stderr = TestSupport.captureStderr(() ->
                result[0] = WordParser.parse(file.getPath(), WordParser.options()));

        assertTrue(stderr.contains("表格嵌套超过"), stderr);
        assertFalse(TestSupport.tables(result[0]).isEmpty(), "外层表格照常解析，只是更深的层次被跳过");
    }

    @Test
    void exactlyEightNestedTablesAreAllReadWithoutAWarning(@TempDir File directory) throws Exception {
        // 上限是 8 层：第 8 层没有再套子表，是正常结束，不该报「已跳过」
        File file = TestSupport.docx(directory, "eight-tables.docx", nestedTables(8));

        final Map<String, Object>[] result = holder();
        String stderr = TestSupport.captureStderr(() ->
                result[0] = WordParser.parse(file.getPath(), WordParser.options()));

        assertEquals("", stderr, "8 层正好在上限内，不该告警");
        assertEquals(8, TestSupport.tables(result[0]).size(), "8 层各成一组");
    }

    @Test
    void ninthNestedTableIsSkippedWithExactlyOneWarning(@TempDir File directory) throws Exception {
        // 第 9 层才越界：前 8 层照常成组，第 9 层整张跳过，提示只打一次
        File file = TestSupport.docx(directory, "nine-tables.docx", nestedTables(9));

        final Map<String, Object>[] result = holder();
        String stderr = TestSupport.captureStderr(() ->
                result[0] = WordParser.parse(file.getPath(), WordParser.options()));

        assertTrue(stderr.contains("表格嵌套超过 8 层"), stderr);
        assertEquals(1, occurrences(stderr, "表格嵌套超过"), "同一层多个子表也只提示一次：" + stderr);
        assertEquals(8, TestSupport.tables(result[0]).size(), "前 8 层各成一组");
    }

    @Test
    void containersNestedBeyondTheDepthLimitAreSkippedWithAWarning(@TempDir File directory) throws Exception {
        // 内容控件套内容控件同理：Body 之下的容器最多钻 16 层
        String inner = TestSupport.p("深处的段落");
        for (int i = 0; i < 20; i++) {
            inner = TestSupport.sdt(inner);
        }
        File file = TestSupport.docx(directory, "deep-sdt.docx", inner);

        final Map<String, Object>[] result = holder();
        String stderr = TestSupport.captureStderr(() ->
                result[0] = WordParser.parse(file.getPath(), WordParser.options()));

        assertTrue(stderr.contains("节点嵌套超过"), stderr);
        assertFalse(TestSupport.text(result[0]).contains("深处的段落"), "超过上限的那一支整个跳过");
    }

    /* ---------------- 合并单元格 ---------------- */

    @Test
    void verticallyMergedDataRowIsNotAMistakenSeparator(@TempDir File directory) throws Exception {
        // 回归：isSeparator 曾经把纵向延续格也算作「被合并掉」，于是这一行只剩一个起始格、被当成分组小标题，
        // KEEP 下第二行的值被抹成空串、DROP 下整行被删。判据改成只按横向延续格判定
        File file = TestSupport.docx(directory, "vmerge.docx", TestSupport.tbl(2,
                TestSupport.tr(TestSupport.tc("列A"), TestSupport.tc("列B")),
                TestSupport.tr(TestSupport.tcVerticalStart("第一"), TestSupport.tc("值一")),
                TestSupport.tr(TestSupport.tcVerticalContinue(), TestSupport.tc("值二"))));

        Map<String, Object> keep = WordParser.parse(file.getPath(), WordParser.options());
        Map<String, Object> drop = WordParser.parse(file.getPath(),
                WordParser.options().fullRow(WordParser.FullRow.DROP));

        assertEquals(2, TestSupport.rows(TestSupport.table(keep, 0)).size(), "两行数据都还在");
        assertEquals(asList("第一", "值一"), TestSupport.values(keep, 0, 0));
        assertEquals(asList("第一", "值二"), TestSupport.values(keep, 0, 1), "延续格回填起始格的值");
        assertEquals(2, TestSupport.rows(TestSupport.table(drop, 0)).size(), "DROP 也不该删掉数据行");
    }

    @Test
    void trueFullRowMergeIsStillASeparator(@TempDir File directory) throws Exception {
        // 真·跨整行合并的行判定不变：DROP 仍比 KEEP 少一行
        // （分组行放最后：它后面紧跟的行按表头处理，会把一行数据吃掉，那一条规则另有测试覆盖）
        File file = TestSupport.docx(directory, "full-row.docx", TestSupport.tbl(2,
                TestSupport.tr(TestSupport.tc("列A"), TestSupport.tc("列B")),
                TestSupport.tr(TestSupport.tc("甲"), TestSupport.tc("乙")),
                TestSupport.tr(TestSupport.tcSpan("2025 年度", 2))));

        Map<String, Object> keep = WordParser.parse(file.getPath(), WordParser.options());
        Map<String, Object> drop = WordParser.parse(file.getPath(),
                WordParser.options().fullRow(WordParser.FullRow.DROP));

        assertEquals(2, TestSupport.rows(TestSupport.table(keep, 0)).size());
        assertEquals(asList("2025 年度", ""), TestSupport.values(keep, 0, 1), "分组行只保留一个值");
        assertEquals(1, TestSupport.rows(TestSupport.table(drop, 0)).size());
        assertEquals(asList("甲", "乙"), TestSupport.values(drop, 0, 0));
    }

    @Test
    void blankFullWidthMergedRowIsNotASeparator(@TempDir File directory) throws Exception {
        // 排版用的空白跨整行合并行不是分组小标题：当成分组行的话，紧跟它的那一行会按「分组行之后的行是表头」
        // 被吃成表头 —— 一行数据没了、表还被切成两组。pdf 一直是这个判据（toGrid 要求整行不全空），
        // 这条把 docx 对齐；空白行本身留成一行空值，不凭空丢内容
        File file = TestSupport.docx(directory, "blank-full-row.docx", TestSupport.tbl(2,
                TestSupport.tr(TestSupport.tc("列A"), TestSupport.tc("列B")),
                TestSupport.tr(TestSupport.tc("甲"), TestSupport.tc("乙")),
                TestSupport.tr(TestSupport.tcSpan("", 2)),
                TestSupport.tr(TestSupport.tc("丙"), TestSupport.tc("丁"))));

        Map<String, Object> result = WordParser.parse(file.getPath(), WordParser.options());

        assertEquals(1, TestSupport.tables(result).size(), "空白行不该把表切成两组");
        assertEquals(asList("列A", "列B"), TestSupport.headers(TestSupport.table(result, 0)), "丙丁不该被当成表头");
        assertEquals(3, TestSupport.rows(TestSupport.table(result, 0)).size());
        assertEquals(asList("", ""), TestSupport.values(result, 0, 1), "空白行留成一行空值");
        assertEquals(asList("丙", "丁"), TestSupport.values(result, 0, 2));
    }

    /* ---------------- 域 ---------------- */

    @Test
    void cellContainingOnlyAnUncalculatedFieldGetsANumberedColumnName(@TempDir File directory) throws Exception {
        // 域代码一律丢掉，只留缓存的域结果；这一格只有域且结果是空的 → 清洗后是空表头 → 补成序号名，
        // 绝不能让「PAGE」这种排版指令变成列名
        File file = TestSupport.docx(directory, "field-cell.docx", TestSupport.tbl(2,
                TestSupport.tr(TestSupport.tcInner(TestSupport.pField("", "PAGE", "", "")),
                        TestSupport.tc("普通格")),
                TestSupport.tr(TestSupport.tc("1"), TestSupport.tc("2"))));

        Map<String, Object> result = WordParser.parse(file.getPath(), WordParser.options());

        assertEquals(asList("列1", "普通格"), TestSupport.headers(TestSupport.table(result, 0)));
        assertEquals(asList("1", "2"), TestSupport.values(result, 0, 0));
    }

    @Test
    void calculatedFieldResultSurvivesInsideACell(@TempDir File directory) throws Exception {
        File file = TestSupport.docx(directory, "field-result.docx", TestSupport.tbl(2,
                TestSupport.tr(TestSupport.tc("页次"), TestSupport.tc("说明")),
                TestSupport.tr(TestSupport.tcInner(TestSupport.pField("", "PAGE", "3", "")),
                        TestSupport.tc("当前页"))));

        Map<String, Object> result = WordParser.parse(file.getPath(), WordParser.options());

        assertEquals(asList("3", "当前页"), TestSupport.values(result, 0, 0));
    }

    /* ---------------- 软换行 ---------------- */

    @Test
    void softBreakInsideACellBecomesASpace(@TempDir File directory) throws Exception {
        // 单元格里的 Shift+Enter 同样换成空格：丢掉会把两行粘成一句，读起来分不出原来的断行
        File file = TestSupport.docx(directory, "soft-break.docx", TestSupport.tbl(1,
                TestSupport.tr(TestSupport.tc("说明")),
                TestSupport.tr(TestSupport.tcInner(TestSupport.pSoftBreak("第一行", "第二行")))));

        Map<String, Object> result = WordParser.parse(file.getPath(), WordParser.options());

        assertEquals(asList("第一行 第二行"), TestSupport.values(result, 0, 0));
    }

    /* ---------------- 小工具 ---------------- */

    /**
     * {@code levels} 层表格，每层两列一行，第一格先写自己的标签再套下一层。
     * 每层都带自己的文字，{@code tables} 的组数才等于层数。
     */
    private static String nestedTables(int levels) {
        String inner = null;
        for (int level = levels; level >= 1; level--) {
            String firstCell = inner == null
                    ? TestSupport.tc("第" + level + "层")
                    : TestSupport.tcInner(TestSupport.p("第" + level + "层") + inner + TestSupport.emptyP());
            inner = TestSupport.tbl(2, TestSupport.tr(firstCell, TestSupport.tc("值" + level)));
        }
        return inner;
    }

    private static int occurrences(String text, String needle) {
        int count = 0;
        for (int at = text.indexOf(needle); at >= 0; at = text.indexOf(needle, at + needle.length())) {
            count++;
        }
        return count;
    }

    /** 泛型数组不能直接 new，先 new 出可物化的 {@code Map<?,?>[]} 再 unchecked 转；这样 {@code -Xlint:all} 下零告警。 */
    @SuppressWarnings("unchecked")
    private static Map<String, Object>[] holder() {
        return (Map<String, Object>[]) new Map<?, ?>[1];
    }
}
