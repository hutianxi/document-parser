package cn.net.tianxi.document;

import org.junit.jupiter.api.Test;

import java.util.*;

import static org.junit.jupiter.api.Assertions.*;

/**
 * {@link Tables} 的单元测试：表头分组是三种格式共用的唯一一份算法，这里用手工网格直接断言，不经过 Aspose。
 *
 * <p>{@code Tables.group} 的输入是解析器摊平后的三样东西：{@code String[][] rows}（合并覆盖的位置已回填起始格的值）、
 * {@code boolean[] separator}（该行是否跨整行合并）、{@code boolean[] heading}（格式自带的标题行标记，没有就传 null），
 * 加上 {@code firstRowStyled}（表样式是否声明首行为标题，Word 的 {@code tblLook firstRow}、PPT 的
 * {@code table.getFirstRow()}）。PDF 三者都给不出，所以它最依赖注册表头。
 */
class TablesTest {

    /**
     * 便捷入口：每次新建一个 matched 集合（调用方不需要看它时用这个）。
     */
    private static List<Map<String, Object>> group(String[][] rows, boolean[] separator, boolean[] heading,
                                                   WordParser.Options options, boolean firstRowStyled) {
        return Tables.group(rows, separator, heading, options, new LinkedHashSet<Integer>(), firstRowStyled);
    }

    private static List<Map<String, Object>> group(String[][] rows, WordParser.Options options,
                                                   boolean firstRowStyled) {
        return group(rows, new boolean[rows.length], null, options, firstRowStyled);
    }

    private static String[][] rows(String[]... rows) {
        return rows;
    }

    private static boolean[] flags(boolean... values) {
        return values;
    }

    /* ---------------- 列名补全 ---------------- */

    @Test
    void normalizeHeadersFillsBlankAndDuplicateNames() {
        // 空表头补序号名、重名补列号后缀，规则与行对象的键同源，headers 才能直接当列清单用
        assertEquals(Arrays.asList("合计", "合计-2", "名称"),
                Tables.normalizeHeaders(Arrays.asList("合计", "合计", "名称"), 3));
        assertEquals(Arrays.asList("列1", "金额", "列3"),
                Tables.normalizeHeaders(Arrays.asList("", "金额", ""), 3));
    }

    @Test
    void normalizeHeadersPadsShortHeaderListsToTheColumnCount() {
        // 列数少的表头补齐到整表最大列数：那些键本来就存在于每一行，这里只是给它们名字
        assertEquals(Arrays.asList("片区", "目标", "列3", "列4"),
                Tables.normalizeHeaders(Arrays.asList("片区", "目标"), 4));
    }

    @Test
    void rowMapKeysMatchTheHeaderNamesExactly() {
        Map<String, String> row = Tables.rowMap(Arrays.asList("片区", "目标"), Arrays.asList("华东", "2400"));

        assertEquals(Arrays.asList("片区", "目标"), new ArrayList<String>(row.keySet()));
        assertEquals("华东", row.get("片区"));
        assertEquals("2400", row.get("目标"));
    }

    @Test
    void rowMapNamesExtraColumnsWhenHeadersAreShorter() {
        // 表头只给到两列、数据行有三格时，第三格按同一条规则补成序号名，不会静默丢掉
        Map<String, String> row = Tables.rowMap(Arrays.asList("片区", "目标"), Arrays.asList("华东", "2400", "达标"));

        assertEquals(Arrays.asList("片区", "目标", "列3"), new ArrayList<String>(row.keySet()));
        assertEquals("达标", row.get("列3"));
    }

    @Test
    void mergeHeadersLetsTheLowerLevelWinAndInheritBlanks() {
        // 多级表头：下层写了名字的用下层的，下层空着的（纵向合并的列名）沿用上层
        assertEquals(Arrays.asList("客户名称", "合同金额", "已回款", "回款率"),
                Tables.mergeHeaders(Arrays.asList("客户名称", "回款情况", "回款情况", "回款率"),
                        Arrays.asList("", "合同金额", "已回款", "")));
    }

    @Test
    void paddedFillsMissingCellsAndNeverReturnsNull() {
        assertEquals(Arrays.asList("a", "b", ""), Tables.padded(new String[]{"a", "b"}, 3));
        assertEquals(Arrays.asList("a"), Tables.padded(new String[]{"a", "b"}, 1));
        assertEquals(Arrays.asList("", ""), Tables.padded(new String[]{null, null}, 2));
    }

    @Test
    void singleKeepsOnlyTheFirstValueOfAFullRowMerge() {
        assertEquals(Arrays.asList("2025 年度", "", ""), Tables.single("2025 年度", 3));
    }

    @Test
    void firstNonBlankSkipsLeadingEmptyCells() {
        // 跨整行合并的起点不一定在第一列，取第一个非空格才不会拿到空串
        assertEquals("2025 年度", Tables.firstNonBlank(Arrays.asList("", "2025 年度", "")));
        assertEquals("", Tables.firstNonBlank(Arrays.asList("", "  ")));
    }

    /* ---------------- 表头判定 ---------------- */

    @Test
    void firstRowIsAHeaderWhenTheTableStyleSaysSo() {
        List<Map<String, Object>> blocks = group(rows(
                new String[]{"片区", "目标"},
                new String[]{"华东", "2400"}), WordParser.options(), true);

        assertEquals(1, blocks.size());
        assertEquals(Arrays.asList("片区", "目标"), TestSupport.headers(blocks.get(0)));
        assertEquals(Arrays.asList("华东", "2400"), TestSupport.values(TestSupport.rows(blocks.get(0)).get(0)));
    }

    @Test
    void firstRowIsAHeaderInSilentModeEvenWithoutStyleHints() {
        // 调用方没给任何表头依据、格式也没有标题行标记（PDF 就是这种）时，首行兜底当表头
        List<Map<String, Object>> blocks = group(rows(
                new String[]{"片区", "目标"},
                new String[]{"华东", "2400"}), WordParser.options(), false);

        assertEquals(Arrays.asList("片区", "目标"), TestSupport.headers(blocks.get(0)));
        assertEquals(1, TestSupport.rows(blocks.get(0)).size());
    }

    @Test
    void firstRowStaysADataRowOnceHeaderEvidenceIsRegistered() {
        // 给了表头依据就以依据为准：没命中的表首行按数据行处理，列名补成序号 —— 免得把一行数据当表头
        List<Map<String, Object>> blocks = group(rows(
                new String[]{"片区", "目标"},
                new String[]{"华东", "2400"}), WordParser.options().headers("别的", "表头"), false);

        assertEquals(1, blocks.size());
        assertEquals(Arrays.asList("列1", "列2"), TestSupport.headers(blocks.get(0)));
        assertEquals(2, TestSupport.rows(blocks.get(0)).size(), "首行也是一行数据");
    }

    @Test
    void explicitHeadingMarksWinOverSilence() {
        // heading[r] 为 true 的行按表头处理，即使调用方注册了别的表头
        List<Map<String, Object>> blocks = group(rows(
                        new String[]{"片区", "目标"},
                        new String[]{"华东", "2400"}),
                flags(false, false), flags(true, false), WordParser.options().headers("别的", "表头"), false);

        assertEquals(Arrays.asList("片区", "目标"), TestSupport.headers(blocks.get(0)));
        assertEquals(1, TestSupport.rows(blocks.get(0)).size());
    }

    @Test
    void keywordsMakeEveryMatchingRowAHeader() {
        // 一张表里出现两次表头行（例如「小计」后又重复一遍列名），关键字兜底会把它们分成两组
        List<Map<String, Object>> blocks = group(rows(
                new String[]{"列a", "列b"},
                new String[]{"1", "2"},
                new String[]{"列a", "列b"},
                new String[]{"3", "4"}), WordParser.options().keywords("列"), false);

        assertEquals(2, blocks.size(), "两组表头");
        assertEquals(1, TestSupport.rows(blocks.get(0)).size());
        assertEquals(1, TestSupport.rows(blocks.get(1)).size());
        assertEquals(Arrays.asList("3", "4"), TestSupport.values(TestSupport.rows(blocks.get(1)).get(0)));
    }

    @Test
    void rowAfterASeparatorRowIsTreatedAsAHeader() {
        // 分组小标题后面紧跟着的通常是新一段的列名：r>0 且上一行是分组行 → 这一行按表头处理
        List<Map<String, Object>> blocks = group(rows(
                        new String[]{"2025 年度", "", ""},
                        new String[]{"片区", "目标", "完成率"},
                        new String[]{"华东", "2400", "101%"}),
                flags(true, false, false), null, WordParser.options(), false);

        assertEquals(2, blocks.size());
        assertEquals(Arrays.asList("列1", "列2", "列3"), TestSupport.headers(blocks.get(0)));
        assertEquals(Arrays.asList("片区", "目标", "完成率"), TestSupport.headers(blocks.get(1)));
        assertEquals(1, TestSupport.rows(blocks.get(1)).size());
    }

    @Test
    void consecutiveHeaderRowsMergeIntoOneMultiLevelHeader() {
        // 两行表头之间没有数据行 → 合并成一组，下层空列沿用上层的名字
        List<Map<String, Object>> blocks = group(rows(
                        new String[]{"客户名称", "回款情况", "回款情况", "回款率"},
                        new String[]{"", "合同金额", "已回款", ""},
                        new String[]{"上海城建", "486", "486", "100%"}),
                flags(false, false, false), flags(true, true, false), WordParser.options(), true);

        assertEquals(1, blocks.size());
        assertEquals(Arrays.asList("客户名称", "合同金额", "已回款", "回款率"), TestSupport.headers(blocks.get(0)));
        assertEquals(1, TestSupport.rows(blocks.get(0)).size(), "两行表头只算一组表头，不是数据");
    }

    @Test
    void headerRowAfterDataStartsANewGroup() {
        List<Map<String, Object>> blocks = group(rows(
                        new String[]{"片区", "目标"},
                        new String[]{"华东", "2400"},
                        new String[]{"客户", "金额"},
                        new String[]{"上海城建", "486"}),
                flags(false, false, false, false), flags(true, false, true, false), WordParser.options(), true);

        assertEquals(2, blocks.size());
        assertEquals(Arrays.asList("片区", "目标"), TestSupport.headers(blocks.get(0)));
        assertEquals(Arrays.asList("客户", "金额"), TestSupport.headers(blocks.get(1)));
    }

    @Test
    void dataRowsBeforeAnyHeaderGetNumberedColumns() {
        List<Map<String, Object>> blocks = group(rows(
                new String[]{"华东", "2400"}), WordParser.options().headers("片区", "目标"), false);

        assertEquals(Arrays.asList("列1", "列2"), TestSupport.headers(blocks.get(0)));
        assertEquals(1, TestSupport.rows(blocks.get(0)).size());
    }

    /* ---------------- 注册表头：命中、覆盖列名、回填 matched ---------------- */

    @Test
    void registeredHeadersOverrideTheNamesReadFromTheGrid() {
        // 命中之后列名用注册的原名：PDF 单元格换行会把列名断成「列 Bb」这种带空格的形式，
        // 比对忽略空格所以照样命中，用注册名三种格式的键才一致
        Set<Integer> matched = new LinkedHashSet<Integer>();
        List<Map<String, Object>> blocks = Tables.group(rows(
                        new String[]{"列 Bb", "列 Ewew"}),
                new boolean[1], null, WordParser.options().headers("列Bb", "列Ewew"), matched, false);

        assertEquals(Arrays.asList("列Bb", "列Ewew"), TestSupport.headers(blocks.get(0)));
        assertEquals(Arrays.asList(0), new ArrayList<Integer>(matched), "命中的是第 0 组");
    }

    @Test
    void aRegisteredNameIsWrittenOnlyOnTheGroupsThatMatched() {
        // name 跟着「命中第几组注册表头」走：同一张表里两组表头都命中同一组注册时，两组都带上同一个名字；
        // 首行兜底、关键字、分组行之后那一条规则判出来的组一个键都不多
        WordParser.Options options = WordParser.options().headers("片区", "年度目标", "完成率").name("年度目标");
        List<Map<String, Object>> blocks = group(rows(
                        new String[]{"片区", "年度目标", "完成率"},
                        new String[]{"华东", "2400", "92%"},
                        new String[]{"合计", "合计", "合计"},              // 跨整行合并的分组小标题
                        new String[]{"片区", "年度目标", "完成率"},
                        new String[]{"华南", "1800", "88%"}),
                flags(false, false, true, false, false), null, options, false);

        assertEquals(2, blocks.size());
        assertEquals("年度目标", blocks.get(0).get("name"), "第一组表头命中注册集");
        assertEquals("年度目标", blocks.get(1).get("name"), "分组行之后的第二组表头命中的是同一组注册集");
        // 键的先后固定：headers → rows → name（起过名字才有）
        assertEquals(Arrays.asList("headers", "rows", "name"),
                new ArrayList<Object>(blocks.get(0).keySet()));
    }

    @Test
    void unnamedRegistrationsAndFallbackHeadersGetNoNameKey() {
        // 注册了但没起名 = 静默模式之外的普通注册，结果与不注册时一样不带 name 键
        List<Map<String, Object>> named = group(rows(new String[]{"片区", "目标"}, new String[]{"华东", "2400"}),
                WordParser.options().headers("片区", "目标"), false);
        assertEquals(Arrays.asList("headers", "rows"), new ArrayList<Object>(named.get(0).keySet()));

        // 靠首行兜底判出来的表头没有注册集可对应，自然也不会有名字
        List<Map<String, Object>> fallback = group(rows(new String[]{"片区", "目标"}, new String[]{"华东", "2400"}),
                WordParser.options(), false);
        assertEquals(Arrays.asList("headers", "rows"), new ArrayList<Object>(fallback.get(0).keySet()));
    }

    @Test
    void consecutiveHeaderRowsMatchingTwoRegisteredSetsKeepTheFirstName() {
        // 多级表头的上下两行都注册过、且都起了名字时，以第一个命中的为准，
        // 结果里不会出现「一组两个 name」；代价是第二个名字在这份文档上落不了地
        WordParser.Options options = WordParser.options()
                .headers("客户名称", "回款情况", "回款情况", "回款率").name("上层")
                .headers("客户名称", "合同金额", "已回款", "回款率").name("下层");
        List<Map<String, Object>> blocks = group(rows(
                        new String[]{"客户名称", "回款情况", "回款情况", "回款率"},
                        new String[]{"客户名称", "合同金额", "已回款", "回款率"},
                        new String[]{"上海城建", "486", "486", "100%"}),
                options, false);

        assertEquals(1, blocks.size(), "两行表头仍然合成一组");
        assertEquals(3, blocks.get(0).size(), "一组只有 headers/rows/name 三个键，name 不会重复写");
        assertEquals("上层", blocks.get(0).get("name"), "第一个命中的注册集说了算");
        assertEquals(Arrays.asList("客户名称", "合同金额", "已回款", "回款率"), TestSupport.headers(blocks.get(0)),
                "列名仍按多级表头合并：下层填了格的用下层");
        assertEquals(1, TestSupport.rows(blocks.get(0)).size());
    }

    @Test
    void registrationToleratesOneMismatchingColumnByDefault() {
        // 默认容忍 1 列不一致：文档里把「列Ewew」写成「列Efew」这类个别错字不会让整组表头失效
        WordParser.Options options = WordParser.options().headers("列Bb", "列Ewew", "列Bbb");

        assertEquals(0, options.matchHeader(Arrays.asList("列Bb", "列Efew", "列Bbb")), "错一个字仍命中");
        assertEquals(-1, options.matchHeader(Arrays.asList("甲", "乙", "丙")), "错三个字就不算了");
    }

    @Test
    void toleranceCanBeTightenedToAnExactMatch() {
        WordParser.Options options = WordParser.options().headers("列Bb", "列Ewew", "列Bbb").headerTolerance(0);

        assertEquals(-1, options.matchHeader(Arrays.asList("列Bb", "列Efew", "列Bbb")), "容忍度 0 要求逐列相等");
        assertEquals(0, options.matchHeader(Arrays.asList("列Bb", "列Ewew", "列Bbb")));
    }

    @Test
    void matchingIgnoresTrailingBlankCellsAndRequiresTheSameColumnCount() {
        WordParser.Options options = WordParser.options().headers("片区", "年度目标", "完成率");

        assertEquals(0, options.matchHeader(Arrays.asList("片区", "年度目标", "完成率", "")), "行尾空列不参与比较");
        assertEquals(-1, options.matchHeader(Arrays.asList("片区", "年度目标")), "列数不同就不算命中");
        assertEquals(-1, options.matchHeader(Arrays.asList("片区", "年度目标", "完成率", "备注")));
    }

    @Test
    void matchingIgnoresCaseAndEveryKindOfSpace() {
        WordParser.Options options = WordParser.options().headers("ItemId", "数量");

        assertEquals(0, options.matchHeader(Arrays.asList(" ITEMID ", "数\u3000量")));
    }

    @Test
    void toleranceNeverAllowsAWholeRowToMismatch() {
        // 容忍度再大也至少要求一列完全相等、且列数相同，否则数据行会被当成表头
        WordParser.Options options = WordParser.options().headers("a", "b").headerTolerance(9);

        assertEquals(0, options.matchHeader(Arrays.asList("a", "x")), "至少一列相等");
        assertEquals(-1, options.matchHeader(Arrays.asList("x", "y")), "一列都不等就不算");
    }

    @Test
    void keywordMatchRequiresEveryPopulatedCellToHit() {
        WordParser.Options options = WordParser.options().keywords("列");

        assertTrue(options.matchesKeywords(Arrays.asList("列a", "列b", "")), "空单元格不参与判断");
        assertFalse(options.matchesKeywords(Arrays.asList("列a", "值b")), "有一格不含关键字就不算表头");
        assertFalse(options.matchesKeywords(Arrays.asList("", "")), "整行空的也不算");
        assertFalse(WordParser.options().matchesKeywords(Arrays.asList("列a")), "没注册关键字时永远不命中");
    }

    /* ---------------- 跨整行合并的行：三种去向 ---------------- */

    @Test
    void separatorRowGoesToTheCurrentGroupByDefault() {
        // KEEP：分组行留在它上面那组，作为那组的一行数据；它后面那一行按表头处理（见下面那条规则），另起一组
        List<Map<String, Object>> blocks = group(rows(
                        new String[]{"片区", "目标"},
                        new String[]{"华东", "2400"},
                        new String[]{"2026 年度", ""},
                        new String[]{"片区", "目标"},
                        new String[]{"华北", "1800"}),
                flags(false, false, true, false, false), null, WordParser.options(), true);

        assertEquals(2, blocks.size());
        assertEquals(2, TestSupport.rows(blocks.get(0)).size(), "分组行留在它上面那组");
        assertEquals(Arrays.asList("2026 年度", ""), TestSupport.values(TestSupport.rows(blocks.get(0)).get(1)));
        assertEquals(1, TestSupport.rows(blocks.get(1)).size(), "分组行之后的表头行开新的一组");
        assertEquals(Arrays.asList("华北", "1800"), TestSupport.values(TestSupport.rows(blocks.get(1)).get(0)));
    }

    @Test
    void separatorRowBeforeAnyHeaderStartsANumberedGroup() {
        // 还没出现过表头就碰到分组行：另起一组（列名用序号），否则这行文字会无声丢掉
        List<Map<String, Object>> blocks = group(rows(
                        new String[]{"2025 年度", "", ""},
                        new String[]{"片区", "目标", "完成率"},
                        new String[]{"华东", "2400", "101%"}),
                flags(true, false, false), null, WordParser.options(), false);

        assertEquals(2, blocks.size());
        assertEquals(Arrays.asList("列1", "列2", "列3"), TestSupport.headers(blocks.get(0)));
        assertEquals(Arrays.asList("2025 年度", "", ""), TestSupport.values(TestSupport.rows(blocks.get(0)).get(0)));
        assertEquals(Arrays.asList("片区", "目标", "完成率"), TestSupport.headers(blocks.get(1)));
    }

    @Test
    void fullRowFirstCollectsSeparatorsIntoTheTablesFirstGroup() throws Exception {
        // FIRST：整张物理表里的分组行都归到该表的第一组（这里就是那张「还没有表头时另起的序号组」），
        // 数据组里不再混分组行；每组分组行之后的那一行仍按表头处理，所以数据被切成两组
        List<Map<String, Object>> blocks = Tables.group(rows(
                        new String[]{"2025 年度", "", ""},
                        new String[]{"片区", "目标", "完成率"},
                        new String[]{"华东", "2400", "101%"},
                        new String[]{"2026 年度", "", ""},
                        new String[]{"片区", "目标", "完成率"},
                        new String[]{"华东", "2800", "—"}),
                flags(true, false, false, true, false, false), null,
                WordParser.options().fullRow(WordParser.FullRow.FIRST), new LinkedHashSet<Integer>(), false);

        assertEquals(3, blocks.size());
        assertEquals(Arrays.asList("列1", "列2", "列3"), TestSupport.headers(blocks.get(0)));
        assertEquals(2, TestSupport.rows(blocks.get(0)).size(), "两行分组小标题都归到第一组");
        assertEquals(Arrays.asList("2026 年度", "", ""), TestSupport.values(TestSupport.rows(blocks.get(0)).get(1)));
        assertEquals(Arrays.asList("片区", "目标", "完成率"), TestSupport.headers(blocks.get(1)));
        assertEquals(1, TestSupport.rows(blocks.get(1)).size(), "数据组里不再有分组行");
        assertEquals(Arrays.asList("华东", "2800", "—"), TestSupport.values(TestSupport.rows(blocks.get(2)).get(0)));
    }

    @Test
    void fullRowDropDiscardsSeparators() {
        List<Map<String, Object>> blocks = Tables.group(rows(
                        new String[]{"片区", "目标"},
                        new String[]{"华东", "2400"},
                        new String[]{"2026 年度", ""}),
                flags(false, false, true), null, WordParser.options().fullRow(WordParser.FullRow.DROP),
                new LinkedHashSet<Integer>(), true);

        assertEquals(1, blocks.size());
        assertEquals(1, TestSupport.rows(blocks.get(0)).size());
    }

    @Test
    void separatorRowsStillCloseTheHeaderAreaEvenWhenDropped() {
        // 分组行即使被丢掉，也标志着「上面那组表头区结束了」：紧跟其后的行不会被并进上一层表头
        List<Map<String, Object>> blocks = Tables.group(rows(
                        new String[]{"片区", "目标"},
                        new String[]{"小标题", ""},
                        new String[]{"片区", "目标"},
                        new String[]{"华东", "2400"}),
                flags(false, true, false, false), null, WordParser.options().fullRow(WordParser.FullRow.DROP),
                new LinkedHashSet<Integer>(), true);

        assertEquals(2, blocks.size());
        assertEquals(0, TestSupport.rows(blocks.get(0)).size(), "第一组只有表头，没有数据");
        assertEquals(1, TestSupport.rows(blocks.get(1)).size());
    }

    /* ---------------- 网格形状 ---------------- */

    @Test
    void rowsOfDifferentLengthsArePaddedToTheWidestRow() {
        List<Map<String, Object>> blocks = group(rows(
                new String[]{"片区", "目标", "完成率"},
                new String[]{"华东", "2400"}), WordParser.options(), true);

        assertEquals(Arrays.asList("片区", "目标", "完成率"), TestSupport.headers(blocks.get(0)));
        assertEquals(Arrays.asList("华东", "2400", ""), TestSupport.values(TestSupport.rows(blocks.get(0)).get(0)),
                "短行补空串，保证每行键数量一致");
    }

    @Test
    void emptyGridYieldsNoGroups() {
        assertEquals(0, group(new String[0][], WordParser.options(), true).size());
    }
}
