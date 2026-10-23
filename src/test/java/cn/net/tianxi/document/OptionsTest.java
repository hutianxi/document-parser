package cn.net.tianxi.document;

import org.junit.jupiter.api.Test;

import java.io.File;
import java.util.Arrays;
import java.util.List;

import static org.junit.jupiter.api.Assertions.*;

/**
 * {@link WordParser.Options} 的单元测试：默认值、null 容错、注册表头的两种写法、模式判定。
 *
 * <p>匹配算法本身（列数、容忍度、行尾空列、大小写与空格）在 {@code TablesTest} 里跟着分组一起测；
 * 三种格式共用同一个 Options 实例见 {@code CrossFormatTest}。
 */
class OptionsTest {

    @Test
    void defaultsAreTheCheapestAndLeastSurprisingChoices() {
        WordParser.Options options = WordParser.options();

        assertEquals(WordParser.FullRow.KEEP, options.fullRow(), "分组行默认保留在它上面那组");
        assertEquals(WordParser.ImageMode.META, options.images(), "图片默认只给元信息，不占体积");
        assertNull(options.imageDir(), "默认不落盘");
        assertEquals(0, options.headerGroupCount(), "默认没注册表头");
        assertTrue(options.isSilent(), "既没注册表头也没关键字，就是静默模式");
        assertFalse(options.isInlineImageMode(), "META 不内嵌内容");
    }

    @Test
    void nullModesFallBackToTheDefaults() {
        // 链式选项常被配置项驱动（读不到配置就是 null），null 不该变成 NPE 或意外行为
        WordParser.Options options = WordParser.options().fullRow(null).images(null);

        assertEquals(WordParser.FullRow.KEEP, options.fullRow());
        assertEquals(WordParser.ImageMode.META, options.images());
    }

    @Test
    void blankAndNullKeywordsAreIgnored() {
        WordParser.Options options = WordParser.options().keywords("列", null, "  ");

        assertFalse(options.isSilent(), "给了有效关键字就不再是静默模式");
        // 关键字判定是「行内每个非空格都要含关键字之一」，所以整行都姓「列」才算表头
        assertTrue(options.matchesKeywords(Arrays.asList("列a", "列b")), "有效关键字照常生效");
        assertFalse(options.matchesKeywords(Arrays.asList("列a", "值b")), "有一格不含关键字就不算表头");
        // 空串会被 contains 无条件命中：null 和 "  " 真被收进去的话这一行也会判成表头
        assertFalse(options.matchesKeywords(Arrays.asList("客户", "金额")), "空关键字要被丢掉");
    }

    @Test
    void keywordsOnlyIsNotSilentSoTheFirstRowFallbackIsOff() {
        // isSilent 决定「首行能不能兜底当表头」：只要给了任何表头依据，兜底就关掉
        assertFalse(WordParser.options().keywords("列").isSilent());
        assertFalse(WordParser.options().headers("列a").isSilent());
        assertTrue(WordParser.options().isSilent());
    }

    @Test
    void headersCanBeRegisteredOneGroupAtATimeOrInBulk() {
        WordParser.Options oneByOne = WordParser.options()
                .headers("片区", "目标")
                .headers("客户", "金额");

        List<List<String>> groups = Arrays.asList(Arrays.asList("片区", "目标"), Arrays.asList("客户", "金额"));
        WordParser.Options bulk = WordParser.options().headerGroups(groups);

        assertEquals(2, oneByOne.headerGroupCount());
        assertEquals(2, bulk.headerGroupCount());
        // 两种写法等价：归一化后的比对名与用于输出的原名都一致
        assertEquals(oneByOne.headerGroup(1), bulk.headerGroup(1));
        assertEquals(oneByOne.headerTitles(1), bulk.headerTitles(1));
        assertEquals(1, oneByOne.matchHeader(Arrays.asList("客户", "金额")), "命中的是第 1 组");

        // 起名绑的是「最近一次注册的那一组」，批量注册内部就是按顺序逐组注册 —— 所以两种写法里名字都落在最后一组
        oneByOne.name("台账");
        bulk.name("台账");
        assertEquals("台账", oneByOne.headerName(1));
        assertEquals("", oneByOne.headerName(0), "前面的组不带名字");
        assertEquals(bulk.headerName(1), oneByOne.headerName(1), "批量与逐组在这里也等价");
        assertEquals("", bulk.headerName(0));
    }

    @Test
    void registeredNamesKeepTheirOriginalCaseWhileMatchingIsNormalized() {
        WordParser.Options options = WordParser.options().headers(" ItemId ", "数量");

        // 比对用归一化名（小写去空格），输出用注册时写的原名（只去首尾空格）
        assertEquals(Arrays.asList("itemid", "数量"), options.headerGroup(0));
        assertEquals(Arrays.asList("ItemId", "数量"), options.headerTitles(0));
        assertEquals(0, options.matchHeader(Arrays.asList("ITEMID", "数 量")));
    }

    @Test
    void nullTitlesBecomeEmptyNamesButStillCountAsColumns() {
        WordParser.Options options = WordParser.options().headers("片区", null);

        assertEquals(2, options.headerGroup(0).size(), "null 也算一列，否则列数比对会错位");
        assertEquals(Arrays.asList("片区", ""), options.headerTitles(0));
        // 已知取舍：行尾空列不参与比较，所以末列是空的注册组永远命不中 —— 注册时别把空列名放在最后
        assertEquals(-1, options.matchHeader(Arrays.asList("片区", "")));
    }

    @Test
    void nameBindsToTheGroupRegisteredJustBeforeIt() {
        WordParser.Options options = WordParser.options()
                .headers("片区", "目标").name("年度目标")
                .headers("客户", "金额").name("回款台账")
                .headers("工单", "数量");

        assertEquals("年度目标", options.headerName(0), "name 归它前面最近那次 headers 注册的那一组");
        assertEquals("回款台账", options.headerName(1));
        assertEquals("", options.headerName(2), "没起名字的组是空串，结果里不会出现 name 键");

        // 连写两次以后者为准：配置驱动的写法里「先给默认名再覆盖」很常见
        options.name("工单统计");
        assertEquals("工单统计", options.headerName(2), "覆盖的是最后注册的那一组");
        assertEquals("回款台账", options.headerName(1), "前面几组不受影响");
    }

    @Test
    void blankNullOrPrematureNameLeavesTheGroupUnnamed() {
        // 一组都没注册时起名字：没有归属，提示一句并忽略，不抛异常（与 imageDir 的处理方式一致）
        WordParser.Options before = WordParser.options().name("先起名后注册");
        assertEquals(0, before.headerGroupCount(), "name 不会凭空造出一组表头");

        WordParser.Options options = WordParser.options()
                .headers("片区", "目标").name("  年度目标  ")
                .headers("客户", "金额").name("   ")
                .headers("工单", "数量").name(null);

        assertEquals("年度目标", options.headerName(0), "只去首尾空格，中间的空格保留");
        assertEquals("", options.headerName(1), "全空格等于没起名");
        assertEquals("", options.headerName(2), "null 等于没起名");
    }

    @Test
    void negativeToleranceIsClampedToAnExactMatch() {
        WordParser.Options options = WordParser.options().headers("片区", "目标").headerTolerance(-3);

        assertEquals(-1, options.matchHeader(Arrays.asList("片区", "别的")), "负数按 0 处理：必须逐列相等");
        assertEquals(0, options.matchHeader(Arrays.asList("片区", "目标")));
    }

    @Test
    void inlineImageModesAreExactlyBase64AndBytes() {
        assertTrue(WordParser.options().images(WordParser.ImageMode.BASE64).isInlineImageMode());
        assertTrue(WordParser.options().images(WordParser.ImageMode.BYTES).isInlineImageMode());
        assertFalse(WordParser.options().images(WordParser.ImageMode.FILE).isInlineImageMode(),
                "FILE 每张都要落盘，不能去重");
        assertFalse(WordParser.options().images(WordParser.ImageMode.META).isInlineImageMode());
    }

    @Test
    void imageDirIsKeptAsGiven() {
        File directory = new File("target/images");
        assertEquals(directory, WordParser.options().imageDir(directory).imageDir());
    }

    @Test
    void sourceIsWrittenByDefaultAndCanBeSwitchedOff() {
        assertTrue(WordParser.options().writesSource(), "默认带 source：核对原文要翻页码");
        assertFalse(WordParser.options().source(false).writesSource());
        assertTrue(WordParser.options().source(false).source(true).writesSource(), "开关可以再来回去");
    }

    @Test
    void unmatchedHeadersIsOffByDefaultAndCanBeSwitchedOn() {
        // 与 source 相反，这个键默认关：它是顶层键，多一个会让按 POJO 反序列化的下游撞上未知字段
        assertFalse(WordParser.options().writesUnmatchedHeaders(), "默认不带 unmatchedHeaders");
        assertTrue(WordParser.options().unmatchedHeaders(true).writesUnmatchedHeaders());
        assertFalse(WordParser.options().unmatchedHeaders(true).unmatchedHeaders(false).writesUnmatchedHeaders(),
                "开关可以再来回去");
        assertTrue(WordParser.options().source(false).unmatchedHeaders(true).writesUnmatchedHeaders(),
                "与 source 开关互不影响");
    }
}
