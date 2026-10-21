package cn.net.tianxi.document;

import org.junit.jupiter.api.Test;

import java.io.File;
import java.util.Arrays;
import java.util.List;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertNull;
import static org.junit.jupiter.api.Assertions.assertTrue;

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
}
