package cn.net.tianxi.document;

import org.junit.jupiter.api.Test;

import java.util.Arrays;

import static org.junit.jupiter.api.Assertions.*;

/**
 * 文本框拆段落的单测。
 *
 * <p>不走 {@code PptParser.parse}：评估模式一个文本框只交第一段，端到端样例照不到「多段被压成一行」。
 * {@link PptParser#textLines(String)} 是拆行与清洗本身，不需要许可证。
 */
class PptParserTextTest {

    @Test
    void allThreeAsposeLineEndingsSplitParagraphs() {
        // Aspose.Slides 的 ITextFrame.getText() 用 \r 分段；只按 \n 拆时整框会经 Results.clean 变成一行
        assertEquals(Arrays.asList("第一段", "第二段", "第三段", "第四段"),
                PptParser.textLines("第一段\r第二段\r\n第三段\n第四段"));
    }

    @Test
    void blankParagraphsAndNullContributeNothing() {
        assertTrue(PptParser.textLines(null).isEmpty());
        assertEquals(Arrays.asList("只有这一段"), PptParser.textLines("\r\n\r\n只有这一段\r\n"));
    }
}
