package cn.net.tianxi.document;

import org.junit.jupiter.api.Assumptions;
import org.junit.jupiter.api.Test;

import java.io.File;
import java.nio.charset.Charset;
import java.nio.file.Files;
import java.util.ArrayList;
import java.util.List;
import java.util.Locale;
import java.util.Map;

import static org.junit.jupiter.api.Assertions.assertNotNull;
import static org.junit.jupiter.api.Assertions.assertTrue;

/**
 * 「看结果」用的测试：把解析结果原样打到控制台，不做逐格断言。
 *
 * <p>其余 9 个测试类都在断言实测值，改代码时能挡住回归，但看不出「这个文档到底解析成了什么样」。
 * 本类补的就是这一面：跑一条测试，控制台直接给出正文、每组表的表头与数据行、图片元信息和完整 JSON。
 * 只保留一条底线断言（{@link TestSupport#assertHeadersMatchRowKeys}），因为契约破了的话输出也没有参考价值。
 *
 * <h3>怎么跑</h3>
 * <pre>
 * mvn -B test -Dtest=PrintResultTest                                  # 三份样例 + 一组选项对比
 * mvn -B test -Dtest=PrintResultTest#printWordSample                  # 只看 docx
 * mvn -B test -Dtest=PrintResultTest#printAnyDocument -Ddoc="D:/你的文档.docx"   # 看自己的文档
 * </pre>
 * 在 IDE 里直接点某一条测试的运行按钮也一样。
 *
 * <h3>控制台编码</h3>
 * 直接看控制台就行：控制台输出按 JVM 的 {@code file.encoding} 编码，cmd 里跑 {@code mvn test} 是 GBK、
 * IDE 里通常是 UTF-8，两种情况下中文都正常显示。只有把输出重定向到文件、再按另一种编码打开时才会花
 * （例如 {@code mvn test > out.txt} 之后用 UTF-8 打开）。这时候看 {@code target/print/<样例名>-<选项>.txt} ——
 * 同一份输出的 UTF-8 副本，每条测试跑完都会在末尾打印它的绝对路径；它在 {@code target/} 下，{@code mvn clean} 会带走。
 * 要强制控制台走 UTF-8，加 {@code -DargLine="-Dfile.encoding=UTF-8"}（实测加不加都是全绿，只影响编码）。
 *
 * <h3>为什么不打 BASE64 模式</h3>
 * 默认 {@code META} 模式下三份样例的 JSON 分别是 4409／6186／3262 个字符，读得完；
 * 换成 {@code BASE64} 那张 670939 字节的 PNG 会把 JSON 撑到 899039 个字符，刷屏且看不出东西。
 * 要看图片内容请用 {@code ImageMode.FILE} + {@code imageDir(File)}，输出里会带落盘路径。
 */
class PrintResultTest {

    private static final Charset UTF_8 = Charset.forName("UTF-8");

    /** 输出的 UTF-8 副本落在这里；{@code mvn clean} 一起清掉。 */
    private static final File OUT_DIR = new File("target/print");

    /* ---------------- 三份样例 ---------------- */

    @Test
    void printWordSample() throws Exception {
        print(TestSupport.WORD, WordParser.options(), "静默");
    }

    @Test
    void printPdfSample() throws Exception {
        print(TestSupport.PDF, WordParser.options(), "静默");
    }

    @Test
    void printPptSample() throws Exception {
        print(TestSupport.PPT, WordParser.options(), "静默");
    }

    /** 同一份 docx 换选项后的差别：注册 5 组真表头 + 丢掉分组小标题行（7 组 30 行 → 6 组 28 行）。 */
    @Test
    void printWordSampleWithRegisteredHeadersAndDrop() throws Exception {
        print(TestSupport.WORD, TestSupport.registered().fullRow(WordParser.FullRow.DROP), "注册表头-DROP");
    }

    /* ---------------- 任意文档 ---------------- */

    /**
     * 看自己的文档：{@code -Ddoc="D:/a.pdf"}，按扩展名自动选解析器（doc/docx/pdf/ppt/pptx）。
     *
     * <p>不给 {@code doc} 就跳过（{@code Assumptions}）而不是失败 —— 这条只是入口，不是回归测试。
     */
    @Test
    void printAnyDocument() throws Exception {
        String path = System.getProperty("doc");
        Assumptions.assumeTrue(path != null && path.trim().length() > 0,
                "跳过：用 -Ddoc=\"D:/你的文档.docx\" 指定要看的那一份（支持 doc/docx/pdf/ppt/pptx）");
        print(path.trim(), WordParser.options(), "静默");
    }

    /* ---------------- 输出 ---------------- */

    /** 解析 → 底线契约断言 → 拼一份可读报告 → 打到控制台并存一份 UTF-8 文件。 */
    private static void print(String path, WordParser.Options options, String tag) throws Exception {
        File file = new File(path);
        assertTrue(file.isFile(), "文档不存在：" + file.getAbsolutePath());

        Map<String, Object> result = parse(path, options);
        TestSupport.assertHeadersMatchRowKeys(result);          // 契约破了，下面的输出也没有参考价值
        String json = WordParser.toJson(result);
        assertNotNull(json, "toJson 返回 null");

        String report = report(file, options, tag, result, json);
        System.out.println(report);

        assertTrue(OUT_DIR.isDirectory() || OUT_DIR.mkdirs(), "创建输出目录失败：" + OUT_DIR.getAbsolutePath());
        File copy = new File(OUT_DIR, fileName(file, tag) + ".txt");
        Files.write(copy.toPath(), report.getBytes(UTF_8));
        System.out.println("[PrintResultTest] 同一份输出的 UTF-8 文件：" + copy.getAbsolutePath());
    }

    /** 按扩展名选解析器；三个入口的返回结构一致，所以报告与 {@code toJson} 可以共用一份。 */
    private static Map<String, Object> parse(String path, WordParser.Options options) throws Exception {
        String name = path.toLowerCase(Locale.ROOT);
        if (name.endsWith(".pdf")) {
            return PdfParser.parse(path, options);
        }
        if (name.endsWith(".ppt") || name.endsWith(".pptx")) {
            return PptParser.parse(path, options);
        }
        if (name.endsWith(".doc") || name.endsWith(".docx")) {
            return WordParser.parse(path, options);
        }
        throw new IllegalArgumentException("只支持 doc/docx/pdf/ppt/pptx，收到：" + path);
    }

    /** 输出文件名：{@code sales-report.docx-静默.txt}。tag 里的空格与路径分隔符换成下划线。 */
    private static String fileName(File file, String tag) {
        return file.getName() + "-" + tag.replaceAll("[\\\\/:*?\"<>|\\s]+", "_");
    }

    private static String report(File file, WordParser.Options options, String tag,
                                 Map<String, Object> result, String json) {
        List<String> paragraphs = splitLines(TestSupport.text(result));
        List<Map<String, Object>> tables = TestSupport.tables(result);
        List<Map<String, Object>> images = TestSupport.images(result);

        StringBuilder sb = new StringBuilder(64 + json.length());
        line(sb, "==============================================================================");
        line(sb, "文档：" + file.getAbsolutePath() + "（" + file.length() + " 字节）");
        line(sb, "选项：" + tag + " —— " + describe(options));
        line(sb, "规模：正文 " + paragraphs.size() + " 段，表格 " + tables.size() + " 组 / "
                + TestSupport.totalRows(result) + " 行，图片 " + images.size() + " 张");
        line(sb, "==============================================================================");

        sb.append('\n');
        line(sb, "---- text（" + paragraphs.size() + " 段，表格里的文字不在这里）----");
        if (paragraphs.isEmpty()) {
            line(sb, "（空：这份文档的正文全落在表格矩形内，按约定只在 tables 里出现）");
        }
        for (int i = 0; i < paragraphs.size(); i++) {
            line(sb, "[" + (i + 1) + "] " + paragraphs.get(i));
        }

        sb.append('\n');
        line(sb, "---- tables（" + tables.size() + " 组）----");
        for (int t = 0; t < tables.size(); t++) {
            Map<String, Object> table = tables.get(t);
            List<String> headers = TestSupport.headers(table);
            List<Map<String, String>> rows = TestSupport.rows(table);
            sb.append('\n');
            line(sb, "#" + (t + 1) + "  " + headers.size() + " 列 × " + rows.size() + " 行");
            line(sb, "headers = " + headers);
            for (int r = 0; r < rows.size(); r++) {
                line(sb, "  行" + (r + 1) + ": " + cells(rows.get(r)));
            }
        }

        sb.append('\n');
        line(sb, "---- images（" + images.size() + " 张，META 模式只有元信息）----");
        for (int i = 0; i < images.size(); i++) {
            line(sb, "[" + (i + 1) + "] " + images.get(i));
        }

        sb.append('\n');
        line(sb, "---- JSON（" + json.length() + " 个字符）----");
        sb.append(json).append('\n');
        return sb.toString();
    }

    /** 选项摘要：静默与否、注册了几组表头、分组行去向、图片模式。 */
    private static String describe(WordParser.Options options) {
        StringBuilder sb = new StringBuilder();
        if (options.isSilent()) {
            sb.append("静默模式（未注册表头、未给关键字，允许首行兜底当表头）");
        } else {
            sb.append("注册了 ").append(options.headerGroupCount()).append(" 组表头");
            for (int i = 0; i < options.headerGroupCount(); i++) {
                sb.append("\n       第").append(i + 1).append("组 = ").append(options.headerTitles(i));
            }
        }
        sb.append("\n       FullRow = ").append(options.fullRow())
                .append("，ImageMode = ").append(options.images());
        if (options.imageDir() != null) {
            sb.append("，imageDir = ").append(options.imageDir().getAbsolutePath());
        }
        return sb.toString();
    }

    /** 一行数据按 {@code 列名=值} 拼出来，顺序就是 {@code headers} 的顺序（行对象是 {@code LinkedHashMap}）。 */
    private static String cells(Map<String, String> row) {
        StringBuilder sb = new StringBuilder();
        for (Map.Entry<String, String> cell : row.entrySet()) {
            if (sb.length() > 0) {
                sb.append(" | ");
            }
            sb.append(cell.getKey()).append('=').append(cell.getValue());
        }
        return sb.toString();
    }

    /** 按行切开；空串（例如 PDF 样例的 {@code text}）不给任何行，而不是给一个空行。 */
    private static List<String> splitLines(String text) {
        List<String> lines = new ArrayList<String>();
        if (text == null || text.length() == 0) {
            return lines;
        }
        for (String line : text.split("\n", -1)) {
            lines.add(line);
        }
        return lines;
    }

    private static void line(StringBuilder sb, String text) {
        sb.append(text).append('\n');
    }
}
