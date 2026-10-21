package cn.net.tianxi.document;

import org.junit.jupiter.api.function.Executable;

import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.io.PrintStream;
import java.nio.charset.Charset;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;
import java.util.zip.ZipEntry;
import java.util.zip.ZipOutputStream;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertNotNull;
import static org.junit.jupiter.api.Assertions.assertTrue;

/**
 * 测试共用工具：取样例路径、从结果 Map 里取三个集合、现场拼最小 OOXML 部件、收标准错误输出。
 *
 * <h3>三份样例是同一份业务内容</h3>
 * 「2025 年度华东区销售业绩报告」：季度销售明细（核心表）、客户回款台账（长表）、售后工单统计（含单元格内嵌套表）、
 * 分片区年度目标（跨整行合并的分组行 + 无标题列）、系统看板截图（一张图）。三种格式各存一份，
 * 因此「同一份内容三种格式读出来一致吗」可以直接逐格比对（见 {@code CrossFormatTest}）。
 *
 * <h3>为什么期望值可以写死（有 lic / 没 lic 都成立）</h3>
 * 评估模式的痕迹按产品不同，样例是照着这些实测边界造的（依据见 docs.aspose.com 各产品的 Licensing 页）：
 * <ul>
 *   <li><b>Aspose.Words</b>：评估模式注入水印段落、并把规模限制在几百个段落。水印由
 *       {@code Results.cleanContent} 过滤掉，样例只有 13 段正文，所以 {@code text} 与 licensed 一致；</li>
 *   <li><b>Aspose.PDF</b>：评估模式限制的是文字层。实测这一条落在段落抽取上（{@code ParagraphAbsorber} 每页只给
 *       前几段），表格吸收器与图片不受影响：31 行的台账整表读全了、1 张图片的字节也完整。
 *       样例正文全部在表格里，页面之间只用空段落撑开（空段落不进 {@code text}），所以 {@code text} 恒为 {@code ""}，
 *       这条限制在本样例上量不出具体数字；</li>
 *   <li><b>Aspose.Slides</b>：评估模式读文本框时只给<b>第一段</b>、且截到 5 个字符再追加提示。样例里每个文本框
 *       都只有一段、不超过 5 个字符（{@code 2025年}、{@code 销售业绩}、{@code 华东大区}…），提示由
 *       {@code PptParser.stripEvaluationNotice} 去掉，所以 licensed 下读到的也是同一串。</li>
 * </ul>
 *
 * <p>代价说清楚：PPT 样例的文本框为了迁就评估模式写得很短，它验证得了「形状/组合形状/表格/图片都读到了、顺序对」，
 * 验证不了「长文本能不能完整读出来」—— 那一条由 docx 样例覆盖（它的段落有 60 多个字符，读出来是完整的）。
 * PDF 样例是<b>真 Word 导出</b>的（不是 Aspose 转的），文件里没有烤进去的评估水印。
 */
final class TestSupport {

    /** 报告正文（13 段）+ 四张表（五组表头）+ 一张看板截图：结构最全的样例。 */
    static final String WORD = "src/test/resources/word/sales-report.docx";

    /** 同一份内容用 Word 导出的 3 页 PDF：31 行的台账跨页断开，靠几何判据并回一张；正文全在表格里。 */
    static final String PDF = "src/test/resources/pdf/sales-report.pdf";

    /** 同一份内容的 3 页幻灯片：文本框、组合形状、四张表（跨行/跨列合并）、一张图片。 */
    static final String PPT = "src/test/resources/ppt/sales-report.pptx";

    /** 全部样例，契约类断言（列名对齐、图片元信息）逐个跑一遍。 */
    static final String[] ALL_SAMPLES = {WORD, PDF, PPT};

    /**
     * 样例里出现的五组表头，逐组注册。
     *
     * <p>核心表两组同名列（{@code 销售额}/{@code 回款率} 各出现两次）注册时故意写成一样的，
     * 补名规则应当把它们还原成 {@code 销售额-5}/{@code 回款率-6}，与静默模式读到的列名逐字相同。
     */
    static WordParser.Options registered() {
        return WordParser.options()
                .headers("片区", "区域", "销售额", "回款率", "销售额", "回款率", "备注")
                .headers("客户名称", "合同金额", "已回款", "回款率", "备注")
                .headers("工单类型", "工单数", "处理说明")
                .headers("故障部件", "工单数")
                .headers("片区", "年度目标", "完成率");
    }

    private static final Charset UTF_8 = Charset.forName("UTF-8");

    private TestSupport() {
    }

    /* ---------------- 读结果里的三个集合 ---------------- */

    static String text(Map<String, Object> result) {
        return (String) result.get("text");
    }

    @SuppressWarnings("unchecked")
    static List<Map<String, Object>> tables(Map<String, Object> result) {
        return (List<Map<String, Object>>) result.get("tables");
    }

    @SuppressWarnings("unchecked")
    static List<Map<String, Object>> images(Map<String, Object> result) {
        return (List<Map<String, Object>>) result.get("images");
    }

    static Map<String, Object> table(Map<String, Object> result, int index) {
        List<Map<String, Object>> tables = tables(result);
        assertTrue(index < tables.size(), "只有 " + tables.size() + " 组表头，取不到第 " + index + " 组");
        return tables.get(index);
    }

    @SuppressWarnings("unchecked")
    static List<String> headers(Map<String, Object> table) {
        return (List<String>) table.get("headers");
    }

    @SuppressWarnings("unchecked")
    static List<Map<String, String>> rows(Map<String, Object> table) {
        return (List<Map<String, String>>) table.get("rows");
    }

    static Map<String, Object> image(Map<String, Object> result, int index) {
        List<Map<String, Object>> images = images(result);
        assertTrue(index < images.size(), "只有 " + images.size() + " 张图片，取不到第 " + index + " 张");
        return images.get(index);
    }

    /** 一行数据按列顺序的值（{@code LinkedHashMap} 保证顺序 = {@code headers} 顺序）。 */
    static List<String> values(Map<String, String> row) {
        return new ArrayList<String>(row.values());
    }

    /** 第 {@code index} 组第 {@code row} 行的值，按列顺序。 */
    static List<String> values(Map<String, Object> result, int index, int row) {
        List<Map<String, String>> rows = rows(table(result, index));
        assertTrue(row < rows.size(), "第 " + index + " 组只有 " + rows.size() + " 行，取不到第 " + row + " 行");
        return values(rows.get(row));
    }

    /** 所有组的行数之和。 */
    static int totalRows(Map<String, Object> result) {
        int total = 0;
        for (Map<String, Object> table : tables(result)) {
            total += rows(table).size();
        }
        return total;
    }

    /* ---------------- 契约断言 ---------------- */

    /**
     * 结果契约：每一组的 {@code headers} 逐个等于它每一行的键（顺序与数量都一致），且值不为 null。
     *
     * <p>这是「可以只读 {@code headers} 当列清单去建表／映射字段」的前提，由 {@code Tables.normalizeHeaders}
     * 与 {@code Tables.rowMap} 共用同一条补名规则来保证（空表头补 {@code 列N}、同名列补 {@code 名-列号}）。
     */
    static void assertHeadersMatchRowKeys(Map<String, Object> result) {
        for (Map<String, Object> table : tables(result)) {
            List<String> headers = headers(table);
            for (String header : headers) {
                assertNotNull(header, "headers 里出现 null");
                assertTrue(header.length() > 0, "headers 里出现空列名：" + headers);
            }
            for (Map<String, String> row : rows(table)) {
                assertEquals(headers, new ArrayList<String>(row.keySet()),
                        "行键与 headers 不一致，headers=" + headers + "，行=" + row);
                for (Map.Entry<String, String> cell : row.entrySet()) {
                    assertNotNull(cell.getValue(), "单元格 " + cell.getKey() + " 的值是 null");
                }
            }
        }
    }

    /** 图片元信息契约：内嵌图片一定有文件名、类型、宽高、字节数与内容指纹；外链图片只给 link。 */
    static void assertImageMeta(Map<String, Object> image) {
        assertNotNull(image.get("fileName"), "图片缺 fileName：" + image);
        assertNotNull(image.get("type"), "图片缺 type：" + image);
        assertNotNull(image.get("width"), "图片缺 width：" + image);
        assertNotNull(image.get("height"), "图片缺 height：" + image);
        if (image.containsKey("link")) {
            return;                                             // 外链图片的字节不在文档里，没有 size/contentId
        }
        assertNotNull(image.get("size"), "图片缺 size：" + image);
        String contentId = String.valueOf(image.get("contentId"));
        assertEquals(40, contentId.length(), "contentId 应该是 SHA-1 十六进制（40 位）：" + contentId);
        assertTrue(contentId.matches("[0-9a-f]{40}"), "contentId 不是十六进制：" + contentId);
    }

    /* ---------------- 标准错误 ---------------- */

    /**
     * 执行一段代码并返回它写到标准错误的内容（未匹配表头、图片模式配错这类提示都走 stderr）。
     *
     * <p>{@link Executable#execute()} 声明 {@code throws Throwable}，这里把断言失败原样抛出（否则失败原因会被
     * 包一层，看不到 JUnit 的期望值对比），其余受检异常按原类型抛出。
     */
    static String captureStderr(Executable body) throws Exception {
        PrintStream original = System.err;
        ByteArrayOutputStream buffer = new ByteArrayOutputStream();
        PrintStream captured = new PrintStream(buffer, true, "UTF-8");
        System.setErr(captured);
        try {
            body.execute();
        } catch (Throwable thrown) {
            if (thrown instanceof AssertionError) {
                throw (AssertionError) thrown;
            }
            if (thrown instanceof Exception) {
                throw (Exception) thrown;
            }
            throw new IllegalStateException(thrown);
        } finally {
            captured.flush();
            System.setErr(original);
        }
        return new String(buffer.toByteArray(), "UTF-8");
    }

    /* ---------------- 现场拼最小 docx ---------------- */
    /* 只有 4 个部件的最小 OOXML 包，Aspose.Words 能直接加载，用来造样例文件里不值得专门存一份的边角结构
       （单元格里的内容控件、嵌套表、域、软换行）。不用 Aspose 自己保存样例 —— 没有 lic 时它保存阶段就会把文本截断。 */

    private static final String CONTENT_TYPES = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
            + "<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">"
            + "<Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/>"
            + "<Default Extension=\"xml\" ContentType=\"application/xml\"/>"
            + "<Override PartName=\"/word/document.xml\" ContentType=\"application/vnd.openxmlformats-officedocument"
            + ".wordprocessingml.document.main+xml\"/></Types>";

    private static final String ROOT_RELS = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
            + "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">"
            + "<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006"
            + "/relationships/officeDocument\" Target=\"word/document.xml\"/></Relationships>";

    private static final String DOCUMENT_RELS = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
            + "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"/>";

    private static final String DOCUMENT_HEAD = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
            + "<w:document xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"><w:body>";

    /**
     * 把若干 body 级部件写成一个最小 docx。
     *
     * @param directory 输出目录（用 JUnit 的 {@code @TempDir}），文件已存在时覆盖
     * @param fileName  文件名，例如 {@code nested.docx}
     * @param bodyParts body 里的内容，用 {@link #p}、{@link #tbl}、{@link #sdt} 这些拼
     */
    static File docx(File directory, String fileName, String... bodyParts) throws IOException {
        StringBuilder body = new StringBuilder(DOCUMENT_HEAD);
        for (String part : bodyParts) {
            body.append(part);
        }
        body.append("</w:body></w:document>");

        File file = new File(directory, fileName);
        OutputStream out = new FileOutputStream(file);
        try {
            ZipOutputStream zip = new ZipOutputStream(out, UTF_8);
            entry(zip, "[Content_Types].xml", CONTENT_TYPES);
            entry(zip, "_rels/.rels", ROOT_RELS);
            entry(zip, "word/_rels/document.xml.rels", DOCUMENT_RELS);
            entry(zip, "word/document.xml", body.toString());
            zip.finish();
        } finally {
            out.close();
        }
        return file;
    }

    private static void entry(ZipOutputStream zip, String name, String content) throws IOException {
        zip.putNextEntry(new ZipEntry(name));
        zip.write(content.getBytes(UTF_8));
        zip.closeEntry();
    }

    /** 一个只有一段文字的段落。 */
    static String p(String text) {
        return "<w:p><w:r><w:t xml:space=\"preserve\">" + escape(text) + "</w:t></w:r></w:p>";
    }

    /** 空段落。单元格里的嵌套表后面必须跟一个段落，OOXML 的硬性要求。 */
    static String emptyP() {
        return "<w:p/>";
    }

    /**
     * 带域的段落：{@code 前文 + 域 + 后文}。
     *
     * <p>Word 把域存成 begin / instrText / separate / 域结果 / end 五段，{@code Node.getText()} 读出来是
     * {@code \u0013 域代码 \u0014 域结果 \u0015}，{@code Results.clean} 负责丢掉代码只留结果。
     *
     * @param instruction 域代码，例如 {@code PAGE}
     * @param cached      文档里缓存的域结果；没算过的域这里是空的
     */
    static String pField(String before, String instruction, String cached, String after) {
        return "<w:p><w:r><w:t xml:space=\"preserve\">" + escape(before) + "</w:t></w:r>"
                + "<w:r><w:fldChar w:fldCharType=\"begin\"/></w:r>"
                + "<w:r><w:instrText xml:space=\"preserve\">" + escape(instruction) + "</w:instrText></w:r>"
                + "<w:r><w:fldChar w:fldCharType=\"separate\"/></w:r>"
                + "<w:r><w:t xml:space=\"preserve\">" + escape(cached) + "</w:t></w:r>"
                + "<w:r><w:fldChar w:fldCharType=\"end\"/></w:r>"
                + "<w:r><w:t xml:space=\"preserve\">" + escape(after) + "</w:t></w:r></w:p>";
    }

    /** 用软换行（Shift+Enter，{@code <w:br/>}）分成两行的段落。 */
    static String pSoftBreak(String firstLine, String secondLine) {
        return "<w:p><w:r><w:t xml:space=\"preserve\">" + escape(firstLine) + "</w:t>"
                + "<w:br/>"
                + "<w:t xml:space=\"preserve\">" + escape(secondLine) + "</w:t></w:r></w:p>";
    }

    /** 一张表；{@code tblLook firstRow=1} 与 Word 里新建表格的默认值一致（首行按标题行处理）。 */
    static String tbl(int columns, String... rows) {
        StringBuilder xml = new StringBuilder("<w:tbl><w:tblPr><w:tblStyle w:val=\"TableGrid\"/>"
                + "<w:tblW w:w=\"0\" w:type=\"auto\"/>"
                + "<w:tblLook w:val=\"0020\" w:firstRow=\"1\" w:lastRow=\"0\" w:firstColumn=\"0\" w:lastColumn=\"0\""
                + " w:noHBand=\"0\" w:noVBand=\"0\"/></w:tblPr><w:tblGrid>");
        for (int c = 0; c < columns; c++) {
            xml.append("<w:gridCol w:w=\"2000\"/>");
        }
        xml.append("</w:tblGrid>");
        for (String row : rows) {
            xml.append(row);
        }
        return xml.append("</w:tbl>").toString();
    }

    static String tr(String... cells) {
        StringBuilder xml = new StringBuilder("<w:tr>");
        for (String cell : cells) {
            xml.append(cell);
        }
        return xml.append("</w:tr>").toString();
    }

    /** 普通单元格。 */
    static String tc(String text) {
        return tcInner(p(text));
    }

    /** 跨 {@code gridSpan} 列合并的起始格；整行只有一个起始格时即「跨整行合并」的分组行。 */
    static String tcSpan(String text, int gridSpan) {
        return "<w:tc><w:tcPr><w:tcW w:w=\"2000\" w:type=\"dxa\"/><w:gridSpan w:val=\"" + gridSpan + "\"/></w:tcPr>"
                + p(text) + "</w:tc>";
    }

    /** 纵向合并的起始格。 */
    static String tcVerticalStart(String text) {
        return "<w:tc><w:tcPr><w:tcW w:w=\"2000\" w:type=\"dxa\"/><w:vMerge w:val=\"restart\"/></w:tcPr>"
                + p(text) + "</w:tc>";
    }

    /** 纵向合并的延续格：自己没有内容，解析时回填起始格的值。 */
    static String tcVerticalContinue() {
        return "<w:tc><w:tcPr><w:tcW w:w=\"2000\" w:type=\"dxa\"/><w:vMerge/></w:tcPr><w:p/></w:tc>";
    }

    /** 单元格里放任意部件（嵌套表、内容控件），调用方自己保证以段落结尾。 */
    static String tcInner(String innerXml) {
        return "<w:tc><w:tcPr><w:tcW w:w=\"2000\" w:type=\"dxa\"/></w:tcPr>" + innerXml + "</w:tc>";
    }

    /** 内容控件（{@code w:sdt}）：Word 的模板和表单常用，Body 之下、单元格之内都能出现。 */
    static String sdt(String... content) {
        StringBuilder xml = new StringBuilder("<w:sdt><w:sdtPr><w:id w:val=\"1\"/></w:sdtPr><w:sdtContent>");
        for (String part : content) {
            xml.append(part);
        }
        return xml.append("</w:sdtContent></w:sdt>").toString();
    }

    /**
     * 没填写内容的「纯文本」控件，连 Word 自己写的占位文字（{@code w:placeholderText}）一起拼出来。
     *
     * <p>实测：那句占位文字不会进结果 —— 它属于 Word 的显示层，不是文档内容。
     */
    static String emptyTextSdt() {
        return "<w:sdt><w:sdtPr><w:id w:val=\"2\"/>"
                + "<w:placeholderText w:val=\"Click here to enter text.\"/><w:text/></w:sdtPr>"
                + "<w:sdtContent><w:p/></w:sdtContent></w:sdt>";
    }

    private static String escape(String text) {
        return text.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;");
    }
}
