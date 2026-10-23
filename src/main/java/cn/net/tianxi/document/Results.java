package cn.net.tianxi.document;

import com.fasterxml.jackson.databind.ObjectMapper;
import com.fasterxml.jackson.databind.SerializationFeature;

import java.io.File;
import java.io.FileInputStream;
import java.io.InputStream;
import java.util.*;
import java.util.regex.Pattern;

/**
 * 三种格式（Word/PDF/PPT）共用的输出与文本处理工具：JSON 序列化、图片落盘、许可证注册、文本清洗、表头归一化。
 *
 * <p>抽出来的目的：{@code WordParser.toJson}、许可证加载、文本清洗这类逻辑如果每个解析器各抄一份，
 * 后面改行为就要改三处，容易漏。本类全部是包内可见的静态方法，不对外暴露；三个解析器对外暴露的
 * {@code toJson} 与 {@code writeImages}（签名一致）只是转发到 {@link #toJson} 与 {@link #writeImages}。
 *
 * <h2>文本清洗：样例 docx 的三段正文</h2>
 * {@link #clean} 的三条规则在 {@code src/test/resources/word/sales-report.docx} 的实测正文里各有对应
 * （{@code WordParserTest.textIsTheThirteenBodyParagraphs} 把这 13 段逐字断言下来）：
 * <table border="1">
 *   <caption>样例 docx 正文与清洗规则</caption>
 *   <tr><th>清洗后</th><th>命中的规则</th></tr>
 *   <tr><td>{@code 报告页次：3 页}</td>
 *       <td>{@code PAGE} 域算过，缓存的域结果 {@code 3} 保留，域代码 {@code PAGE} 丢弃；
 *           原样输出会是 {@code 报告页次：PAGE 3 页}</td></tr>
 *   <tr><td>{@code 归档编号：}</td>
 *       <td>{@code DOCPROPERTY} 域从没算过、域结果为空，整块消失，只剩域前面的文字</td></tr>
 *   <tr><td>{@code …统计口径为含税销售额。 数据来源：销售管理系统…}</td>
 *       <td>Shift+Enter 的软换行 {@code \u000b} 换成一个空格，所以句号后面那个空格是它留下的痕迹；
 *           直接丢弃会粘成 {@code 含税销售额。数据来源：}</td></tr>
 * </table>
 *
 * <p>评估水印的过滤同样在这条链上：没有 lic 时 Aspose.Words 会把
 * {@code Evaluation Only. Created with Aspose.Words. Copyright 2001-2023 Aspose Pty Ltd.} 注进页眉页脚，
 * 而样例正文的 13 段里没有一句水印 —— 走的是 {@link #cleanContent}（{@link #clean} + {@link #isEvaluationNotice}）。
 *
 * <h2>表头归一化</h2>
 * {@link #normalize} 只服务表头比对：去一切空格 + {@link Locale#ROOT} 小写。样例三份文档的表头都不需要它
 * （docx 与 pdf 的 {@code tables} 在静默模式下 JSON 逐字相同，说明列名本来就对得上），
 * 它防的是 PDF 单元格里因换行多出来的空格、以及 Word 里手打的全角空格 —— 这些在样例里没有实例，
 * 覆盖它的是 {@code ResultsTest} 的 28 条单元测试（半角/全角/{@code \u00a0}/{@code \u2007}/{@code \u202f}/制表符）。
 *
 * <h2>输出体积</h2>
 * {@link #toJson} 用的是缩进输出，体积差异全部来自图片内容那一项。三份样例实测（字符数）：
 * <table border="1">
 *   <caption>{@code toJson} 结果长度：默认 META vs BASE64</caption>
 *   <tr><th>样例</th><th>{@code META}（默认）</th><th>{@code BASE64}</th></tr>
 *   <tr><td>sales-report.docx（PNG 1024×768，670939 字节）</td><td>4740</td><td>899370</td></tr>
 *   <tr><td>sales-report.pdf（JPEG 699×524，38668 字节）</td><td>6473</td><td>58076</td></tr>
 *   <tr><td>sales-report.pptx（同 docx 那张 PNG）</td><td>3526</td><td>898156</td></tr>
 * </table>
 * 这就是 {@link WordParser.ImageMode} 默认选 {@code META} 的理由：一份 670KB 的截图会让 JSON 从 4.7KB 涨到 899KB。
 *
 * <p>单元测试见 {@code ResultsTest}（28 条），其中 {@link #safeName} 与 {@link #writeImages} 的越权写入、
 * 重名覆盖、无内容告警各有专门用例。
 */
final class Results {

    /**
     * ObjectMapper 线程安全，建一次复用即可；原来每次 {@code toJson} 都新建一个。
     */
    private static final ObjectMapper MAPPER = new ObjectMapper().enable(SerializationFeature.INDENT_OUTPUT);

    /**
     * {@link #clean} 用：连续空格压成一个。
     */
    private static final Pattern TEXT_SPACES = Pattern.compile(" {2,}");

    /**
     * {@link #normalize} 用：表头比对不关心任何空格，半角/全角/不间断/零宽类空格一律去掉。
     *
     * <p>PDF 抽出来的单元格文字经常因换行多出空格，还可能带 {@code \u00a0}；Word 里也常用全角空格。
     * 只去半角空格的话，{@code 客户 名称} 就匹配不上注册的 {@code 客户名称}。
     * 样例三份文档的表头本身没有这类空格（docx 与 pdf 的 tables 逐字相同），这条规则的覆盖在
     * {@code ResultsTest.normalizeIgnoresCaseAndEveryKindOfSpace}。
     */
    private static final Pattern ALL_SPACES = Pattern.compile("[\\s\u3000\u00a0\u2007\u202f]+");

    /**
     * 已经成功注册过许可证的产品名，见 {@link #applyLicense}。
     */
    private static final Set<String> LICENSED = new HashSet<String>();

    /**
     * {@link #isEvaluationNotice} 用：评估水印去掉空格、转小写之后的特征片段。
     */
    private static final String[] EVALUATION_MARKS = {
            "evaluationonly",                    // Evaluation Only.
            "createdwithaspose",                 // Created with Aspose.Words / Aspose.PDF for Java ...
            "createdwithanevaluationcopy",       // Created with an evaluation copy of Aspose ...
            "todiscoverthefullversions",         // To discover the full versions of our APIs ...
            "ptyltd",                            // Copyright 2001-2023 Aspose Pty Ltd.（含被换行切剩的孤行）
            "onlydeliverablewithalicense"        // 部分产品页脚水印的尾巴
    };

    private Results() {
    }

    /**
     * 序列化成缩进 JSON。
     */
    static String toJson(Object value) throws Exception {
        return MAPPER.writeValueAsString(value);
    }

    /**
     * 把 {@link WordParser.ImageMode#BASE64} 结果里的图片落成文件。
     *
     * <p>文件名取自结果 Map，所以既要做 {@link #safeName} 消毒，也要做 {@link Images#uniqueFile} 查重，
     * 免得调用方传进来的名字覆盖已有文件。
     *
     * <p>只有 {@code BASE64} 模式的结果里带 {@code base64} 字段，其他模式（含默认的 {@code META}）
     * 这里无内容可写，会在标准错误提示一句而不是安静地什么都不做。
     */
    static void writeImages(Map<String, Object> result, File directory) throws Exception {
        Object list = result.get("images");
        if (!(list instanceof List)) {
            return;
        }
        int written = 0;
        for (Object item : (List<?>) list) {
            Map<?, ?> image = (Map<?, ?>) item;
            String base64 = (String) image.get("base64");
            if (base64 == null) {
                continue;
            }
            byte[] bytes = Base64.getDecoder().decode(base64.substring(base64.indexOf(',') + 1));
            Images.write(Images.uniqueFile(directory, String.valueOf(image.get("fileName"))), bytes);
            written++;
        }
        if (written == 0 && !((List<?>) list).isEmpty()) {
            System.err.println("[writeImages] 结果里的图片没有 base64 内容，一张都没导出 —— "
                    + "解析时用 WordParser.ImageMode.BASE64（默认是 META，只给元信息），"
                    + "或者改用 ImageMode.FILE 配合 imageDir(File) 边解析边落盘");
        }
    }

    /**
     * 只取文件名部分，丢掉调用方传进来的任何目录成分。
     *
     * <p>{@link #writeImages} 的文件名来自结果 Map，而这个 Map 完全可能是从外部（数据库、前端、消息）反序列化
     * 回来的；不加这道工序，{@code fileName} 写成 {@code ../../} 开头的路径就能把文件写到目录之外。
     */
    static String safeName(String fileName) {
        if (fileName == null || fileName.length() == 0) {
            return "image.bin";
        }
        int slash = Math.max(fileName.lastIndexOf('/'), fileName.lastIndexOf('\\'));
        String name = slash < 0 ? fileName : fileName.substring(slash + 1);
        return name.length() == 0 || ".".equals(name) || "..".equals(name) ? "image.bin" : name;
    }

    /**
     * 找许可证文件：先看 classpath 根目录，再看进程工作目录。
     *
     * <p>Aspose 各产品（words/pdf/slides）许可证类互不通用，但文件查找方式一样，所以只抽这一半；
     * 拿到流之后由各解析器调用自己产品的 {@code License.setLicense}。
     *
     * @return 找不到时返回 null，调用方按评估模式继续
     */
    static InputStream licenseStream(String classpathResource, String fileName) {
        InputStream stream = Results.class.getResourceAsStream(classpathResource);
        if (stream != null) {
            return stream;
        }
        try {
            File file = new File(fileName);
            return file.isFile() ? new FileInputStream(file) : null;
        } catch (Exception ignored) {
            return null;
        }
    }

    /**
     * 由各解析器实现：用自家产品的 {@code License} 注册这个流。
     */
    interface LicenseApplier {
        void setLicense(InputStream stream) throws Exception;
    }

    /**
     * 注册许可证，一个产品只注册一次。
     *
     * <p>Aspose 的许可证是进程级全局状态，原来每调一次 {@code parse} 就重新读一遍 lic 文件；批量解析几百个
     * 文档时这些 IO 全是白费的。另外原来注册失败被静默 {@code catch} 掉，lic 过期或写错时现象是「输出莫名其妙
     * 少了内容」，排查半天，所以现在失败要提示一行。
     *
     * @param product 产品名（aspose-words / aspose-pdf / aspose-slides），用于去重和提示信息
     */
    static synchronized void applyLicense(String product, String classpathResource, String fileName,
                                          LicenseApplier applier) {
        if (LICENSED.contains(product)) {
            return;
        }
        InputStream stream = licenseStream(classpathResource, fileName);
        if (stream == null) {
            return;                                         // 没有 lic 文件：正常按评估模式走
        }
        try {
            applier.setLicense(stream);
            LICENSED.add(product);
        } catch (Exception e) {
            System.err.println("[" + product + "] 许可证加载失败，按评估模式解析（内容可能被截断）：" + e.getMessage());
        } finally {
            close(stream);
        }
    }

    /**
     * 关掉流，失败当作没发生：许可证文件已经读完，关不掉也不能让解析挂掉。
     */
    static void close(InputStream stream) {
        if (stream == null) {
            return;
        }
        try {
            stream.close();
        } catch (Exception ignored) {
            // 忽略关闭异常
        }
    }

    /**
     * 清理 Aspose 节点文本里的结构标记。
     *
     * <h3>处理规则</h3>
     * <ul>
     *   <li>换行类（{@code \r}、{@code \n}、软换行 {@code \u000b}、分页 {@code \u000c}）与不换行空格
     *       {@code \u00a0} 统一成空格。软换行必须算「分隔」而不是「丢弃」：Shift+Enter 把一段话切成两行时，
     *       丢了这两个字符就变成「第一行第二行」粘连，内容看着还在、实际读不出来；不换行空格换成普通空格
     *       是为了后面「连续空格压成一个」和表头归一化能一视同仁（{@code \u00a0} 在 Word 里看着就是空格）；</li>
     *   <li>域代码丢弃、域结果保留（见下）；</li>
     *   <li>{@code \u0007}（单元格结束）等其它控制字符丢弃；</li>
     *   <li>连续空格压成一个，最后 trim。</li>
     * </ul>
     *
     * <h3>为什么要单独处理 Word 的域</h3>
     * {@code Node.getText()} 把域写成三段：{@code \u0013 域代码 \u0014 域结果 \u0015}。原样输出时，
     * 「本页码是 3」这种句子会变成 {@code 前文 PAGE 后文}，落在表格单元格里更糟 —— {@code PAGE} 会被当成
     * 该列的表头。所以这里按状态机走一遍：{@code \u0013}～{@code \u0014} 之间（域代码）不输出，
     * {@code \u0014}～{@code \u0015} 之间（域结果）照常输出。域可以嵌套（{@code TOC} 里套 {@code PAGE}），
     * 用两个计数器记当前层数，内层域整体落在外层的代码区时仍然一起丢掉。
     *
     * <p>域结果没算过时就是空的（新建/转换出来的文档常见），此时这段文字整块消失，只剩前后文 —— 这是对的：
     * 那是 Word 的排版指令，不是文档内容。想要真实结果得先 {@code Document.updateFields()}，本模块不做版面计算。
     */
    static String clean(String raw) {
        if (raw == null) {
            return "";
        }
        StringBuilder sb = new StringBuilder(raw.length());
        int fieldLevel = 0;       // 当前套在第几层域里
        int codeLevel = 0;        // 其中还在「代码区」的层数，大于 0 表示当前字符要丢弃
        for (int i = 0; i < raw.length(); i++) {
            char ch = raw.charAt(i);
            if (ch == 0x13) {                                     // 域开始
                fieldLevel++;
                codeLevel++;
                continue;
            }
            if (ch == 0x14) {                                     // 域代码 / 域结果的分界
                if (codeLevel > 0) {
                    codeLevel--;
                }
                continue;
            }
            if (ch == 0x15) {                                     // 域结束，连带退出它所在的代码区
                if (fieldLevel > 0) {
                    fieldLevel--;
                }
                codeLevel = Math.min(codeLevel, fieldLevel);
                continue;
            }
            if (codeLevel > 0) {
                continue;                                          // 域代码本身不是内容
            }
            if (ch == '\r' || ch == '\n' || ch == '\u00a0' || ch == '\u000b' || ch == '\u000c') {
                sb.append(' ');
            } else if (ch >= 0x20) {                              // \u0007（单元格结束）等控制字符全部丢弃
                sb.append(ch);
            }
        }
        // 一次正则压掉连续空格。原来用 while (contains("  ")) + replace，一串长空格就要把整个文本重抄十几遍，
        // 实测 300 万空格的单元格：旧写法 95ms，这里 10ms
        return TEXT_SPACES.matcher(sb).replaceAll(" ").trim();
    }

    /**
     * 判断这段文字是不是 Aspose 评估水印。
     *
     * <p>水印句子会被排版的换行切成好几段（PDF 尤其明显：{@code Evaluation Only. Created with Aspose.Words.
     * Copyright 2001-2023 Aspose Pty Ltd.} 会剩一个 {@code Pty Ltd.} 的孤行），所以不能只比对行首，
     * 这里去掉所有空格、转小写之后比对若干特征片段，跨行断开的残句也能认出来。
     *
     * <p>代价：正文里真的写到 "Aspose Pty Ltd" 这类字样时会被当成水印丢掉。评估水印只出现在没有 lic 的环境，
     * 生产接入 lic 后这段过滤不触发任何内容。
     */
    static boolean isEvaluationNotice(String text) {
        if (text == null || text.length() == 0) {
            return false;
        }
        String value = ALL_SPACES.matcher(text).replaceAll("").toLowerCase(Locale.ROOT);
        for (String mark : EVALUATION_MARKS) {
            if (value.contains(mark)) {
                return true;
            }
        }
        return false;
    }

    /**
     * {@link #clean} + {@link #isEvaluationNotice}：整段是评估水印时返回空串，调用方按空文本跳过即可。
     */
    static String cleanContent(String raw) {
        String clean = clean(raw);
        return isEvaluationNotice(clean) ? "" : clean;
    }

    /**
     * 表头比对用的归一化：忽略大小写与一切空格，让 {@code "列 1"}、{@code "列　1"}、{@code "列1"} 等价。
     *
     * <p>用 {@link Locale#ROOT} 而不是默认 Locale：土耳其语环境下 {@code toLowerCase()} 会把 {@code I} 变成
     * {@code ı}，注册表头里的 {@code ItemId} 就再也匹配不上文档里的 {@code itemId} 了。
     */
    static String normalize(String value) {
        if (value == null) {
            return "";
        }
        return ALL_SPACES.matcher(value).replaceAll("").toLowerCase(Locale.ROOT);
    }

    static String join(List<String> values, String separator) {
        StringBuilder sb = new StringBuilder();
        for (int i = 0; i < values.size(); i++) {
            if (i > 0) {
                sb.append(separator);
            }
            sb.append(values.get(i));
        }
        return sb.toString();
    }
}
