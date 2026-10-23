package cn.net.tianxi.document;

import com.aspose.slides.*;

import java.io.File;
import java.io.InputStream;
import java.util.*;

/**
 * ppt/pptx 解析工具类：一次调用返回三个集合 —— 文本、表格集合、图片集合，返回结构与 {@link WordParser} 完全一致。
 *
 * <pre>
 * Map&lt;String, Object&gt; result = PptParser.parse("D:/a.pptx", "片区");
 * Map&lt;String, Object&gt; result = PptParser.parse("D:/a.pptx", options);   // 与 WordParser 同一个 Options
 * String json = PptParser.toJson(result);
 * </pre>
 *
 * <p>下面的说明与数字全部来自仓库里的样例 {@code src/test/resources/ppt/sales-report.pptx}
 * （同一份「2025 年度华东区销售业绩报告」的 3 页幻灯片，手工拼的 OOXML：合并单元格属性写在 {@code <a:tc>} 上），
 * 环境是 JDK 8、<b>没有 lic</b>（评估模式）。
 *
 * <h2>样例实测基线</h2>
 * {@code text} 是 9 行文本框文字，按形状在页面里出现的顺序：
 * {@code 2025年 / 销售业绩 / 华东大区 / 回款率 / 86% / 达标 / 数据来源 / 销售系统 / 财务对账}
 * （最后三个在一个<b>组合形状</b>里，不递归展开就整组丢掉）；{@code tables} 6 个表头组共 22 行；
 * {@code images} 1 张 PNG（1024×768、670939 字节、{@code contentId} = {@code 487e915db2c53907e8c1551aeaffea235423210c}，
 * 与 docx 样例里那张是同一份字节，指纹因此相同）。默认 {@code META} 模式整串 JSON 3526 个字符，
 * {@code BASE64} 模式 898156 个字符。每组的 {@code source} 里 {@code page} 是<b>幻灯片序号</b>（样例实测
 * 第 1／2／3 页各出现，见 {@link Tables#markSource}）；幻灯片不跨页，所以不会有 {@code lastPage}，
 * 不要这个键时 {@code Options.source(false)} 同样把它去掉（三种格式一起生效）。
 * 换选项：注册 5 组真表头 6 组／21 行、{@code DROP} 5 组／19 行、
 * {@code FIRST} 6 组／22 行。注册时给三组起名（{@code Options.name(String)}）是 6 组／21 行／3466 个字符，
 * 带 {@code name} 的是第 1、2、4、5 组（台账在幻灯片网格里出现的也是合并后那一行，所以命中）；核心表与 docx／pdf <b>逐格相同</b>（列名
 * {@code ["片区","区域","销售额","回款率","销售额-5","回款率-6","备注"]}，4 行数据一字不差）。
 * <i>（测试：{@code PptParserTest}、{@code PptParserTextTest}、{@code CrossFormatTest}）</i>
 *
 * <h2>形状怎么分发</h2>
 * 一页幻灯片就是一段文档流，形状按幻灯片上的顺序依次进入结果 ——
 * <ol>
 *   <li>{@link ITable} 读成网格：每个单元格用 {@code getFirstRowIndex/getFirstColumnIndex} 找到自己合并的起点，
 *       起点格的文本按 {@code rowSpan}/{@code colSpan} 铺满它覆盖的所有格子，与 {@link WordParser} 展开
 *       {@code gridSpan}/{@code vMerge} 之后的网格一致；整行只余一个「不是横向合并延续格」的格子、且这一行
 *       不是全空时即跨整行的合并行（小标题、分组行），去向由 {@link WordParser.Options#fullRow(WordParser.FullRow)} 决定。
 *       样例目标表的 {@code 2025年}
 *       出现在任何表头之前，于是另起一组、列名补成 {@code ["列1","列2","列3","列4"]}；</li>
 *   <li>{@link IAutoShape}（文本框、形状里的文字）按段落进 {@code text}；</li>
 *   <li>{@link IPictureFrame} 的 {@link IPPImage} 二进制进 {@code images}；</li>
 *   <li>{@link IGroupShape} 组合形状递归展开（上限 {@value #MAX_GROUP_DEPTH} 层），
 *       里面的表格、文本、图片按同样规则处理。</li>
 * </ol>
 *
 * <p>这四条分支互不重叠：实测八个形状类逐个跑 {@code isAssignableFrom}，只有 {@code AutoShape} 是
 * {@link IAutoShape}，{@code Table} 只实现 {@code ITable}、{@code GroupShape} 只实现 {@code IGroupShape}、
 * {@code PictureFrame} 与 {@code VideoFrame} 只实现 {@code IPictureFrame}，所以判断顺序不影响结果
 * （仍把最宽的 {@link IAutoShape} 留到最后兜底）。{@code Chart}、{@code OleObjectFrame}、{@code SmartArtShape}
 * 一个接口都不沾，整形状被静默跳过 —— 这就是「图表数据／OLE 对象内容／SmartArt 文字读不到」的由来。
 *
 * <h2>两行表头：建议注册表头</h2>
 * 幻灯片没有「重复标题行」这种标记，表格只有样式上的首行标题（{@code getFirstRow()}），它只是弱线索。
 * 样例台账在 Word 里是两行表头（上层 {@code 回款情况} 跨两列），静默模式下层表头行被当成一行数据：
 * 列名成了 {@code ["客户名称","回款情况","回款情况-3","回款率","备注"]}、8 行（第一行就是下层表头）。
 * 注册 {@code headers("客户名称","合同金额","已回款","回款率","备注")} 之后那一行命中注册表头、
 * 与上一行连续表头合并，列名正确、7 行数据（最后一行是 {@code ["合计","1586","1384","87%",""]}）。
 * 这就是「PPT 建议注册表头」的理由，PDF 同理。
 *
 * <h2>评估模式</h2>
 * 没有 {@code aspose-slides.lic} 时 Aspose.Slides 读文本框<b>只给第一段</b>、且截到 5 个字符再追加
 * {@value #TRUNCATION_NOTICE} 提示；表格单元格走的也是文本框那一条，所以单元格同样被截，
 * {@link #stripEvaluationNotice} 会把提示去掉（被截掉的内容回不来）。
 * 样例因此把每处文字都写成一段、不超过 5 个字符，有没有 lic 读到的都是同一串
 * （<i>{@code PptParserTest.everyTextFrameInTheSampleStaysWithinTheEvaluationLimit} 守着这条</i>）；
 * 代价是这份样例验证不了「长文本能不能完整读出来」，那一条由 docx 样例覆盖（它的段落有 60 多个字符，读出来是完整的）。
 * 幻灯片本身若是用 Aspose 之类导出的，页面上还会带着 {@code Evaluation Only. Created with ...} 的水印文字，
 * 这部分和 Word/PDF 一样由 {@link Results#cleanContent} 统一过滤。
 * 放一个 {@code aspose-slides.lic} 到 classpath 或当前目录即可解除限制。
 *
 * <p><b>读不到的部分</b>：备注页／母版／版式页的文字（只读正片 {@code getSlides()}）、SmartArt 图形里的文字、
 * 图表数据、嵌入的 OLE 对象（Excel 工作表、Word 文档）里的内容、形状填充的底图。这些都在形状文字之外，
 * 需要时按 {@link #readShapes} 的分发方式各加一支。视频是个例外：{@code VideoFrame} 继承 {@code PictureFrame}，
 * 走的是图片那一支，进 {@code images} 的是它的封面图，视频内容本身不读。
 *
 * <p>本类只有解析相关的静态方法，没有命令行入口；要单独验证一个文档，直接在测试类里调
 * {@link #parse(String, String...)} 或 {@link #parse(String, WordParser.Options)}。
 */
public final class PptParser {

    /**
     * Aspose.Slides 评估模式截断文本时留下的提示，不是内容。
     */
    private static final String TRUNCATION_NOTICE = "... text has been truncated due to evaluation version limitation.";

    /**
     * 组合形状最多递归展开的层数：正常文件只有几层，深的多半是损坏或故意构造出来撑爆栈的。
     */
    private static final int MAX_GROUP_DEPTH = 32;

    private PptParser() {
    }

    /**
     * 按文件路径解析 ppt/pptx。
     *
     * @param headerKeywords 可选的表头关键字，含义同 {@link WordParser#parse(String, String...)}
     * @return {text 全文, tables 表格集合, images 图片集合}
     */
    public static Map<String, Object> parse(String filePath, String... headerKeywords) throws Exception {
        return parse(open(filePath), WordParser.options().keywords(headerKeywords));
    }

    /**
     * 从输入流解析，语义同 {@link #parse(String, String...)}。
     */
    public static Map<String, Object> parse(InputStream inputStream, String... headerKeywords) throws Exception {
        return parse(open(inputStream), WordParser.options().keywords(headerKeywords));
    }

    /**
     * 按选项解析，可注册完整表头集合、整行合并行的去向与图片输出方式，同 {@link WordParser#parse(String, WordParser.Options)}。
     */
    public static Map<String, Object> parse(String filePath, WordParser.Options options) throws Exception {
        return parse(open(filePath), options);
    }

    /**
     * 从输入流按选项解析，语义同 {@link #parse(String, WordParser.Options)}；流本身由调用方关闭。
     */
    public static Map<String, Object> parse(InputStream inputStream, WordParser.Options options) throws Exception {
        return parse(open(inputStream), options);
    }

    /* ---------------- 解析主流程 ---------------- */

    private static Map<String, Object> parse(Presentation presentation, WordParser.Options options) throws Exception {
        try {
            return read(presentation, options);
        } finally {
            presentation.dispose();          // 演示文稿持有文件与内存中的大图，不释放会一直占着
        }
    }

    /**
     * 逐页幻灯片按形状顺序读，表格／文本／图片各进各的集合。
     *
     * <h3>为什么不用排序</h3>
     * 幻灯片里 {@code getShapes()} 的顺序就是绘制顺序，也就是人读的顺序，不像 PDF 需要按坐标重排；
     * 于是 {@code tables} 和 {@code images} 的编号天然按文档流走。
     *
     * <h3>调用策略</h3>
     * 只读 {@code getSlides()}（正片）。备注页、母版、版式页里的文字属于版式附属内容，不进结果；
     * 一张图的 {@code images} 编号顺序与它在幻灯片上出现的先后一致，跨页重复引用同一张图的字节由
     * {@link Images} 用内容指纹标 {@code duplicateOf}。
     */
    private static Map<String, Object> read(Presentation presentation, WordParser.Options options) throws Exception {
        Images images = new Images(options);
        List<String> paragraphs = new ArrayList<String>();
        List<Map<String, Object>> tables = new ArrayList<Map<String, Object>>();
        Set<Integer> matchedGroups = new LinkedHashSet<Integer>();   // 实际命中过的表头集合下标

        for (int s = 0; s < presentation.getSlides().size(); s++) {
            readShapes(presentation.getSlides().get_Item(s).getShapes(), options, paragraphs, tables, images,
                    matchedGroups, s + 1, 0);
        }
        List<Map<String, Object>> unmatched = Tables.warnUnmatchedHeaders("[PptParser] ", options, matchedGroups);

        Map<String, Object> result = new LinkedHashMap<String, Object>();
        result.put("text", Results.join(paragraphs, "\n"));
        result.put("tables", tables);
        result.put("images", images.list());
        if (options.writesUnmatchedHeaders()) {
            result.put("unmatchedHeaders", unmatched);
        }
        return result;
    }

    /**
     * 按幻灯片上的形状顺序读取：表格进 {@code tables}，文本进 {@code text}，图片进 {@code images}。
     *
     * <h3>四类形状的处理</h3>
     * <ul>
     *   <li>{@link ITable} —— 交给 {@link #parseTable} 拆合并、分组；</li>
     *   <li>{@link IPictureFrame} —— 位图，交给 {@link #addImage}；</li>
     *   <li>{@link IGroupShape} —— 组合形状，递归展开，里面的表格／文本／图片按同样规则处理；</li>
     *   <li>{@link IAutoShape} —— 文本框、占位符、几何形状的文字都在这里（Aspose.Slides 里带
     *       {@code getTextFrame()} 的形状接口就是它）；文本框可能为 null（占位符被删过之类），跳过而不是报错。</li>
     * </ul>
     *
     * <h3>判断顺序</h3>
     * 实测这四个接口在 Aspose.Slides 23.12 里互不兼容：{@code Table} 只实现 {@code ITable}、
     * {@code GroupShape} 只实现 {@code IGroupShape}、{@code PictureFrame} 只实现 {@code IPictureFrame}，
     * 都不是 {@code IAutoShape}，所以四支谁先谁后不影响结果。仍然把最宽的 {@code IAutoShape} 放在最后兜底：
     * Aspose 哪天调整了层次（例如让表格也实现 {@code IAutoShape}），先判它会把表格文字当普通文本读走。
     *
     * <h3>边界</h3>
     * 组合嵌套超过 {@value #MAX_GROUP_DEPTH} 层时只跳过<b>这一组</b>，同层其余形状照常读；
     * {@code Chart}、{@code OleObjectFrame}、{@code SmartArtShape} 三个类既不实现上面任何一个接口、
     * 也不是 {@code IAutoShape}，四条分支全不认，整个形状被静默跳过（图表数据、OLE 对象内容、SmartArt 文字
     * 都读不到）；形状填充底图同理，属于已知范围外。
     */
    private static void readShapes(IShapeCollection shapes, WordParser.Options options,
                                   List<String> paragraphs, List<Map<String, Object>> tables,
                                   Images images, Set<Integer> matchedGroups, int slide, int depth) throws Exception {
        for (int i = 0; i < shapes.size(); i++) {
            IShape shape = shapes.get_Item(i);
            if (shape instanceof ITable) {
                List<Map<String, Object>> groups = parseTable((ITable) shape, options, matchedGroups);
                // 一张表只在一页幻灯片上（pptx 没有「跨页表」这回事），起止页写同一个序号
                if (options.writesSource()) {
                    Tables.markSource(groups, slide, slide);
                }
                tables.addAll(groups);
            } else if (shape instanceof IPictureFrame) {
                addImage(images, (IPictureFrame) shape);
            } else if (shape instanceof IGroupShape) {
                if (depth >= MAX_GROUP_DEPTH) {
                    // 只跳过这一组，同一层后面的形状照常读；整层 return 会让一页的其余内容凭空消失
                    System.err.println("[PptParser] 组合形状嵌套超过 " + MAX_GROUP_DEPTH + " 层，该组已跳过");
                    continue;
                }
                readShapes(((IGroupShape) shape).getShapes(), options, paragraphs, tables, images,
                        matchedGroups, slide, depth + 1);
            } else if (shape instanceof IAutoShape) {
                ITextFrame frame = ((IAutoShape) shape).getTextFrame();
                if (frame != null) {                     // 占位符被删过之类的形状可能没有文本框
                    addText(paragraphs, frame.getText());
                }
            }
        }
    }

    /* ---------------- 文本 ---------------- */

    /**
     * 一个文本框按段落拆成多行，行内空白统一清洗，与 {@link WordParser} 的段落文本一致。
     */
    private static void addText(List<String> paragraphs, String text) {
        paragraphs.addAll(textLines(text));
    }

    /**
     * 把 {@link ITextFrame#getText()} 的整段文字拆成清洗过的非空行。
     *
     * <p>Aspose.Slides 用 {@code \r} 分隔段落（Windows 换行则是 {@code \r\n}），不能只按 {@code \n} 拆：
     * {@link Results#clean} 会把 {@code \r} 转成空格，多段文本框就会被压成一行。三种换行都认。
     *
     * <p>包级可见只为单测：评估模式一个文本框只给第一段，端到端测试照不到多段的情况。
     */
    static List<String> textLines(String text) {
        List<String> lines = new ArrayList<String>();
        if (text == null) {
            return lines;
        }
        for (String line : text.split("\r\n|\r|\n")) {
            String clean = stripEvaluationNotice(Results.cleanContent(line));
            if (clean.length() > 0) {
                lines.add(clean);
            }
        }
        return lines;
    }

    /* ---------------- 表格：一个表头对应它下面的行 ---------------- */

    /**
     * 读一张表成网格，再交给 {@link Tables#group} 按表头分组。
     *
     * <h3>解析过程</h3>
     * <ol>
     *   <li>逐格用 {@code getFirstRowIndex/getFirstColumnIndex} 找到它合并的起点，不是起点的格子跳过；</li>
     *   <li>起点格的文本按 {@code rowSpan}/{@code colSpan} 铺满它覆盖的所有格子，得到与 {@link WordParser}
     *       展开 {@code gridSpan}/{@code vMerge} 之后一致的网格；</li>
     *   <li>统计每行有几个<b>不是横向合并延续格</b>的格子（{@code getFirstColumnIndex() == c}）：整行只剩一个、
     *       且这一行不是全空，才算跨整行的合并行（小标题、分组行）。空白合并行只是排版间隔，当成分组行的话
     *       紧跟它的那一行会被当成表头 —— 实测两列表中间插一行空白合并行时一张表被切成两组、{@code 丙丁}
     *       那一行变成了列名（{@code PptParserTest.blankFullWidthMergedRowIsNotASeparator}）；</li>
     *   <li>标题行线索用 {@code getFirstRow()}（样式上的首行标题），它只是弱线索，跨多组表头时仍建议注册完整表头。</li>
     * </ol>
     *
     * <h3>为什么纵向合并的延续格也要算进「这一行自己的格子」</h3>
     * 判据只排除<b>横向</b>延续格，与 {@link WordParser} 的 {@code Grid.isSeparator}、{@link PdfParser} 的
     * {@code toGrid} 是同一条：某一列纵向合并到上一行时（大分类列、序号列很常见），这一行其余格子仍是自己的数据。
     * 早先这里数的是「起点格」（{@code getFirstRowIndex() == r} 且 {@code getFirstColumnIndex() == c}），
     * 于是「A 列纵向合并跨两行 + B 列每行各有值」的两列表里，第二行只剩 B 一个起点格、被当成跨整行的分组行：
     * 实测 {@code KEEP} 把 {@code 值二} 抹成空串（{@code [{列A:第一,列B:值一},{列A:第一,列B:}]}）、
     * {@code DROP} 整行删掉，同一份版面的 docx 两种模式都读到 {@code 值二}。改成只排除横向延续格后
     * 两种模式都与 docx 一致（{@code PptParserTest.verticallyMergedDataRowIsNotAMistakenSeparator}）。
     *
     * <h3>两处防御</h3>
     * {@code isMergedCell()} 对起始格同样返回 true，不能用来区分，所以这里改用「起点下标等于自己」判断；
     * 非常规单元格（老文件里的图片格之类）可能没有文本框，取到 null 就当空格处理，不让整篇解析失败。
     */
    private static List<Map<String, Object>> parseTable(ITable table, WordParser.Options options,
                                                        Set<Integer> matchedGroups) {
        int rowCount = table.getRows().size();
        int columnCount = table.getColumns().size();
        String[][] values = new String[rowCount][columnCount];
        for (String[] row : values) {
            Arrays.fill(row, "");
        }
        int[] origins = new int[rowCount];          // 每行「不是横向合并延续格」的格子数

        for (int r = 0; r < rowCount; r++) {
            for (int c = 0; c < columnCount; c++) {
                ICell cell = table.getRows().get_Item(r).get_Item(c);
                if (cell.getFirstColumnIndex() == c) {
                    origins[r]++;                   // 纵向合并的延续格也算：见下面 separator 的说明
                }
                if (cell.getFirstRowIndex() != r || cell.getFirstColumnIndex() != c) {
                    continue;                       // 被别的格合并覆盖，内容在起始格里
                }
                ITextFrame cellFrame = cell.getTextFrame();      // 老文件里的非常规单元格可能没有文本框
                String text = cellFrame == null ? "" : stripEvaluationNotice(Results.cleanContent(cellFrame.getText()));
                int spannedRows = Math.max(1, cell.getRowSpan());
                int spannedColumns = Math.max(1, cell.getColSpan());
                for (int rr = r; rr < Math.min(rowCount, r + spannedRows); rr++) {
                    for (int cc = c; cc < Math.min(columnCount, c + spannedColumns); cc++) {
                        values[rr][cc] = text;
                    }
                }
            }
        }

        boolean[] separator = new boolean[rowCount];
        for (int r = 0; r < rowCount; r++) {
            boolean blank = true;
            for (int c = 0; c < columnCount; c++) {
                blank &= values[r][c].length() == 0;
            }
            // 整行只剩一个「不是横向合并延续格」的格子、又不是全空 —— 跨整行的合并行（小标题、分组行）
            separator[r] = columnCount > 1 && origins[r] == 1 && !blank;
        }
        return Tables.group(values, separator, null, options, matchedGroups, table.getFirstRow());
    }

    /* ---------------- 图片 ---------------- */

    /**
     * 记录一张图片。
     *
     * <p>{@code getLinkPathLong()} 有值且二进制为空时是「链接到文件」的图片，字节不在 pptx 包里，只给路径。
     */
    private static void addImage(Images images, IPictureFrame frame) throws Exception {
        IPictureFillFormat format = frame.getPictureFormat();
        ISlidesPicture picture = format == null ? null : format.getPicture();
        IPPImage image = picture == null ? null : picture.getImage();
        if (image == null) {
            return;
        }
        byte[] bytes = image.getBinaryData();
        String link = picture.getLinkPathLong();
        if ((bytes == null || bytes.length == 0) && link != null && link.length() > 0) {
            images.link(image.getContentType(), image.getWidth(), image.getHeight(), link);
            return;
        }
        images.add(image.getContentType(), image.getWidth(), image.getHeight(), bytes);
    }

    /* ---------------- 输出 ---------------- */

    /**
     * 序列化成缩进 JSON，同 {@link WordParser#toJson(Object)}。
     */
    public static String toJson(Object value) throws Exception {
        return Results.toJson(value);
    }

    /**
     * 把 {@link WordParser.ImageMode#BASE64} 结果里的图片落成文件，同 {@link WordParser#writeImages}。
     */
    public static void writeImages(Map<String, Object> result, File directory) throws Exception {
        Results.writeImages(result, directory);
    }

    /* ---------------- 通用 ---------------- */

    /**
     * classpath 或当前目录存在 aspose-slides.lic 时自动注册（一个进程只需一次），避免评估模式截断文本。
     */
    private static void applyLicenseIfPresent() {
        Results.applyLicense("aspose-slides", "/aspose-slides.lic", "aspose-slides.lic",
                stream -> new License().setLicense(stream));
    }

    /**
     * 先注册许可证、再加载演示文稿。
     *
     * <p>顺序不能反：Aspose.Slides 在构造 {@code Presentation} 时就按评估模式给每段文字打上截断标记，
     * lic 晚注册对已加载的文件无效（{@link #stripEvaluationNotice} 只能去掉标记，被截掉的内容回不来）。
     */
    private static Presentation open(String filePath) throws Exception {
        applyLicenseIfPresent();
        return new Presentation(filePath);
    }

    /**
     * 同上，从输入流加载；{@link #parse} 会负责 {@code dispose}，传入的流本身由调用方关闭。
     */
    private static Presentation open(InputStream inputStream) throws Exception {
        applyLicenseIfPresent();
        return new Presentation(inputStream);
    }

    /**
     * 去掉评估模式追加在文本末尾的截断提示；剩下的内容若是水印则当成空。
     */
    private static String stripEvaluationNotice(String text) {
        int index = text.indexOf(TRUNCATION_NOTICE);
        return index < 0 ? text : Results.cleanContent(text.substring(0, index));
    }
}
