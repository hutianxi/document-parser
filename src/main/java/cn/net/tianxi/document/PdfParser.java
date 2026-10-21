package cn.net.tianxi.document;

import com.aspose.pdf.AbsorbedCell;
import com.aspose.pdf.AbsorbedRow;
import com.aspose.pdf.AbsorbedTable;
import com.aspose.pdf.Document;
import com.aspose.pdf.ImagePlacement;
import com.aspose.pdf.ImagePlacementAbsorber;
import com.aspose.pdf.License;
import com.aspose.pdf.MarkupParagraph;
import com.aspose.pdf.MarkupSection;
import com.aspose.pdf.Page;
import com.aspose.pdf.PageCollection;
import com.aspose.pdf.PageMarkup;
import com.aspose.pdf.ParagraphAbsorber;
import com.aspose.pdf.Rectangle;
import com.aspose.pdf.TableAbsorber;
import com.aspose.pdf.TextFragment;
import com.aspose.pdf.XImage;

import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Collections;
import java.util.Comparator;
import java.util.LinkedHashMap;
import java.util.LinkedHashSet;
import java.util.List;
import java.util.Map;
import java.util.Set;

/**
 * pdf 解析工具类：一次调用返回三个集合 —— 文本、表格集合、图片集合，返回结构与 {@link WordParser} 完全一致。
 *
 * <pre>
 * Map&lt;String, Object&gt; result = PdfParser.parse("D:/a.pdf", "客户");
 * Map&lt;String, Object&gt; result = PdfParser.parse("D:/a.pdf", options);   // 与 WordParser 同一个 Options
 * String json = PdfParser.toJson(result);
 * </pre>
 *
 * <p>下面的说明与数字全部来自仓库里的样例 {@code src/test/resources/pdf/sales-report.pdf}
 * （同一份「2025 年度华东区销售业绩报告」用<b>真 Microsoft Word 2021</b> 导出的 3 页 PDF，
 * 文件里的 {@code Producer} 就是它，不是 Aspose 转的，所以文件本身不带评估水印），
 * 环境是 JDK 8、<b>没有 lic</b>（评估模式）。
 *
 * <h2>样例实测基线</h2>
 * {@code text} 是<b>空串</b>（这份 PDF 的正文全部落在表格里，页面之间只有撑开用的空段落；
 * 空段落与评估水印都不进 {@code text}，所以空串是正常输出，不是「什么都没读到」）；
 * {@code tables} 6 个表头组共 45 行；{@code images} 1 张 JPEG（699×524、38668 字节、
 * {@code contentId} = {@code 80c8e65be48c09a0f173845e4dfc072bc541fb68}）。
 * 默认 {@code META} 模式整串 JSON 6186 个字符，{@code BASE64} 模式 57789 个字符。
 * 换选项：注册 5 组真表头后 JSON 与静默模式逐字节相同；{@code DROP} 5 组／43 行；{@code FIRST} 6 组／45 行。
 * <i>（测试：{@code PdfParserTest}）</i>
 *
 * <h2>表格靠几何还原</h2>
 * PDF 里不存在「表格」对象，只有画出来的线段和散落的所有文字，所以读法与 docx 完全不同：
 * <ol>
 *   <li>{@link TableAbsorber} 按线段识别出候选表格，一张表格常常被切成好几块（合并单元格造成的内部线段
 *       缺失会让 Aspose 认为那是两张表），矩形相接的块会先并回一张表，见 {@link #cluster}。
 *       样例核心表的 {@code 片区} 列纵向合并、合并处没有横线，「江苏」「安徽」两行就被切成了独立小块，
 *       聚回来后 4 行都在同一组，列名与值与 Word 读到的<b>逐格相同</b>
 *       （{@code ["片区","区域","销售额","回款率","销售额-5","回款率-6","备注"]}）；
 *       工单表的「{@code 硬件故障 | 12 | 更换配件后复测通过}」那一行同样被切出去又聚回来，仍是完整 4 行；</li>
 *   <li>把所有单元格的左右边界、上下边界各自聚类成列线、行线（容差 {@value #TOLERANCE} pt），网格就出来了；</li>
 *   <li>单元格矩形覆盖到几列几行，就等于横向/纵向合并跨了几格，按 {@link WordParser} 的同一套规则回填值；</li>
 *   <li>整行只被一个单元格覆盖的行，即跨整行的合并行（小标题、分组行），交给
 *       {@link WordParser.Options#fullRow(WordParser.FullRow)} 决定去向。样例目标表的
 *       {@code 2025 年度（已结账）} 出现在任何表头之前，于是另起一组、列名补成
 *       {@code ["列1","列2","列3","列4"]}；</li>
 *   <li>长表格跨页断开时，续接的那一截没有自己的表头行、又正好接在上一页末尾，会并回上一页那张表，
 *       四个判据见 {@link #continues}；调用方没注册表头依据时，「找不出表头行」那条换成「两截都贴着页边界」
 *       （{@value #PAGE_EDGE_REACH} 倍页高以内）这条几何证据顶上，免得两张恰好等列数的独立表格被错并。
 *       跨三页以上的长表由 {@link #mergeSplitTables} 逐页接着并。
 *       样例的台账 31 行在第 1／2 页之间断开：第 1 页画到 {@code 泰州医药城}（第 20 行），
 *       第 2 页从 {@code 湖州建材贸易}（第 21 行）接着画，{@code 合计} 行在最后 —— 并回了一组，
 *       没并的话会变成两组、各带一个表头。</li>
 * </ol>
 *
 * <h2>与 docx 的两处已知差异</h2>
 * <ul>
 *   <li><b>域结果可能丢</b>：导出这份 PDF 用的那份加长版源 docx（台账 30 家客户，源文件不在仓库里）
 *       合计行两格是 {@code SUM(ABOVE)} 的缓存结果（当时用探针核对过 Aspose.Words 读到
 *       {@code 2582.30}/{@code 2226.60}），Word 导出 PDF 时重算了这个域，
 *       PDF 里那两格只剩一个空格，解析如实读出空串 {@code ["合计","","","86%",""]}。要金额请用 docx 那一份
 *       （仓库里的 docx 样例台账是 12 家客户，合计读的是缓存值 {@code ["合计","2096.30","1796.10","86%",""]}）；</li>
 *   <li><b>嵌套表不在了</b>：docx 工单表里那张 {@code 故障部件/工单数} 小表，导出成 PDF 后已经摊平进版面，
 *       所以注册这组表头会在标准错误收到「表头集合未匹配到任何行」的提示（提示只是线索，不是错误）。</li>
 * </ul>
 *
 * <p>另外，PDF 没有 Word 的 {@code w:tblHeader} 标记可依赖，一旦调用方给了表头依据，「首行兜底当表头」就关掉：
 * 只注册台账那一组时，核心表的表头行按数据行处理（那一组 5 行 = 表头行 + 4 行数据）、列名补成 {@code 列N}。
 * 这是有意的 —— 给了依据就以依据为准，免得把一行数据当表头。
 *
 * <h2>正文文本</h2>
 * 用 {@link ParagraphAbsorber} 按段落读出，落在表格矩形内的段落会被跳过 —— 与 Word 的约定一致，
 * 表格里的字只出现在 {@code tables}，不会在 {@code text} 里重复一遍。
 *
 * <h2>评估模式</h2>
 * 没有 {@code aspose-pdf.lic} 时 Aspose.PDF 按评估模式解析。<b>实测这份样例上被限制的是文字层</b>：
 * 段落读不全，但当前样例的正文全落在表格矩形内、{@code text} 是空串，所以这条限制在样例上看不出来
 * （没有实测数字，接入 lic 后请自行核对）；<b>表格与图片没有被截</b> —— 台账 30 行客户 + 合计、
 * 全篇 45 行数据、1 张图片都完整读出，{@code PdfParserTest} 逐字断言了这些行。
 * 页面上可能留着 Aspose.Words 导出的水印文字，由 {@link Results#cleanContent} 统一过滤
 * （被换行切成孤行的那半截也能认出来）。放一个 {@code aspose-pdf.lic} 到 classpath 或当前目录即可解除限制。
 *
 * <p><b>读不到的部分</b>：没有线段边框的表格（只有底色或空白对齐的版面）几何还原不出来，会以普通段落进
 * {@code text}；扫描件／图片型 PDF 没有文字层，需要 OCR，本模块不做。
 *
 * <p>本类只有解析相关的静态方法，没有命令行入口；要单独验证一个文档，直接在测试类里调
 * {@link #parse(String, String...)} 或 {@link #parse(String, WordParser.Options)}。
 */
public final class PdfParser {

    /** 几何聚类容差（pt）：两个边界坐标相差不到这个值就当成同一条列线/行线。 */
    private static final double TOLERANCE = 2.0;

    /**
     * 「贴着页边界」的判定范围，按页高取比例（A4 约 168pt）：够装下一页边距（Word 默认上下 2.54cm ≈ 72pt，
     * 手动放大到 5cm 也还剩余量），又远小于一张完整表落在页面中部时的位置。
     * 用比例而不是固定点数，A5、Letter 这些页高不同的文档才用同一个阈值。
     */
    private static final double PAGE_EDGE_REACH = 0.2;

    private PdfParser() {
    }

    /**
     * 按文件路径解析 pdf。
     *
     * @param headerKeywords 可选的表头关键字，含义同 {@link WordParser#parse(String, String...)}
     * @return {text 全文, tables 表格集合, images 图片集合}
     */
    public static Map<String, Object> parse(String filePath, String... headerKeywords) throws Exception {
        return parse(open(filePath), WordParser.options().keywords(headerKeywords));
    }

    /** 从输入流解析，语义同 {@link #parse(String, String...)}。 */
    public static Map<String, Object> parse(InputStream inputStream, String... headerKeywords) throws Exception {
        return parse(open(inputStream), WordParser.options().keywords(headerKeywords));
    }

    /** 按选项解析，可注册完整表头集合、整行合并行的去向与图片输出方式，同 {@link WordParser#parse(String, WordParser.Options)}。 */
    public static Map<String, Object> parse(String filePath, WordParser.Options options) throws Exception {
        return parse(open(filePath), options);
    }

    /** 从输入流按选项解析，语义同 {@link #parse(String, WordParser.Options)}；解析结束会关闭文档，流本身由调用方关闭。 */
    public static Map<String, Object> parse(InputStream inputStream, WordParser.Options options) throws Exception {
        return parse(open(inputStream), options);
    }

    /* ---------------- 解析主流程 ---------------- */

    private static Map<String, Object> parse(Document document, WordParser.Options options) throws Exception {
        try {
            return read(document, options);
        } finally {
            // Aspose.PDF 持有文件句柄，不关闭在 Windows 上就删不掉、也盖不了原文件
            close(document);
        }
    }

    /**
     * 全流程：先把每一页的表格／段落／图片统一成带位置的 {@link Flow}，再按「页码 + 纵向位置」排一次序，
     * 最后按顺序写进三个集合。
     *
     * <h3>为什么要先收集再排序</h3>
     * PDF 的内容流顺序不等于人读的顺序（先画底图、再画文字很常见），而 {@code text}、{@code tables}、
     * {@code images} 的编号都必须按文档里看到的先后走。所以每一页先各自算出内容项，排序（同一页内按上边界
     * 从大到小，PDF 的 y 轴朝上）之后才交给 {@link Tables#group} 和 {@link Images}。
     *
     * <h3>跨页表格在这一步并回去</h3>
     * 排序之后、分组之前调用 {@link #mergeSplitTables}：一张长表跨页断成两截时，两截各自成网格会各出一个表头组，
     * 分组必须整体做一次。合并的判据见该方法。
     *
     * <h3>调用策略</h3>
     * {@code document} 由 {@link #open(String)} 造出来（许可证先注册），本方法负责在 {@code finally} 里
     * {@code close} —— Aspose.PDF 持有文件句柄，不关闭在 Windows 上就删不掉原文件。
     */
    private static Map<String, Object> read(Document document, WordParser.Options options) throws Exception {
        Images images = new Images(options);
        List<Flow> flows = new ArrayList<Flow>();       // 全文内容，按页码 + 纵向位置排好

        PageCollection pages = document.getPages();
        for (int p = 1; p <= pages.size(); p++) {
            Page page = pages.get_Item(p);
            flows.addAll(pageFlows(page, p));
        }
        Collections.sort(flows, new Comparator<Flow>() {
            @Override
            public int compare(Flow a, Flow b) {
                return a.page != b.page ? Integer.compare(a.page, b.page) : Double.compare(b.top, a.top);
            }
        });

        mergeSplitTables(flows, options);

        List<String> paragraphs = new ArrayList<String>();
        List<Map<String, Object>> tables = new ArrayList<Map<String, Object>>();
        Set<Integer> matchedGroups = new LinkedHashSet<Integer>();   // 实际命中过的表头集合下标
        for (Flow flow : flows) {
            if (flow.grid != null) {
                tables.addAll(Tables.group(flow.grid.values, flow.grid.separator, null,
                        options, matchedGroups, false));
            } else if (flow.text != null) {
                paragraphs.add(flow.text);
            } else if (flow.image != null) {
                flow.image.appendTo(images);
            }
        }
        warnUnmatchedHeaders(options, matchedGroups);

        Map<String, Object> result = new LinkedHashMap<String, Object>();
        result.put("text", Results.join(paragraphs, "\n"));
        result.put("tables", tables);
        result.put("images", images.list());
        return result;
    }

    /** 注册了却一行都没命中的表头集合提示到标准错误输出，规则同 {@link WordParser}。 */
    private static void warnUnmatchedHeaders(WordParser.Options options, Set<Integer> matchedGroups) {
        for (int i = 0; i < options.headerGroupCount(); i++) {
            if (!matchedGroups.contains(i)) {
                System.err.println("[PdfParser] 表头集合未匹配到任何行：" + options.headerGroup(i));
            }
        }
    }

    /**
     * 一页的内容：表格（可能多块）、正文段落、图片，各自带上纵向位置。
     *
     * <h3>处理顺序有依赖</h3>
     * 表格必须先算：正文要靠「段落中心点是否落在表格矩形里」来决定跳过，早于表格就没法判重。
     * 之后 {@link #textFlows} 取段落、{@link #imageFlows} 取图片，三者都只返回带位置的 {@link Flow}，
     * 页与页之间、页内各块之间的顺序统一交给 {@link #read} 排序。
     *
     * <h3>调用策略</h3>
     * 只处理 {@code getPages()} 里的正文页。PDF 里<b>没有线段边框的表格</b>（只有底色或空白对齐）识别不出来 ——
     * 这类版面会被当成普通段落进 {@code text}，是几何还原方式的固有限制。
     */
    private static List<Flow> pageFlows(Page page, int pageNumber) {
        List<Flow> flows = new ArrayList<Flow>();
        List<Rectangle> tableRects = new ArrayList<Rectangle>();   // 正文要跳过的区域
        Rectangle pageRect = page.getRect();
        double pageHeight = pageRect.getURY() - pageRect.getLLY();

        TableAbsorber absorber = new TableAbsorber();
        absorber.visit(page);
        for (List<AbsorbedTable> group : cluster(absorber.getTableList())) {
            Grid grid = toGrid(group);
            if (grid == null) {
                continue;
            }
            for (AbsorbedTable table : group) {
                if (table.getRectangle() != null) {
                    tableRects.add(table.getRectangle());
                }
            }
            Flow flow = new Flow();
            flow.page = pageNumber;
            flow.top = maxURY(grid.bounds);
            flow.left = minLLX(grid.bounds);      // 跨页续接要靠「左边界对齐」判断，取最左的那条竖线
            flow.topGap = (pageRect.getURY() - flow.top) / pageHeight;    // 距页首多远，跨页续接判据用
            flow.bottomGap = (minLLY(grid.bounds) - pageRect.getLLY()) / pageHeight;
            flow.grid = grid;
            flows.add(flow);
        }

        flows.addAll(textFlows(page, tableRects, pageNumber));
        flows.addAll(imageFlows(page, pageNumber));
        return flows;
    }

    /* ---------------- 跨页续接 ---------------- */

    /**
     * 把跨页断开的表格并回一张。
     *
     * <p>PDF 的表格是按页画出来的线段，Word 里一张长表跨页后，下一页那一截通常不再重复表头行，几何上也没有
     * 任何标记说明「它属于上一页那张表」。这里用四个条件判断：相邻的两页、列数相同且左边对齐、
     * <b>上一截是上一页最后一段内容、下一截是下一页第一段内容</b>（被分页切断的表只会是这种位置；
     * 中间夹着正文或后面还有内容的，就是各自独立的两张表）、并且后一截里找不出任何一行能当表头
     * （既没命中注册的表头集合，也没命中关键字）。四条同时成立才并，避免把两页恰好都是三列的无关表格错并到一起。
     *
     * <p>最后那条靠表头的判据，前提是调用方注册了表头集合或关键字。两者都没注册时（{@link
     * WordParser.Options#isSilent()}），任何行都算不上表头，判据永远成立，只剩几何条件撑着，两页各有一张
     * 三列等宽的独立表格就会被并成一张。这种情况下改要一条不依赖表头的证据：前一截的下边界贴着页底、
     * 后一截的上边界顶着页首（各占页高 {@value #PAGE_EDGE_REACH} 以内）。被页底切断的表必然是这个形状；
     * 一张画到页面中部就结束的表说明它是完整的，不再往下并。
     *
     * <p>跨三页以上的长表靠「并完一截就把 {@code previous} 的页码和下边界换成后一截的」继续往下接：
     * {@code continues} 要求页码相邻，{@code previous} 若一直停在第一截，第三页就跟它比不上了
     * （实测三页长表会剩两截：第 1+2 页并成一张、第 3 页另起一张，静默选项下第 3 页那截的首行还会被当成表头）。
     * 左边界不跟着换：始终与第一截比，免得每页 {@value #TOLERANCE}pt 的容差一页页累积成漂移。
     */
    private static void mergeSplitTables(List<Flow> flows, WordParser.Options options) {
        for (int i = 0; i < flows.size(); i++) {
            Flow flow = flows.get(i);
            // flows 已按「页码升序、同一页自上而下」排好，所以前后邻居就能定出这一页的第一段、最后一段
            flow.firstOnPage = i == 0 || flows.get(i - 1).page != flow.page;
            flow.lastOnPage = i + 1 == flows.size() || flows.get(i + 1).page != flow.page;
        }
        Flow previous = null;
        for (int i = 0; i < flows.size(); i++) {
            Flow flow = flows.get(i);
            if (flow.grid == null) {
                continue;
            }
            if (previous != null && continues(previous, flow, options)) {
                previous.grid = previous.grid.concat(flow.grid);
                // 并进来的这一截成了新的「尾巴」：三页以上的长表还要继续跟下一页比，页码和离页底的距离都得用它的
                previous.page = flow.page;
                previous.bottomGap = flow.bottomGap;
                flow.grid = null;
            } else {
                previous = flow;
            }
        }
    }

    private static boolean continues(Flow previous, Flow next, WordParser.Options options) {
        if (previous.page + 1 != next.page
                || previous.grid.columnCount != next.grid.columnCount
                || Math.abs(previous.left - next.left) > TOLERANCE
                || !previous.lastOnPage || !next.firstOnPage) {
            return false;
        }
        if (hasHeaderCandidate(next.grid, options)) {
            return false;
        }
        // 注册过表头依据时上面那条就是证据；静默选项下它永远不成立，再要一条纯几何的「被页底切断」证据。
        // 代价：静默选项下，上一页最后画到的那一行很高（下边界离页底超过页高的 1/5）时判不出续接，
        // 拆成两组。宁可拆开重复一个表头，也不能把两张无关的表并成一张把数据搞混。
        return !options.isSilent() || (previous.bottomGap <= PAGE_EDGE_REACH && next.topGap <= PAGE_EDGE_REACH);
    }

    /** 这一截表里有没有能当表头的行 —— 有表头说明是新的一张表，不该并到上一页的表里。 */
    private static boolean hasHeaderCandidate(Grid grid, WordParser.Options options) {
        for (int r = 0; r < grid.rowCount; r++) {
            if (grid.separator[r]) {
                continue;
            }
            List<String> cells = Arrays.asList(grid.values[r]);
            if (options.matchHeader(cells) >= 0 || options.matchesKeywords(cells)) {
                return true;
            }
        }
        return false;
    }

    /* ---------------- 表格：几何还原成网格 ---------------- */

    /**
     * 把同一页上矩形相接的 {@link AbsorbedTable} 并成一组，一组当成一张物理表。
     *
     * <p>PDF 没有表格对象，Aspose 只能按可见线段推断：合并单元格让内部线段缺失，一张表因此常被切成多块，
     * 每块自己看行列数都对不上。并回去之后列线、行线才是按整张表聚类的，跨块的值也不会漏。
     *
     * <p>两两比较是 {@code O(n^2)}（n = 本页被切出的块数）：正常一页也就几块到几十块，代价可以忽略；
     * 真要处理几千块一面的极端版面，得先按 y 坐标分桶再比较。
     */
    private static List<List<AbsorbedTable>> cluster(List<AbsorbedTable> absorbed) {
        int n = absorbed.size();
        int[] parent = new int[n];
        for (int i = 0; i < n; i++) {
            parent[i] = i;
        }
        for (int i = 0; i < n; i++) {
            for (int j = i + 1; j < n; j++) {
                if (adjacent(absorbed.get(i).getRectangle(), absorbed.get(j).getRectangle())) {
                    union(parent, i, j);
                }
            }
        }
        Map<Integer, List<AbsorbedTable>> byRoot = new LinkedHashMap<Integer, List<AbsorbedTable>>();
        List<List<AbsorbedTable>> groups = new ArrayList<List<AbsorbedTable>>();
        for (int i = 0; i < n; i++) {
            Integer root = Integer.valueOf(find(parent, i));
            List<AbsorbedTable> group = byRoot.get(root);
            if (group == null) {
                group = new ArrayList<AbsorbedTable>();
                byRoot.put(root, group);
                groups.add(group);        // 分组顺序按组内第一个块的出现顺序，即页面上从上往下
            }
            group.add(absorbed.get(i));
        }
        return groups;
    }

    /** 两个表格矩形在横向和纵向上都挨着（允许 {@value #TOLERANCE}pt 的缝隙）即认为属于同一张表。 */
    private static boolean adjacent(Rectangle a, Rectangle b) {
        if (a == null || b == null) {
            return false;
        }
        boolean xShared = a.getLLX() <= b.getURX() + TOLERANCE && b.getLLX() <= a.getURX() + TOLERANCE;
        boolean yShared = a.getLLY() <= b.getURY() + TOLERANCE && b.getLLY() <= a.getURY() + TOLERANCE;
        return xShared && yShared;
    }

    private static int find(int[] parent, int i) {
        while (parent[i] != i) {
            i = parent[i] = parent[parent[i]];
        }
        return i;
    }

    private static void union(int[] parent, int i, int j) {
        parent[find(parent, i)] = find(parent, j);
    }

    /**
     * 一组表格块还原成网格。
     *
     * <h3>解析过程</h3>
     * <ol>
     *   <li>把组内所有单元格矩形收集起来，左右边界聚类成列线、上下边界聚类成行线（容差
     *       {@value #TOLERANCE}pt），行列数就是线段数减一；</li>
     *   <li>每个单元格按自己的矩形落在哪几格写值：跨几列等于横向合并几格，跨几行等于纵向合并几行，
     *       与 {@link WordParser} 展开 {@code gridSpan}/{@code vMerge} 之后的网格一致；</li>
     *   <li>行线自下而上编号（PDF 的 y 轴朝上），换算成「自上而下」的行下标；</li>
     *   <li>整行只被一个单元格覆盖、且不是空行的，标成跨整行的合并行（小标题、分组行）。</li>
     * </ol>
     *
     * <h3>两处容忍</h3>
     * 面积为零或坐标反了的矩形直接跳过（这类单元格没有内容可言）；一个单元格覆盖到已经有内容的格子时不覆盖，
     * 先写进来的那个赢，免得重叠矩形把短标签冲掉。
     *
     * @return 没有任何有效单元格时返回 null，调用方按「这一组不是表」处理
     */
    private static Grid toGrid(List<AbsorbedTable> group) {
        List<Rectangle> cells = new ArrayList<Rectangle>();
        List<String> texts = new ArrayList<String>();
        List<Double> xs = new ArrayList<Double>();
        List<Double> ys = new ArrayList<Double>();
        for (AbsorbedTable table : group) {
            for (AbsorbedRow row : table.getRowList()) {
                for (AbsorbedCell cell : row.getCellList()) {
                    Rectangle rect = cell.getRectangle();
                    if (rect == null || rect.getURX() <= rect.getLLX() || rect.getURY() <= rect.getLLY()) {
                        continue;
                    }
                    cells.add(rect);
                    texts.add(cellText(cell));
                    xs.add(Double.valueOf(rect.getLLX()));
                    xs.add(Double.valueOf(rect.getURX()));
                    ys.add(Double.valueOf(rect.getLLY()));
                    ys.add(Double.valueOf(rect.getURY()));
                }
            }
        }
        if (cells.isEmpty()) {
            return null;
        }

        double[] columns = clusterEdges(xs);     // 升序，长度 = 列数 + 1
        double[] rowEdges = clusterEdges(ys);    // 升序（自下而上），长度 = 行数 + 1
        int columnCount = columns.length - 1;
        int rowCount = rowEdges.length - 1;
        String[][] values = new String[rowCount][columnCount];
        boolean[][] origin = new boolean[rowCount][cells.size()];
        for (String[] row : values) {
            Arrays.fill(row, "");
        }

        for (int i = 0; i < cells.size(); i++) {
            Rectangle rect = cells.get(i);
            int left = edgeIndex(columns, rect.getLLX());
            int right = edgeIndex(columns, rect.getURX());
            int bottom = edgeIndex(rowEdges, rect.getLLY());
            int top = edgeIndex(rowEdges, rect.getURY());
            if (right <= left) {
                right = Math.min(left + 1, columnCount);
                left = right - 1;
            }
            if (top <= bottom) {
                top = Math.min(bottom + 1, rowCount);
                bottom = top - 1;
            }
            // 行线是自下而上编号的（PDF 的 y 轴朝上），转成「自上而下」的行下标
            int firstRow = rowCount - top;
            int lastRow = rowCount - 1 - bottom;
            for (int r = firstRow; r <= lastRow; r++) {
                origin[r][i] = true;
                for (int c = left; c < right && c < columnCount; c++) {
                    if (values[r][c].length() == 0) {
                        values[r][c] = texts.get(i);
                    }
                }
            }
        }

        boolean[] separator = new boolean[rowCount];
        for (int r = 0; r < rowCount; r++) {
            int origins = 0;
            for (int i = 0; i < cells.size(); i++) {
                if (origin[r][i]) {
                    origins++;
                }
            }
            boolean blank = true;
            for (int c = 0; c < columnCount; c++) {
                blank &= values[r][c].length() == 0;
            }
            // 整行只有一个单元格、又不全空 —— 跨整行的合并行（小标题、分组行）
            separator[r] = columnCount > 1 && origins == 1 && !blank;
        }
        return new Grid(values, separator, columnCount, cells);
    }

    /** 单元格文本：文字片段按顺序拼接，再统一清洗空白与控制字符，评估水印整段直接当成空。 */
    private static String cellText(AbsorbedCell cell) {
        StringBuilder sb = new StringBuilder();
        for (TextFragment fragment : cell.getTextFragments()) {
            sb.append(fragment.getText());
        }
        return Results.cleanContent(sb.toString());
    }

    /** 边界聚类：排序后把相差不到 {@value #TOLERANCE}pt 的坐标并成一条线。 */
    private static double[] clusterEdges(List<Double> values) {
        double[] sorted = new double[values.size()];
        for (int i = 0; i < sorted.length; i++) {
            sorted[i] = values.get(i).doubleValue();
        }
        Arrays.sort(sorted);
        List<Double> merged = new ArrayList<Double>();
        for (double value : sorted) {
            if (!merged.isEmpty() && Math.abs(merged.get(merged.size() - 1).doubleValue() - value) <= TOLERANCE) {
                continue;
            }
            merged.add(Double.valueOf(value));
        }
        if (merged.size() < 2) {
            merged.add(Double.valueOf(merged.get(0).doubleValue() + 1));   // 只有一条线时补出至少一列/一行
        }
        double[] edges = new double[merged.size()];
        for (int i = 0; i < edges.length; i++) {
            edges[i] = merged.get(i).doubleValue();
        }
        return edges;
    }

    /** 找出与该坐标最接近的边界下标。 */
    private static int edgeIndex(double[] edges, double value) {
        int best = 0;
        double bestDiff = Math.abs(edges[0] - value);
        for (int i = 1; i < edges.length; i++) {
            double diff = Math.abs(edges[i] - value);
            if (diff < bestDiff) {
                bestDiff = diff;
                best = i;
            }
        }
        return best;
    }

    /**
     * 还原出来的一张表的网格。
     *
     * <p>{@code values[r][c]} 单元格文本，{@code separator[r]} 该行是否跨整行合并，{@code bounds} 组成它的
     * 单元格矩形（按页面上的顺序，用于取表格的上下位置和左边界）。
     */
    private static final class Grid {
        final String[][] values;
        final boolean[] separator;
        final int rowCount;
        final int columnCount;
        final List<Rectangle> bounds;

        Grid(String[][] values, boolean[] separator, int columnCount, List<Rectangle> bounds) {
            this.values = values;
            this.separator = separator;
            this.rowCount = values.length;
            this.columnCount = columnCount;
            this.bounds = bounds;
        }

        /** 接到另一张表后面：PDF 的表格按页断开，但分组要在整张表上做一次，所以先并行拼接再分组。 */
        Grid concat(Grid other) {
            String[][] rows = new String[rowCount + other.rowCount][];
            boolean[] merged = new boolean[rowCount + other.rowCount];
            System.arraycopy(values, 0, rows, 0, rowCount);
            System.arraycopy(other.values, 0, rows, rowCount, other.rowCount);
            System.arraycopy(separator, 0, merged, 0, rowCount);
            System.arraycopy(other.separator, 0, merged, rowCount, other.rowCount);
            List<Rectangle> allBounds = new ArrayList<Rectangle>(bounds);
            allBounds.addAll(other.bounds);
            return new Grid(rows, merged, Math.max(columnCount, other.columnCount), allBounds);
        }
    }

    /* ---------------- 正文文本 ---------------- */

    /**
     * 按段落读出正文，跳过表格区域内的段落（表格文字只进 {@code tables}）和 Aspose 的评估水印。
     *
     * <h3>解析过程</h3>
     * {@link ParagraphAbsorber} 把一页切成 {@link PageMarkup}，每个 markup 再切成节、节切成分段；
     * 段落要从 {@link MarkupSection} 里取 —— {@link PageMarkup#getParagraphs()} 在由 Word 转换出来的 PDF 上
     * 常常是空的，只取它会让整页正文消失。每段用 {@link #bounds} 求出文字片段的并集矩形，
     * 中心点落在某个表格矩形里就认为属于表格，跳过。
     *
     * <h3>位置拿不到时也要留内容</h3>
     * 文字片段全都没有矩形（嵌入字体缺 ToUnicode、个别生成器的畸形 PDF 会这样）时，这段文字排在页尾继续输出，
     * 而不是直接丢掉。宁可顺序不完美，也不能让正文凭空少一段。
     *
     * @param pageNumber 页码，用于和其它内容一起排序
     */
    private static List<Flow> textFlows(Page page, List<Rectangle> tableRects, int pageNumber) {
        List<Flow> flows = new ArrayList<Flow>();
        ParagraphAbsorber absorber = new ParagraphAbsorber();
        absorber.visit(page);
        for (PageMarkup markup : absorber.getPageMarkups()) {
            List<MarkupParagraph> paragraphs = new ArrayList<MarkupParagraph>();
            for (MarkupSection section : markup.getSections()) {
                paragraphs.addAll(section.getParagraphs());
            }
            if (paragraphs.isEmpty()) {
                paragraphs.addAll(markup.getParagraphs());
            }
            for (MarkupParagraph paragraph : paragraphs) {
                String text = Results.cleanContent(paragraph.getText());   // 评估水印清洗后即为空串
                if (text.length() == 0) {
                    continue;
                }
                Rectangle rect = bounds(paragraph);
                if (rect != null && inAny(rect, tableRects)) {
                    continue;                                  // 落在表格里的段落已经在 tables 里了
                }
                Flow flow = new Flow();
                flow.page = pageNumber;
                flow.top = rect == null ? 0 : rect.getURY();   // 位置未知的段落排到本页最后，但内容要留
                flow.text = text;
                flows.add(flow);
            }
        }
        return flows;
    }

    /** 段落的位置 = 它所有文字片段矩形的并集。 */
    private static Rectangle bounds(MarkupParagraph paragraph) {
        double llx = Double.MAX_VALUE;
        double lly = Double.MAX_VALUE;
        double urx = -Double.MAX_VALUE;
        double ury = -Double.MAX_VALUE;
        for (TextFragment fragment : paragraph.getFragments()) {
            Rectangle rect = fragment.getRectangle();
            if (rect == null) {
                continue;
            }
            llx = Math.min(llx, rect.getLLX());
            lly = Math.min(lly, rect.getLLY());
            urx = Math.max(urx, rect.getURX());
            ury = Math.max(ury, rect.getURY());
        }
        return urx <= llx || ury <= lly ? null : new Rectangle(llx, lly, urx, ury);
    }

    /** 段落中心点落在任一表格矩形内即认为属于表格。 */
    private static boolean inAny(Rectangle rect, List<Rectangle> tableRects) {
        double centerX = (rect.getLLX() + rect.getURX()) / 2;
        double centerY = (rect.getLLY() + rect.getURY()) / 2;
        for (Rectangle table : tableRects) {
            if (centerX >= table.getLLX() - TOLERANCE && centerX <= table.getURX() + TOLERANCE
                    && centerY >= table.getLLY() - TOLERANCE && centerY <= table.getURY() + TOLERANCE) {
                return true;
            }
        }
        return false;
    }

    /** 一组矩形里最靠下的上边界，用作表格在页面上的纵向位置（PDF 的 y 轴朝上，越大越靠上）。 */
    private static double maxURY(List<Rectangle> rects) {
        double max = -Double.MAX_VALUE;
        for (Rectangle rect : rects) {
            max = Math.max(max, rect.getURY());
        }
        return max;
    }

    /** 一组矩形里最左的左边界；{@link #continues} 用它判断跨页的两截是不是同一条竖线对齐。 */
    private static double minLLX(List<Rectangle> rects) {
        double min = Double.MAX_VALUE;
        for (Rectangle rect : rects) {
            min = Math.min(min, rect.getLLX());
        }
        return min;
    }

    /** 一组矩形里最靠下的下边界，即这张表在页面上画到的最低位置。 */
    private static double minLLY(List<Rectangle> rects) {
        double min = Double.MAX_VALUE;
        for (Rectangle rect : rects) {
            min = Math.min(min, rect.getLLY());
        }
        return min;
    }

    /* ---------------- 图片 ---------------- */

    /**
     * 图片用 {@link ImagePlacementAbsorber} 逐页取。
     *
     * <p>PDF 里同一张图可能被多页重复引用（页眉页脚 logo 等），字节只有一份：内容相同的图片由 {@link Images}
     * 按 SHA-1 指纹标记 {@code duplicateOf}。
     */
    private static List<Flow> imageFlows(Page page, int pageNumber) {
        List<Flow> flows = new ArrayList<Flow>();
        ImagePlacementAbsorber absorber = new ImagePlacementAbsorber();
        absorber.visit(page);
        for (ImagePlacement placement : absorber.getImagePlacements()) {
            XImage image = placement.getImage();
            if (image == null) {
                continue;
            }
            ByteArrayOutputStream out = new ByteArrayOutputStream();
            try {
                image.save(out);
            } catch (Exception ignored) {
                continue;        // 取不出字节的图片（加密或特殊编码）只能跳过
            }
            byte[] bytes = out.toByteArray();
            String type = Images.sniff(bytes);
            if (type == null) {
                // CCITT/JBIG2 之类浏览器不认的原始编码，统一转成 PNG 再交给调用方
                ByteArrayOutputStream png = new ByteArrayOutputStream();
                try {
                    image.save(png, com.aspose.pdf.ImageType.getPng());
                    bytes = png.toByteArray();
                    type = "PNG";
                } catch (Exception ignored) {
                    // 转换失败就保留原始字节
                }
            }
            Flow flow = new Flow();
            flow.page = pageNumber;
            flow.top = placement.getRectangle() == null ? 0 : placement.getRectangle().getURY();
            flow.image = new PendingImage(type, image.getWidth(), image.getHeight(), bytes);
            flows.add(flow);
        }
        return flows;
    }

    /* ---------------- 输出 ---------------- */

    /** 序列化成缩进 JSON，同 {@link WordParser#toJson(Object)}。 */
    public static String toJson(Object value) throws Exception {
        return Results.toJson(value);
    }

    /** 把 {@link WordParser.ImageMode#BASE64} 结果里的图片落成文件，同 {@link WordParser#writeImages}。 */
    public static void writeImages(Map<String, Object> result, File directory) throws Exception {
        Results.writeImages(result, directory);
    }

    private static void close(Document document) {
        try {
            document.close();
        } catch (Exception ignored) {
            // 关闭失败不影响已经解析出来的结果
        }
    }

    /* ---------------- 通用 ---------------- */

    /** classpath 或当前目录存在 aspose-pdf.lic 时自动注册（一个进程只需一次），避免评估模式的文字截断和水印。 */
    private static void applyLicenseIfPresent() {
        Results.applyLicense("aspose-pdf", "/aspose-pdf.lic", "aspose-pdf.lic",
                stream -> new License().setLicense(stream));
    }

    /**
     * 先注册许可证、再加载文档。
     *
     * <p>顺序不能反：Aspose.PDF 在构造 {@code Document} 时就决定了评估限制（文字层读不全），
     * lic 晚注册对已加载的文档无效。
     */
    private static Document open(String filePath) throws Exception {
        applyLicenseIfPresent();
        return new Document(filePath);
    }

    /** 同上，从输入流加载；{@link #parse} 会负责 {@code close}，传入的流本身由调用方关闭。 */
    private static Document open(InputStream inputStream) throws Exception {
        applyLicenseIfPresent();
        return new Document(inputStream);
    }

    /* ---------------- 页面内容项 ---------------- */

    /**
     * 文档里的一段内容，{@code grid}、{@code text}、{@code image} 三者只有一个非空；
     * 按页码 + 纵向位置排序后依次写入结果的三个集合。
     *
     * <p>排序做完之后 {@link #mergeSplitTables} 会把跨页的表格段并起来，被并掉的那一段 {@code grid} 置 null；
     * 留下来那一段的 {@code page}/{@code bottomGap} 改成最后一截的，好接着跟下一页比（{@code top}、{@code left}
     * 仍是第一截的，位置与左边界以它为准）。
     */
    private static final class Flow {
        int page;
        double top;
        double left;
        /**
         * 这一段内容离页首／页底有多远，占页高的比例（{@code 0} = 正贴着边界）。
         * 只有表格段会赋值；静默选项下的跨页续接判据用，见 {@link #continues}。
         */
        double topGap;
        double bottomGap;
        /** 这一段是不是它那一页的最上面一段 / 最下面一段，跨页续接判断用（见 {@link #mergeSplitTables}）。 */
        boolean firstOnPage;
        boolean lastOnPage;
        Grid grid;
        String text;
        PendingImage image;
    }

    /** 图片先攒着，等内容排好序再交给 {@link Images} 记录，编号顺序才和文档里看到的先后一致。 */
    private static final class PendingImage {
        private final String type;
        private final int width;
        private final int height;
        private final byte[] bytes;

        PendingImage(String type, int width, int height, byte[] bytes) {
            this.type = type;
            this.width = width;
            this.height = height;
            this.bytes = bytes;
        }

        void appendTo(Images images) throws Exception {
            images.add(type, width, height, bytes);
        }
    }
}
