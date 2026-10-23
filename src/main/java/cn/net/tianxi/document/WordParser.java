package cn.net.tianxi.document;

import com.aspose.words.*;

import java.io.File;
import java.io.InputStream;
import java.util.*;
import java.util.List;

/**
 * word 解析工具类（doc / docx）：一次调用返回三个集合 —— 文本、表格集合、图片集合。
 *
 * <pre>
 * Map&lt;String, Object&gt; result = WordParser.parse("D:/a.docx", "销售");   // 关键字入口
 * Map&lt;String, Object&gt; result = WordParser.parse("D:/a.docx", options);   // 选项入口，见 Options
 * String json = WordParser.toJson(result);
 * </pre>
 *
 * <p>下面的说明与数字全部来自仓库里的样例 {@code src/test/resources/word/sales-report.docx}
 * （2025 年度华东区销售业绩报告：13 段正文 + 4 张物理表 + 1 张看板截图，手工拼的 OOXML 包），
 * 环境是 JDK 8、<b>没有 lic</b>（评估模式）。每条都有测试守着，测试类见各段末尾。
 *
 * <h2>返回结构</h2>
 * <pre>
 * {
 *   "text": "全文文本，段落之间用换行拼接；表格里的文字不在这里重复一遍",
 *   "tables": [
 *     { "headers": ["片区","区域","销售额","回款率","销售额-5","回款率-6","备注"],
 *       "rows": [ {"片区":"华东大区","区域":"上海","销售额":"1280","回款率":"92%",
 *                  "销售额-5":"1465","回款率-6":"95%","备注":"超额完成"} ],
 *       "name": "季度明细",                     // 注册时 Options.name(...) 起过名字才有这一行
 *       "source": {"page":1} }
 *   ],
 *   "images": [ {"fileName":"image-1.png","type":"PNG","width":1024,"height":768,
 *               "size":670939,"contentId":"487e915db2c53907e8c1551aeaffea235423210c"} ],
 *   "unmatchedHeaders": [ {"headers":["故障部件","工单数"]} ]  // Options.unmatchedHeaders(true) 才有这一行
 * }
 * </pre>
 *
 * <p>样例实测：{@code text} 是 13 段正文（最长那段 62 个字符，完整读出，评估模式没截 Words 的正文），
 * {@code tables} 7 个表头组共 30 行，{@code images} 1 张。默认（{@link ImageMode#META}）整串 JSON 4740 个字符；
 * 换 {@link ImageMode#BASE64} 变成 899370 个字符 —— 那张 670939 字节的 PNG 编码后占 894610 个字符。
 * {@code FILE} 模式多一个 {@code "path"}，{@code BYTES} 多一个 {@code "bytes"}。
 * 每组都带 {@code source}（起始页，跨页的台账另带 {@code lastPage}），见 {@link Tables#markSource}；
 * 这份页码要问版面，不想要就 {@code Options.source(false)}（见该方法的实测耗时）。
 * {@code name} 只在注册表头时用 {@link Options#name(String)} 起过名字、且这一行确实命中那组注册表头才写，
 * 所以默认输出里没有它（实测：注册五组不起名仍是 4740 个字符；给其中三组起名后有<b>四</b>组带上了这个键
 * —— {@code 年度目标} 命中两组 —— 4828 个字符，每组多 22 个字符）；
 * 起了名字的组可以用 {@link #tablesByName(Map, String)}／{@link #rowsByName(Map, String)}
 * 按名字直接取，不必记住第几组。四个键的先后固定是 {@code headers} → {@code rows} → {@code name} → {@code source}。
 * 顶层键默认只有 {@code text}／{@code tables}／{@code images} 三个，打开
 * {@link Options#unmatchedHeaders(boolean)} 才多一个 {@code unmatchedHeaders}（注册了却一组都没命中的表头集合，
 * 与标准错误那条提示同一份规则），默认关是为了不给下游添未知键。
 * <i>（测试：{@code WordParserTest}）</i>
 *
 * <h2>表格：一个表头对应它下面的行</h2>
 * 一张物理表里有几组表头，{@code tables} 里就有几个元素；每个元素的 {@code rows} 只包含属于它上面那个表头的行，
 * 因此不会串组。样例的 4 张物理表出来是 7 组：
 * <ol>
 *   <li><b>季度销售明细</b>（1 组）：单行表头，{@code 销售额}/{@code 回款率} 各出现两次（上半年／全年），
 *       摊平后补成 {@code 销售额-5}/{@code 回款率-6}；{@code 片区} 列纵向合并，江苏、安徽两行自己没有内容，
 *       回填成起始格的 {@code 华东大区}/{@code 华东南区}，4 行数据一行不少（没被误判成分组行）；</li>
 *   <li><b>客户回款台账</b>（1 组）：两行表头，上层的 {@code 回款情况} 跨两列、下层只填了
 *       {@code 合同金额}/{@code 已回款}，连续两行表头合并成
 *       {@code ["客户名称","合同金额","已回款","回款率","备注"]}；13 行 = 12 家客户 + 合计行，
 *       合计行的金额是 {@code SUM(ABOVE)} 域<b>缓存的结果</b> {@code ["合计","2096.30","1796.10","86%",""]}；</li>
 *   <li><b>售后工单统计</b>（2 组）：{@code 硬件故障} 那格里再套一张两列小表，母表那一格不摊平子表文字
 *       （读出来是 {@code ["硬件故障","42",""]}），子表 {@code ["故障部件","工单数"]} 另成一组排在母表之后；</li>
 *   <li><b>分片区年度目标</b>（3 组）：第一行就是跨整行合并的 {@code 2025 年度（已结账）}，
 *       此时还没有任何表头，于是另起一组、列名补成 {@code ["列1","列2","列3","列4"]}（否则这行文字会无声丢掉）；
 *       接着是 {@code ["片区","年度目标","完成率","列4"]}（第四列在文档里没标题），
 *       表中间的 {@code 2026 年度（预算下达）} 又是一行分组合并行，它结束上一组、开出第三组。</li>
 * </ol>
 *
 * <p>表头行的判定优先级（任一命中即视为表头）与跨整行合并行的三种去向写在 {@link Tables}，
 * 由 {@link Options#fullRow(WordParser.FullRow)} 选：样例上 {@code KEEP}（默认）7 组／30 行，
 * {@code FIRST} 组数行数都不变、只是把 {@code 2026 年度（预算下达）} 挪到这张表的第一组，
 * {@code DROP} 6 组／28 行（{@code 2025 年度（已结账）} 那组整组消失）。
 * <i>（测试：{@code WordParserTest}、{@code TablesTest}）</i>
 *
 * <h2>列名与行的键一定对齐</h2>
 * {@code headers} 逐个等于每一行的键：表头是空格的那一列在两边都叫 {@code 列N}，同名列都加列号后缀
 * （{@code 销售额-5}），所以可以只读 {@code headers} 当列清单去建表或映射字段。合并单元格一律拆开：
 * 横向合并按列展开（列数取各行单元格数的最大值），纵向合并的延续行回填起始格的值，
 * 于是每行键数量一致、合并里填的内容不丢。命中 {@link Options#headers(String...)} 注册的表头时，
 * 列名用注册的原名（默认容忍 1 列不一致：把 {@code 备注} 注册成 {@code 说明} 仍命中，结果里就叫 {@code 说明}）。
 * <i>（测试：{@code TablesTest}、{@code CrossFormatTest}）</i>
 *
 * <h2>正文清洗</h2>
 * 文本过 {@link Results#clean} 一道：域代码 {@code \u0013…\u0014} 丢掉、域<b>结果</b>保留
 * （样例第 2 段 {@code 报告页次：3 页} 里的 {@code 3} 是 {@code PAGE} 域缓存值；第 3 段的 {@code DOCPROPERTY}
 * 没算过，整块消失，只剩 {@code 归档编号：}）；软换行 {@code \u000b}、分页符 {@code \u000c} 换成空格而不是丢弃
 * （样例第 4 段是 Shift+Enter 分的两行，读出来是 {@code …含税销售额。 数据来源：…}）；
 * 单元格结束符等控制字符丢弃，连续空格压成一个。本类不做版面计算，不调 {@code updateFields}，
 * 所以没算过的域就是空的。
 *
 * <p>块级内容控件（{@code w:sdt}，Word 模板和表单常用）包住的段落表格都读得到 —— 分发交给 {@link #walk}
 * 递归进容器，样例里 {@code 编制部门：销售管理部 审核：财务总监} 那一段就在控件里；
 * 套在 {@code w:tc} 里的控件，{@code Grid#readRow} 的浅层段落查找也看穿得到。
 * <b>没填写的控件不会给出占位文字</b>：实测 4 种写法（{@code <w:p/>} 内容、空 {@code <w:sdtContent/>}、
 * 没有 {@code w:sdtPr}、带 {@code w:placeholderText}）{@code text} 里都没有
 * {@code Click here to enter text.} —— 那句提示属于 Word 的显示层，不是文档内容。
 * <i>（测试：{@code WordParserTest}、{@code WordEdgeCasesTest}）</i>
 *
 * <h2>图片</h2>
 * 只遍历各节 Body 之下的节点收图（评估水印注入在页眉页脚，扫整篇会把它们一起收进来），
 * 按文档流顺序编号 {@code image-1.png}、{@code image-2.png}……元信息与四种输出方式见 {@link Images} 与
 * {@link ImageMode}；外链图片（Word「链接到文件」）只回报 {@code link} 路径，本类不去读那个文件、也不联网。
 * <i>（测试：{@code WordParserTest} 的图片那 7 条、{@code ImagesTest}）</i>
 *
 * <h2>读不到的部分</h2>
 * 页眉页脚、脚注尾注、批注（版式附属内容，评估水印也注入在那里）；域结果没算过时是空的；
 * 矢量图 EMF/WMF/PICT/EPS 原样报告类型，不转 PNG（浏览器不认这几种，需要时拿内容后用
 * {@code ImageData.toImage()} 自行转换）。
 *
 * <h2>评估模式与深度上限</h2>
 * 没有 {@code aspose-words.lic} 时 Aspose.Words 会往页眉页脚注入水印段落：本类只走 Body，
 * 加上 {@link Results#cleanContent} 过滤，样例的 13 段正文里没有一句水印。递归深度都有上限，
 * 超过就跳过那一支并在标准错误提示一次，避免畸形文件把栈撑爆：表格最多读
 * {@value #MAX_NESTED_TABLE_DEPTH} 层（第 9 层起跳过）、Body 之下容器节点最多钻
 * {@value #MAX_CONTAINER_DEPTH} 层。提示只在真要跳内容时出现：恰好 8 层的表读全 8 组、标准错误为空，
 * 第 9 层才提示一次；实测 12 层嵌套表触发 {@code 表格嵌套超过}、20 层嵌套控件触发 {@code 节点嵌套超过}，
 * 外层都照常解析。
 *
 * <p>{@link PdfParser} 与 {@link PptParser} 复用本类的 {@link Options}、分组算法和输出结构，
 * 三种格式的返回结果形状完全一致；同一份报告的核心表在三种格式下逐格相同
 * （<i>测试：{@code CrossFormatTest}</i>）。
 *
 * <p>本类只有解析相关的静态方法，没有命令行入口；要单独验证一个文档，直接在测试类里调
 * {@link #parse(String, String...)} 或 {@link #parse(String, Options)}。
 */
public final class WordParser {

    /**
     * 表格嵌套解析的最深层数：正常文件一两层，深到这个数以上的多半是畸形文件。
     */
    private static final int MAX_NESTED_TABLE_DEPTH = 8;

    /**
     * Body 之下容器节点（内容控件一类）最多钻的层数，超过就跳过这一支，避免把栈撑爆。
     */
    private static final int MAX_CONTAINER_DEPTH = 16;

    private WordParser() {
    }

    /**
     * 按文件路径解析 doc/docx。
     *
     * @param headerKeywords 可选的表头关键字。一行的所有非空单元格都包含其中任意一个关键字，该行即判定为表头。
     *                       建议传各列表头的共同部分：样例那张核心表的列名是
     *                       {@code 片区/区域/销售额/回款率/销售额/回款率/备注}，传 {@code "销售"} 只能命中含
     *                       {@code 销售额} 的行、其余列不含它就不算表头（判定是「每个非空格都要含关键字之一」），
     *                       所以要按整表列名的公共字来传，或者干脆用 {@link Options#headers(String...)} 注册完整表头。
     *                       不传任何关键字就是静默模式：允许「首行兜底当表头」，样例上出来 7 组／30 行。
     *                       传了关键字兜底就关掉，但样例的四张表都声明了 {@code tblLook firstRow=1}，
     *                       首行仍按表头处理，所以 {@code parse(path, "销售")} 与静默模式的 JSON 逐字节相同。
     * @return {text 全文, tables 表格集合, images 图片集合}
     */
    public static Map<String, Object> parse(String filePath, String... headerKeywords) throws Exception {
        return parse(open(filePath), options().keywords(headerKeywords));
    }

    /**
     * 从输入流解析，语义同 {@link #parse(String, String...)}。
     */
    public static Map<String, Object> parse(InputStream inputStream, String... headerKeywords) throws Exception {
        return parse(open(inputStream), options().keywords(headerKeywords));
    }

    /**
     * 按选项解析，可注册完整表头集合、整行合并行的去向与图片输出方式。
     *
     * <pre>
     * WordParser.Options options = WordParser.options()
     *         // 一组完整表头，可重复注册（下面两组来自 src/test/resources 的样例）
     *         .headers("片区", "区域", "销售额", "回款率", "销售额", "回款率", "备注")
     *         .headers("客户名称", "合同金额", "已回款", "回款率", "备注")
     *         .keywords("销售")                        // 可选，宽松兜底
     *         .fullRow(WordParser.FullRow.DROP)        // 跨整行合并的行：保留/归到首个表头/丢弃
     *         .images(WordParser.ImageMode.FILE)       // 图片直接落盘，结果里只留路径
     *         .imageDir(new File("D:/out/images"));
     * Map&lt;String, Object&gt; result = WordParser.parse("D:/a.docx", options);
     * </pre>
     *
     * @return {text 全文, tables 表格集合, images 图片集合}
     */
    public static Map<String, Object> parse(String filePath, Options options) throws Exception {
        return parse(open(filePath), options);
    }

    /**
     * 从输入流按选项解析，语义同 {@link #parse(String, Options)}；流由调用方关闭。
     */
    public static Map<String, Object> parse(InputStream inputStream, Options options) throws Exception {
        return parse(open(inputStream), options);
    }

    /**
     * 新建一份选项。
     */
    public static Options options() {
        return new Options();
    }

    /**
     * 解析的主循环：逐节走一遍 Body 的节点树，段落进 {@code text}，表格进 {@code tables}，图片进 {@code images}。
     *
     * <h3>为什么不写死「Body 只有段落和表格」</h3>
     * Body 直接子节点绝大多数是 PARAGRAPH 或 TABLE（模型限死：{@code body.appendChild(shape)} 会抛
     * {@code Cannot insert a node of this type at this location}），但块级内容控件（{@code w:sdt}，Word 里
     * 「纯文本/组」控件，模板和表单常用）、文档子块一类容器也能出现在这一层，它们把里面的段落和表格包住。
     * 所以分发交给 {@link #walk}：认得的节点按类型处理，认不出的容器往下钻，这样内容不会因为一层包装而整块丢掉。
     *
     * <h3>调用策略</h3>
     * 私有方法，只能由 {@code open(...)} 系列入口调用：许可证一定在 {@code new Document(...)} 之前注册，
     * 否则这次解析仍是评估模式。页眉页脚不在走到的范围内 —— 它们是版式附属内容，而且评估水印就注入在那里。
     */
    private static Map<String, Object> parse(Document document, Options options) throws Exception {
        Content content = new Content(options, document);

        for (int s = 0; s < document.getSections().getCount(); s++) {
            CompositeNode<?> body = document.getSections().get(s).getBody();
            for (Node node = body.getFirstChild(); node != null; node = node.getNextSibling()) {
                walk(node, content, 0);
            }
        }
        List<Map<String, Object>> unmatched =
                Tables.warnUnmatchedHeaders("[WordParser] ", options, content.matchedGroups);

        Map<String, Object> result = new LinkedHashMap<String, Object>();
        result.put("text", Results.join(content.paragraphs, "\n"));
        result.put("tables", content.tables);
        result.put("images", content.images.list());
        if (options.writesUnmatchedHeaders()) {
            result.put("unmatchedHeaders", unmatched);
        }
        return result;
    }

    /**
     * 节点分发：文本、表格、图片各归各位，容器节点递归下去。
     *
     * <h3>四类节点的处理</h3>
     * <ul>
     *   <li><b>PARAGRAPH</b>：{@code getText()} 给出整段文字（含段内文本框、域结果，域代码由
     *       {@link Results#clean} 剔除），清洗后为空就跳过；段内的图片交给 {@link #collectImages} 按出现顺序收走。
     *       列表项也是 PARAGRAPH（项目符号不进文本），浮动图片、文本框同样挂在某个段落里，所以不需要为它们另开分支，
     *       更不会读第二遍。</li>
     *   <li><b>TABLE</b>：交给 {@link #parseTable}，一张物理表可能切出多个表头组，还可能带出嵌套表。</li>
     *   <li><b>SHAPE</b>：只有脱离段落出现在容器里才会走到这里（正常文件里它在段落内，已由上一条收走）；
     *       是图片就记录，同时继续往里读，因为文本框的形状里就是段落。</li>
     *   <li><b>其它容器</b>（内容控件、书签类节点等）：本身没有内容，递归遍历子节点。
     *       深度上限 {@value #MAX_CONTAINER_DEPTH} 层，超过就提示并放弃这一支，免得畸形文件把栈撑爆。</li>
     * </ul>
     */
    private static void walk(Node node, Content content, int depth) throws Exception {
        if (node.getNodeType() == NodeType.PARAGRAPH) {
            String text = Results.cleanContent(node.getText());   // 评估水印段落清洗后即为空串
            if (text.length() > 0) {
                content.paragraphs.add(text);
            }
            collectImages((CompositeNode<?>) node, content.images);
            return;
        }
        if (node.getNodeType() == NodeType.TABLE) {
            content.tables.addAll(parseTable((Table) node, content, 0));
            return;
        }
        if (node.getNodeType() == NodeType.SHAPE && ((Shape) node).hasImage()) {
            addImage(content.images, (Shape) node);
        }
        // 叶子节点没有子节点可钻，深度再大也不会继续递归；先判 isComposite 再判深度，
        // 免得深层但无内容的书签类节点被误报成「节点嵌套超过」
        if (!node.isComposite()) {
            return;
        }
        if (depth >= MAX_CONTAINER_DEPTH) {
            System.err.println("[WordParser] 节点嵌套超过 " + MAX_CONTAINER_DEPTH + " 层，该分支已跳过："
                    + NodeType.getName(node.getNodeType()));
            return;
        }
        CompositeNode<?> composite = (CompositeNode<?>) node;
        for (Node child = composite.getFirstChild(); child != null; child = child.getNextSibling()) {
            walk(child, content, depth + 1);
        }
    }

    /**
     * 一次解析的过程状态：三个集合加上表头命中之类需要跨节点积累的账。
     */
    private static final class Content {
        final List<String> paragraphs = new ArrayList<String>();
        final List<Map<String, Object>> tables = new ArrayList<Map<String, Object>>();
        final Images images;
        /**
         * 实际命中过的表头集合下标，用来提示「注册了却没匹配上」。
         */
        final Set<Integer> matchedGroups = new LinkedHashSet<Integer>();
        final Options options;
        final Document document;
        /**
         * 版面信息，第一次遇到表格才建 —— 没有表格的文件不用付这份开销。
         */
        private LayoutCollector pages;

        Content(Options options, Document document) {
            this.options = options;
            this.document = document;
            this.images = new Images(options);
        }

        /**
         * 给这张表切出的各个表头组补来源页。
         *
         * <p>实测（{@code sales-report.docx}）：不先调 {@code updatePageLayout()}，{@code LayoutCollector}
         * 给的起止页和刷过版面之后<b>完全一样</b>（5 张表 1-1／1-2／2-2／2-2／2-2），而刷版面本身要 197ms
         * （JDK 21、无 lic），所以这里直接用打开时的版面。跨页的台账因此带上 {@code lastPage=2}。
         */
        void markSource(List<Map<String, Object>> groups, Table table) throws Exception {
            if (!options.writesSource()) {
                return;                     // 关开关就是连版面都不算：LayoutCollector 一次都不创建
            }
            if (pages == null) {
                pages = new LayoutCollector(document);
            }
            Tables.markSource(groups, pages.getStartPageIndex(table), pages.getEndPageIndex(table));
        }
    }

    /* ---------------- 表格：一个表头对应它下面的行 ---------------- */

    /**
     * 读取一张物理表成网格，再交给 {@link Tables#group} 按表头分组；单元格里的嵌套表单独成组。
     *
     * <h3>解析过程</h3>
     * <ol>
     *   <li>{@code convertToHorizontallyMergedCells} 展开横向合并，让每行单元格数与列数对齐；</li>
     *   <li>建立 {@link Grid}：逐格读文本、收图片，再回填横向/纵向合并覆盖位置的值，保证逐行读取不丢信息；</li>
     *   <li>把网格摊平成 {@code String[][]} + 「是否跨整行合并」+「是否标题行」三个数组 —— 这正是
     *       {@link Tables#group} 唯一的输入形式，三种格式都归到这一份网格；</li>
     *   <li>分组；</li>
     *   <li>最后逐个单元格找嵌套表（{@code Cell} 里直接挂的 {@code Table}），用同一套流程递归解析，
     *       结果追加在本表分组之后。</li>
     * </ol>
     *
     * <h3>为什么嵌套表要单独成组</h3>
     * {@code Cell.getText()} 会把单元格内嵌套表的文字一起摊平成一长串，母表那一格因此变成几百字的巨串，
     * 子表的行列结构全丢。改成「本格只取直接子段落的文本 + 子表另成一组」，两边的信息都保得住。
     *
     * <h3>调用策略</h3>
     * 内部递归入口是 {@code depth}，外层调用请传 0；嵌套深度超过 {@value #MAX_NESTED_TABLE_DEPTH} 层会跳过并提示，
     * 免得畸形文件把栈撑爆。图片统一进同一个 {@link Images}，所以编号仍按文档流顺序。
     */
    private static List<Map<String, Object>> parseTable(Table table, Content content, int depth) throws Exception {
        Options options = content.options;
        Set<Integer> matchedGroups = content.matchedGroups;
        Images images = content.images;
        table.convertToHorizontallyMergedCells();

        RowCollection rowNodes = table.getRows();
        int rowCount = rowNodes.getCount();
        int columnCount = 0;
        for (int r = 0; r < rowCount; r++) {
            columnCount = Math.max(columnCount, rowNodes.get(r).getCells().getCount());
        }

        Grid grid = new Grid(rowCount, columnCount);
        for (int r = 0; r < rowCount; r++) {
            grid.readRow(rowNodes.get(r), r, images);
        }
        grid.fillMergedValues(rowNodes);

        String[][] rows = new String[rowCount][];
        boolean[] separator = new boolean[rowCount];
        boolean[] heading = new boolean[rowCount];
        for (int r = 0; r < rowCount; r++) {
            rows[r] = grid.row(r).toArray(new String[columnCount]);
            separator[r] = grid.isSeparator(r);
            heading[r] = grid.heading(r, rowNodes);
        }

        // 表样式声明首行是标题（tblLook firstRow），Word 里新建表格默认带着
        boolean firstRowStyled = (table.getStyleOptions() & TableStyleOptions.FIRST_ROW) != 0;
        List<Map<String, Object>> groups =
                new ArrayList<Map<String, Object>>(Tables.group(rows, separator, heading, options, matchedGroups,
                        firstRowStyled));
        // 来源页先按本表标定，再解析嵌套表 —— 子表在自己的递归里标自己的页
        content.markSource(groups, table);
        groups.addAll(parseNestedTables(rowNodes, content, depth));
        return groups;
    }

    /**
     * 本表各单元格里直接挂着的嵌套表，逐张递归解析；超过深度上限的跳过并提示。
     */
    private static List<Map<String, Object>> parseNestedTables(RowCollection rowNodes, Content content, int depth)
            throws Exception {
        List<Map<String, Object>> nested = new ArrayList<Map<String, Object>>();
        // 本表是第 depth + 1 层；子表会是第 depth + 2 层，超过上限就只提示、不往下读。
        // 提示必须等到真有一个子表再打：到了上限但没有子表的表是正常结束，不该报「已跳过」
        boolean warned = false;
        for (int r = 0; r < rowNodes.getCount(); r++) {
            CellCollection cells = rowNodes.get(r).getCells();
            for (int c = 0; c < cells.getCount(); c++) {
                Cell cell = cells.get(c);
                if (cell == null) {
                    continue;
                }
                // 只取直接子表：更深的层次由下一次递归负责，避免同一张表读两遍
                NodeCollection<?> tables = cell.getChildNodes(NodeType.TABLE, false);
                for (int i = 0; i < tables.getCount(); i++) {
                    if (depth + 1 >= MAX_NESTED_TABLE_DEPTH) {
                        if (!warned) {
                            System.err.println("[WordParser] 表格嵌套超过 " + MAX_NESTED_TABLE_DEPTH
                                    + " 层，更里面的内容已跳过");
                            warned = true;
                        }
                        continue;
                    }
                    nested.addAll(parseTable((Table) tables.get(i), content, depth + 1));
                }
            }
        }
        return nested;
    }

    /**
     * 跨整行合并的行（小标题、分组行）如何处理。
     *
     * <p>「跨整行合并」= 整行只剩一个起始格（{@code WordParser} 看 {@code gridSpan}、{@code PdfParser} 看
     * 单元格矩形覆盖了几列、{@code PptParser} 看 {@code colSpan}），业务上通常是 {@code 2025 年度（已结账）}
     * 这类分组小标题。样例 docx 的目标表里有两行这种行，三种模式的实测结果：
     * <ul>
     *   <li>{@link #KEEP}（默认）：7 组／30 行，分组行留在它上面最近的那一组，作为该组的一行数据
     *       （其余列是空串），入库时能看出这一行属于哪个年度；</li>
     *   <li>{@link #FIRST}：组数行数与 KEEP 完全相同，只是把分组行挪到<b>同一张物理表</b>的第一组 ——
     *       样例里 {@code 2026 年度（预算下达）} 从目标表那组挪到了 {@code 2025 年度（已结账）} 开出来的
     *       序号列组（{@code ["列1","列2","列3","列4"]}），那组因此有 2 行；</li>
     *   <li>{@link #DROP}：6 组／28 行，两行分组小标题都没了，{@code 2025 年度（已结账）} 那组整组消失。</li>
     * </ul>
     *
     * <p>注意分组行<b>后面紧跟的那一行按表头处理</b>（新一段的列名，见 {@link Tables}），
     * 所以分组行之后直接放数据行会多切出一组。
     */
    public enum FullRow {
        /**
         * 保留，归到它上面最近的表头组（默认）。
         */
        KEEP,
        /**
         * 保留，统一归到同一张物理表的第一个表头组。
         */
        FIRST,
        /**
         * 直接丢弃。
         */
        DROP
    }

    /**
     * 图片内容的输出方式。
     *
     * <p>一张 3MB 的截图 base64 之后约 4MB，十几张图就能让 {@link #toJson(Object)} 的结果大到无法直接
     * 打印、入库或发给大模型，所以内容去向交给调用方决定；元信息（文件名、类型、宽高、字节数、内容指纹）
     * 任何模式下都会给。
     *
     * <p>样例上的实测差距：{@code sales-report.docx} 只有一张 670939 字节的 PNG，{@code META} 整串 JSON
     * 4740 个字符，{@code BASE64} 899370 个字符（那段 data URI 本身 894610 个字符，占 99.5%）；
     * 同一张图在 PDF 样例里被 Word 重编成 38668 字节的 JPEG，{@code META} 6473 → {@code BASE64} 58076 个字符。
     *
     * <p>内存上的差别：{@code BASE64}/{@code BYTES} 会把整篇文档的图片字节都留在结果里，图片多的文档请选
     * {@code FILE}（边解析边落盘，结果里只有路径）或 {@code META}。
     *
     * <p>默认是 {@code META}：解析的常见用途是把文字和表格入库，图片内容用不上，没必要为它付体积和内存。
     * 要把图片直接塞进 {@code <img src>} 时才显式选 {@code BASE64}，要转存或送 OCR 选 {@code FILE}/{@code BYTES}。
     */
    public enum ImageMode {
        /**
         * 结果里放 {@code data:image/png;base64,...}，前端可以直接塞进 {@code <img src>}；体积最大，要显式选。
         */
        BASE64,
        /**
         * 解析时直接把图片写进 {@link Options#imageDir(File)}，结果里只留 {@code fileName}/{@code path}。
         */
        FILE,
        /**
         * 结果里放原始 {@code byte[]}（Jackson 会编成不带 data: 前缀的 base64），适合 Java 侧转存或送 OCR。
         */
        BYTES,
        /**
         * 只要元信息、不导出内容（默认）：只关心文字和表格时最省。
         */
        META
    }

    /**
     * 解析选项：表头怎么认、跨整行合并行怎么归组、图片内容怎么输出。
     *
     * <p>三种格式共用这一份选项，{@link PdfParser} 和 {@link PptParser} 也接收它；配好一个实例可以反复用于
     * 多次解析（解析过程不改它，图片编号每次都从 1 开始）。
     *
     * <p>默认值与它们在三份样例上的效果：
     * <table border="1">
     *   <caption>选项默认值</caption>
     *   <tr><th>选项</th><th>默认</th><th>说明</th></tr>
     *   <tr><td>{@link #headers(String...)}</td><td>不注册</td>
     *       <td>不注册且不给关键字就是「静默模式」，允许首行兜底当表头：docx 7 组／30 行、pdf 6 组／45 行、
     *           pptx 6 组／22 行。注册样例的 5 组真表头后，docx 与 pdf 的 JSON 与静默模式逐字节相同，
     *           pptx 变成 6 组／21 行（下层表头行不再被当成数据）</td></tr>
     *   <tr><td>{@link #headerTolerance(int)}</td><td>1</td>
     *       <td>表头比对允许几列不一致；0 表示逐列相等。无论取值多少都至少要求一列完全相等且列数相同</td></tr>
     *   <tr><td>{@link #name(String)}</td><td>不起名</td>
     *       <td>给最近一次注册的那一组起名，命中才写 {@code name} 键；不起名时结果与静默模式逐字节相同。
     *           起过名字后就能用 {@link #tablesByName(Map, String)}／{@link #rowsByName(Map, String)} 按名字取组取行</td></tr>
     *   <tr><td>{@link #keywords(String...)}</td><td>不给</td>
     *       <td>宽松兜底：行内每个非空单元格都含关键字之一才算表头</td></tr>
     *   <tr><td>{@link #fullRow(cn.net.tianxi.document.WordParser.FullRow) fullRow(FullRow)}</td>
     *       <td>{@link FullRow#KEEP}</td>
     *       <td>见 {@link FullRow}：docx 上 KEEP 7 组／30 行、FIRST 同数量只改归属、DROP 6 组／28 行</td></tr>
     *   <tr><td>{@link #images(cn.net.tianxi.document.WordParser.ImageMode) images(ImageMode)}</td>
     *       <td>{@link ImageMode#META}</td>
     *       <td>见 {@link ImageMode}：docx 上 META 4740 字符、BASE64 899370 字符</td></tr>
     *   <tr><td>{@link #imageDir(File)}</td><td>null</td>
     *       <td>只在 {@code FILE} 模式下生效；{@code FILE} 没配它会在解析到第一张图时抛
     *           {@code IllegalStateException}，配了却不是 {@code FILE} 会在标准错误提示</td></tr>
     *   <tr><td>{@link #source(boolean)}</td><td>{@code true}</td>
     *       <td>每组表带不带 {@code source}（起始页／幻灯片序号），三种格式一起生效。
     *           docx 关掉它就不付版面构建：同一份样例开／关交替实测 74–254ms（开）对 11–35ms（关）</td></tr>
     *   <tr><td>{@link #unmatchedHeaders(boolean)}</td><td>{@code false}</td>
     *       <td>打开后结果多一个顶层键 {@code unmatchedHeaders}，列出注册了却一组都没命中的表头集合
     *           （与标准错误那条提示同一份规则）。默认关，关掉时输出逐字节不变</td></tr>
     * </table>
     *
     * <p>可变对象，链式调用即改即生效；没有做同步，配置期间别跨线程共享。配置好后并发用于多次解析是可以的
     * （{@link Images} 每次解析新建，选项本身只读），但 {@code imageDir} 指向同一目录时，重名图片由
     * {@link Images#uniqueFile} 加序号，不会互相覆盖。
     */
    public static final class Options {
        private final List<List<String>> headerGroups = new ArrayList<List<String>>();
        /**
         * 与 {@code headerGroups} 一一对应的原名，比对用小写去空格，输出列名用这一份。
         */
        private final List<List<String>> headerGroupNames = new ArrayList<List<String>>();
        /**
         * 与 {@code headerGroups} 一一对应的组名，由 {@link #name(String)} 写在最近一次注册的那一组上；没起名就是空串。
         */
        private final List<String> headerLabels = new ArrayList<String>();
        private final List<String> keywords = new ArrayList<String>();
        private FullRow fullRow = FullRow.KEEP;
        private int headerTolerance = 1;
        private ImageMode imageMode = ImageMode.META;
        private File imageDir;
        private boolean source = true;
        private boolean unmatchedHeaders;

        /**
         * 注册一组完整表头，行文本与之逐项相等（忽略大小写与空格）即判定为表头；可多次调用注册多组。
         *
         * <p>默认容忍一列不一致（见 {@link #headerTolerance(int)}），所以文档里把 {@code 备注} 写成 {@code 说明}
         * 这类个别差别不会让整组表头失效（{@code WordParserTest.registeredHeadersOverrideTheNamesReadFromTheDocument}
         * 就是这条）。
         *
         * <p>命中之后，{@code tables} 里的列名直接用这里注册的名字，不用文档里读出来的那一行 ——
         * PDF 的单元格文字会因为换行多出空格（比对时按 {@link Results#normalize} 忽略空格，所以照样命中），
         * 用注册名可以保证三种格式的键一致。
         * 注册名本身重名或为空时，仍按 {@link Tables} 的统一规则补成 {@code 名-列号} / {@code 列N}，
         * 让 {@code headers} 与行键保持对齐，所以注册时给够不重名的列名最省事。
         */
        public Options headers(String... titles) {
            List<String> group = new ArrayList<String>();
            List<String> names = new ArrayList<String>();
            for (String title : titles) {
                group.add(Results.normalize(title));
                names.add(title == null ? "" : title.trim());
            }
            headerGroups.add(group);
            headerGroupNames.add(names);
            headerLabels.add("");
            return this;
        }

        /**
         * 注册多组完整表头。
         */
        public Options headerGroups(List<List<String>> groups) {
            for (List<String> group : groups) {
                headers(group.toArray(new String[group.size()]));
            }
            return this;
        }

        /**
         * 给<b>最近一次 {@link #headers(String...)} 注册的那一组</b>起一个名字，命中之后写进结果的 {@code name} 键。
         *
         * <h3>用意</h3>
         * <p>一份报告里十几张表，通常只有几张要入库。用法是先跑一次 {@code PrintResultTest}（或 {@code -Ddoc}
         * 指自己的文档）看每组打印出来的 {@code headers}，挑出要入库的那几组，把它们的列名注册进来并顺手起个名字：
         * <pre>
         * WordParser.options()
         *     .headers("片区", "年度目标", "完成率").name("年度目标")
         *     .headers("客户名称", "合同金额", "已回款", "回款率", "备注").name("回款台账")
         *     .headers("工单类型", "工单数", "处理说明");           // 不入库的表不起名字，结果里就没有 name 键
         * </pre>
         * 之后取行不必再按下标猜第几组：{@link WordParser#tablesByName(Map, String)} 与
         * {@link WordParser#rowsByName(Map, String)} 按名字直接拿，三种格式的结果共用同一份名字。
         *
         * <h4>注意事项</h4>
         * <ul>
         *   <li><b>顺序有关</b>：{@code name} 绑定在它<b>前面</b>最近一次 {@code headers} 注册的那一组，
         *       连写两次以后者为准；用 {@link #headerGroups(List)} 批量注册时，它绑定的是<b>那一批的最后一组</b>
         *       （批量注册等价于按顺序逐组 {@code headers}），要给多组起名就改用 {@code headers(...).name(...)} 的逐组写法。
         *       一组都没注册过时调用它会在标准错误提示一句并忽略，不抛异常
         *       （与 {@link #imageDir(File)} 的处理方式一致）。</li>
         *   <li><b>只有命中才写键</b>：{@code name} 写在分组算法命中这一组注册表头的时候，
         *       所以靠 {@link #keywords(String...)}、{@code w:tblHeader}、首行兜底或分组小标题后那一条规则
         *       判出来的表头组都没有这个键。没命中的组保持 {@code headers}／{@code rows}／{@code source} 三个键不变。</li>
         *   <li><b>一个名字可以命中多组</b>：同一张物理表里出现两次表头、或跨页断开后并回一张表再切出多个表头组时，
         *       这几组都带同一个 {@code name}（实测 docx 样例的第 6、7 组是同一张年度目标表的两个表头组，
         *       列名与来源页完全相同）。这不是错误，是「组按表头切、不按物理表切」的必然结果，
         *       所以 {@code name} 不能当主键用；要一次拿全就用 {@link #rowsByName(Map, String)}，它把同名各组的行拼平。</li>
         *   <li><b>命中判定沿用 {@link #headers(String...)} 的容差</b>（默认允许一列不一致），
         *       名字跟着组走，也就跟着这条容差走 —— 文档里把 {@code 备注} 写成 {@code 说明} 照样命中并带上名字，
         *       但一行数据恰好凑得够近时也会「误命中」，收紧成 {@link #headerTolerance(int) headerTolerance(0)} 即可。</li>
         *   <li><b>多级表头照「合并之后那一行」注册，三种格式都命中</b>：分组算法先拿网格里的<b>每一行</b>单独比，
         *       两行表头合并完还会拿<b>合并结果</b>再比一次。实测注册
         *       {@code 客户名称/合同金额/已回款/回款率/备注}（即两行表头合并之后的列名）时，
         *       docx／pdf／pptx 三份样例都带上 {@code name}，把 {@link #headerTolerance(int) headerTolerance}
         *       收到 0 也一样命中，台账行数分别是 13／31／7（三份样例的客户数本来就不同）。
         *       三种格式给出的网格行不一样：docx 与 pptx 是「上层 {@code 回款情况}」「下层
         *       {@code 合同金额·已回款}」两行（docx 两行都带 {@code w:tblHeader}，靠这个标记才连着合并），
         *       pdf 直接给一行合并后的结果。所以反过来<b>只注册上层那一行</b>就有差异：docx 照样 13 行，
         *       pptx 的下层那行不再算表头、漏进数据变成 8 行，pdf 命不中（容忍度调到 2 才命中）。
         *       跨格式共用一份注册时，把上下两种写法各注册一次、起<b>同一个名字</b>最稳
         *       （实测组数与行数都不变：7 组 30 行／6 组 45 行／6 组 21 行）。
         *       同一个名字之下只要有一份注册命中了，另一份命不中也<b>不再提示</b>「表头集合未匹配到任何行」
         *       （提示的规则在 {@link Tables#warnUnmatchedHeaders}；没起名字的注册不受这条影响）。</li>
         *   <li><b>改名是契约变更</b>：{@code name} 的字符串完全由调用方决定，改了它，下游按名字取数的代码就断了。
         *       把它当字段名一样对待，进版本库。</li>
         * </ul>
         *
         * <p>结果里键的先后固定为 {@code headers} → {@code rows} → {@code name}（起过名字才有）
         * → {@code source}（{@link #source(boolean)} 开着才有）。
         *
         * @param label 组名；首尾空格会被去掉，空串或 {@code null} 等于不给这一组起名
         */
        public Options name(String label) {
            if (headerGroups.isEmpty()) {
                System.err.println("[name] 还没有注册任何表头集合，这一次 name(\"" + label + "\") 没有归属，已忽略 —— "
                        + "它绑定的是写在它前面最近一次 headers(...) 注册的那一组");
                return this;
            }
            headerLabels.set(headerLabels.size() - 1, label == null ? "" : label.trim());
            return this;
        }

        /**
         * 表头集合比对时允许多少列不一致，默认 1；传 0 表示必须逐列完全相等。
         * 无论取值多少都至少要求一列完全相等，且列数必须相同，避免容忍度过宽把数据行当成表头。
         */
        public Options headerTolerance(int columns) {
            this.headerTolerance = Math.max(0, columns);
            return this;
        }

        /**
         * 宽松兜底：行内所有非空单元格都含关键字之一即判定为表头。
         */
        public Options keywords(String... values) {
            for (String value : values) {
                if (value != null && value.trim().length() > 0) {
                    keywords.add(Results.normalize(value));
                }
            }
            return this;
        }

        /**
         * 跨整行合并的行（分组小标题）怎么归组，默认 {@link FullRow#KEEP}；传 null 等同于 KEEP。
         */
        public Options fullRow(FullRow mode) {
            this.fullRow = mode == null ? FullRow.KEEP : mode;
            return this;
        }

        /**
         * 图片内容怎么输出，默认 {@link ImageMode#META}（只给元信息）；传 null 等同于用默认值。
         */
        public Options images(ImageMode mode) {
            this.imageMode = mode == null ? ImageMode.META : mode;
            return this;
        }

        /**
         * {@link ImageMode#FILE} 时图片写到的目录，不存在会在解析时创建。
         */
        public Options imageDir(File directory) {
            this.imageDir = directory;
            return this;
        }

        /**
         * 每组表格要不要带 {@code source} 键（起始页 / 幻灯片序号），默认带。
         *
         * <p>关掉它是给「只要文字和表格、不回填原文位置」的调用方省时间的：docx 的页码要问
         * {@link LayoutCollector}，而它第一次回答时会把<b>整篇版面算一遍</b> —— 同一份
         * {@code sales-report.docx} 实测单次解析从 23–30ms（上一版没有 source）涨到 83–227ms，
         * 拆开看是加载 16–24ms、建 collector 0ms、5 张表 17 次取页码 98–131ms。
         * 传 {@code false} 之后 Word 侧不再创建 {@code LayoutCollector}，同一份样例实测回到
         * <b>11–35ms</b>。探针是同一 JVM 里开／关<b>交替</b>跑的（每种各 6 次、连跑两轮），避免后跑的那一种
         * 白捡 JIT 预热：开 74–254ms 对 关 11–35ms，两轮里「关」的最大值都低于同一轮「开」的最小值；
         * 跑完就删，仓库里没有它。
         *
         * <p>PDF 与 PPT 的页码是解析过程中本来就有的（{@code Page} 的下标、幻灯片的下标），不额外花钱，
         * 但这个键名与含义三种格式一致，所以开关对三种格式<b>一起生效</b>：关掉以后三种格式都不写 {@code source}，
         * 结果里每组只剩 {@code headers} 与 {@code rows} 两个键。
         */
        public Options source(boolean enabled) {
            this.source = enabled;
            return this;
        }

        /**
         * 结果里要不要多一个顶层键 {@code unmatchedHeaders}，列出「注册了却一组都没命中」的表头集合，默认<b>不写</b>。
         *
         * <p>这份清单与标准错误那条 {@code 表头集合未匹配到任何行} 提示是同一份规则（{@link Tables#warnUnmatchedHeaders}），
         * 每一项是 {@code {headers: [注册时写的列名]}}，注册时起过名字的另带 {@code name}；全部命中时是空列表
         * （开着就一定有这个键，不会是 {@code null}）。键的先后是 {@code text} → {@code tables} → {@code images}
         * → {@code unmatchedHeaders}，三种格式一致。
         *
         * <h3>为什么默认关</h3>
         * <p>顶层新增一个键对下游是有代价的：把结果映射成 POJO 的调用方，Jackson 默认
         * {@code FAIL_ON_UNKNOWN_PROPERTIES=true}，遇到没声明的键会抛 {@code UnrecognizedPropertyException}。
         * 所以它做成显式打开：不打开时输出与上一版<b>逐字节相同</b>
         * （{@code CrossFormatTest.unmatchedHeadersIsAnOptInTopLevelKey} 把三种格式关闭状态下的顶层键钉死成
         * {@code [text, tables, images]}，打开后是这四个键且前三个的内容一字不变；
         * {@code CrossFormatTest.documentedJsonLengthsAreTheRealOnes} 钉的是关闭状态下的 JSON 字符数）。
         *
         * <h4>什么时候值得打开</h4>
         * <p>提示只落在标准错误里，跑在容器／定时任务里往往没人看，注册写错的表现是<b>静默丢数据</b>：
         * 那一组的行会整段并进上一组表头，{@code rowsByName} 拿回来是空的，而结果本身看不出异常。
         * 把这份清单随结果一起返回，调用方就能自己决定要不要告警、失败重试还是记日志。
         *
         * <p>同一个 {@code name} 之下已有一份注册命中时，命不中的那一份<b>不进清单</b>（那一组已经找到了，再报是噪声）；
         * 没起名字的注册不受这条影响。
         *
         * @param enabled {@code true} 才写这个顶层键
         */
        public Options unmatchedHeaders(boolean enabled) {
            this.unmatchedHeaders = enabled;
            return this;
        }

        /* 供 Tables / Images 读取的包内访问器 */

        FullRow fullRow() {
            return fullRow;
        }

        ImageMode images() {
            return imageMode;
        }

        File imageDir() {
            return imageDir;
        }

        /**
         * 结果里每组要不要带 {@code source} 键，由 {@link Options#source(boolean)} 决定，默认要。
         */
        boolean writesSource() {
            return source;
        }

        /**
         * 结果里要不要带顶层键 {@code unmatchedHeaders}，由 {@link Options#unmatchedHeaders(boolean)} 决定，默认不带。
         * 三个解析器都要读它，所以放在这里而不是各写一份。
         */
        boolean writesUnmatchedHeaders() {
            return unmatchedHeaders;
        }

        /**
         * 内容是否内嵌在结果里。只有内嵌模式才做重复图片去重；
         * {@code FILE} 模式下每张图都要落成自己的文件，{@code META} 模式本来就不带内容。
         */
        boolean isInlineImageMode() {
            return imageMode == ImageMode.BASE64 || imageMode == ImageMode.BYTES;
        }

        /**
         * 调用方没给任何表头依据时，允许用首行兜底。
         */
        boolean isSilent() {
            return headerGroups.isEmpty() && keywords.isEmpty();
        }

        int headerGroupCount() {
            return headerGroups.size();
        }

        List<String> headerGroup(int index) {
            return headerGroups.get(index);
        }

        /**
         * 命中的表头集合中注册的原名，用作结果里的列名（比对用的是 {@link #headerGroup(int)} 的归一化名）。
         */
        List<String> headerTitles(int index) {
            return headerGroupNames.get(index);
        }

        /**
         * 这一组注册时起的名字，没起过是空串；由 {@link Tables} 在命中时写成结果的 {@code name} 键。
         */
        String headerName(int index) {
            return index < headerLabels.size() ? headerLabels.get(index) : "";
        }

        /**
         * 返回该行命中的表头集合下标，未命中返回 -1。
         * <p>
         * 比对规则：列数相同（忽略行尾空列）且不一致的列数不超过 {@code min(headerTolerance, 列数-1)}。
         */
        int matchHeader(List<String> cells) {
            List<String> row = new ArrayList<String>();
            for (String cell : cells) {
                row.add(Results.normalize(cell));
            }
            while (!row.isEmpty() && row.get(row.size() - 1).length() == 0) {
                row.remove(row.size() - 1);                // 末尾空列不参与比较
            }
            for (int i = 0; i < headerGroups.size(); i++) {
                List<String> group = headerGroups.get(i);
                if (group.size() != row.size()) {
                    continue;
                }
                int mismatch = 0;
                for (int c = 0; c < group.size(); c++) {
                    if (!group.get(c).equals(row.get(c))) {
                        mismatch++;
                    }
                }
                if (mismatch <= Math.min(headerTolerance, group.size() - 1)) {
                    return i;
                }
            }
            return -1;
        }

        boolean matchesKeywords(List<String> cells) {
            if (keywords.isEmpty()) {
                return false;
            }
            boolean populated = false;
            for (String cell : cells) {
                String value = Results.normalize(cell);
                if (value.length() == 0) {
                    continue;                              // 空单元格不参与判断
                }
                populated = true;
                boolean hit = false;
                for (String keyword : keywords) {
                    hit |= value.contains(keyword);
                }
                if (!hit) {
                    return false;
                }
            }
            return populated;
        }
    }

    /* ---------------- 图片 ---------------- */

    /**
     * 收集图片。
     *
     * <p>只遍历各节 Body 下的节点：评估模式下 Aspose 会往页眉页脚注入水印图片，扫整篇文档会把它们一起收进来。
     * <p>图片按文档流顺序收集 —— 段落里的由 {@link #collectImages(CompositeNode, Images)} 收，
     * 单元格里的由 {@link Grid#readRow} 收，所以 {@code images} 里的编号和文档里看到的先后一致。
     * <p>EMF/WMF/PICT/EPS 属于矢量图，浏览器不认；需要转成 PNG 时拿到内容后用
     * {@code ImageData.toImage()} 自行转换（无头服务器记得加 {@code -Djava.awt.headless=true}）。
     */
    private static void collectImages(CompositeNode<?> node, Images images) throws Exception {
        NodeCollection<?> shapes = node.getChildNodes(NodeType.SHAPE, true);
        for (int i = 0; i < shapes.getCount(); i++) {
            Shape shape = (Shape) shapes.get(i);
            if (shape.hasImage()) {
                addImage(images, shape);
            }
        }
    }

    /**
     * 记录一张图片：类型、宽高、字节，或者「链接到文件」的外链路径。
     *
     * <p>调用方只有 {@link #walk} 和 {@link Grid}，两者都已经在 {@link Images} 上按文档流顺序追加，
     * 所以这里的顺序就是图片在文档里出现的先后。
     *
     * <p>{@code getImageSize()} 对个别矢量/异常图片会返回 null，宽高按 {@link Images} 的约定传 0，
     * 不能因为一个尺寸读不出来就让整篇文档解析失败。外链图片（{@code isLinkOnly()}）的字节不在 docx 里，
     * 只回报路径，本模块不会去读那个文件，也不会联网。
     */
    private static void addImage(Images images, Shape shape) throws Exception {
        ImageData data = shape.getImageData();
        String type = ImageType.getName(data.getImageType());
        // getImageSize() 对个别矢量/异常图片会返回 null，宽高传 0 即可（Images 的约定就是「未知传 0」），
        // 不能因为一个尺寸读不出来就整篇文档解析失败
        ImageSize size = data.getImageSize();
        int width = size == null ? 0 : size.getWidthPixels();
        int height = size == null ? 0 : size.getHeightPixels();
        if (data.isLinkOnly()) {
            images.link(type, width, height, data.getSourceFullName());
        } else {
            images.add(type, width, height, data.getImageBytes());
        }
    }

    /* ---------------- 输出 ---------------- */

    /**
     * 输出成缩进 JSON，{@code parse} 的返回值直接序列化即可。
     */
    public static String toJson(Object value) throws Exception {
        return Results.toJson(value);
    }

    /**
     * 把 {@link ImageMode#BASE64} 结果里的图片落成文件；用 {@link ImageMode#FILE} 解析时不需要再调它。
     *
     * @param result    {@link #parse} 的返回值
     * @param directory 输出目录，不存在时创建
     */
    public static void writeImages(Map<String, Object> result, File directory) throws Exception {
        Results.writeImages(result, directory);
    }

    /* ---------------- 按 name 取组取行 ---------------- */

    /**
     * 取 {@code name} 命中的表头组（返回结果里 {@code tables} 中的同一批对象，不是副本）。
     *
     * <p>与 {@link #toJson}、{@link #writeImages} 一样放在本类：三种格式的返回结构相同，所以取值入口只需要一份，
     * {@link PdfParser#parse} 与 {@link PptParser#parse} 的结果同样传得进来。
     *
     * <p>同名多组是常态而不是意外：同一张物理表里两组表头、或跨页并回一张表再切出多个表头组时都会同名
     * （实测 docx 样例的 {@code 年度目标} 命中第 6、7 两组，列名与来源页完全相同）。返回列表保留各组的边界，
     * 想一次拿平铺的行用 {@link #rowsByName(Map, String)}。
     *
     * <p>返回的是结果里的原对象，改它等于改结果；要改前先自己拷一份。
     *
     * @param result {@link #parse}、{@link PdfParser#parse} 或 {@link PptParser#parse} 的返回值
     * @param name   注册时 {@link Options#name(String)} 起的那个名字，比对用 {@code equals}（不做大小写与空格归一化）
     * @return 命中的组，顺序与 {@code tables} 里一致；一个都没命中（含没起过名字、{@code result} 为 null、
     * {@code name} 为 null）返回<b>空集合</b>而不是 null，也不抛异常
     */
    @SuppressWarnings("unchecked")
    public static List<Map<String, Object>> tablesByName(Map<String, Object> result, String name) {
        List<Map<String, Object>> hits = new ArrayList<Map<String, Object>>();
        if (result == null || name == null) {
            return hits;
        }
        Object tables = result.get("tables");
        if (!(tables instanceof List)) {
            return hits;
        }
        String wanted = name.trim();
        if (wanted.length() == 0) {
            return hits;
        }
        for (Object item : (List<?>) tables) {
            if (!(item instanceof Map)) {
                continue;
            }
            Map<String, Object> table = (Map<String, Object>) item;
            if (wanted.equals(table.get("name"))) {
                hits.add(table);
            }
        }
        return hits;
    }

    /**
     * 取 {@code name} 命中各组的行，按 {@code tables} 里的组顺序拼平。
     *
     * <p>入库时直接循环这些行即可，不必关心「一张表被切成了几组」。每行还是原来的行对象（键 = 那一组的列名），
     * 所以同名各组的列名不一致时行与行的键可能不同 —— 三种格式的样例里同名列名一致，跨文档批处理时建议先问一次
     * {@link #tablesByName(Map, String)} 确认列名再落库。
     *
     * @return 命中各组的行依次拼起来的列表；没命中时为空列表
     */
    @SuppressWarnings("unchecked")
    public static List<Map<String, String>> rowsByName(Map<String, Object> result, String name) {
        List<Map<String, String>> rows = new ArrayList<Map<String, String>>();
        for (Map<String, Object> table : tablesByName(result, name)) {
            Object listed = table.get("rows");
            if (listed instanceof List) {
                for (Object row : (List<?>) listed) {
                    if (row instanceof Map) {
                        rows.add((Map<String, String>) row);
                    }
                }
            }
        }
        return rows;
    }

    /* ---------------- 通用 ---------------- */

    /**
     * classpath 或当前目录存在 aspose-words.lic 时自动注册（一个进程只需一次），避免评估水印。
     */
    private static void applyLicenseIfPresent() {
        Results.applyLicense("aspose-words", "/aspose-words.lic", "aspose-words.lic",
                stream -> new License().setLicense(stream));
    }

    /**
     * 先注册许可证、再加载文档。
     *
     * <p>顺序不能反：Aspose 在构造 {@code Document} 时就把评估限制写进了文档对象，lic 晚注册对已经加载的
     * 文档无效。原来的写法（先 new Document 再注册）会导致进程内第一次解析必带水印，之后才正常。
     */
    private static Document open(String filePath) throws Exception {
        applyLicenseIfPresent();
        return new Document(filePath);
    }

    /**
     * 同上，从输入流加载。传入的流由调用方负责关闭，本类不会关闭它。
     */
    private static Document open(InputStream inputStream) throws Exception {
        applyLicenseIfPresent();
        return new Document(inputStream);
    }

    /* ---------------- 归一化后的表格网格 ---------------- */

    /**
     * 每行 cells 数量等于列数，合并覆盖的位置补上起始格的值。
     */
    private static final class Grid {
        private final String[][] values;
        /**
         * horizontal[r][c] = true 表示该格是被左边单元格横向合并覆盖的延续格。
         */
        private final boolean[][] horizontal;
        /**
         * vertical[r][c] = true 表示该格是被上方单元格纵向合并覆盖的延续格。
         */
        private final boolean[][] vertical;
        private final int rowCount;
        private final int columnCount;

        Grid(int rowCount, int columnCount) {
            this.rowCount = rowCount;
            this.columnCount = columnCount;
            this.values = new String[rowCount][columnCount];
            this.horizontal = new boolean[rowCount][columnCount];
            this.vertical = new boolean[rowCount][columnCount];
            for (String[] row : values) {
                Arrays.fill(row, "");
            }
        }

        /**
         * 按单元格原始文本填一格，合并覆盖的位置稍后回填；顺带收集单元格里的图片。
         *
         * <p>文本和图片都只看格子自己那一层的段落（浅查找，不进嵌套表）：格子里嵌了另一张表时，那些内容属于嵌套表
         * 自己那一组（见 {@link WordParser#parseNestedTables}），在这里读会既把子表文字摊成一长串塞进母表格子，
         * 又让图片出现两遍。
         *
         * <p>格子里的内容控件（{@code w:sdt}，Word 表单和模板常见）不会漏：{@code getChildNodes(PARAGRAPH, false)}
         * 会看穿控件外壳拿到里面的段落（实测该格的直接子节点是 {@code STRUCTURED_DOCUMENT_TAG} + {@code PARAGRAPH}，
         * 浅查找仍返回 2 个段落），控件里的文字照常进这一格。
         */
        void readRow(Row row, int index, Images images) throws Exception {
            CellCollection cells = row.getCells();
            for (int c = 0; c < columnCount && c < cells.getCount(); c++) {
                Cell cell = cells.get(c);
                NodeCollection<?> ownParagraphs = cell.getChildNodes(NodeType.PARAGRAPH, false);
                StringBuilder text = new StringBuilder();
                for (int i = 0; i < ownParagraphs.getCount(); i++) {
                    Paragraph paragraph = (Paragraph) ownParagraphs.get(i);
                    if (i > 0) {
                        text.append(' ');
                    }
                    text.append(paragraph.getText());
                    collectImages(paragraph, images);
                }
                values[index][c] = Results.cleanContent(text.toString());
            }
        }

        /**
         * 标记 + 回填，分两步做，因为合并覆盖关系要先看完整行/整列才能确定：
         * <ol>
         *   <li>逐列自上而下走一遍，{@code CellMerge.PREVIOUS} 的格标成纵向延续格，并记住当前延续段起始行的下标；</li>
         *   <li>每行内把横向延续格写成正左邻的值（横向合并的文本只存在起始格里）；</li>
         *   <li>每列内把纵向延续格写成本列上一个起始格的值 —— 竖向合并的内容因此在每一行都重复出现，不丢信息。</li>
         * </ol>
         * 顺序不能颠倒：纵向回填依赖横向先把跨列的值铺开。
         */
        void fillMergedValues(RowCollection rowNodes) {
            for (int c = 0; c < columnCount; c++) {
                int originRow = -1;
                for (int r = 0; r < rowCount; r++) {
                    CellFormat format = formatAt(rowNodes.get(r).getCells(), c);
                    if (format == null) {
                        originRow = r;
                        continue;
                    }
                    if (format.getHorizontalMerge() == CellMerge.PREVIOUS) {
                        horizontal[r][c] = true;
                    }
                    if (format.getVerticalMerge() == CellMerge.PREVIOUS && originRow >= 0) {
                        vertical[r][c] = true;
                    } else {
                        originRow = r;
                    }
                }
            }
            for (int r = 0; r < rowCount; r++) {
                for (int c = 1; c < columnCount; c++) {
                    if (horizontal[r][c]) {
                        values[r][c] = values[r][c - 1];
                    }
                }
            }
            for (int c = 0; c < columnCount; c++) {
                for (int r = 1; r < rowCount; r++) {
                    if (vertical[r][c]) {
                        values[r][c] = values[r - 1][c];
                    }
                }
            }
        }

        /**
         * 是否跨整行的合并行：整行只剩一个起始格，其余格都是横向延续。
         * 这种行通常是分组小标题（样例里是 {@code 2025 年度（已结账）}、{@code 2026 年度（预算下达）}），
         * 不能当普通数据行，也不能当表头。
         *
         * <p>只排除横向延续格，<b>不</b>排除纵向延续格：某一列纵向合并到上一行时（序号列、大分类列很常见），
         * 这一行其余格子仍是自己的数据。把它们一起排除会让这种行被当成分组行 ——
         * 实测两列表「{@code 列A} 跨两行纵向合并 + {@code 列B} 每行各有值」时，第二行按旧判据「只剩一个起始格」
         * 被判成分组行，{@code KEEP} 把 {@code 值二} 抹成空串、{@code DROP} 干脆整行删掉；
         * 改后两种模式都拿到 {@code [{列A:"第一",列B:"值一"},{列A:"第一",列B:"值二"}]}。
         *
         * <p>样例核心表的 {@code 片区} 列同样是纵向合并（{@code 华东大区} 跨上海、江苏两行），
         * 那两行除了合并列还有 5 个自己的格子，新旧判据都不会误判；4 行数据一行不少，
         * 合并列的值回填成 {@code 华东大区}/{@code 华东南区}。
         *
         * <p>整行合并<b>又一格内容都没有</b>的空白行不算分组行（{@code !blank}，与 {@code PdfParser} 的
         * {@code toGrid}、{@code PptParser} 的 {@code parseTable} 同一条判据）：排版里常有这种空白合并行当间隔，
         * 把它当分组行的话，紧跟它的那一行会按「分组行之后的行是表头」被吃成表头 —— 实测两列表中间插一行空白
         * 合并行时旧判据把一张表切成两组、{@code 丙丁} 那一行变成了列名。空白行本身留成一行空值，不凭空丢内容。
         * <i>（测试：{@code WordEdgeCasesTest.verticallyMergedDataRowIsNotAMistakenSeparator}、
         * {@code WordEdgeCasesTest.blankFullWidthMergedRowIsNotASeparator}、
         * {@code WordParserTest.verticallyMergedColumnIsFilledFromItsOriginRow}）</i>
         */
        boolean isSeparator(int row) {
            if (columnCount <= 1) {
                return false;
            }
            int origins = 0;
            boolean blank = true;
            for (int c = 0; c < columnCount; c++) {
                if (!horizontal[row][c]) {
                    origins++;
                }
                blank &= values[row][c].length() == 0;
            }
            return origins == 1 && !blank;
        }

        /**
         * Word 的"重复标题行"标记 {@code w:tblHeader}，作者手工设置过的表头最可信。
         */
        boolean heading(int row, RowCollection rowNodes) {
            return Boolean.TRUE.equals(rowNodes.get(row).getRowFormat().getHeadingFormat());
        }

        /**
         * 一行的列值（合并覆盖的位置已由 {@link #fillMergedValues} 补齐）。
         */
        List<String> row(int index) {
            return Arrays.asList(Arrays.copyOf(values[index], columnCount));
        }

        /**
         * 取某列的单元格格式；该行单元格数不足（跨整行合并等）时返回 null。
         */
        private static CellFormat formatAt(CellCollection cells, int column) {
            if (column >= cells.getCount()) {
                return null;
            }
            Cell cell = cells.get(column);
            return cell == null ? null : cell.getCellFormat();
        }
    }
}
