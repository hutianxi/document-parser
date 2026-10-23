package cn.net.tianxi.document;

import org.junit.jupiter.api.Assumptions;
import org.junit.jupiter.api.Test;

import java.io.File;
import java.io.FileInputStream;
import java.io.InputStream;
import java.nio.charset.Charset;
import java.nio.file.Files;
import java.security.MessageDigest;
import java.util.*;
import java.util.concurrent.atomic.AtomicReference;

import static org.junit.jupiter.api.Assertions.assertNotNull;
import static org.junit.jupiter.api.Assertions.assertTrue;

/**
 * 「看结果」用的测试：把解析结果原样打到控制台，不做逐格断言。
 *
 * <p>其余测试类都在断言实测值，改代码时能挡住回归，但看不出「这个文档到底解析成了什么样」。
 * 本类补的就是这一面：跑一条测试，控制台直接给出正文、每组表的注册名（起过名字才有）、来源页与表头、
 * 全部数据行、图片元信息和完整 JSON。只保留一条底线断言（{@link TestSupport#assertHeadersMatchRowKeys}），
 * 因为契约破了的话输出也没有参考价值。
 *
 * <h3>报告清单（按用途分四组，共 18 份；一条命令全部跑）</h3>
 * <pre>
 * mvn -B test -Dtest=PrintResultTest                                  # 下面 18 份报告全部打印
 *
 * 一、三份样例的完整结果（7 份）
 * printWordSample                  # docx 静默：正文 / 7 组表 / 图片 / 完整 JSON
 * printPdfSample                   # pdf  静默：6 组 45 行，text 恒空
 * printPptSample                   # pptx 静默：6 组 22 行（两级表头下层当数据）
 * printWordSampleWithRegisteredHeadersAndDrop   # docx 注册 + DROP：7 组 30 行 → 6 组 28 行
 * printAllSamplesWithRegisteredNames            # 三份样例 ×「注册 + 起名」各一份
 *
 * 二、公共接口方法逐个跑（4 份）
 * printPublicApiOverloads          # 3 格式 × 四个 parse 重载（路径/流 × 关键字/选项）
 * printOptionDefaultsAndSetters    # options() 默认值、headerTolerance、headerGroups、name、keywords
 * printSerializationAndImageExport # 三个类的 toJson 与 writeImages（BASE64，只打长度）
 * printNamedMatchesAcrossFormats   # tablesByName / rowsByName：命中哪几组、拼平多少行
 *
 * 三、表头规则与选项矩阵（5 份）
 * printLedgerNameRemedies          # docx 两行表头台账：四种注册写法逐个看
 * printOptionMatrix                # 三份样例 × 7 种选项的规模 / JSON 长度 / 键
 * printAllModeJson                 # 三份样例 × 7 种选项的完整 JSON，另存 21 份 .json
 * printUnmatchedHeaders            # unmatchedHeaders 开关与未命中清单
 * printHeaderAndFieldRules         # 重名列/空列名的补名规则 + SUM 域缓存值（含现造文档）
 *
 * 四、图片与错误路径（2 份）
 * printImageModes                  # META / BYTES / BASE64 / FILE 四种图片模式
 * printErrorPaths                  # 缺文件、FILE 少 imageDir 的异常类型
 *
 * 另：printAnyDocument -Ddoc="D:/你的文档.docx"   # 看自己的文档（不给 -Ddoc 时跳过）
 * </pre>
 * 在 IDE 里直接点某一条测试的运行按钮也一样。每条用例都只用 {@link TestSupport#assertHeadersMatchRowKeys}
 * 做底线断言（契约破了输出没有参考价值），其余靠「把真实结果打出来」供人工核对；需要防回归的数值在
 * {@code CrossFormatTest}、{@code TableByNameTest}、{@code OptionsTest} 等类里另有断言。
 *
 * <h3>新报告的看点</h3>
 * <ul>
 *   <li>{@code printAllModeJson}：三份样例 × 7 种选项的<b>完整 JSON</b>。每个模式先打一行摘要
 *       （组数／行数、顶层键、标准错误行数），再原样打出整份 JSON，并另存 21 份 UTF-8 文件到
 *       {@code target/print/json/<样例名>-<模式别名>.json}。{@code printOptionMatrix} 只给长度，
 *       要逐字段核对或拿去做 diff 时看这一份；同一批选项的组数／行数与矩阵报告一致。</li>
 *   <li>{@code printPublicApiOverloads}：同一文件、同一选项下路径与流两种入参的 JSON <b>逐字符相同</b>
 *       （三份样例全打 {@code true}）。关键字入参不是选项的别名：它关掉「首行兜底」，pdf 上用
 *       {@code 片区/区域/销售额/回款率/备注} 这组关键字时是 <b>6 组 47 行</b>（静默 45 行，多出的两行是
 *       没被关键字命中的表把表头当成了数据）；docx／pptx 与静默相同（表样式本身声明了首行是标题行）。</li>
 *   <li>{@code printOptionDefaultsAndSetters}：默认 {@code isSilent=true}、{@code fullRow=KEEP}、
 *       {@code images=META}、{@code source=true}、{@code unmatchedHeaders=false}；{@code headerTolerance}
 *       没有公开 getter，用「备注写成说明」验证：默认 1 命中、收到 0 落空并提示未匹配；
 *       {@code headerGroups(List)} 与逐组 {@code headers(...)} 的 JSON 逐字符相同；{@code name(...)}
 *       只绑定它前面最近一次注册，第一组没起名时 {@code rowsByName("片区")} 是 0 行（不是报错）。</li>
 *   <li>{@code printHeaderAndFieldRules}：把「{@code 销售额-5}／{@code 回款率-6}」和「合计格有值」
 *       这两类疑问拿源文档复现 —— 重名列按 <b>物理列号</b> 补后缀（不是「第几个重名」），空表头补 {@code 列N}；
 *       现造「1 + 2 但 SUM 域缓存写 300.00」的文档读出来就是 300.00，缓存为空则给空串，证明解析器只读缓存、
 *       不执行域代码。三份样例的合计行分别打在报告第 4 节（docx 2096.30／1796.10、pdf 两格为空、
 *       ppt 1586／1384），差异来自源文件本身。</li>
 *   <li>{@code printSerializationAndImageExport}：三个类的 {@code toJson} 输出逐字符相同；
 *       {@code BASE64} 模式下 docx 899370、pdf 58076、pptx 898156 个字符（只打长度，不打正文）；
 *       {@code writeImages} 写出的文件 SHA-1 与图片元信息里的 {@code contentId} 一致；
 *       {@code META} 结果调用时只提示、不创建目录、不写文件。</li>
 * </ul>
 *
 * <h3>给要入库的那几张表起名时，就看这份输出</h3>
 * 每组标题行是 {@code #6  4 列 × 3 行，name = 年度目标，来源 {page=2}} 加一行 {@code headers = [...]}。
 * 用法是：跑 {@code printAnyDocument} 看自己的文档 → 把要入库的那几组的 {@code headers} 抄进
 * {@code WordParser.options().headers(...).name("...")} → 再跑一次确认那几组带上了名字 →
 * 入库代码用 {@code WordParser.rowsByName(result, "名字")} 取行，不必记住第几组。
 * <b>抄列名时别把 {@code 列N} 抄进去</b>：那是原文那一格为空时补出来的序号名，不是文档里的字
 * （{@code 片区, 年度目标, 完成率, 列4} 这一组注册时只写前三个）。
 *
 * <h3>控制台编码</h3>
 * 直接看控制台就行：控制台输出按 JVM 的 {@code file.encoding} 编码，cmd 里跑 {@code mvn test} 是 GBK、
 * IDE 里通常是 UTF-8，两种情况下中文都正常显示。只有把输出重定向到文件、再按另一种编码打开时才会花
 * （例如 {@code mvn test > out.txt} 之后用 UTF-8 打开）。这时候看 {@code target/print/<样例名>-<选项>.txt} ——
 * 同一份输出的 UTF-8 副本，每条测试跑完都会在末尾打印它的绝对路径；它在 {@code target/} 下，{@code mvn clean} 会带走。
 * 要强制控制台走 UTF-8，加 {@code -DargLine="-Dfile.encoding=UTF-8"}（实测加不加都是全绿，只影响编码）。
 *
 * <h3>为什么不打 BASE64 模式</h3>
 * 默认 {@code META} 模式下三份样例的 JSON 分别是 4740／6473／3526 个字符，读得完；
 * 换成 {@code BASE64} 那张 670939 字节的 PNG 会把 docx 的 JSON 撑到 899370 个字符，刷屏且看不出东西。
 * 要看图片内容请用 {@code ImageMode.FILE} + {@code imageDir(File)}，见 {@code printImageModes} 的
 * 「四种模式」报告：它只打每种模式的体积、字段与落盘路径，不打 BASE64 正文。三个类的
 * {@code toJson}／{@code writeImages} 另见 {@code printSerializationAndImageExport}：那里同样只打长度与
 * 导出文件的 SHA-1 校验，不打 BASE64 正文。
 */
class PrintResultTest {

    private static final Charset UTF_8 = Charset.forName("UTF-8");

    /**
     * 输出的 UTF-8 副本落在这里；{@code mvn clean} 一起清掉。
     */
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

    /**
     * 同一份 docx 换选项后的差别：注册 5 组真表头 + 丢掉分组小标题行（7 组 30 行 → 6 组 28 行）。
     */
    @Test
    void printWordSampleWithRegisteredHeadersAndDrop() throws Exception {
        print(TestSupport.WORD, TestSupport.registered().fullRow(WordParser.FullRow.DROP), "注册表头-DROP");
    }

    /**
     * 三份样例各跑一遍「注册 5 组表头、其中 3 组起名」（{@link TestSupport#named()}），每份一份完整报告。
     *
     * <p>实测带 {@code name} 的组（组号从 1 起）：
     * <ul>
     *   <li>docx：第 <b>1、2、6、7</b> 组 —— {@code 季度明细}／{@code 回款台账}／{@code 年度目标}／{@code 年度目标}；</li>
     *   <li>pdf：第 1、2、5、6 组；pptx：第 1、2、4、5 组。</li>
     * </ul>
     * 第 3、4 组（工单类型／故障部件）虽然注册了但没起名，所以没有 {@code name} 键；第 5 组是文档里的
     * 小标题行（{@code 列1..列4}），没有注册它。第 6、7 组是同一张物理表的两个表头组，带<b>同一个</b>
     * {@code name} —— 看这两行输出就能明白 {@code name} 不能当主键，取数要用
     * {@code WordParser.rowsByName(result, "年度目标")} 一次拿平。
     *
     * <p>docx 的台账现在<b>五种注册全部命中</b>：分组算法在多级表头合并完之后会拿合并结果再比一次，
     * 所以注册「合并之后那一行」也能命中，标准错误上一个字都不打。
     */
    @Test
    void printAllSamplesWithRegisteredNames() throws Exception {
        for (String path : TestSupport.ALL_SAMPLES) {
            print(path, TestSupport.named(), "注册表头-起名");
        }
    }

    /* ---------------- 起了名字之后拿到的数据 ---------------- */

    /**
     * 三份样例各跑一遍 {@link TestSupport#named()}，<b>按名字</b>把命中的组和全部数据行打出来 ——
     * 「注册 → 起名 → {@code rowsByName} 取数」这条链路上真正到手的数据长什么样，看这一条就够。
     *
     * <p>每个名字打三样东西：命中的是第几组（在 {@code tables} 里的下标，1 起）、那几组各自的表头与全部行、
     * {@code rowsByName} 拼平之后的行数。拼平结果与「各组按组顺序接起来」是不是同一批对象也一并打出来
     * （逐行比对象引用，不是比值），这样「同名多组要一次拿全就用 {@code rowsByName}」这句话有据可查。
     *
     * <p>末尾三个名字是<b>故意取不到的</b>：没注册过的名字、空串、{@code null}，都打出行数与内容，
     * 用来看清楚「取不到给空列表、不给 null、不抛异常」这条约定。
     */
    @Test
    void printNamedMatchesAcrossFormats() throws Exception {
        String[] names = {"季度明细", "回款台账", "年度目标"};
        StringBuilder sb = new StringBuilder(16384);
        for (String path : TestSupport.ALL_SAMPLES) {
            Map<String, Object> result = parse(path, TestSupport.named());
            TestSupport.assertHeadersMatchRowKeys(result);
            List<Map<String, Object>> tables = TestSupport.tables(result);

            line(sb, "==============================================================================");
            line(sb, "文档：" + path);
            line(sb, "选项：注册 5 组表头，其中 3 组起了名字（季度明细／回款台账／年度目标），另 2 组只注册不起名");
            line(sb, "规模：" + tables.size() + " 组 / " + TestSupport.totalRows(result) + " 行");
            line(sb, "带 name 的组：" + namedGroups(tables) + "（其余组的键里没有 name）");
            line(sb, "==============================================================================");

            for (String name : names) {
                List<Map<String, Object>> hit = WordParser.tablesByName(result, name);
                List<Map<String, String>> flat = WordParser.rowsByName(result, name);
                sb.append('\n');
                line(sb, "---- name = 「" + name + "」：命中 " + hit.size() + " 组，rowsByName 拼平 "
                        + flat.size() + " 行 ----");
                List<Map<String, String>> concatenated = new ArrayList<Map<String, String>>();
                for (Map<String, Object> table : hit) {
                    line(sb, "  tables 里的第 " + (tables.indexOf(table) + 1) + " 组，来源 " + table.get("source")
                            + "，" + TestSupport.headers(table).size() + " 列 × "
                            + TestSupport.rows(table).size() + " 行");
                    line(sb, "  headers = " + TestSupport.headers(table));
                    List<Map<String, String>> rows = TestSupport.rows(table);
                    for (int r = 0; r < rows.size(); r++) {
                        line(sb, "    行" + (r + 1) + ": " + cells(rows.get(r)));
                    }
                    concatenated.addAll(rows);
                }
                line(sb, "  rowsByName 与上面各组按组顺序接起来是同一批对象吗："
                        + sameObjects(concatenated, flat) + "（拼平 " + flat.size() + " 行 / 接起来 "
                        + concatenated.size() + " 行）");
            }

            sb.append('\n');
            line(sb, "---- 取不到的名字：给空列表，不给 null、不抛异常 ----");
            line(sb, "  rowsByName(result, \"没注册过的名字\") = " + WordParser.rowsByName(result, "没注册过的名字")
                    + "，行数 " + WordParser.rowsByName(result, "没注册过的名字").size());
            line(sb, "  rowsByName(result, \"\")             = " + WordParser.rowsByName(result, "")
                    + "，行数 " + WordParser.rowsByName(result, "").size());
            line(sb, "  rowsByName(result, null)          = " + WordParser.rowsByName(result, null)
                    + "，行数 " + WordParser.rowsByName(result, null).size());
            line(sb, "  tablesByName(result, \"没注册过的名字\") = " + WordParser.tablesByName(result, "没注册过的名字"));
            line(sb, "  只注册不起名的那两组（工单类型／故障部件）用名字取不到："
                    + WordParser.rowsByName(result, "工单类型").size() + " 行 —— 它们的结果里没有 name 键");
            line(sb, "  JSON（起名后）：" + WordParser.toJson(result).length() + " 个字符");
        }
        dump("按名字取数-三份样例", sb.toString());
    }

    /**
     * docx 那张两行表头的「客户回款台账」用四种写法注册，逐个打出命中结果与全部数据行 ——
     * 两行表头在网格里长什么样、四种写法各自靠哪一行命中、拿到多少行，一条测试看完。
     *
     * <p>四种写法都只注册台账这一组（不带其余四组），所以输出里的组数、每组的表头与行数可以直接横向比：
     * 写法只影响「台账那组带不带 {@code name}」和「标准错误上提不提示」，不该影响分组本身。
     * 每种打法都把 7 组的表头与名字全列出来，方便看有没有把别的组也带跑偏。
     *
     * <p>实测四种写法在 docx 上都命中、都是 13 行（12 家客户 + 合计）、标准错误都是空的：
     * ①② 靠多级表头合并完再比一次，③ 直接命中上层那一行（两行都带 {@code w:tblHeader} 才连着合并），
     * ④ 先命中上层那一行、再由同名规则盖住。差异在别的格式上：pdf 只有合并后那一行（③ 命不中），
     * pptx 上下两行都在但只注册上层时下层会漏进数据（③ 变 8 行）。
     */
    @Test
    void printLedgerNameRemedies() throws Exception {
        List<String> ledger = Arrays.asList("客户名称", "合同金额", "已回款", "回款率", "备注");
        StringBuilder sb = new StringBuilder(16384);
        line(sb, "==============================================================================");
        line(sb, "docx 客户回款台账：四种注册写法逐个看命中与数据");
        line(sb, "台账在文档里是两行表头 —— 上层「客户名称｜回款情况（跨两列）｜回款率｜备注」，");
        line(sb, "下层空两格填「合同金额｜已回款」；合并之后才是这 5 个列名：" + ledger);
        line(sb, "==============================================================================");

        printLedger(sb, "① 默认容忍度 1，照合并之后那一行注册（README 里最自然的写法）",
                WordParser.options().headers("客户名称", "合同金额", "已回款", "回款率", "备注").name("回款台账"),
                ledger);
        printLedger(sb, "② 容忍度调到 2，同一份注册",
                WordParser.options().headers("客户名称", "合同金额", "已回款", "回款率", "备注")
                        .name("回款台账").headerTolerance(2),
                ledger);
        printLedger(sb, "③ 照文档里真实的上层那一行注册（跨列的那格写两遍）",
                WordParser.options().headers("客户名称", "回款情况", "回款情况", "回款率", "备注").name("回款台账"),
                ledger);
        printLedger(sb, "④ 上下两种写法都注册、起同一个名字（跨 docx／pdf／pptx 共用一份注册时推荐）",
                WordParser.options()
                        .headers("客户名称", "回款情况", "回款情况", "回款率", "备注").name("回款台账")
                        .headers("客户名称", "合同金额", "已回款", "回款率", "备注").name("回款台账"),
                ledger);
        dump("docx-台账四种注册写法", sb.toString());
    }

    /**
     * 打一种注册写法的结果：全部组的表头与名字、台账那组的全部行、按名字取到的行数、标准错误原文。
     */
    private static void printLedger(StringBuilder sb, String title, final WordParser.Options options,
                                    List<String> ledger) throws Exception {
        final AtomicReference<Map<String, Object>> parsed = new AtomicReference<Map<String, Object>>();
        String stderr = TestSupport.captureStderr(() -> parsed.set(WordParser.parse(TestSupport.WORD, options)));
        Map<String, Object> result = parsed.get();
        TestSupport.assertHeadersMatchRowKeys(result);
        List<Map<String, Object>> tables = TestSupport.tables(result);

        sb.append('\n');
        line(sb, "---- " + title + " ----");
        line(sb, "  " + tables.size() + " 组 / " + TestSupport.totalRows(result) + " 行；带 name 的组："
                + namedGroups(tables));
        for (int i = 0; i < tables.size(); i++) {
            Map<String, Object> table = tables.get(i);
            line(sb, "  #" + (i + 1) + " " + TestSupport.headers(table) + name(table)
                    + "，来源 " + table.get("source"));
        }

        Map<String, Object> ledgerTable = null;
        for (Map<String, Object> table : tables) {
            if (ledger.equals(TestSupport.headers(table))) {
                ledgerTable = table;
            }
        }
        sb.append('\n');
        if (ledgerTable == null) {
            line(sb, "  台账那组没找到（没有任何一组的 headers 等于 " + ledger + "）");
        } else {
            List<Map<String, String>> rows = TestSupport.rows(ledgerTable);
            line(sb, "  台账那组：键 " + ledgerTable.keySet() + "，" + rows.size() + " 行数据");
            for (int r = 0; r < rows.size(); r++) {
                line(sb, "    行" + (r + 1) + ": " + cells(rows.get(r)));
            }
        }
        line(sb, "  rowsByName(result, \"回款台账\") 拿到 "
                + WordParser.rowsByName(result, "回款台账").size() + " 行");
        line(sb, "  标准错误：" + (stderr.trim().length() == 0 ? "（空，没有任何提示）" : stderr.trim()));
    }

    /**
     * 带了 {@code name} 的组：{@code #1=季度明细, #2=回款台账, #6=年度目标, #7=年度目标}；一个都没有时给一句说明。
     */
    private static String namedGroups(List<Map<String, Object>> tables) {
        StringBuilder sb = new StringBuilder();
        for (int i = 0; i < tables.size(); i++) {
            Object name = tables.get(i).get("name");
            if (name != null) {
                if (sb.length() > 0) {
                    sb.append(", ");
                }
                sb.append('#').append(i + 1).append('=').append(name);
            }
        }
        return sb.length() == 0 ? "（一个都没有）" : sb.toString();
    }

    /**
     * 两个行列表是不是同一批对象（逐个比引用，不比值）—— 用来说明 {@code rowsByName} 不复制数据。
     */
    private static boolean sameObjects(List<Map<String, String>> left, List<Map<String, String>> right) {
        if (left.size() != right.size()) {
            return false;
        }
        for (int i = 0; i < left.size(); i++) {
            if (left.get(i) != right.get(i)) {
                return false;
            }
        }
        return true;
    }

    /* ---------------- 公共接口方法：四个 parse 重载逐个跑 ---------------- */

    /**
     * 三种格式 × 四个 {@code parse} 重载：路径/流 × 关键字/选项，每条都把规模与 JSON 长度打出来。
     *
     * <p>这一条补的是「接口方法有没有逐个跑过」：其余报告都用同一个入口（路径 + {@link WordParser.Options}），
     * 看不出另外三个重载是否也通、结果是否一致。同一份文件、同一份选项下，路径与流两种入参应当给出
     * <b>逐字符相同</b>的 JSON；关键字入参与选项入参的差别只在「表头怎么认」，报告里直接打出来。
     *
     * <p>关键字传的是第一张表的列名 {@code 片区/区域/销售额/回款率/备注}：判定规则是「每个非空单元格都含
     * 其中之一」，所以这组关键字能命中核心表；传了关键字就关掉「首行兜底」，没被任何依据命中的表不会
     * 凭空把首行当表头 —— 这一点在 pdf 上会直接反映成组数／行数变化，报告里照实打出来。
     */
    @Test
    void printPublicApiOverloads() throws Exception {
        StringBuilder sb = new StringBuilder(16384);
        line(sb, "==============================================================================");
        line(sb, "公共入口：三种格式 × 四个 parse 重载（路径/流 × 关键字/选项）");
        line(sb, "同一份文件、同一份选项：路径与流两种入参的 JSON 应当逐字符相同。");
        line(sb, "关键字入参 = [片区, 区域, 销售额, 回款率, 备注]；传了关键字就关掉首行兜底。");
        line(sb, "==============================================================================");

        String[] keywords = {"片区", "区域", "销售额", "回款率", "备注"};
        for (String path : TestSupport.ALL_SAMPLES) {
            WordParser.Options options = TestSupport.named();
            Map<String, Object> silent = parse(path, WordParser.options());
            Map<String, Object> pathOptions = parse(path, options);
            Map<String, Object> pathKeywords = parseKeywords(path, keywords);
            Map<String, Object> streamOptions = parseStream(path, options);
            Map<String, Object> streamKeywords = parseStreamKeywords(path, keywords);

            sb.append('\n');
            line(sb, "---- " + path + "（入口类 " + parserClass(path) + "）----");
            overloadRow(sb, "parse(String, Options)", pathOptions, null);
            overloadRow(sb, "parse(InputStream, Options)", streamOptions, pathOptions);
            overloadRow(sb, "parse(String, String...)", pathKeywords, null);
            overloadRow(sb, "parse(InputStream, String...)", streamKeywords, pathKeywords);
            overloadRow(sb, "parse(String, 零个关键字)（对照组）", silent, null);
            line(sb, "  路径+Options 与 流+Options 的 JSON 逐字符相同：" + sameJson(pathOptions, streamOptions));
            line(sb, "  路径+关键字 与 流+关键字 的 JSON 逐字符相同：" + sameJson(pathKeywords, streamKeywords));
            line(sb, "  关键字结果与零关键字结果的 JSON 相同：" + sameJson(pathKeywords, silent));
        }
        dump("公共入口-四个重载", sb.toString());
    }

    /**
     * 打一个重载的规模；给了对照结果时再打「与对照是否逐字符相同」。
     */
    private static void overloadRow(StringBuilder sb, String label, Map<String, Object> result,
                                    Map<String, Object> reference) throws Exception {
        TestSupport.assertHeadersMatchRowKeys(result);
        line(sb, "  " + label + "：" + TestSupport.tables(result).size() + " 组 / "
                + TestSupport.totalRows(result) + " 行，正文 "
                + splitLines(TestSupport.text(result)).size() + " 段，图片 "
                + TestSupport.images(result).size() + " 张，JSON "
                + WordParser.toJson(result).length() + " 字符");
        if (reference != null) {
            line(sb, "    与上面的路径入参逐字符相同：" + sameJson(reference, result));
        }
    }

    /* ---------------- 选项默认值与每个 setter 的实际效果 ---------------- */

    /**
     * {@link WordParser#options()} 的默认值，以及 {@code headerTolerance}、{@code headerGroups}、
     * {@code name}、{@code keywords} 四个容易配错的方法在三份样例上的实际效果。
     *
     * <p>默认值直接读包内访问器打出来；{@code headerTolerance} 没有公开 getter，所以用行为验证：
     * 把核心表的「备注」注册成「说明」（一列不一致），默认容忍 1 时命中，收到 0 时不再命中。
     * 这一条也是「容忍度不是越小越好、也不是越大越好」的依据 —— 收紧到 0 会让个别写法的表头落空。
     */
    @Test
    void printOptionDefaultsAndSetters() throws Exception {
        StringBuilder sb = new StringBuilder(16384);
        line(sb, "==============================================================================");
        line(sb, "Options 默认值与 setter 的实际效果（三份样例）");
        line(sb, "==============================================================================");

        WordParser.Options defaults = WordParser.options();
        line(sb, "\n---- options() 默认值 ----");
        line(sb, "  isSilent = " + defaults.isSilent() + "（没注册表头、没给关键字）");
        line(sb, "  headerGroupCount = " + defaults.headerGroupCount());
        line(sb, "  fullRow = " + defaults.fullRow() + "（跨整行合并行保留并归到当前组）");
        line(sb, "  images = " + defaults.images() + "（只要元信息）");
        line(sb, "  imageDir = " + defaults.imageDir());
        line(sb, "  source = " + defaults.writesSource() + "（每组带 page / lastPage）");
        line(sb, "  unmatchedHeaders = " + defaults.writesUnmatchedHeaders()
                + "（不写顶层键；打开后是 text/tables/images/unmatchedHeaders 四个键）");
        line(sb, "  headerTolerance = 1（没有公开 getter，下面用行为验证）");

        sb.append('\n');
        line(sb, "---- headerTolerance：核心表注册时把「备注」写成「说明」（差 1 列）----");
        probeOption(sb, "默认容忍 1", WordParser.options()
                        .headers("片区", "区域", "销售额", "回款率", "销售额", "回款率", "说明").name("季度明细"),
                "季度明细");
        probeOption(sb, "headerTolerance(0)（要求逐列相等）", WordParser.options()
                        .headers("片区", "区域", "销售额", "回款率", "销售额", "回款率", "说明")
                        .name("季度明细").headerTolerance(0),
                "季度明细");

        sb.append('\n');
        line(sb, "---- headerGroups(List) 与逐组 headers(...)：同一批表头两种写法 ----");
        List<List<String>> groups = Arrays.asList(
                Arrays.asList("片区", "区域", "销售额", "回款率", "销售额", "回款率", "备注"),
                Arrays.asList("客户名称", "合同金额", "已回款", "回款率", "备注"));
        Map<String, Object> batch = parse(TestSupport.WORD, WordParser.options().headerGroups(groups));
        Map<String, Object> sequential = parse(TestSupport.WORD, WordParser.options()
                .headers("片区", "区域", "销售额", "回款率", "销售额", "回款率", "备注")
                .headers("客户名称", "合同金额", "已回款", "回款率", "备注"));
        line(sb, "  headerGroups： " + TestSupport.tables(batch).size() + " 组 / "
                + TestSupport.totalRows(batch) + " 行，JSON " + WordParser.toJson(batch).length() + " 字符");
        line(sb, "  逐组 headers： " + TestSupport.tables(sequential).size() + " 组 / "
                + TestSupport.totalRows(sequential) + " 行，JSON " + WordParser.toJson(sequential).length() + " 字符");
        line(sb, "  两种写法的 JSON 逐字符相同：" + sameJson(batch, sequential));

        sb.append('\n');
        line(sb, "---- name(String)：只给紧挨在它前面的那一组起名 ----");
        WordParser.Options binding = WordParser.options()
                .headers("片区", "区域", "销售额", "回款率", "销售额", "回款率", "备注")
                .headers("客户名称", "合同金额", "已回款", "回款率", "备注").name("只给第二组");
        Map<String, Object> bound = parse(TestSupport.WORD, binding);
        line(sb, "  注册两组、只在第二组后写 name(\"只给第二组\")：带 name 的组 = "
                + namedGroups(TestSupport.tables(bound)));
        line(sb, "  rowsByName(\"只给第二组\") = " + WordParser.rowsByName(bound, "只给第二组").size() + " 行");
        line(sb, "  rowsByName(\"片区\")（第一组没起名，取不到）= "
                + WordParser.rowsByName(bound, "片区").size() + " 行");
        final AtomicReference<String> strayName = new AtomicReference<String>("");
        strayName.set(TestSupport.captureStderr(new org.junit.jupiter.api.function.Executable() {
            @Override
            public void execute() {
                WordParser.options().name("没有注册组");
            }
        }));
        line(sb, "  一组都没注册就调 name(...)：不抛异常，标准错误 = " + oneLine(strayName.get()));

        sb.append('\n');
        line(sb, "---- keywords(String...)：docx 上 keywords(\"销售\") 的实测效果 ----");
        Map<String, Object> silent = parse(TestSupport.WORD, WordParser.options());
        Map<String, Object> keyword = parseKeywords(TestSupport.WORD, "销售");
        line(sb, "  静默： " + TestSupport.tables(silent).size() + " 组 / " + TestSupport.totalRows(silent)
                + " 行，JSON " + WordParser.toJson(silent).length() + " 字符");
        line(sb, "  keywords(\"销售\")： " + TestSupport.tables(keyword).size() + " 组 / "
                + TestSupport.totalRows(keyword) + " 行，JSON " + WordParser.toJson(keyword).length() + " 字符");
        line(sb, "  两者 JSON 逐字符相同：" + sameJson(keyword, silent)
                + "（样例四张表都声明了首行是标题行，所以关键字不是唯一依据）");
        dump("选项默认值与setter", sb.toString());
    }

    /**
     * 一种选项写法在 docx 上的命中情况：组数／行数、带 name 的组、按名字取到的行数、标准错误原文。
     */
    private static void probeOption(StringBuilder sb, String label, final WordParser.Options options,
                                    final String name) throws Exception {
        final AtomicReference<Map<String, Object>> parsed = new AtomicReference<Map<String, Object>>();
        String stderr = TestSupport.captureStderr(new org.junit.jupiter.api.function.Executable() {
            @Override
            public void execute() throws Exception {
                parsed.set(WordParser.parse(TestSupport.WORD, options));
            }
        });
        Map<String, Object> result = parsed.get();
        TestSupport.assertHeadersMatchRowKeys(result);
        List<Map<String, Object>> tables = TestSupport.tables(result);
        line(sb, "  " + label + "：" + tables.size() + " 组 / " + TestSupport.totalRows(result)
                + " 行，带 name 的组 = " + namedGroups(tables)
                + "，rowsByName(\"" + name + "\") = " + WordParser.rowsByName(result, name).size() + " 行");
        line(sb, "    标准错误：" + (stderr.trim().length() == 0 ? "（空）" : stderr.trim()));
    }

    /* ---------------- 表头补名规则与单元格里的域 ---------------- */

    /**
     * 两个常被当成 bug 的实际行为，各给一份「源文档里到底是什么」的证据：
     *
     * <ol>
     *   <li><b>重名列为什么变成 {@code 销售额-5}／{@code 回款率-6}</b>：源表头第一行确实是
     *       {@code 片区, 区域, 销售额, 回款率, 销售额, 回款率, 备注}（见 {@link TestSupport#registered()} 第一组），
     *       两个名字各出现两次。JSON 对象的键不能重复，所以第二列补成 {@code 名-绝对列号}；后缀是
     *       <b>第 5、6 列在物理表里的列号</b>，不是「第几个重名」。空表头补 {@code 列N} 用的是同一条规则
     *       （{@code Tables.columnName}），两者都由 {@code headers} 与行键共用，所以键始终对齐。</li>
     *   <li><b>合计行为什么有值</b>：docx 的「合计」格子里存的是 Word 的 {@code SUM(ABOVE)} 域，文档里带着
     *       上次计算的缓存结果（2096.30／1796.10），{@code 86%} 是普通文字。解析器读的是缓存值，不执行域代码、
     *       不重算。现造一份「1 + 2 但缓存写 300.00」的文档可以反证：读出来就是 300.00；缓存为空就给空串。
     *       PDF 是 Word 导出的，导出时前两个域没有带缓存（合同金额／已回款为空），86% 仍是字面量；
     *       PPT 的合计是作者直接写的字面量。三份样例数值不同，是因为源文件里存的内容本来就不同。</li>
     * </ol>
     */
    @Test
    void printHeaderAndFieldRules() throws Exception {
        File tmp = new File(OUT_DIR, "synthetic");
        deleteRecursively(tmp);
        assertTrue(tmp.mkdirs() || tmp.isDirectory(), "创建临时目录失败：" + tmp.getAbsolutePath());
        StringBuilder sb = new StringBuilder(16384);
        line(sb, "==============================================================================");
        line(sb, "表头补名规则与单元格里的域：两个常被当成 bug 的行为，拿源文档复现一遍");
        line(sb, "==============================================================================");

        line(sb, "\n---- 1. 三份样例里哪些列名是解析器补出来的 ----");
        for (String path : TestSupport.ALL_SAMPLES) {
            Map<String, Object> result = parse(path, WordParser.options());
            TestSupport.assertHeadersMatchRowKeys(result);
            line(sb, "  " + path);
            List<Map<String, Object>> tables = TestSupport.tables(result);
            for (int i = 0; i < tables.size(); i++) {
                List<String> headers = TestSupport.headers(tables.get(i));
                line(sb, "    #" + (i + 1) + " " + headers + "，补出来的列名 = " + synthesized(headers));
            }
        }
        line(sb, "  源表头第一行 = [片区, 区域, 销售额, 回款率, 销售额, 回款率, 备注]：两个名字各出现两次，");
        line(sb, "  第二列补成「名-绝对列号」（-5/-6 是物理表第 5、6 列），不是「第几个重名」，也不是丢了内容。");

        line(sb, "\n---- 2. 现造 docx：重名列与空列名的最小复现 ----");
        File duplicate = TestSupport.docx(tmp, "duplicate-header.docx",
                TestSupport.p("重名与空列名的最小复现"),
                TestSupport.tbl(3,
                        TestSupport.tr(TestSupport.tc("项目"), TestSupport.tc("数值"), TestSupport.tc("数值")),
                        TestSupport.tr(TestSupport.tc("甲"), TestSupport.tc("1"), TestSupport.tc("2"))));
        printSmallDoc(sb, duplicate, "表头三格依次是：项目 / 数值 / 数值");
        File blank = TestSupport.docx(tmp, "blank-header.docx",
                TestSupport.p("空表头的最小复现"),
                TestSupport.tbl(3,
                        TestSupport.tr(TestSupport.tc("项目"), TestSupport.tc(""), TestSupport.tc("备注")),
                        TestSupport.tr(TestSupport.tc("甲"), TestSupport.tc("1"), TestSupport.tc("2"))));
        printSmallDoc(sb, blank, "表头三格依次是：项目 / 空 / 备注");

        line(sb, "\n---- 3. 现造 docx：单元格里的 SUM(ABOVE) 域，读缓存值还是重算？ ----");
        File cached = TestSupport.docx(tmp, "cached-field.docx",
                TestSupport.p("数据是 1 和 2，合计格缓存写 300.00"),
                TestSupport.tbl(2,
                        TestSupport.tr(TestSupport.tc("项目"), TestSupport.tc("金额")),
                        TestSupport.tr(TestSupport.tc("甲"), TestSupport.tc("1")),
                        TestSupport.tr(TestSupport.tc("乙"), TestSupport.tc("2")),
                        TestSupport.tr(TestSupport.tc("合计"), TestSupport.tcInner(
                                TestSupport.pField("", " SUM(ABOVE) ", "300.00", "")))));
        printSmallDoc(sb, cached, "合计格 = SUM(ABOVE)，缓存结果 300.00；1 + 2 = 3");
        File cachedEmpty = TestSupport.docx(tmp, "cached-field-empty.docx",
                TestSupport.p("数据是 1 和 2，合计格的域没有缓存结果"),
                TestSupport.tbl(2,
                        TestSupport.tr(TestSupport.tc("项目"), TestSupport.tc("金额")),
                        TestSupport.tr(TestSupport.tc("甲"), TestSupport.tc("1")),
                        TestSupport.tr(TestSupport.tc("乙"), TestSupport.tc("2")),
                        TestSupport.tr(TestSupport.tc("合计"), TestSupport.tcInner(
                                TestSupport.pField("", " SUM(ABOVE) ", "", "")))));
        printSmallDoc(sb, cachedEmpty, "合计格 = SUM(ABOVE)，没有缓存结果");
        line(sb, "  结论：读的是 fldChar separate 与 end 之间的缓存结果，不执行 SUM、不重算；缓存为空就给空串。");

        line(sb, "\n---- 4. 三份样例的台账合计行（源文件里存的内容就不一样）----");
        for (String path : TestSupport.ALL_SAMPLES) {
            Map<String, Object> result = parse(path, WordParser.options());
            TestSupport.assertHeadersMatchRowKeys(result);
            for (Map<String, Object> table : TestSupport.tables(result)) {
                List<String> headers = TestSupport.headers(table);
                if (!headers.isEmpty() && "客户名称".equals(headers.get(0))) {
                    List<Map<String, String>> rows = TestSupport.rows(table);
                    line(sb, "  " + path + "：" + headers.size() + " 列 × " + rows.size() + " 行，最后一行 = "
                            + cells(rows.get(rows.size() - 1)));
                }
            }
        }
        dump("表头补名与域缓存", sb.toString());
    }

    /**
     * 解析一份现造的小文档并打出每组表头与全部行（静默模式，首行按表头处理）。
     */
    private static void printSmallDoc(StringBuilder sb, File file, String note) throws Exception {
        Map<String, Object> result = parse(file.getPath(), WordParser.options());
        TestSupport.assertHeadersMatchRowKeys(result);
        line(sb, "  " + file.getName() + "（" + note + "）：" + TestSupport.tables(result).size() + " 组 / "
                + TestSupport.totalRows(result) + " 行");
        for (Map<String, Object> table : TestSupport.tables(result)) {
            line(sb, "    headers = " + TestSupport.headers(table));
            List<Map<String, String>> rows = TestSupport.rows(table);
            for (int r = 0; r < rows.size(); r++) {
                line(sb, "      行" + (r + 1) + ": " + cells(rows.get(r)));
            }
        }
    }

    /**
     * 列名里带 {@code -编号} 或 {@code 列N} 的，视为解析器按规则补出来的名字。
     */
    private static String synthesized(List<String> headers) {
        List<String> names = new ArrayList<String>();
        for (String header : headers) {
            if (header.matches("^列\\d+$") || header.matches(".*-\\d+$")) {
                names.add(header);
            }
        }
        return names.isEmpty() ? "（无）" : names.toString();
    }

    /* ---------------- 序列化与图片导出 ---------------- */

    /**
     * 三个解析器各自的 {@code toJson}／{@code writeImages}：三份样例 × {@code BASE64} 模式，
     * 打 JSON 字符数、两个入口的 JSON 是否逐字符相同、导出文件数与每个文件的字节数／SHA-1 是否等于 {@code contentId}。
     *
     * <p>BASE64 的 JSON 会到 90 万字符量级（docx 那张 670939 字节的 PNG），所以这里<b>只打长度，不打正文</b>。
     * 另外用 {@code META} 结果调一次 {@code writeImages}，把「没有 base64 内容时只提示、不写文件」也打出来。
     */
    @Test
    void printSerializationAndImageExport() throws Exception {
        StringBuilder sb = new StringBuilder(8192);
        File exportRoot = new File(OUT_DIR, "write-images");
        deleteRecursively(exportRoot);
        line(sb, "==============================================================================");
        line(sb, "序列化与图片导出：三个类的 toJson / writeImages（BASE64 模式，只打长度）");
        line(sb, "==============================================================================");

        for (String path : TestSupport.ALL_SAMPLES) {
            Map<String, Object> result = parse(path, WordParser.options().images(WordParser.ImageMode.BASE64));
            TestSupport.assertHeadersMatchRowKeys(result);
            String shared = WordParser.toJson(result);
            String own = toJson(path, result);
            File directory = new File(exportRoot, new File(path).getName());
            writeImages(path, result, directory);

            sb.append('\n');
            line(sb, "---- " + path + "（" + parserClass(path) + "）----");
            line(sb, "  图片 " + TestSupport.images(result).size() + " 张；JSON：WordParser.toJson = "
                    + shared.length() + " 字符，本类的 toJson = " + own.length()
                    + " 字符，逐字符相同：" + shared.equals(own));
            line(sb, "  writeImages 输出目录：" + directory.getAbsolutePath());
            File[] written = directory.listFiles();
            if (written == null || written.length == 0) {
                line(sb, "  目录内没有文件（与结果里的图片张数不符，检查这条用例）");
            } else {
                Arrays.sort(written, new Comparator<File>() {
                    @Override
                    public int compare(File left, File right) {
                        return left.getName().compareTo(right.getName());
                    }
                });
                for (int i = 0; i < written.length; i++) {
                    Map<String, Object> image = TestSupport.image(result, i);
                    String contentId = String.valueOf(image.get("contentId"));
                    line(sb, "  " + written[i].getName() + "：" + written[i].length() + " 字节，SHA-1 = "
                            + sha1(written[i]) + "，与 contentId 相同：" + contentId.equals(sha1(written[i])));
                }
            }
        }

        final AtomicReference<String> metaStderr = new AtomicReference<String>("");
        metaStderr.set(TestSupport.captureStderr(new org.junit.jupiter.api.function.Executable() {
            @Override
            public void execute() throws Exception {
                WordParser.writeImages(parse(TestSupport.WORD, WordParser.options()),
                        new File(exportRoot, "meta-no-content"));
            }
        }));
        File metaDir = new File(exportRoot, "meta-no-content");
        line(sb, "\n---- META 结果调用 writeImages（没有 base64 内容）----");
        line(sb, "  标准错误 = " + oneLine(metaStderr.get()));
        line(sb, "  目录存在 = " + metaDir.isDirectory() + "，文件数 = "
                + (metaDir.list() == null ? 0 : metaDir.list().length) + "（不应写出任何文件）");
        dump("序列化与图片导出", sb.toString());
    }

    /* ---------------- 选项矩阵、未命中清单、图片模式、错误路径 ---------------- */

    /**
     * 选项矩阵与完整 JSON 共用的七种组合。
     *
     * <p>每次调用都新建一份 {@link WordParser.Options}：{@code Options} 是可变的，解析时按当前配置读，
     * 不能让两条用例或两次解析共享同一个实例。{@code slug} 只用于 JSON 文件名，避免中文和空格。
     */
    private static List<Mode> modes() {
        List<Mode> modes = new ArrayList<Mode>(7);
        modes.add(new Mode("静默 KEEP", "silent", WordParser.options()));
        modes.add(new Mode("注册 KEEP", "registered", TestSupport.registered()));
        modes.add(new Mode("注册 FIRST", "registered-first",
                TestSupport.registered().fullRow(WordParser.FullRow.FIRST)));
        modes.add(new Mode("注册 DROP", "registered-drop",
                TestSupport.registered().fullRow(WordParser.FullRow.DROP)));
        modes.add(new Mode("起名 KEEP", "named", TestSupport.named()));
        modes.add(new Mode("起名 + source(false)", "named-no-source", TestSupport.named().source(false)));
        modes.add(new Mode("起名 + unmatchedHeaders(true)", "named-unmatched",
                TestSupport.named().unmatchedHeaders(true)));
        return modes;
    }

    /**
     * 选项矩阵的一格：标签、JSON 文件名用的短名、这一格的选项。
     */
    private static final class Mode {
        final String label;
        final String slug;
        final WordParser.Options options;

        Mode(String label, String slug, WordParser.Options options) {
            this.label = label;
            this.slug = slug;
            this.options = options;
        }
    }

    /**
     * 三份样例 × 七种选项的横向对照，一行一个组合：组数／行数、JSON 字符数、带 name 的组、
     * {@code unmatchedHeaders} 组数、标准错误行数、顶层键。想看某一格的具体内容，再跑
     * {@link #printAllModeJson()}。
     *
     * <p>七个组合是：静默、注册、注册 + {@code FIRST}、注册 + {@code DROP}、注册 + 起名、
     * 起名 + {@code source(false)}、起名 + {@code unmatchedHeaders(true)}。这一条专门用来核对
     * 「换选项到底改了什么、没改什么」：例如 docx 与 pdf 的注册模式 JSON 与静默逐字节相同，
     * pptx 的注册模式会少一行（两级表头的下层）。
     */
    @Test
    void printOptionMatrix() throws Exception {
        StringBuilder sb = new StringBuilder(16384);
        for (String path : TestSupport.ALL_SAMPLES) {
            line(sb, "==============================================================================");
            line(sb, "文档：" + path);
            line(sb, "==============================================================================");
            for (Mode mode : modes()) {
                matrixRow(sb, path, mode.label, mode.options);
            }
            sb.append('\n');
        }
        dump("选项矩阵-三份样例", sb.toString());
    }

    /**
     * 三份样例 × 七种选项的<b>完整 JSON</b>：与 {@link #printOptionMatrix()} 的七个组合一一对应，
     * 只是把正文原样打出来，方便直接比较「换选项到底改了哪些字符」。
     *
     * <p>控制台会有十万字符量级，这条是排查用的，不是日常回归。每份 JSON 还会单独写一份到
     * {@code target/print/json/<样例名>-<slug>.json}：控制台按 GBK 重定向到文件时中文会花，
     * 而这三份样例的 JSON 里有中文键和中文值，直接开 UTF-8 文件更可靠。
     *
     * <p>图片的四种 {@code ImageMode} 不在这条里：{@code BASE64} 会把 docx 的 JSON 撑到 899370 个字符，
     * 那不是「看结果」而是「刷屏」；四种模式只打体积与路径，见 {@link #printImageModes()}。
     */
    @Test
    void printAllModeJson() throws Exception {
        StringBuilder sb = new StringBuilder(1 << 18);
        int written = 0;
        for (String path : TestSupport.ALL_SAMPLES) {
            File file = new File(path);
            sb.append('\n');
            line(sb, "==============================================================================");
            line(sb, "文档：" + path + "（" + file.length() + " 字节）");
            line(sb, "==============================================================================");
            for (Mode mode : modes()) {
                final AtomicReference<Map<String, Object>> parsed = new AtomicReference<Map<String, Object>>();
                String stderr = TestSupport.captureStderr(() -> parsed.set(parse(path, mode.options)));
                Map<String, Object> result = parsed.get();
                TestSupport.assertHeadersMatchRowKeys(result);
                String json = WordParser.toJson(result);

                sb.append('\n');
                line(sb, "---- 模式：" + mode.label + " ----");
                line(sb, "  " + TestSupport.tables(result).size() + " 组 / "
                        + TestSupport.totalRows(result) + " 行，顶层键 " + result.keySet()
                        + "，stderr " + (stderr.trim().isEmpty() ? "空" : splitLines(stderr).size() + " 行"));
                line(sb, "  JSON（" + json.length() + " 个字符）→ "
                        + writeModeJson(file, mode, json).getAbsolutePath());
                sb.append(json).append('\n');
                written++;
            }
        }
        line(sb, "\n共写出 " + written + " 份 JSON 到 " + new File(OUT_DIR, "json").getAbsolutePath());
        dump("全部模式-完整JSON", sb.toString());
    }

    /**
     * 打选项矩阵的一行；标准错误被捕获，只报行数，避免预期内的「未匹配」提示刷屏。
     */
    private static void matrixRow(StringBuilder sb, final String path, String label,
                                  final WordParser.Options options) throws Exception {
        final AtomicReference<Map<String, Object>> parsed = new AtomicReference<Map<String, Object>>();
        String stderr = TestSupport.captureStderr(() -> parsed.set(parse(path, options)));
        Map<String, Object> result = parsed.get();
        TestSupport.assertHeadersMatchRowKeys(result);
        List<Map<String, Object>> tables = TestSupport.tables(result);
        int jsonLength = WordParser.toJson(result).length();
        Object unmatched = result.get("unmatchedHeaders");
        String unmatchedText = unmatched == null ? "无此键"
                : ((List<?>) unmatched).size() + " 组";
        int stderrLines = 0;
        for (String line : stderr.split("\n")) {
            if (line.trim().length() > 0) {
                stderrLines++;
            }
        }
        line(sb, "  " + label + "：" + tables.size() + " 组 / " + TestSupport.totalRows(result) + " 行，JSON "
                + jsonLength + " 字符，name=" + namedGroups(tables) + "，unmatchedHeaders=" + unmatchedText
                + "，stderr=" + (stderrLines == 0 ? "空" : stderrLines + " 行")
                + "，顶层键=" + result.keySet());
    }

    /**
     * {@code Options.unmatchedHeaders(boolean)} 的直接对照：默认关闭时顶层没有这个键；
     * 打开后列出「注册了却一组都没命中」的表头集合，与标准错误上那条提示是同一份规则。
     *
     * <p>两组注册各看一遍：{@link TestSupport#named()}（五组真表头）与「只注册 {@code 故障部件} 并起名」。
     * 前者在 docx 上五组全命中、清单是空列表；pdf／pptx 上没有嵌套小表，清单里是 {@code 故障部件}。
     * 后者在 docx 上第 4 组命中，pdf／pptx 上除了 {@code headers} 还会带上注册时的 {@code name}。
     */
    @Test
    void printUnmatchedHeaders() throws Exception {
        final WordParser.Options onlyParts = WordParser.options()
                .headers("故障部件", "工单数").name("部件统计").unmatchedHeaders(true);
        StringBuilder sb = new StringBuilder(8192);
        for (final String path : TestSupport.ALL_SAMPLES) {
            line(sb, "==============================================================================");
            line(sb, "文档：" + path);
            line(sb, "==============================================================================");

            final AtomicReference<Map<String, Object>> off = new AtomicReference<Map<String, Object>>();
            String offErr = TestSupport.captureStderr(() -> off.set(parse(path, TestSupport.named())));
            line(sb, "开关关闭（named()）：顶层键=" + off.get().keySet()
                    + "，unmatchedHeaders 存在？" + off.get().containsKey("unmatchedHeaders"));
            appendStderr(sb, offErr);

            final AtomicReference<Map<String, Object>> on = new AtomicReference<Map<String, Object>>();
            String onErr = TestSupport.captureStderr(
                    () -> on.set(parse(path, TestSupport.named().unmatchedHeaders(true))));
            Map<String, Object> onResult = on.get();
            line(sb, "开关打开（named()）：顶层键=" + onResult.keySet() + "，未命中 "
                    + ((List<?>) onResult.get("unmatchedHeaders")).size() + " 组");
            line(sb, "  清单原文：");
            sb.append(wordParserJson(onResult.get("unmatchedHeaders")));
            appendStderr(sb, onErr);

            final AtomicReference<Map<String, Object>> parts = new AtomicReference<Map<String, Object>>();
            String partsErr = TestSupport.captureStderr(() -> parts.set(parse(path, onlyParts)));
            line(sb, "只注册 [故障部件, 工单数]（name=部件统计）：未命中 "
                    + ((List<?>) parts.get().get("unmatchedHeaders")).size() + " 组");
            line(sb, "  清单原文：");
            sb.append(wordParserJson(parts.get().get("unmatchedHeaders")));
            appendStderr(sb, partsErr);
            sb.append('\n');
        }
        dump("未匹配表头-三份样例", sb.toString());
    }

    /**
     * docx 那张 670939 字节的 PNG 在四种图片模式下的差别：JSON 体积、每组图片字段、落盘路径与文件大小。
     *
     * <p>BASE64 只打长度与前缀，不打正文（它会把这行报告撑到近 90 万字符）；BYTES 只打字节数组长度；
     * FILE 落到 {@code target/print/images-docx/} 并逐个核对文件存在、大小与元信息一致；
     * 另跑一遍「META + imageDir」确认只在标准错误提示、<b>不新增任何文件</b>。注意此时目录里已经有
     * FILE 模式写下的 {@code image-1.png}，所以判据是「运行前后文件清单不变」，不是「目录为空」。
     */
    @Test
    void printImageModes() throws Exception {
        File imageDir = new File(OUT_DIR, "images-docx");
        deleteRecursively(imageDir);
        StringBuilder sb = new StringBuilder(8192);
        line(sb, "==============================================================================");
        line(sb, "docx 图片四种模式：同一张 1024×768 PNG（670939 字节）");
        line(sb, "==============================================================================");

        WordParser.ImageMode[] modes = {WordParser.ImageMode.META, WordParser.ImageMode.BYTES,
                WordParser.ImageMode.BASE64, WordParser.ImageMode.FILE};
        for (WordParser.ImageMode mode : modes) {
            WordParser.Options options = WordParser.options().images(mode);
            if (mode == WordParser.ImageMode.FILE) {
                options = options.imageDir(imageDir);
            }
            Map<String, Object> result = parse(TestSupport.WORD, options);
            TestSupport.assertHeadersMatchRowKeys(result);
            Map<String, Object> image = TestSupport.image(result, 0);
            line(sb, "\n---- " + mode + " ----");
            line(sb, "  JSON：" + WordParser.toJson(result).length() + " 字符，图片 " + TestSupport.images(result).size() + " 张");
            line(sb, "  字段：" + imageSummary(image));
            if (mode == WordParser.ImageMode.FILE) {
                File written = new File(String.valueOf(image.get("path")));
                line(sb, "  落盘：" + written.getAbsolutePath() + "，存在=" + written.isFile()
                        + "，字节=" + written.length());
                String[] names = imageDir.list();
                line(sb, "  目录内文件：" + (names == null ? "（目录不存在）" : Arrays.toString(names)));
            }
        }

        final AtomicReference<Map<String, Object>> meta = new AtomicReference<Map<String, Object>>();
        List<String> before = listFiles(imageDir);
        String stderr = TestSupport.captureStderr(
                () -> meta.set(parse(TestSupport.WORD,
                        WordParser.options().images(WordParser.ImageMode.META).imageDir(imageDir))));
        List<String> after = listFiles(imageDir);
        List<String> added = new ArrayList<String>(after);
        added.removeAll(before);
        line(sb, "\n---- META + imageDir（应只提示、不新增文件；目录里已有 FILE 模式写的文件）----");
        line(sb, "  图片字段：" + imageSummary(TestSupport.image(meta.get(), 0)));
        appendStderr(sb, stderr);
        line(sb, "  目录内文件：运行前 " + before + "，运行后 " + after + "，新增 " + added.size() + " 个"
                + (added.isEmpty() ? "" : " -> " + added));
        dump("docx-图片四种模式", sb.toString());
    }

    /**
     * 错误路径的实测异常类型：三个解析器对「文件不存在」抛的不是同一种异常，图片 {@code FILE} 模式
     * 少配 {@code imageDir} 时抛 {@link IllegalStateException}。这里原样打出异常类名与消息，
     * 方便调用方决定捕获哪一类，也方便核对文档里的说明。
     */
    @Test
    void printErrorPaths() throws Exception {
        StringBuilder sb = new StringBuilder(8192);
        line(sb, "==============================================================================");
        line(sb, "错误路径：只看异常类型与消息，不修改任何文件");
        line(sb, "==============================================================================");

        String missing = "target/print/definitely-missing";
        probe(sb, "WordParser.parse(不存在的 .docx)", () -> WordParser.parse(missing + ".docx", WordParser.options()));
        probe(sb, "PdfParser.parse(不存在的 .pdf)", () -> PdfParser.parse(missing + ".pdf", WordParser.options()));
        probe(sb, "PptParser.parse(不存在的 .pptx)", () -> PptParser.parse(missing + ".pptx", WordParser.options()));

        sb.append('\n');
        probe(sb, "WordParser.parse(有图文档, FILE 模式但没给 imageDir)",
                () -> WordParser.parse(TestSupport.WORD, WordParser.options().images(WordParser.ImageMode.FILE)));
        probe(sb, "PdfParser.parse(有图文档, FILE 模式但没给 imageDir)",
                () -> PdfParser.parse(TestSupport.PDF, WordParser.options().images(WordParser.ImageMode.FILE)));
        probe(sb, "PptParser.parse(有图文档, FILE 模式但没给 imageDir)",
                () -> PptParser.parse(TestSupport.PPT, WordParser.options().images(WordParser.ImageMode.FILE)));
        dump("错误路径-三份样例", sb.toString());
    }

    /**
     * 跑一次注定失败的调用，打出异常全名与消息，并附上标准错误的前两行（Aspose 有时把细节写在那里）。
     */
    private static void probe(StringBuilder sb, String title, ThrowingRunnable action) {
        line(sb, "---- " + title + " ----");
        final AtomicReference<String> stderr = new AtomicReference<String>("");
        try {
            stderr.set(TestSupport.captureStderr(action::run));
            line(sb, "  没有抛异常（与预期不符，检查这条用例）");
        } catch (Throwable thrown) {
            line(sb, "  异常：" + thrown.getClass().getName());
            line(sb, "  消息：" + oneLine(thrown.getMessage()));
        }
        appendStderr(sb, stderr.get(), 2);
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

    /* ---------------- 按入参形式分派：路径/流 × 关键字/选项 ---------------- */

    /**
     * 按扩展名选解析器，走「路径 + 关键字」重载。
     */
    private static Map<String, Object> parseKeywords(String path, String... keywords) throws Exception {
        String name = path.toLowerCase(Locale.ROOT);
        if (name.endsWith(".pdf")) {
            return PdfParser.parse(path, keywords);
        }
        if (name.endsWith(".ppt") || name.endsWith(".pptx")) {
            return PptParser.parse(path, keywords);
        }
        if (name.endsWith(".doc") || name.endsWith(".docx")) {
            return WordParser.parse(path, keywords);
        }
        throw new IllegalArgumentException("只支持 doc/docx/pdf/ppt/pptx，收到：" + path);
    }

    /**
     * 按扩展名选解析器，走「流 + 选项」重载；流由本方法打开并关闭（解析器不接管调用方传入的流）。
     */
    private static Map<String, Object> parseStream(String path, WordParser.Options options) throws Exception {
        InputStream stream = new FileInputStream(path);
        try {
            return parseStream(path, stream, options);
        } finally {
            stream.close();
        }
    }

    /**
     * 按扩展名选解析器，走「流 + 关键字」重载；流由本方法打开并关闭。
     */
    private static Map<String, Object> parseStreamKeywords(String path, String... keywords) throws Exception {
        InputStream stream = new FileInputStream(path);
        try {
            String name = path.toLowerCase(Locale.ROOT);
            if (name.endsWith(".pdf")) {
                return PdfParser.parse(stream, keywords);
            }
            if (name.endsWith(".ppt") || name.endsWith(".pptx")) {
                return PptParser.parse(stream, keywords);
            }
            if (name.endsWith(".doc") || name.endsWith(".docx")) {
                return WordParser.parse(stream, keywords);
            }
            throw new IllegalArgumentException("只支持 doc/docx/pdf/ppt/pptx，收到：" + path);
        } finally {
            stream.close();
        }
    }

    private static Map<String, Object> parseStream(String path, InputStream stream,
                                                   WordParser.Options options) throws Exception {
        String name = path.toLowerCase(Locale.ROOT);
        if (name.endsWith(".pdf")) {
            return PdfParser.parse(stream, options);
        }
        if (name.endsWith(".ppt") || name.endsWith(".pptx")) {
            return PptParser.parse(stream, options);
        }
        if (name.endsWith(".doc") || name.endsWith(".docx")) {
            return WordParser.parse(stream, options);
        }
        throw new IllegalArgumentException("只支持 doc/docx/pdf/ppt/pptx，收到：" + path);
    }

    /**
     * 报告里显示的入口类名。
     */
    private static String parserClass(String path) {
        String name = path.toLowerCase(Locale.ROOT);
        if (name.endsWith(".pdf")) {
            return "PdfParser";
        }
        if (name.endsWith(".ppt") || name.endsWith(".pptx")) {
            return "PptParser";
        }
        return "WordParser";
    }

    private static boolean sameJson(Map<String, Object> left, Map<String, Object> right) throws Exception {
        return WordParser.toJson(left).equals(WordParser.toJson(right));
    }

    /**
     * 调该格式自己的 {@code toJson}（三个类的方法签名相同、实现共用 {@code Results.toJson}）。
     */
    private static String toJson(String path, Map<String, Object> result) throws Exception {
        String name = path.toLowerCase(Locale.ROOT);
        if (name.endsWith(".pdf")) {
            return PdfParser.toJson(result);
        }
        if (name.endsWith(".ppt") || name.endsWith(".pptx")) {
            return PptParser.toJson(result);
        }
        return WordParser.toJson(result);
    }

    /**
     * 调该格式自己的 {@code writeImages}。
     */
    private static void writeImages(String path, Map<String, Object> result, File directory) throws Exception {
        String name = path.toLowerCase(Locale.ROOT);
        if (name.endsWith(".pdf")) {
            PdfParser.writeImages(result, directory);
        } else if (name.endsWith(".ppt") || name.endsWith(".pptx")) {
            PptParser.writeImages(result, directory);
        } else {
            WordParser.writeImages(result, directory);
        }
    }

    private static String sha1(File file) throws Exception {
        MessageDigest digest = MessageDigest.getInstance("SHA-1");
        InputStream stream = new FileInputStream(file);
        try {
            byte[] buffer = new byte[8192];
            int read;
            while ((read = stream.read(buffer)) > 0) {
                digest.update(buffer, 0, read);
            }
        } finally {
            stream.close();
        }
        StringBuilder hex = new StringBuilder(40);
        for (byte value : digest.digest()) {
            hex.append(Character.forDigit((value >> 4) & 0xf, 16));
            hex.append(Character.forDigit(value & 0xf, 16));
        }
        return hex.toString();
    }

    /* ---------------- 输出 ---------------- */

    /**
     * 解析 → 底线契约断言 → 拼一份可读报告 → 打到控制台并存一份 UTF-8 文件。
     */
    private static void print(String path, WordParser.Options options, String tag) throws Exception {
        File file = new File(path);
        assertTrue(file.isFile(), "文档不存在：" + file.getAbsolutePath());

        Map<String, Object> result = parse(path, options);
        TestSupport.assertHeadersMatchRowKeys(result);          // 契约破了，下面的输出也没有参考价值
        String json = WordParser.toJson(result);
        assertNotNull(json, "toJson 返回 null");

        String report = report(file, options, tag, result, json);
        dump(fileName(file, tag), report);
    }

    /**
     * 把一种模式的 JSON 单独落一份 UTF-8 文件，返回该文件。
     */
    private static File writeModeJson(File sample, Mode mode, String json) throws Exception {
        File dir = new File(OUT_DIR, "json");
        assertTrue(dir.isDirectory() || dir.mkdirs(), "创建输出目录失败：" + dir.getAbsolutePath());
        File out = new File(dir, sample.getName() + "-" + mode.slug + ".json");
        Files.write(out.toPath(), json.getBytes(UTF_8));
        return out;
    }

    /**
     * 打到控制台，并在 {@code target/print} 存一份同名 UTF-8 副本（控制台按 GBK 编码时重定向到文件会花）。
     */
    private static void dump(String name, String report) throws Exception {
        System.out.println(report);
        assertTrue(OUT_DIR.isDirectory() || OUT_DIR.mkdirs(), "创建输出目录失败：" + OUT_DIR.getAbsolutePath());
        File copy = new File(OUT_DIR, name + ".txt");
        Files.write(copy.toPath(), report.getBytes(UTF_8));
        System.out.println("[PrintResultTest] 同一份输出的 UTF-8 文件：" + copy.getAbsolutePath());
    }

    /**
     * 按扩展名选解析器；三个入口的返回结构一致，所以报告与 {@code toJson} 可以共用一份。
     */
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

    /**
     * 输出文件名：{@code sales-report.docx-静默.txt}。tag 里的空格与路径分隔符换成下划线。
     */
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
        line(sb, "顶层键：" + result.keySet() + "（顺序即插入顺序）");
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
            line(sb, "#" + (t + 1) + "  " + headers.size() + " 列 × " + rows.size() + " 行"
                    + name(table) + "，来源 " + table.get("source"));
            line(sb, "键 = " + table.keySet());
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

        Object unmatched = result.get("unmatchedHeaders");
        if (unmatched instanceof List) {
            sb.append('\n');
            List<?> list = (List<?>) unmatched;
            line(sb, "---- unmatchedHeaders（" + list.size() + " 组，注册了却一组都没命中）----");
            for (Object item : list) {
                line(sb, "  " + item);
            }
        }

        sb.append('\n');
        line(sb, "---- JSON（" + json.length() + " 个字符）----");
        sb.append(json).append('\n');
        return sb.toString();
    }

    /**
     * 选项摘要：静默与否、注册了几组表头、分组行去向、图片模式。
     */
    private static String describe(WordParser.Options options) {
        StringBuilder sb = new StringBuilder();
        if (options.isSilent()) {
            sb.append("静默模式（未注册表头、未给关键字，允许首行兜底当表头）");
        } else {
            sb.append("注册了 ").append(options.headerGroupCount()).append(" 组表头");
            for (int i = 0; i < options.headerGroupCount(); i++) {
                sb.append("\n       第").append(i + 1).append("组 = ").append(options.headerTitles(i));
                String name = options.headerName(i);
                if (name.length() > 0) {
                    sb.append("，name = ").append(name);
                }
            }
        }
        sb.append("\n       FullRow = ").append(options.fullRow())
                .append("，ImageMode = ").append(options.images());
        if (options.imageDir() != null) {
            sb.append("，imageDir = ").append(options.imageDir().getAbsolutePath());
        }
        sb.append("，source = ").append(options.writesSource())
                .append("，unmatchedHeaders = ").append(options.writesUnmatchedHeaders());
        return sb.toString();
    }

    /**
     * 一组的注册名后缀：{@code ，name = 年度目标}；没起过名字时是空串，这一行与改动前完全一样。
     */
    private static String name(Map<String, Object> table) {
        Object name = table.get("name");
        return name == null ? "" : "，name = " + name;
    }

    /**
     * 一行数据按 {@code 列名=值} 拼出来，顺序就是 {@code headers} 的顺序（行对象是 {@code LinkedHashMap}）。
     */
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

    /**
     * 图片字段的可读摘要：{@code bytes} 只打长度，{@code base64} 只打长度与前缀，避免刷屏。
     */
    private static String imageSummary(Map<String, Object> image) {
        StringBuilder sb = new StringBuilder();
        for (Map.Entry<String, Object> field : image.entrySet()) {
            if (sb.length() > 0) {
                sb.append("；");
            }
            if (field.getValue() instanceof byte[]) {
                sb.append(field.getKey()).append("=byte[").append(((byte[]) field.getValue()).length).append(']');
            } else if ("base64".equals(field.getKey())) {
                String value = String.valueOf(field.getValue());
                sb.append("base64[").append(value.length()).append("] 前缀=")
                        .append(value.substring(0, Math.min(30, value.length()))).append("…");
            } else {
                sb.append(field.getKey()).append('=').append(field.getValue());
            }
        }
        return sb.toString();
    }

    /**
     * 把清单转成缩进好的 JSON；失败时退回 {@code String.valueOf}，这里只是打印，不该因为格式化失败中断。
     */
    private static String wordParserJson(Object value) {
        try {
            String json = WordParser.toJson(value);
            return json.replaceAll("(?m)^", "  ") + "\n";
        } catch (Exception failure) {
            return "  " + String.valueOf(value) + "\n";
        }
    }

    /**
     * 按行切开；空串（例如 PDF 样例的 {@code text}）不给任何行，而不是给一个空行。
     */
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

    /**
     * 把捕获到的标准错误原样附到报告里（每行缩进四格）。
     */
    private static void appendStderr(StringBuilder sb, String stderr) {
        appendStderr(sb, stderr, Integer.MAX_VALUE);
    }

    /**
     * 只附前 {@code maxLines} 行非空标准错误；空内容打「（空）」。
     */
    private static void appendStderr(StringBuilder sb, String stderr, int maxLines) {
        String[] lines = stderr == null ? new String[0] : stderr.split("\n");
        int printed = 0;
        for (String line : lines) {
            if (line.trim().length() == 0) {
                continue;
            }
            if (printed >= maxLines) {
                line(sb, "  stderr：…（还有更多，已截断）");
                break;
            }
            line(sb, "  stderr：" + line);
            printed++;
        }
        if (printed == 0) {
            line(sb, "  stderr：（空）");
        }
    }

    /**
     * 异常消息压成一行，太长时截到 300 字符 —— 这里只看类型与关键信息。
     */
    private static String oneLine(String message) {
        if (message == null) {
            return "（null）";
        }
        String value = message.replace('\r', ' ').replace('\n', ' ');
        return value.length() <= 300 ? value : value.substring(0, 300) + "…";
    }

    /**
     * 目录里的文件名（排序后），目录不存在时给空列表。用来比较一次解析前后是否多出文件。
     */
    private static List<String> listFiles(File dir) {
        String[] names = dir.list();
        if (names == null) {
            return Collections.emptyList();
        }
        List<String> sorted = new ArrayList<String>(Arrays.asList(names));
        Collections.sort(sorted);
        return sorted;
    }

    /**
     * 删除目录及其内容；不存在时直接返回。目标只在 {@code target/print} 下，{@code mvn clean} 也会带走。
     */
    private static void deleteRecursively(File file) {
        if (file == null || !file.exists()) {
            return;
        }
        File[] children = file.listFiles();
        if (children != null) {
            for (File child : children) {
                deleteRecursively(child);
            }
        }
        // 删不掉也不影响断言：uniqueFile 会换名，报告里能看到实际路径
        //noinspection ResultOfMethodCallIgnored
        file.delete();
    }

    /**
     * 允许 {@link #probe} 接收会抛异常的动作。
     */
    private interface ThrowingRunnable {
        void run() throws Exception;
    }
}
