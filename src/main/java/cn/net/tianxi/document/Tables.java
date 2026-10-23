package cn.net.tianxi.document;

import java.util.*;

/**
 * 表头分组的共用算法：把一张表的行网格切成「一个表头 + 归属它的行」。
 *
 * <p>Word、PDF、PPT 三种格式各自负责把自己的表格读成 {@code String[][]} 网格（合并单元格的位置必须已经
 * 回填成起始格的值，并标出哪些行是跨整行的合并行），分组规则只在这里实现一份，所以三种格式的行为必然一致。
 *
 * <h2>表头行的判定优先级</h2>
 * 任一命中即视为表头：
 * <ol>
 *   <li>命中 {@link WordParser.Options#headers(String...)} 注册的完整表头集合 —— 最优先，
 *       并且列名直接采用注册的名字，不用文档里读出来的那一行；</li>
 *   <li>命中 {@link WordParser.Options#keywords(String...)} 关键字（行内每个非空单元格都要含关键字之一）；</li>
 *   <li>格式自带的标题行标记（Word 的 {@code w:tblHeader}）；</li>
 *   <li><b>紧跟在跨整行合并行之后的行</b> —— 分组小标题后面通常是新一段的列名；</li>
 *   <li>表样式声明首行为标题时的第一行（Word 的 {@code tblLook firstRow}、PPT 的 {@code getFirstRow()}）；</li>
 *   <li>调用方没给任何表头依据、且整表没有标题标记时，第一行兜底。</li>
 * </ol>
 *
 * <p>第 6 条只在「静默模式」（{@link WordParser.Options#isSilent()}：既没注册表头、也没给关键字）下生效，
 * 给了依据就以依据为准，免得把一行数据当表头 —— PDF 没有标题行标记可依赖，这条在它身上最明显。
 * 第 4 条有个容易踩的后果：分组行之后直接放数据行，那行数据会被当成表头、多切出一组。
 *
 * <h2>样例上的效果</h2>
 * {@code sales-report.docx} 的 4 张物理表出来是 7 组，每一组各压中一条规则：
 * <ul>
 *   <li>季度销售明细：第 5 条（表样式声明首行为标题），列名 {@code ["片区","区域","销售额","回款率",
 *       "销售额-5","回款率-6","备注"]} —— 两组同名列由 {@link #normalizeHeaders} 补上列号后缀；</li>
 *   <li>客户回款台账：两行连续表头且中间没有数据行，由 {@link #mergeHeaders} 合并成
 *       {@code ["客户名称","合同金额","已回款","回款率","备注"]}（下层空着的列沿用上层名字）；</li>
 *   <li>售后工单统计：母表一组，单元格里的嵌套小表 {@code ["故障部件","工单数"]} 另成一组；</li>
 *   <li>分片区年度目标：第一行就是跨整行合并的 {@code 2025 年度（已结账）}，此时还没有任何表头，
 *       于是另起一组、列名补成 {@code ["列1","列2","列3","列4"]}（否则这行文字会无声丢掉）；
 *       接着 {@code ["片区","年度目标","完成率","列4"]}（第四列在文档里没标题）；
 *       表中间的 {@code 2026 年度（预算下达）} 结束上一组、开出第三组 —— 同一张物理表里两组表头。</li>
 * </ul>
 *
 * <p>跨整行合并的行本身按 {@link WordParser.Options#fullRow(WordParser.FullRow)} 决定去向：
 * {@code KEEP} 留在当前组（docx 样例 7 组／30 行）、{@code FIRST} 归到同一张物理表的第一组
 * （组数行数不变，只改归属）、{@code DROP} 丢弃（6 组／28 行）。
 * 它同时标志着「上面那组的表头区结束了」：即使被丢掉，紧跟其后的行也不会被并进上一层表头。
 *
 * <h2>结果契约</h2>
 * 每组 = {@code {headers: [列名], rows: [{列名: 值}]}}，{@code headers} 与每一行的键<b>逐个相等</b>
 * （空表头补成 {@code 列N}、同名列补成 {@code 名-列号}，两处规则由 {@link #normalizeHeaders} 与
 * {@link #rowMap} 共用同一个 {@link #columnName}），所以调用方可以只读 {@code headers} 拿列清单去建表或映射字段。
 * 长短不一的行由 {@link #padded} 补齐到整表最大列数，保证每行键数量一致。
 * 分组本身只有这两个键；注册时 {@link WordParser.Options#name(String)} 起过名字的那一组多一个 {@code name}
 * （写在 {@link #group} 命中注册表头的那一刻，见 {@link #markName}）；来源位置 {@code source} 由解析器在分组之后
 * 用 {@link #markSource} 补进来（本网格里没有版面信息，也不该有）。
 * 四个键的先后固定为 {@code headers} → {@code rows} → {@code name} → {@code source}。
 *
 * <p><i>测试：{@code TablesTest}（34 条，直接喂网格，不经过 Aspose）+ 三个解析器测试里的端到端断言。</i>
 */
final class Tables {

    private Tables() {
    }

    /**
     * @param rows           行网格，每行长度 = 列数；合并覆盖的位置要提前回填起始格的值
     * @param separator      每行是否跨整行合并（小标题、分组行），长度与 {@code rows} 相同
     * @param heading        每行是否被格式标记为标题行；格式本身没有这种标记时传 null
     * @param matched        出参：本次实际命中的表头集合下标，供调用方提示未命中的集合
     * @param firstRowStyled 表样式是否声明首行为标题（Word 的 {@code tblLook firstRow}）
     * @return 每组 = {@code {headers: [列名], rows: [{列名: 值}]}}；{@code headers} 与每一行的键逐个相等
     * （空表头补成 {@code 列N}，同名列补成 {@code 名-列号}），所以调用方可以只读 {@code headers} 拿列清单；
     * 命中了带名字的注册表头时再多一个 {@code name} 键（{@link WordParser.Options#name(String)}，见 {@link #markName}）；
     * 来源位置不在这里，由调用方（各解析器）随后用 {@link #markSource} 补 {@code source} 键
     */
    static List<Map<String, Object>> group(String[][] rows, boolean[] separator, boolean[] heading,
                                           WordParser.Options options, Set<Integer> matched, boolean firstRowStyled) {
        int columnCount = 0;
        for (String[] row : rows) {
            columnCount = Math.max(columnCount, row.length);
        }

        // 整表没有任何标题标记、且调用方没给表头依据时，才允许首行兜底
        boolean noExplicitHeader = true;
        if (heading != null) {
            for (boolean flag : heading) {
                noExplicitHeader &= !flag;
            }
        }
        boolean noHeaderHint = options.isSilent() && noExplicitHeader;

        List<Map<String, Object>> blocks = new ArrayList<Map<String, Object>>();
        Map<String, Object> block = null;                 // 当前表头组
        List<String> blockHeaders = new ArrayList<String>();
        boolean blockHasData = false;                     // 当前组是否已经收过数据行
        for (int r = 0; r < rows.length; r++) {
            List<String> cells = padded(rows[r], columnCount);
            boolean separatorRow = separator[r];
            int group = separatorRow ? -1 : options.matchHeader(cells);
            boolean header = group >= 0 || (!separatorRow && (options.matchesKeywords(cells)
                    || (heading != null && heading[r])
                    || (r > 0 && separator[r - 1])
                    || (r == 0 && (firstRowStyled || noHeaderHint))));
            if (header) {
                if (group >= 0) {
                    matched.add(group);
                    // 列名用调用方注册的原名：PDF 单元格换行会让抽出来的文字多出空格，用注册名三种格式的键才一致
                    cells = padded(options.headerTitles(group).toArray(new String[0]), columnCount);
                }
                if (block == null || blockHasData) {
                    // 上面没有组、或当前组已收过数据 —— 开新的一组
                    block = newBlock(cells, columnCount);
                    blockHeaders = headersOf(block);
                    blocks.add(block);
                    blockHasData = false;
                } else {
                    // 紧接上一行表头且中间没有数据：多级表头，低层级空列沿用高层级列名
                    blockHeaders = normalizeHeaders(mergeHeaders(blockHeaders, cells), columnCount);
                    // 合并之后才等于注册列名的情况（docx 的两行表头：上下两行单独比都对不上，合并后逐字相同），
                    // 拿合并结果再比一次；命中就照单行命中的规矩换成注册名并写 name。
                    // 只在当前这一行本身没命中注册时才比 —— 命中的话名字由下面那句按「第一个名字为准」写，
                    // 这里不去抢，免得改掉多级表头两行各自命中不同注册时的既有优先级
                    if (group < 0) {
                        int mergedGroup = options.matchHeader(blockHeaders);
                        if (mergedGroup >= 0) {
                            matched.add(mergedGroup);
                            blockHeaders = normalizeHeaders(
                                    padded(options.headerTitles(mergedGroup).toArray(new String[0]), columnCount),
                                    columnCount);
                            markName(block, options.headerName(mergedGroup));
                        }
                    }
                    block.put("headers", blockHeaders);
                }
                if (group >= 0) {
                    markName(block, options.headerName(group));
                }
                continue;
            }
            if (separatorRow) {
                // 跨整行的合并行（小标题）由调用方决定去向；它同时结束了当前组的表头区
                if (options.fullRow() != WordParser.FullRow.DROP) {
                    Map<String, Object> target = options.fullRow() == WordParser.FullRow.FIRST && !blocks.isEmpty()
                            ? blocks.get(0) : block;
                    if (target == null) {
                        // 还没出现过表头：另起一组（列名用序号），否则这条分组行的文字就无声丢了
                        block = newBlock(new ArrayList<String>(), columnCount);
                        blockHeaders = headersOf(block);
                        blocks.add(block);
                        target = block;
                    }
                    rowsOf(target).add(rowMap(headersOf(target), single(firstNonBlank(cells), columnCount)));
                }
                blockHasData = true;
                continue;
            }
            // 数据行挂在上面最近的一组表头下
            if (block == null) {
                block = newBlock(blockHeaders, columnCount);   // 表头之前就出现的行，列名用序号兜底
                blockHeaders = headersOf(block);           // 补名之后，行对象的键与 headers 同源
                blocks.add(block);
            }
            rowsOf(block).add(rowMap(blockHeaders, cells));
            blockHasData = true;
        }
        return blocks;
    }

    static Map<String, Object> newBlock(List<String> headers, int columnCount) {
        Map<String, Object> block = new LinkedHashMap<String, Object>();
        block.put("headers", normalizeHeaders(headers, columnCount));
        block.put("rows", new ArrayList<Map<String, String>>());
        return block;
    }

    /**
     * 注册时给这一组起过名字才写 {@code name} 键，没起名时一个键都不多。
     *
     * <p>为什么在 {@link #group} 里面写、而 {@code source} 要在外面补：名字跟着「命中第几组注册表头」走，
     * 这个信息只有分组算法知道（{@link WordParser.Options#matchHeader} 的返回值），版面完全掺不进来；
     * 反过来页码只有各格式知道。两个键的写入点分开，分组算法的入参不会因为加键而变多。
     *
     * <p>同一组连续命中两个不同的注册集（多级表头且两行各自命中一处注册）时，以第一个名字为准、不覆盖 ——
     * 结果里不会出现「一行一个 name」的歧义，代价是第二个注册集的名字在这份文档上落不了地。
     * 三份样例都没有这种叠法（docx 台账那两行连续表头一行都对不上注册名，实测说明见 {@code TableByNameTest}），
     * 这条分支由 {@code TablesTest.consecutiveHeaderRowsMatchingTwoRegisteredSetsKeepTheFirstName} 直接喂网格覆盖。
     */
    private static void markName(Map<String, Object> block, String name) {
        if (name.length() > 0 && !block.containsKey("name")) {
            block.put("name", name);
        }
    }

    /**
     * 给这一组补来源位置：{@code source = {page: 起始页}}，跨页的表再多一个 {@code lastPage}。
     *
     * <p>三种格式共用同一个键名和同一个字段：docx 是 Aspose.Words {@code LayoutCollector} 给出的页码、
     * pdf 是页码、pptx 是幻灯片序号（都是 1 起）。回原文核对时按这个数字翻到那一页/那一页幻灯片即可。
     *
     * <p>为什么在 {@link #group} 之外单独补：分组只认网格，网格里没有版面信息；页码是各格式自己知道的
     * （docx 问 LayoutCollector、pdf 在 {@code Flow.page}、pptx 是幻灯片下标），所以由解析器在分组之后
     * 把这一份信息写进来，分组算法不因此多一个参数。
     *
     * <p>同一张物理表切出的多个表头组，来源都写这张表的起止页 —— 组是按表头切的，不是按页切的。
     * {@code page} 为 0 表示这一格式没给出位置（实测三份样例的每一组都不是 0，这条分支还没被触发过），
     * 此时整键不写，省得调用方把「第 0 页」当成一个真实页次去翻。
     *
     * <p>调用方（三个解析器）先问过 {@code Options.writesSource()}：
     * {@link WordParser.Options#source(boolean)} 传 {@code false} 时根本不会走到这里，三种格式的结果都不带
     * {@code source} 键。Word 侧还把这一判断放在创建 {@code LayoutCollector} 之前，所以关掉开关连版面计算都省了。
     */
    static void markSource(List<Map<String, Object>> groups, int page, int lastPage) {
        if (page <= 0) {
            return;
        }
        for (Map<String, Object> group : groups) {
            Map<String, Object> source = new LinkedHashMap<String, Object>();
            source.put("page", Integer.valueOf(page));
            if (lastPage > page) {
                source.put("lastPage", Integer.valueOf(lastPage));
            }
            group.put("source", source);
        }
    }

    /**
     * 注册了却一行都没命中的表头集合：提示到标准错误输出，同时收集成返回值交给调用方，
     * 三个解析器共用这一份规则（只有前缀不同）。
     *
     * <p>为什么要提示：这种注册写错了（错字、少列、列数不同）会让后面的数据行整段并进上一组表头，
     * 只从结果里很难看出来。提示只是线索、不是错误 —— 结果照给。
     *
     * <h3>哪些注册算「没命中」</h3>
     * 比对是「同一行内逐列相等」，多级表头另有一条：两行表头<b>合并之后</b>等于注册列名时也算命中
     * （{@link #group} 里合并完会拿合并结果再比一次），所以注册「合并后那一行」在三种格式都命得中。
     * 仍然命不中的典型是三种：写错字、列数不同（末尾空列在比对前被裁掉，容忍度放宽的是格子的值、从不放宽列数）、
     * 以及注册的是文档里真实的<b>上层</b>那一行而这个格式的网格里没有它（实测 pdf 就没有）。
     *
     * <p>同一个 {@code name} 之下只要有一份注册命中了，另一份命不中就<b>既不提示、也不进返回值</b> ——
     * 那一组表头已经找到了，再报就是噪声（{@link #coveredBySameName}）。没有起名字的注册不受这条影响，照样报。
     *
     * @param tag 日志前缀，例如 {@code [WordParser] }
     * @return 未命中的注册，每项是 {@code {headers: [注册时写的列名]}}，起过名字的另带 {@code name}；
     * 全部命中时返回空列表，<b>不返回 null</b>。要不要把它写进结果由
     * {@link WordParser.Options#unmatchedHeaders(boolean)} 决定（默认不写，本方法照样提示到标准错误）。
     * {@code headers} 与提示用的都是注册时写的<b>原文</b>（{@code headerTitles}），
     * 不是比对用的归一化形式（小写去空格），这样调用方能拿它直接去改自己的注册代码
     */
    static List<Map<String, Object>> warnUnmatchedHeaders(String tag, WordParser.Options options,
                                                          Set<Integer> matchedGroups) {
        List<Map<String, Object>> unmatched = new ArrayList<Map<String, Object>>();
        for (int i = 0; i < options.headerGroupCount(); i++) {
            if (matchedGroups.contains(i) || coveredBySameName(options, i, matchedGroups)) {
                continue;
            }
            System.err.println(tag + "表头集合未匹配到任何行：" + options.headerTitles(i));
            Map<String, Object> item = new LinkedHashMap<String, Object>();
            item.put("headers", new ArrayList<String>(options.headerTitles(i)));
            String name = options.headerName(i);
            if (name.length() > 0) {
                item.put("name", name);
            }
            unmatched.add(item);
        }
        return unmatched;
    }

    /**
     * 同一个 {@code name} 下已经有别的注册命中了吗（命中了就说明这一组表头已经找到了，不必再提示）。
     */
    private static boolean coveredBySameName(WordParser.Options options, int index, Set<Integer> matchedGroups) {
        String name = options.headerName(index);
        if (name.length() == 0) {
            return false;
        }
        for (int j = 0; j < options.headerGroupCount(); j++) {
            if (j != index && matchedGroups.contains(j) && name.equals(options.headerName(j))) {
                return true;
            }
        }
        return false;
    }

    /**
     * 列名补全到满列数：空表头用 {@code 列N}，同名再补列号后缀，规则与 {@link #rowMap} 是同一条。
     *
     * <p>这样 {@code headers} 里的每一项都等于各行对象里对应位置的键，调用方可以直接把 {@code headers}
     * 当列清单用（生成建表语句、映射实体字段、拼 CSV 表头），不用再猜「这一列会不会因为表头是空格而改叫 {@code 列3}」。
     * 代价有两处：列数少的表也会补齐到整表最大列数（那些键本来就存在于每一行，这里只是给它们名字）；
     * 表头是空的列拿到的是合成名 {@code 列N}，回原文核对时要按列号去找。
     */
    static List<String> normalizeHeaders(List<String> headers, int columnCount) {
        List<String> titles = padded(headers.toArray(new String[0]), columnCount);
        List<String> names = new ArrayList<String>(columnCount);
        for (int c = 0; c < titles.size(); c++) {
            names.add(columnName(names, titles.get(c), c));
        }
        return names;
    }

    /**
     * 单列列名：表头为空用序号兜底，与前面的列重名时加列号后缀（合并单元格摊平后就会有重名）。
     */
    private static String columnName(List<String> used, String title, int index) {
        String name = title.length() > 0 ? title : "列" + (index + 1);
        return used.contains(name) ? name + "-" + (index + 1) : name;
    }

    /**
     * 一行数据 = 表头列名 -&gt; 单元格内容；表头没给到名字的那一列按 {@link #normalizeHeaders} 的同一条规则补，保证键与 headers 对齐。
     */
    static Map<String, String> rowMap(List<String> headers, List<String> cells) {
        Map<String, String> row = new LinkedHashMap<String, String>();
        List<String> names = new ArrayList<String>(cells.size());
        for (int c = 0; c < cells.size(); c++) {
            String name = columnName(names, c < headers.size() ? headers.get(c) : "", c);
            names.add(name);
            row.put(name, cells.get(c));
        }
        return row;
    }

    /**
     * 多级表头：低层级缺名时用高层级同列的名称补上（纵向合并的列名）。
     */
    static List<String> mergeHeaders(List<String> upper, List<String> lower) {
        List<String> merged = new ArrayList<String>(lower);
        for (int c = 0; c < merged.size(); c++) {
            if (merged.get(c).length() == 0 && c < upper.size()) {
                merged.set(c, upper.get(c));
            }
        }
        return merged;
    }

    /**
     * 跨整行的合并行只保留一个值，其余列留空。
     */
    static List<String> single(String value, int columnCount) {
        String[] cells = new String[columnCount];
        Arrays.fill(cells, "");
        cells[0] = value;
        return Arrays.asList(cells);
    }

    /**
     * 跨整行合并行的文本：取第一个非空格，避免合并起点不在第一列时取错。
     */
    static String firstNonBlank(List<String> cells) {
        for (String cell : cells) {
            if (cell != null && cell.trim().length() > 0) {
                return cell;
            }
        }
        return "";
    }

    /**
     * 把长短不一的行补齐到同一列数，保证每行键数量一致。
     */
    static List<String> padded(String[] row, int columnCount) {
        String[] cells = new String[columnCount];
        Arrays.fill(cells, "");
        System.arraycopy(row, 0, cells, 0, Math.min(row.length, columnCount));
        for (int i = 0; i < cells.length; i++) {
            if (cells[i] == null) {
                cells[i] = "";
            }
        }
        return Arrays.asList(cells);
    }

    @SuppressWarnings("unchecked")
    static List<String> headersOf(Map<String, Object> block) {
        return (List<String>) block.get("headers");
    }

    @SuppressWarnings("unchecked")
    static List<Map<String, String>> rowsOf(Map<String, Object> block) {
        return (List<Map<String, String>>) block.get("rows");
    }
}
