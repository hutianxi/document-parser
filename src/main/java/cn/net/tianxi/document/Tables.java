package cn.net.tianxi.document;

import java.util.ArrayList;
import java.util.Arrays;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;
import java.util.Set;

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
 *
 * <p><i>测试：{@code TablesTest}（31 条，直接喂网格，不经过 Aspose）+ 三个解析器测试里的端到端断言。</i>
 */
final class Tables {

    private Tables() {
    }

    /**
     * @param rows         行网格，每行长度 = 列数；合并覆盖的位置要提前回填起始格的值
     * @param separator    每行是否跨整行合并（小标题、分组行），长度与 {@code rows} 相同
     * @param heading      每行是否被格式标记为标题行；格式本身没有这种标记时传 null
     * @param matched      出参：本次实际命中的表头集合下标，供调用方提示未命中的集合
     * @param firstRowStyled 表样式是否声明首行为标题（Word 的 {@code tblLook firstRow}）
     * @return 每组 = {@code {headers: [列名], rows: [{列名: 值}]}}；{@code headers} 与每一行的键逐个相等
     *         （空表头补成 {@code 列N}，同名列补成 {@code 名-列号}），所以调用方可以只读 {@code headers} 拿列清单
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
                    block.put("headers", blockHeaders);
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

    /** 单列列名：表头为空用序号兜底，与前面的列重名时加列号后缀（合并单元格摊平后就会有重名）。 */
    private static String columnName(List<String> used, String title, int index) {
        String name = title.length() > 0 ? title : "列" + (index + 1);
        return used.contains(name) ? name + "-" + (index + 1) : name;
    }

    /** 一行数据 = 表头列名 -&gt; 单元格内容；表头没给到名字的那一列按 {@link #normalizeHeaders} 的同一条规则补，保证键与 headers 对齐。 */
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

    /** 多级表头：低层级缺名时用高层级同列的名称补上（纵向合并的列名）。 */
    static List<String> mergeHeaders(List<String> upper, List<String> lower) {
        List<String> merged = new ArrayList<String>(lower);
        for (int c = 0; c < merged.size(); c++) {
            if (merged.get(c).length() == 0 && c < upper.size()) {
                merged.set(c, upper.get(c));
            }
        }
        return merged;
    }

    /** 跨整行的合并行只保留一个值，其余列留空。 */
    static List<String> single(String value, int columnCount) {
        String[] cells = new String[columnCount];
        Arrays.fill(cells, "");
        cells[0] = value;
        return Arrays.asList(cells);
    }

    /** 跨整行合并行的文本：取第一个非空格，避免合并起点不在第一列时取错。 */
    static String firstNonBlank(List<String> cells) {
        for (String cell : cells) {
            if (cell != null && cell.trim().length() > 0) {
                return cell;
            }
        }
        return "";
    }

    /** 把长短不一的行补齐到同一列数，保证每行键数量一致。 */
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
