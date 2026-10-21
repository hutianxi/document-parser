/**
 * 文档解析模块：把 Word / PDF / PPT 解析成同样的三个集合 —— {@code text}、{@code tables}、{@code images}。
 *
 * <p>入口是三个类，方法签名一一对应，换格式只换类名：
 * <ul>
 *   <li>{@link cn.net.tianxi.document.WordParser} —— doc / docx；</li>
 *   <li>{@link cn.net.tianxi.document.PdfParser} —— pdf；</li>
 *   <li>{@link cn.net.tianxi.document.PptParser} —— ppt / pptx。</li>
 * </ul>
 *
 * <pre>{@code
 * Map<String, Object> result = WordParser.parse("D:/a.docx", "列");
 * String json = WordParser.toJson(result);
 * }</pre>
 *
 * <h2>返回结构</h2>
 * 下面是 {@code src/test/resources/word/sales-report.docx} 的实测输出（默认 {@code ImageMode.META}，
 * 所以图片那一项只有元信息）：
 * <pre>
 * {
 *   "text": "2025 年度华东区销售业绩报告\n报告页次：3 页\n归档编号：\n…\n以上数据截至 2025 年 12 月 31 日，最终以财务系统结账数为准。",
 *   "tables": [
 *     { "headers": ["片区","区域","销售额","回款率","销售额-5","回款率-6","备注"],
 *       "rows": [ {"片区":"华东大区","区域":"上海","销售额":"1280","回款率":"92%",
 *                  "销售额-5":"1465","回款率-6":"95%","备注":"超额完成"} ] }
 *   ],
 *   "images": [ {"fileName":"image-1.png","type":"PNG","width":1024,"height":768,"size":670939,
 *                "contentId":"487e915db2c53907e8c1551aeaffea235423210c"} ]
 * }
 * </pre>
 * 这份样例静默模式下的实测规模：正文 13 段、7 个表头组共 30 行、1 张图片；
 * {@code toJson} 出来 4409 个字符，改用 {@code ImageMode.BASE64} 是 899039 个字符。
 * 同源的 {@code sales-report.pdf}（3 页）与 {@code sales-report.pptx}（3 页）规模见各自的解析器类说明。
 *
 * <h2>表格：一个表头对应它下面的行</h2>
 * 一张物理表里有几组表头，{@code tables} 里就有几个元素，每个元素的 {@code rows} 只包含属于它上面那个表头的行，
 * 所以不会串组。表头行的判定优先级（任一命中即视为表头）写在 {@link cn.net.tianxi.document.Tables}：
 * 注册的完整表头集合 &rarr; 关键字 &rarr; 格式自带的标题行标记 &rarr; 紧跟跨整行合并行之后的行 &rarr;
 * 表样式声明的首行 &rarr; 什么都没给时的第一行。
 *
 * <p>合并单元格一律「拆开」：横向、纵向合并覆盖的位置都回填起始格的值，于是每行的键数量一致、合并里填的内容不丢。
 * {@code headers} 与每一行的键逐个相等（表头是空格的列补成 {@code 列N}，同名列补成 {@code 名-列号}），
 * 因此可以只读 {@code headers} 当列清单去建表或映射字段。
 * 跨整行的合并行（分组小标题，例如「2025 年度」）由
 * {@link cn.net.tianxi.document.WordParser.Options#fullRow(cn.net.tianxi.document.WordParser.FullRow)
 * Options.fullRow(FullRow)} 决定保留还是丢弃。Word 里单元格内再套一张表时，子表不进母表格子的文本，
 * 作为 {@code tables} 里额外的元素排在母表分组之后。
 *
 * <h2>读到什么、读不到什么</h2>
 * <ul>
 *   <li><b>Word</b>：各节 Body 的段落、表格、图片，文本框文字与块级内容控件（{@code w:sdt}）包住的段落表格都能读到；
 *       控件没填写时它自带的占位文字（{@code Click here to enter text.}）不会进 {@code text} ——
 *       那句提示存在 {@code w:sdtPr/w:placeholderText} 里，属于 Word 的显示层，不是文档内容；
 *       页眉页脚、脚注尾注、批注不读（版式附属内容，评估水印就注入在那里）；</li>
 *   <li><b>PDF</b>：有线段边框的表格靠几何还原，跨页断开的长表并回一张（跨三页以上也是一路并下去；
 *       没注册表头依据时靠「一截贴页底、一截贴页首」认定是同一张表被切断，所以上一页最后那一行很高时会被拆成两组）；
 *       <u>没有边框线的表格</u>识别不出来，
 *       会被当成普通段落进 {@code text}；扫描件／图片型 PDF 没有文字层，需要 OCR；</li>
 *   <li><b>PPT</b>：正片各页的形状文字、表格、图片（组合形状递归展开；视频帧走图片那一支，收的是封面图）；
 *       备注页／母版、SmartArt 图形内的文字、图表数据、嵌入的 OLE 对象里的内容、形状填充底图不读。</li>
 * </ul>
 *
 * <h2>文本清洗</h2>
 * 三种格式的文本都过 {@link cn.net.tianxi.document.Results} 的 {@code clean} 一道：Word 用三段控制字符把域包成
 * 「域代码 + 域结果」，丢掉代码只留结果（否则单元格里会多出 {@code PAGE}，还可能被当成表头）；软换行、分页符
 * 换成空格而不是丢弃（丢了会把 Shift+Enter 分的两行粘成一句）；单元格结束符等不可见控制字符丢弃，连续空格压成一个。
 * 域结果没算过时是空的 —— 本模块不做版面计算，不调 {@code updateFields}。
 *
 * <h2>图片</h2>
 * 元信息任何模式都给；内容去向由 {@link cn.net.tianxi.document.WordParser.ImageMode} 选：
 * 只要元信息（默认，解析出来主要是为了入库，没必要为图片付体积和内存）、base64 内嵌、直接落盘只留路径、
 * 原始字节。内嵌模式下内容相同的图片用 {@code contentId}（SHA-1 指纹）标 {@code duplicateOf}，不重复内嵌。
 *
 * <h2>许可证与评估模式</h2>
 * 底层是 Aspose.Words / Aspose.PDF / Aspose.Slides。把 {@code aspose-words.lic}、{@code aspose-pdf.lic}、
 * {@code aspose-slides.lic} 放在 classpath 根目录或进程工作目录，各产品在一个进程内只注册一次，且都在
 * 加载文档<u>之前</u>注册（晚注册对已加载的文档无效）。没有 lic 时按评估模式解析，三个产品的限制并不相同：
 * <ul>
 *   <li><b>Words</b>：往页眉页脚注入水印段落。样例 docx 的正文 13 段里没有一句水印，
 *       因为页眉页脚本来就不读，水印段落走 {@code Results.cleanContent} 被过滤；</li>
 *   <li><b>Slides</b>：每个文本框只给第一段，并截到 5 个字符，后面接
 *       {@code ... text has been truncated due to evaluation version limitation.}（表格单元格同样受影响），
 *       所以样例 pptx 实测只读到 9 行文本，且每行都很短 —— 这部分内容回不来；</li>
 *   <li><b>PDF</b>：受限的是<u>文字层</u>（每页只给出前若干段），
 *       {@code TableAbsorber} 读到的表格行与图片<b>不截断</b>：样例 sales-report.pdf 在评估模式下
 *       读到了完整的 6 个表头组、45 行、1 张图片。样例的 {@code text} 是空串（段落都落在表格矩形内被跳过），
 *       所以文字层这条限制在这份样例上观察不到。</li>
 * </ul>
 * 评估模式留下的痕迹统一过滤：{@link cn.net.tianxi.document.Results#cleanContent} 负责去掉水印文字
 * （正文段落、表格单元格、幻灯片文本框都过这一道，被换行切成孤行的半截水印也能识别），
 * {@code PptParser} 的 {@code stripEvaluationNotice} 负责去掉截断提示。
 * 但<u>被截断的内容回不来</u>，所以评估模式的输出只能当样例看，不能当结果用。
 *
 * <h2>安全与边界</h2>
 * <ul>
 *   <li>解析的是外部文档，输出的 {@code fileName}、{@code path} 一律经过处理：写文件前只取文件名部分
 *       （{@link cn.net.tianxi.document.Results} 的 {@code safeName}），目录里重名再加序号，
 *       所以文档自带的怪名字不可能把文件写到目标目录之外；</li>
 *   <li>外链图片（Word「链接到文件」）只回报路径，本模块不会去读那个文件，也不会联网；</li>
 *   <li>{@code contentId} 是内容指纹，只为去重和缓存服务，不是防篡改摘要；</li>
 *   <li>递归深度都有上限，超过就跳过并提示，避免畸形文件把栈撑爆：PPT 组合形状 32 层
 *       （{@link cn.net.tianxi.document.PptParser}），Word Body 之下容器节点 16 层、表格嵌套 8 层
 *       （{@link cn.net.tianxi.document.WordParser}）；</li>
 *   <li>大文件与内存：base64/bytes 模式会把图片字节都留在结果里，图片多的文档请用 {@code FILE} 或
 *       {@code META}；OOXML（docx/pptx）是压缩包，解压后全量进内存，压缩比极高的畸形文件可能直接把进程撑爆，
 *       所以上传/落盘环节要有大小上限，解析调用最好带超时与独立线程池；</li>
 *   <li>三个解析器都是无状态静态方法，可以并发调用；{@code Options} 本身可变态，配置期间不要跨线程共享。</li>
 * </ul>
 *
 * <h2>运行环境</h2>
 * Java 8（Aspose 用的是 {@code jdk16} classifier 的构件），无头服务器记得加
 * {@code -Djava.awt.headless=true}。三个解析类都是纯静态方法，不依赖任何容器，也没有命令行入口：
 * 单独验证一个文档就在测试类里给个路径调用即可，例如
 * {@code WordParser.parse("src/test/resources/word/sales-report.docx", WordParser.options().keywords("销售"))}。
 */
package cn.net.tianxi.document;
