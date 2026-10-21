package cn.net.tianxi.document;

import java.io.File;
import java.io.OutputStream;
import java.security.MessageDigest;
import java.util.ArrayList;
import java.util.Base64;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Locale;
import java.util.Map;
import java.util.regex.Pattern;

/**
 * 图片集合的共用收集器：统一文件名、元信息、内容指纹、重复图片处理和四种 {@link WordParser.ImageMode} 的输出。
 *
 * <p>三种格式的来源不同（Word 的 {@code ImageData}、PDF 的 {@code XImage}、PPT 的 {@code IPPImage}），
 * 但结果里每条的形状一样，所以只在这里实现一份。下面是三份样例的实测条目（默认 {@code META} 模式）：
 * <pre>
 * // sales-report.docx 与 sales-report.pptx 里是同一份 PNG 字节，所以指纹相同
 * {"fileName":"image-1.png","type":"PNG","width":1024,"height":768,
 *  "size":670939,"contentId":"487e915db2c53907e8c1551aeaffea235423210c"}
 *
 * // sales-report.pdf：Word 导出时把那张图重编成了 JPEG，解析照实报告文件里的字节，不做还原
 * {"fileName":"image-1.jpg","type":"JPEG","width":699,"height":524,
 *  "size":38668,"contentId":"80c8e65be48c09a0f173845e4dfc072bc541fb68"}
 * </pre>
 *
 * <h2>四种模式的实测体积</h2>
 * {@code META}（默认）只有上面这些元信息；{@code BASE64} 多一个 {@code "base64":"data:image/png;base64,..."}，
 * {@code BYTES} 多一个 {@code "bytes"}（原始 {@code byte[]}，Jackson 会编成不带 {@code data:} 前缀的 base64），
 * {@code FILE} 多一个 {@code "path"}（落盘后的绝对路径）且结果里不留内容。整串 JSON 的长度：
 * docx 4409 → 899039、pdf 6186 → 57789、pptx 3262 → 897892 ——
 * docx 那张 670939 字节的 PNG 编码后是 894610 个字符，占了 99.5%，这就是默认不给内容的原因。
 *
 * <h2>文件名与类型</h2>
 * 按文档流顺序编号 {@code image-1}、{@code image-2}……扩展名跟着实际类型走（PNG → {@code .png}、
 * JPEG → {@code .jpg}，认不出来的类型 → {@code .bin}）。类型串来自文档自带的 content-type，
 * 先去掉 {@code image/} 这类前缀，再把第一个不是字母/数字的位置往后全部丢掉 ——
 * {@code png/../../evil} 会被剪成认不出的类型，所以文档自带的怪名字不可能让 {@code fileName} 带上路径。
 * 类型为空时按字节的魔数嗅探（PNG/JPEG/GIF/BMP/TIFF/EMF/WMF 都认得），嗅不出来记 {@code UNKNOWN}。
 *
 * <h2>重复图片</h2>
 * 只有内嵌模式（{@code BASE64}/{@code BYTES}）做去重：内容相同的第二张不重复内嵌，
 * 只留一个 {@code "duplicateOf": 0} 指向前一条的下标，{@code contentId} 仍是它自己的指纹。
 * {@code META} 本来就不带内容，不标去重；{@code FILE} 模式下每张图都要落成自己的文件，
 * 重名由 {@link #uniqueFile} 加 {@code -2}、{@code -3} 后缀，绝不覆盖已有文件。
 * 三份样例各只有一张图，去重这条由 {@code ImagesTest} 直接喂字节验证。
 *
 * <h2>外链图片与配错选项</h2>
 * Word「链接到文件」、PPT「仅链接」的图片字节不在文档包里，{@link #link} 只回报元信息与 {@code link} 路径，
 * 没有 {@code size}/{@code contentId}，本模块不去读那个文件、也不联网。
 * {@code FILE} 模式没配 {@link WordParser.Options#imageDir(java.io.File)} 会在解析到第一张图时抛
 * {@code IllegalStateException}（静默失败最难受：解析跑完了、图片一张也没落地）；
 * 反过来配了 {@code imageDir} 却不是 {@code FILE} 模式，构造时在标准错误提示且一个文件也不写。
 *
 * <p>两点使用上的约定：
 * <ul>
 *   <li>一个实例只服务一次解析，非线程安全（{@code index}、{@code firstSeen} 都是普通字段），
 *       不要在多个文档之间复用同一个 {@code Images}；</li>
 *   <li>{@code contentId} 是内容指纹，用 SHA-1 只为「同一份字节认成同一张图」和让调用方按内容缓存，
 *       不承担防篡改职责，别拿它当安全摘要。</li>
 * </ul>
 *
 * <p><i>测试：{@code ImagesTest}（24 条，直接喂字节，不经过 Aspose）+ 三个解析器测试里的端到端断言。</i>
 */
final class Images {

    /**
     * 类型串里第一个不是字母/数字的位置往后全部丢掉。
     *
     * <p>类型来自文档自带的 content-type，可能写成 {@code png/../../evil} 这种；它会被拼进
     * {@code fileName}，所以先把分隔符和点号剪掉，免得给调用方一个带路径的文件名。
     */
    private static final Pattern TYPE_TAIL = Pattern.compile("[^a-z0-9].*$");

    private final WordParser.Options options;
    private final List<Map<String, Object>> entries = new ArrayList<Map<String, Object>>();
    /** 内容指纹 -&gt; 该内容第一次出现时在 {@code entries} 里的下标。 */
    private final Map<String, Integer> firstSeen = new LinkedHashMap<String, Integer>();
    private int index;

    Images(WordParser.Options options) {
        this.options = options;
        if (options.imageDir() != null && options.images() != WordParser.ImageMode.FILE) {
            System.err.println("[WordParser] imageDir 只在 ImageMode.FILE 下生效，当前是 " + options.images()
                    + "，" + options.imageDir() + " 不会写入图片");
        }
    }

    List<Map<String, Object>> list() {
        return entries;
    }

    /**
     * 记录一张嵌入图片。
     *
     * @param type   图片类型名（PNG/JPEG/…），传 null 时按字节嗅探
     * @param width  像素宽，未知传 0
     * @param height 像素高，未知传 0
     */
    void add(String type, int width, int height, byte[] bytes) throws Exception {
        if (bytes == null) {
            return;
        }
        String normalized = normalizeType(type, bytes);
        Map<String, Object> image = meta(normalized, width, height);
        image.put("size", Integer.valueOf(bytes.length));
        String contentId = contentId(bytes);
        image.put("contentId", contentId);
        if (options.isInlineImageMode()) {
            // 内嵌模式下内容重复的图片不再重复内嵌，只指向前一条
            Integer duplicateOf = firstSeen.get(contentId);
            if (duplicateOf != null) {
                image.put("duplicateOf", duplicateOf);
                entries.add(image);
                return;
            }
            firstSeen.put(contentId, Integer.valueOf(entries.size()));
        }
        putContent(image, bytes, normalized);
        entries.add(image);
    }

    /** 记录一张外链图片（Word 里"链接到文件"、PPT 里"仅链接"）：字节不在文档内，只给路径。 */
    void link(String type, int width, int height, String source) {
        Map<String, Object> image = meta(normalizeType(type, null), width, height);
        image.put("link", source);
        entries.add(image);
    }

    private Map<String, Object> meta(String type, int width, int height) {
        Map<String, Object> image = new LinkedHashMap<String, Object>();
        image.put("fileName", "image-" + (++index) + "." + extension(type));
        image.put("type", type);
        image.put("width", Integer.valueOf(width));
        image.put("height", Integer.valueOf(height));
        return image;
    }

    /** 按 {@link WordParser.ImageMode} 决定内容放哪儿；{@code META} 模式什么都不放。 */
    private void putContent(Map<String, Object> image, byte[] bytes, String type) throws Exception {
        if (options.images() == WordParser.ImageMode.BASE64) {
            image.put("base64", "data:" + contentType(type) + ";base64," + Base64.getEncoder().encodeToString(bytes));
        } else if (options.images() == WordParser.ImageMode.BYTES) {
            image.put("bytes", bytes);
        } else if (options.images() == WordParser.ImageMode.FILE) {
            File directory = options.imageDir();
            if (directory == null) {
                throw new IllegalStateException("ImageMode.FILE 需要先调用 imageDir(File) 指定导出目录");
            }
            File file = uniqueFile(directory, String.valueOf(image.get("fileName")));
            write(file, bytes);
            image.put("fileName", file.getName());       // 实际落盘的名字可能加了序号，结果里要写真实的那个
            image.put("path", file.getAbsolutePath());
        }
    }

    /**
     * 在目录里给文件名找一个不被占用的名字。
     *
     * <p>图片一律叫 {@code image-N.扩展名}，每次解析都从 1 开始编号。批量解析多个文档、又写到同一个目录时，
     * 第二个文档的 {@code image-1.png} 会直接把第一个文档的同名文件覆盖掉 —— 而且结果里看不出出了问题。
     * 这里遇到重名就加序号（{@code image-1-2.png}），保证谁的图片都不丢。
     *
     * <p>「先看存不存在、再创建」中间有个窗口：两个线程同时往同一个目录写还可能撞上同一个名字。要在并发场景
     * 下导出，请给每次解析一个自己的 {@link WordParser.Options#imageDir(File)}，或者用 {@code CREATE_NEW}
     * 那套 NIO 写入自己控制。
     */
    static File uniqueFile(File directory, String fileName) {
        String name = Results.safeName(fileName);
        File file = new File(directory, name);
        if (!file.exists()) {
            return file;
        }
        int dot = name.lastIndexOf('.');
        String stem = dot < 0 ? name : name.substring(0, dot);
        String tail = dot < 0 ? "" : name.substring(dot);
        for (int i = 2; i < Integer.MAX_VALUE; i++) {
            file = new File(directory, stem + "-" + i + tail);
            if (!file.exists()) {
                return file;
            }
        }
        throw new IllegalStateException("目录里重名图片过多: " + directory);
    }

    static void write(File file, byte[] bytes) throws Exception {
        File parent = file.getParentFile();
        if (parent != null && !parent.exists() && !parent.mkdirs()) {
            throw new java.io.IOException("无法创建目录: " + parent);
        }
        OutputStream out = new java.io.FileOutputStream(file);
        try {
            out.write(bytes);
        } finally {
            out.close();
        }
    }

    /** 内容指纹：SHA-1 十六进制，用于识别重复图片、也让调用方能按内容缓存。 */
    static String contentId(byte[] bytes) throws Exception {
        byte[] digest = MessageDigest.getInstance("SHA-1").digest(bytes);
        StringBuilder hex = new StringBuilder(digest.length * 2);
        for (byte b : digest) {
            hex.append(Character.forDigit((b >> 4) & 0xF, 16)).append(Character.forDigit(b & 0xF, 16));
        }
        return hex.toString();
    }

    /** 类型名统一成大写：{@code image/jpeg}、{@code jpg}、{@code JPEG} 都会变成 {@code JPEG}。 */
    private static String normalizeType(String type, byte[] bytes) {
        String value = type == null ? "" : type.trim().toLowerCase(Locale.ROOT);
        int slash = value.indexOf('/');
        if (slash >= 0) {
            value = value.substring(slash + 1);                 // image/png -> png
        }
        int semicolon = value.indexOf(';');
        if (semicolon >= 0) {
            value = value.substring(0, semicolon);               // 丢掉 image/png;charset=... 的参数
        }
        value = TYPE_TAIL.matcher(value).replaceAll("");
        if (value.length() == 0 || "jpg".equals(value) || "jpeg".equals(value)) {
            String sniffed = sniff(bytes);                       // 类型缺失或是 JPEG 别名时按文件头确认
            if (sniffed != null) {
                return sniffed;
            }
        }
        return value.length() == 0 ? "UNKNOWN" : value.toUpperCase(Locale.ROOT);
    }

    /** 按文件头魔数判断类型，PDF 的 XImage 拿不到干净的类型名时用这个。 */
    static String sniff(byte[] bytes) {
        if (bytes == null || bytes.length < 4) {
            return null;
        }
        if ((bytes[0] & 0xFF) == 0x89 && (bytes[1] & 0xFF) == 0x50) {
            return "PNG";
        }
        if ((bytes[0] & 0xFF) == 0xFF && (bytes[1] & 0xFF) == 0xD8) {
            return "JPEG";
        }
        if ((bytes[0] & 0xFF) == 0x47 && (bytes[1] & 0xFF) == 0x49) {
            return "GIF";
        }
        if ((bytes[0] & 0xFF) == 0x42 && (bytes[1] & 0xFF) == 0x4D) {
            return "BMP";
        }
        if ((bytes[0] & 0xFF) == 0x49 && (bytes[1] & 0xFF) == 0x49) {
            return "TIFF";
        }
        if ((bytes[0] & 0xFF) == 0x4D && (bytes[1] & 0xFF) == 0x4D) {
            return "TIFF";
        }
        if (bytes.length > 44 && (bytes[0] & 0xFF) == 0x01 && (bytes[1] & 0xFF) == 0x00
                && (bytes[2] & 0xFF) == 0x00 && (bytes[3] & 0xFF) == 0x00) {
            return "EMF";
        }
        if ((bytes[0] & 0xFF) == 0xD7 && (bytes[1] & 0xFF) == 0xCD && (bytes[2] & 0xFF) == 0xC6) {
            return "WMF";
        }
        return null;
    }

    /** 图片扩展名：未知类型用 bin。 */
    static String extension(String type) {
        if ("JPEG".equals(type)) {
            return "jpg";
        }
        return "UNKNOWN".equals(type) ? "bin" : type.toLowerCase();
    }

    static String contentType(String type) {
        return "UNKNOWN".equals(type) ? "application/octet-stream" : "image/" + type.toLowerCase(Locale.ROOT);
    }
}
