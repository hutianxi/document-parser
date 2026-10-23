package cn.net.tianxi.document;

import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

import java.io.File;
import java.util.Arrays;
import java.util.Locale;
import java.util.Map;

import static org.junit.jupiter.api.Assertions.*;

/**
 * {@link Images} 的单元测试：文件名编号、类型归一化、内容指纹、重复图片、四种输出模式、外链图片。
 *
 * <p>图片集合是三种格式共用的唯一一份实现，这里直接喂字节断言，不经过 Aspose —— 想造「同一份字节出现两次」
 * 「类型串里带路径」这类输入，用真实文档反而费劲。端到端的图片断言在 {@code WordParserTest} /
 * {@code PdfParserTest} / {@code PptParserTest} 里。
 */
class ImagesTest {

    /**
     * PNG 文件头 + 几个字节。
     */
    private static final byte[] PNG = {(byte) 0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A, 1, 2, 3};
    /**
     * JPEG 文件头 + 几个字节（内容与 PNG 不同，指纹也不同）。
     */
    private static final byte[] JPEG = {(byte) 0xFF, (byte) 0xD8, (byte) 0xFF, (byte) 0xE0, 9, 8, 7};
    /**
     * 认不出来的字节：既不是图片魔数，类型串也给不出。
     */
    private static final byte[] OPAQUE = {1, 2, 3, 4, 5, 6};

    private static Images images(WordParser.ImageMode mode) {
        return new Images(WordParser.options().images(mode));
    }

    /* ---------------- 元信息与编号 ---------------- */

    @Test
    void metaModeCarriesMetadataOnly() throws Exception {
        Images images = images(WordParser.ImageMode.META);
        images.add("PNG", 1024, 768, PNG);

        Map<String, Object> image = images.list().get(0);
        assertEquals("image-1.png", image.get("fileName"));
        assertEquals("PNG", image.get("type"));
        assertEquals(1024, image.get("width"));
        assertEquals(768, image.get("height"));
        assertEquals(PNG.length, image.get("size"));
        assertEquals(40, String.valueOf(image.get("contentId")).length(), "contentId 是 SHA-1 十六进制");
        assertFalse(image.containsKey("base64"), "META 模式不带内容");
        assertFalse(image.containsKey("bytes"), "META 模式不带内容");
        assertFalse(image.containsKey("path"), "META 模式不落盘");
        TestSupport.assertImageMeta(image);
    }

    @Test
    void fileNamesAreNumberedInDocumentOrder() throws Exception {
        Images images = images(WordParser.ImageMode.META);
        images.add("PNG", 1, 1, PNG);
        images.add("JPEG", 2, 2, JPEG);
        images.link("PNG", 3, 3, "file:///D:/out/link.png");

        assertEquals(Arrays.asList("image-1.png", "image-2.jpg", "image-3.png"),
                Arrays.asList(images.list().get(0).get("fileName"), images.list().get(1).get("fileName"),
                        images.list().get(2).get("fileName")));
    }

    @Test
    void nullBytesAreIgnored() throws Exception {
        Images images = images(WordParser.ImageMode.META);
        images.add("PNG", 1, 1, null);

        assertEquals(0, images.list().size(), "拿不到字节的图片不该占一个编号");
    }

    @Test
    void unknownDimensionsAreReportedAsZero() throws Exception {
        // PDF 的 XImage 常常给不出显示尺寸，此时如实写 0，不猜
        Images images = images(WordParser.ImageMode.META);
        images.add("PNG", 0, 0, PNG);

        assertEquals(0, images.list().get(0).get("width"));
        assertEquals(0, images.list().get(0).get("height"));
    }

    /* ---------------- 类型归一化 ---------------- */

    @Test
    void typeIsNormalizedToUpperCaseWithoutMimePrefix() throws Exception {
        Images images = images(WordParser.ImageMode.META);
        images.add("image/png", 1, 1, PNG);
        images.add(" JPG ", 1, 1, JPEG);
        images.add("gif", 1, 1, new byte[]{0x47, 0x49, 0x46, 0x38, 1});

        assertEquals("PNG", images.list().get(0).get("type"));
        assertEquals("JPEG", images.list().get(1).get("type"), "jpg 是别名，按文件头确认成 JPEG");
        assertEquals("GIF", images.list().get(2).get("type"));
    }

    @Test
    void missingTypeIsSniffedFromTheBytes() throws Exception {
        Images images = images(WordParser.ImageMode.META);
        images.add(null, 1, 1, JPEG);
        images.add(null, 1, 1, OPAQUE);

        assertEquals("JPEG", images.list().get(0).get("type"));
        assertEquals("image-1.jpg", images.list().get(0).get("fileName"));
        assertEquals("UNKNOWN", images.list().get(1).get("type"), "认不出来就写 UNKNOWN，不瞎猜");
        assertEquals("image-2.bin", images.list().get(1).get("fileName"), "未知类型的扩展名是 bin");
    }

    @Test
    void typeStringCannotSmuggleAPathIntoTheFileName() throws Exception {
        // 类型串来自文档自带的 content-type，可能是 png/../../evil 这种；它会被拼进 fileName，
        // 所以第一个不是字母数字的字符往后全部丢掉
        Images images = images(WordParser.ImageMode.META);
        images.add("png/../../evil", 1, 1, OPAQUE);

        String fileName = String.valueOf(images.list().get(0).get("fileName"));
        assertFalse(fileName.contains("/"), fileName);
        assertFalse(fileName.contains(".."), fileName);
        assertEquals("image-1.bin", fileName);
    }

    @Test
    void sniffRecognizesTheCommonMagicNumbers() {
        assertEquals("PNG", Images.sniff(PNG));
        assertEquals("JPEG", Images.sniff(JPEG));
        assertEquals("GIF", Images.sniff(new byte[]{0x47, 0x49, 0x46, 0x38}));
        assertEquals("BMP", Images.sniff(new byte[]{0x42, 0x4D, 1, 2}));
        assertEquals("TIFF", Images.sniff(new byte[]{0x49, 0x49, 0x2A, 0x00}));
        assertEquals("TIFF", Images.sniff(new byte[]{0x4D, 0x4D, 0x00, 0x2A}));
        assertEquals("WMF", Images.sniff(new byte[]{(byte) 0xD7, (byte) 0xCD, (byte) 0xC6, (byte) 0x9A}));
        assertEquals(null, Images.sniff(null));
        assertEquals(null, Images.sniff(new byte[]{1, 2}), "不足 4 字节不认");
        assertEquals(null, Images.sniff(OPAQUE));
    }

    @Test
    void extensionAndContentTypeFollowTheType() {
        assertEquals("jpg", Images.extension("JPEG"));
        assertEquals("png", Images.extension("PNG"));
        assertEquals("bin", Images.extension("UNKNOWN"));
        assertEquals("image/png", Images.contentType("PNG"));
        assertEquals("application/octet-stream", Images.contentType("UNKNOWN"), "未知类型不能编成 image/unknown");
    }

    @Test
    void extensionDoesNotFollowTheDefaultLocale() {
        // 回归：extension 曾经用不带 Locale 的 toLowerCase()，土耳其语环境下 "TIFF" 小写成 tıff（无点 ı），
        // 文件名成了 image-1.tıff —— 任何图片查看器都不认。JVM 的默认 Locale 由部署环境决定，库不能跟着变
        Locale original = Locale.getDefault();
        try {
            Locale.setDefault(new Locale("tr", "TR"));
            assertEquals("tiff", Images.extension("TIFF"));
            assertEquals("png", Images.extension("PNG"));
            assertEquals("jpg", Images.extension("JPEG"));
            assertEquals("image/tiff", Images.contentType("TIFF"), "MIME 子类型同样不能被 Locale 改掉");
        } finally {
            Locale.setDefault(original);
        }
    }

    /* ---------------- 内容指纹 ---------------- */

    @Test
    void contentIdIsStableForTheSameBytesAndDiffersForDifferentOnes() throws Exception {
        assertEquals(Images.contentId(PNG), Images.contentId(PNG.clone()));
        assertNotEquals(Images.contentId(PNG), Images.contentId(JPEG));
        assertTrue(Images.contentId(PNG).matches("[0-9a-f]{40}"), Images.contentId(PNG));
    }

    /* ---------------- 三种带内容的模式 ---------------- */

    @Test
    void base64ModeBuildsADataUriWithTheRightMimeType() throws Exception {
        Images images = images(WordParser.ImageMode.BASE64);
        images.add("PNG", 1, 1, PNG);
        images.add(null, 1, 1, OPAQUE);

        String payload = String.valueOf(images.list().get(0).get("base64"));
        assertTrue(payload.startsWith("data:image/png;base64,"), payload);
        assertEquals(PNG.length, java.util.Base64.getDecoder()
                .decode(payload.substring(payload.indexOf(',') + 1)).length);
        assertTrue(String.valueOf(images.list().get(1).get("base64"))
                .startsWith("data:application/octet-stream;base64,"), "未知类型用通用 MIME");
    }

    @Test
    void bytesModeCarriesTheOriginalBytes() throws Exception {
        Images images = images(WordParser.ImageMode.BYTES);
        images.add("PNG", 1, 1, PNG);

        Object bytes = images.list().get(0).get("bytes");
        assertTrue(bytes instanceof byte[]);
        assertTrue(Arrays.equals(PNG, (byte[]) bytes), "字节应与原图一致");
        assertFalse(images.list().get(0).containsKey("base64"), "BYTES 模式不再给一份 base64");
    }

    @Test
    void fileModeWritesTheImageAndReportsTheRealName(@TempDir File directory) throws Exception {
        Images images = new Images(WordParser.options().images(WordParser.ImageMode.FILE).imageDir(directory));
        images.add("PNG", 1, 1, PNG);

        Map<String, Object> image = images.list().get(0);
        File written = new File(directory, String.valueOf(image.get("fileName")));
        assertTrue(written.isFile(), "图片应该落盘");
        assertEquals(PNG.length, written.length());
        assertEquals(written.getAbsolutePath(), image.get("path"));
        assertFalse(image.containsKey("base64"), "FILE 模式结果里不留内容");
        assertFalse(image.containsKey("bytes"), "FILE 模式结果里不留内容");
    }

    @Test
    void fileModeCreatesMissingDirectories(@TempDir File directory) throws Exception {
        File nested = new File(new File(directory, "out"), "images");
        Images images = new Images(WordParser.options().images(WordParser.ImageMode.FILE).imageDir(nested));
        images.add("PNG", 1, 1, PNG);

        assertTrue(new File(nested, "image-1.png").isFile(), "目录不存在时应自动创建");
    }

    @Test
    void fileModeWithoutDirectoryFails() {
        // 静默失败最难受：解析跑完了、图片一张也没落地，所以直接抛
        final Images images = images(WordParser.ImageMode.FILE);
        IllegalStateException thrown = assertThrows(IllegalStateException.class,
                () -> images.add("PNG", 1, 1, PNG));
        assertTrue(thrown.getMessage().contains("imageDir"), thrown.getMessage());
    }

    @Test
    void imageDirWithoutFileModeWarnsOnStderr(@TempDir File directory) throws Exception {
        // 配错组合要提示，否则调用方会以为图已经导出去了
        String stderr = TestSupport.captureStderr(() ->
                new Images(WordParser.options().imageDir(directory)));

        assertTrue(stderr.contains("imageDir"), stderr);
        assertTrue(stderr.contains("META"), stderr);
        assertTrue(stderr.contains(directory.getPath()), stderr);
    }

    @Test
    void fileModeNeverOverwritesAnExistingFile(@TempDir File directory) throws Exception {
        // 图片一律叫 image-N，每次解析都从 1 开始编号：批量解析写到同一个目录时，第二个文档不能覆盖第一个的图
        Images first = new Images(WordParser.options().images(WordParser.ImageMode.FILE).imageDir(directory));
        first.add("PNG", 1, 1, PNG);
        Images second = new Images(WordParser.options().images(WordParser.ImageMode.FILE).imageDir(directory));
        second.add("PNG", 1, 1, PNG);

        assertEquals("image-1.png", first.list().get(0).get("fileName"));
        assertEquals("image-1-2.png", second.list().get(0).get("fileName"), "重名时加序号");
        assertTrue(new File(directory, "image-1.png").isFile());
        assertTrue(new File(directory, "image-1-2.png").isFile());
        assertEquals(2, directory.listFiles().length);
    }

    @Test
    void uniqueFileAddsAnIncreasingSuffix(@TempDir File directory) throws Exception {
        assertEquals("image-1.png", Images.uniqueFile(directory, "image-1.png").getName());
        assertTrue(new File(directory, "image-1.png").createNewFile());
        assertEquals("image-1-2.png", Images.uniqueFile(directory, "image-1.png").getName());
        assertTrue(new File(directory, "image-1-2.png").createNewFile());
        assertEquals("image-1-3.png", Images.uniqueFile(directory, "image-1.png").getName());
    }

    @Test
    void uniqueFileSanitizesTheName(@TempDir File directory) {
        // 文件名可能来自外部反序列化的结果，../../ 开头也必须只当文件名用
        assertEquals("evil.png", Images.uniqueFile(directory, "../../evil.png").getName());
        assertEquals(directory, Images.uniqueFile(directory, "../../evil.png").getParentFile());
    }

    /* ---------------- 重复图片 ---------------- */

    @Test
    void inlineModesEmbedDuplicateContentOnlyOnce() throws Exception {
        Images images = images(WordParser.ImageMode.BASE64);
        images.add("PNG", 1, 1, PNG);
        images.add("PNG", 1, 1, PNG.clone());

        Map<String, Object> second = images.list().get(1);
        assertEquals(0, second.get("duplicateOf"), "指向前一条的下标");
        assertFalse(second.containsKey("base64"), "重复的内容不再内嵌一遍");
        assertEquals(images.list().get(0).get("contentId"), second.get("contentId"), "指纹是它自己的");
        assertEquals("image-2.png", second.get("fileName"), "文件名仍按顺序编号");
        assertEquals(PNG.length, second.get("size"), "元信息一样不少");
    }

    @Test
    void metaModeDoesNotTrackDuplicates() throws Exception {
        // META 模式本来就不带内容，没有「重复内嵌」可省，也就不写 duplicateOf
        Images images = images(WordParser.ImageMode.META);
        images.add("PNG", 1, 1, PNG);
        images.add("PNG", 1, 1, PNG.clone());

        assertFalse(images.list().get(1).containsKey("duplicateOf"));
        assertEquals(2, images.list().size());
    }

    @Test
    void fileModeWritesEveryDuplicateAsItsOwnFile(@TempDir File directory) throws Exception {
        // FILE 模式每张图都要落成自己的文件，去重会把该导出的文件少掉一个
        Images images = new Images(WordParser.options().images(WordParser.ImageMode.FILE).imageDir(directory));
        images.add("PNG", 1, 1, PNG);
        images.add("PNG", 1, 1, PNG.clone());

        assertEquals(2, directory.listFiles().length, "两张图各一个文件");
        assertFalse(images.list().get(1).containsKey("duplicateOf"));
    }

    /* ---------------- 外链图片 ---------------- */

    @Test
    void linkedImagesCarryOnlyTheSourcePath() {
        // Word 的「链接到文件」、PPT 的「仅链接」：字节不在文档包里，本模块不去读那个文件也不联网
        Images images = images(WordParser.ImageMode.BASE64);
        images.link("PNG", 800, 600, "file:///D:/images/chart.png");

        Map<String, Object> image = images.list().get(0);
        assertEquals("file:///D:/images/chart.png", image.get("link"));
        assertEquals("PNG", image.get("type"));
        assertEquals(800, image.get("width"));
        assertFalse(image.containsKey("size"), "外链图片没有字节数");
        assertFalse(image.containsKey("contentId"), "外链图片没有指纹");
        assertFalse(image.containsKey("base64"), "外链图片不内嵌内容");
        TestSupport.assertImageMeta(image);
    }

    @Test
    void linkedImageTypeIsSniffedAsUnknownWithoutBytes() {
        Images images = images(WordParser.ImageMode.META);
        images.link(null, 0, 0, "chart.emf");

        assertEquals("UNKNOWN", images.list().get(0).get("type"), "没有字节可嗅探时不猜类型");
    }
}
