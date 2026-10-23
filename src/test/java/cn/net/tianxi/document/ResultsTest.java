package cn.net.tianxi.document;

import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

import java.io.File;
import java.util.*;
import java.util.concurrent.atomic.AtomicInteger;

import static org.junit.jupiter.api.Assertions.*;

/**
 * {@link Results} 的单元测试：文本清洗、评估水印识别、表头归一化、文件名消毒、许可证加载、写图片。
 *
 * <p>这一层是三种格式共用的，出错会同时影响 Word/PDF/PPT，所以全部用最小输入直接断言，不经过 Aspose。
 * 控制字符在注释里写成 {@code <XX>}（十六进制），源码字符串里用反斜杠 u 转义 —— 直接敲进注释会被编译器当成真字符。
 */
class ResultsTest {

    /* ---------------- clean：域代码、软换行、控制字符 ---------------- */

    @Test
    void cleanDropsFieldCodeAndKeepsFieldResult() {
        // <13> 域开始、<14> 代码与结果的分界、<15> 域结束
        assertEquals("前文 3 后文", Results.clean("前文 \u0013 PAGE \u00143\u0015 后文"));
    }

    @Test
    void cleanDropsWholeFieldWhenResultNotCalculated() {
        // 新建或转换出来的文档里域结果常常是空的，这段文字整块消失是对的：那是排版指令，不是内容
        assertEquals("前文 后文", Results.clean("前文 \u0013 PAGE \u0014\u0015 后文"));
    }

    @Test
    void cleanDropsNestedFieldCode() {
        // TOC 里套 PAGE：内层域整体落在外层的结果区，内层自己的代码仍要丢掉
        assertEquals("目录 7 结尾", Results.clean("目录 \u0013 TOC \u0014\u0013 PAGE \u00147\u0015\u0015 结尾"));
    }

    @Test
    void cleanTurnsLineBreaksIntoSpacesInsteadOfDroppingThem() {
        // 软换行 <0B>、分页符 <0C>、段落结束 <0D>：换成空格而不是丢掉，否则 Shift+Enter 分的两行会粘成一句
        assertEquals("第一行 第二行", Results.clean("第一行\u000b第二行\r"));
        assertEquals("上一页 下一页", Results.clean("上一页\u000c下一页"));
        // <0A> 也走同一条规则：clean 是三个解析器共用的那一道，换行的写法不止 <0D> 一种
        assertEquals("标题 副标题", Results.clean("标题\n副标题"));
    }

    @Test
    void cleanDropsCellTerminatorAndOtherControlCharacters() {
        // <07> 是单元格结束符，<01>/<02> 之类是 Aspose 的内部占位，都不该进结果
        assertEquals("普通格", Results.clean("\u0013PAGE\u0014\u0015\u0007普通格\u0007"));
        assertEquals("ab", Results.clean("a\u0001\u0002b"));
    }

    @Test
    void cleanCollapsesSpacesAndTrims() {
        assertEquals("a b", Results.clean("  a    b  "));
        assertEquals("", Results.clean("   "));
        assertEquals("", Results.clean(null));
    }

    @Test
    void cleanKeepsNonBreakingSpaceInWords() {
        // 不间断空格 <A0> 在正文里是有意义的分隔（clean 换成普通空格），但表头比对要完全无视它，见 normalize
        assertEquals("列 1", Results.clean("列\u00a01"));
    }

    /* ---------------- 评估水印 ---------------- */

    @Test
    void recognizesEvaluationWatermark() {
        assertTrue(Results.isEvaluationNotice("Evaluation Only. Created with Aspose.Words. Copyright 2001-2023 Aspose Pty Ltd."));
        assertTrue(Results.isEvaluationNotice("Created with an evaluation copy of Aspose.Slides"));
        assertTrue(Results.isEvaluationNotice("To discover the full versions of our APIs, please visit our website"));
    }

    @Test
    void recognizesWatermarkSplitAcrossLinesAndInAnyCase() {
        // PDF 里水印被排版切成好几段，只剩半截也要认得出来；比对前去掉了所有空格并转小写
        assertTrue(Results.isEvaluationNotice("Pty Ltd."));
        assertTrue(Results.isEvaluationNotice("evaluationonly.createdwithaspose"));
        assertTrue(Results.isEvaluationNotice("EVALUATION   ONLY"));
    }

    @Test
    void ordinaryTextIsNotWatermark() {
        assertFalse(Results.isEvaluationNotice("2025 年度销售情况"));
        assertFalse(Results.isEvaluationNotice(""));
        assertFalse(Results.isEvaluationNotice(null));
    }

    @Test
    void cleanContentReturnsEmptyForWatermarkParagraph() {
        assertEquals("", Results.cleanContent("Evaluation Only. Created with Aspose.PDF. Copyright 2001-2023 Aspose Pty Ltd."));
        assertEquals("正文", Results.cleanContent("正文\r"));
    }

    /* ---------------- normalize：表头比对 ---------------- */

    @Test
    void normalizeIgnoresCaseAndEveryKindOfSpace() {
        // 半角、全角 <3000>、不间断 <A0>、数字空格 <2007>、窄不换行空格 <202F>、制表符都算空格
        assertEquals("列bb", Results.normalize("列 Bb"));
        assertEquals("列bb", Results.normalize("列\u3000Bb"));
        assertEquals("列bb", Results.normalize(" 列\u00a0B\u2007b\u202f "));
        assertEquals("列bb", Results.normalize("列\tBb\n"));
        assertEquals("", Results.normalize(null));
    }

    @Test
    void normalizeUsesRootLocaleSoTurkishEnvironmentCannotBreakIt() {
        // 默认 Locale 是土耳其语时 toLowerCase() 会把 I 变成 ı，注册表头 ItemId 就再也匹配不上 itemId
        assertEquals("itemid", Results.normalize(" ItemId "));
    }

    /* ---------------- safeName：写文件前的文件名消毒 ---------------- */

    @Test
    void safeNameDropsDirectoryParts() {
        assertEquals("evil.png", Results.safeName("../../evil.png"));
        assertEquals("evil.png", Results.safeName("..\\..\\evil.png"));
        assertEquals("evil.png", Results.safeName("/etc/evil.png"));
        assertEquals("image-1.png", Results.safeName("image-1.png"));
    }

    @Test
    void safeNameFallsBackForBlankAndDotNames() {
        assertEquals("image.bin", Results.safeName(null));
        assertEquals("image.bin", Results.safeName(""));
        assertEquals("image.bin", Results.safeName("."));
        assertEquals("image.bin", Results.safeName(".."));
        assertEquals("image.bin", Results.safeName("/"));
    }

    /* ---------------- join / toJson ---------------- */

    @Test
    void joinInsertsSeparatorBetweenItemsOnly() {
        assertEquals("", Results.join(new ArrayList<String>(), "\n"));
        assertEquals("a", Results.join(Arrays.asList("a"), "\n"));
        assertEquals("a\nb\nc", Results.join(Arrays.asList("a", "b", "c"), "\n"));
    }

    @Test
    void toJsonIndentsAndEncodesByteArrays() throws Exception {
        Map<String, Object> value = new LinkedHashMap<String, Object>();
        value.put("text", "正文");
        value.put("bytes", new byte[]{1, 2, 3});
        String json = Results.toJson(value);
        assertTrue(json.contains("\n"), "应该是缩进输出：" + json);
        assertTrue(json.contains("\"text\" : \"正文\""), json);
        // Jackson 把 byte[] 编成不带 data: 前缀的 base64
        assertTrue(json.contains(Base64.getEncoder().encodeToString(new byte[]{1, 2, 3})), json);
    }

    /* ---------------- writeImages ---------------- */

    @Test
    void writeImagesDecodesBase64IntoFiles(@TempDir File directory) throws Exception {
        byte[] bytes = {(byte) 0x89, 0x50, 0x4E, 0x47, 1, 2, 3};
        Map<String, Object> result = resultWithBase64Image("image-1.png", bytes);

        Results.writeImages(result, directory);

        File written = new File(directory, "image-1.png");
        assertTrue(written.isFile(), "图片应该落盘：" + written);
        assertEquals(bytes.length, written.length());
    }

    @Test
    void writeImagesCannotEscapeTargetDirectory(@TempDir File directory) throws Exception {
        // fileName 可能来自外部反序列化的结果，../../ 开头也必须只当文件名用
        Map<String, Object> result = resultWithBase64Image("../../evil.png", new byte[]{1, 2});

        Results.writeImages(result, directory);

        assertTrue(new File(directory, "evil.png").isFile(), "应该写在目标目录里");
        assertFalse(new File(directory.getParentFile(), "evil.png").exists(), "不能写到目录之外");
    }

    @Test
    void writeImagesDoesNotOverwriteExistingFiles(@TempDir File directory) throws Exception {
        // 批量解析多个文档、导出到同一个目录时，重名图片由 Images.uniqueFile 加序号
        Results.writeImages(resultWithBase64Image("image-1.png", new byte[]{1}), directory);
        Results.writeImages(resultWithBase64Image("image-1.png", new byte[]{2}), directory);

        assertTrue(new File(directory, "image-1.png").isFile());
        assertTrue(new File(directory, "image-1-2.png").isFile(), "第二张应该改名而不是覆盖第一张");
    }

    @Test
    void writeImagesWarnsWhenResultCarriesNoContent(@TempDir File directory) throws Exception {
        // 默认的 META 模式没有 base64 字段，一张也导不出 —— 要在 stderr 说清楚，不能安静地什么都不做
        Map<String, Object> metaOnly = new LinkedHashMap<String, Object>();
        List<Map<String, Object>> images = new ArrayList<Map<String, Object>>();
        Map<String, Object> image = new LinkedHashMap<String, Object>();
        image.put("fileName", "image-1.png");
        images.add(image);
        metaOnly.put("images", images);

        String stderr = TestSupport.captureStderr(() -> Results.writeImages(metaOnly, directory));

        assertEquals(0, directory.listFiles().length, "META 模式没有内容可写");
        assertTrue(stderr.contains("BASE64"), "应该提示改用哪种模式：" + stderr);
    }

    @Test
    void writeImagesIgnoresResultWithoutImageList(@TempDir File directory) throws Exception {
        Results.writeImages(new LinkedHashMap<String, Object>(), directory);
        assertEquals(0, directory.listFiles().length);
    }

    /* ---------------- 许可证 ---------------- */

    @Test
    void licenseStreamReturnsNullWhenNothingFound() {
        assertNull(Results.licenseStream("/definitely-not-here.lic", "definitely-not-here.lic"));
    }

    @Test
    void licenseStreamFindsClasspathResource() {
        // 测试资源就在 classpath 上，拿它当「找得到」的例子
        java.io.InputStream stream = Results.licenseStream("/word/sales-report.docx", "whatever.lic");
        assertTrue(stream != null, "classpath 里存在的资源应该能拿到流");
        Results.close(stream);
    }

    @Test
    void applyLicenseRegistersEachProductOnlyOnce() throws Exception {
        final AtomicInteger calls = new AtomicInteger();
        Results.LicenseApplier applier = new Results.LicenseApplier() {
            @Override
            public void setLicense(java.io.InputStream stream) {
                calls.incrementAndGet();
            }
        };
        Results.applyLicense("unit-test-product", "/word/sales-report.docx", "whatever.lic", applier);
        Results.applyLicense("unit-test-product", "/word/sales-report.docx", "whatever.lic", applier);
        assertEquals(1, calls.get(), "同一产品在一个进程里只注册一次");
    }

    @Test
    void applyLicenseReportsFailureInsteadOfSwallowingIt() throws Exception {
        // lic 过期或写错时，过去的现象是「输出莫名其妙少了内容」，所以失败要提示一行
        String stderr = TestSupport.captureStderr(() -> Results.applyLicense("unit-test-broken",
                "/word/sales-report.docx", "whatever.lic",
                stream -> {
                    throw new IllegalStateException("lic 已过期");
                }));
        assertTrue(stderr.contains("unit-test-broken"), stderr);
        assertTrue(stderr.contains("lic 已过期"), stderr);
    }

    @Test
    void applyLicenseStaysQuietWhenThereIsNoLicenseFile() throws Exception {
        final AtomicInteger calls = new AtomicInteger();
        String stderr = TestSupport.captureStderr(() ->
                Results.applyLicense("unit-test-missing", "/definitely-not-here.lic", "definitely-not-here.lic",
                        stream -> calls.incrementAndGet()));
        assertEquals(0, calls.get(), "没有 lic 文件就按评估模式走");
        assertEquals("", stderr, "没有 lic 是正常情况，不该提示");
    }

    @Test
    void closeToleratesNullAndFailure() {
        Results.close(null);
        Results.close(new java.io.InputStream() {
            @Override
            public int read() throws java.io.IOException {
                throw new java.io.IOException("关不掉");
            }

            @Override
            public void close() throws java.io.IOException {
                throw new java.io.IOException("关不掉");
            }
        });
    }

    private static Map<String, Object> resultWithBase64Image(String fileName, byte[] bytes) {
        Map<String, Object> image = new LinkedHashMap<String, Object>();
        image.put("fileName", fileName);
        image.put("base64", "data:image/png;base64," + Base64.getEncoder().encodeToString(bytes));
        List<Map<String, Object>> images = new ArrayList<Map<String, Object>>();
        images.add(image);
        Map<String, Object> result = new LinkedHashMap<String, Object>();
        result.put("images", images);
        return result;
    }
}
