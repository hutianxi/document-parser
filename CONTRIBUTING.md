# 贡献指南（CONTRIBUTING）

欢迎提 issue、提 PR。本项目采用 [Apache License 2.0](LICENSE)：你可以自由使用、修改、再分发，
按 Apache 2.0 第 5 条，**你提交到本仓库的贡献默认也按 Apache 2.0 授权**，无需另签协议；
署名要求见 [NOTICE](NOTICE)。

代码托管在两处，内容一致：**主仓库** <https://github.com/hutianxi/document-parser>（issue 与 PR 都走这里），
**国内镜像** <https://gitee.com/mailtianxi/document-parser>（由主仓库推送同步，方便访问不到 GitHub 的人拉代码）。

---

## 1. 环境

| 项 | 要求 | 说明 |
| --- | --- | --- |
| JDK | 8 及以上 | 源码与字节码级别都是 1.8（`java.version=1.8` → `-source/-target 1.8`），所以用高版本 JDK 编译产物仍是 `major version: 52`；本机在 JDK `1.8.0_191` 与 `21.0.1` 下各跑一遍完整测试都是 183 全过、1 跳过 |
| Maven | 3.6+（本机实测 3.8.8） | 首次构建需联网，Aspose 构件来自 `https://repository.aspose.com/repo/` |
| 网络 | 首次必须能访问 Aspose 仓库 | 之后可用 `mvn -o` 离线构建 |

`aspose-*.lic` 放在 classpath 根目录或进程工作目录即可解除评估模式限制（见 README 第 12.2 节）。
**lic 文件是私有凭证，永远不要提交进仓库**——`.gitignore` 已经拦住了 `*.lic`。

## 2. 构建与测试

```
mvn -B clean test          # 当前基线：183 个测试，0 失败，1 跳过（-Ddoc 入口）；JDK 8 与 JDK 21 都实测全绿
mvn -B clean package       # 产物 target/document-parser-1.1.1.jar
```

想直接看某个样例的解析结果（打印到控制台，同时落一份 UTF-8 文本到 `target/print/`）：

```
mvn -B test -Dtest=PrintResultTest#printWordSample
mvn -B test -Dtest=PrintResultTest#printAnyDocument -Ddoc="D:/你的文档.docx"
```

javadoc 的 doclint 核对（本机离线仓库没有 `maven-javadoc-plugin`，所以用 JDK 自带的 `javadoc`）：

```
mvn -B dependency:build-classpath -Dmdep.outputFile=target/cp.txt
javadoc -encoding UTF-8 -charset UTF-8 -Xdoclint:all,-missing \
        -cp "$(cat target/cp.txt)" -d target/site/apidocs $(find src/main/java -name '*.java')
```

## 3. 代码约定（改动前请先读）

1. **纯静态工具类**：不加启动类、不加 `main`、不加命令行参数解析、不引入 Spring 容器。
2. **三种格式一套结构**：`WordParser` / `PdfParser` / `PptParser` 方法签名一一对应，返回 `Map`，
   键固定是 `text` / `tables` / `images`。表头分组只写在 `Tables`、图片只写在 `Images`、
   文本清洗与 JSON 只写在 `Results`，**不要在解析器里各抄一份**。
3. **结果契约不能破**（`CrossFormatTest` 守着）：`headers` 逐个等于每一行的键；
   空表头补 `列N`、同名列补 `名-列号`；表格里的文字不在 `text` 里重复；图片顺序 = 文档流顺序。
4. **javadoc 里的数字必须是实测的**：本项目的类说明大量引用 `src/test/resources/` 下三份样例的真实输出
   （段数、行数、列名、JSON 字符数）。改了行为或样例，就把对应数字重新测一遍再写进去；
   没有样例可测的防御分支，要在注释里明说「无样例覆盖，由 `XxxTest` 的单元测试守着」。
   一次性探针脚本用完就删，不留进仓库。
5. **加新格式**：写一个 `XxxParser`，把表格读成 `String[][]` + `boolean[] separator` + 可选 `boolean[] heading`，
   然后调 `Tables.group(...)`；图片交给 `new Images(options).add(...)`；`Options`、`toJson`、`writeImages` 复用。

## 4. 提 PR 的流程

1. Fork 本仓库，从 `main` 切一个短分支（`fix/pdf-nested-table`、`feat/excel-parser` 这种）。
2. 一个 PR 只做一件事。行为变更请**同时**给测试；判据（表头识别、水印过滤、列名补全等）变更请附
   **修改前后的端到端 JSON 对比**，只写「已修复」不够。
3. 本地跑通 `mvn -B clean test`（全绿）与 doclint（0 error），把实测数字写进 PR 描述。
4. 不要提交：`target/`、`.idea/`、`*.iml`、`*.lic`、任何真实客户文档。
   新样例文档请自己用 OOXML 手工拼或自己导出，**不要放含有真实业务数据的文件**。
5. 维护者会按第 3 节的约定复审；合并方式是 squash，所以 PR 内的中间提交不必刻意整理。

## 5. 提 issue 时请带上

- 文档格式与来源（Word 导出的 PDF？扫描件？扫描件没有文字层，本模块不做 OCR，读不出来是预期的）；
- 调用方式：用的哪个解析器、`Options` 怎么配的（有没有注册表头、`fullRow`、`images` 模式）；
- 是否已加载 Aspose lic：没有 lic 时 PPT 文本框只给前 5 个字符、PDF 文字层读不全，这类现象属于评估模式限制，
  不是 bug；
- 期望输出与实际输出（贴 `PrintResultTest` 打出来的那一份最直观）。

## 6. 许可红线

- 本仓库代码按 Apache 2.0 授权；**Aspose 的三个 jar 是商业软件，不随本仓库分发，也不受 Apache 2.0 覆盖**。
  任何 PR 都不要把 Aspose 的二进制、`.lic` 文件或破解/绕过评估限制的代码放进来。
- 修改后再分发时，请保留 `LICENSE`、`NOTICE` 与源文件里的版权信息，并按 Apache 2.0 第 4(b) 条
  在被改过的文件里注明「已修改」。
