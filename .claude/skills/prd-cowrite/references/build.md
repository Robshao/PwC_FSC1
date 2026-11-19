# PRD HTML 构建与有限图形检查

只手写 Markdown 正文源；评审改动回改源再生成 HTML。工具仅使用 Python 标准库，无需安装依赖；通过 `uv run --no-project --offline python` 运行。已经安装的 Python 必须可用，离线运行不会替你补齐运行环境。

## 运行

以下 `<skill>` 是占位符，必须替换为真实安装目录的绝对路径，例如 `/绝对路径/prd-cowrite`；安装时保持 `scripts/`、`assets/`、`references/` 同级。无须读脚本全文，先读本文；参数速查用 `--help`。

```sh
# 不要求先切换到 skill 目录；默认图库随脚本定位。
uv run --no-project --offline python "<skill>/scripts/build_prd_html.py" "<skill>/assets/example.md" ./prd-fragment.html --standalone ./prd-full.html --check-chairs

# 使用自己的正文、图库和单图目录。请将 /绝对路径/... 换成实际文件。
uv run --no-project --offline python "<skill>/scripts/build_prd_html.py" /绝对路径/源.md /绝对路径/片段.html --standalone /绝对路径/完整.html --figures /绝对路径/图库.json --figs-dir /绝对路径/单图目录 --check-chairs

uv run --no-project --offline python "<skill>/scripts/svg_lint.py" ./prd-full.html
uv run --no-project --offline python "<skill>/scripts/build_prd_html.py" --help
```

输出目录须预先存在。片段没有 HTML 外壳，供有明确接入约定的宿主页使用；`--standalone` 额外写出带外壳的独立页面。独立页本身不加载外部渲染库。

未给 `--figures` 时，读取脚本所属 skill 根的 `assets/figures.example.json`，不受当前工作目录影响。显式 `--figures` 优先；相对输入、输出、图库、单图目录及 `@args.txt` 路径仍相对调用时的工作目录，兼容旧行为。长参数可用 `@args.txt` 传入，一行一个 token。

其他参数有 `--title`（默认取 Markdown 首行 `# ` 标题）、`--h1`（默认同 title）、`--kicker`、`--subtitle`、`--legend-html`、`--chairs-lede`、`--footer-src-name`（默认源文件名）、`--footer-extra`。

## 图库与单图合并

图库形状为 `{"css": "页面基础样式", "figs": {"f0": {"title": "图名", "caption": "说明", "body": "内联 SVG 或 HTML"}}}`。示例样式和图见 `assets/figures.example.json`。

`--figs-dir` 读取目录中的 `*.json`，用文件名去掉 `.json` 作为 fid；同名覆盖、新名新增，其他扩展名忽略。单图文件内容为 `{"title": "图名", "caption": "说明", "body": "图内容"}`。合并仅发生在内存，不回写基础图库或单图文件；输入内容键没有额外校验。

占位符使用小写 `f` 加一位或多位数字，例如 `{{fig:f0}}`、`{{fig:f12}}`、`{{fig:f123}}`，独占一行。基础图库与单图目录都支持多位图号：`f12.json` 可由 `{{fig:f12}}` 引用；授权标记前的提图入口、正文和折叠块内的入口都支持。图号不要求连续；缺少被引用的图会构建失败。任意文件名虽能被合并，只有上述 fid 格式能被图占位符引用。

## Markdown 约定

- 首行写 `# 标题`。必填 `<!-- meta` 换行块（`状态: ...`、`受众: ...`、`权威: ...`、`变更: ...`）和 `<!-- audience` 换行块（内放 Markdown 表），均以 `-->` 结束。只校验两块是否存在，不校验四键齐全。
- 章用 `## 一、章名`，节用 `### 1.1 节名`，小标题用 `#### 名`。无顿号的章按附录处理。章/节后的 `> 结论：…` 显示在折叠外。
- 节内含 `<!-- auth -->` 才渲染该节正文；无此标记时显示待授权占位。标记前的图提到节首，标记后的图留在正文原位；这是排版约定，不是权限或安全控制。
- `<!-- fold -->` 前直接展示、后进入“展开细节”。bullet 后紧跟四空格缩进块会成为“详细”折叠；不认 tab，发现行首 tab 会警告。
- `<mark>…</mark>` 产生黄底待拍项、页顶数量和页尾索引；`!!…!!` 为强调条款，支持 `**粗体**`、行内代码和 `§一` 章引用。表格、有序列表、引用块及 `> 可判定项：…` 可参考 `assets/example.md`；不支持通用 Markdown 的全部语法。
- `mermaid` 代码围栏会转为转义后的 `<pre class="mermaid">`。**独立页不会自动渲染 Mermaid**，只显示源码；需要图时，由已明确配置渲染器的宿主页处理，或先离线生成受信任的 SVG 再放入图库。builder 不下载、不初始化 Mermaid。

## 验收退出状态

`--check-chairs` 比较三数：源 Markdown 的字面 `<mark>` 数、页顶“待拍 N”、生成 HTML 中 `class="chairtext"` 的出现数减去图例里的样例数。一致返回 0，不一致返回 3；检查发生在写产物之后，失败时片段和独立页仍可能存在。源码计数含未授权节、注释或示例中的字面标记；页面数是字符串统计，未验证浏览器中的真实可见性。不要用产物存在代替退出状态。

`svg_lint.py` 的状态如下。读取或解析失败也会非零退出，须结合诊断输出判断，不能把所有非零都当成几何问题。

| 退出码 | 含义 | 后续判断 |
|---|---|---|
| 0 | 至少检出一个内联 SVG，有限规则未报问题 | 仍须真实渲染看屏 |
| 1 | 有已报告的几何问题；或读取/解析异常 | 查看输出并处理，不据此验收通过 |
| 2 | 未检出内联 SVG，检查对象数为 0 | 属于未覆盖，不是通过 |

lint 用正则粗查小写 `<svg>...</svg>`：属性按双引号解析，预期有效 `viewBox` 和直接数值；主要识别自闭合 `<rect .../>`、`<path .../>` 中的大写 M/H/V 命令及 `<text>`，检查线穿框、文字压线、文字横向出界和文字互压。含大写 L 的 path 整条跳过；其他命令不是完整解析，不应将部分读取当完整覆盖。宽度至少 900 的矩形不参与线穿框检查，class 恰为 `pill` 的矩形和文字也会跳过。

它不计算完整 SVG/CSS 布局、transform、真实字体度量、tspan 布局、曲线、斜线或文字纵向出界；横向边界计算也仅适合 x 原点为 0 的 viewBox。`<g>` 继承仅粗略处理 font-size/text-anchor。外链 SVG、图片及未渲染 Mermaid 不在覆盖范围，复杂或不匹配的写法可能漏查、误报或解析失败。**即使返回 0，也需要在实际宿主/浏览器中看屏，按本次改动检查字号、遮挡、溢出和折叠交互。**

## HTML / JavaScript 信任边界

这两个脚本不是 HTML 清洗器、安全扫描器或沙箱。图库的 `css`、图的 `title/caption/body`，以及 `--legend-html` 等页面参数会直接拼入 HTML；部分 Markdown 标题等也直接用于页面外壳。普通正文有转义，不代表整体输入可信。

只在输入来源可信时生成并打开页面。恶意 HTML/SVG、事件处理器、脚本或 CSS/远程资源引用可随生成结果进入浏览器；打开时是否执行脚本、发送请求取决于宿主和 CSP。builder 自身不执行页面 JavaScript、不自动联网加载库；lint 不消除这些风险。处理不可信材料时，先在另行确定的清洗或隔离流程中处理，不能把 `<!-- auth -->`、`--check-chairs` 或 lint 作为安全许可。添加渲染器、放开脚本或加载外部资源须由调用方明确决定。
