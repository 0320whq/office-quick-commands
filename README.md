# 快捷指令 · Office 加载项

为 Microsoft Office 提供一组「快捷指令」：命令面板、键盘快捷键与一键常用操作，支持 **Excel / Word**（Windows、Mac、网页版均可）。

- **命令面板**：在「快捷指令」功能区单击「打开命令面板」，或在面板中直接点击 / 搜索执行命令。
- **键盘快捷键**：无需鼠标，按下组合键即可触发常用操作。
- **功能区按钮**：在「快捷指令」选项卡中提供一键操作按钮（Word / Excel 各自一套，并按主题归组为菜单）。
- **Word 专属符号面板**：在命令面板底部点击即可插入常用特殊符号。
- **护眼模式**：切换命令面板柔和配色（仅影响本面板，不影响文档），状态自动记忆。

---

## 功能一览

> 命令面板按当前主机自动筛选可用指令；键盘快捷键与功能区按钮为精选子集。
> Mac 上 `Ctrl+Alt` 对应 `Command+Option`；菜单项无单独快捷键，请用命令面板或功能区按钮。

### 通用（Word + Excel 均可）

| 指令 | 说明 | 快捷键 | 功能区 |
| --- | --- | --- | --- |
| 插入日期 | Excel：选区填充当前日期；Word：光标处插入日期时间 | `Ctrl+Alt+D` | 插入日期 |
| 加粗 / 斜体 / 下划线 | 切换选区对应格式 | `Ctrl+Alt+B` / `Ctrl+Alt+I` / `Ctrl+Alt+U` | 加粗 / 斜体 / 下划线 |
| 居中 / 左 / 右 / 两端对齐 | 段落 / 单元格对齐 | `Ctrl+Alt+C`（居中） | 居中 |
| 突出显示 | Excel 填充黄色 / Word 高亮选区 | `Ctrl+Alt+H` | 突出显示 |
| 清除格式 | 清除选区格式 | `Ctrl+Alt+K` | 清除格式 |
| 导出 PDF | 将当前文档导出为 PDF 并下载 | `Ctrl+Alt+P` | 导出 PDF |
| 护眼模式 | 切换面板柔和配色（记忆状态） | — | 护眼模式 |

### 仅 Word（21 项）

| 指令 | 说明 | 快捷键 |
| --- | --- | --- |
| 标题 1 / 标题 2 / 正文 | 套用内置样式 | `Ctrl+Alt+J`（标题1） |
| 项目符号 / 编号 / 取消列表 | 列表格式 | `Ctrl+Alt+L` / `Ctrl+Alt+N` |
| 1.5 倍 / 2 倍行距 | 行距 | — |
| 首行缩进 / 取消缩进 | 段落缩进 | — |
| 段前分页 | 为所选段落设置“段前分页” | — |
| 转为大写 / 小写 / 标题 / 句首 / 切换 | 文本大小写全套（标题/句首/切换基于算法推断） | `Ctrl+Alt+Shift+U`（大写） |
| 插入 3×3 表格 | 光标后插入表格 | — |
| 插入分页符 / 分节符 | 光标后插入分隔符 | — |
| 删除空段落 / 清理多余空格 | 删空段 / 合并连续空格并去首尾 | — |
| 清除高亮 | 取消高亮颜色 | — |
| 字数统计 | 全文词数 + 选区字数/字符数 | `Ctrl+Alt+W` |
| 删除空白页 | 清掉空段落 / 分页占位导致的空白页 | — |
| 表格转文本 / 文本转表格 | 表格 ↔ 文本互转 | — |
| 插入页码 / 页眉 / 页脚 | 每个节页脚/页眉写入 PAGE 域或文字 | — |
| 多级列表 | 按段落前导空格自动分级编号 | — |
| 转为繁体 / 转为简体 | 选区简繁互转（opencc-js 全量词典，含多义字校正） | `Ctrl+Alt+Shift+T` / `Ctrl+Alt+Shift+S` |
| 高级替换 | 选区查找替换，可选区分大小写 / 全字匹配（记忆上次输入） | `Ctrl+Alt+Shift+R` |
| 插入特殊符号 | 面板底部符号网格点击插入（§ ¶ © ® ° × ÷ ± … ★ 等） | — |

### 仅 Excel（29 项）

| 指令 | 说明 | 快捷键 | 功能区 |
| --- | --- | --- | --- |
| 循环填充色 | 调色板循环切换选区填充色 | `Ctrl+Alt+Q` | 循环填充色 |
| 去空格（Trim） | 去除单元格首尾及多余空格 | `Ctrl+Alt+T` | 去空格 |
| 删除空白行 | 删除选区内整行空白行 | — | — |
| 清除内容 | 仅清值保留格式 | `Ctrl+Alt+X` | — |
| 自动调整 | 自动调整列宽与行高 | `Ctrl+Alt+A` | 自动调整 |
| 冻结首行 / 首列 / 取消冻结 | 冻结窗格 | `Ctrl+Alt+F` / `Ctrl+Alt+Shift+F` | 冻结首行 / 取消冻结 |
| 插入行 / 插入列 | 选区上方 / 左侧插入 | `Ctrl+Alt+R`（行） | 插入行 |
| 合并 / 拆分单元格 | 合并 / 取消合并 | `Ctrl+Alt+M`（合并） | 合并单元格 |
| 选中数据区 | 选中工作表已用区域 | — | — |
| 求和 / 平均值 | 选区每列下方插入 SUM / AVERAGE | `Ctrl+Alt+S`（求和） | 求和 |
| 升序 / 降序排序 | 按第一列排序 | `Ctrl+Alt+E`（升序） | 升序排序 |
| 生成图表 | 选区数据生成聚类柱形图 | `Ctrl+Alt+G` | 生成图表 |
| 数据条 | 选区添加数据条条件格式 | — | — |
| 转为表格 | 选区转为带样式表格 | — | — |
| 删除重复行 | 按首列删重（保留表头） | — | — |
| 按颜色求和 | 以活动单元格填充色为基准对全表同色数值求和 | `Ctrl+Alt+Y` | 按颜色求和 |
| 查找重复值 | 标红重复单元格（保留首个） | `Ctrl+Alt+O` | — |
| 数字格式：货币 / 百分比 / 千分位 / 日期 | 批量设置数字格式 | — | — |
| 加边框 / 去边框 | 选区加/去外侧+内侧边框 | `Ctrl+Alt+Shift+B` | — |
| 隐藏 / 取消隐藏 行与列 | 行列显隐 | — | — |
| 转置 | 选区转置到其右侧 | `Ctrl+Alt+Z` | — |
| 公式转值 | 公式替换为计算值（保留数字格式） | `Ctrl+Alt+V` | — |
| 文本分列 | 按分隔符（逗号/空格/Tab/竖线）拆首列到右侧多列 | — | 文本分列 |
| 定位空值 | 选中选区内所有空白单元格 | — | 定位空值 |
| 数据有效性 | 选区添加下拉列表（逗号分隔选项） | — | 数据有效性 |
| 高级替换 | 选区查找替换，可选区分大小写 / 全字匹配（记忆上次输入） | `Ctrl+Alt+Shift+R` | 高级替换 |

> 命令面板 `COMMANDS` 总计 **84 项**（含共享/仅 Word / 仅 Excel，其中「高级替换」在 Word 与 Excel 各一项）。功能区按钮经菜单归组，保持选项卡整洁。

---

## 目录结构

```
office-quick-commands/
├── manifest.xml          # 加载项清单（指向 GitHub Pages，无需本地服务器）
├── manifest-local.xml    # 本地调试用清单（指向 localhost:3000，与 manifest.xml 结构一致）
├── shortcuts.json         # 键盘快捷键与动作的映射（含 Mac 变体）
├── setup.js              # 一键把 manifest.xml 指向你的 GitHub Pages
├── publish.js            # 一键校验（语法/JSON/XML/一致性）并 git 发布
├── server.js             # 本地 HTTPS 静态服务器（Node 标准库，零依赖）
├── assets/               # 图标（icon-16/32/80.png，占位图，可替换）
├── src/
│   ├── taskpane.html     # 命令面板界面（含 opencc 引用）
│   ├── taskpane.css      # 面板样式
│   ├── taskpane.js       # 命令实现 + 动作注册（Office.actions.associate）
│   └── lib/
│       └── opencc-bundle.js  # 简繁转换词典（opencc-js UMD 打包，离线内置，约 1.18MB）
├── tools/
│   └── make_icons.py      # 生成图标的脚本（标准库，可选重跑）
└── README.md
```

---

## 部署（任选其一，推荐方案一）

Office 要求加载项内容通过 **HTTPS** 提供。下面两条路都能用，**方案一不用在你电脑上常驻任何服务器**，最省心。

### 方案一：GitHub Pages（推荐，免服务器）

1. **配置地址（一次性）**：在项目目录运行
   ```bash
   node setup.js
   ```
   按提示输入你的 GitHub 用户名，脚本会把 `manifest.xml` 里的地址改成
   `https://<你的用户名>.github.io/office-quick-commands/`。
2. **上传到 GitHub**：
   ```bash
   git init
   git add -A
   git commit -m "快捷指令 Office 加载项"
   git branch -M main
   git remote add origin https://github.com/<你的用户名>/office-quick-commands.git
   git push -u origin main
   ```
3. **开启 Pages**：仓库页面 → `Settings` → `Pages` → `Source` 选 `main` 分支 `/ (root)`，保存。**约 1 分钟后**生效。
4. **旁加载（只需一次）**：打开 Office →「插入」→「我的加载项」→「管理我的加载项」→「上传我的加载项」，选择本目录的 **`manifest.xml`**。
   之后每次打开 Office 都会出现「快捷指令」选项卡，无需再运行任何服务器。

> 没有 GitHub 账号？用方案二；或把 `src/`、`assets/`、`shortcuts.json` 托管到任意 HTTPS 站点后，手动把 `manifest.xml` 里的地址改成你的域名即可。

### 方案二：本地服务器（无需账号，但需常驻运行）

Office 要求本地也走 HTTPS 且证书受信任。

1. **准备受信任的本地证书**
   ```bash
   npx --yes office-addin-dev-certs install
   ```
   （备选：`openssl req -x509 -newkey rsa:2048 -nodes -keyout localhost.key -out localhost.crt -days 365 -subj "/CN=localhost"`，再把 `localhost.crt` 导入「受信任的根证书颁发机构」。）
2. **启动服务器**
   ```bash
   node server.js
   ```
   访问 https://localhost:3000/src/taskpane.html 确认可正常打开（浏览器提示证书风险属正常）。
3. **旁加载**：用 **`manifest-local.xml`** 走上面的「上传我的加载项」步骤。
   ⚠️ 每次使用加载项前都要先启动 `node server.js`（保持窗口运行）。

---

### 一键校验与发布（publish.js）

项目自带 `publish.js`，自动完成完整性校验并（可选）推送到 git：
```bash
node publish.js --check   # 仅校验：JS 语法 / JSON / 两份 manifest XML / KB·RIBBON↔快捷键↔功能区 一致性
node publish.js --dry     # 校验后打印将执行的 git 命令（不真正推送）
node publish.js           # 校验通过即 git add / commit / push 到 origin
```
校验不通过时不会推送，可放心在每次改动后运行。

## 自定义指令

1. 在 `src/taskpane.js` 中新增一个 `async function doXxx()` 命令实现（按 `Office.context.host` 区分主机，仅对该主机执行、其余静默跳过）。
2. 若需键盘快捷键：在 `shortcuts.json` 的 `actions` 与 `shortcuts` 中分别登记动作 id 与按键（含 Mac 变体）；并在 `taskpane.js` 的 `KB` 数组中登记 `["KbXxx", doXxx]`，`registerActions()` 会自动 `Office.actions.associate("KbXxx", wrapKb(doXxx))`。
3. 若需功能区按钮：在 `manifest.xml`（及 `manifest-local.xml`）对应主机的 `<Group>` 中加一个 `<Control xsi:type="Button|Menu">`，`Action` 用 `ExecuteFunction` 并指定 `FunctionName="RibbonXxx"`；再在 `taskpane.js` 的 `RIBBON` 数组中登记 `["RibbonXxx", doXxx]`，`registerActions()` 会自动 `Office.actions.associate("RibbonXxx", wrapRibbon(doXxx))`。
4. 在 `taskpane.js` 的 `COMMANDS` 数组中补充面板按钮（含 `hosts: ["word"]` / `["excel"]` / `["word","excel"]` 与可选的 `kb`），供命令面板调用与筛选。

> 修改清单后，两个 manifest 必须保持 `FunctionName` 集合一致（`manifest-local.xml` 由 `manifest.xml` 经 URL 替换自动再生，请勿手动分叉）。

---

## 说明 / 已知限制

- 键盘快捷键需要 Office 较新版本（Windows：Excel 2102+、Word 2408+；Mac 对应版本），并依赖共享运行时（已在清单中通过 `lifetime="long"` 启用）。
- 清单要求 `ExcelApi 1.9` 与 `WordApi 1.4` 作为最低版本（页码 `insertField` 需 1.4）；旧版 Office 可能无法加载（现代 Microsoft 365 订阅均满足）。
- **简繁转换**由 `src/lib/opencc-bundle.js`（opencc-js）提供，**离线内置完整词典**，可正确处理多义字（如「干→乾/幹」「里→裡/裏」按上下文判别），无需联网与任何运行时依赖。
- 「求和 / 平均值」在**选区每列正下方**写入公式；选区需为连续区域。
- 「按颜色求和」以活动单元格填充色为基准，扫描工作表已用区域中同色数值单元格求和，结果写入活动单元格右侧；选区空白或白色时请先点选一个带色单元格。
- 「文本分列」按指定分隔符将选区**首列**拆到右侧多列，原列保留；分隔符支持逗号、空格、Tab、竖线。
- 「数据有效性」为选区添加下拉列表，选项需用英文逗号分隔输入（如 `合格,不合格,待复检`）。
- 「删除空白行 / 删除空段落 / 清理多余空格」作用于**当前选区**，未选中时作用于整篇文档 / 整个已用区域。
- 「导出 PDF」依赖 Office 的 `getFileAsync(FileType.Pdf)`，桌面版支持良好；个别网页版环境可能不支持，失败时面板会给出提示。
- 命令按主机区分：仅 Excel 的命令在 Word 下静默跳过，反之亦然。
- `assets/` 下的图标为占位图，正式分发前可替换为自己的品牌图标（需 16/32/80 像素 PNG）。

---

## English summary

**Quick Commands** is an Office Add-in (Office Web Add-in, XML manifest + HTML/JS + Office.js) that brings a command palette, keyboard shortcuts, and one-click actions to **Word and Excel** (Windows / Mac / web).

- **84 commands** across shared, Word-only, and Excel-only categories (see tables above); a sidebar palette filters by the current host.
- **Deployment without a server**: run `node setup.js` to point `manifest.xml` at your GitHub Pages site, push to GitHub, enable Pages, then sideload `manifest.xml` once.
- **Local fallback**: `node server.js` + `manifest-local.xml` (needs a trusted localhost HTTPS cert).
- **One-click publish**: `node publish.js` validates JS syntax, `shortcuts.json`, both manifests (XML), and id consistency (KB / RIBBON ↔ shortcuts ↔ ribbon buttons), then optionally pushes to git.
- **Simplified ↔ Traditional Chinese** conversion uses the bundled `src/lib/opencc-bundle.js` (opencc-js, full offline dictionary, context-aware for polysemous characters).
- Icons in `assets/` are placeholders; replace with your own 16/32/80 PNGs before distribution.
