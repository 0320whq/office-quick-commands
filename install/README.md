# 快捷指令 QuickCommands - VBA 版

适用于所有 Office 桌面版（2007-2019-365），不依赖 Web Add-in。

## 包含内容

| 文件 | 说明 |
|------|------|
| `word.bas` | Word VBA 宏源码（18个快捷指令） |
| `excel.bas` | Excel VBA 宏源码（22个快捷指令） |
| `install.cmd` | 一键配置脚本（设置宏安全+信任位置+引导操作） |
| `word-ribbon.xml` | Word 功能区 XML（可选，用于自定义 Ribbon） |
| `excel-ribbon.xml` | Excel 功能区 XML（可选，用于自定义 Ribbon） |

## 安装方法

### 方式一：使用安装脚本（推荐）

1. 双击运行 `install.cmd`
2. 脚本会自动配置宏安全设置和信任位置
3. 按提示打开 Word/Excel 导入 `.bas` 文件
4. 重启 Office 即可使用

### 方式二：手动安装

#### Word

1. 打开 Word，按 `Alt+F11` 打开 VBA 编辑器
2. 左侧"工程资源管理器"中找到 **Normal** 项目
3. 右键 → **导入文件** → 选择 `word.bas`
4. `Ctrl+S` 保存
5. 关闭 Word，重新打开
6. 顶部出现 QuickCommands 工具栏

#### Excel

1. 打开 Excel，按 `Alt+F11` 打开 VBA 编辑器
2. 如果左侧没有 **VBAProject(PERSONAL.XLSB)**：
   - 视图 → 宏 → 录制宏
   - "保存在"选 **个人宏工作簿**
   - 点确定后立即停止录制
3. 右键 VBAProject(PERSONAL.XLSB) → **导入文件** → 选择 `excel.bas`
4. `Ctrl+S` 保存
5. 关闭 Excel，重新打开
6. 顶部出现 QuickCommands 工具栏

## 可用指令

### Word（18个）

| 按钮 | 功能 |
|------|------|
| 日期时间 | 插入当前日期时间 |
| 加粗/斜体/下划线 | 切换格式 |
| 居中/左对齐/右对齐/两端对齐 | 段落对齐 |
| 高亮 | 切换黄色高亮 |
| 清格式 | 清除所有格式 |
| 删空段 | 删除空白段落 |
| 合并空格 | 多个空格合并为一个 |
| 页码 | 插入页码 |
| 导出PDF | 导出为 PDF |
| 标题1/标题2/正文 | 应用样式 |
| 项目符号/编号 | 列表格式 |

### Excel（22个）

| 按钮 | 功能 |
|------|------|
| 日期 | 插入当前日期 |
| 加粗/斜体/下划线 | 切换格式 |
| 居中/左对齐/右对齐 | 单元格对齐 |
| 高亮 | 切换黄色填充 |
| 清格式/清内容 | 清除 |
| 插行/插列 | 插入行列 |
| 升序/降序 | 排序 |
| 求和/求均值 | 插入公式 |
| 转值 | 公式转数值 |
| 去重 | 删除重复项 |
| 筛选 | 切换筛选 |
| 冻结首行 | 冻结窗格 |
| 导出CSV/导出PDF | 格式转换 |

## 卸载方法

### Word
1. `Alt+F11` → Normal → QuickCommands → 右键 → 移除
2. `Ctrl+S` 保存

### Excel
1. `Alt+F11` → VBAProject(PERSONAL.XLSB) → QuickCommands → 右键 → 移除
2. `Ctrl+S` 保存

## 技术说明

- **不依赖** Web Add-in，兼容 Office 2007 及以上所有桌面版本
- 使用 `CommandBars` API 创建工具栏，`AutoExec`/`Auto_Open` 实现自动加载
- 不需要修改注册表中的 AccessVBOM 设置
- 不需要 VSTO 或任何运行时环境
- 纯 VBA，零依赖

## 与 Web Add-in 版本的关系

本项目同时提供 Web Add-in 版本（见 `../src/`），适用于 Office 365 和 Office 2019（需正确配置旁加载）。VBA 版本是 Web Add-in 不可用时的替代方案，覆盖所有桌面 Office 版本。
