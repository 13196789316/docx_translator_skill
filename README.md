# 📖 Word 高保真双语对照翻译工具包

本工具包（Skill）是一套专门针对 **Microsoft Word (.docx)** 文档的**中英双语对照自动生成与美化工具**。

它最初专为《印尼服装 SRM 系统操作手册》定制开发，内置了高水平的服装供应链与 SRM 行业术语翻译。由于采用了通用的 Word 底层 XML 渲染机制，它同样能**无损复用于您的任何其他 Word 手册或文档**（完美保护图片、列表层级、局部粗体、跨行/跨列合并表格等排版格式）。

---

## ⚡ 快速开始：一键生成双语手册

如果您想直接将桌面的操作手册生成为精美的双语对照版本，只需在您的终端（Terminal）执行以下两步：

### 1. 安装基础依赖
在 Mac 终端运行以下命令安装 Word 处理库：
```bash
pip install python-docx
```

### 2. 进入目录并一键执行
```bash
# 进入本文件夹
cd /Users/shenweitao/Desktop/docx_translator_skill

# 合并翻译并注入 Word 样式
python3 scripts/merge.py && python3 scripts/inject.py
```
> 🎉 **生成结果**：生成的双对照手册将自动保存在您的桌面上：`/Users/shenweitao/Desktop/印尼服装SRM系统操作手册_双语.docx`。

---

## 🎨 进阶玩法：随心所欲定制英文外观

为了防止中英文混在一起导致视觉疲劳，工具包默认将英文设置为**字号小一号**、**优雅的暗蓝色**以及**斜体字**。

您可以通过给脚本传入不同的参数，一键换装您喜欢的样式：

| 方案 | 风格质感 | 一键运行命令 |
| :--- | :--- | :--- |
| **方案 A** | **优雅暗蓝** (默认) | `python3 scripts/inject.py --color "31,78,121"` |
| **方案 B** | **极简莫兰迪灰** | `python3 scripts/inject.py --color "100,100,100"` |
| **方案 C** | **科技感浅青绿** | `python3 scripts/inject.py --color "0,128,128" --no-italic` |
| **方案 D** | **经典纯黑** | `python3 scripts/inject.py --color "0,0,0" --no-italic` |

*注：您也可以使用 `--para-size`（正文字号）和 `--table-size`（表格字号）参数来微调大小。*

---

## 📂 渐进式目录向导

根据**渐进式披露原则**，我们将工具包的组件进行了清晰的分层，您可以按需深入查阅：

*   📁 **[scripts/](file:///Users/shenweitao/Desktop/docx_translator_skill/scripts/) (核心工具箱)**：
    *   `extract.py`：负责把 Word 里的文字抽成 JSON。
    *   `merge.py`：把多卷翻译无缝拼合。
    *   `inject.py`：负责将翻译写回 Word 并应用漂亮样式。
*   📁 **[examples/](file:///Users/shenweitao/Desktop/docx_translator_skill/examples/) (数据样例)**：
    *   收录了本次提取出的 348 个中文段落源数据和最终翻译完成的中英映射 JSON 文件。
*   📁 **[references/](file:///Users/shenweitao/Desktop/docx_translator_skill/references/) (底层技术白皮书)**：
    *   **[合并单元格防重复指南](file:///Users/shenweitao/Desktop/docx_translator_skill/references/table_deduplication.md)**：为什么普通脚本操作合并单元格会追加多次英文？带您了解我们是如何使用物理 `_tc` 句柄进行全局表格去重的。
    *   **[Runs 文字块无损控制细节](file:///Users/shenweitao/Desktop/docx_translator_skill/references/style_customization.md)**：揭秘为什么本工具能在换行追加英文的同时，绝对保护您中文原有的高亮、加粗、特殊颜色等格式不受一丝破坏。
*   📄 **[SKILL.md](file:///Users/shenweitao/Desktop/docx_translator_skill/SKILL.md)**：
    *   专门提供给其他人工智能助手（AI Agent）读取的技能主描述文件。
