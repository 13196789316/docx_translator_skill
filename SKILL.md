---
name: docx-bilingual-translator
description: A high-fidelity Word document bilingual translator with automatic layout protection and professional terminology mapping.
---

# 📖 Word 高保真双语翻译器技能指南

本技能（Skill）提供了一套无损的 Word（.docx）文档“中英双语对照化”全自动解决方案。通过在原有中文行下方**追加格式化英文 runs** 的方式，实现完美保真排版（图片、字体加粗、表格合并等格式不受损）。

---

## ⚡ 极速入门 (TL;DR)

本技能的运转由三大脚本协作完成，已被收录在桌面 `docx_translator_skill/` 中。

只需在终端运行以下一行组合命令，即可直接将您的 Word 文档转化为**格式精美、对照整齐的双语手册**：

```bash
# 1. 进入技能文件夹
cd ~/Desktop/docx_translator_skill

# 2. 一键执行合并并注入双语文本 (默认读取桌面的 manual.docx，并输出为 manual_bilingual.docx)
python3 scripts/merge.py && python3 scripts/inject.py
```

---

## 📂 技能目录结构与渐进式披露

为了让您和任何未来的 AI 助手极速上手且保持视界清爽，本技能采用**渐进式披露原则**进行文件夹架构设计：

```
docx_translator_skill/
├── SKILL.md                          # 👈 【本文件】第一层：最简化主入口与核心操作 (极简披露)
├── README.md                         # 👈 【使用说明】第一层：极速上手与配色微调说明书
├── scripts/                          # 👈 【核心工具】第二层：各司其职的模块化运行脚本 (按需查阅)
│   ├── extract.py                    # 文本无损提取器 (支持参数)
│   ├── merge.py                      # 翻译分片合并器 (支持参数)
│   └── inject.py                     # 双语样式注入器 (支持参数)
├── examples/                         # 👈 【参考实例】第二层：标准格式输入输出数据 (按需查阅)
│   ├── sample_source.json            # 提取出的待译中文源数据
│   └── sample_result.json            # 翻译完毕的完整英汉映射数据
└── references/                       # 👈 【深度技术】第三层：隐去的底层避坑机制与高阶文档 (深层披露)
    ├── table_deduplication.md        # 表格物理合并单元格去重底层原理
    └── style_customization.md        # Word 段落 Runs 字体、字号与颜色自定义细节
```

---

## 🛠️ 高级配置：随心所欲定制英文字体与颜色

`inject.py` 支持高度灵活的命令行参数，您可以自由调整英文的大小、颜色和斜体样式：

| 命令行参数 | 默认值 | 作用与示例 |
| :--- | :--- | :--- |
| `-i`, `--input` | `~/Desktop/manual.docx` | 输入的原始 DOCX 文件路径 |
| `-t`, `--translation` | `examples/sample_result.json` | 翻译好的 JSON 映射路径 |
| `-o`, `--output` | `~/Desktop/manual_bilingual.docx` | 输出的双语对照版 DOCX 路径 |
| `--para-size` | `10.0` | 正文中英文的字号大小（单位：pt） |
| `--table-size` | `9.5` | 表格中英文的字号大小（单位：pt） |
| `--color` | `"31,78,121"` | 英文的 RGB 颜色（如 `"100,100,100"` 对应极简灰色） |
| `--no-italic` | `False` | 禁用斜体字（默认英文使用斜体字进行视觉区分） |

### 🎨 示例：生成极简灰色、常规字体的双语手册
```bash
python3 scripts/inject.py \
  --input "~/Desktop/my_document.docx" \
  --output "~/Desktop/my_document_bilingual.docx" \
  --para-size 10.5 \
  --table-size 10.0 \
  --color "100,100,100" \
  --no-italic
```

---

## 📚 深度参考与技术底层

如需探究本技能底层最核心的技术细节和机制，请随时按需查阅以下技术文档：

1.  **[合并单元格全局去重算法](file:///Users/shenweitao/Desktop/docx_translator_skill/references/table_deduplication.md)**：解密 python-docx 遍历合并单元格时产生多重文本的物理机制，以及我们引入 `_tc` 句柄进行全局去重的高效方案。
2.  **[Run 级字符样式无损渲染](file:///Users/shenweitao/Desktop/docx_translator_skill/references/style_customization.md)**：解析为什么 `add_run()` 方式能够在换行的同时绝对保护段落原始字体大小、粗体加粗、高亮等样式不受侵害。
