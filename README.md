# better-docx

**[English](README_EN.md) | 中文**

> 基于 [python-docx](https://github.com/python-openxml/python-docx) 的 Word 文档高级处理库 —— 用 42 个函数覆盖「表格、图片、排版、分节、目录、解析、定位」全流程，产出**真正的原生 Word 对象**（可编辑表格 `<w:tbl>`、嵌入图片），而不是文本符号模拟。

[![Python](https://img.shields.io/badge/Python-3.8%2B-blue)](https://www.python.org/)
[![License](https://img.shields.io/badge/License-MIT-green)](LICENSE)
[![Version](https://img.shields.io/badge/version-1.0.0-brightgreen)](pyproject.toml)

---

## 为什么做这个

`python-docx` 是底层库：很多常见需求得直接操作 OOXML XML 才能实现，而且坑极多（字体回退、底纹全黑、插图报错、分页失控……）。better-docx 把这些踩过的坑封装成**即用函数**，让你像操作高层对象一样写 Word：

| 需求 | python-docx 原生 | better-docx |
|---|---|---|
| 中文字体 | 必须手动设 `w:eastAsia`，否则宋体回退成默认西文字体 | `set_font(run, "宋体")` 一键双设 |
| 单元格底纹 | 手写 `w:shd`，`val` 写错整个变黑 | `shade_cell(cell, "D9E2F3")` |
| 动态页码 | 手写 begin/instrText/end 三段式域 | `add_page_number_field(footer.paragraphs[0])` |
| 自动目录 | 无法直接生成 | `add_toc_field(doc)` 插入 TOC 域，Word 里 F9 更新 |
| 横竖混排 | 分节 + 手动交换宽高 + 隐藏陷阱 | `set_orientation(sec, landscape=True)` 幂等 |
| 插图 | `add_picture` 老报 `part` 错误 | `add_image(doc, "x.png")` 一次搞定 |
| 定位段落 | 只能线性遍历 | `find_heading(doc, "第三章 结论")` 按标题锚点 |

## 特性

- **原生表格**：边框样式、单元格底纹、合并单元格、固定列宽、跨页重复表头、行不拆散
- **图片插入**：正文 / 表格单元格插图、居中题注、尺寸控制、安全替换旧图
- **中文排版**：宋体 / 黑体双字体槽一键设置、首行缩进 2 字符（字符单位，随字号缩放）、行距与段间距
- **页面布局**：多节横竖混排、各节独立页眉页脚、首页 / 奇偶页不同、PAGE / NUMPAGES 动态页码域、A4 页边距
- **段落列表**：项目符号、数字编号、多级嵌套（内置样式体系）
- **样式系统**：Heading 1–9 大纲级别、全局样式覆盖、自定义样式、TOC 域自动目录
- **导航感知**：docx → JSON 树 / Markdown、锚点定位、定点插入段落 / 表格、书签与超链接读写
- **排错工具**：分页异常（一页一段）的诊断与修复

## 安装

```bash
pip install better-docx
```

或从源码安装（开发模式）：

```bash
git clone https://github.com/Wersky/better-docx.git
cd better-docx
pip install -e .
```

仅需 `python-docx>=1.1.0` 一个依赖。

## 快速开始

```python
from docx import Document
from better_docx import set_font, shade_cell, set_cell_text, add_toc_field, add_page_number_field, set_orientation

doc = Document()
sec = doc.sections[0]
set_orientation(sec, landscape=False)                 # A4 竖版（自动交换宽高，幂等）
add_page_number_field(sec.footer.paragraphs[0])       # 页脚动态页码「第 X 页」
add_toc_field(doc)                                    # 目录域（正文用 Heading 后 F9 生成）

doc.add_heading("一、概述", level=1)                   # 进目录的大纲标题
t = doc.add_table(rows=2, cols=2)                     # 原生表格
t.style = "Table Grid"
set_cell_text(t.rows[0].cells[0], "表头", bold=True)   # 居中 + 垂直居中
shade_cell(t.rows[0].cells[0], "D9E2F3")              # 表头浅蓝底纹

doc.save("report.docx")
```

## 功能概览（42 个函数，三个模块）

`from better_docx import *` 扁平导出全部函数，也可按子模块导入。

### `better_docx.helpers` —— 基础元素

| 函数 | 作用 |
|---|---|
| `set_font(run, name, size, bold, italic, color)` | 设置字体（中文自动补 eastAsia） |
| `add_para(doc, text, ...)` / `add_heading(doc, text, level)` | 段落 / 黑体标题 |
| `shade_cell(cell, hex)` | 单元格底纹 |
| `set_cell_text(cell, text, ...)` / `add_para_to_cell(cell, ...)` | 单元格写文本 / 追加段落 |
| `new_image_paragraph(doc)` / `add_image(doc, path, ...)` / `add_image_to_cell(cell, ...)` | 插图（正文 / 单元格） |
| `remove_images_from_paragraph(p)` / `global_replace(doc, old, new)` | 替换旧图 / 全局保格式替换 |
| `set_cell_widths(table, widths)` / `set_repeat_header(table)` / `set_row_no_split(row)` | 列宽 / 跨页表头 / 行不拆散 |
| `diagnose_pagination(doc)` / `fix_one_paragraph_per_page(doc)` | 分页异常诊断 / 修复 |

### `better_docx.layout` —— 页面与排版

| 函数 | 作用 |
|---|---|
| `set_page_a4(section, margins)` | A4 纸张 + 页边距 |
| `set_orientation(section, landscape)` | 横 / 竖切换（幂等，自动交换宽高） |
| `enable_even_odd_headers(section)` | 启用奇偶页不同（XML 级） |
| `add_page_number_field(paragraph, ...)` / `add_numpages_field(paragraph, ...)` | PAGE / NUMPAGES 动态页码域 |
| `set_first_line_indent_chars(paragraph, chars)` | 中文首行缩进 N 字符 |
| `set_line_spacing(paragraph, value, rule)` | 行距（单倍 / 1.5 / 固定值） |
| `add_bullets(doc, items)` / `add_numbered(doc, pairs)` | 项目符号 / 编号（含嵌套） |
| `restyle_heading(style, ...)` | 覆盖内置 Heading 样式 |
| `add_toc_field(doc, levels)` | 插入 TOC 域（Word 打开 F9 生成目录） |

### `better_docx.navigator` —— 解析与定位

| 函数 | 作用 |
|---|---|
| `docx_to_json(doc)` / `docx_to_markdown(doc)` | 全文档结构化解析 |
| `build_heading_index(doc)` / `find_heading(doc, text)` / `find_paragraph_containing(doc, sub)` | 锚点定位 |
| `insert_paragraph_after(p, ...)` / `insert_paragraph_before_cjk(p, ...)` | 定点插段落 |
| `insert_table_before(p, ...)` / `insert_table_after(p, ...)` | 定点插表格 |
| `add_bookmark(p, name)` / `find_bookmark_paragraph(doc, name)` / `append_text_to_bookmark(doc, name, text)` | 书签读写 |
| `add_hyperlink(p, url, text)` / `list_hyperlinks(doc)` / `list_bookmarks(doc)` | 超链接 / 书签枚举 |

## 完整示例

生成一份**多节研究报告**（封面 → 目录 → 正文 → 横向附录 → 竖向附录），见 [`examples/make_report.py`](examples/make_report.py)：

```python
from docx import Document
from docx.enum.section import WD_SECTION
from better_docx import *

doc = Document()
add_toc_field(doc)                      # 目录域

sec1 = doc.sections[0]                  # 竖向节：封面(首页无页眉) + 正文
set_page_a4(sec1)
sec1.different_first_page_header_footer = True
enable_even_odd_headers(sec1)
sec1.header.is_linked_to_previous = False
sec1.header.paragraphs[0].text = "2026 研究报告"
add_page_number_field(sec1.footer.paragraphs[0]); add_numpages_field(sec1.footer.paragraphs[0])

sec2 = doc.add_section(WD_SECTION.NEW_PAGE)   # 横向节：放宽表格
set_page_a4(sec2); set_orientation(sec2, landscape=True)
sec2.header.is_linked_to_previous = False
sec2.header.paragraphs[0].text = "附录 A · 横向数据表"

# ... 正文标题、缩进、列表、表格、图片、书签、超链接

# 用 navigator 定位并精确修改（Agent 式闭环）
docx = Document("report.docx")
assert find_heading(docx, "三、结论") is not None
insert_table_before(find_heading(docx, "三、结论"), rows=2, cols=3)
print(docx_to_markdown(docx)[:200])     # 解析回 Markdown 预览
```

## 测试

内置测试覆盖全部 42 个函数（含边界与偏执场景），零额外依赖即可运行（推荐配合 pytest）：

```bash
pip install -e ".[dev]"     # 装 pytest
pytest tests/ -v            # 或 python -m pytest tests/ -v
```

## 设计说明

- **真正的原生对象**：表格是 `<w:tbl>`、图片是嵌入 drawing、页码是域代码（`fldChar`），Word 打开后可正常编辑、更新、打印，绝非 markdown 符号拼凑。
- **中文字体双设**：所有涉及字体的函数都同时设置西文槽与 `w:eastAsia` 槽，规避中文回退问题。
- **幂等与健壮**：方向切换、奇偶页启用等操作均幂等；图片缺失、样式名非法等边界不抛异常而是安全降级。
- **可解析**：内置 `docx_to_json` / `docx_to_markdown`，方便在 Agent 工作流里「感知 → 定位 → 修改 → 复查」。

## FAQ

**和 python-docx 的关系？** 它是 better-docx 的唯一依赖，本库是它的高层封装，不 fork、不魔改，可随时降级回原生 API。

**支持中文吗？** 内置对中文的一等支持：宋体 / 黑体 / 仿宋双字体槽、字符单位缩进、中文页码格式（`CHINESENUM1`）。

**能读已有文档吗？** 能。`docx_to_json` / `docx_to_markdown` 解析任意 docx；`find_heading` / `find_paragraph_containing` 定位；注意 `.doc`（旧二进制）需先转 `.docx`。

## License

[MIT](LICENSE) © 2026 Wersky