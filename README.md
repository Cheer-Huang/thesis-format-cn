# thesis-format-cn

> 一键将 Word 文档格式化为符合中国高校学位论文规范的格式

## 为什么需要这个工具

用 AI 编程助手（CodeBuddy、Cursor、Copilot 等）或脚本一键生成 Word 报告时，产出的文档通常存在大量格式问题：

- 标题字体是默认蓝色而非黑色，字号不符合规范
- 正文没有首行缩进，行距随意
- 表格是全边框网格而非三线表
- 图表没有按章编号，缺少表题
- 参考文献没有悬挂缩进
- 没有页码、页眉、目录

手动逐项调整这些格式耗时且容易遗漏。**thesis-format-cn** 在生成初稿后一键运行，自动修正上述全部格式问题，让 AI 产出的报告直接达到学位论文提交标准。

## 两种使用方式

### 方式一：作为 AI Agent Skill（推荐）

本项目是一个标准 Skill 包，可安装到各类 AI 编程助手中。安装后，用 AI 生成 Word 报告，然后直接说"按清华论文格式排版"即可自动触发格式化，无需手动运行脚本。

#### 安装指南

**CodeBuddy：**

```bash
git clone https://github.com/Cheer-Huang/thesis-format-cn.git ~/.codebuddy/skills/thu-thesis-format
```

安装后在 CodeBuddy 对话中提到"论文格式"、"三线表"、"清华论文"等关键词即可自动触发。

**Claude Code：**

```bash
git clone https://github.com/Cheer-Huang/thesis-format-cn.git ~/.claude/skills/thu-thesis-format
```

在项目根目录的 `CLAUDE.md` 中添加引用：
```markdown
## Skills
- 论文格式化: ~/.claude/skills/thu-thesis-format/SKILL.md
```

**Trae / 其他 Agent：**

将本项目克隆到任意目录，在 Agent 的系统提示或项目配置中引用 `SKILL.md`：

```bash
git clone https://github.com/Cheer-Huang/thesis-format-cn.git ~/skills/thesis-format-cn
```

然后在 Agent 配置中添加：
```
当用户提到"论文格式"、"学位论文"、"毕业论文"、"三线表"、"thesis format"等关键词时，
参考 ~/skills/thesis-format-cn/SKILL.md 中的指引，运行 scripts/ 下的脚本完成格式化。
```

**通用方式（任何 Agent 均可）：**

直接告诉 Agent：
> 请阅读 `~/skills/thesis-format-cn/SKILL.md`，按照其中的工作流程帮我格式化论文。

### 方式二：命令行工具

作为独立的 Python 命令行工具使用，适用于任何工作流：

```bash
pip install python-docx
python scripts/format_thu_thesis.py input.docx output.docx
```

基于 [python-docx](https://github.com/python-openxml/python-docx) 实现，无需安装 Microsoft Word。

目前支持 **清华大学** 研究生学位论文格式（依据《清华大学研究生学位论文写作指南》202605 版），架构设计支持扩展其他高校规范。

## 功能一览

| 格式要素 | 规范要求 | 自动处理 |
|---------|---------|---------|
| 页面设置 | A4，页边距 3cm，页眉页脚 2.2cm | ✅ |
| 章标题 | 三号(16pt)黑体居中，"第X章"格式 | ✅ |
| 节标题 | 四号(14pt)黑体居左 | ✅ |
| 正文 | 小四(12pt)宋体，首行缩进2字符，固定值20磅 | ✅ |
| 图题 | 11pt宋体居中，按章编号(图X.Y) | ✅ |
| 表题 | 11pt宋体居中，按章编号(表X.Y) | ✅ |
| 三线表 | 上下线1.5磅，表头线1磅，无竖线 | ✅ |
| 参考文献 | 五号(10.5pt)，悬挂缩进2字符 | ✅ |
| 字体颜色 | 全文统一黑色 | ✅ |
| 页眉页脚 | 五号居中篇眉 + 页码 | ✅ |
| 目录 | TOC域自动生成 | ✅ |
| 分页符 | 每章另起页 | ✅ |
| 格式审计 | 逐项检查并输出报告 | ✅ |

## 推荐工作流

格式化是**最后一步**，请在内容全部定稿后再运行。正确的使用顺序：

```
1. AI 生成 / 编辑内容（增删段落、修改文字、插入图表）
2. 内容定稿，确认无需再改
3. 运行格式化脚本（本工具）
4. 在 Word 中右键目录"更新域"
5. 提交
```

> **重要**：格式化后请勿再让 AI 编辑文档内容。AI 编辑器在插入新段落、修改文字时可能继承默认样式，破坏已设置的字体、缩进、行距等格式。如需修改内容，请先改好再重新运行格式化脚本。

## 快速开始

### 安装依赖

```bash
pip install python-docx
```

### 一键格式化

```bash
python scripts/format_thu_thesis.py input.docx output.docx
```

### 格式审计

```bash
python scripts/audit_thu_thesis.py input.docx
```

### 在 Python 中调用

```python
from scripts.format_thu_thesis import format_thu_thesis

# 一键格式化
format_thu_thesis("input.docx", "output.docx")

# 或按需调用单个函数
from docx import Document
from scripts.format_thu_thesis import fix_page_setup, fix_three_line_tables

doc = Document("input.docx")
fix_page_setup(doc)
fix_three_line_tables(doc)
doc.save("output.docx")
```

### 格式化后

打开 Word 文档，右键目录区域选择 **"更新域"** 即可生成目录内容。

## 格式规范速查

详细规范见 [references/thu_format_spec.md](references/thu_format_spec.md)。

| 元素 | 中文字体 | 英文字体 | 字号 | 对齐 | 行距 | 段前 | 段后 |
|------|---------|---------|------|------|------|------|------|
| 章标题 | 黑体 | Arial | 16pt | 居中 | 单倍 | 24磅 | 18磅 |
| 节标题 | 黑体 | Arial | 14pt | 居左 | 固定值20磅 | 24磅 | 6磅 |
| 正文 | 宋体 | Times New Roman | 12pt | 两端对齐 | 固定值20磅 | 0 | 0 |
| 图题 | 宋体 | Times New Roman | 11pt | 居中 | 单倍 | 6磅 | 12磅 |
| 表题 | 宋体 | Times New Roman | 11pt | 居中 | 单倍 | 12磅 | 6磅 |
| 参考文献 | 宋体 | Times New Roman | 10.5pt | 两端对齐 | 固定值16磅 | 3磅 | 0 |

三线表规范：
- 顶线（上边线）：1.5 磅
- 表头线（第三条线）：1 磅
- 底线（下边线）：1.5 磅
- 无竖线，中间行无横线

## 支持的高校

| 高校 | 状态 | 规范来源 |
|------|------|---------|
| 清华大学 | ✅ 已支持 | 研究生学位论文写作指南（202605）|
| 其他高校 | 🔜 规划中 | 欢迎贡献 |

### 添加新高校支持

1. 在 `references/` 下添加该校的格式规范文档
2. 在 `scripts/` 下添加对应的 `format_xxx.py` 脚本
3. 参考清华规范的实现方式，调整字体、字号、间距等参数

欢迎通过 PR 贡献新高校的格式支持！

## 技术细节

### python-docx 单位换算

| 单位 | 换算关系 |
|------|---------|
| 1 磅 (pt) | 12700 EMU / 20 twips |
| 1 cm | 360000 EMU |
| 1 汉字符 | 约 12pt = 240 twips |
| 字号 sz 属性 | 半磅，如 16pt → sz=32 |
| 边框 sz 属性 | 八分之一磅，如 1.5磅 → sz=12 |

### 常见坑

1. **Heading 样式蓝色**：Word 内置 Heading 1/2/3 样式默认有蓝色 (365F91, 4F81BD)，必须同时修改样式定义和 run 级别颜色
2. **Pt(0) 是 falsy**：python-docx 中 `Pt(0)` 为 falsy 值，审计时需直接检查 XML 的 `w:spacing` 属性
3. **TOC 域需手动更新**：python-docx 无法自动生成目录内容，插入 TOC 域后需在 Word 中右键"更新域"
4. **PAGE 域**：页码通过 Word PAGE 域实现，python-docx 检测文本为空但 Word 中正常显示

## 免责声明

本项目为独立开源工具，与清华大学官方无任何关联。项目中的格式参数基于公开的《研究生学位论文写作指南》提取，仅供参考。实际提交论文时，请以学校官方发布的最新指南和院系要求为准。如格式规范有更新，欢迎提交 Issue 或 PR。

## License

MIT License - 详见 [LICENSE](LICENSE)

## 致谢

- 格式规范来源：清华大学研究生院公开发布的《研究生学位论文写作指南》
- 参考实现：[thuthesis](https://github.com/tuna/thuthesis) LaTeX 模板（MIT/LPPL 许可）
- 依赖库：[python-docx](https://github.com/python-openxml/python-docx)
