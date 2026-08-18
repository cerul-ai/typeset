# typeset

把 Markdown 生成排版规范的中文 `.docx`——合同、协议、服务确认单、方案书、正式函件。

这是一个 [Agent Skill](https://code.claude.com/docs/en/skills)，给 Claude Code / Claude 用。也可以直接当命令行工具使。

## 为什么需要它

`pandoc x.md -o x.docx` 能跑通，但产物是给英文博客用的：

| | pandoc 默认 | 中文正式文书需要 |
| --- | --- | --- |
| 纸型 | Letter | A4 |
| 标题 | `#0F4761` 青蓝、不加粗 | 纯黑、加粗、黑体 |
| 西文字体 | Aptos | Times New Roman |
| 中文字体 | 主题回退，跨机器不一致 | 宋体 / 黑体显式指定 |
| 行距 | 单倍 | 1.5 倍（留批注空间） |
| 页眉页脚 | 无 | 横线 + 第 X 页 / 共 Y 页 |

更麻烦的是三类只有打开 Word 才看得见的问题：表格跨页被劈成两半、条款标题孤零零留在页底、签章区甲方在上一页乙方在下一页。

## 用法

```bash
brew install pandoc poppler
brew install --cask libreoffice

python3 scripts/build.py 合同.md                # 生成
python3 scripts/verify.py 合同.docx --render    # 结构 lint + 转 PDF + 逐页渲染
```

然后**真的去看渲染出来的 jpg**。前面每一步都通过、文档在 Word 里能打开，版面依然可能很难看。

注意 LibreOffice 匹配不到「宋体 / 黑体」，会回退到 Arial Unicode MS。替换字的度量不同，**连页数都会变**——同一份合同 Word 出 17 页、LibreOffice 出 27 页。所以 LibreOffice 只用来验结构（表格有没有被劈开、签章区有没有散），**最终分页与字体外观要在 Word 里确认**。`verify.py` 会自动检测字体有没有被替换并提示。

从模板起步：

```bash
cp templates/contract.md 我的合同.md
```

## 四套封面与签章方案

```bash
python3 scripts/build.py 合同.md --all    # 四套各出一份，渲染封面页挑一个
```

| 方案 | 特点 |
| --- | --- |
| A 复刻参考版 | 国内公司合同最常见的范式，默认选它 |
| B 严格对齐版 | 标签-值两列 + 填空线，适合打印手填 |
| C 公文庄重版 | 信息块加外框、标题放大，最有仪式感 |
| D 现代简洁版 | 左对齐 + 细线包夹，最好看但无页眉线 |

## 目录

```
SKILL.md                      给模型看的入口
scripts/build.py              Markdown → .docx
scripts/verify.py             结构 lint + 渲染检查
references/house-style.md     排版参数与取值依据
references/openxml-gotchas.md XSD 元素顺序、分页控制、页眉页脚注入
references/contract-zh.md     中文合同的结构与惯例
templates/contract.md         合同骨架
```

## 安装

SKILL.md 是 [Agent Skills 开放标准](https://agentskills.io)，Claude Code、Codex CLI、
Gemini CLI、Cursor 等都能读。两个工具的发现路径不同：

| 工具 | 路径 |
| --- | --- |
| Codex CLI | `~/.agents/skills/` |
| Claude Code | `~/.claude/skills/` |

装一份、两处都能用：

```bash
git clone https://github.com/cerul-ai/typeset.git ~/.agents/skills/typeset
ln -s ../../.agents/skills/typeset ~/.claude/skills/typeset
```

之后跟 Claude 或 Codex 说"帮我把这份合同做成 Word"就会自动用上。

## 改风格

排版参数集中在 `scripts/build.py` 顶部的常量块。想换仿宋、换字号、换边距，改那里，不要散到各处。取值依据见 `references/house-style.md`。
