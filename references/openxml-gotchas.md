# OpenXML 陷阱

`.docx` 是一个 ZIP，里面是一堆 XML。只改 `scripts/build.py` 的常量不需要读这篇；要手写或修改 OpenXML 才需要。

## 一、子元素顺序由 XSD 强制

**这是最容易犯、最难 debug 的错。** 顺序写错，Word 打开报"文档已损坏"或"部分内容有问题"，而错误信息**完全不提元素顺序**。

### `<w:pPr>` 的顺序

```
pStyle → keepNext → keepLines → pageBreakBefore → framePr → widowControl →
numPr → suppressLineNumbers → pBdr → shd → tabs → … → spacing → ind →
contextualSpacing → … → jc → textDirection → textAlignment → … → outlineLvl → rPr
```

常见错误：

| 错误写法 | 问题 |
| --- | --- |
| `<w:pPr><w:spacing/><w:keepNext/>` | `spacing` 必须在 `keepNext` 之后 |
| `<w:pPr><w:jc/><w:ind/>` | `jc` 必须在 `ind` 之后 |
| `<w:pPr><w:shd/><w:pBdr/>` | `pBdr` 必须在 `shd` 之前 |

**别在已有的 `pPr` 里插元素**——你不知道它原本有什么、顺序如何。整段重建：

```python
new_ppr = ppr(keepNext=True, spacing=spacing(280, 140), jc=jc("center"))
blk = re.sub(r"<w:pPr>.*?</w:pPr>", new_ppr, blk, count=1, flags=re.S)
```

### `<w:style>` 的顺序

```
name → aliases → basedOn → next → link → … → qFormat → locked → … →
pPr → rPr → tblPr → trPr → tcPr → tblStylePr
```

常见错误：把 `tblPr` 写在 `pPr` 前面。表格样式的正确骨架：

```xml
<w:style w:type="table" w:styleId="Table">
  <w:name w:val="Table"/>
  <w:pPr>…</w:pPr>
  <w:rPr>…</w:rPr>
  <w:tblPr>…</w:tblPr>
  <w:trPr><w:cantSplit/></w:trPr>
  <w:tblStylePr w:type="firstRow">
    <w:pPr>…</w:pPr><w:rPr>…</w:rPr><w:tcPr>…</w:tcPr>
  </w:tblStylePr>
</w:style>
```

`<w:tblStylePr>` 内部也有顺序：`pPr → rPr → tblPr → trPr → tcPr`。

`scripts/verify.py` 的 lint 会检查这两类顺序，生成后先跑一遍。

## 二、分页控制

| 元素 | 作用 | 放哪 |
| --- | --- | --- |
| `<w:keepNext/>` | 本段与下一段同页 | 段落 pPr |
| `<w:keepLines/>` | 本段内部不分页 | 段落 pPr |
| `<w:widowControl/>` | 禁止单行孤行落在页首页尾 | docDefaults 全局开 |
| `<w:cantSplit/>` | 表格行不跨页断开 | `<w:trPr>` |
| `<w:tblHeader/>` | 表头在续页重复 | 首行 `<w:trPr>` |

**签章区整块 keepNext。** 甲方组在上一页、乙方组在下一页是硬伤。给块内除最后一段外的所有段落加 `keepNext`，Word 放不下就会把整块推到下一页而不是拆开。最后一段不能加，否则会粘住后文。

**引导句 keepNext。** 「三、费用：」留在页底、(a)(b)(c) 在下一页，看起来很随意。

**长表格独占一页**：`cantSplit` 只保证单行不劈开，整张表仍可能跨页。要整表一页就在表前插分页符。

分页符（放在 markdown 里的 raw block）：

```
```{=openxml}
<w:p><w:r><w:br w:type="page"/></w:r></w:p>
```
```

## 三、页眉页脚要改四个地方

只写 `header1.xml` 不够，Word 不认。四处都要动：

1. `word/header1.xml` — 内容本身
2. `[Content_Types].xml` — 加 `<Override PartName="/word/header1.xml" ContentType="…wordprocessingml.header+xml"/>`
3. `word/_rels/document.xml.rels` — 加 `<Relationship Id="rIdX" Type="…/header" Target="header1.xml"/>`
4. `word/document.xml` 的 `<w:sectPr>` — 加 `<w:headerReference w:type="default" r:id="rIdX"/>`

pandoc 默认的 reference.docx 里 `<w:sectPr />` 是**空的**，可以整个替换掉来一次注入纸型、边距、页眉页脚引用。注意 `document.xml` 根元素必须已声明 `xmlns:r`（pandoc 的默认文档已声明）。

页码用域而不是文本：

```xml
<w:r><w:fldChar w:fldCharType="begin"/></w:r>
<w:r><w:instrText xml:space="preserve">PAGE</w:instrText></w:r>
<w:r><w:fldChar w:fldCharType="end"/></w:r>
```

`NUMPAGES` 同理。页数变了 Word 自动更新。

## 四、字体的两层机制

字体有两条路径，都要处理：

1. **显式指定**：`<w:rFonts w:ascii="…" w:hAnsi="…" w:eastAsia="…" w:cs="…"/>`。中文和西文分开指定，各用各的。
2. **主题回退**：没显式指定的元素走 `word/theme/theme1.xml` 的 `majorFont`（标题）/ `minorFont`（正文）。pandoc 默认这两个的 `<a:ea>` 是空的，中文会走系统回退，跨机器不一致。

两层都设才稳。`build.py` 里 `patch_styles()` 管第一层、主题 patch 管第二层。

## 五、pandoc 的两个坑

**`- (a) 文字` 变成空项目符号 + 嵌套列表。** `fancy_lists` 扩展把 `(a)` 当有序列表标记。转义成 `\(a\)` 就不会，配合自定义悬挂缩进样式：

```markdown
::: {custom-style="SubItem"}
\(a\) 本条所称……
:::
```

**智能引号破坏中文引号。** `"文字"` 经 smart 扩展后开引号可能也变成右引号。直接写 `“文字”`。

## 六、raw OpenXML 块

pandoc 的 markdown 支持直接嵌入 OpenXML：

```
```{=openxml}
<w:p>…</w:p>
```
```

用来做封面、签章区这类版式精确的东西——比用 markdown 表达再靠样式修正可靠得多。注意：

- 表格（`<w:tbl>`）之后必须跟一个 `<w:p/>`，否则 Word 可能报错
- 块内不能用 markdown 语法，`**加粗**` 要写成 `<w:b/>`
- 自定义段落样式用 `::: {custom-style="X"}` div，能保留 markdown 行内格式，比 raw block 更适合有格式的正文
