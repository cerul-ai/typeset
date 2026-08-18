#!/usr/bin/env python3
"""Markdown → 中文正式文书排版的 .docx。

pandoc 自带的 reference.docx 是给英文博客用的：标题是 #0F4761 青蓝色、20/16/14pt
且不加粗，纸型 Letter，西文字体 Aptos。直接拿来做中文合同会很难看。这个脚本重建
一份 reference.docx，把版式换成中文正式文书的通行参数（见 references/house-style.md），
再把封面与签章区按选定方案生成为 raw OpenXML。

用法：
    python3 scripts/build.py 合同.md                 # 默认方案 A
    python3 scripts/build.py 合同.md --style C       # 换封面/签章方案
    python3 scripts/build.py 合同.md -o 输出.docx
    python3 scripts/build.py 合同.md --all           # 四套方案各出一份

markdown 里的三个占位符由本脚本填充：
    @@COVER@@       封面（随后自动分页）
    @@SIGNATURE@@   签章区（可出现多次）
    @@PAGEBREAK@@   手动分页

封面信息取自 markdown 的 YAML frontmatter，缺省值见 DEFAULT_META。
"""
import argparse
import pathlib
import re
import shutil
import subprocess
import sys
import zipfile

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
FOOTER_RID, HEADER_RID = "rIdZhFooter", "rIdZhHeader"

# ── 版式参数（twips；1cm = 567）见 references/house-style.md ──────────
PAGE = dict(w=11906, h=16838)                                   # A4
MARGIN = dict(top=1440, bottom=1440, left=1701, right=1701,     # 上下2.54 左右3.00cm
              header=1134, footer=850)                          # 页眉2.0 页脚1.5cm
CONTENT_W = PAGE["w"] - MARGIN["left"] - MARGIN["right"]
BODY_SZ, TABLE_SZ, SMALL_SZ = 24, 21, 18   # 半磅：12pt 正文 / 10.5pt 表格 / 9pt 页脚
LINE = 360                                  # 1.5 倍行距
EA_BODY, EA_HEAD, LATIN = "宋体", "黑体", "Times New Roman"
HEAD_SZ = {1: 32, 2: 28, 3: 24}             # 16 / 14 / 12 pt
RULE, SHADE_HEAD, SHADE_CALL = "BFBFBF", "F2F2F2", "F7F7F7"

DEFAULT_META = dict(
    title="", party_a="【甲方全称】", party_b="【乙方全称】",
    party_a_label="甲方", party_b_label="乙方", style="A",
)

# CT_PPr 的子元素顺序由 XSD 强制。顺序写错 Word 会报"文档已损坏"，而且报错信息
# 完全不提元素顺序，极难 debug —— 所有 pPr 都必须经 ppr() 生成。
_PPR_ORDER = ("keepNext", "keepLines", "widowControl", "pBdr", "shd", "tabs",
              "spacing", "ind", "jc", "outlineLvl")


# ══ XML 小工具 ═══════════════════════════════════════════════════════
def rpr(sz=BODY_SZ, ea=EA_BODY, bold=False, color="000000"):
    b = "<w:b/>" if bold else ""
    return (f'<w:rPr><w:rFonts w:ascii="{LATIN}" w:hAnsi="{LATIN}" w:eastAsia="{ea}" '
            f'w:cs="{LATIN}"/>{b}<w:color w:val="{color}"/>'
            f'<w:sz w:val="{sz}"/><w:szCs w:val="{sz}"/></w:rPr>')


def ppr(**parts):
    out = []
    for tag in _PPR_ORDER:
        v = parts.get(tag)
        if v is None or v is False:
            continue
        out.append(f"<w:{tag}/>" if v is True else v)
    return f'<w:pPr>{"".join(out)}</w:pPr>'


def spacing(before=0, after=0, line=LINE):
    return f'<w:spacing w:before="{before}" w:after="{after}" w:line="{line}" w:lineRule="auto"/>'


def run(text, **kw):
    return f'<w:r>{rpr(**kw)}<w:t xml:space="preserve">{text}</w:t></w:r>'


def para(*runs, **pprkw):
    return f'<w:p>{ppr(**pprkw)}{"".join(runs)}</w:p>'


def blank(n=1):
    return para(spacing=spacing()) * n


def jc(v):
    return f'<w:jc w:val="{v}"/>'


def ind(left=None, right=None, hanging=None):
    a = "".join(f' w:{k}="{v}"' for k, v in
                (("left", left), ("right", right), ("hanging", hanging)) if v is not None)
    return f"<w:ind{a}/>"


def border(edge, sz=6, color="000000", space=1, val="single"):
    return f'<w:{edge} w:val="{val}" w:sz="{sz}" w:space="{space}" w:color="{color}"/>'


def cell(content, w, borders=None, valign="center"):
    b = f"<w:tcBorders>{borders}</w:tcBorders>" if borders else ""
    return (f'<w:tc><w:tcPr><w:tcW w:w="{w}" w:type="dxa"/>{b}'
            f'<w:vAlign w:val="{valign}"/></w:tcPr>{content}</w:tc>')


def table(rows, widths, align="center", borders=None):
    edges = ("top", "left", "bottom", "right", "insideH", "insideV")
    tb = borders or "".join(border(e, val="none", sz=0) for e in edges)
    grid = "".join(f'<w:gridCol w:w="{w}"/>' for w in widths)
    # cantSplit：整行不跨页断开，长表格才不会把一行劈成两半
    trs = "".join(f"<w:tr><w:trPr><w:cantSplit/></w:trPr>{r}</w:tr>" for r in rows)
    return (f'<w:tbl><w:tblPr><w:tblW w:w="{sum(widths)}" w:type="dxa"/>{jc(align)}'
            f'<w:tblBorders>{tb}</w:tblBorders><w:tblLayout w:type="fixed"/></w:tblPr>'
            f"<w:tblGrid>{grid}</w:tblGrid>{trs}</w:tbl>") + para()


UL = "＿" * 12                       # 全角下划线，在宋体里比 ASCII 下划线整齐
DATE_BLANK = "　　　　　年　　　月　　　日"
NO_TEXT_BELOW = "（本行以下无正文，仅供签章之用）"


# ══ 四套封面 ═════════════════════════════════════════════════════════
# 中文合同封面靠「不同对齐」区分信息组：编号靠右、缔约方靠左且互相对齐、
# 日期单独下沉。全部居中反而会把这种区分抹平。
def cover_A(m):
    """复刻常见范式：编号靠右、甲乙左对齐、日期下沉。"""
    return (blank(2)
            + para(run("合同编号：【　　　　　　　　】"), jc=jc("right"), spacing=spacing(after=120))
            + blank(4)
            + para(run(f'{m["party_a_label"]}：{m["party_a"]}', bold=True), spacing=spacing(after=240))
            + blank(1)
            + para(run(f'{m["party_b_label"]}：{m["party_b"]}', bold=True), spacing=spacing(after=240))
            + blank(5)
            + para(run(f"签订日期：【{DATE_BLANK}】"), ind=ind(left=1400)))


def cover_B(m):
    """严格对齐：标签与值分两列，值带填空下边线。"""
    lw, vw = 1900, 5200
    under = border("bottom", sz=6, color="808080")
    rows = [cell(para(run(lab, bold=True), spacing=spacing(60, 60)), lw)
            + cell(para(run(val), spacing=spacing(60, 60)), vw, borders=under)
            for lab, val in (("合同编号", "　"),
                             (m["party_a_label"], m["party_a"]),
                             (m["party_b_label"], m["party_b"]),
                             ("签订日期", DATE_BLANK))]
    return blank(3) + table(rows, [lw, vw])


def cover_C(m):
    """公文庄重：信息块整体加外框。配合更大的标题字号使用。"""
    box_w = 6600
    inner = "".join(para(run(f"{lab}：{val}", bold=b), spacing=spacing(80, 80))
                    for lab, val, b in (("合同编号", UL, False),
                                        (m["party_a_label"], m["party_a"], True),
                                        (m["party_b_label"], m["party_b"], True),
                                        ("签订日期", DATE_BLANK, False)))
    frame = "".join(border(e, sz=8) for e in ("top", "left", "bottom", "right"))
    return (blank(2) + table([cell(inner, box_w, borders=frame, valign="top")], [box_w])
            + blank(3)
            + para(run("（本合同经双方签字并盖章后生效）", sz=TABLE_SZ, color="595959"), jc=jc("center")))


def cover_D(m):
    """现代简洁：左对齐信息块，上下细线包夹。"""
    rule = para(pBdr=f'<w:pBdr>{border("bottom", sz=4, color="A6A6A6")}</w:pBdr>',
                spacing=spacing(after=200))
    body = "".join(para(run(lab, bold=True), run(f"　　{val}"), spacing=spacing(after=140))
                   for lab, val in (("合同编号", "【　　　　　　　　】"),
                                    (m["party_a_label"], m["party_a"]),
                                    (m["party_b_label"], m["party_b"]),
                                    ("签订日期", f"【{DATE_BLANK}】")))
    return blank(3) + rule + body + rule


# ══ 四套签章区 ═══════════════════════════════════════════════════════
# 签章区最怕被分页劈成两半（甲方在上一页、乙方在下一页）。整块 keepNext 之后，
# Word 放不下就会把整块推到下一页，而不是拆开。
def _stacked(m, fills=("　", "　"), seal=False):
    specs = [((), dict(spacing=spacing()))] * 2
    specs.append(((run(NO_TEXT_BELOW, sz=TABLE_SZ),), dict(spacing=spacing(after=360))))
    for label, fill in ((m["party_a_label"], fills[0]), (m["party_b_label"], fills[1])):
        for lab, val in ((label, fill), ("授权代表", fill), ("日期", DATE_BLANK)):
            specs.append(((run(f"{lab}：", bold=(lab == label)), run(val)),
                          dict(spacing=spacing(after=200))))
        if seal:
            specs.append(((run("（此处加盖公章）", sz=TABLE_SZ, color="595959"),),
                          dict(ind=ind(left=567), spacing=spacing(after=120))))
        specs += [((), dict(spacing=spacing()))] * 2

    out = ""
    for i, (runs, kw) in enumerate(specs):
        if i < len(specs) - 1:          # 末段不加 keepNext，否则会粘住后文
            kw = {**kw, "keepNext": True}
        out += para(*runs, **kw)
    return out


def sig_A(m):
    """竖排，仅标签，留白供盖章。"""
    return _stacked(m)


def sig_B(m):
    """无框两列表格，甲乙并排，字段带填空线。"""
    cw = CONTENT_W // 2
    under = border("bottom", sz=6, color="808080")
    rows = [cell(para(run(f'{m["party_a_label"]}（签字 / 盖章）', bold=True), spacing=spacing(after=120)), cw)
            + cell(para(run(f'{m["party_b_label"]}（签字 / 盖章）', bold=True), spacing=spacing(after=120)), cw)]
    for lab in ("单位名称", "授权代表", "日期"):
        c = cell(para(run(f"{lab}：{UL}"), spacing=spacing(160, 160)), cw, borders=under)
        rows.append(c + c)
    return (para(spacing=spacing(), keepNext=True) * 2
            + para(run(NO_TEXT_BELOW, sz=TABLE_SZ), spacing=spacing(after=360), keepNext=True)
            + table(rows, [cw, cw], align="left"))


def sig_C(m):
    """竖排 + 独立盖章区提示。"""
    return _stacked(m, fills=(UL, UL), seal=True)


def sig_D(m):
    """竖排 + 行内（盖章）标注。"""
    return _stacked(m, fills=(f"{UL}　（盖章）", f"{UL}　（盖章）"))


STYLES = {
    "A": dict(name="复刻参考版", cover=cover_A, sig=sig_A, header_rule=True, title_sz=32),
    "B": dict(name="严格对齐版", cover=cover_B, sig=sig_B, header_rule=True, title_sz=32),
    "C": dict(name="公文庄重版", cover=cover_C, sig=sig_C, header_rule=True, title_sz=44),
    "D": dict(name="现代简洁版", cover=cover_D, sig=sig_D, header_rule=False, title_sz=32),
}


# ══ 页眉 / 页脚 / 节属性 ═══════════════════════════════════════════════
def footer_xml():
    """第 X 页 / 共 Y 页。合同该有页码，防抽换页。"""
    def fld(x):
        return f"<w:r>{rpr(sz=SMALL_SZ)}{x}</w:r>"
    parts = [run("第 ", sz=SMALL_SZ)]
    for f in ("PAGE", "NUMPAGES"):
        parts += [fld('<w:fldChar w:fldCharType="begin"/>'),
                  fld(f'<w:instrText xml:space="preserve">{f}</w:instrText>'),
                  fld('<w:fldChar w:fldCharType="end"/>')]
        parts.append(run(" 页 / 共 " if f == "PAGE" else " 页", sz=SMALL_SZ))
    return (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<w:ftr xmlns:w="{W}">'
            f'{para(*parts, jc=jc("center"), spacing=spacing(line=240))}</w:ftr>')


def header_xml(rule):
    p = para(pBdr=f'<w:pBdr>{border("bottom", sz=6)}</w:pBdr>', spacing=spacing(line=240)) \
        if rule else para(spacing=spacing(line=240))
    return f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<w:hdr xmlns:w="{W}">{p}</w:hdr>'


SECTPR = (
    "<w:sectPr>"
    f'<w:headerReference w:type="default" r:id="{HEADER_RID}"/>'
    f'<w:footerReference w:type="default" r:id="{FOOTER_RID}"/>'
    f'<w:pgSz w:w="{PAGE["w"]}" w:h="{PAGE["h"]}"/>'
    f'<w:pgMar w:top="{MARGIN["top"]}" w:right="{MARGIN["right"]}" w:bottom="{MARGIN["bottom"]}"'
    f' w:left="{MARGIN["left"]}" w:header="{MARGIN["header"]}" w:footer="{MARGIN["footer"]}" w:gutter="0"/>'
    '<w:docGrid w:type="lines" w:linePitch="312"/></w:sectPr>'
)

EXTRA_STYLES = (
    # (a)(b)(c) 子项：悬挂缩进，折行文字对齐到标号之后
    '<w:style w:type="paragraph" w:customStyle="1" w:styleId="SubItem">'
    '<w:name w:val="SubItem"/><w:basedOn w:val="Normal"/>'
    + ppr(spacing=spacing(60, 60), ind=ind(left=840, hanging=480)) + "</w:style>"
    # 引导句 keepNext：防止「三、费用：」被单独留在页底、列表却在下一页
    '<w:style w:type="paragraph" w:customStyle="1" w:styleId="LeadIn">'
    '<w:name w:val="LeadIn"/><w:basedOn w:val="Normal"/>'
    + ppr(keepNext=True, spacing=spacing(after=120)) + "</w:style>"
    # 整段不跨页：用于价格、赔偿上限这类不能被劈开的条款
    '<w:style w:type="paragraph" w:customStyle="1" w:styleId="Together">'
    '<w:name w:val="Together"/><w:basedOn w:val="Normal"/>'
    + ppr(keepLines=True, spacing=spacing(after=120)) + "</w:style>"
    # 封面信息行
    '<w:style w:type="paragraph" w:customStyle="1" w:styleId="Cover">'
    '<w:name w:val="Cover"/><w:basedOn w:val="Normal"/>'
    + ppr(spacing=spacing(120, 120), jc=jc("center")) + rpr(sz=28) + "</w:style>"
)


def patch_styles(path, title_sz):
    s = path.read_text()

    # 文档默认：宋体 12pt、1.5 倍行距、开 widowControl（禁止单行孤行落在页首页尾）
    s = re.sub(r"<w:docDefaults>.*?</w:docDefaults>",
               f"<w:docDefaults><w:rPrDefault>{rpr()}</w:rPrDefault><w:pPrDefault>"
               f"{ppr(widowControl=True, spacing=spacing(after=120))}</w:pPrDefault></w:docDefaults>",
               s, count=1, flags=re.S)

    # 标题：黑体、纯黑、加粗。pandoc 默认是 #0F4761 青蓝且不加粗，那是博客风格
    for lvl, sz in {**HEAD_SZ, 1: title_sz}.items():
        m = re.search(rf'<w:style [^>]*w:styleId="Heading{lvl}".*?</w:style>', s, re.S)
        if not m:
            continue
        blk = m.group(0)
        before, after = (360, 200) if lvl == 1 else (280, 140)
        new_ppr = ppr(keepNext=True, keepLines=True, spacing=spacing(before, after),
                      jc=jc("center") if lvl == 1 else None,
                      outlineLvl=f'<w:outlineLvl w:val="{lvl - 1}"/>')
        new_rpr = rpr(sz=sz, ea=EA_HEAD, bold=True)
        for pat, rep in ((r"<w:pPr>.*?</w:pPr>", new_ppr), (r"<w:rPr>.*?</w:rPr>", new_rpr)):
            blk2 = re.sub(pat, rep, blk, count=1, flags=re.S)
            blk = blk2 if blk2 != blk else blk.replace("</w:style>", rep + "</w:style>", 1)
        s = s.replace(m.group(0), blk, 1)

    # 表格：全边框 + 表头底纹 + 内收一号字 + 行不跨页
    m = re.search(r'<w:style w:type="table"[^>]*w:styleId="Table".*?</w:style>', s, re.S)
    if m:
        grid = "".join(border(e, sz=4, color=RULE, space=0)
                       for e in ("top", "left", "bottom", "right", "insideH", "insideV"))
        s = s.replace(m.group(0),
                      '<w:style w:type="table" w:styleId="Table"><w:name w:val="Table"/>'
                      # w:style 的子元素顺序：name → pPr → rPr → tblPr → trPr → tblStylePr
                      + ppr(spacing=spacing(20, 20, line=280)) + rpr(sz=TABLE_SZ)
                      + f"<w:tblPr><w:tblBorders>{grid}</w:tblBorders>"
                      '<w:tblCellMar><w:top w:w="60" w:type="dxa"/><w:left w:w="90" w:type="dxa"/>'
                      '<w:bottom w:w="60" w:type="dxa"/><w:right w:w="90" w:type="dxa"/>'
                      "</w:tblCellMar></w:tblPr><w:trPr><w:cantSplit/></w:trPr>"
                      '<w:tblStylePr w:type="firstRow">' + ppr(jc=jc("left"))
                      + rpr(sz=TABLE_SZ, ea=EA_HEAD, bold=True)
                      + f'<w:tcPr><w:shd w:val="clear" w:color="auto" w:fill="{SHADE_HEAD}"/></w:tcPr>'
                      "</w:tblStylePr></w:style>", 1)

    # 引用块：浅底 + 左侧粗边，用来突出声明性条款
    m = re.search(r'<w:style [^>]*w:styleId="BlockText".*?</w:style>', s, re.S)
    if m:
        s = s.replace(m.group(0),
                      '<w:style w:type="paragraph" w:styleId="BlockText">'
                      '<w:name w:val="Block Text"/><w:basedOn w:val="Normal"/>'
                      + ppr(pBdr=f'<w:pBdr>{border("left", sz=18, color="808080", space=8)}</w:pBdr>',
                            shd=f'<w:shd w:val="clear" w:color="auto" w:fill="{SHADE_CALL}"/>',
                            spacing=spacing(120, 120), ind=ind(left=284, right=284))
                      + rpr(sz=TABLE_SZ) + "</w:style>", 1)

    # 公式 / 代码块：浅底 + 缩进居中，别像贴上去的终端输出
    m = re.search(r'<w:style [^>]*w:styleId="SourceCode".*?</w:style>', s, re.S)
    if m:
        s = s.replace(m.group(0),
                      '<w:style w:type="paragraph" w:styleId="SourceCode">'
                      '<w:name w:val="Source Code"/><w:basedOn w:val="Normal"/>'
                      + ppr(keepNext=True,
                            shd=f'<w:shd w:val="clear" w:color="auto" w:fill="{SHADE_CALL}"/>',
                            spacing=spacing(120, 120, line=280),
                            ind=ind(left=567, right=567), jc=jc("center"))
                      + rpr(sz=TABLE_SZ) + "</w:style>", 1)

    path.write_text(s.replace("</w:styles>", EXTRA_STYLES + "</w:styles>", 1))


def build_reference(build_dir, style):
    work = build_dir / "ref"
    if work.exists():
        shutil.rmtree(work)
    work.mkdir(parents=True)

    default = build_dir / "pandoc-default.docx"
    if not default.exists():
        default.write_bytes(subprocess.run(
            ["pandoc", "--print-default-data-file", "reference.docx"],
            check=True, capture_output=True).stdout)
    with zipfile.ZipFile(default) as z:
        z.extractall(work)

    patch_styles(work / "word/styles.xml", style["title_sz"])

    # 主题字体兜底：未显式指定字体的元素走主题
    theme = work / "word/theme/theme1.xml"
    t = theme.read_text()
    for tag, ea in (("majorFont", EA_HEAD), ("minorFont", EA_BODY)):
        blk = re.search(rf"<a:{tag}>.*?</a:{tag}>", t, re.S).group(0)
        q = re.sub(r'<a:latin typeface="[^"]*"', f'<a:latin typeface="{LATIN}"', blk, count=1)
        q = re.sub(r'<a:ea typeface="[^"]*"', f'<a:ea typeface="{ea}"', q, count=1)
        t = t.replace(blk, q, 1)
    theme.write_text(t)

    (work / "word/footer1.xml").write_text(footer_xml())
    (work / "word/header1.xml").write_text(header_xml(style["header_rule"]))

    ct = work / "[Content_Types].xml"
    base = "application/vnd.openxmlformats-officedocument.wordprocessingml"
    ct.write_text(ct.read_text().replace(
        "</Types>",
        f'<Override PartName="/word/footer1.xml" ContentType="{base}.footer+xml"/>'
        f'<Override PartName="/word/header1.xml" ContentType="{base}.header+xml"/></Types>', 1))

    rels = work / "word/_rels/document.xml.rels"
    rt = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
    rels.write_text(rels.read_text().replace(
        "</Relationships>",
        f'<Relationship Id="{FOOTER_RID}" Type="{rt}/footer" Target="footer1.xml"/>'
        f'<Relationship Id="{HEADER_RID}" Type="{rt}/header" Target="header1.xml"/></Relationships>', 1))

    doc = work / "word/document.xml"
    d = doc.read_text()
    if "<w:sectPr />" not in d:
        sys.exit("pandoc 参考文档里找不到空 sectPr，无法注入 A4 与页眉页脚")
    doc.write_text(d.replace("<w:sectPr />", SECTPR, 1))

    ref = build_dir / "reference-zh.docx"
    ref.unlink(missing_ok=True)
    with zipfile.ZipFile(ref, "w", zipfile.ZIP_DEFLATED) as z:
        for f in sorted(work.rglob("*")):
            if f.is_file():
                z.write(f, f.relative_to(work).as_posix())
    return ref


# ══ markdown 预处理 ═══════════════════════════════════════════════════
def raw(xml):
    return f"```{{=openxml}}\n{xml}\n```"


PAGE_BREAK = raw('<w:p><w:r><w:br w:type="page"/></w:r></w:p>')
SUBITEM_OPEN = '::: {custom-style="SubItem"}'


def parse_meta(text):
    m = dict(DEFAULT_META)
    fm = re.match(r"^---\n(.*?)\n---\n", text, re.S)
    if not fm:
        return m, text
    for line in fm.group(1).split("\n"):
        if ":" in line:
            k, v = line.split(":", 1)
            k = k.strip()
            if k in m:
                m[k] = v.strip().strip("\"'")
    return m, text[fm.end():]


def preprocess(text, style, meta):
    # `- (a) 文字` 会被 pandoc 的 fancy_lists 当成有序列表标记，渲染成
    # "空项目符号 + 嵌套 (a) 列表"。转义括号并套悬挂缩进样式。
    text = re.sub(r"^- \(([a-z0-9]+)\) (.+)$",
                  lambda m: f'{SUBITEM_OPEN}\n\\({m.group(1)}\\) {m.group(2)}\n:::',
                  text, flags=re.M)

    lines, out = text.split("\n"), []
    for i, ln in enumerate(lines):
        t = ln.strip()
        if t == "@@COVER@@":
            out += ["", raw(style["cover"](meta)), "", PAGE_BREAK, "", raw(blank(1)), ""]
        elif t == "@@SIGNATURE@@":
            out += ["", raw(style["sig"](meta)), ""]
        elif t == "@@PAGEBREAK@@":
            out += ["", PAGE_BREAK, ""]
        else:
            # 子项组的引导句套 LeadIn，避免它单独留在页底
            nxt = next((x.strip() for x in lines[i + 1:] if x.strip()), "")
            if t and nxt == SUBITEM_OPEN and not t.startswith((":", "#", "|", "`", "-", ">")):
                out += ['::: {custom-style="LeadIn"}', ln, ":::"]
            else:
                out.append(ln)
    return "\n".join(out)


def build(src, style_key, out_path=None):
    style = STYLES[style_key]
    src = pathlib.Path(src)
    build_dir = src.parent / ".docx-build"
    build_dir.mkdir(exist_ok=True)

    meta, body = parse_meta(src.read_text())
    if meta["title"] and not body.lstrip().startswith("# "):
        body = f'# {meta["title"]}\n\n' + body
    md = build_dir / "preprocessed.md"
    md.write_text(preprocess(body, style, meta))

    ref = build_reference(build_dir, style)
    out = pathlib.Path(out_path) if out_path else src.with_suffix(".docx")
    if not out_path and len(STYLES) > 1 and style_key != "A":
        out = src.with_name(f'{src.stem}-方案{style_key}.docx')
    subprocess.run(["pandoc", str(md), "-o", str(out), "--reference-doc", str(ref),
                    "--from", "markdown", "--to", "docx"], check=True)
    print(f"OK  {out}  ({out.stat().st_size / 1024:.0f} KB)  方案{style_key}·{style['name']}")
    return out


def main():
    ap = argparse.ArgumentParser(description="Markdown → 中文排版 .docx")
    ap.add_argument("source", help="markdown 源文件")
    ap.add_argument("--style", choices=list(STYLES), help="封面/签章方案（默认取 frontmatter，再默认 A）")
    ap.add_argument("-o", "--output", help="输出路径")
    ap.add_argument("--all", action="store_true", help="四套方案各出一份")
    a = ap.parse_args()

    meta, _ = parse_meta(pathlib.Path(a.source).read_text())
    keys = list(STYLES) if a.all else [a.style or meta["style"]]
    for k in keys:
        build(a.source, k, a.output if len(keys) == 1 else None)


if __name__ == "__main__":
    main()
