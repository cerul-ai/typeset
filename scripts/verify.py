#!/usr/bin/env python3
"""检查生成的 .docx：结构 lint + 转 PDF + 逐页渲染 + 版面体检。

排版问题几乎都是"看了才知道"的——表格被劈成两半、标题孤零零留在页底、
签章区甲乙方分到两页。所以这个脚本的重点不是"生成成功了吗"，而是
"把每一页变成图片，让你真的看一眼"。

用法：
    python3 scripts/verify.py 合同.docx              # lint + 版面体检
    python3 scripts/verify.py 合同.docx --render     # 另外转 PDF 并渲染成图片
"""
import argparse
import pathlib
import shutil
import subprocess
import sys
import xml.etree.ElementTree as ET
import zipfile

W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"

# XSD 强制的子元素顺序。写错了 Word 会报"文档已损坏"，而且不告诉你是顺序问题。
PPR_ORDER = ["pStyle", "keepNext", "keepLines", "pageBreakBefore", "framePr", "widowControl",
             "numPr", "suppressLineNumbers", "pBdr", "shd", "tabs", "suppressAutoHyphens",
             "kinsoku", "wordWrap", "overflowPunct", "topLinePunct", "autoSpaceDE",
             "autoSpaceDN", "bidi", "adjustRightInd", "snapToGrid", "spacing", "ind",
             "contextualSpacing", "mirrorIndents", "suppressOverlap", "jc", "textDirection",
             "textAlignment", "textboxTightWrap", "outlineLvl", "divId", "cnfStyle", "rPr",
             "sectPr", "pPrChange"]
STYLE_ORDER = ["name", "aliases", "basedOn", "next", "link", "autoRedefine", "hidden",
               "uiPriority", "semiHidden", "unhideWhenUsed", "qFormat", "locked", "personal",
               "personalCompose", "personalReply", "rsid", "pPr", "rPr", "tblPr", "trPr",
               "tcPr", "tblStylePr"]


def lint(path):
    """检查 pPr 与 style 的子元素顺序。这是最容易犯、最难 debug 的错。

    必须按**直接子元素**判断。用正则抓标签会把孙元素也算进来——比如
    <w:style> 里 <w:tblStylePr> 内嵌的 <w:pPr>，会被误判成 style 的子元素
    排在 tblStylePr 之后。所以这里用 ElementTree 走真正的树结构。
    """
    problems = []
    checks = {"pPr": PPR_ORDER, "style": STYLE_ORDER, "tblStylePr": STYLE_ORDER}
    with zipfile.ZipFile(path) as z:
        for part in ("word/document.xml", "word/styles.xml"):
            if part not in z.namelist():
                continue
            root = ET.fromstring(z.read(part))
            for el in root.iter():
                order = checks.get(el.tag.replace(W, ""))
                if not order:
                    continue
                seen, idx = [], -1
                for child in el:
                    name = child.tag.replace(W, "")
                    if name not in order:
                        continue
                    i = order.index(name)
                    if i < idx:
                        problems.append(f"{part}: <w:{el.tag.replace(W, '')}> 里 "
                                        f"<w:{name}> 排在 <w:{seen[-1]}> 之后，应在其之前")
                        break
                    idx = i
                    seen.append(name)
    return problems


def to_pdf(path):
    """LibreOffice 优先；macOS 上没有就调 Word。

    Word 走 AppleScript 有两个坑：默认 AppleEvent 超时只有 60 秒（大文档或 Word
    正忙时会报 -1712），而且它是 GUI 应用，会短暂抢占屏幕。所以要用
    `with timeout` 包住，并且首选 LibreOffice。
    """
    out = path.with_suffix(".pdf")
    out.unlink(missing_ok=True)

    if shutil.which("soffice"):
        r = subprocess.run(["soffice", "--headless", "--convert-to", "pdf",
                            "--outdir", str(path.parent), str(path)],
                           capture_output=True, text=True, timeout=300)
        if out.exists():
            return out
        print(f"LibreOffice 转换失败：{r.stderr.strip()[:200]}")

    if sys.platform == "darwin" and pathlib.Path("/Applications/Microsoft Word.app").exists():
        script = f'''with timeout of 600 seconds
  tell application "Microsoft Word"
    set d to open file name POSIX file "{path.resolve()}" as string
    save as d file name (POSIX file "{out.resolve()}" as string) file format format PDF
    close d saving no
  end tell
end timeout'''
        try:
            r = subprocess.run(["osascript", "-e", script],
                               capture_output=True, text=True, timeout=660)
        except subprocess.TimeoutExpired:
            print("Word 转换超时。它可能弹了对话框——切到 Word 看一眼，或改用 LibreOffice。")
            return None
        if out.exists():
            return out
        print(f"Word 转换失败：{r.stderr.strip()[:200]}")
    return None


def survey(pdf):
    """逐页统计行数与首行，用来发现"半空的页""被劈开的表格"。"""
    txt = subprocess.run(["pdftotext", "-layout", str(pdf), "-"],
                         capture_output=True, text=True).stdout
    pages = txt.split("\f")[:-1]
    print(f"\n共 {len(pages)} 页")
    for i, p in enumerate(pages, 1):
        lines = [l.strip() for l in p.split("\n") if l.strip()]
        flag = "  ← 偏空，检查是否被强行分页" if 0 < len(lines) < 10 else ""
        print(f"  p{i:02d} [{len(lines):2d} 行] {(lines[0][:46] if lines else '(空)')}{flag}")
    return len(pages)


def check_fonts(pdf):
    """检查渲染时字体是否被替换。

    LibreOffice 匹配不到「宋体 / 黑体」这类中文字体名（即使系统里有），会回退到
    Arial Unicode MS 之类。替换字的度量不同，**页数会显著变化**——实测同一份合同
    Word 出 17 页、LibreOffice 出 27 页。所以字体一旦被替换，渲染图只能用来看
    "机制有没有生效"（表格行有没有被劈开、签章区有没有散），不能用来判断分页。
    """
    if not shutil.which("pdffonts"):
        return
    out = subprocess.run(["pdffonts", str(pdf)], capture_output=True, text=True).stdout
    want = {"宋体": ("SimSun", "Songti"), "黑体": ("SimHei", "Heiti")}
    missing = [zh for zh, names in want.items() if not any(n in out for n in names)]
    if not missing:
        print("\n✓ 中文字体解析正确，渲染图与 Word 里的样子一致")
        return
    print(f"\n⚠ 渲染器没解析出 {' / '.join(missing)}，已回退到替代字体。这意味着：")
    print("   · 页数与分页位置**不代表** Word 里的结果（实测可差 50% 以上）")
    print("   · 仍可验证：表格行是否被劈开、签章区是否完整、内容有无丢失")
    print("   · 字体外观与最终分页，请在 Microsoft Word 里确认——它自带 SimSun/SimHei，")
    print("     与对方 Windows Word 看到的一致")


def main():
    ap = argparse.ArgumentParser(description="检查 .docx 的结构与版面")
    ap.add_argument("docx")
    ap.add_argument("--render", action="store_true", help="转 PDF 并把每页渲染成 jpg")
    ap.add_argument("--dpi", type=int, default=100)
    a = ap.parse_args()
    path = pathlib.Path(a.docx)

    problems = lint(path)
    if problems:
        print("✗ 元素顺序有问题（Word 可能报文档已损坏）：")
        for p in problems[:10]:
            print("   ", p)
    else:
        print("✓ pPr / style 子元素顺序正常")

    if not a.render:
        return
    pdf = to_pdf(path)
    if not pdf:
        print("✗ 无法转 PDF：装 LibreOffice（brew install --cask libreoffice）或 Microsoft Word")
        return
    n = survey(pdf)
    if shutil.which("pdftoppm"):
        prefix = path.parent / f"{path.stem}-p"
        subprocess.run(["pdftoppm", "-jpeg", "-r", str(a.dpi), str(pdf), str(prefix)], check=True)
        print(f"\n已渲染 {n} 页 → {prefix}-*.jpg")
        print("现在真的去看几页图，尤其是表格页与签章页。")
    else:
        print("\n装 poppler 才能渲染成图片：brew install poppler")
    check_fonts(pdf)


if __name__ == "__main__":
    main()
