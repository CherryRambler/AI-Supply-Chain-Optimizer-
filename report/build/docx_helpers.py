"""python-docx helpers that implement the SRM report formatting guidelines."""
import copy, re, subprocess, tempfile, os, zipfile
from lxml import etree
from docx import Document
from docx.shared import Pt, Inches, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_TAB_ALIGNMENT, WD_BREAK
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.section import WD_SECTION
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

TNR = "Times New Roman"
BLUE = RGBColor(0x00, 0x70, 0xC0)
TEXT_W = 5.76          # inches: A4 width 8.27 − 1.5 − 1.0
J, C, L, R = (WD_ALIGN_PARAGRAPH.JUSTIFY, WD_ALIGN_PARAGRAPH.CENTER,
              WD_ALIGN_PARAGRAPH.LEFT, WD_ALIGN_PARAGRAPH.RIGHT)


def set_font(run, size=12, bold=None, italic=None, color=None, name=TNR):
    run.font.size = Pt(size)
    run.font.name = name
    rpr = run._r.get_or_add_rPr()
    rf = rpr.find(qn("w:rFonts"))
    if rf is None:
        rf = OxmlElement("w:rFonts"); rpr.insert(0, rf)
    for a in ("w:ascii", "w:hAnsi", "w:cs", "w:eastAsia"):
        rf.set(qn(a), name)
    for a in ("w:asciiTheme", "w:hAnsiTheme", "w:cstheme", "w:eastAsiaTheme"):
        if rf.get(qn(a)) is not None:
            del rf.attrib[qn(a)]
    if bold is not None: run.font.bold = bold
    if italic is not None: run.font.italic = italic
    run.font.color.rgb = color if color is not None else RGBColor(0, 0, 0)
    return run


def fmt(p, align=J, ls=1.5, before=0, after=6, keep_next=False, page_break=False, left=None, first=None, hanging=None, right=None):
    f = p.paragraph_format
    f.alignment = align; f.line_spacing = ls
    f.space_before = Pt(before); f.space_after = Pt(after)
    f.keep_with_next = keep_next
    f.page_break_before = page_break
    f.widow_control = True
    if left is not None: f.left_indent = Inches(left)
    if right is not None: f.right_indent = Inches(right)
    if first is not None: f.first_line_indent = Inches(first)
    if hanging is not None:
        f.left_indent = Inches(hanging); f.first_line_indent = Inches(-hanging)
    return p


TOKEN = re.compile(r"(\*\*.+?\*\*|\*[^*\s][^*]*?\*|_\{.+?\}|\^\{.+?\}|`.+?`)")


def add_rich(p, text, size=12, bold=False, color=None):
    """**bold**, *italic*, _{sub}, ^{sup}, `code`."""
    for tok in TOKEN.split(text):
        if not tok:
            continue
        b, i, sub, sup, code = bold, False, False, False, False
        if tok.startswith("**"):
            tok, b = tok[2:-2], True
        elif tok.startswith("*") and tok.endswith("*") and len(tok) > 2:
            tok, i = tok[1:-1], True
        elif tok.startswith("_{"):
            tok, sub, i = tok[2:-1], True, True
        elif tok.startswith("^{"):
            tok, sup = tok[2:-1], True
        elif tok.startswith("`"):
            tok, code = tok[1:-1], True
        r = p.add_run(tok)
        set_font(r, size - 1 if code else size, bold=b, italic=i, color=color, name="Consolas" if code else TNR)
        if sub: r.font.subscript = True
        if sup: r.font.superscript = True
    return p


def page_field(paragraph, size=12):
    r = paragraph.add_run(); set_font(r, size)
    for tag, attr in (("w:fldChar", {"w:fldCharType": "begin"}), ("w:instrText", None), ("w:fldChar", {"w:fldCharType": "separate"})):
        el = OxmlElement(tag)
        if attr:
            for k, v in attr.items(): el.set(qn(k), v)
        else:
            el.set(qn("xml:space"), "preserve"); el.text = " PAGE "
        r._r.append(el)
    r2 = paragraph.add_run("1"); set_font(r2, size)
    r3 = paragraph.add_run(); set_font(r3, size)
    e = OxmlElement("w:fldChar"); e.set(qn("w:fldCharType"), "end"); r3._r.append(e)


def set_cell_bg(cell, hexcolor):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd"); shd.set(qn("w:val"), "clear"); shd.set(qn("w:color"), "auto"); shd.set(qn("w:fill"), hexcolor)
    tcPr.append(shd)


def set_cell_margins(table, top=40, bottom=40, left=80, right=80):
    tblPr = table._tbl.tblPr
    m = OxmlElement("w:tblCellMar")
    for k, v in (("top", top), ("left", left), ("bottom", bottom), ("right", right)):
        e = OxmlElement(f"w:{k}"); e.set(qn("w:w"), str(v)); e.set(qn("w:type"), "dxa"); m.append(e)
    tblPr.append(m)


def set_borders(table, color="808080", sz=4):
    tblPr = table._tbl.tblPr
    b = OxmlElement("w:tblBorders")
    for k in ("top", "left", "bottom", "right", "insideH", "insideV"):
        e = OxmlElement(f"w:{k}"); e.set(qn("w:val"), "single"); e.set(qn("w:sz"), str(sz))
        e.set(qn("w:space"), "0"); e.set(qn("w:color"), color); b.append(e)
    tblPr.append(b)


def latex_to_omml(latex_list, workdir):
    """Convert LaTeX equations to OMML elements with pandoc (one docx round-trip)."""
    md = "\n\n".join(f"$${x}$$" for x in latex_list)
    src, out = os.path.join(workdir, "eq.md"), os.path.join(workdir, "eq.docx")
    open(src, "w").write(md)
    subprocess.run(["pandoc", src, "-o", out], check=True)
    xml = zipfile.ZipFile(out).read("word/document.xml")
    root = etree.fromstring(xml)
    M = "http://schemas.openxmlformats.org/officeDocument/2006/math"
    paras = root.findall(f".//{{{M}}}oMathPara")
    maths = [p.find(f"{{{M}}}oMath") for p in paras]
    assert len(maths) == len(latex_list), (len(maths), len(latex_list))
    return maths
