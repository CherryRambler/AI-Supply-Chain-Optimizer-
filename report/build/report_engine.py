"""Report engine: loads the SRM template, keeps the cover page, and writes the report body."""
import copy, os, re
from docx import Document
from docx.shared import Pt, Inches, RGBColor
from docx.enum.text import WD_TAB_ALIGNMENT, WD_BREAK
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.section import WD_SECTION
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from docx_helpers import *

ROMAN_START = 5   # template: ABSTRACT = page v


class Report:
    def __init__(self, template, title):
        self.doc = d = Document(template)
        body = d.element.body
        cover_end = d.paragraphs[34]._p                     # holds the cover's section break
        node = cover_end.getnext()
        while node is not None:
            nxt = node.getnext()
            if node.tag != qn("w:sectPr"):
                body.remove(node)
            node = nxt
        # cover title (18 pt bold, centred)
        p = d.paragraphs[1]
        for r in p.runs[1:]:
            r._r.getparent().remove(r._r)
        p.runs[0].text = title
        set_font(p.runs[0], 18, bold=True, color=RGBColor(0, 0, 0))
        cover = d.paragraphs[:34]
        for cp in cover:                                     # centre every cover line, no stray indents
            if cp.text.strip():
                started = False
                for r in cp.runs:
                    if not started:
                        r.text = r.text.lstrip()
                        started = bool(r.text)
                if cp.runs: cp.runs[-1].text = cp.runs[-1].text.rstrip()
                cp.alignment = C
                cp.paragraph_format.left_indent = cp.paragraph_format.right_indent = Inches(0)
                cp.paragraph_format.first_line_indent = Inches(0)
        for i in (26, 23, 16, 9):                            # fewer blank lines so the cover fits one page
            el = cover[i]._p
            el.getparent().remove(el)
        self.cover, self.cover_end, self.title = cover, cover_end, title
        self.front_sections = 1
        self.chap = 0
        self.n_fig = self.n_tab = self.n_eq = 0
        self.toc, self.lof, self.lot = [], [], []
        self.eqs = []          # (paragraph, latex)
        self.placeholders = {}

    # ------------------------------------------------------------------ cover page and bonafide (copied from the SRM sample)
    def front_matter(self, info, sample):
        """Cover page and bonafide certificate copied element by element from the department's sample
        (front_pages_sample.docx): same paragraphs, styles, spacing, margins and logos. Only the
        placeholders are filled in; the sample's <Font ...> instruction lines are emptied but kept,
        so the vertical spacing stays exactly as in the sample."""
        import io
        I = info
        src = Document(sample)
        s_body = src.element.body
        els = [e for e in s_body if e.tag != qn("w:sectPr")]
        final_sect = s_body.find(qn("w:sectPr"))
        # drop the template's own cover (everything before the front-matter content)
        body = self.doc.element.body
        for node in list(body):
            if node.tag != qn("w:sectPr"):
                body.remove(node)
        # logos: copy the sample's images into this document and re-point the drawings
        rid_map = {}
        for rid, rel in src.part.rels.items():
            if "image" in rel.reltype:
                rid_map[rid], _ = self.doc.part.get_or_add_image(io.BytesIO(rel.target_part.blob))
        R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
        W = lambda t: qn("w:" + t)

        def texts(el):
            return [t for t in el.iter(W("t"))]

        def blank(el):                                       # keep the paragraph (spacing), drop its text
            for t in texts(el): t.text = ""

        def put(el, text, keep=1):                           # text into the first `keep` run(s), others emptied
            ts = texts(el)
            ts[0].text = text
            for t in ts[1:]: t.text = ""

        def full(el):
            return "".join(t.text or "" for t in texts(el))

        copies = [copy.deepcopy(e) for e in els]
        did = 9000
        for e in copies:                                     # unique drawing ids, local image rIds
            for d in e.iter("{http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing}docPr"):
                did += 1; d.set("id", str(did))
            for b in e.iter("{http://schemas.openxmlformats.org/drawingml/2006/main}blip"):
                old = b.get(f"{{{R_NS}}}embed")
                if old in rid_map: b.set(f"{{{R_NS}}}embed", rid_map[old])
            for sp in e.iter(W("sectPr")):                   # sample footer is empty: no footer on these pages
                for fr in sp.findall(W("footerReference")): sp.remove(fr)
        for e in copies:
            t = full(e).strip()
            if e.tag != W("p"):
                continue
            if t == "<Font Style Times New Roman - Bold>":     # project title line
                put(e, self.title)
                for r in e.iter(W("r")):
                    rpr = r.find(W("rPr"))
                    if rpr is None:
                        rpr = OxmlElement("w:rPr"); r.insert(0, rpr)
                    for tag in ("b", "bCs"):
                        if rpr.find(W(tag)) is None: rpr.append(OxmlElement("w:" + tag))
                    for tag in ("sz", "szCs"):
                        old = rpr.find(W(tag))
                        if old is not None: rpr.remove(old)
                        z = OxmlElement("w:" + tag); z.set(W("val"), "32"); rpr.append(z)
                    for sp_ in rpr.findall(W("spacing")): rpr.remove(sp_)
            elif t.startswith("<") and t.endswith(">") and "Font" in t:
                blank(e)                                     # instruction line: keep its height, drop the text
            elif t.startswith("NOVEMBER"):
                put(e, I["month"])
            elif t == "FACULTY NAME":
                put(e, I["guide"].upper())
            elif t == "Designation":
                put(e, I["guide_desig"])
        # student lines on the cover: STUDENT1 / STUDENT2 / STUDENT2 -> one line per team member
        studs = [e for e in copies if e.tag == W("p") and "STUDENT" in full(e) and "[REG NUM]" in full(e) and "Certified" not in full(e)]
        lines = [f"{n.upper()} [{r}]" for n, r in I["team"]]
        while len(studs) < len(lines):
            el = copy.deepcopy(studs[-1])
            copies.insert(copies.index(studs[-1]) + 1, el); studs.append(el)
        for el, line in zip(studs, lines):
            put(el, line)
        for el in studs[len(lines):]:
            copies.remove(el)
        # bonafide sentence: names run
        names = ", ".join(lines)
        for e in copies:
            if e.tag == W("p") and full(e).startswith("Certified that"):
                for x in texts(e):
                    if "STUDENT1" in (x.text or ""):
                        x.text = names + "” "
        # signature table placeholders
        for tbl in [e for e in copies if e.tag == W("tbl")]:
            for p in tbl.iter(W("p")):
                ts = texts(p)
                for x in ts:
                    if x.text in ("<", ">"): x.text = ""
                    elif x.text == "Faculty Name": x.text = I["guide"]
                    elif x.text == "Designation": x.text = I["guide_desig"]
                ppr = p.find(W("pPr"))                       # sample "Normal (Web)" style is justified; keep that explicit
                if ppr is not None and ppr.find(W("jc")) is None:
                    jc = OxmlElement("w:jc"); jc.set(W("val"), "both"); ppr.append(jc)
        # the bonafide's section break: the sample's last section properties on its last paragraph
        last_p = [e for e in copies if e.tag == W("p")][-1]
        ppr = last_p.find(W("pPr"))
        if ppr is None:
            ppr = OxmlElement("w:pPr"); last_p.insert(0, ppr)
        ppr.append(copy.deepcopy(final_sect))
        anchor = body.find(W("sectPr"))
        for e in copies:
            anchor.addprevious(e)
        self.front_sections = 2

    # ------------------------------------------------------------------ structure
    def prelim_heading(self, text, page_break=True, key=None):
        p = self.doc.add_paragraph()
        fmt(p, C, ls=1.0, before=0, after=18, page_break=page_break, keep_next=True)
        set_font(p.add_run(text), 16, bold=True)
        return p

    def placeholder(self, name):
        p = self.doc.add_paragraph(); fmt(p, C, ls=1.0, after=0)
        self.placeholders[name] = p
        return p

    def new_main_section(self):
        self.doc.add_section(WD_SECTION.NEW_PAGE)

    def chapter(self, title, first=False):
        self.chap += 1
        self.n_fig = self.n_tab = self.n_eq = 0
        p = self.doc.add_paragraph()
        fmt(p, C, ls=1.0, before=0, after=6, page_break=not first, keep_next=True)
        set_font(p.add_run(f"CHAPTER {self.chap}"), 16, bold=True, color=RGBColor(0, 0, 0))
        h = self.doc.add_paragraph(style="Heading 1")
        fmt(h, C, ls=1.15, before=6, after=18, keep_next=True, left=0, first=0, right=0)
        set_font(h.add_run(f"{self.chap}. {title.upper()}"), 16, bold=True)
        self.toc.append(("chapter", str(self.chap), title.upper(), f"CHAPTER {self.chap}"))

    def back_heading(self, text, toc_title=None):
        """REFERENCES / APPENDIX — same look as a chapter label, new page."""
        p = self.doc.add_paragraph(style="Heading 1")
        fmt(p, C, ls=1.0, before=0, after=18, page_break=True, keep_next=True, left=0, first=0, right=0)
        set_font(p.add_run(text), 16, bold=True, color=RGBColor(0, 0, 0))
        self.toc.append(("back", "", toc_title or text, text))
        return p

    def h2(self, text):
        p = self.doc.add_paragraph(style="Heading 2")
        fmt(p, L, ls=1.15, before=10, after=4, keep_next=True, left=0, first=0, right=0)
        set_font(p.add_run(text), 14, bold=True)

    def h3(self, text):
        p = self.doc.add_paragraph(style="Heading 3")
        fmt(p, L, ls=1.15, before=8, after=4, keep_next=True, left=0, first=0, right=0)
        set_font(p.add_run(text), 12, bold=True)

    # ------------------------------------------------------------------ text
    def p(self, text, align=J, after=6, keep_next=False, size=12):
        p = self.doc.add_paragraph()
        fmt(p, align, ls=1.5, after=after, keep_next=keep_next, left=0, first=0, right=0)
        add_rich(p, text, size=size)
        return p

    def bullets(self, items, size=12):
        for it in items:
            p = self.doc.add_paragraph(style="List Bullet")
            fmt(p, J, ls=1.5, after=3, hanging=0.25, right=0)
            add_rich(p, it, size=size)
        self.doc.paragraphs[-1].paragraph_format.space_after = Pt(6)

    def numbered(self, items, size=12):
        for i, it in enumerate(items, 1):
            p = self.doc.add_paragraph()
            fmt(p, J, ls=1.5, after=3, hanging=0.3, right=0)
            p.paragraph_format.tab_stops.add_tab_stop(Inches(0.3))
            add_rich(p, f"{i}.\t" + it, size=size)
        self.doc.paragraphs[-1].paragraph_format.space_after = Pt(6)

    def eq(self, latex):
        self.n_eq += 1
        p = self.doc.add_paragraph()
        fmt(p, L, ls=1.15, before=3, after=3, left=0, first=0, right=0)
        ts = p.paragraph_format.tab_stops
        ts.add_tab_stop(Inches(TEXT_W / 2), WD_TAB_ALIGNMENT.CENTER)
        ts.add_tab_stop(Inches(TEXT_W), WD_TAB_ALIGNMENT.RIGHT)
        set_font(p.add_run("\t"), 12)
        anchor = p.add_run(); set_font(anchor, 12)
        set_font(p.add_run(f"\t({self.chap}.{self.n_eq})"), 12)
        self.eqs.append((anchor, latex))
        return f"({self.chap}.{self.n_eq})"

    def where(self, text):
        return self.p(text, after=6)

    # ------------------------------------------------------------------ figures and tables
    def fig(self, path, caption, width=TEXT_W):
        self.n_fig += 1
        num = f"{self.chap}.{self.n_fig}"
        p = self.doc.add_paragraph(); fmt(p, C, ls=1.0, before=4, after=2, keep_next=True, left=0, first=0, right=0)
        p.add_run().add_picture(path, width=Inches(width))
        c = self.doc.add_paragraph(); fmt(c, C, ls=1.0, before=2, after=8, left=0, first=0, right=0)
        set_font(c.add_run(f"Figure {num}: {caption}"), 11, bold=True)
        self.lof.append((num, caption))
        return f"Figure {num}"

    def next_fig(self):
        return f"Figure {self.chap}.{self.n_fig + 1}"

    def next_tab(self):
        return f"Table {self.chap}.{self.n_tab + 1}"

    def table(self, caption, header, rows, widths, align=None, size=11, bold_first=False, numbered=True):
        size = 11                                            # guideline: table content 11 pt
        if numbered:
            self.n_tab += 1
            num = f"{self.chap}.{self.n_tab}"
            c = self.doc.add_paragraph(); fmt(c, C, ls=1.0, before=6, after=4, keep_next=True, left=0, first=0, right=0)
            set_font(c.add_run(f"Table {num}: {caption}"), 11, bold=True)
            self.lot.append((num, caption))
        t = self.doc.add_table(rows=1 + len(rows), cols=len(header))
        self._style_table(t, widths)
        align = align or ([L] + [C] * (len(header) - 1))
        for j, h in enumerate(header):
            cell = t.rows[0].cells[j]; set_cell_bg(cell, "D9E2F3")
            self._cell_text(cell, h, size, True, C)
        trPr = t.rows[0]._tr.get_or_add_trPr(); th = OxmlElement("w:tblHeader"); th.set(qn("w:val"), "true"); trPr.append(th)
        for i, row in enumerate(rows, 1):
            for j, v in enumerate(row):
                self._cell_text(t.rows[i].cells[j], str(v), size, bold_first and j == 0, align[j])
        for row in list(t.rows)[:-1]:                        # keep short tables on one page (last row
            if len(rows) <= 14:                              # not chained to the following text)
                for cell in row.cells:
                    for par in cell.paragraphs: par.paragraph_format.keep_with_next = True
        sp = self.doc.add_paragraph(); fmt(sp, L, ls=1.0, after=4)
        sp.paragraph_format.space_after = Pt(2)
        return f"Table {self.n_tab}" if not numbered else f"Table {num}"

    def _style_table(self, t, widths):
        t.alignment = WD_TABLE_ALIGNMENT.CENTER
        t.autofit = False
        set_borders(t); set_cell_margins(t)
        tblPr = t._tbl.tblPr
        for row in t.rows:
            trPr = row._tr.get_or_add_trPr()
            cs = OxmlElement("w:cantSplit"); cs.set(qn("w:val"), "true"); trPr.append(cs)
            for j, cell in enumerate(row.cells):
                cell.width = Inches(widths[j])
        tblPr = t._tbl.tblPr
        tw = tblPr.find(qn("w:tblW"))
        if tw is None:
            tw = OxmlElement("w:tblW"); tblPr.append(tw)
        tw.set(qn("w:w"), str(int(sum(widths) * 1440))); tw.set(qn("w:type"), "dxa")
        for row in t.rows:
            for cell in row.cells:
                va = OxmlElement("w:vAlign"); va.set(qn("w:val"), "center"); cell._tc.get_or_add_tcPr().append(va)
        grid = t._tbl.tblGrid
        for j, gc in enumerate(grid.findall(qn("w:gridCol"))):
            gc.set(qn("w:w"), str(int(widths[j] * 1440)))

    def _cell_text(self, cell, text, size, bold, align):
        cell.text = ""
        par = cell.paragraphs[0]
        lines = text.split("\n")
        for k, line in enumerate(lines):
            if k:
                par = cell.add_paragraph()
            fmt(par, align, ls=1.0, before=0, after=0)
            add_rich(par, line, size=size, bold=bold)

    def code(self, text, size=10, shade=True, after=8):
        lines = text.rstrip("\n").split("\n")
        p = self.doc.add_paragraph()
        fmt(p, L, ls=1.0, before=4, after=after, left=0, first=0, right=0)
        if shade:
            pPr = p._p.get_or_add_pPr()
            shd = OxmlElement("w:shd"); shd.set(qn("w:val"), "clear"); shd.set(qn("w:color"), "auto"); shd.set(qn("w:fill"), "F2F2F2")
            pPr.append(shd)
        for k, line in enumerate(lines):
            r = p.add_run(line.replace("\t", "    "))
            set_font(r, size, name="Consolas")
            if k < len(lines) - 1:
                r.add_break()
        return p

    def output(self, text):
        """Printed notebook output (Consolas 10, single spacing, no shading, left rule)."""
        p = self.code(text, size=10, shade=False)
        pPr = p._p.get_or_add_pPr()
        bdr = OxmlElement("w:pBdr"); left = OxmlElement("w:left")
        for k, v in (("w:val", "single"), ("w:sz", "12"), ("w:space", "6"), ("w:color", "2A78D6")):
            left.set(qn(k), v)
        bdr.append(left); pPr.insert(0, bdr)
        return p

    # ------------------------------------------------------------------ finishing
    def finish_equations(self, workdir):
        maths = latex_to_omml([l for _, l in self.eqs], workdir)
        M = "{http://schemas.openxmlformats.org/officeDocument/2006/math}"
        for (anchor, _), m in zip(self.eqs, maths):
            m = copy.deepcopy(m)
            for r in m.iter(M + "r"):                         # 12 pt Cambria Math for every math run
                rpr = OxmlElement("w:rPr")
                rf = OxmlElement("w:rFonts"); rf.set(qn("w:ascii"), "Cambria Math"); rf.set(qn("w:hAnsi"), "Cambria Math")
                sz = OxmlElement("w:sz"); sz.set(qn("w:val"), "24")
                rpr.append(rf); rpr.append(sz)
                mrpr = r.find(M + "rPr")
                if mrpr is not None: mrpr.addnext(rpr)
                else: r.insert(0, rpr)
            anchor._r.addnext(m)

    def sections_and_footers(self):
        secs = self.doc.sections
        for s in list(secs)[self.front_sections:]:          # same margins on every page (front pages keep the sample's)
            s.left_margin, s.right_margin = Inches(1.5), Inches(1.0)
            s.top_margin, s.bottom_margin = Inches(1.0), Inches(1.0)
            s.footer_distance = Inches(0.5)
        for s, fmt_, start in ((secs[-2], "lowerRoman", ROMAN_START), (secs[-1], "decimal", 1)):
            sp = s._sectPr
            for fr in sp.findall(qn("w:footerReference")):
                sp.remove(fr)
            s.start_type = WD_SECTION.NEW_PAGE
            pg = OxmlElement("w:pgNumType"); pg.set(qn("w:fmt"), fmt_); pg.set(qn("w:start"), str(start))
            sp.find(qn("w:cols")).addprevious(pg)
            s.footer.is_linked_to_previous = False
            fp = s.footer.paragraphs[0]
            fp.text = ""
            fmt(fp, C, ls=1.0, after=0)
            page_field(fp, 12)
        secs[0].start_type = WD_SECTION.NEW_PAGE

    def build_index_tables(self, pages=None):
        """TOC, list of figures, list of tables at their placeholders. pages: key -> page label."""
        pages = pages or {}
        get = lambda key: pages.get(key, "")
        # --- table of contents
        rows = [("ABSTRACT", "abstract"), ("TABLE OF CONTENTS", "toc"), ("LIST OF FIGURES", "lof"),
                ("LIST OF TABLES", "lot"), ("ABBREVIATIONS", "abbr")]
        body = []
        for kind, no, title, key in self.toc:
            body.append((kind, no, title, key))
        n = len(rows) + 1 + len(body)
        t = self.doc.add_table(rows=n, cols=3)
        W = [1.0, 3.96, 0.8]
        self._style_table(t, W)
        for i, (title, key) in enumerate(rows):
            a, b, c = t.rows[i].cells
            a.merge(b)
            self._cell_text(a, title, 11, True, L); self._cell_text(c, get(key), 11, False, C)
        hdr = t.rows[len(rows)].cells
        for j, h in enumerate(("CHAPTER NO.", "TITLE", "PAGE NO.")):
            set_cell_bg(hdr[j], "D9E2F3"); self._cell_text(hdr[j], h, 11, True, C)
        for i, (kind, no, title, key) in enumerate(body, len(rows) + 1):
            a, b, c = t.rows[i].cells
            bold = kind in ("chapter", "back")
            self._cell_text(a, no, 11, bold, C)
            self._cell_text(b, title if kind != "section" else "    " + title, 11, bold, L)
            self._cell_text(c, get(key), 11, False, C)
        self.placeholders["toc"]._p.addnext(t._tbl)
        # --- figures / tables
        for name, items, label in (("lof", self.lof, "FIGURE NO."), ("lot", self.lot, "TABLE NO.")):
            t = self.doc.add_table(rows=1 + len(items), cols=3)
            W = [1.0, 3.96, 0.8]
            self._style_table(t, W)
            for j, h in enumerate((label, "TITLE", "PAGE NO.")):
                set_cell_bg(t.rows[0].cells[j], "D9E2F3"); self._cell_text(t.rows[0].cells[j], h, 11, True, C)
            kind = "Figure" if name == "lof" else "Table"
            for i, (num, cap) in enumerate(items, 1):
                cells = t.rows[i].cells
                self._cell_text(cells[0], num, 11, False, C)
                self._cell_text(cells[1], cap, 11, False, L)
                self._cell_text(cells[2], get(f"{kind} {num}:"), 11, False, C)
            self.placeholders[name]._p.addnext(t._tbl)


# ---------------------------------------------------------------------- schema-order normaliser
_ORDER = {
    "pPr": ["pStyle", "keepNext", "keepLines", "pageBreakBefore", "framePr", "widowControl", "numPr",
            "suppressLineNumbers", "pBdr", "shd", "tabs", "suppressAutoHyphens", "kinsoku", "wordWrap",
            "overflowPunct", "topLinePunct", "autoSpaceDE", "autoSpaceDN", "bidi", "adjustRightInd", "snapToGrid",
            "spacing", "ind", "contextualSpacing", "mirrorIndents", "suppressOverlap", "jc", "textDirection",
            "textAlignment", "textboxTightWrap", "outlineLvl", "divId", "cnfStyle", "rPr", "sectPr", "pPrChange"],
    "tblPr": ["tblStyle", "tblpPr", "tblOverlap", "bidiVisual", "tblStyleRowBandSize", "tblStyleColBandSize", "tblW",
              "jc", "tblCellSpacing", "tblInd", "tblBorders", "shd", "tblLayout", "tblCellMar", "tblLook",
              "tblCaption", "tblDescription"],
    "tcPr": ["cnfStyle", "tcW", "gridSpan", "hMerge", "vMerge", "tcBorders", "shd", "noWrap", "tcMar",
             "textDirection", "tcFitText", "vAlign", "hideMark"],
}
_MORDER = {"dPr": ["begChr", "sepChr", "endChr", "grow", "shp", "ctrlPr"]}


def normalise(doc):
    W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
    M = "{http://schemas.openxmlformats.org/officeDocument/2006/math}"
    parts = [doc.element] + [s.footer._element for s in doc.sections if not s.footer.is_linked_to_previous]
    for root in parts:
        for tag, order in _ORDER.items():
            for el in root.iter(W + tag):
                kids = list(el)
                rank = {n: i for i, n in enumerate(order)}
                kids.sort(key=lambda k: rank.get(k.tag.split("}")[1], 999))
                for k in kids:
                    el.remove(k)
                for k in kids:
                    el.append(k)
        for el in root.iter(M + "dPr"):
            kids = list(el); rank = {n: i for i, n in enumerate(_MORDER["dPr"])}
            kids.sort(key=lambda k: rank.get(k.tag.split("}")[1], 999))
            for k in kids: el.remove(k)
            for k in kids: el.append(k)
        for el in root.iter(M + "rPr"):
            kids = list(el); rank = {n: i for i, n in enumerate(["lit", "nor", "scr", "sty", "brk", "aln"])}
            kids.sort(key=lambda k: rank.get(k.tag.split("}")[1], 999))
            for k in kids: el.remove(k)
            for k in kids: el.append(k)
            if el.find(M + "nor") is not None:
                for t in ("sty", "scr"):
                    if el.find(M + t) is not None:
                        el.remove(el.find(M + t))
