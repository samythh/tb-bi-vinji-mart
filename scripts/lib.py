import copy, re
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.text.paragraph import Paragraph
from docx.shared import Cm

TOTAL_W = 8306


def rp(x):
    neg = x < 0
    s = f'{abs(round(x)):,.0f}'.replace(',', '.')
    return ('−Rp ' if neg else 'Rp ') + s


def num(x, d=0):
    s = f'{x:,.{d}f}'
    return s.replace(',', 'X').replace('.', ',').replace('X', '.')


def pct(x, d=1):
    return num(x, d) + '%'


class Builder:
    def __init__(self, doc):
        self.doc = doc
        body = doc.element.body
        kids = list(body.iterchildren())
        # keep cover + TOC + first empty Heading2 (index 0..15)
        sect = body.find(qn('w:sectPr'))
        for el in kids[16:]:
            if el is not sect:
                body.remove(el)
        self.body = body
        self.sect = sect
        self.num_base = self._find_num_base()

    # ---------- low level ----------
    def _p(self, style, jc=None, first_line=None, keep_next=False, spacing_before0=False):
        p = OxmlElement('w:p')
        ppr = OxmlElement('w:pPr')
        ps = OxmlElement('w:pStyle'); ps.set(qn('w:val'), style); ppr.append(ps)
        if keep_next:
            ppr.append(OxmlElement('w:keepNext'))
        if spacing_before0:
            sp = OxmlElement('w:spacing'); sp.set(qn('w:before'), '0'); ppr.append(sp)
        if first_line is not None:
            ind = OxmlElement('w:ind'); ind.set(qn('w:firstLine'), str(first_line)); ppr.append(ind)
        if jc:
            j = OxmlElement('w:jc'); j.set(qn('w:val'), jc); ppr.append(j)
        p.append(ppr)
        return p

    def _runs(self, p, text, bold=False):
        # supports **bold** and [placeholder] (yellow)
        tokens = re.split(r'(\*\*.+?\*\*|\[[^\]]+\])', text)
        for tk in tokens:
            if not tk:
                continue
            b = bold; hl = False
            if tk.startswith('**') and tk.endswith('**'):
                tk = tk[2:-2]; b = True
            elif tk.startswith('[') and tk.endswith(']'):
                hl = True
            r = OxmlElement('w:r')
            rpr = OxmlElement('w:rPr')
            if b:
                rpr.append(OxmlElement('w:b'))
            if hl:
                h = OxmlElement('w:highlight'); h.set(qn('w:val'), 'yellow'); rpr.append(h)
            if len(rpr):
                r.append(rpr)
            parts = tk.split('\t')
            for i, part in enumerate(parts):
                if i:
                    r.append(OxmlElement('w:tab'))
                t = OxmlElement('w:t'); t.set(qn('xml:space'), 'preserve'); t.text = part
                r.append(t)
            p.append(r)
        return p

    def _add(self, el):
        self.sect.addprevious(el)
        return el

    # ---------- blocks ----------
    def h(self, level, text):
        p = self._p(f'Heading{level}', spacing_before0=(level == 4))
        self._runs(p, text)
        return self._add(p)

    def p(self, text, indent=True):
        p = self._p('IsiLaporan', first_line=567 if indent else None)
        self._runs(p, text)
        return self._add(p)

    def _find_num_base(self):
        numbering = self.doc.part.numbering_part.element
        for n in numbering.findall(qn('w:num')):
            if n.get(qn('w:numId')) == '20':
                return n.find(qn('w:abstractNumId')).get(qn('w:val'))
        return None

    def _new_num(self):
        numbering = self.doc.part.numbering_part.element
        ids = [int(n.get(qn('w:numId'))) for n in numbering.findall(qn('w:num'))]
        nid = max(ids) + 1
        n = OxmlElement('w:num'); n.set(qn('w:numId'), str(nid))
        a = OxmlElement('w:abstractNumId'); a.set(qn('w:val'), self.num_base); n.append(a)
        lo = OxmlElement('w:lvlOverride'); lo.set(qn('w:ilvl'), '0')
        so = OxmlElement('w:startOverride'); so.set(qn('w:val'), '1'); lo.append(so); n.append(lo)
        numbering.append(n)
        return nid

    def ol(self, items):
        nid = self._new_num()
        for it in items:
            p = self._p('IsiLaporan')
            ppr = p.find(qn('w:pPr'))
            npr = OxmlElement('w:numPr')
            il = OxmlElement('w:ilvl'); il.set(qn('w:val'), '0'); npr.append(il)
            ni = OxmlElement('w:numId'); ni.set(qn('w:val'), str(nid)); npr.append(ni)
            ppr.append(npr)
            self._runs(p, it)
            self._add(p)

    def cap(self, text, table=True):
        p = self._p('Caption', keep_next=table)
        self._runs(p, text)
        return self._add(p)

    def fig(self, path, caption, width_cm=15.5):
        p = self._p('IsiLaporan', jc='center', keep_next=True)
        self._add(p)
        Paragraph(p, self.doc._body).add_run().add_picture(path, width=Cm(width_cm))
        self.cap(caption, table=False)

    def table(self, headers, rows, widths, sz=None):
        widths = list(widths)
        s = sum(widths)
        widths = [round(w * TOTAL_W / s) for w in widths]
        widths[-1] += TOTAL_W - sum(widths)
        tbl = OxmlElement('w:tbl')
        tpr = OxmlElement('w:tblPr')
        st = OxmlElement('w:tblStyle'); st.set(qn('w:val'), 'TableGrid'); tpr.append(st)
        tw = OxmlElement('w:tblW'); tw.set(qn('w:w'), str(TOTAL_W)); tw.set(qn('w:type'), 'dxa'); tpr.append(tw)
        lay = OxmlElement('w:tblLayout'); lay.set(qn('w:type'), 'fixed'); tpr.append(lay)
        look = OxmlElement('w:tblLook'); look.set(qn('w:val'), '0420'); look.set(qn('w:firstRow'), '1'); look.set(qn('w:noVBand'), '1'); tpr.append(look)
        tbl.append(tpr)
        grid = OxmlElement('w:tblGrid')
        for w in widths:
            g = OxmlElement('w:gridCol'); g.set(qn('w:w'), str(w)); grid.append(g)
        tbl.append(grid)
        for ri, row in enumerate([headers] + rows):
            tr = OxmlElement('w:tr')
            if ri == 0:
                trpr = OxmlElement('w:trPr'); trpr.append(OxmlElement('w:cantSplit')); trpr.append(OxmlElement('w:tblHeader')); tr.append(trpr)
            for ci, (val, w) in enumerate(zip(row, widths)):
                tc = OxmlElement('w:tc')
                tcpr = OxmlElement('w:tcPr')
                cw = OxmlElement('w:tcW'); cw.set(qn('w:w'), str(w)); cw.set(qn('w:type'), 'dxa'); tcpr.append(cw)
                if ri == 0:
                    sh = OxmlElement('w:shd'); sh.set(qn('w:val'), 'clear'); sh.set(qn('w:color'), 'auto'); sh.set(qn('w:fill'), 'F2F2F2'); tcpr.append(sh)
                tc.append(tcpr)
                lines = str(val).split('\n')
                for ln in lines:
                    p = self._p('IsiTabel', keep_next=(ri == 0))
                    self._runs(p, ln, bold=(ri == 0))
                    tc.append(p)
                tr.append(tc)
            tbl.append(tr)
        if sz:
            mar = OxmlElement('w:tblCellMar')
            for side in ('left', 'right'):
                e = OxmlElement('w:' + side); e.set(qn('w:w'), '45'); e.set(qn('w:type'), 'dxa'); mar.append(e)
            tpr.insert(list(tpr).index(look), mar)
            for r in tbl.iter(qn('w:r')):
                rpr = r.find(qn('w:rPr'))
                if rpr is None:
                    rpr = OxmlElement('w:rPr'); r.insert(0, rpr)
                for tag in ('w:sz', 'w:szCs'):
                    e = OxmlElement(tag); e.set(qn('w:val'), str(sz)); rpr.append(e)
        self._add(tbl)
        # empty paragraph after table for spacing
        self._add(self._p('IsiLaporan'))
        return tbl


def finalize(doc):
    body = doc.element.body
    sect = body.find(qn('w:sectPr'))
    last = sect.getprevious()
    if last is not None and last.tag == qn('w:p') and not ''.join(x.text or '' for x in last.iter(qn('w:t'))):
        ppr = last.find(qn('w:pPr'))
        sp = OxmlElement('w:spacing'); sp.set(qn('w:before'), '0'); sp.set(qn('w:after'), '0'); sp.set(qn('w:line'), '20'); sp.set(qn('w:lineRule'), 'exact')
        ppr.append(sp)
        rpr = OxmlElement('w:rPr'); sz = OxmlElement('w:sz'); sz.set(qn('w:val'), '2'); rpr.append(sz); ppr.append(rpr)


def set_cover(doc, judul_dok, judul_proyek, nama_dok):
    ps = list(doc.element.body.iterchildren())
    def setp(el, text):
        p = Paragraph(el, doc._body)
        runs = p.runs
        runs[0].text = text
        for r in runs[1:]:
            r._r.getparent().remove(r._r)
        return p
    for el in ps[:9]:
        t = ''.join(x.text or '' for x in el.iter(qn('w:t')))
        if t.startswith('Judul Dokumen'):
            setp(el, 'Judul Dokumen\t: ' + judul_dok)
        elif t.startswith('Judul Proyek'):
            setp(el, 'Judul Proyek\t: ' + judul_proyek)
        elif t.startswith('Nama dokumen'):
            setp(el, 'Nama dokumen\t: ' + nama_dok)
        elif t.startswith('Tanggal diserahkan'):
            setp(el, 'Tanggal diserahkan\t: 19 November 2026')
        elif t.startswith('Jumlah Halaman'):
            setp(el, 'Jumlah Halaman\t: {PAGES}')
