# -*- coding: utf-8 -*-
"""Quét tệp Word (.docx) / Excel (.xlsx, .xlsm) lưu trên đĩa."""
from pathlib import Path
from spell_engine import WORD_RE, apply_case

try:
    from docx import Document
except ImportError:
    Document = None
try:
    import openpyxl
except ImportError:
    openpyxl = None


def check_text(checker, text):
    """Trả về [(start, end, word, suggestions, kind)] cho một đoạn văn bản."""
    toks = list(WORD_RE.finditer(text))
    words = [m.group() for m in toks]
    out = []
    for i, m in enumerate(toks):
        prev = words[i - 1].lower() if i and text[toks[i - 1].end():m.start()].isspace() else None
        nxt = (words[i + 1].lower()
               if i + 1 < len(toks) and text[m.end():toks[i + 1].start()].isspace() else None)
        kind, sugs = checker.check_word(m.group(), prev, nxt)
        if kind:
            out.append((m.start(), m.end(), m.group(),
                        [apply_case(m.group(), x) for x in sugs], kind))
    return out


class FileIssue:
    """Một lỗi trong tệp trên đĩa. start/end được cập nhật sau mỗi lần sửa."""
    source = "file"

    def __init__(self, file, location, container, group, word, sugs, kind, start, end):
        self.file = file
        self.doc_name = Path(file).name
        self.location = location
        self.container = container          # run (docx) hoặc cell (xlsx)
        self.group = group                  # các lỗi cùng container
        self.original = word
        self.suggestions = sugs
        self.kind = kind
        self.start, self.end = start, end
        self.selected = sugs[0] if sugs else ""
        self.save = None                    # gán sau: hàm lưu bản _checked

    def _get(self):
        c = self.container
        return c.text if hasattr(c, "text") else c.value

    def _set(self, v):
        c = self.container
        if hasattr(c, "text"):
            c.text = v
        else:
            c.value = v

    def apply(self, new):
        text = self._get()
        old_end = self.end
        self._set(text[:self.start] + new + text[self.end:])
        delta = len(new) - (self.end - self.start)
        self.end = self.start + len(new)
        for y in self.group:                # dịch vị trí các lỗi còn lại
            if y is not self and y.start >= old_end:
                y.start += delta
                y.end += delta


def _target(fn):
    p = Path(fn)
    return p.with_name(p.stem + "_checked" + p.suffix)


def scan_docx(checker, fn):
    if Document is None:
        raise RuntimeError("Thiếu thư viện python-docx.")
    doc = Document(fn)
    out = []

    def scan_par(p, loc):
        for ri, run in enumerate(p.runs):          # quét từng run để giữ định dạng
            group = []
            for s, e, w, sugs, kind in check_text(checker, run.text):
                x = FileIssue(fn, f"{loc} / Run {ri + 1}: ký tự {s + 1}", run, group,
                              w, sugs, kind, s, e)
                group.append(x)
                out.append(x)

    for pi, p in enumerate(doc.paragraphs, 1):
        scan_par(p, f"Đoạn {pi}")
    for ti, t in enumerate(doc.tables, 1):
        for ri, row in enumerate(t.rows, 1):
            for ci, cell in enumerate(row.cells, 1):
                for pi, p in enumerate(cell.paragraphs, 1):
                    scan_par(p, f"Bảng {ti}, hàng {ri}, cột {ci}, đoạn {pi}")
    for si, sec in enumerate(doc.sections, 1):
        for name, paras in (("Đầu trang", sec.header.paragraphs),
                            ("Chân trang", sec.footer.paragraphs)):
            for pi, p in enumerate(paras, 1):
                scan_par(p, f"Trang {si} - {name}, đoạn {pi}")

    def save():
        t = _target(fn)
        doc.save(t)
        return str(t)

    for x in out:
        x.save = save
    return out


def scan_excel(checker, fn):
    if openpyxl is None:
        raise RuntimeError("Thiếu thư viện openpyxl.")
    wb = openpyxl.load_workbook(fn, keep_vba=fn.lower().endswith(".xlsm"))
    out = []
    for ws in wb.worksheets:
        for row in ws.iter_rows():
            for cell in row:
                v = cell.value
                if isinstance(v, str) and v and not v.startswith("="):
                    group = []
                    for s, e, w, sugs, kind in check_text(checker, v):
                        x = FileIssue(fn, f"{ws.title}!{cell.coordinate}", cell, group,
                                      w, sugs, kind, s, e)
                        group.append(x)
                        out.append(x)

    def save():
        t = _target(fn)
        wb.save(t)
        return str(t)

    for x in out:
        x.save = save
    return out
