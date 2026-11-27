# -*- coding: utf-8 -*-
"""Làm việc trực tiếp với Word/Excel đang mở (COM). Không phụ thuộc Tkinter.
Cần Windows + Microsoft Office desktop + pywin32. Mọi lệnh COM chạy trên thread giao diện.
"""
import os

RED = 255                    # RGB(255,0,0)
MIXED = 9999999              # Word: nhiều màu trong cùng vùng chọn


class LiveIssue:
    def __init__(self, source, doc, doc_name, doc_dir, location, original, suggestions,
                 kind, start, end, sheet=None, cell=None):
        self.source = source              # "word" | "excel"
        self.doc = doc                    # đối tượng COM Document / Workbook
        self.doc_name = doc_name
        self.doc_dir = doc_dir
        self.file = doc_name
        self.location = location
        self.original = original
        self.suggestions = suggestions
        self.kind = kind
        self.start, self.end = start, end
        self.sheet, self.cell = sheet, cell
        self.selected = suggestions[0] if suggestions else ""


class OfficeLive:
    def __init__(self):
        self.marks = []
        self._com_ready = False

    # ── kết nối ──────────────────────────────────────────────
    def _apps(self):
        try:
            import pythoncom
            import win32com.client as wc
        except ImportError as e:
            raise RuntimeError("Tính năng quét Office đang mở cần Windows, "
                               "Microsoft Office và thư viện pywin32.") from e
        if not self._com_ready:
            pythoncom.CoInitialize()
            self._com_ready = True

        def get(prog):
            try:
                return wc.GetActiveObject(prog)
            except Exception:
                return None
        return get("Word.Application"), get("Excel.Application")

    def _docs(self):
        word, excel = self._apps()
        out = []
        if word is not None:
            try:
                for i in range(1, int(word.Documents.Count) + 1):
                    d = word.Documents.Item(i)
                    out.append(("word", d, str(d.Name)))
            except Exception:
                pass
        if excel is not None:
            try:
                for i in range(1, int(excel.Workbooks.Count) + 1):
                    wb = excel.Workbooks.Item(i)
                    try:
                        if int(wb.Windows.Count) == 0:      # workbook ẩn (PERSONAL.XLSB...)
                            continue
                    except Exception:
                        pass
                    out.append(("excel", wb, str(wb.Name)))
            except Exception:
                pass
        return out

    @staticmethod
    def _info(docs, errors=()):
        return {"word": [n for k, _, n in docs if k == "word"],
                "excel": [n for k, _, n in docs if k == "excel"],
                "total": len(docs), "errors": list(errors)}

    def list_open(self):
        return self._info(self._docs())

    def fingerprint(self):
        fp = {}
        for kind, obj, name in self._docs():
            try:
                if kind == "word":
                    fp[(kind, name)] = hash(str(obj.Content.Text))
                else:
                    parts = []
                    for i in range(1, int(obj.Worksheets.Count) + 1):
                        ws = obj.Worksheets.Item(i)
                        parts.append(str(ws.Name) + str(ws.UsedRange.Value2))
                    fp[(kind, name)] = hash("|".join(parts))
            except Exception:
                pass
        return fp

    # ── đánh dấu đỏ ──────────────────────────────────────────
    def _mark(self, obj):
        try:
            old = obj.Font.Color
            if old is None or old == MIXED:
                return
            self.marks.append((obj, old))
            obj.Font.Color = RED
        except Exception:
            pass

    def restore_marks(self):
        for obj, old in reversed(self.marks):
            try:
                obj.Font.Color = old
            except Exception:
                pass
        self.marks = []

    # ── quét ─────────────────────────────────────────────────
    def scan(self, check):
        """check(text) -> [(start,end,word,sugs,kind)]. Trả về (issues, info)."""
        self.restore_marks()
        docs, issues, errors = self._docs(), [], []
        for kind, obj, name in docs:
            try:
                if kind == "word":
                    issues += self._scan_word(obj, name, check)
                else:
                    issues += self._scan_excel(obj, name, check)
            except Exception as e:
                errors.append(f"{name}: {e}")
        return issues, self._info(docs, errors)

    @staticmethod
    def _dir(obj):
        try:
            return str(obj.Path)
        except Exception:
            return ""

    def _scan_word(self, doc, name, check):
        out, ddir = [], self._dir(doc)
        paras = doc.Paragraphs
        for pi in range(1, int(paras.Count) + 1):
            rng = paras.Item(pi).Range
            txt = str(rng.Text or "").rstrip("\r\x07")
            if not txt:
                continue
            base = int(rng.Start)
            for s, e, w, sugs, kind in check(txt):
                self._mark(doc.Range(base + s, base + e))
                out.append(LiveIssue("word", doc, name, ddir, f"Đoạn {pi}, ký tự {s + 1}",
                                     w, sugs, kind, base + s, base + e))
        return out

    def _scan_excel(self, wb, name, check):
        out, ddir = [], self._dir(wb)
        for si in range(1, int(wb.Worksheets.Count) + 1):
            ws = wb.Worksheets.Item(si)
            used = ws.UsedRange
            vals, forms = used.Value2, used.Formula     # đọc cả vùng một lần: nhanh hơn từng ô
            if vals is None:
                continue
            if not isinstance(vals, tuple):
                vals, forms = ((vals,),), ((forms,),)
            r0, c0 = int(used.Row), int(used.Column)
            for ri, row in enumerate(vals):
                for ci, val in enumerate(row):
                    if not isinstance(val, str) or not val:
                        continue
                    f = forms[ri][ci]
                    if isinstance(f, str) and f.startswith("="):
                        continue
                    hits = check(val)
                    if not hits:
                        continue
                    cell = ws.Cells(r0 + ri, c0 + ci)
                    addr = str(cell.Address(False, False))
                    for s, e, w, sugs, kind in hits:
                        try:
                            self._mark(cell.Characters(s + 1, e - s))
                        except Exception:
                            pass
                        out.append(LiveIssue("excel", wb, name, ddir, f"{ws.Name}!{addr}",
                                             w, sugs, kind, s, e, sheet=str(ws.Name), cell=addr))
        return out

    # ── sửa: luôn lưu thành bản sao, không ghi đè tệp gốc ────
    def fix_copy(self, source, doc, issues, target):
        target = os.path.normpath(target)
        try:
            if os.path.normcase(target) == os.path.normcase(os.path.normpath(str(doc.FullName))):
                raise RuntimeError("Không được ghi đè tệp gốc. Hãy chọn tên khác.")
        except AttributeError:
            pass
        if source == "word":
            doc.SaveAs2(FileName=target, FileFormat=16)       # 16 = .docx
            for x in sorted(issues, key=lambda z: z.start, reverse=True):
                if x.selected:
                    doc.Range(int(x.start), int(x.end)).Text = x.selected
            doc.Save()
        else:
            app = doc.Application
            doc.SaveCopyAs(target)
            copy_wb = app.Workbooks.Open(target)
            grouped = {}
            for x in issues:
                if x.selected:
                    grouped.setdefault((x.sheet, x.cell), []).append(x)
            for (sheet, addr), items in grouped.items():
                cell = copy_wb.Worksheets(sheet).Range(addr)
                value = str(cell.Value2 or "")
                for x in sorted(items, key=lambda z: z.start, reverse=True):
                    value = value[:x.start] + x.selected + value[x.end:]
                cell.Value2 = value
            copy_wb.Save()
