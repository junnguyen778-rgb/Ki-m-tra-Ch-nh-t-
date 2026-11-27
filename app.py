# -*- coding: utf-8 -*-
"""ViSpellCheck Pro: quét Word/Excel ĐANG MỞ chỉ với một cú bấm (không cần thêm tệp).
Tùy chọn: kiểm tra thêm tệp .docx/.xlsx/.xlsm trên đĩa."""
import csv
import tkinter as tk
from pathlib import Path
from tkinter import ttk, filedialog, messagebox

from spell_engine import Checker
from doc_scan import check_text, scan_docx, scan_excel
from office_live import OfficeLive

# ───────────────────────── Bảng màu ─────────────────────────
PAL = dict(bg="#e8eef8", card="#ffffff", shadow="#b4c0d6", text="#1e293b",
           muted="#64748b", footer="#0f172a")
BLUE = ("#60a5fa", "#1d4ed8")
GREEN = ("#4ade80", "#15803d")
ORANGE = ("#fbbf24", "#d97706")
DEEP_ORANGE = ("#fb923c", "#c2410c")
PURPLE = ("#c4b5fd", "#6d28d9")
TEAL = ("#5eead4", "#0f766e")
SLATE = ("#cbd5e1", "#475569")
RED = ("#fca5a5", "#b91c1c")
FONT = "Segoe UI"


def _rgb(h):
    h = h.lstrip("#")
    return [int(h[i:i + 2], 16) for i in (0, 2, 4)]


def mix(c1, c2, t):
    a, b = _rgb(c1), _rgb(c2)
    return "#%02x%02x%02x" % tuple(max(0, min(255, int(a[i] + (b[i] - a[i]) * t))) for i in range(3))


def shade(c, f):
    """f<0: tối hơn, f>0: sáng hơn."""
    return mix(c, "#000000", -f) if f < 0 else mix(c, "#ffffff", f)


def round_pts(x0, y0, x1, y1, r):
    r = max(1, min(r, (x1 - x0) / 2, (y1 - y0) / 2))
    return [x0 + r, y0, x1 - r, y0, x1, y0, x1, y0 + r, x1, y1 - r, x1, y1,
            x1 - r, y1, x0 + r, y1, x0, y1, x0, y1 - r, x0, y0 + r, x0, y0]


def draw_gradient(cv, x0, y0, x1, y1, c1, c2, vertical=True, step=2):
    if vertical:
        n = max(1, int(y1 - y0))
        for i in range(0, n, 1):
            cv.create_line(x0, y0 + i, x1, y0 + i, fill=mix(c1, c2, i / n))
    else:
        n = max(1, int(x1 - x0))
        for i in range(0, n, step):
            cv.create_line(x0 + i, y0, x0 + i, y1, fill=mix(c1, c2, i / n), width=step)


# ───────────────────────── Nút 3D ─────────────────────────
class Button3D(tk.Canvas):
    """Nút có độ nổi: viền đáy đậm, mặt bóng, lún xuống khi bấm."""
    DEPTH = 4

    def __init__(self, master, text, command=None, colors=BLUE, height=40, width=160,
                 font=(FONT, 10, "bold"), bg=PAL["card"]):
        super().__init__(master, width=width, height=height + self.DEPTH + 2, bg=bg,
                         highlightthickness=0, bd=0, cursor="hand2")
        self.b_text, self.b_cmd, self.b_colors = text, command, colors
        self.b_w, self.b_h, self.b_font = width, height, font
        self.b_hover = self.b_pressed = False
        self.b_enabled = True
        self.bind("<Configure>", self._on_resize)
        self.bind("<Enter>", lambda e: self._set(hover=True))
        self.bind("<Leave>", lambda e: self._set(hover=False, pressed=False))
        self.bind("<ButtonPress-1>", lambda e: self._set(pressed=True))
        self.bind("<ButtonRelease-1>", self._release)
        self._draw()

    def _on_resize(self, e):
        self.b_w = e.width
        self._draw()

    def _set(self, **kw):
        if "hover" in kw:
            self.b_hover = kw["hover"]
        if "pressed" in kw:
            self.b_pressed = kw["pressed"]
        self._draw()

    def _release(self, e):
        was = self.b_pressed
        self.b_pressed = False
        self._draw()
        inside = 0 <= e.x <= self.b_w and 0 <= e.y <= self.b_h + self.DEPTH
        if was and inside and self.b_enabled and self.b_cmd:
            self.b_cmd()

    def update_button(self, text=None, colors=None, enabled=None):
        if text is not None:
            self.b_text = text
        if colors is not None:
            self.b_colors = colors
        if enabled is not None:
            self.b_enabled = enabled
        self._draw()

    def _draw(self):
        self.delete("all")
        w, h, d = self.b_w, self.b_h, self.DEPTH
        top, bot = self.b_colors
        if not self.b_enabled:
            top, bot = "#e2e8f0", "#94a3b8"
        elif self.b_hover:
            top, bot = shade(top, .15), shade(bot, .10)
        dy = d - 1 if self.b_pressed else 0
        self.create_polygon(round_pts(1, d, w - 1, self.b_h + d, 11), smooth=True,
                            fill=shade(bot, -.38), outline="")
        self.create_polygon(round_pts(1, dy, w - 1, h + dy, 11), smooth=True,
                            fill=bot, outline=shade(bot, -.15))
        self.create_polygon(round_pts(5, dy + 2, w - 5, dy + h * .5, 9), smooth=True,
                            fill=mix(bot, top, .55), outline="")
        self.create_text(w / 2, dy + h / 2, text=self.b_text, font=self.b_font,
                         fill="#ffffff" if self.b_enabled else "#f1f5f9")


# ───────────────────────── Khung chức năng nổi ─────────────────────────
def make_card(parent, title, colors):
    """Khung có bóng đổ (phải/dưới), thanh tiêu đề chuyển màu. Trả về (khung_ngoài, thân)."""
    outer = tk.Frame(parent, bg=PAL["shadow"])
    card = tk.Frame(outer, bg=PAL["card"], highlightthickness=1,
                    highlightbackground=shade(colors[1], .45))
    card.pack(fill="both", expand=True, padx=(0, 4), pady=(0, 4))
    head = tk.Canvas(card, height=34, highlightthickness=0, bd=0, bg=colors[1])
    head.pack(fill="x")

    def paint(e):
        head.delete("all")
        draw_gradient(head, 0, 0, e.width, 34, colors[0], colors[1], vertical=True)
        head.create_line(0, 0, e.width, 0, fill=shade(colors[0], .55))
        head.create_text(13, 18, anchor="w", text=title, fill="#ffffff",
                         font=(FONT, 10, "bold"))
    head.bind("<Configure>", paint)
    body = tk.Frame(card, bg=PAL["card"], padx=12, pady=10)
    body.pack(fill="both", expand=True)
    return outer, body


def desc_label(body, text, size=9):
    lb = tk.Label(body, text=text, bg=PAL["card"], fg=PAL["muted"], justify="left",
                  anchor="w", font=(FONT, size))
    lb.pack(fill="x")
    body.bind("<Configure>", lambda e: lb.configure(wraplength=max(120, e.width - 28)))
    return lb


# ───────────────────────── Ứng dụng ─────────────────────────
class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("ViSpellCheck Pro - Kiểm tra chính tả tiếng Việt")
        self.geometry("1200x780")
        self.minsize(1000, 680)
        self.configure(bg=PAL["bg"])
        self.checker = Checker()
        self.live = OfficeLive()
        self.live_issues, self.file_issues = [], []
        self.fp = None
        self.auto = False
        self.conn = (0, 0)               # (số tài liệu Word, số workbook Excel)
        self._style()
        self._build()
        self.protocol("WM_DELETE_WINDOW", self._close)
        self.after(800, self._tick)

    # ── giao diện ─────────────────────────────────────────
    @property
    def issues(self):
        return self.live_issues + self.file_issues

    def _style(self):
        st = ttk.Style(self)
        try:
            st.theme_use("clam")
        except tk.TclError:
            pass
        st.configure("Treeview", rowheight=30, font=(FONT, 10), background="#ffffff",
                     fieldbackground="#ffffff", foreground=PAL["text"], borderwidth=0)
        st.configure("Treeview.Heading", font=(FONT, 10, "bold"), background="#1e3a8a",
                     foreground="#ffffff", relief="raised", padding=7, borderwidth=2)
        st.map("Treeview.Heading", background=[("active", "#2563eb")])
        st.map("Treeview", background=[("selected", "#bfdbfe")],
               foreground=[("selected", "#0f172a")])

    def _build(self):
        self.columnconfigure(0, weight=1)
        self.rowconfigure(2, weight=1)

        # Banner chuyển màu + đèn trạng thái Word/Excel
        self.header = tk.Canvas(self, height=84, highlightthickness=0, bd=0, bg="#1e3a8a")
        self.header.grid(row=0, column=0, sticky="ew")
        self.header.bind("<Configure>", self._paint_header)

        # Hàng 4 khung chức năng
        row = tk.Frame(self, bg=PAL["bg"])
        row.grid(row=1, column=0, sticky="ew", padx=14, pady=(14, 6))
        for i in range(4):
            row.columnconfigure(i, weight=1, uniform="cards")

        c1, b1 = make_card(row, "◉  QUÉT OFFICE ĐANG MỞ", BLUE)
        desc_label(b1, "Tự nhận diện mọi tài liệu Word và Excel đang mở. "
                       "Không cần thêm tệp.")
        self.btn_scan = Button3D(b1, "▶  QUÉT NGAY", self.scan_open, BLUE, height=46)
        self.btn_scan.pack(fill="x", pady=(10, 0))
        self.lbl_docs = tk.Label(b1, text="Chưa kết nối Office", bg=PAL["card"],
                                 fg=PAL["muted"], anchor="w", font=(FONT, 9, "italic"))
        self.lbl_docs.pack(fill="x", pady=(6, 0))
        c1.grid(row=0, column=0, sticky="nsew", padx=6)

        c2, b2 = make_card(row, "⟳  THEO DÕI TỰ ĐỘNG", GREEN)
        desc_label(b2, "Tự quét lại khi nội dung tài liệu thay đổi (mỗi 3 giây).")
        self.btn_auto = Button3D(b2, "ĐANG TẮT", self.toggle_auto, SLATE, height=46)
        self.btn_auto.pack(fill="x", pady=(10, 0))
        c2.grid(row=0, column=1, sticky="nsew", padx=6)

        c3, b3 = make_card(row, "✎  SỬA LỖI", DEEP_ORANGE)
        desc_label(b3, "Bản đã sửa luôn lưu thành tệp mới (_checked), tệp gốc được giữ nguyên.")
        Button3D(b3, "SỬA MỤC ĐÃ CHỌN", self.fix_selected, ORANGE, height=36).pack(
            fill="x", pady=(8, 0))
        Button3D(b3, "SỬA TẤT CẢ (GỢI Ý ĐẦU)", self.fix_all, DEEP_ORANGE, height=36).pack(
            fill="x", pady=(4, 0))
        c3.grid(row=0, column=2, sticky="nsew", padx=6)

        c4, b4 = make_card(row, "⇩  BÁO CÁO & TỆP", PURPLE)
        desc_label(b4, "Xuất danh sách lỗi, hoặc kiểm tra thêm tệp chưa mở (tùy chọn).")
        Button3D(b4, "XUẤT BÁO CÁO CSV", self.export, PURPLE, height=36).pack(
            fill="x", pady=(8, 0))
        Button3D(b4, "MỞ TỆP TỪ ĐĨA…", self.add_files, SLATE, height=36).pack(
            fill="x", pady=(4, 0))
        c4.grid(row=0, column=3, sticky="nsew", padx=6)

        # Bảng kết quả
        tc, tb = make_card(self, "DANH SÁCH LỖI NGHI NGỜ", ("#475569", "#1e293b"))
        tc.grid(row=2, column=0, sticky="nsew", padx=20, pady=(8, 4))
        top = tk.Frame(tb, bg=PAL["card"])
        top.pack(fill="x", pady=(0, 6))
        tk.Label(top, text="Bấm vào một dòng để xem gợi ý và sửa nhanh bên dưới.",
                 bg=PAL["card"], fg=PAL["muted"], font=(FONT, 9)).pack(side="left")
        self.badge_ctx = tk.Label(top, text="Ngữ cảnh: 0", bg="#d97706", fg="white",
                                  relief="raised", bd=2, padx=10, font=(FONT, 9, "bold"))
        self.badge_ctx.pack(side="right", padx=(6, 0))
        self.badge_spell = tk.Label(top, text="Chính tả: 0", bg="#dc2626", fg="white",
                                    relief="raised", bd=2, padx=10, font=(FONT, 9, "bold"))
        self.badge_spell.pack(side="right")

        wrap = tk.Frame(tb, bg=PAL["card"], relief="sunken", bd=2)
        wrap.pack(fill="both", expand=True)
        cols = ("src", "doc", "loc", "orig", "sug", "kind", "rep")
        heads = ("Nguồn", "Tài liệu", "Vị trí", "Từ nghi sai", "Gợi ý", "Loại lỗi", "Thay bằng")
        widths = (70, 190, 200, 130, 280, 90, 130)
        self.tree = ttk.Treeview(wrap, columns=cols, show="headings", selectmode="extended")
        for c, h, w in zip(cols, heads, widths):
            self.tree.heading(c, text=h)
            self.tree.column(c, width=w, anchor="w", stretch=(c == "sug"))
        sb = ttk.Scrollbar(wrap, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=sb.set)
        self.tree.pack(side="left", fill="both", expand=True)
        sb.pack(side="right", fill="y")
        self.tree.tag_configure("spell", background="#fef2f2")
        self.tree.tag_configure("context", background="#fffbeb")
        self.tree.bind("<<TreeviewSelect>>", self.on_select)

        # Chi tiết & sửa nhanh
        dc, db = make_card(self, "CHI TIẾT & SỬA NHANH", TEAL)
        dc.grid(row=3, column=0, sticky="ew", padx=20, pady=(4, 8))
        r1 = tk.Frame(db, bg=PAL["card"])
        r1.pack(fill="x")
        tk.Label(r1, text="Từ nghi sai:", bg=PAL["card"], fg=PAL["muted"],
                 font=(FONT, 10)).pack(side="left")
        self.lbl_word = tk.Label(r1, text="—", bg=PAL["card"], fg="#b91c1c",
                                 font=(FONT, 15, "bold"))
        self.lbl_word.pack(side="left", padx=10)
        self.lbl_loc = tk.Label(r1, text="", bg=PAL["card"], fg=PAL["muted"],
                                font=(FONT, 9))
        self.lbl_loc.pack(side="left")
        r2 = tk.Frame(db, bg=PAL["card"])
        r2.pack(fill="x", pady=(6, 0))
        tk.Label(r2, text="Gợi ý:", bg=PAL["card"], fg=PAL["muted"],
                 font=(FONT, 10), width=11, anchor="w").pack(side="left")
        self.chips = tk.Frame(r2, bg=PAL["card"])
        self.chips.pack(side="left", fill="x")
        r3 = tk.Frame(db, bg=PAL["card"])
        r3.pack(fill="x", pady=(6, 0))
        tk.Label(r3, text="Thay bằng:", bg=PAL["card"], fg=PAL["muted"],
                 font=(FONT, 10), width=11, anchor="w").pack(side="left")
        self.rep_var = tk.StringVar()
        tk.Entry(r3, textvariable=self.rep_var, relief="sunken", bd=3, font=(FONT, 11),
                 bg="#f8fafc").pack(side="left", fill="x", expand=True, ipady=3)
        Button3D(r3, "ÁP DỤNG", self.fix_manual, GREEN, height=32, width=130).pack(
            side="left", padx=(10, 0))

        # Thanh trạng thái
        bar = tk.Frame(self, bg=PAL["footer"])
        bar.grid(row=4, column=0, sticky="ew")
        self.dot = tk.Label(bar, text="●", fg="#22c55e", bg=PAL["footer"], font=(FONT, 12))
        self.dot.pack(side="left", padx=(12, 4), pady=6)
        self.status = tk.Label(bar, text="", bg=PAL["footer"], fg="#e2e8f0", anchor="w",
                               font=(FONT, 10))
        self.status.pack(side="left", fill="x", expand=True)
        self.set_status("Sẵn sàng. Mở tài liệu Word/Excel rồi bấm QUÉT NGAY.")

    def _paint_header(self, e=None):
        c = self.header
        c.delete("all")
        w = c.winfo_width() or 1200
        h = 84
        draw_gradient(c, 0, 0, w, h, "#1e3a8a", "#0891b2", vertical=False, step=3)
        c.create_line(0, 0, w, 0, fill="#93c5fd")
        c.create_line(0, h - 1, w, h - 1, fill="#0b1f4d")
        c.create_text(24, 34, anchor="w", text="ViSpellCheck Pro", fill="#ffffff",
                      font=(FONT, 24, "bold"))
        c.create_text(26, 63, anchor="w", fill="#bae6fd", font=(FONT, 10),
                      text="Kiểm tra chính tả tiếng Việt cho Word và Excel đang mở")
        x = w - 20
        for label, n in (("Excel", self.conn[1]), ("Word", self.conn[0])):
            pw = 138
            x0, y0, x1, y1 = x - pw, 26, x, 58
            c.create_polygon(round_pts(x0 + 2, y0 + 3, x1 + 2, y1 + 3, 14), smooth=True,
                             fill="#0a2a5e", outline="")
            c.create_polygon(round_pts(x0, y0, x1, y1, 14), smooth=True,
                             fill="#16407f", outline="#60a5fa")
            c.create_oval(x0 + 12, y0 + 10, x0 + 24, y0 + 22,
                          fill="#22c55e" if n else "#94a3b8", outline="#ffffff")
            c.create_text(x0 + 32, (y0 + y1) / 2, anchor="w", fill="#ffffff",
                          font=(FONT, 10, "bold"), text=f"{label}: {n} tệp")
            x = x0 - 10

    def set_status(self, text, level="ok"):
        self.dot.configure(fg={"ok": "#22c55e", "busy": "#fbbf24", "error": "#ef4444"}[level])
        self.status.configure(text=text)

    def _update_conn(self, info):
        conn = (len(info["word"]), len(info["excel"]))
        names = info["word"] + info["excel"]
        self.lbl_docs.configure(
            text=("Đang mở: " + "; ".join(names)) if names else "Chưa phát hiện Word/Excel đang mở")
        if conn != self.conn:
            self.conn = conn
            self._paint_header()

    # ── quét Office đang mở ───────────────────────────────
    def scan_open(self, quiet=False):
        self.set_status("Đang quét Word/Excel đang mở...", "busy")
        self.update_idletasks()
        try:
            found, info = self.live.scan(lambda t: check_text(self.checker, t))
        except Exception as e:  # noqa: BLE001
            self.set_status(str(e), "error")
            if not quiet:
                messagebox.showerror("Office", str(e))
            return
        self.live_issues = found
        self.fp = self.live.fingerprint()
        self._update_conn(info)
        self.refresh()
        if not info["total"]:
            self.set_status("Chưa phát hiện Word/Excel nào đang mở. Hãy mở tài liệu "
                            "rồi bấm QUÉT NGAY.", "error")
            if not quiet:
                messagebox.showinfo("ViSpellCheck", "Chưa có tài liệu Word/Excel nào đang mở.")
            return
        msg = (f"Đã quét {info['total']} tài liệu ({len(info['word'])} Word, "
               f"{len(info['excel'])} Excel): {len(found)} vị trí nghi lỗi — "
               f"từ sai được tô đỏ trong tài liệu.")
        if info["errors"]:
            msg += "  Bỏ qua: " + "; ".join(info["errors"])
        self.set_status(msg, "ok")

    def toggle_auto(self):
        self.auto = not self.auto
        self.btn_auto.update_button("ĐANG BẬT" if self.auto else "ĐANG TẮT",
                                    GREEN if self.auto else SLATE)
        if self.auto:
            self.fp = None                 # buộc quét ngay ở lần kế tiếp
            self.set_status("Đã bật theo dõi tự động.")
        else:
            self.set_status("Đã tắt theo dõi tự động.")

    def _tick(self):
        try:
            info = self.live.list_open()
            self._update_conn(info)
            if self.auto and info["total"] and self.live.fingerprint() != self.fp:
                self.scan_open(quiet=True)
        except Exception:  # noqa: BLE001  (Office bận hoặc chưa có pywin32)
            pass
        self.after(3000, self._tick)

    # ── tệp trên đĩa (tùy chọn) ───────────────────────────
    def add_files(self):
        ps = filedialog.askopenfilenames(
            filetypes=[("Tài liệu Word/Excel", "*.docx *.xlsx *.xlsm"), ("Tất cả", "*.*")])
        if not ps:
            return
        self.set_status("Đang quét tệp...", "busy")
        self.update_idletasks()
        found = []
        for fn in ps:
            ext = Path(fn).suffix.lower()
            try:
                if ext == ".docx":
                    found += scan_docx(self.checker, fn)
                elif ext in (".xlsx", ".xlsm"):
                    found += scan_excel(self.checker, fn)
            except Exception as e:  # noqa: BLE001
                messagebox.showerror("Lỗi đọc tệp", f"{Path(fn).name}\n{e}")
        self.file_issues = found
        self.refresh()
        self.set_status(f"Đã quét {len(ps)} tệp trên đĩa: {len(found)} vị trí nghi lỗi.")

    # ── bảng kết quả ──────────────────────────────────────
    def refresh(self):
        self.tree.delete(*self.tree.get_children())
        names = {"word": "Word", "excel": "Excel", "file": "Tệp"}
        for i, x in enumerate(self.issues):
            self.tree.insert("", "end", iid=str(i), tags=(x.kind,), values=(
                names[x.source], x.doc_name, x.location, x.original,
                ", ".join(x.suggestions),
                "Chính tả" if x.kind == "spell" else "Ngữ cảnh", x.selected))
        sp = sum(1 for x in self.issues if x.kind == "spell")
        self.badge_spell.configure(text=f"Chính tả: {sp}")
        self.badge_ctx.configure(text=f"Ngữ cảnh: {len(self.issues) - sp}")
        self._show_detail(None)

    def _show_detail(self, x):
        for w in self.chips.winfo_children():
            w.destroy()
        if x is None:
            self.lbl_word.configure(text="—")
            self.lbl_loc.configure(text="")
            self.rep_var.set("")
            return
        self.lbl_word.configure(text=x.original)
        self.lbl_loc.configure(text=f"{x.doc_name} · {x.location}")
        self.rep_var.set(x.selected)
        for s in x.suggestions:
            Button3D(self.chips, s, lambda s=s: self.rep_var.set(s), TEAL, height=28,
                     width=max(70, 11 * len(s) + 26), font=(FONT, 10, "bold")).pack(
                side="left", padx=(0, 6))
        if not x.suggestions:
            tk.Label(self.chips, text="(không có gợi ý — hãy nhập tay)", bg=PAL["card"],
                     fg=PAL["muted"], font=(FONT, 9, "italic")).pack(side="left")

    def on_select(self, _=None):
        sel = self.tree.selection()
        self._show_detail(self.issues[int(sel[0])] if sel else None)

    # ── sửa lỗi ───────────────────────────────────────────
    def fix_selected(self):
        sel = self.tree.selection()
        if not sel:
            messagebox.showinfo("Sửa lỗi", "Hãy chọn ít nhất một dòng trong danh sách.")
            return
        self._do_fix([self.issues[int(i)] for i in sel])

    def fix_manual(self):
        sel = self.tree.selection()
        new = self.rep_var.get().strip()
        if not sel:
            messagebox.showinfo("Sửa lỗi", "Hãy chọn một dòng trong danh sách.")
            return
        if not new:
            messagebox.showwarning("Sửa lỗi", "Hãy nhập từ thay thế.")
            return
        x = self.issues[int(sel[0])]
        x.selected = new
        self._do_fix([x])

    def fix_all(self):
        todo = [x for x in self.issues if x.suggestions or x.selected]
        if not todo:
            return
        if not messagebox.askyesno("Xác nhận", f"Sửa {len(todo)} vị trí theo gợi ý đầu tiên?"):
            return
        for x in todo:
            x.selected = x.selected or x.suggestions[0]
        self._do_fix(todo)

    def _do_fix(self, chosen):
        chosen = [x for x in chosen if x.selected]
        if not chosen:
            return
        saved, errors = [], []
        files = [x for x in chosen if x.source == "file"]
        live = [x for x in chosen if x.source != "file"]
        if files:
            for x in files:
                x.apply(x.selected)
            for x in {id(y.save): y for y in files}.values():
                saved.append(x.save())
            self.file_issues = [y for y in self.file_issues if y not in files]
        if live:
            self.live.restore_marks()
            groups = {}
            for x in live:
                groups.setdefault(id(x.doc), []).append(x)
            for grp in groups.values():
                first = grp[0]
                ext = ".docx" if first.source == "word" else (
                    ".xlsm" if first.doc_name.lower().endswith(".xlsm") else ".xlsx")
                opts = dict(parent=self, title="Lưu bản đã sửa thành tệp mới",
                            initialfile=Path(first.doc_name).stem + "_checked" + ext,
                            defaultextension=ext,
                            filetypes=[("Tài liệu Office", "*" + ext), ("Tất cả", "*.*")])
                if first.doc_dir:
                    opts["initialdir"] = first.doc_dir
                target = filedialog.asksaveasfilename(**opts)
                if not target:
                    continue
                try:
                    self.live.fix_copy(first.source, first.doc, grp, target)
                    saved.append(target)
                except Exception as e:  # noqa: BLE001
                    errors.append(f"{first.doc_name}: {e}")
        if live:
            self.scan_open(quiet=True)
        else:
            self.refresh()
        if saved:
            self.set_status("Đã lưu bản đã sửa: " + "; ".join(Path(s).name for s in saved))
            messagebox.showinfo("Đã lưu", "Tệp gốc được giữ nguyên.\nBản đã sửa:\n"
                                + "\n".join(saved))
        if errors:
            self.set_status("Có lỗi khi sửa: " + "; ".join(errors), "error")
            messagebox.showerror("Không thể lưu bản sửa", "\n".join(errors))

    # ── báo cáo / đóng ────────────────────────────────────
    def export(self):
        if not self.issues:
            messagebox.showinfo("Báo cáo", "Chưa có kết quả kiểm tra.")
            return
        p = filedialog.asksaveasfilename(defaultextension=".csv", filetypes=[("CSV", "*.csv")],
                                         initialfile="ViSpellCheck_Report.csv")
        if not p:
            return
        with open(p, "w", newline="", encoding="utf-8-sig") as f:
            w = csv.writer(f)
            w.writerow(["Nguồn", "Tài liệu", "Vị trí", "Nội dung lỗi", "Gợi ý", "Loại lỗi",
                        "Thay bằng"])
            for x in self.issues:
                w.writerow([x.source, x.doc_name, x.location, x.original,
                            "; ".join(x.suggestions), x.kind, x.selected])
        messagebox.showinfo("Báo cáo", f"Đã xuất: {p}")

    def _close(self):
        try:
            self.live.restore_marks()      # trả lại màu chữ gốc trong tài liệu đang mở
        except Exception:  # noqa: BLE001
            pass
        self.destroy()


if __name__ == "__main__":
    App().mainloop()
