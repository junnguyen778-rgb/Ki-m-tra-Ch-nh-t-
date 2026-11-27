# -*- coding: utf-8 -*-
"""ViSpellCheck Pro: kiểm tra Word DOCX và Excel XLSX/XLSM."""
import csv, os, threading, tkinter as tk
from tkinter import ttk, filedialog, messagebox
from pathlib import Path
from spell_engine import Checker, WORD_RE, apply_case

try:
    from docx import Document
except ImportError: Document = None
try:
    import openpyxl
except ImportError: openpyxl = None

class Issue:
    def __init__(self, file, location, original, suggestions, kind, setter, group_key="", start=0, end=0):
        self.file=file; self.location=location; self.original=original
        self.suggestions=suggestions; self.kind=kind; self.setter=setter
        self.group_key=group_key; self.start=start; self.end=end
        self.selected=suggestions[0] if suggestions else ""

def check_text(checker, text):
    toks=list(WORD_RE.finditer(text)); out=[]
    words=[m.group() for m in toks]
    for i,m in enumerate(toks):
        prev=words[i-1].lower() if i and text[toks[i-1].end():m.start()].isspace() else None
        nxt=words[i+1].lower() if i+1<len(toks) and text[m.end():toks[i+1].start()].isspace() else None
        kind,sugs=checker.check_word(m.group(),prev,nxt)
        if kind:
            out.append((m.start(),m.end(),m.group(),[apply_case(m.group(),x) for x in sugs],kind))
    return out

class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("ViSpellCheck Pro - Kiểm tra chính tả tiếng Việt")
        self.geometry("1120x700"); self.minsize(900,560)
        self.checker=Checker(); self.issues=[]; self.files=[]
        self._ui()
    def _ui(self):
        top=ttk.Frame(self,padding=10); top.pack(fill="x")
        ttk.Label(top,text="ViSpellCheck Pro",font=("Segoe UI",18,"bold")).pack(side="left")
        ttk.Button(top,text="Office đang mở",command=self.open_live).pack(side="right",padx=4)
        ttk.Button(top,text="Thêm tệp",command=self.add_files).pack(side="right",padx=4)
        ttk.Button(top,text="Kiểm tra",command=self.start_check).pack(side="right",padx=4)
        ttk.Button(top,text="Xuất báo cáo CSV",command=self.export).pack(side="right",padx=4)
        ttk.Button(top,text="Sửa lỗi đã chọn",command=self.fix_selected).pack(side="right",padx=4)
        ttk.Button(top,text="Sửa tất cả theo gợi ý đầu",command=self.fix_all).pack(side="right",padx=4)
        self.file_label=ttk.Label(self,text="Chưa chọn tệp",padding=(12,0)); self.file_label.pack(anchor="w")
        cols=("file","location","original","suggestions","kind","replacement")
        self.tree=ttk.Treeview(self,columns=cols,show="headings",selectmode="extended")
        heads={"file":"Tệp","location":"Vị trí","original":"Từ/cụm từ","suggestions":"Gợi ý","kind":"Loại lỗi","replacement":"Thay bằng"}
        widths={"file":190,"location":130,"original":130,"suggestions":270,"kind":130,"replacement":160}
        for c in cols:self.tree.heading(c,text=heads[c]); self.tree.column(c,width=widths[c],anchor="w")
        self.tree.pack(fill="both",expand=True,padx=10,pady=8)
        bottom=ttk.Frame(self,padding=10); bottom.pack(fill="x")
        ttk.Label(bottom,text="Chọn một dòng rồi nhập nội dung cột 'Thay bằng' ở ô bên dưới để sửa thủ công:").pack(anchor="w")
        row=ttk.Frame(bottom); row.pack(fill="x",pady=4)
        self.replacement=tk.StringVar()
        ttk.Entry(row,textvariable=self.replacement).pack(side="left",fill="x",expand=True)
        ttk.Button(row,text="Áp dụng cho dòng chọn",command=self.fix_manual).pack(side="left",padx=6)
        self.status=tk.StringVar(value="Sẵn sàng. Tệp gốc không bị thay đổi; bản đã sửa được lưu thành tệp mới.")
        ttk.Label(self,textvariable=self.status,relief="sunken",anchor="w",padding=5).pack(fill="x",side="bottom")
        self.tree.bind("<<TreeviewSelect>>",self.on_select)
    def open_live(self):
        try:
            from office_live import OfficeLiveWindow
            OfficeLiveWindow(self, self.checker)
        except Exception as e:
            messagebox.showerror("Office Live", str(e))

    def add_files(self):
        ps=filedialog.askopenfilenames(filetypes=[("Tài liệu Word/Excel","*.docx *.xlsx *.xlsm"),("Tất cả","*.*")])
        if ps:
            self.files=list(ps); self.file_label.config(text=f"Đã chọn {len(self.files)} tệp: "+"; ".join(Path(x).name for x in self.files))
            self.issues=[]; self.refresh()
    def start_check(self):
        if not self.files:
            messagebox.showinfo("ViSpellCheck","Hãy chọn tệp Word hoặc Excel trước."); return
        self.status.set("Đang kiểm tra...")
        threading.Thread(target=self._scan,daemon=True).start()
    def _scan(self):
        issues=[]
        for fn in self.files:
            ext=Path(fn).suffix.lower()
            try:
                if ext==".docx": issues.extend(self.scan_docx(fn))
                elif ext in (".xlsx",".xlsm"): issues.extend(self.scan_excel(fn))
            except Exception as e:
                self.after(0,lambda f=fn,err=str(e):messagebox.showerror("Lỗi đọc tệp",f"{Path(f).name}\n{err}"))
        self.issues=issues
        self.after(0,self._scan_done)
    def scan_docx(self,fn):
        if Document is None: raise RuntimeError("Thiếu python-docx.")
        doc=Document(fn); out=[]
        def scan_paragraph(p,loc):
            # Mỗi run được quét riêng để giữ định dạng khi thay; lỗi bị chia qua nhiều run có thể không nhận diện.
            for ri,run in enumerate(p.runs):
                base=run.text
                for s,e,w,sugs,kind in check_text(self.checker,base):
                    def setter(new, r=run, a=s, b=e, old=w):
                        r.text=r.text[:a]+new+r.text[b:]
                    out.append(Issue(fn,f"{loc} / Run {ri+1}: ký tự {s+1}",w,sugs,kind,setter,
                                     group_key=f"{loc}|run{ri}",start=s,end=e))
        for pi,p in enumerate(doc.paragraphs,1):scan_paragraph(p,f"Đoạn {pi}")
        for ti,table in enumerate(doc.tables,1):
            for ri,row in enumerate(table.rows,1):
                for ci,cell in enumerate(row.cells,1):
                    for pi,p in enumerate(cell.paragraphs,1):scan_paragraph(p,f"Bảng {ti}, hàng {ri}, cột {ci}, đoạn {pi}")
        # Header/footer
        for si,sec in enumerate(doc.sections,1):
            for kind,paras in (("Đầu trang",sec.header.paragraphs),("Chân trang",sec.footer.paragraphs)):
                for pi,p in enumerate(paras,1):scan_paragraph(p,f"Trang {si} - {kind}, đoạn {pi}")
        def save():
            target=Path(fn).with_name(Path(fn).stem+"_checked"+Path(fn).suffix); doc.save(target); return str(target)
        # setters mutate in-memory doc; one shared save callback is attached below
        for issue in out: issue.save_document=save
        return out
    def scan_excel(self,fn):
        if openpyxl is None:raise RuntimeError("Thiếu openpyxl.")
        wb=openpyxl.load_workbook(fn,keep_vba=fn.lower().endswith(".xlsm"))
        out=[]
        for ws in wb.worksheets:
            for row in ws.iter_rows():
                for cell in row:
                    if isinstance(cell.value,str) and cell.value and not cell.value.startswith("="):
                        val=cell.value
                        for s,e,w,sugs,kind in check_text(self.checker,val):
                            def setter(new,c=cell,a=s,b=e):
                                c.value=c.value[:a]+new+c.value[b:]
                            out.append(Issue(fn,f"{ws.title}!{cell.coordinate}",w,sugs,kind,setter,
                                              group_key=f"{ws.title}!{cell.coordinate}",start=s,end=e))
        for issue in out:issue.workbook=wb
        # store workbook per issue; group-saving occurs after fixes
        return out
    def _scan_done(self):
        # retain source document/workbook objects through callbacks and save by source
        self.refresh()
        self.status.set(f"Hoàn tất: phát hiện {len(self.issues)} vị trí nghi lỗi.")
        if not self.issues:messagebox.showinfo("Kết quả","Không phát hiện lỗi theo từ điển hiện tại.")
    def refresh(self):
        self.tree.delete(*self.tree.get_children())
        for i,x in enumerate(self.issues):
            self.tree.insert("", "end", iid=str(i), values=(Path(x.file).name,x.location,x.original,
                ", ".join(x.suggestions), "Chính tả" if x.kind=="spell" else "Ngữ cảnh",x.selected))
    def on_select(self,_):
        sel=self.tree.selection()
        if sel:self.replacement.set(self.issues[int(sel[0])].selected)
    def fix_selected(self):
        sel=self.tree.selection()
        if not sel:messagebox.showinfo("Sửa lỗi","Hãy chọn ít nhất một dòng.");return
        chosen=[self.issues[int(i)] for i in sel if self.issues[int(i)].selected]
        # Áp dụng từ phải sang trái trong cùng một đoạn/ô để không lệch chỉ số ký tự.
        for x in sorted(chosen,key=lambda z:(z.file,z.group_key,-z.start)):
            x.setter(x.selected)
        self.save_fixed({x.file for x in chosen})
    def fix_manual(self):
        sel=self.tree.selection()
        if not sel:return
        x=self.issues[int(sel[0])]; new=self.replacement.get().strip()
        if not new:messagebox.showwarning("Sửa lỗi","Hãy nhập từ thay thế.");return
        x.setter(new); x.selected=new
        self.save_fixed({x.file})
    def fix_all(self):
        if not self.issues:return
        if not messagebox.askyesno("Xác nhận",f"Sửa {len(self.issues)} vị trí theo gợi ý đầu tiên?"):return
        chosen=[x for x in self.issues if x.suggestions]
        for x in sorted(chosen,key=lambda z:(z.file,z.group_key,-z.start)):
            x.selected=x.suggestions[0]
            x.setter(x.selected)
        self.save_fixed({x.file for x in chosen})
    def save_fixed(self,files):
        saved=[]
        for fn in files:
            related=[x for x in self.issues if x.file==fn]
            ext=Path(fn).suffix.lower()
            target=Path(fn).with_name(Path(fn).stem+"_checked"+Path(fn).suffix)
            if ext==".docx":
                if related and hasattr(related[0],"save_document"):related[0].save_document()
            else:
                if related and hasattr(related[0],"workbook"):
                    related[0].workbook.save(target)
            saved.append(str(target))
        self.refresh()
        self.status.set("Đã lưu bản đã sửa: "+"; ".join(Path(x).name for x in saved))
        messagebox.showinfo("Đã lưu","Tệp gốc được giữ nguyên.\nBản đã sửa:\n"+"\n".join(saved))
    def export(self):
        if not self.issues:messagebox.showinfo("Báo cáo","Chưa có kết quả kiểm tra.");return
        p=filedialog.asksaveasfilename(defaultextension=".csv",filetypes=[("CSV","*.csv")],initialfile="ViSpellCheck_Report.csv")
        if not p:return
        with open(p,"w",newline="",encoding="utf-8-sig") as f:
            w=csv.writer(f);w.writerow(["Tệp","Vị trí","Nội dung lỗi","Gợi ý","Loại lỗi","Thay bằng"])
            for x in self.issues:w.writerow([x.file,x.location,x.original,"; ".join(x.suggestions),x.kind,x.selected])
        messagebox.showinfo("Báo cáo",f"Đã xuất: {p}")

if __name__=="__main__":
    App().mainloop()
