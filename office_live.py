# -*- coding: utf-8 -*-
"""Theo doi Word/Excel dang mo, danh dau chu sai mau do, sua va luu ban sao.
Can Microsoft Office desktop + pywin32. Tat ca COM calls chay tren thread Tkinter.
"""
import re, os, tkinter as tk
from tkinter import ttk, messagebox, filedialog
from pathlib import Path
from spell_engine import Checker, WORD_RE, apply_case

RED = 255  # RGB(255,0,0), dung cho Word va Excel

class LiveIssue:
    def __init__(self, app_kind, app_obj, location, original, suggestions, kind,
                 start=0, end=0, sheet=None, cell=None, paragraph=None):
        self.app_kind=app_kind; self.app_obj=app_obj; self.location=location
        self.original=original; self.suggestions=suggestions; self.kind=kind
        self.start=start; self.end=end; self.sheet=sheet; self.cell=cell; self.paragraph=paragraph
        self.replacement=suggestions[0] if suggestions else ''

class OfficeLiveWindow(tk.Toplevel):
    """Cua so giam sat cac tai lieu dang active trong Word va Excel."""
    def __init__(self, master, checker=None):
        super().__init__(master)
        self.title('ViSpellCheck Pro - Office Live')
        self.geometry('1080x650'); self.minsize(850,500)
        self.checker=checker or Checker()
        self.issues=[]; self.marks=[]; self.last_fingerprint={}
        self.auto=tk.BooleanVar(value=True)
        self._build()
        self.after(1200,self._tick)

    def _build(self):
        bar=ttk.Frame(self,padding=8); bar.pack(fill='x')
        ttk.Button(bar,text='Kiểm tra ngay',command=self.scan_active).pack(side='left',padx=3)
        ttk.Checkbutton(bar,text='Tự động kiểm tra (mỗi 3 giây)',variable=self.auto).pack(side='left',padx=10)
        ttk.Button(bar,text='Sửa lỗi chọn và lưu bản sao',command=self.fix_selected).pack(side='right',padx=3)
        ttk.Button(bar,text='Làm mới danh sách',command=self.scan_active).pack(side='right',padx=3)
        self.status=tk.StringVar(value='Đang chờ Word hoặc Excel. Hãy mở tài liệu trong Microsoft Office.')
        ttk.Label(self,textvariable=self.status,padding=(10,4)).pack(fill='x')
        cols=('app','location','word','suggestions','type','replacement')
        self.tree=ttk.Treeview(self,columns=cols,show='headings',selectmode='extended')
        labels={'app':'Ứng dụng','location':'Vị trí','word':'Từ nghi sai','suggestions':'Gợi ý','type':'Loại lỗi','replacement':'Thay bằng'}
        widths={'app':90,'location':240,'word':130,'suggestions':260,'type':110,'replacement':160}
        for c in cols:self.tree.heading(c,text=labels[c]);self.tree.column(c,width=widths[c],anchor='w')
        self.tree.pack(fill='both',expand=True,padx=8,pady=6)
        edit=ttk.Frame(self,padding=8);edit.pack(fill='x')
        ttk.Label(edit,text='Từ thay thế cho dòng đang chọn:').pack(side='left')
        self.replacement=tk.StringVar()
        ttk.Entry(edit,textvariable=self.replacement).pack(side='left',fill='x',expand=True,padx=8)
        ttk.Button(edit,text='Áp dụng từ nhập',command=self.apply_manual).pack(side='right')
        self.tree.bind('<<TreeviewSelect>>',self._selected)
        note=('Đánh dấu màu đỏ chỉ tồn tại trong phiên Office đang mở. '
              'Khi lưu bản sửa, chương trình tạo tệp _checked; tệp nguồn trên đĩa không bị ghi đè.')
        ttk.Label(self,text=note,wraplength=1020,foreground='#666',padding=8).pack(fill='x')

    def _com_apps(self):
        try:
            import pythoncom, win32com.client
            pythoncom.CoInitialize()
            try: word=win32com.client.GetActiveObject('Word.Application')
            except Exception: word=None
            try: excel=win32com.client.GetActiveObject('Excel.Application')
            except Exception: excel=None
            return word,excel
        except Exception as e:
            raise RuntimeError('Không kết nối được Office COM. Hãy cài Microsoft Office desktop và pywin32. Chi tiết: '+str(e))

    def _restore_marks(self):
        # Khoi phuc mau chu goc truoc khi quet lai, tranh de lai mau do cu.
        for obj,old in reversed(self.marks):
            try: obj.Font.Color=old
            except Exception: pass
        self.marks=[]

    def _fingerprint(self, word, excel):
        vals=[]
        if word is not None:
            try:
                d=word.ActiveDocument; vals.append(('word',str(d.FullName),str(d.Content.Text)))
            except Exception: pass
        if excel is not None:
            try:
                w=excel.ActiveWorkbook; parts=[str(w.FullName)]
                for sh in w.Worksheets:
                    ur=sh.UsedRange
                    parts.append(str(sh.Name)+':'+str(ur.Value2))
                vals.append(('excel',*parts))
            except Exception: pass
        return {x[0]:hash(tuple(x[1:])) for x in vals}

    def _tick(self):
        if not self.winfo_exists():return
        if self.auto.get():
            try:
                word,excel=self._com_apps(); fp=self._fingerprint(word,excel)
                if fp and fp!=self.last_fingerprint:self.scan_active()
            except Exception as e:self.status.set(str(e))
        self.after(3000,self._tick)

    def scan_active(self):
        try: word,excel=self._com_apps()
        except Exception as e:
            self.status.set(str(e));return
        self._restore_marks(); self.issues=[]
        if word is not None:
            try:self._scan_word(word)
            except Exception as e:self.status.set('Lỗi đọc Word: '+str(e))
        if excel is not None:
            try:self._scan_excel(excel)
            except Exception as e:self.status.set('Lỗi đọc Excel: '+str(e))
        self.last_fingerprint=self._fingerprint(word,excel)
        self._refresh()
        self.status.set(f'Đã kiểm tra Office đang mở: {len(self.issues)} vị trí nghi lỗi. Từ sai đã được đánh dấu đỏ.')

    def _check_string(self,text,prev=None,nxt=None):
        matches=list(WORD_RE.finditer(text)); words=[m.group() for m in matches]; out=[]
        for i,m in enumerate(matches):
            p=words[i-1].lower() if i and text[matches[i-1].end():m.start()].isspace() else None
            n=words[i+1].lower() if i+1<len(matches) and text[m.end():matches[i+1].start()].isspace() else None
            kind,sugs=self.checker.check_word(m.group(),p,n)
            if kind:out.append((m.start(),m.end(),m.group(),[apply_case(m.group(),x) for x in sugs],kind))
        return out

    def _mark_word_range(self,rng):
        try:
            old=rng.Font.Color
            self.marks.append((rng,old));rng.Font.Color=RED
        except Exception:pass

    def _scan_word(self,app):
        doc=app.ActiveDocument
        # Paragraph.Range bao gom ca cac doan trong bang o noi dung chinh.
        paragraphs=doc.Paragraphs
        for pi in range(1,paragraphs.Count+1):
            para=paragraphs.Item(pi); txt=str(para.Range.Text or '').rstrip('\r\x07')
            if not txt:continue
            base=int(para.Range.Start)
            for s,e,w,sugs,kind in self._check_string(txt):
                rg=doc.Range(base+s,base+e)
                self._mark_word_range(rg)
                self.issues.append(LiveIssue('Word',doc,f'Đoạn {pi}, ký tự {s+1}',w,sugs,kind,
                                             start=base+s,end=base+e,paragraph=pi))

    def _scan_excel(self,app):
        wb=app.ActiveWorkbook
        for si in range(1,wb.Worksheets.Count+1):
            ws=wb.Worksheets.Item(si); used=ws.UsedRange
            # Duyet tung o trong UsedRange, chi lay gia tri chu; bo qua cong thuc.
            for row in range(1,used.Rows.Count+1):
                for col in range(1,used.Columns.Count+1):
                    cell=used.Cells(row,col); val=cell.Value2
                    if not isinstance(val,str) or not val or val.startswith('='):continue
                    for s,e,w,sugs,kind in self._check_string(val):
                        try:
                            chars=cell.Characters(s+1,e-s); old=chars.Font.Color
                            self.marks.append((chars,old));chars.Font.Color=RED
                        except Exception:pass
                        self.issues.append(LiveIssue('Excel',wb,f'{ws.Name}!{cell.Address(False,False)}',w,sugs,kind,
                                                     start=s,end=e,sheet=str(ws.Name),cell=str(cell.Address(False,False))))

    def _refresh(self):
        self.tree.delete(*self.tree.get_children())
        for i,x in enumerate(self.issues):
            self.tree.insert('', 'end',iid=str(i),values=(x.app_kind,x.location,x.original,
                ', '.join(x.suggestions),'Chính tả' if x.kind=='spell' else 'Ngữ cảnh',x.replacement))

    def _selected(self,_=None):
        sel=self.tree.selection()
        if sel:self.replacement.set(self.issues[int(sel[0])].replacement)

    def _target_path(self, source, ext):
        p=Path(source) if source else None
        initial=(p.stem+'_checked'+ext) if p and p.name else 'Tai_lieu_checked'+ext
        return filedialog.asksaveasfilename(parent=self,title='Lưu bản đã sửa thành tệp mới',
            initialfile=initial,defaultextension=ext,
            filetypes=[('Tài liệu Office', '*'+ext),('Tất cả','*.*')])

    def _apply_word_copy(self, doc, selected, target):
        # SaveAs2 sang ten moi truoc khi sua: tep goc tren dia khong bi ghi de.
        doc.SaveAs2(FileName=target,FileFormat=16)
        for x in sorted(selected,key=lambda z:z.start,reverse=True):
            if not x.replacement:continue
            rg=doc.Range(int(x.start),int(x.end));rg.Text=x.replacement
        doc.Save()

    def _apply_excel_copy(self, app, wb, selected, target):
        # SaveCopyAs tao ban sao tu workbook dang mo; mo ban sao de sua, workbook goc khong doi.
        wb.SaveCopyAs(target)
        copy_wb=app.Workbooks.Open(target)
        grouped={}
        for x in selected:
            if x.replacement:grouped.setdefault((x.sheet,x.cell),[]).append(x)
        for (sheet,addr),items in grouped.items():
            cell=copy_wb.Worksheets(sheet).Range(addr)
            value=str(cell.Value2 or '')
            for x in sorted(items,key=lambda z:z.start,reverse=True):
                value=value[:x.start]+x.replacement+value[x.end:]
            cell.Value2=value
        copy_wb.Save()
        try:copy_wb.Activate()
        except Exception:pass

    def _fix(self, selected):
        if not selected:return
        # Mot lan thao tac chi sua mot loai Office de tranh luu nham tai lieu.
        kinds={x.app_kind for x in selected}
        if len(kinds)>1:
            messagebox.showwarning('Chọn riêng Word hoặc Excel','Hãy chọn các lỗi thuộc cùng một ứng dụng trong mỗi lần lưu bản sao.');return
        kind=next(iter(kinds)); self._restore_marks()
        try:
            word,excel=self._com_apps()
            if kind=='Word':
                doc=word.ActiveDocument
                if any(x.app_obj.FullName!=doc.FullName for x in selected):raise RuntimeError('Tài liệu Word đang active đã thay đổi. Hãy kiểm tra lại.')
                target=self._target_path(doc.FullName,'.docx')
                if not target:return
                self._apply_word_copy(doc,selected,target)
            else:
                wb=excel.ActiveWorkbook
                if any(x.app_obj.FullName!=wb.FullName for x in selected):raise RuntimeError('Workbook Excel đang active đã thay đổi. Hãy kiểm tra lại.')
                target=self._target_path(wb.FullName,'.xlsm' if str(wb.Name).lower().endswith('.xlsm') else '.xlsx')
                if not target:return
                self._apply_excel_copy(excel,wb,selected,target)
            messagebox.showinfo('Đã tạo bản sửa','Đã lưu bản sao đã sửa tại:\n'+target+'\n\nTệp gốc không bị ghi đè.')
            self.scan_active()
        except Exception as e:
            messagebox.showerror('Không thể lưu bản sửa',str(e))

    def fix_selected(self):
        ids=self.tree.selection()
        if not ids:messagebox.showinfo('Sửa lỗi','Hãy chọn ít nhất một lỗi.');return
        self._fix([self.issues[int(i)] for i in ids])

    def apply_manual(self):
        ids=self.tree.selection()
        if not ids:return
        new=self.replacement.get().strip()
        if not new:messagebox.showwarning('Sửa lỗi','Hãy nhập từ thay thế.');return
        chosen=[self.issues[int(i)] for i in ids]
        if len(chosen)!=1:
            messagebox.showinfo('Sửa thủ công','Chế độ nhập từ thay thế chỉ áp dụng cho một lỗi mỗi lần.');return
        chosen[0].replacement=new;self._fix(chosen)

