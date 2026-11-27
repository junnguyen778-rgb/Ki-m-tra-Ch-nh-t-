# -*- coding: utf-8 -*-
"""Bộ kiểm tra âm tiết tiếng Việt và gợi ý theo ngữ cảnh."""
import json, re, unicodedata
from collections import Counter
from pathlib import Path

DATA_DIR = Path.home() / ".vi_spellcheck"
MODEL_PATH = DATA_DIR / "model.json"
USER_WORDS_PATH = DATA_DIR / "user_words.txt"
WORD_RE = re.compile(r"[^\W\d_]+", re.UNICODE)
TONES = ["", "\u0300", "\u0309", "\u0303", "\u0301", "\u0323"]
STOP_TONES = ["\u0301", "\u0323"]
MODIFIED = set("ăâêôơư")
ONSETS = ["", "b", "c", "ch", "d", "đ", "g", "gh", "gi", "h", "k", "kh", "l", "m",
          "n", "ng", "ngh", "nh", "p", "ph", "qu", "r", "s", "t", "th", "tr", "v", "x"]
OPEN = ("a e ê i y o ô ơ u ư ia ua ưa ai ay ây ao au âu eo êu iu oi ôi ơi ui ưi ưu "
        "uy oa oe uê iêu uôi ươi ươu oai oay uây oeo uya").split()
CLOSED = {
    "a":"c ch m n ng nh p t", "ă":"c m n ng p t", "â":"c m n ng p t",
    "e":"c m n ng p t", "ê":"ch m n nh p t", "i":"ch m n nh p t",
    "o":"c m n ng p t", "ô":"c m n ng p t", "ơ":"m n p t", "u":"c m n ng p t",
    "ư":"c ng t", "iê":"c m n ng p t", "uô":"c m n ng t", "ươ":"c m n ng p t",
    "oa":"c ch m n ng nh p t", "oă":"c m n ng t", "oe":"m n t", "uâ":"n ng t",
    "uê":"ch nh", "uy":"ch n nh p t", "uyê":"n t",
}
LABIO = {"oa","oe","oă","oai","oay","oeo","uâ","uê","uy","uyê","uây"}
ONSET_GROUPS = [["ngh","ng"],["tr","ch"],["gi","d","r"],["s","x"],["l","n"],
                ["c","k","q"],["g","gh"]]

def nfc(s): return unicodedata.normalize("NFC", s)
def strip_tone(s):
    s = unicodedata.normalize("NFD", s.lower().replace("đ","d"))
    return "".join(c for c in s if not unicodedata.combining(c))
def apply_case(src, cand):
    if len(src)>1 and src.isupper(): return cand.upper()
    if src[:1].isupper(): return cand[:1].upper()+cand[1:]
    return cand

def tone_positions(n, coda):
    mod=[i for i,c in enumerate(n) if c in MODIFIED]
    if mod: return [mod[-1]]
    if n in ("oa","oe","uy"): return [1] if coda else [0,1]
    if n in ("ia","ua"): return [0]
    if n=="uya": return [1]
    if coda or len(n)==1: return [len(n)-1]
    return [1] if n in ("oai","oay","oeo") else [0]

def build_syllables():
    rimes=[(n,"") for n in OPEN]
    rimes += [(n,c) for n,cs in CLOSED.items() for c in cs.split()]
    out=set()
    for onset in ONSETS:
        for nuc,coda in rimes:
            n,ons=nuc,onset
            if onset=="qu":
                if n[0] in "ou":
                    if n not in LABIO: continue
                    n=n[1:]
            elif onset=="gi":
                if n=="i": ons="g"
                elif n.startswith("iê"): n="ê"+n[2:]
                elif n in ("ia","iu"): n=n[1:]
                elif n[0] in "iy": continue
            elif onset=="" and n.startswith("iê"): n="yê"+n[2:]
            front=n[0] in "eêiy"
            if onset in ("k","gh","ngh") and not front: continue
            if onset in ("c","g","ng") and front: continue
            tones=STOP_TONES if coda in ("c","ch","p","t") else TONES
            for pos in tone_positions(n,coda):
                for tone in tones:
                    chars=list(n); chars[pos]=nfc(chars[pos]+tone)
                    out.add(nfc(ons+"".join(chars)+coda))
    return out

SEED = """tôi là sinh viên tôi là người việt nam tôi là giáo viên tôi là học sinh
tôi thích ăn cơm tôi thích ăn phở tôi thích uống cà phê hôm nay trời đẹp
hôm nay tôi đi làm chúng tôi đang học tiếng việt xin chào các bạn cảm ơn bạn
hà nội là thủ đô của việt nam tôi muốn học tiếng việt tôi không biết tôi không hiểu
người việt nam rất thân thiện sinh viên đang học bài ngày mai tôi sẽ đi học
gia đình tôi có bốn người công việc của tôi rất bận tôi đang làm việc
bảo vệ bí mật nhà nước an toàn thông tin an ninh mạng chuyển đổi số
cán bộ chiến sĩ công an nhân dân cảnh sát cơ động văn bản quy phạm pháp luật
cơ yếu mật mã quốc gia chữ ký số dữ liệu số hệ thống thông tin"""

class Model:
    def __init__(self): self.uni,self.bi=Counter(),Counter()
    def total(self): return sum(self.uni.values())
    def train(self,text):
        for chunk in re.split(r"[.,;:!?()\[\]\"“”‘’\n\r]+",text):
            toks=[nfc(t.lower()) for t in WORD_RE.findall(chunk)]
            self.uni.update(toks); self.bi.update(f"{a} {b}" for a,b in zip(toks,toks[1:]))
    def b(self,a,w): return self.bi.get(f"{a} {w}",0) if a and w else 0
    def u(self,w): return self.uni.get(w,0)
    def save(self,path=MODEL_PATH):
        path.parent.mkdir(parents=True,exist_ok=True)
        path.write_text(json.dumps({"uni":self.uni,"bi":self.bi},ensure_ascii=False),encoding="utf-8")
    @classmethod
    def load_default(cls):
        m=cls()
        try:
            d=json.loads(MODEL_PATH.read_text(encoding="utf-8")); m.uni.update(d["uni"]); m.bi.update(d["bi"])
        except (OSError,ValueError,KeyError): m.train(SEED)
        return m

def edit1(a,b):
    if abs(len(a)-len(b))>1:return False
    if len(a)>len(b):a,b=b,a
    i=0
    while i<len(a) and a[i]==b[i]:i+=1
    if len(a)==len(b):return a[i+1:]==b[i+1:]
    return a[i:]==b[i+1:]

class Checker:
    def __init__(self,extra=()):
        self.words=build_syllables()
        self.words|={nfc(w.lower()) for w in extra}
        if USER_WORDS_PATH.exists():
            self.words|={nfc(w.lower()) for w in USER_WORDS_PATH.read_text(encoding="utf-8").split()}
        self.model=Model.load_default()
        self.by_plain={}
        for w in self.words:self.by_plain.setdefault(strip_tone(w),set()).add(w)
    def candidates(self,word,edit=True):
        w=nfc(word.lower()); plain=strip_tone(w); out=[]
        def add(c):
            if c!=w and c not in out:out.append(c)
        for c in sorted(self.by_plain.get(plain,())):add(c)
        for group in ONSET_GROUPS:
            for a in sorted(group,key=len,reverse=True):
                if plain.startswith(a):
                    for b in group:
                        for c in sorted(self.by_plain.get(b+plain[len(a):],())):add(c)
                    break
        for x,y in (("i","y"),("y","i")):
            if plain.endswith(x):
                for c in sorted(self.by_plain.get(plain[:-1]+y,())):add(c)
        if edit:
            for p,cs in self.by_plain.items():
                if edit1(plain,p):
                    for c in sorted(cs):add(c)
        return out
    def rank(self,cands,prev=None,nxt=None):
        return sorted(cands,key=lambda c:(-(self.model.b(prev,c)+self.model.b(c,nxt)),-self.model.u(c)))
    def check_word(self,word,prev=None,nxt=None,context=True):
        w=nfc(word.lower())
        if w not in self.words:
            return "spell",self.rank(self.candidates(w),prev,nxt)[:5]
        if context and self.model.total()>=200:
            alts=self.candidates(w,edit=False)
            cur=self.model.b(prev,w)+self.model.b(w,nxt)
            if cur==0:
                scored=[(self.model.b(prev,c)+self.model.b(c,nxt),c) for c in alts]
                scored=[x for x in scored if x[0]>=2]
                if scored:
                    scored.sort(key=lambda x:(-x[0],-self.model.u(x[1])))
                    return "context",[c for _,c in scored[:5]]
        return None,[]
    def add_word(self,w):
        w=nfc(w.lower()); self.words.add(w); self.by_plain.setdefault(strip_tone(w),set()).add(w)
        DATA_DIR.mkdir(parents=True,exist_ok=True)
        with USER_WORDS_PATH.open("a",encoding="utf-8") as f:f.write(w+"\n")
