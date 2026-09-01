import os, re, shutil, urllib.parse, webbrowser, tkinter as tk
from tkinter import filedialog, messagebox
from pathlib import Path
from datetime import datetime

import customtkinter as ctk
from openpyxl import load_workbook, Workbook
import requests

try:
    from PIL import Image, ImageOps, ImageFilter, ImageEnhance
except Exception:
    Image = None

try:
    import pytesseract
except Exception:
    pytesseract = None

APP_DIR = Path(__file__).parent
ECU_DB = APP_DIR / "ECU_BAZA.xlsx"
ECU_IMG_DIR = APP_DIR / "ECU_SLIKE"
ECU_IMG_DIR.mkdir(exist_ok=True)

SKIP = {"ISKALNIK","NAVODILA","BAZA","POVZETEK","AI_WEB_CENE","ECU_BAZA"}

def clean(v):
    if v is None: return ""
    s=str(v).strip()
    return "" if s.lower() in ("none","nan","nat") else s

def norm(v):
    return clean(v).upper()

def brand(v):
    s=norm(v)
    return {
        "VOLKSWAGEN":"VW","WV":"VW",
        "MERCEDES":"MERCEDES-BENZ","MERCEDES BENZ":"MERCEDES-BENZ",
        "PEUGOT":"PEUGEOT","CITROËN":"CITROEN","ŠKODA":"SKODA"
    }.get(s,s)

def model(v):
    s=norm(v)
    return {
        "DASTER":"DUSTER","PASAT":"PASSAT","MEGAN":"MEGANE",
        "OCTAVIJA":"OCTAVIA","FOKUS":"FOCUS","TRANZIT":"TRANSIT",
        "KAŠKAI":"QASHQAI"
    }.get(s,s)

# ---------------- REZERVNI DELI ----------------

def find_header(rows):
    for i,r in enumerate(rows[:30]):
        t=" | ".join(norm(x) for x in r)
        if sum(k in t for k in ["MODEL","TIP","LETNIK","REZERVNI","CENA","ZALOGA"])>=2:
            return i
    for i,r in enumerate(rows[:30]):
        if sum(1 for x in r if clean(x))>=3:
            return i
    return None

def find(headers, keys, exclude=()):
    for i,h in enumerate(headers):
        hu=norm(h)
        if any(e in hu for e in exclude): continue
        if any(k in hu for k in keys): return i
    return None

def workbook_sheets(wb):
    if "VSI_PODATKI_SORTIRANO" in wb.sheetnames:
        return [wb["VSI_PODATKI_SORTIRANO"]]
    if "BAZA" in wb.sheetnames:
        return [wb["BAZA"]]
    return [w for w in wb.worksheets if norm(w.title) not in SKIP]

def read_parts(path, cb=None):
    wb=load_workbook(path, data_only=True, read_only=True)
    out=[]
    for ws in workbook_sheets(wb):
        if cb: cb("Berem: "+ws.title)
        rows=list(ws.iter_rows(values_only=True))
        if not rows: continue
        hi=find_header(rows)
        if hi is None: continue
        h=[clean(x) for x in rows[hi]]
        bi=find(h,["ZNAMKA","MODEL VOZILA"],["TIP"])
        mi=find(h,["TIP"],["MOTOR"])
        yi=find(h,["LETNIK"])
        ei=find(h,["TIP MOTORJA","MOTOR"])
        pi=find(h,["REZERVNI DEL"])
        ci=find(h,["CENA"])
        zi=find(h,["ZALOGA","STATUS","STANJE"])
        di=find(h,["DATUM"])
        if bi is None and len(h)>2: bi=2
        if mi is None and len(h)>3: mi=3
        if yi is None and len(h)>4: yi=4
        if ei is None and len(h)>5: ei=5
        if pi is None and len(h)>7: pi=7
        if ci is None and len(h)>8: ci=8
        for r in rows[hi+1:]:
            if not any(clean(x) for x in r): continue
            def get(i): return clean(r[i]) if i is not None and i<len(r) else ""
            rec={
                "zaloga":get(zi),
                "znamka":brand(get(bi)),
                "model":model(get(mi)),
                "letnik":get(yi),
                "motor":norm(get(ei)),
                "del":get(pi),
                "cena":get(ci),
                "datum":get(di) or ws.title
            }
            if rec["znamka"] or rec["model"] or rec["del"]:
                out.append(rec)
    wb.close()
    return out

def query_part(rec):
    return " ".join([
        rec.get("znamka",""), rec.get("model",""), rec.get("letnik",""),
        rec.get("motor",""), rec.get("del",""), "rabljeni rezervni del cena"
    ])

def open_part_web(rec):
    q=query_part(rec)
    webbrowser.open("https://www.google.com/search?q="+urllib.parse.quote_plus(q))
    webbrowser.open("https://www.ebay.com/sch/i.html?_nkw="+urllib.parse.quote_plus(q))
    webbrowser.open("https://www.google.com/search?q="+urllib.parse.quote_plus(q+" site:ovoko.com OR site:rrr.lt OR site:ebay.com"))

def parse_price(text):
    vals=[]
    for pat in [r"(\d{2,5}(?:[.,]\d{1,2})?)\s*(?:€|EUR|eur)", r"(?:€|EUR|eur)\s*(\d{2,5}(?:[.,]\d{1,2})?)"]:
        for m in re.findall(pat, text or ""):
            try:
                v=float(m.replace(".","").replace(",","."))
                if 5<=v<=10000: vals.append(v)
            except Exception:
                pass
    if not vals: return ""
    vals.sort()
    return str(round(vals[len(vals)//2],2)).replace(".",",")+" €"

def serpapi_price(rec, key):
    if not key: return ""
    q=query_part(rec)
    r=requests.get("https://serpapi.com/search.json", params={
        "engine":"google","q":q,"hl":"sl","gl":"si","api_key":key
    }, timeout=20)
    if r.status_code!=200:
        raise RuntimeError("SerpApi napaka "+str(r.status_code))
    data=r.json()
    snippets=[]
    for item in data.get("organic_results",[])[:10]:
        snippets.append(clean(item.get("title"))+" "+clean(item.get("snippet")))
    return parse_price(" ".join(snippets))

# ---------------- PAMETNI ECU OCR ----------------

def setup_tesseract():
    if pytesseract is None:
        return False, "pytesseract ni nameščen."
    possible=[
        r"C:\Program Files\Tesseract-OCR\tesseract.exe",
        r"C:\Program Files (x86)\Tesseract-OCR\tesseract.exe",
    ]
    for p in possible:
        if os.path.exists(p):
            pytesseract.pytesseract.tesseract_cmd=p
            return True,p
    try:
        pytesseract.get_tesseract_version()
        return True,"Tesseract v PATH"
    except Exception:
        return False,"Tesseract OCR ni najden."

def ensure_ecu_db():
    if ECU_DB.exists(): return
    wb=Workbook()
    ws=wb.active
    ws.title="ECU_BAZA"
    ws.append([
        "ID","DATUM","ZNAMKA","MODEL","LETNIK","MOTOR",
        "GLAVNA_KATALOSKA","VSE_OZNAKE","PROIZVAJALEC",
        "CENA","ZALOGA","OPOMBA","OCR_TEXT","SLIKA"
    ])
    wb.save(ECU_DB)

def next_ecu_id(ws):
    m=0
    for r in ws.iter_rows(min_row=2, values_only=True):
        try: m=max(m,int(r[0]))
        except Exception: pass
    return m+1

def read_ecu():
    ensure_ecu_db()
    wb=load_workbook(ECU_DB, data_only=True)
    ws=wb["ECU_BAZA"]
    out=[]
    for r in ws.iter_rows(min_row=2, values_only=True):
        if not any(clean(x) for x in r): continue
        out.append({
            "id":clean(r[0]),"datum":clean(r[1]),"znamka":clean(r[2]),"model":clean(r[3]),
            "letnik":clean(r[4]),"motor":clean(r[5]),"kat":clean(r[6]),"vse":clean(r[7]),
            "proizvajalec":clean(r[8]),"cena":clean(r[9]),"zaloga":clean(r[10]),
            "opomba":clean(r[11]),"ocr":clean(r[12]),"slika":clean(r[13])
        })
    wb.close()
    return out

def normalize_ocr_text(t):
    t = norm(t)
    # OCR mistakes common on ECU labels
    t = t.replace("O", "0")
    t = t.replace("I", "1")
    t = t.replace("|", "1")
    # keep spaces for VAG patterns
    t = re.sub(r"[^A-Z0-9\s\-/\.]", " ", t)
    t = re.sub(r"\s+", " ", t)
    return t

def detect_producer(text):
    t=norm(text)
    producers = ["BOSCH","MAGNETI MARELLI","MARELLI","SIEMENS","DELPHI","DENSO","CONTINENTAL","VALEO","VISTEON","TRW","IAW"]
    for p in producers:
        if p in t:
            if p == "MARELLI": return "MAGNETI MARELLI"
            return p
    return ""

def detect_brand_from_text(text):
    t = norm(text)
    if "VW" in t or "VOLKSWAGEN" in t or re.search(r"\b[0-9][A-Z0]\d\s*9\d{2}\s*\d{3}\s*[A-Z]{0,3}\b", normalize_ocr_text(text)):
        return "VW"
    if "AUDI" in t: return "AUDI"
    if "SEAT" in t: return "SEAT"
    if "SKODA" in t or "ŠKODA" in t: return "SKODA"
    return ""

def format_vag_number(s):
    s = re.sub(r"[^A-Z0-9]", "", norm(s).replace("O","0"))
    # e.g. 6Q0906034AN -> 6Q0 906 034 AN
    if len(s) >= 9 and re.match(r"^[0-9][A-Z0-9][0-9][0-9]{6}[A-Z0-9]{0,3}$", s):
        return f"{s[:3]} {s[3:6]} {s[6:9]} {s[9:] if len(s)>9 else ''}".strip()
    return s

def extract_catalog_numbers(text, filename=""):
    raw = (text or "") + " " + (filename or "")
    t = normalize_ocr_text(raw)
    compact = re.sub(r"[\s\-/\.]", "", t)

    candidates = set()

    # VAG OEM, e.g. 6Q0 906 034 AN, 03G 906 016, 038 906 019
    vag_patterns = [
        r"\b([0-9][A-Z0-9][0-9]\s*9[0-9]{2}\s*[0-9]{3}\s*[A-Z0-9]{0,3})\b",
        r"\b([0-9]{3}\s*9[0-9]{2}\s*[0-9]{3}\s*[A-Z0-9]{0,3})\b",
        r"\b([0-9][A-Z0-9][0-9]9[0-9]{2}[0-9]{3}[A-Z0-9]{0,3})\b",
        r"\b([0-9]{3}9[0-9]{2}[0-9]{3}[A-Z0-9]{0,3})\b",
    ]
    for pat in vag_patterns:
        for m in re.findall(pat, t):
            candidates.add(format_vag_number(m))

    # Also from compact text
    for m in re.findall(r"\b([0-9][A-Z0-9][0-9]9[0-9]{2}[0-9]{3}[A-Z0-9]{0,3})\b", compact):
        candidates.add(format_vag_number(m))
    for m in re.findall(r"\b([0-9]{3}9[0-9]{2}[0-9]{3}[A-Z0-9]{0,3})\b", compact):
        candidates.add(format_vag_number(m))

    # Bosch / common ECU numbers
    common_patterns = [
        r"\b0\s*281\s*\d{3}\s*\d{3}\b",
        r"\b0281\d{6}\b",
        r"\b0\s*261\s*\d{3}\s*\d{3}\b",
        r"\b0261\d{6}\b",
        r"\bIAW\s*[A-Z0-9\.\-]{4,12}\b",
        r"\bMJD\s*[A-Z0-9\.\-]{4,12}\b",
        r"\bMJT\s*[A-Z0-9\.\-]{4,12}\b",
        r"\b\d{7,12}[A-Z]{0,3}\b",
        r"\b[A-Z]{1,4}\d{5,12}[A-Z0-9]{0,4}\b",
    ]
    for pat in common_patterns:
        for m in re.findall(pat, t):
            val = re.sub(r"[\s\-\/\.]", "", m)
            if 6 <= len(val) <= 18:
                candidates.add(val)

    # Filter obvious non-catalog noise
    bad_exact = {"000000","111111","123456","999999","160300","2893","076","00","01","45"}
    filtered = []
    for c in candidates:
        cc = re.sub(r"[\s]", "", c)
        if cc in bad_exact: 
            continue
        if len(cc) < 6:
            continue
        filtered.append(c)

    def score(x):
        xs = re.sub(r"\s","",x)
        if re.match(r"^[0-9][A-Z0-9][0-9]9[0-9]{2}[0-9]{3}[A-Z0-9]{0,3}$", xs): return 0
        if re.match(r"^[0-9]{3}9[0-9]{2}[0-9]{3}[A-Z0-9]{0,3}$", xs): return 1
        if xs.startswith(("0281","0261")): return 2
        if xs.startswith(("IAW","MJD","MJT")): return 3
        return 5

    filtered = sorted(set(filtered), key=lambda x:(score(x), len(x), x))
    return filtered

def preprocess_versions(path):
    img=Image.open(path)
    img=ImageOps.exif_transpose(img)
    versions=[]

    base=img.convert("L")
    w,h=base.size
    if max(w,h)<2500:
        base=base.resize((w*2,h*2))
    base=ImageOps.autocontrast(base)
    versions.append(base.filter(ImageFilter.SHARPEN))

    # high contrast
    hc=ImageEnhance.Contrast(base).enhance(2.2)
    versions.append(hc.filter(ImageFilter.SHARPEN))

    # inverted sometimes helps for labels
    inv=ImageOps.invert(hc)
    versions.append(inv)

    # rotate 90/270 if user photographed sideways
    versions.append(base.rotate(90, expand=True))
    versions.append(base.rotate(270, expand=True))
    return versions

def ocr_image(path):
    ok,msg=setup_tesseract()
    if not ok or Image is None:
        nums=extract_catalog_numbers("",Path(path).stem)
        return "",nums,msg

    all_text=[]
    all_nums=set()
    try:
        for img in preprocess_versions(path):
            for psm in [6, 11, 12]:
                config=f"--psm {psm} -c tessedit_char_whitelist=ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789-/._ "
                text=pytesseract.image_to_string(img, lang="eng", config=config)
                all_text.append(text)
                for n in extract_catalog_numbers(text,Path(path).stem):
                    all_nums.add(n)
        joined="\n".join(all_text)
        nums=extract_catalog_numbers(joined,Path(path).stem)
        for n in all_nums:
            if n not in nums: nums.append(n)
        return joined,nums,"OK"
    except Exception as e:
        nums=extract_catalog_numbers("",Path(path).stem)
        return "",nums,"OCR napaka: "+str(e)

def add_ecu(data, image_path="", ocr_text=""):
    ensure_ecu_db()
    wb=load_workbook(ECU_DB)
    ws=wb["ECU_BAZA"]
    eid=next_ecu_id(ws)
    saved=""
    if image_path:
        src=Path(image_path)
        if src.exists():
            safe=re.sub(r"[^A-Za-z0-9_-]+","_",data.get("kat") or f"ECU_{eid}")
            dst=ECU_IMG_DIR / f"{eid}_{safe}{src.suffix.lower()}"
            shutil.copy2(src,dst)
            saved=str(dst)
    ws.append([
        eid,datetime.now().strftime("%Y-%m-%d %H:%M"),
        data.get("znamka",""),data.get("model",""),data.get("letnik",""),data.get("motor",""),
        data.get("kat",""),data.get("vse",""),data.get("proizvajalec",""),data.get("cena",""),
        data.get("zaloga","NA ZALOGI"),data.get("opomba",""),ocr_text,saved
    ])
    wb.save(ECU_DB)
    return eid

def import_ecu_folder(folder, log=None):
    imgs=[]
    for ext in ("*.jpg","*.jpeg","*.png","*.webp","*.bmp"):
        imgs += list(Path(folder).glob(ext))
    found=0; manual=0
    for img in imgs:
        if log: log("OCR: "+img.name)
        text,nums,status=ocr_image(str(img))
        prod=detect_producer(text+" "+img.stem)
        detected_brand=detect_brand_from_text(text+" "+img.stem)
        if nums:
            # one record per image; main number first, all numbers stored
            add_ecu({
                "znamka":detected_brand,
                "kat":nums[0],
                "vse":" | ".join(nums),
                "proizvajalec":prod,
                "zaloga":"NA ZALOGI",
                "opomba":"Pametni OCR uvoz"
            },str(img),text)
            found+=1
        else:
            add_ecu({
                "kat":"",
                "vse":"",
                "proizvajalec":prod,
                "zaloga":"PREVERI",
                "opomba":"OCR ni našel kataloške - ročni pregled"
            },str(img),text)
            manual+=1
    return found,manual,len(imgs)

# ---------------- UI ----------------

class App(ctk.CTk):
    def __init__(self):
        super().__init__()
        ctk.set_appearance_mode("dark"); ctk.set_default_color_theme("blue")
        self.title("GLOBAL AVTO | ZONTA KOMPLET - pametni ECU OCR")
        self.geometry("1400x850")
        self.file=tk.StringVar()
        self.api=tk.StringVar()
        self.parts=[]; self.parts_results=[]; self.ecu=[]; self.ecu_results=[]
        self.qb=tk.StringVar(); self.qm=tk.StringVar(); self.qy=tk.StringVar(); self.qe=tk.StringVar(); self.qp=tk.StringVar()
        self.eb=tk.StringVar(); self.em=tk.StringVar(); self.ey=tk.StringVar(); self.ee=tk.StringVar(); self.ek=tk.StringVar()
        self.build()

    def build(self):
        self.grid_columnconfigure(0,weight=1); self.grid_rowconfigure(1,weight=1)
        ctk.CTkLabel(self,text="GLOBAL AVTO - iskanje delov + pametni ECU OCR",font=("Segoe UI",28,"bold")).grid(row=0,column=0,padx=20,pady=15,sticky="w")
        self.tabs=ctk.CTkTabview(self)
        self.tabs.grid(row=1,column=0,sticky="nsew",padx=20,pady=(0,20))
        self.t_parts=self.tabs.add("Iskanje rezervnih delov")
        self.t_ecu=self.tabs.add("Motorni računalniki ECU")
        self.build_parts_tab()
        self.build_ecu_tab()

    def make_tree(self,parent,cols,row):
        frame=ctk.CTkFrame(parent,fg_color="#0d141c")
        frame.grid(row=row,column=0,sticky="nsew",padx=10,pady=10)
        frame.grid_columnconfigure(0,weight=1); frame.grid_rowconfigure(0,weight=1)
        tr=tk.ttk.Treeview(frame,columns=cols,show="headings")
        st=tk.ttk.Style()
        try: st.theme_use("clam")
        except Exception: pass
        st.configure("Treeview",background="#0d141c",foreground="#e8eef5",fieldbackground="#0d141c",rowheight=32)
        st.configure("Treeview.Heading",background="#1d2b38",foreground="white")
        for c in cols:
            tr.heading(c,text=c.upper())
            w=330 if c in ("del","slika","opomba","vse") else 115
            if c in ("kat","kataloska"): w=190
            tr.column(c,width=w,anchor="w")
        tr.grid(row=0,column=0,sticky="nsew")
        y=tk.ttk.Scrollbar(frame,orient="vertical",command=tr.yview)
        tr.configure(yscrollcommand=y.set)
        y.grid(row=0,column=1,sticky="ns")
        return tr

    def build_parts_tab(self):
        t=self.t_parts
        t.grid_columnconfigure(0,weight=1); t.grid_rowconfigure(4,weight=1)
        top=ctk.CTkFrame(t); top.grid(row=0,column=0,sticky="ew",padx=10,pady=8); top.grid_columnconfigure(0,weight=1)
        ctk.CTkEntry(top,textvariable=self.file,placeholder_text="Glavna Excel evidenca",height=38).grid(row=0,column=0,sticky="ew",padx=8,pady=8)
        ctk.CTkButton(top,text="Izberi evidenco",command=self.pick_file).grid(row=0,column=1,padx=6)
        ctk.CTkButton(top,text="Naloži",fg_color="#19a75f",command=self.load_parts).grid(row=0,column=2,padx=6)
        ctk.CTkEntry(top,textvariable=self.api,placeholder_text="SerpApi ključ za spletno ceno (neobvezno)",height=38).grid(row=1,column=0,columnspan=3,sticky="ew",padx=8,pady=(0,8))

        s=ctk.CTkFrame(t); s.grid(row=1,column=0,sticky="ew",padx=10,pady=8)
        fields=[("Znamka",self.qb),("Model",self.qm),("Letnik",self.qy),("Motor",self.qe),("Rezervni del",self.qp)]
        for i,(lab,var) in enumerate(fields):
            f=ctk.CTkFrame(s,fg_color="transparent")
            f.grid(row=0,column=i,padx=5,sticky="ew")
            s.grid_columnconfigure(i,weight=1)
            ctk.CTkLabel(f,text=lab,text_color="#9aa8b8").pack(anchor="w")
            ctk.CTkEntry(f,textvariable=var,fg_color="#cfe2f3",text_color="#0b1117").pack(fill="x")
        ctk.CTkButton(s,text="IŠČI",command=self.search_parts).grid(row=0,column=5,padx=8,pady=(20,0))
        ctk.CTkButton(s,text="SPLETNA CENA",fg_color="#19a75f",command=self.web_part).grid(row=0,column=6,padx=8,pady=(20,0))

        self.parts_info=ctk.CTkLabel(t,text="Naloži evidenco.",text_color="#9aa8b8")
        self.parts_info.grid(row=2,column=0,sticky="w",padx=12)
        self.price_info=ctk.CTkLabel(t,text="",text_color="#69d28f",font=("Segoe UI",14,"bold"))
        self.price_info.grid(row=3,column=0,sticky="w",padx=12)
        self.parts_tree=self.make_tree(t,("zaloga","znamka","model","letnik","motor","del","cena","datum"),4)

    def build_ecu_tab(self):
        t=self.t_ecu
        t.grid_columnconfigure(0,weight=1); t.grid_rowconfigure(4,weight=1)
        a=ctk.CTkFrame(t); a.grid(row=0,column=0,sticky="ew",padx=10,pady=8)
        ctk.CTkButton(a,text="UVOZI MAPO SLIK + PAMETNI OCR",fg_color="#19a75f",command=self.import_ecu_images).grid(row=0,column=0,padx=8,pady=8)
        ctk.CTkButton(a,text="DODAJ ENO SLIKO",command=self.add_one_ecu).grid(row=0,column=1,padx=8)
        ctk.CTkButton(a,text="ODPRI SLIKO",command=self.open_ecu_image).grid(row=0,column=2,padx=8)
        ctk.CTkButton(a,text="OSVEŽI ECU",command=self.load_ecu).grid(row=0,column=3,padx=8)

        s=ctk.CTkFrame(t); s.grid(row=1,column=0,sticky="ew",padx=10,pady=8)
        fields=[("Znamka",self.eb),("Model",self.em),("Letnik",self.ey),("Motor",self.ee),("Kataloška / katerakoli oznaka",self.ek)]
        for i,(lab,var) in enumerate(fields):
            f=ctk.CTkFrame(s,fg_color="transparent")
            f.grid(row=0,column=i,padx=5,sticky="ew")
            s.grid_columnconfigure(i,weight=1)
            ctk.CTkLabel(f,text=lab,text_color="#9aa8b8").pack(anchor="w")
            ctk.CTkEntry(f,textvariable=var,fg_color="#cfe2f3",text_color="#0b1117").pack(fill="x")
        ctk.CTkButton(s,text="IŠČI ECU",command=self.search_ecu).grid(row=0,column=5,padx=8,pady=(20,0))

        self.ecu_info=ctk.CTkLabel(t,text=f"ECU baza: {ECU_DB}",text_color="#9aa8b8")
        self.ecu_info.grid(row=2,column=0,sticky="w",padx=12)
        self.logbox=ctk.CTkTextbox(t,height=90)
        self.logbox.grid(row=3,column=0,sticky="ew",padx=10,pady=5)
        self.ecu_tree=self.make_tree(t,("id","zaloga","kat","vse","proizvajalec","znamka","model","letnik","motor","cena","opomba","slika"),4)
        self.load_ecu()

    def log(self,msg):
        self.logbox.insert("end",msg+"\n")
        self.logbox.see("end")
        self.update()

    def pick_file(self):
        p=filedialog.askopenfilename(filetypes=[("Excel","*.xlsx")])
        if p: self.file.set(p)

    def load_parts(self):
        p=self.file.get()
        if not p or not os.path.exists(p):
            self.pick_file(); p=self.file.get()
        if not p: return
        self.parts_info.configure(text="Nalagam ..."); self.update()
        try:
            self.parts=read_parts(p,lambda msg:self.parts_info.configure(text=msg))
            self.parts_info.configure(text=f"Naloženo: {len(self.parts)} zapisov")
            self.show_parts(self.parts[:500])
        except Exception as e:
            messagebox.showerror("Napaka",str(e))

    def show_parts(self,rows):
        self.parts_results=rows
        for i in self.parts_tree.get_children(): self.parts_tree.delete(i)
        for idx,r in enumerate(rows):
            self.parts_tree.insert("", "end", iid=str(idx), values=[r["zaloga"],r["znamka"],r["model"],r["letnik"],r["motor"],r["del"],r["cena"],r["datum"]])

    def search_parts(self):
        if not self.parts: self.load_parts()
        qb,qm,qy,qe,qp=brand(self.qb.get()),model(self.qm.get()),norm(self.qy.get()),norm(self.qe.get()),norm(self.qp.get())
        res=[]
        for r in self.parts:
            if qb and qb not in norm(r["znamka"]): continue
            if qm and qm not in norm(r["model"]): continue
            if qy and qy not in norm(r["letnik"]): continue
            if qe and qe not in norm(r["motor"]): continue
            if qp and qp not in norm(r["del"]): continue
            res.append(r)
        res.sort(key=lambda r:(1 if "NI" in norm(r.get("zaloga","")) or "PRODANO" in norm(r.get("zaloga","")) else 0,norm(r["znamka"]),norm(r["model"]),norm(r["del"])))
        self.show_parts(res[:1000])
        self.parts_info.configure(text=f"Najdenih: {len(res)} | prikazanih: {min(len(res),1000)}")

    def selected_part(self):
        sel=self.parts_tree.selection()
        if sel:
            try: return self.parts_results[int(sel[0])]
            except Exception: pass
        return {"znamka":brand(self.qb.get()),"model":model(self.qm.get()),"letnik":self.qy.get(),"motor":norm(self.qe.get()),"del":self.qp.get()}

    def web_part(self):
        r=self.selected_part()
        key=self.api.get().strip()
        if key:
            try:
                self.price_info.configure(text="Iščem spletno ceno ..."); self.update()
                p=serpapi_price(r,key)
                if p:
                    self.price_info.configure(text="Spletna ocena: "+p)
                    return
            except Exception:
                self.price_info.configure(text="API ni uspel, odpiram povezave.")
        else:
            self.price_info.configure(text="Odpiram spletne povezave.")
        open_part_web(r)

    def load_ecu(self):
        self.ecu=read_ecu()
        self.show_ecu(self.ecu[:1000])
        self.ecu_info.configure(text=f"ECU zapisov: {len(self.ecu)} | baza: {ECU_DB}")

    def show_ecu(self,rows):
        self.ecu_results=rows
        for i in self.ecu_tree.get_children(): self.ecu_tree.delete(i)
        for idx,r in enumerate(rows):
            self.ecu_tree.insert("", "end", iid=str(idx), values=[
                r["id"],r["zaloga"],r["kat"],r["vse"],r["proizvajalec"],r["znamka"],r["model"],
                r["letnik"],r["motor"],r["cena"],r["opomba"],r["slika"]
            ])

    def search_ecu(self):
        qk,qb,qm,qy,qe=norm(self.ek.get()),brand(self.eb.get()),model(self.em.get()),norm(self.ey.get()),norm(self.ee.get())
        res=[]
        for r in self.ecu:
            blob=norm(r["kat"]+" "+r.get("vse","")+" "+r.get("ocr",""))
            if qk and qk not in blob: continue
            if qb and qb not in norm(r["znamka"]): continue
            if qm and qm not in norm(r["model"]): continue
            if qy and qy not in norm(r["letnik"]): continue
            if qe and qe not in norm(r["motor"]): continue
            res.append(r)
        res.sort(key=lambda r:(norm(r["kat"]),norm(r["znamka"]),norm(r["model"])))
        self.show_ecu(res[:1000])
        self.ecu_info.configure(text=f"Najdenih ECU: {len(res)}")

    def import_ecu_images(self):
        folder=filedialog.askdirectory(title="Izberi mapo s slikami ECU")
        if not folder: return
        ok,msg=setup_tesseract()
        if not ok:
            messagebox.showwarning("Tesseract manjka", msg+"\n\nProgram bo poskusil brati številke iz imen datotek.")
        found,manual,total=import_ecu_folder(folder,self.log)
        messagebox.showinfo("Uvoz končan",f"Slike: {total}\nNajdene kataloške: {found}\nZa ročni pregled: {manual}")
        self.load_ecu()

    def add_one_ecu(self):
        img=filedialog.askopenfilename(filetypes=[("Slike","*.jpg *.jpeg *.png *.webp *.bmp")])
        if not img: return
        text,nums,status=ocr_image(img)
        kat=nums[0] if nums else self.ek.get()
        prod=detect_producer(text+" "+Path(img).stem)
        detected_brand=detect_brand_from_text(text+" "+Path(img).stem)
        eid=add_ecu({
            "znamka":detected_brand,
            "kat":kat,
            "vse":" | ".join(nums),
            "proizvajalec":prod,
            "zaloga":"NA ZALOGI" if kat else "PREVERI",
            "opomba":"Ročni/pametni OCR vnos"
        },img,text)
        messagebox.showinfo("Dodano",f"ECU dodan. ID: {eid}\nGlavna kataloška: {kat or 'ni prebrano'}\nVse oznake: {' | '.join(nums)}")
        self.load_ecu()

    def open_ecu_image(self):
        sel=self.ecu_tree.selection()
        if not sel:
            messagebox.showinfo("Ni izbire","Izberi ECU v tabeli.")
            return
        r=self.ecu_results[int(sel[0])]
        p=r.get("slika","")
        if p and os.path.exists(p):
            os.startfile(p)
        else:
            messagebox.showerror("Napaka","Slika ni najdena.")

if __name__=="__main__":
    App().mainloop()
