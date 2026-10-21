# -*- coding: utf-8 -*-
"""Editörler için pencere arayüzü (Tkinter).

Komut satırı kullanmadan: Word dosyalarını seç, makale bilgilerini gir, "Dönüştür"e bas;
PDF, kontrol raporu, zip ve Crossref/JATS/BibTeX dosyaları üretilir. Bütün bir sayı da
aynı pencereden işlenebilir. Son girilen bilgiler bir sonraki açılışta hatırlanır.
"""
import datetime
import json
import os
import queue
import re
import subprocess
import sys
import threading

from . import __version__
from .template import MONTHS

SETTINGS_PATH = os.path.join(os.path.expanduser("~"), ".sjuoz_gui.json")
ARTICLE_TYPES = ["Original Article", "Review Article", "Short Communication", "Case Report"]
FIELDS = [  # (anahtar, etiket, varsayılan)
    ("article_id", "Makale no", ""),
    ("volume", "Cilt", ""),
    ("issue", "Sayı", ""),
    ("year", "Yıl", str(datetime.date.today().year)),
    ("first_page", "İlk sayfa", "1"),
    ("received", "Geliş tarihi", "XX"),
    ("accepted", "Kabul tarihi", "XX"),
    ("published", "Yayın tarihi", "XX"),
    ("depositor_email", "Crossref e-postası", ""),
]


def build_args(values, mode="makale"):
    """Formdaki değerlerden komut satırı seçenekleri üretir (ekran gerektirmez, test edilebilir).

    values: {"manuscript", "cover", "issue_folder", "outdir", "article_id", "volume", "issue", "year",
             "month", "first_page", "received", "accepted", "published", "article_type",
             "depositor_email", "compile", "zip", "export", "links", "link_blue"}
    """
    v = {k: (val.strip() if isinstance(val, str) else val) for k, val in values.items()}
    if mode == "sayi":
        if not v.get("issue_folder"):
            raise ValueError("Sayı klasörünü seçin.")
        args = ["--sayi", v["issue_folder"]]
    else:
        if not v.get("manuscript"):
            raise ValueError("Makale Word dosyasını seçin.")
        args = [v["manuscript"]] + ([v["cover"]] if v.get("cover") else [])
        for key, flag in (("article_id", "--article-id"), ("first_page", "--first-page"),
                          ("received", "--received"), ("accepted", "--accepted"),
                          ("published", "--published"), ("article_type", "--article-type")):
            if v.get(key):
                args += [flag, v[key]]
    for key, flag in (("volume", "--volume"), ("issue", "--issue"), ("year", "--year")):
        if v.get(key):
            if not re.fullmatch(r"\d+", v[key]):
                raise ValueError("%s sayı olmalı: %r" % ({"volume": "Cilt", "issue": "Sayı", "year": "Yıl"}[key], v[key]))
            args += [flag, v[key]]
    if mode != "sayi" and v.get("first_page") and not re.fullmatch(r"\d+", v["first_page"]):
        raise ValueError("İlk sayfa sayı olmalı: %r" % v["first_page"])
    if v.get("month"):
        args += ["--month", v["month"]]
    if v.get("outdir"):
        args += ["-o", v["outdir"]]
    if v.get("depositor_email"):
        args += ["--depositor-email", v["depositor_email"]]
    if not v.get("compile", True):
        args.append("--no-compile")
    if not v.get("zip", True):
        args.append("--no-zip")
    if not v.get("export", True):
        args.append("--no-export")
    if not v.get("links", True):
        args.append("--no-links")
    elif v.get("link_blue"):
        args += ["--link-color", "blue"]
    return args


def guess_article_id(path):
    mt = re.search(r"(\d{3,})", os.path.basename(path or ""))
    return mt.group(1) if mt else ""


def open_path(path):
    """Dosyayı/klasörü işletim sisteminin varsayılan programıyla açar."""
    if not path or not os.path.exists(path):
        return
    if sys.platform.startswith("win"):
        os.startfile(path)  # noqa  (yalnız Windows)
    elif sys.platform == "darwin":
        subprocess.Popen(["open", path])
    else:
        subprocess.Popen(["xdg-open", path])


class _Stream:
    """print() çıktısını pencereye aktarır."""

    def __init__(self, q):
        self.q = q

    def write(self, text):
        if text:
            self.q.put(("log", text))

    def flush(self):
        pass


class App:
    def __init__(self, root):
        import tkinter as tk
        from tkinter import ttk
        self.tk, self.ttk, self.root = tk, ttk, root
        root.title("SJUOZ Word → LaTeX  %s" % __version__)
        root.minsize(760, 640)
        self.q = queue.Queue()
        self.result = None
        self.vars = {}
        settings = self.load_settings()

        main = ttk.Frame(root, padding=12)
        main.pack(fill="both", expand=True)
        self.tabs = ttk.Notebook(main)
        self.tabs.pack(fill="x")

        # --- Tek makale
        one = ttk.Frame(self.tabs, padding=8)
        self.tabs.add(one, text="Tek makale")
        self.vars["manuscript"] = tk.StringVar()
        self.vars["cover"] = tk.StringVar()
        self._file_row(one, 0, "Makale (.docx / .pdf / .tex)", "manuscript", self.pick_manuscript)
        self._file_row(one, 1, "Kapak mektubu (.docx / .pdf)", "cover",
                       lambda: self.pick_file("cover", "Kapak mektubu"))
        ttk.Label(one, text="Kapak mektubu isteğe bağlıdır; ORCID ve e-postalar oradan alınır.",
                  foreground="#555").grid(row=2, column=1, columnspan=2, sticky="w")

        # --- Bütün sayı
        many = ttk.Frame(self.tabs, padding=8)
        self.tabs.add(many, text="Bütün sayı")
        self.vars["issue_folder"] = tk.StringVar(value=settings.get("issue_folder", ""))
        self._file_row(many, 0, "Sayı klasörü", "issue_folder", self.pick_issue_folder)
        ttk.Label(many, text="Her alt klasör bir makaledir (klasör adı = makale no). İsteğe bağlı sayi.json\n"
                             "makale sırasını ve tarihleri belirler; sayfa numaraları zincirlenir.",
                  foreground="#555").grid(row=1, column=1, columnspan=2, sticky="w")

        # --- Bilgiler
        info = ttk.LabelFrame(main, text="Makale / sayı bilgileri", padding=8)
        info.pack(fill="x", pady=(10, 0))
        for k, (key, label, default) in enumerate(FIELDS):
            r, c = divmod(k, 3)
            ttk.Label(info, text=label).grid(row=r, column=c * 2, sticky="w", padx=(0, 4), pady=2)
            var = tk.StringVar(value=settings.get(key, default) if key not in ("article_id", "received",
                                                                                 "accepted", "published")
                               else default)
            self.vars[key] = var
            ttk.Entry(info, textvariable=var, width=18).grid(row=r, column=c * 2 + 1, sticky="w", padx=(0, 12))
        r = len(FIELDS) // 3 + 1
        ttk.Label(info, text="Ay").grid(row=r, column=0, sticky="w")
        self.vars["month"] = tk.StringVar(value=settings.get("month", MONTHS[datetime.date.today().month - 1]))
        ttk.Combobox(info, textvariable=self.vars["month"], values=MONTHS, width=15,
                     state="readonly").grid(row=r, column=1, sticky="w")
        ttk.Label(info, text="Makale türü").grid(row=r, column=2, sticky="w")
        self.vars["article_type"] = tk.StringVar(value="")
        ttk.Combobox(info, textvariable=self.vars["article_type"], values=[""] + ARTICLE_TYPES,
                     width=18).grid(row=r, column=3, sticky="w")
        ttk.Label(info, text="Tarih örneği: 12 May 2026", foreground="#555").grid(row=r, column=4, columnspan=2,
                                                                                  sticky="w")

        # --- Çıktı ve seçenekler
        out = ttk.LabelFrame(main, text="Çıktı", padding=8)
        out.pack(fill="x", pady=(10, 0))
        self.vars["outdir"] = tk.StringVar(value=settings.get("outdir", ""))
        self._file_row(out, 0, "Çıktı klasörü", "outdir", self.pick_outdir)
        opts = ttk.Frame(out)
        opts.grid(row=1, column=1, columnspan=2, sticky="w", pady=(4, 0))
        for k, (key, label) in enumerate((("compile", "PDF derle"), ("zip", "Overleaf zip"),
                                          ("export", "Crossref / JATS / BibTeX"), ("links", "Tıklanabilir bağlantılar"),
                                          ("link_blue", "Bağlantılar mavi"))):
            self.vars[key] = tk.BooleanVar(value=settings.get(key, key != "link_blue"))
            ttk.Checkbutton(opts, text=label, variable=self.vars[key]).grid(row=0, column=k, padx=(0, 14))

        # --- Düğmeler
        bar = ttk.Frame(main)
        bar.pack(fill="x", pady=10)
        self.run_btn = ttk.Button(bar, text="Dönüştür", command=self.start)
        self.run_btn.pack(side="left")
        self.progress = ttk.Progressbar(bar, mode="indeterminate", length=160)
        self.progress.pack(side="left", padx=10)
        self.status = ttk.Label(bar, text="Hazır", font=("TkDefaultFont", 10, "bold"))
        self.status.pack(side="left", padx=6)
        self.open_btns = {}
        for key, label in (("folder", "Klasörü aç"), ("pdf", "PDF'i aç"), ("report", "Raporu aç")):
            b = ttk.Button(bar, text=label, state="disabled", command=lambda k=key: self.open_result(k))
            b.pack(side="right", padx=(6, 0))
            self.open_btns[key] = b

        # --- Günlük
        logf = ttk.LabelFrame(main, text="İşlem günlüğü", padding=4)
        logf.pack(fill="both", expand=True)
        self.log = tk.Text(logf, height=14, wrap="word", font=("TkFixedFont", 9))
        sb = ttk.Scrollbar(logf, command=self.log.yview)
        self.log.configure(yscrollcommand=sb.set)
        self.log.pack(side="left", fill="both", expand=True)
        sb.pack(side="right", fill="y")
        self.log.tag_configure("warn", foreground="#b54708")
        self.log.tag_configure("err", foreground="#b42318")
        self.log.tag_configure("ok", foreground="#067647")
        root.after(100, self.poll)

    # ---- yardımcılar
    def _file_row(self, parent, row, label, key, command):
        ttk = self.ttk
        ttk.Label(parent, text=label).grid(row=row, column=0, sticky="w", pady=3)
        ttk.Entry(parent, textvariable=self.vars[key], width=70).grid(row=row, column=1, sticky="we", padx=6)
        ttk.Button(parent, text="Seç…", command=command).grid(row=row, column=2)
        parent.columnconfigure(1, weight=1)

    def pick_file(self, key, title):
        from tkinter import filedialog
        path = filedialog.askopenfilename(title=title, filetypes=[("Word, PDF, LaTeX", "*.docx *.pdf *.tex *.zip"), ("Word", "*.docx"), ("PDF", "*.pdf"),
                                                     ("LaTeX (.tex / .zip)", "*.tex *.zip")])
        if path:
            self.vars[key].set(path)
        return path

    def pick_manuscript(self):
        path = self.pick_file("manuscript", "Makale")
        if path and not self.vars["article_id"].get():
            self.vars["article_id"].set(guess_article_id(path))
        if path and not self.vars["outdir"].get():
            self.vars["outdir"].set(os.path.join(os.path.dirname(path), "latex_" + (guess_article_id(path) or "cikti")))

    def pick_issue_folder(self):
        from tkinter import filedialog
        path = filedialog.askdirectory(title="Sayı klasörü")
        if path:
            self.vars["issue_folder"].set(path)

    def pick_outdir(self):
        from tkinter import filedialog
        path = filedialog.askdirectory(title="Çıktı klasörü")
        if path:
            self.vars["outdir"].set(path)

    def values(self):
        return {k: v.get() for k, v in self.vars.items()}

    def mode(self):
        return "sayi" if self.tabs.index(self.tabs.select()) == 1 else "makale"

    def load_settings(self):
        try:
            with open(SETTINGS_PATH, encoding="utf-8") as f:
                return json.load(f)
        except (OSError, ValueError):
            return {}

    def save_settings(self):
        keep = ("volume", "issue", "year", "month", "depositor_email", "outdir", "issue_folder",
                "compile", "zip", "export", "links", "link_blue")
        try:
            with open(SETTINGS_PATH, "w", encoding="utf-8") as f:
                json.dump({k: v for k, v in self.values().items() if k in keep}, f, ensure_ascii=False, indent=1)
        except OSError:
            pass

    # ---- çalıştırma
    def start(self):
        from tkinter import messagebox
        try:
            args = build_args(self.values(), self.mode())
        except ValueError as e:
            messagebox.showwarning("Eksik bilgi", str(e))
            return
        self.save_settings()
        self.log.delete("1.0", "end")
        self.run_btn.configure(state="disabled")
        for b in self.open_btns.values():
            b.configure(state="disabled")
        self.progress.start(12)
        self.status.configure(text="Çalışıyor…", foreground="#175cd3")
        threading.Thread(target=self.work, args=(args, self.mode()), daemon=True).start()

    def work(self, args, mode):
        from .issue import convert_issue
        from .pipeline import build_parser, convert
        old = sys.stdout, sys.stderr
        sys.stdout = sys.stderr = _Stream(self.q)
        try:
            opts = build_parser().parse_args(args)
            res = convert_issue(opts.sayi, opts) if mode == "sayi" else convert(opts)
            self.q.put(("done", (mode, res)))
        except SystemExit as e:
            self.q.put(("error", "Geçersiz seçenek: %s" % e))
        except Exception as e:  # noqa  — arayüz çökmemeli, hata günlüğe yazılır
            import traceback
            self.q.put(("error", "%s\n%s" % (e, traceback.format_exc())))
        finally:
            sys.stdout, sys.stderr = old

    def poll(self):
        try:
            while True:
                kind, payload = self.q.get_nowait()
                if kind == "log":
                    tag = "err" if "HATA" in payload else "warn" if ("UYARI" in payload or "KONTROL" in payload) \
                        else "ok" if ("PDF üretildi" in payload or "TAMAM" in payload) else None
                    self.log.insert("end", payload, tag)
                    self.log.see("end")
                elif kind == "done":
                    self.finish(*payload)
                elif kind == "error":
                    self.finish(None, None, payload)
        except queue.Empty:
            pass
        self.root.after(100, self.poll)

    def finish(self, mode, res, error=None):
        self.progress.stop()
        self.run_btn.configure(state="normal")
        colors = {"TAMAM": "#067647", "KONTROL": "#b54708", "HATA": "#b42318"}
        if error or not res:
            self.status.configure(text="Hata", foreground=colors["HATA"])
            self.log.insert("end", "\n" + (error or "Makale dosyası bulunamadı.") + "\n", "err")
            return
        self.result = (mode, res)
        self.status.configure(text="Bitti — %s" % res["status"], foreground=colors.get(res["status"], "black"))
        self.open_btns["folder"].configure(state="normal")
        has_pdf = bool(res.get("full_pdf") if mode == "sayi" else res.get("pdf_path"))
        self.open_btns["pdf"].configure(state="normal" if has_pdf else "disabled")
        self.open_btns["report"].configure(state="normal")

    def open_result(self, key):
        if not self.result:
            return
        mode, res = self.result
        if mode == "sayi":
            path = {"folder": os.path.dirname(res["summary"]), "pdf": res.get("full_pdf"),
                    "report": res["summary"]}[key]
        else:
            path = {"folder": os.path.dirname(os.path.abspath(res["tex_path"])), "pdf": res.get("pdf_path"),
                    "report": res.get("report_path")}[key]
        open_path(path)


def main():
    try:
        import tkinter as tk
    except ImportError:
        sys.exit("Pencere arayüzü için Tkinter gerekli (Linux: sudo apt install python3-tk).")
    root = tk.Tk()
    try:
        from tkinter import ttk
        style = ttk.Style(root)
        if "vista" in style.theme_names():
            style.theme_use("vista")
        elif "clam" in style.theme_names():
            style.theme_use("clam")
    except Exception:  # noqa
        pass
    App(root)
    root.mainloop()
    return 0


if __name__ == "__main__":
    sys.exit(main())
