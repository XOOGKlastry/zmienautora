"""ZmienAutora - hurtowa zmiana autora i śladów edycji w plikach Word i Excel.

Obsługuje: .docx .docm .dotx .xlsx .xlsm .xltx
Wybierz folder, wpisz autora, kliknij Start.
"""
import os
import re
import shutil
import threading
import zipfile
from datetime import datetime, timezone
from xml.sax.saxutils import escape

EXT_WORD = {".docx", ".docm", ".dotx", ".dotm"}
EXT_EXCEL = {".xlsx", ".xlsm", ".xltx", ".xltm"}
BACKUP_DIR = "_kopia_oryginalow"


def esc_attr(s):
    return escape(s, {'"': "&quot;"})


def sub_attr(text, name, value):
    """Zamienia wartość atrybutu name="..." (name może mieć prefiks)."""
    pat = re.compile(r'(\s%s=")[^"]*(")' % re.escape(name))
    new, n = pat.subn(lambda m: m.group(1) + esc_attr(value) + m.group(2), text)
    return new, n


def sub_elem(text, tag, value):
    """Zamienia zawartość <tag ...>...</tag> lub <tag/>."""
    t = re.escape(tag)
    v = escape(value)
    n_total = 0
    text, n = re.subn(r"(<%s(?:\s[^>]*)?>)(.*?)(</%s>)" % (t, t),
                      lambda m: m.group(1) + v + m.group(3), text, flags=re.S)
    n_total += n
    text, n = re.subn(r"<%s((?:\s[^>]*)?)/>" % t,
                      lambda m: "<%s%s>%s</%s>" % (tag, m.group(1), v, tag), text)
    n_total += n
    return text, n_total


def process_xml(name, text, o):
    """Zwraca (nowy_tekst, liczba_zmian) dla jednej części pliku."""
    a, ini, dt, company = o["author"], o["initials"], o["date"], o["company"]
    c = 0

    def do(res):
        nonlocal text, c
        text, n = res
        c += n

    if name == "docProps/core.xml":
        do(sub_elem(text, "dc:creator", a))
        do(sub_elem(text, "cp:lastModifiedBy", a))
        do(sub_elem(text, "cp:revision", "1"))
        do(sub_elem(text, "dcterms:created", dt))
        do(sub_elem(text, "dcterms:modified", dt))
        do(sub_elem(text, "cp:lastPrinted", dt))
        return text, c

    if name == "docProps/app.xml":
        do(sub_elem(text, "Company", company))
        do(sub_elem(text, "Manager", ""))
        do(sub_elem(text, "TotalTime", "0"))
        return text, c

    if name == "docProps/custom.xml":
        # usuń wszystkie właściwości niestandardowe
        text, n = re.subn(r"<property\b.*?</property>", "", text, flags=re.S)
        return text, n

    if name.startswith("word/"):
        do(sub_attr(text, "w:author", a))
        do(sub_attr(text, "w:initials", ini))
        do(sub_attr(text, "w:date", dt))
        do(sub_attr(text, "w16du:dateUtc", dt))
        # identyfikatory sesji edycji (rsid)
        text, n = re.subn(r'\sw:rsid\w*="[^"]*"', "", text)
        c += n
        text, n = re.subn(r"<w:rsids>.*?</w:rsids>", "", text, flags=re.S)
        c += n
        if name == "word/people.xml":
            do(sub_attr(text, "w15:author", a))
            do(sub_attr(text, "w15:userId", a))
            do(sub_attr(text, "w15:providerId", "None"))
        return text, c

    if name == "word/_rels/settings.xml.rels":
        text, n = re.subn(r'(Type="[^"]*attachedTemplate"[^>]*Target=")[^"]*(")',
                          r"\1Normal.dotm\2", text)
        text, n2 = re.subn(r'(Target=")[^"]*("[^>]*Type="[^"]*attachedTemplate")',
                           r"\1Normal.dotm\2", text)
        return text, n + n2

    if name.startswith("xl/"):
        if re.match(r"xl/(comments/)?comments?\d*\.xml$", name):
            do(sub_elem(text, "author", a))
        if name.startswith("xl/persons/"):
            do(sub_attr(text, "displayName", a))
            do(sub_attr(text, "userId", a))
            do(sub_attr(text, "providerId", "None"))
        if name.startswith("xl/threadedComments/"):
            do(sub_attr(text, "dT", dt))
        if name.startswith("xl/revisions/"):
            do(sub_attr(text, "userName", a))
            do(sub_attr(text, "dateTime", dt))
        if name == "xl/workbook.xml":
            # lokalna ścieżka zapisu pliku (zawiera nazwę użytkownika Windows)
            text, n = re.subn(r"<mc:AlternateContent\b(?:(?!</mc:AlternateContent>).)*?absPath.*?</mc:AlternateContent>",
                              "", text, flags=re.S)
            c += n
            text, n = re.subn(r"<x15ac:absPath\b[^>]*/>", "", text)
            c += n
        return text, c

    return text, 0


def process_file(path, o):
    with zipfile.ZipFile(path) as zin:
        items = [(i, zin.read(i.filename)) for i in zin.infolist()]
    total = 0
    out = []
    for info, data in items:
        n = info.filename
        if n.endswith(".xml") or n.endswith(".rels"):
            try:
                text = data.decode("utf-8")
            except UnicodeDecodeError:
                out.append((info, data))
                continue
            new, c = process_xml(n, text, o)
            if c:
                total += c
                data = new.encode("utf-8")
        out.append((info, data))
    if not total:
        return 0
    tmp = path + ".tmp"
    with zipfile.ZipFile(tmp, "w") as zout:
        for info, data in out:
            zi = zipfile.ZipInfo(info.filename, date_time=info.date_time)
            zi.compress_type = info.compress_type
            zi.external_attr = info.external_attr
            zout.writestr(zi, data)
    os.replace(tmp, path)
    return total


def find_files(folder, recursive):
    exts = EXT_WORD | EXT_EXCEL
    for root, dirs, files in os.walk(folder):
        dirs[:] = [d for d in dirs if d != BACKUP_DIR]
        for f in files:
            if f.startswith("~$"):
                continue
            if os.path.splitext(f)[1].lower() in exts:
                yield os.path.join(root, f)
        if not recursive:
            break


def run(folder, o, recursive, backup, log):
    files = list(find_files(folder, recursive))
    if not files:
        log("Nie znaleziono plików Word/Excel.")
        return
    bdir = os.path.join(folder, BACKUP_DIR, datetime.now().strftime("%Y-%m-%d_%H%M%S"))
    ok = err = 0
    for p in files:
        rel = os.path.relpath(p, folder)
        try:
            if backup:
                dst = os.path.join(bdir, rel)
                os.makedirs(os.path.dirname(dst), exist_ok=True)
                shutil.copy2(p, dst)
            n = process_file(p, o)
            log(f"OK   {rel}  ({n} zmian)")
            ok += 1
        except PermissionError:
            log(f"BŁĄD {rel}  - plik jest otwarty w innym programie, zamknij go")
            err += 1
        except zipfile.BadZipFile:
            log(f"BŁĄD {rel}  - plik zaszyfrowany hasłem lub uszkodzony")
            err += 1
        except Exception as e:
            log(f"BŁĄD {rel}  - {e}")
            err += 1
    log(f"\nGotowe. Zmienione: {ok}, błędy: {err}.")
    if backup:
        log(f"Kopie oryginałów: {bdir}")


def make_initials(name):
    return "".join(w[0] for w in name.split() if w).upper()[:4]


def gui():
    import tkinter as tk
    from tkinter import filedialog, messagebox, scrolledtext

    w = tk.Tk()
    w.title("Zmień autora - Word i Excel")
    w.geometry("680x520")
    pad = {"padx": 8, "pady": 4}

    folder = tk.StringVar()
    author = tk.StringVar()
    company = tk.StringVar()
    date = tk.StringVar(value=datetime.now().strftime("%Y-%m-%d %H:%M"))
    recursive = tk.BooleanVar(value=True)
    backup = tk.BooleanVar(value=True)

    f = tk.Frame(w)
    f.pack(fill="x", **pad)
    f.columnconfigure(1, weight=1)
    tk.Label(f, text="Folder:").grid(row=0, column=0, sticky="w")
    tk.Entry(f, textvariable=folder).grid(row=0, column=1, sticky="ew", padx=4)
    tk.Button(f, text="Wybierz...",
              command=lambda: folder.set(filedialog.askdirectory() or folder.get())
              ).grid(row=0, column=2)
    tk.Label(f, text="Autor:").grid(row=1, column=0, sticky="w")
    tk.Entry(f, textvariable=author).grid(row=1, column=1, sticky="ew", padx=4, pady=3)
    tk.Label(f, text="Firma (puste = usuń):").grid(row=2, column=0, sticky="w")
    tk.Entry(f, textvariable=company).grid(row=2, column=1, sticky="ew", padx=4, pady=3)
    tk.Label(f, text="Data (RRRR-MM-DD GG:MM):").grid(row=3, column=0, sticky="w")
    tk.Entry(f, textvariable=date).grid(row=3, column=1, sticky="ew", padx=4, pady=3)
    tk.Checkbutton(f, text="Również podfoldery", variable=recursive).grid(row=4, column=1, sticky="w")
    tk.Checkbutton(f, text=f"Zrób kopię oryginałów (folder {BACKUP_DIR})",
                   variable=backup).grid(row=5, column=1, sticky="w")

    out = scrolledtext.ScrolledText(w, height=16, font=("Consolas", 9))

    def log(msg):
        w.after(0, lambda: (out.insert("end", msg + "\n"), out.see("end")))

    def start():
        if not os.path.isdir(folder.get()):
            messagebox.showerror("Błąd", "Wybierz istniejący folder.")
            return
        if not author.get().strip():
            messagebox.showerror("Błąd", "Wpisz autora.")
            return
        try:
            d = datetime.strptime(date.get().strip(), "%Y-%m-%d %H:%M")
        except ValueError:
            messagebox.showerror("Błąd", "Data w formacie RRRR-MM-DD GG:MM")
            return
        d = d.astimezone(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")
        a = author.get().strip()
        o = {"author": a, "initials": make_initials(a), "date": d,
             "company": company.get().strip()}
        btn.config(state="disabled")
        out.delete("1.0", "end")

        def job():
            try:
                run(folder.get(), o, recursive.get(), backup.get(), log)
            finally:
                w.after(0, lambda: btn.config(state="normal"))
        threading.Thread(target=job, daemon=True).start()

    btn = tk.Button(w, text="Start", width=20, command=start, bg="#2d6cdf", fg="white")
    btn.pack(**pad)
    out.pack(fill="both", expand=True, **pad)
    w.mainloop()


if __name__ == "__main__":
    gui()
