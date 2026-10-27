"""fraudecheck: finds similarities between student submissions (Canvas export).

Signals per pair of students:
  images    same or nearly the same image (md5 + 16x16 grayscale), also from embedded PDFs
  text      shared sentences (lightly reworded), paraphrased sentences (embeddings),
            JPlag score, text inside images (OCR)
  code      Dolos, similarity across the whole project (only outliers relative to the class)
  links     same Miro board
  metadata  same suspicious generator (python-docx, headless browser, ...) shortly after each other
Signals per student:
  links copied from an AI chat (utm_source=chatgpt.com, ...), Miro invite instead of a board link,
  visible link != actual link, document last edited by someone else, a lot of text in little
  editing time, screenshot of an AI chat, different tool/name than in earlier assignments.
Groups of 3+ students sharing the same piece are reported as clusters; whatever appears in a large
part of the class (template, assignment text, logo) is treated as course material and ignored.

Usage: python fraudecheck.py <folder> [--template t.docx] [--old last-year-folder] [-o report.html]
See README.md for all options.
"""
import argparse, base64, csv, fnmatch, glob, hashlib, html, io, json, math, os, re, shutil, subprocess, sys, tempfile, zipfile
from collections import defaultdict
from datetime import datetime, timezone
from itertools import combinations
from urllib.parse import urlparse
import numpy as np
from PIL import Image

TEXT_EXT = {".txt", ".md", ".html", ".css", ".json", ".csv"}
CODE_EXT = {".py", ".js", ".ts", ".java", ".c", ".h", ".cpp", ".cs", ".php", ".rb", ".go", ".rs", ".sh"}  # no .sql: mostly generated exports
IMG_EXT = {".png", ".jpg", ".jpeg", ".gif", ".bmp", ".webp"}
OFFICE_EXT = {".docx", ".pptx", ".xlsx"}
SKIP_DIRS = {"venv", ".venv", "env", "__pycache__", ".git", "node_modules", ".idea", ".vscode", ".vercel"}
SUSPECT_GENERATORS = ["python-docx", "python-pptx", "openpyxl", "headlesschrome", "reportlab",
                      "wkhtmltopdf", "weasyprint", "pandoc", "chatgpt", "openai", "claude"]
AI_LINKS = ["utm_source=chatgpt", "utm_source=openai", "utm_source=perplexity", "utm_source=copilot",
            "chatgpt.com/", "chat.openai.com", "claude.ai/", "gemini.google.com", "perplexity.ai", "copilot.microsoft.com"]
AI_SCREEN = ["chatgpt", "copy code", "kopieer code", "message chatgpt", "ask anything", "regenerate",  # UI text, also Dutch
             "claude.ai", "gemini", "copilot"]

MIN_IMG = 100        # ignore smaller images (icons)
OCR_MIN = 300        # only run OCR on images at least this wide
IMG_DIST = 0.03      # mean pixel difference on 16x16 grayscale: copy 0.001, nearest other diagram 0.06
MIN_WORDS = 6        # shorter sentences don't count
SHINGLE_SIM = 0.6    # Jaccard on word 3-grams: (lightly reworded) same sentence
PARA_SIM = 0.9       # cosine on sentence embeddings: same content, different words
NEAR_MINUTES = 60    # same suspicious generator within this many minutes
CLUSTER_MIN = 3      # a group of this many students is a cluster
COMMON_FRACTION = 0.15 # in >= 15% of the class (min. 4) = course material; copy rings are rarely larger
PASTE_WORDS, PASTE_MINUTES = 300, 10   # this many words in this little editing time = pasted?
WORD = r"[\wÀ-ÿ]+(?:['-][\wÀ-ÿ]+)*"   # words; stray punctuation doesn't count
EMBED_MODEL = "sentence-transformers/paraphrase-multilingual-MiniLM-L12-v2"
OCR_LANGS = "nld+eng"


def log(*a):
    print(*a, file=sys.stderr)

def run(cmd, **kw):
    return subprocess.run(cmd, capture_output=True, text=True, encoding="utf-8", errors="replace",
                          stdin=subprocess.DEVNULL, **kw).stdout


# ---------- extraction ----------

def empty_doc():
    return {"text": "", "images": [], "meta": {}, "links": []}

def office_extract(path, data=None):
    doc = empty_doc()
    with zipfile.ZipFile(io.BytesIO(data) if data else path) as z:
        names = z.namelist()
        parts = sorted(n for n in names if re.match(
            r"(word/(document|footnotes|endnotes|header\d*|footer\d*)|ppt/slides/slide\d+|xl/sharedStrings)\.xml$", n))
        for part in parts:
            xml = z.read(part).decode("utf-8", "replace")
            rels_name = re.sub(r"([^/]+)$", r"_rels/\1.rels", part)
            rels = {}
            if rels_name in names:
                r = z.read(rels_name).decode("utf-8", "replace")
                rels = dict(re.findall(r'Id="([^"]+)"[^>]*Target="([^"]+)"[^>]*TargetMode="External"', r))
                rels.update(re.findall(r'Id="([^"]+)"[^>]*TargetMode="External"[^>]*Target="([^"]+)"', r))
            seen = set()
            for rid, inner in re.findall(r'<w:hyperlink[^>]*r:id="([^"]+)"[^>]*>(.*?)</w:hyperlink>', xml, re.S):
                if rid in rels:
                    doc["links"].append((html.unescape(rels[rid]), html.unescape(re.sub(r"<[^>]+>", "", inner)).strip()))
                    seen.add(rid)
            doc["links"] += [(html.unescape(t), "") for rid, t in rels.items() if rid not in seen and t.startswith("http")]
            xml = re.sub(r"</w:p>|</a:p>|<w:br/>|<w:tab/>", "\n", xml)
            doc["text"] += "\n" + html.unescape(re.sub(r"<[^>]+>", "", xml))
        doc["images"] = [(n, z.read(n)) for n in names if "/media/" in n and os.path.splitext(n)[1].lower() in IMG_EXT]
        for name in ("docProps/core.xml", "docProps/app.xml"):
            if name in names:
                for k, v in re.findall(r"<(?:\w+:)?(creator|lastModifiedBy|created|modified|Application|TotalTime|Words|revision)[^>]*>([^<]*)<",
                                       z.read(name).decode("utf-8", "replace")):
                    doc["meta"][k] = v
        # embedded documents (e.g. a PDF flowchart inside a docx): include their text, images and links
        for n in names:
            if "/embeddings/" not in n:
                continue
            blob, inner = z.read(n), None
            if b"%PDF" in blob:
                inner = extract_bytes(blob[blob.index(b"%PDF"):blob.rfind(b"%%EOF") + 5], ".pdf")
            elif n.lower().endswith(tuple(OFFICE_EXT)):
                inner = office_extract(None, blob)
            if inner:
                doc["text"] += "\n" + inner["text"]
                doc["images"] += [(n + ":" + i, b) for i, b in inner["images"]]
                doc["links"] += inner["links"]
    return doc

def pdf_extract(path):
    doc = empty_doc()
    doc["text"] = run(["pdftotext", path, "-"])
    for line in run(["pdfinfo", path]).splitlines():
        k, _, v = line.partition(":")
        if k in ("Creator", "Producer", "Author", "CreationDate", "ModDate"):
            doc["meta"][k] = v.strip()
    xml = run(["pdftohtml", "-xml", "-i", "-stdout", "-q", path])
    doc["links"] = [(html.unescape(h), html.unescape(re.sub(r"<[^>]+>", "", t)).strip())
                    for h, t in re.findall(r'<a href="([^"]+)">([^<]*)', xml) if h.startswith("http")]
    # a visible URL as plain text right below a link (e.g. petsafe.com shown, petlibro.com linked);
    # only when the link above is a title, not a URL. Miro links are the visualisation, not a source.
    linked = {domain(h) for h, _ in doc["links"]}
    prev = None                        # href of the previous text element if it was a title link
    for el in re.findall(r"<text[^>]*>(.*?)</text>", xml):
        links = [(html.unescape(h), html.unescape(t)) for h, t in re.findall(r'<a href="(http[^"]+)">([^<]*)', el)]
        plain = html.unescape(re.sub(r"<a [^>]*>.*?</a>|<[^>]+>", "", el)).strip()
        for u in re.findall(r"(?:https?://|www\.)[^\s<]+", plain):
            d = domain(u)
            if prev and d not in linked and d != "miro.com" and looks_like_domain(d):
                doc["links"].append((prev, u))
        titles = [h for h, t in links if t.strip() and not re.search(r"https?://|www\.|\.(com|nl|be|org)\b", t.lower())]
        prev = titles[-1] if titles else (prev if not plain and links else None)
    with tempfile.TemporaryDirectory() as d:
        subprocess.run(["pdfimages", "-png", path, os.path.join(d, "i")], capture_output=True, stdin=subprocess.DEVNULL)
        doc["images"] = [(f, open(os.path.join(d, f), "rb").read()) for f in sorted(os.listdir(d))]
    return doc

def extract_bytes(data, ext):
    with tempfile.TemporaryDirectory() as d:
        p = os.path.join(d, "embedded" + ext)
        open(p, "wb").write(data)
        return extract(p)

def extract(path):
    """-> {"text", "images": [(name, bytes)], "meta": {}, "links": [(href, visible text)]}"""
    ext = os.path.splitext(path)[1].lower()
    try:
        if ext == ".pdf":
            return pdf_extract(path)
        if ext in OFFICE_EXT:
            return office_extract(path)
    except (zipfile.BadZipFile, OSError) as e:
        log(f"  could not read {os.path.basename(path)}: {e}")
        return empty_doc()
    doc = empty_doc()
    if ext in IMG_EXT:
        doc["images"] = [(os.path.basename(path), open(path, "rb").read())]
    elif ext in TEXT_EXT | CODE_EXT:
        doc["text"] = open(path, encoding="utf-8", errors="replace").read()
    return doc


# ---------- images ----------

def signature(data):
    """16x16 grayscale; None for small or unreadable images.
    (A dHash is too sensitive for diagrams with flat colour blocks.)"""
    try:
        im = Image.open(io.BytesIO(data))
        if min(im.size) < MIN_IMG:
            return None
        return im.convert("L").resize((16, 16), Image.Resampling.BOX).tobytes()
    except Exception:
        return None

def distance(a, b):
    """mean difference per pixel, 0 (equal) .. 1"""
    return float(np.abs(np.frombuffer(a, np.uint8).astype(int) - np.frombuffer(b, np.uint8)).mean() / 255)

def thumb(data, size=160):
    im = Image.open(io.BytesIO(data)).convert("RGB")
    im.thumbnail((size, size))
    buf = io.BytesIO()
    im.save(buf, "JPEG", quality=70)
    return "data:image/jpeg;base64," + base64.b64encode(buf.getvalue()).decode()

_ocr_cache = {}
def ocr(data):
    """text inside an image via tesseract; '' if tesseract is missing."""
    key = hashlib.md5(data).hexdigest()
    if key not in _ocr_cache:
        out = ""
        try:
            im = Image.open(io.BytesIO(data))
            if im.size[0] >= OCR_MIN and shutil.which("tesseract"):
                buf = io.BytesIO()
                im.convert("RGB").save(buf, "PNG")
                out = subprocess.run(["tesseract", "stdin", "stdout", "-l", OCR_LANGS], input=buf.getvalue(),
                                     capture_output=True).stdout.decode("utf-8", "replace")
        except Exception:
            pass
        _ocr_cache[key] = out
    return _ocr_cache[key]


# ---------- text ----------

def sentences(text):
    """-> [(normalised sentence, frozenset of word 3-grams)]"""
    out = []
    text = re.sub(r"(https?://|www\.)\S+", " ", text)   # links are checked separately
    for s in re.split(r"(?<=[.!?])\s+|\n\s*\n|\|", text):
        words = re.findall(WORD, s.lower())
        if len(words) >= MIN_WORDS:
            out.append((" ".join(words), frozenset(zip(words, words[1:], words[2:]))))
    return out

def similar(a, b):
    sa, sb = a[1], b[1]
    if not sa or not sb:
        return a[0] == b[0]
    return len(sa & sb) / len(sa | sb) >= SHINGLE_SIM

def shared_sentences(xs, ys):
    hits, seen = [], set()
    index = defaultdict(list)          # 3-gram -> sentences in ys
    for y in ys:
        for g in y[1]:
            index[g].append(y)
    for x in xs:
        cands = {id(y): y for g in x[1] for y in index.get(g, [])}
        for y in cands.values():
            if similar(x, y) and x[0] not in seen:
                seen.add(x[0])
                hits.append((x[0], y[0]))
                break
    return hits

_model = None
def embed(texts):
    """normalised sentence embeddings (fastembed, local); None if fastembed is missing."""
    global _model
    try:
        from fastembed import TextEmbedding
    except ImportError:
        return None
    if _model is None:
        log(f"  loading embedding model ({EMBED_MODEL}; downloaded on first use)")
        _model = TextEmbedding(EMBED_MODEL)
    v = np.array(list(_model.embed(texts)), dtype=np.float32)
    return v / np.linalg.norm(v, axis=1, keepdims=True)


# ---------- links ----------

def domain(s):
    s = s.strip().lower()
    host = urlparse(s if "://" in s else "http://" + s).netloc.split(":")[0]
    return host[4:] if host.startswith("www.") else host

TLDS = {"com", "nl", "be", "org", "net", "io", "de", "eu", "co", "uk", "fr", "ai", "app", "dev", "info", "edu", "gov", "us", "ca"}

def looks_like_domain(d):
    return bool(re.fullmatch(r"[a-z0-9-]+(\.[a-z0-9-]+)*\.[a-z]+", d)) and d.rsplit(".", 1)[-1] in TLDS

def same_site(a, b):
    """'hue.com' (a broken-off piece of philips-hue.com) counts as the same site."""
    return a == b or a.endswith(b) or b.endswith(a)

def miro_board(href):
    m = re.search(r"miro\.com/app/board/([^/?#]+)", href)
    return m.group(1) if m else None

def link_flags(links):
    flags = []
    for href, shown in links:
        h = href.lower()
        ai = next((p for p in AI_LINKS if p in h), None)
        if ai:
            flags.append(f"link copied from an AI chat ({ai}): {href[:90]}")
        if "miro.com/welcomeonboard" in h:
            flags.append("Miro link is a team invite, not a board link (may not open)")
        s = shown.strip().lower()
        if (s.startswith(("http", "www.")) or re.fullmatch(r"[a-z0-9.-]+(/\S*)?", s or "-")) \
                and looks_like_domain(domain(s)) and not same_site(domain(s), domain(href)):
            flags.append(f"visible link '{shown[:50]}' points to {domain(href)}")
    return list(dict.fromkeys(flags))


# ---------- metadata ----------

def parse_time(v):
    v = (v or "").strip()
    for fmt in ("%Y-%m-%dT%H:%M:%SZ", "%a %b %d %H:%M:%S %Y"):
        try:
            return datetime.strptime(re.split(r" (Romance|CET|CEST|UTC)", v)[0], fmt).replace(tzinfo=timezone.utc)
        except ValueError:
            pass
    return None

def generator(meta):
    return " / ".join(v for k, v in meta.items() if k in ("creator", "Creator", "Producer", "Application") and v)

def tool_family(meta):
    g = generator(meta).lower()
    for s in SUSPECT_GENERATORS:
        if s in g:
            return s
    for k, fam in (("google docs", "Google Docs"), ("libreoffice", "LibreOffice"), ("writer", "LibreOffice"),
                   ("pages", "Pages"), ("quartz", "Pages"), ("word", "Word")):
        if k in g:
            return fam
    return "unknown" if not g else "other"

def authors(meta, ignore):
    out = []
    for k in ("lastModifiedBy", "creator", "Author"):
        who = meta.get(k, "").strip()
        if who and who.lower() not in ignore and "python-docx" not in who.lower():
            out.append((k, who))
    return out

def student_flags(student, meta, ignore):
    flags = []
    fam = tool_family(meta)
    if fam in SUSPECT_GENERATORS:
        flags.append(f"made with {fam}")
    for k, who in authors(meta, ignore):
        if re.search(r"\d", who) or len(who.split()) > 4:   # a document title, not a person
            continue
        tokens = [t for t in re.findall(r"[a-z]+", who.lower()) if len(t) > 2]
        if tokens and not any(t in student for t in tokens):
            flags.append(f"{k} = '{who}' (not the student?)")
    try:
        words, minutes = int(meta.get("Words", 0)), int(meta.get("TotalTime", -1))
        if words >= PASTE_WORDS and 0 <= minutes <= PASTE_MINUTES:
            flags.append(f"{words} words in {minutes} min editing time (pasted?)")
    except ValueError:
        pass
    return flags


# ---------- history across assignments ----------

def history_flags(history, student, assignment, meta, ignore):
    prev = [h for h in history.get(student, []) if h["assignment"] != assignment]
    if not prev:
        return []
    flags, fam = [], tool_family(meta)
    fams = [h["tool"] for h in prev if h["tool"] != "unknown"]
    if fam != "unknown" and fams and fam not in fams:
        counts = ", ".join(f"{f} {fams.count(f)}x" for f in dict.fromkeys(fams))
        flags.append(f"different tool than before: now {fam} (before: {counts})")
    names = {n for h in prev for n in h["names"]}
    now = {who for _, who in authors(meta, ignore)}
    if names and now and not now & names:
        flags.append(f"different name in metadata than before: {', '.join(sorted(now))} (before: {', '.join(sorted(names))})")
    return flags

def history_update(history, student, assignment, meta, ignore):
    entry = {"assignment": assignment, "date": datetime.now().strftime("%Y-%m-%d"),
             "tool": tool_family(meta), "names": sorted({who for _, who in authors(meta, ignore)})}
    history[student] = [h for h in history.get(student, []) if h["assignment"] != assignment] + [entry]


# ---------- jplag and dolos (optional) ----------

def safe(name):
    return re.sub(r"[^\w-]", "_", name)

def jplag_scores(jar, texts, template_text, keep_dir=None, old=()):
    """texts: {student: text}; old: students from earlier years (only new<->old is compared)."""
    if not jar or not shutil.which("java"):
        return {}
    names = {safe(s): s for s in texts}
    with tempfile.TemporaryDirectory() as d:
        for s, t in texts.items():
            root = os.path.join(d, "old" if s in old else "new", safe(s))
            os.makedirs(root)
            open(os.path.join(root, "text.txt"), "w", encoding="utf-8").write(t)
        cmd = ["java", "-jar", jar, "-l", "text", "-M", "RUN", "-n", "-1", "-r", os.path.join(d, "res"), os.path.join(d, "new")]
        if old:
            cmd[-1:-1] = ["-old", os.path.join(d, "old")]
        if template_text:
            os.makedirs(os.path.join(d, "base"))
            open(os.path.join(d, "base", "text.txt"), "w", encoding="utf-8").write(template_text)
            cmd[3:3] = ["-bc", os.path.join(d, "base")]
        subprocess.run(cmd, capture_output=True, stdin=subprocess.DEVNULL)
        res = os.path.join(d, "res.jplag")
        if not os.path.exists(res):
            log("  JPlag produced no result, skipped")
            return {}
        if keep_dir:
            shutil.copy(res, os.path.join(keep_dir, "jplag.jplag"))
        top = json.loads(zipfile.ZipFile(res).read("topComparisons.json"))
        return {frozenset((names.get(c["firstSubmission"], c["firstSubmission"]), names.get(c["secondSubmission"], c["secondSubmission"]))):
                c["similarities"].get("AVG", 0) for c in top}

def dolos_scores(code_files, keep_dir=None):
    """code_files: {student: [paths]} -> {frozenset(pair): (project similarity, file_a, file_b)}.
    One Dolos run per extension. Project similarity = for each file its best match in the other
    student's work, averaged over all files of both students (one tiny boilerplate file barely counts)."""
    dolos = shutil.which("dolos")
    if not dolos:
        return {}
    owner = {os.path.normcase(os.path.abspath(p)): s for s, ps in code_files.items() for p in ps}
    by_ext = defaultdict(list)
    for p in owner:
        by_ext[os.path.splitext(p)[1]].append(p)
    best, per_file = {}, defaultdict(float)
    for ext, files in by_ext.items():
        if len({owner[p] for p in files}) < 2:
            continue
        with tempfile.TemporaryDirectory() as d:
            out = os.path.join(keep_dir, "dolos" + ext.replace(".", "-")) if keep_dir else os.path.join(d, "out")
            shutil.rmtree(out, ignore_errors=True)
            # file list as CSV (on Windows dolos is a .cmd and cmd.exe stops at 8191 characters);
            # Dolos only reads paths relative to the CSV, so the files are copied next to it
            src = {}
            for i, p in enumerate(files):
                name = f"{i}_{os.path.basename(p)}"
                shutil.copy(p, os.path.join(d, name))
                src[name] = p
            info = os.path.join(d, "info.csv")
            with open(info, "w", encoding="utf-8", newline="") as fh:
                w = csv.writer(fh)
                w.writerow(["filename", "full_name", "label"])
                w.writerows([n, n, owner[p]] for n, p in src.items())
            subprocess.run([dolos, "run", "-f", "csv", "-o", out, info], capture_output=True, stdin=subprocess.DEVNULL)
            csv_path = os.path.join(out, "pairs.csv")
            if not os.path.exists(csv_path):
                log(f"  Dolos produced no result for {ext}, skipped")
                continue
            for row in csv.DictReader(open(csv_path, encoding="utf-8")):
                fa = src.get(os.path.basename(row["leftFilePath"]))
                fb = src.get(os.path.basename(row["rightFilePath"]))
                a, b = owner.get(fa), owner.get(fb)
                if not a or not b or a == b:
                    continue
                key, sim = frozenset((a, b)), float(row["similarity"])
                per_file[(fa, b)] = max(per_file[(fa, b)], sim)
                per_file[(fb, a)] = max(per_file[(fb, a)], sim)
                if sim > best.get(key, (0,))[0]:
                    best[key] = (sim, os.path.basename(fa), os.path.basename(fb))
    out = {}
    for key, (_, f1, f2) in best.items():
        a, b = tuple(key)
        vals = [per_file[(f, b)] for f, s in owner.items() if s == a] + [per_file[(f, a)] for f, s in owner.items() if s == b]
        out[key] = (sum(vals) / len(vals), f1, f2)
    return out


# ---------- collecting ----------

def collect(folder, prefix="", exclude=()):
    """{student: [files]}; folders (unpacked zips) are read recursively."""
    students = defaultdict(list)
    for f in sorted(os.listdir(folder)):
        p = os.path.join(folder, f)
        if "_" not in f or any(fnmatch.fnmatch(f, x) for x in exclude):
            continue
        s = prefix + f.split("_")[0]
        if os.path.isfile(p):
            students[s].append(p)
        elif os.path.isdir(p):
            for root, dirs, files in os.walk(p):
                dirs[:] = [x for x in dirs if x not in SKIP_DIRS]
                students[s] += [os.path.join(root, x) for x in sorted(files) if not any(fnmatch.fnmatch(x, e) for e in exclude)]
    return students

class UnionFind:
    def __init__(self, n):
        self.p = list(range(n))
    def find(self, i):
        while self.p[i] != i:
            self.p[i] = self.p[self.p[i]]
            i = self.p[i]
        return i
    def union(self, a, b):
        self.p[self.find(a)] = self.find(b)
    def groups(self):
        g = defaultdict(list)
        for i in range(len(self.p)):
            g[self.find(i)].append(i)
        return list(g.values())

def find_groups(data, current):
    """Groups equal images and sentences across all students.
    Large part of the class -> course material (removed); 3+ students -> cluster."""
    common_min = max(CLUSTER_MIN + 1, math.ceil(COMMON_FRACTION * len(current)))
    clusters, common = [], {"images": 0, "sentences": 0}

    items = [(s, i) for s, d in data.items() for i in d["imgs"]]
    if items:
        sig = np.array([np.frombuffer(i["sig"], np.uint8) for _, i in items], dtype=np.int16)
        uf = UnionFind(len(items))
        for k in range(len(items)):
            dist = np.abs(sig[k + 1:] - sig[k]).mean(axis=1) / 255
            for j in np.nonzero(dist <= IMG_DIST)[0]:
                uf.union(k, k + 1 + j)
        for g in uf.groups():
            who = {items[k][0] for k in g}
            if len(who & current) >= common_min:
                for k in g:
                    items[k][1]["common"] = True
                common["images"] += 1
            elif len(who) >= CLUSTER_MIN:
                clusters.append({"kind": "image", "students": sorted(who), "example": items[g[0]][1]["bytes"]})
        for d in data.values():
            d["imgs"] = [i for i in d["imgs"] if not i.get("common")]

    items = [(s, x) for s, d in data.items() for x in d["sents"]]
    index = defaultdict(list)
    for k, (_, x) in enumerate(items):
        for g in x[1]:
            index[g].append(k)
    uf = UnionFind(len(items))
    for k, (s, x) in enumerate(items):
        for j in {j for g in x[1] for j in index[g] if j > k and items[j][0] != s}:
            if similar(x, items[j][1]):
                uf.union(k, j)
    drop = set()
    for g in uf.groups():
        who = {items[k][0] for k in g}
        if len(who & current) >= common_min:
            drop.update(id(items[k][1]) for k in g)
            common["sentences"] += 1
        elif len(who) >= CLUSTER_MIN:
            clusters.append({"kind": "sentence", "students": sorted(who), "example": items[g[0]][1][0]})
    for d in data.values():
        d["sents"] = [x for x in d["sents"] if id(x) not in drop]
    return clusters, common, common_min

def paraphrases(data, current, common_min, shared):
    """{frozenset(pair): [(sentence_a, sentence_b)]} for sentences with the same content in other words.
    Sentences that match a large part of the class this way don't count (assignment in own words)."""
    items = [(s, x[0]) for s, d in data.items() for x in d["sents"]]
    if len(items) < 2:
        return {}
    vec = embed([t for _, t in items])
    if vec is None:
        log("  fastembed not installed: paraphrase detection skipped")
        return {}
    owner = np.array([s for s, _ in items])
    sims = vec @ vec.T
    out = defaultdict(list)
    for k, (s, t) in enumerate(items):
        hit = np.nonzero((sims[k] >= PARA_SIM) & (owner != s))[0]
        who = set(owner[hit])
        if not hit.size or len(who & current) + 1 >= common_min:
            continue
        for j in hit:
            if j > k or owner[j] == s:
                continue
            key = frozenset((s, owner[j]))
            if (t, items[j][1]) not in shared.get(key, set()) and (items[j][1], t) not in shared.get(key, set()):
                out[key].append((t, items[j][1]))
    return out


# ---------- analysis ----------

def nodered_flow(path):
    """Node-RED export -> list of nodes; None for other (or empty) JSON."""
    try:
        nodes = json.load(open(path, encoding="utf-8-sig"))
    except (ValueError, OSError):
        return None
    if isinstance(nodes, list) and nodes and all(isinstance(n, dict) and "id" in n and "type" in n for n in nodes):
        return nodes
    return None

def handmade_ids(ids):
    """Node-RED makes random ids (16 hex, older 8.6 hex): readable or sequential ids
    (a5f1000000000001, ...02) mean the JSON was written or generated, not exported."""
    hexes = sorted(int(i, 16) for i in ids if re.fullmatch(r"[0-9a-f]{16}", i))
    odd = sum(not re.fullmatch(r"[0-9a-f]{16}|[0-9a-f]{8}\.[0-9a-f]{4,6}", i) for i in ids)
    return odd > len(ids) / 2 or sum(b - a < 1000 for a, b in zip(hexes, hexes[1:])) >= 3

def load_student(s, files, tmpl, ignore, use_ocr, code_dir):
    text, imgs, meta, links = "", [], {}, []
    code = [f for f in files if os.path.splitext(f)[1].lower() in CODE_EXT]
    ids, funcs = set(), {}
    for f in files:
        nodes = nodered_flow(f) if f.lower().endswith(".json") else None
        if nodes is not None:   # Node-RED export: function code goes to Dolos, the JSON itself is no text
            for n in nodes:
                ids.add(str(n["id"]))
                text += "".join("\n" + n[k] for k in ("name", "info") if isinstance(n.get(k), str))
                src = "\n".join(n[k] for k in ("initialize", "func", "finalize") if isinstance(n.get(k), str) and n[k].strip())
                if not src:
                    continue
                norm = re.sub(r"\s+", "", re.sub(r"/\*.*?\*/|//.*", "", src, flags=re.S))   # comments/layout don't count
                if len(norm) >= 80:
                    funcs[hashlib.md5(norm.encode()).hexdigest()] = n.get("name") or str(n["id"])
                p = os.path.join(code_dir, f"{safe(s)}_{len(code)}_{safe(n.get('name') or str(n['id']))}.js")
                open(p, "w", encoding="utf-8").write(src)
                code.append(p)
            continue
        doc = extract(f)
        if os.path.splitext(f)[1].lower() not in CODE_EXT:   # code goes to Dolos
            text += "\n" + doc["text"]
        imgs += [(os.path.basename(f), n, b) for n, b in doc["images"]]
        links += doc["links"]
        if doc["meta"] and not meta:
            meta = dict(doc["meta"])
    t = parse_time(meta.get("created") or meta.get("CreationDate"))
    if t:
        meta["_time"] = t
    hashed, ocr_text, ai_screens = [], "", []
    for f, n, b in imgs:
        h = signature(b)
        if h is None or any(distance(h, x) <= IMG_DIST for x in tmpl["hashes"]):
            continue
        hashed.append({"file": f, "name": n, "bytes": b, "sig": h, "md5": hashlib.md5(b).hexdigest()})
        if use_ocr:
            o = ocr(b)
            ocr_text += "\n" + o
            hits = [w for w in AI_SCREEN if w in o.lower()]
            if hits:
                ai_screens.append(f"{n}: {', '.join(hits)}")
    sents = [x for x in sentences(text + "\n" + ocr_text)
             if not any(similar(x, y) for y in tmpl["sents"])
             and not (x[1] and len(x[1] & tmpl["grams"]) >= 0.6 * len(x[1]))]
    flags = student_flags(s, meta, ignore) + link_flags(links)
    flags += [f"image looks like a screenshot of an AI chat ({x})" for x in ai_screens]
    if ids and handmade_ids(ids):
        flags.append("Node-RED node ids are not random (readable or sequential): flow written or generated, not exported")
    return {"text": text, "sents": sents, "imgs": hashed, "meta": meta, "links": links, "flags": flags,
            "code": code, "ids": ids, "funcs": funcs, "files": [os.path.basename(f) for f in files]}

def analyse(folder, template=None, jar=None, ignore=(), ignore_images=(), old=(), use_ocr=True, use_embed=True,
            use_dolos=True, history=None, keep_dir=None, exclude=()):
    ignore = {i.lower() for i in ignore}
    t = extract(template) if template else empty_doc()
    w = re.findall(WORD, t["text"].lower())
    tmpl = {"text": t["text"], "sents": sentences(t["text"]), "grams": set(zip(w, w[1:], w[2:])),
            "hashes": [h for _, b in t["images"] + [(p, open(p, "rb").read()) for p in ignore_images]
                       if (h := signature(b)) is not None]}
    assignment = os.path.basename(os.path.abspath(folder.rstrip("/\\")))

    students = dict(collect(folder, exclude=exclude))
    current = set(students)
    for o in old:
        students.update(collect(o, prefix=os.path.basename(os.path.abspath(o.rstrip("/\\"))) + "/", exclude=exclude))
    old_names = set(students) - current

    # ponytail: without keep_dir the extracted Node-RED code stays in a temp folder
    code_dir = os.path.join(keep_dir, "nodered") if keep_dir else tempfile.mkdtemp(prefix="fraudecheck-")
    shutil.rmtree(code_dir, ignore_errors=True)
    os.makedirs(code_dir, exist_ok=True)   # rmtree fails silently on Windows while the folder is open
    data = {}
    for s, files in students.items():
        data[s] = load_student(s, files, tmpl, ignore, use_ocr, code_dir)
        if history is not None and s in current:
            data[s]["flags"] += history_flags(history, s, assignment, data[s]["meta"], ignore)
            history_update(history, s, assignment, data[s]["meta"], ignore)
        log(f"  {s}: {len(files)} file(s), {len(data[s]['sents'])} sentences, {len(data[s]['imgs'])} images")

    clusters, common, common_min = find_groups(data, current)
    jp = jplag_scores(jar, {s: d["text"] for s, d in data.items() if d["text"].strip()}, tmpl["text"], keep_dir, old_names)
    dl = dolos_scores({s: d["code"] for s, d in data.items() if d["code"]}, keep_dir) if use_dolos else {}
    # code from the same lessons always looks alike: only outliers relative to the class count
    sims = [v[0] for k, v in dl.items() if k <= current]
    if len(sims) >= 3:
        mean = sum(sims) / len(sims)
        dolos_min = max(0.5, mean + 2 * (sum((x - mean) ** 2 for x in sims) / len(sims)) ** 0.5)
    else:
        dolos_min = 0.8

    shared_all = {}
    for a, b in combinations(sorted(data), 2):
        if a in old_names and b in old_names:
            continue
        shared_all[frozenset((a, b))] = shared_sentences(data[a]["sents"], data[b]["sents"])
    para = paraphrases(data, current, common_min, {k: set(v) for k, v in shared_all.items()}) if use_embed else {}
    # node ids and functions in a large part of the class come from a lesson flow
    id_count = defaultdict(int)
    fn_count = defaultdict(int)
    for s in current:
        for i in data[s]["ids"]:
            id_count[i] += 1
        for h in data[s]["funcs"]:
            fn_count[h] += 1

    pairs = []
    for key, shared in shared_all.items():
        a, b = sorted(key)
        A, B = data[a], data[b]
        reasons, score, imgs = [], 0, []
        for x in A["imgs"]:
            for y in B["imgs"]:
                if x["md5"] == y["md5"] or distance(x["sig"], y["sig"]) <= IMG_DIST:
                    imgs.append((x["bytes"], y["bytes"], "identical file" if x["md5"] == y["md5"]
                                 else f"nearly equal (difference {distance(x['sig'], y['sig']):.3f})"))
        if imgs:
            score += 50
            reasons.append(f"{len(imgs)} matching image(s)")
        if shared:
            score += min(30, 5 * len(shared))
            reasons.append(f"{len(shared)} shared sentence(s)")
        p = para.get(key, [])
        if len(p) >= 2:   # a single paraphrased sentence is too often a coincidence
            score += min(20, 4 * len(p))
            reasons.append(f"{len(p)} paraphrased sentence(s) with the same content")
        j = jp.get(key, 0)
        if j >= 0.1:
            score += round(50 * j)
            reasons.append(f"JPlag {j:.0%}")
        dsim = dl.get(key)
        if dsim and dsim[0] >= dolos_min:
            score += round(60 * dsim[0])
            reasons.append(f"Dolos: code {dsim[0]:.0%} similar across the project (class threshold {dolos_min:.0%}; mainly {dsim[1]} ~ {dsim[2]})")
        ids = [i for i in A["ids"] & B["ids"] if id_count[i] < common_min]
        if len(ids) >= 3:   # random ids only match when one imported the other's flow
            score += 40
            reasons.append(f"{len(ids)} identical Node-RED node ids (flow copied from each other)")
        same = sorted(A["funcs"][h] for h in A["funcs"].keys() & B["funcs"].keys() if fn_count[h] < common_min)
        if same:
            score += min(30, 10 * len(same))
            reasons.append(f"{len(same)} identical function node(s): {', '.join(same[:4])}")
        boards = {miro_board(h) for h, _ in A["links"]} & {miro_board(h) for h, _ in B["links"]} - {None}
        if boards:
            score += 40
            reasons.append(f"same Miro board ({', '.join(boards)})")
        fa, fb = tool_family(A["meta"]), tool_family(B["meta"])
        if fa == fb and fa in SUSPECT_GENERATORS:
            score += 15
            reasons.append(f"both made with {fa}")
            ta, tb = A["meta"].get("_time"), B["meta"].get("_time")
            if ta and tb and abs((ta - tb).total_seconds()) <= NEAR_MINUTES * 60:
                score += 10
                reasons.append(f"created {abs(ta - tb).total_seconds() / 60:.0f} min apart")
        if score:
            pairs.append({"a": a, "b": b, "score": score, "reasons": reasons, "imgs": imgs, "shared": shared, "para": p,
                          "old": a in old_names or b in old_names})
    pairs.sort(key=lambda p: -p["score"])
    code_top = sorted(((v[0], *sorted(k), v[1], v[2]) for k, v in dl.items()), reverse=True)[:10]
    return {"data": data, "pairs": pairs, "clusters": clusters, "common": common, "common_min": common_min,
            "code_top": code_top, "dolos_min": dolos_min, "current": current, "old": old_names}


# ---------- report ----------

def why(p):
    return ("earlier year: " if p["old"] else "") + ", ".join(p["reasons"])

def report(res, out, title, keep_dir=None, jar=None):
    e = html.escape
    data, current = res["data"], res["current"]
    rows = []
    for p in res["pairs"]:
        imgs = "".join(f'<div class="im"><img src="{thumb(x)}"><img src="{thumb(y)}"><br>{e(w)}</div>' for x, y, w in p["imgs"][:4])
        sh = "".join(f"<li>{e(x)}</li>" for x, _ in p["shared"][:8])
        pa = "".join(f"<li>{e(x)}<br><i>~ {e(y)}</i></li>" for x, y in p["para"][:6])
        rows.append(f'<tr class="{"hi" if p["score"] >= 40 else ""}"><td>{p["score"]}</td><td><b>{e(p["a"])}</b><br><b>{e(p["b"])}</b></td>'
                    f'<td>{e(why(p))}{imgs}{"<p>Shared sentences:</p><ul>" + sh + "</ul>" if sh else ""}'
                    f'{"<p>Paraphrased:</p><ul>" + pa + "</ul>" if pa else ""}</td></tr>')
    cl_rows = "".join(
        f'<tr><td>{c["kind"]}</td><td>{e(", ".join(c["students"]))}</td><td>'
        + (f'<img class="t" src="{thumb(c["example"])}">' if c["kind"] == "image" else e(c["example"]))
        + "</td></tr>" for c in res["clusters"])
    code_rows = "".join(f'<tr class="{"hi" if s >= res["dolos_min"] else ""}"><td>{s:.0%}</td><td>{e(a)}<br>{e(b)}</td><td>{e(f1)} ~ {e(f2)}</td></tr>'
                        for s, a, b, f1, f2 in res["code_top"])

    def cell(v):
        return e(v) if v else '<span class="u">unknown</span>'
    meta_rows = []
    for s in sorted(current):
        d = data[s]
        m = d["meta"]
        edit = f'{m["TotalTime"]} min / {m.get("Words", "?")} words' if "TotalTime" in m else ""
        meta_rows.append(f'<tr class="{"hi" if d["flags"] else ""}"><td>{e(s)}</td><td>{cell(generator(m))}</td>'
                         f'<td>{cell(m.get("lastModifiedBy") or m.get("Author", ""))}</td>'
                         f'<td>{cell(m["_time"].strftime("%d/%m %H:%M") if m.get("_time") else "")}</td><td>{cell(edit)}</td>'
                         f'<td>{"<br>".join(e(f) for f in d["flags"])}</td></tr>')
    viewers = ""
    if keep_dir:
        lines = []
        if os.path.exists(os.path.join(keep_dir, "jplag.jplag")):
            lines.append(f'java -jar "{jar}" -M VIEW "{os.path.join(keep_dir, "jplag.jplag")}"')
        lines += [f'dolos serve "{d}"' for d in sorted(glob.glob(os.path.join(keep_dir, "dolos-*")))]
        if lines:
            viewers = "<h2>Details in JPlag / Dolos</h2><p>Open the side-by-side comparisons with:</p><pre>" + e("\n".join(lines)) + "</pre>"
    n_old = len(res["old"])
    page = f"""<!doctype html><meta charset="utf-8"><title>{e(title)}</title>
<style>body{{font:14px system-ui,sans-serif;margin:24px;max-width:1200px}}table{{border-collapse:collapse;width:100%;margin-bottom:32px}}
td,th{{border:1px solid #ccc;padding:6px;vertical-align:top;text-align:left}}tr.hi{{background:#fff1e6}}.u{{color:#999}}
.im{{display:inline-block;margin:6px 12px 0 0;font-size:12px}}.im img,img.t{{max-height:120px;margin-right:4px;border:1px solid #aaa}}
li{{font-size:12px;margin-bottom:4px}}pre{{background:#f4f4f4;padding:8px;overflow-x:auto}}</style>
<h1>{e(title)}</h1>
<p>{len(current)} students{f", compared with {n_old} submissions from earlier years" if n_old else ""}; {len(res["pairs"])} pairs with a signal,
{len(res["clusters"])} clusters. Ignored as course material (in {res["common_min"]}+ students): {res["common"]["images"]} image(s),
{res["common"]["sentences"]} sentence(s). Score: image 50, Miro board 40, shared sentence 5 (max 30), paraphrased sentence 4 (max 20),
JPlag 50×%, Dolos 60×%, same suspicious generator 15 (+10 within {NEAR_MINUTES} min),
3+ identical Node-RED node ids 40, identical function node 10 (max 30).
<b>These are signals, not proof: always check the originals.</b></p>
<h2>Suspicious pairs</h2><table><tr><th>Score</th><th>Students</th><th>Why</th></tr>{"".join(rows) or "<tr><td colspan=3>none</td></tr>"}</table>
<h2>Clusters (3+ students sharing the same piece)</h2><table><tr><th>What</th><th>Students</th><th>Example</th></tr>{cl_rows or "<tr><td colspan=3>none</td></tr>"}</table>
{f'<h2>Code (Dolos)</h2><p>Similarity across the whole project; highlighted from the class threshold {res["dolos_min"]:.0%} (mean + 2 standard deviations, min. 50%).</p><table><tr><th>Similar</th><th>Students</th><th>Most similar files</th></tr>{code_rows}</table>' if code_rows else ""}
<h2>Per student</h2><table><tr><th>Student</th><th>Made with</th><th>Last modified by / author</th><th>Created (UTC)</th><th>Editing time</th><th>Notable</th></tr>{"".join(meta_rows)}</table>
{viewers}"""
    open(out, "w", encoding="utf-8").write(page)

def to_json(res, out):
    data = res["data"]
    j = {"students": {s: {"flags": data[s]["flags"], "generator": generator(data[s]["meta"]), "tool": tool_family(data[s]["meta"]),
                          "files": data[s]["files"], "links": data[s]["links"]} for s in sorted(res["current"])},
         "pairs": [{"a": p["a"], "b": p["b"], "score": p["score"], "earlier_year": p["old"], "reasons": p["reasons"],
                    "images": len(p["imgs"]), "shared": [x for x, _ in p["shared"]], "paraphrased": p["para"]} for p in res["pairs"]],
         "clusters": [{"kind": c["kind"], "students": c["students"], "example": c["example"] if c["kind"] == "sentence" else None}
                      for c in res["clusters"]],
         "code": [{"similarity": s, "a": a, "b": b, "files": [f1, f2]} for s, a, b, f1, f2 in res["code_top"]],
         "dolos_threshold": res["dolos_min"], "common": res["common"]}
    json.dump(j, open(out, "w", encoding="utf-8"), ensure_ascii=False, indent=1)


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("folder", help="folder with submissions (Canvas names: <student>_<id>_<id>_<file>)")
    ap.add_argument("--template", help="blank assignment/template: its text and images are ignored")
    ap.add_argument("--old", action="append", default=[], help="folder with submissions from an earlier year (repeatable)")
    ap.add_argument("--ignore-author", action="append", default=[], help="author name that is expected (e.g. the teacher)")
    ap.add_argument("--ignore-image", action="append", default=[], help="image to ignore (e.g. a school logo)")
    ap.add_argument("--exclude", action="append", default=[], help="file name pattern to skip, e.g. '*_feedback.txt' (repeatable)")
    ap.add_argument("--jplag", help="path to jplag-*.jar (default: newest ~/tools/jplag/jplag-*.jar)")
    ap.add_argument("--history", default=os.path.expanduser("~/.fraudecheck/history.json"),
                    help="file with each student's history across assignments")
    ap.add_argument("--no-history", action="store_true", help="don't read or update the history")
    ap.add_argument("--no-ocr", action="store_true", help="no text recognition in images")
    ap.add_argument("--no-embed", action="store_true", help="no paraphrase detection (embeddings)")
    ap.add_argument("-o", "--output", default="fraudecheck.html")
    a = ap.parse_args()
    sys.stdout.reconfigure(encoding="utf-8")
    sys.stderr.reconfigure(encoding="utf-8")
    if not a.jplag:
        jars = sorted(glob.glob(os.path.expanduser("~/tools/jplag/jplag-*.jar")))
        a.jplag = jars[-1] if jars else None
    log(f"JPlag (text): {a.jplag or 'not found'} | Dolos (code): {shutil.which('dolos') or 'not found'} | "
        f"OCR: {'off' if a.no_ocr else shutil.which('tesseract') or 'not found'}")

    history = None
    if not a.no_history:
        history = json.load(open(a.history, encoding="utf-8")) if os.path.exists(a.history) else {}
    keep_dir = os.path.splitext(a.output)[0] + "_files"
    os.makedirs(keep_dir, exist_ok=True)

    res = analyse(a.folder, a.template, a.jplag, a.ignore_author, a.ignore_image, a.old, not a.no_ocr,
                  not a.no_embed, True, history, keep_dir, a.exclude)
    title = "Fraudecheck – " + os.path.basename(os.path.abspath(a.folder.rstrip("/\\")))
    report(res, a.output, title, keep_dir, a.jplag)
    to_json(res, os.path.splitext(a.output)[0] + ".json")
    if history is not None:
        os.makedirs(os.path.dirname(a.history), exist_ok=True)
        json.dump(history, open(a.history, "w", encoding="utf-8"), ensure_ascii=False, indent=1)

    for p in res["pairs"][:10]:
        print(f"{p['score']:4}  {p['a']} <-> {p['b']}: {why(p)}")
    for c in res["clusters"][:5]:
        print(f"cluster ({c['kind']}): {', '.join(c['students'])}")
    flagged = [s for s in sorted(res["current"]) if res["data"][s]["flags"]]
    print(f"{len(flagged)} students with a notable finding (see report)")
    if res["code_top"]:
        s, x, y, *_ = res["code_top"][0]
        print(f"code: highest project similarity {s:.0%} ({x} <-> {y}), class threshold {res['dolos_min']:.0%}")
    print(f"report: {os.path.abspath(a.output)}")

if __name__ == "__main__":
    main()
