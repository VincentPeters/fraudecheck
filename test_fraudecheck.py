"""Self-test with synthetic submissions: python test_fraudecheck.py

Builds a fake class of 10 students + 1 from an earlier year and checks that every signal
is found, and that template material and the 'clean' students produce nothing.
JPlag, Dolos and embeddings are off, so the test runs the same everywhere; OCR is
checked if Tesseract is installed.
"""
import io, json, os, tempfile, zipfile
from PIL import Image, ImageDraw
import fraudecheck as fc


# ---------- helpers ----------

def png(shapes, size=(400, 300), fmt="PNG", resize=None):
    im = Image.new("RGB", size, "white")
    d = ImageDraw.Draw(im)
    for kind, box, color in shapes:
        (d.rectangle if kind == "r" else d.ellipse)(box, fill=color)
    if resize:
        im = im.resize(resize)
    buf = io.BytesIO()
    im.save(buf, fmt)
    return buf.getvalue()

DIAGRAM = [("r", [60, 20 + i * 70, 340, 70 + i * 70], c) for i, c in enumerate(["blue", "green", "orange", "purple"])]
TEMPLATE_IMG = [("r", [0, 0, 400, 60], "black"), ("e", [150, 100, 250, 200], "red")]

def docx(path, paragraphs, images=(), links=(), creator="Student", modified_by="Student",
         created="2026-10-20T10:00:00Z", app="Microsoft Office Word", minutes=120, words=400):
    """minimal docx: only what fraudecheck reads"""
    rels = "".join(f'<Relationship Id="rIdL{i}" Type="hyperlink" Target="{h}" TargetMode="External"/>' for i, (h, _) in enumerate(links))
    body = "".join(f"<w:p><w:r><w:t>{p}</w:t></w:r></w:p>" for p in paragraphs)
    body += "".join(f'<w:p><w:hyperlink r:id="rIdL{i}"><w:r><w:t>{t}</w:t></w:r></w:hyperlink></w:p>' for i, (_, t) in enumerate(links))
    with zipfile.ZipFile(path, "w") as z:
        z.writestr("word/document.xml", f'<w:document xmlns:w="w" xmlns:r="r"><w:body>{body}</w:body></w:document>')
        z.writestr("word/_rels/document.xml.rels", f"<Relationships>{rels}</Relationships>")
        z.writestr("docProps/core.xml", f"<cp:coreProperties><dc:creator>{creator}</dc:creator>"
                   f"<cp:lastModifiedBy>{modified_by}</cp:lastModifiedBy><dcterms:created>{created}</dcterms:created></cp:coreProperties>")
        z.writestr("docProps/app.xml", f"<Properties><Application>{app}</Application><TotalTime>{minutes}</TotalTime><Words>{words}</Words></Properties>")
        for i, data in enumerate(images):
            z.writestr(f"word/media/image{i}.png", data)

TEMPLATE_TXT = "Step 1 - Purpose of the application: what is the main goal of the application you are analysing?"
SHARED_Q = ["The user opens the app and immediately sees all notifications from the past week.",
            "Every notification shows the time, the sensor and the measured value in a clear table.",
            "When a value is too high the row turns red so that the problem stands out."]

def unique(name, n=3):
    return [f"{name.capitalize()} describes topic number {i} here in their own words about {name}{i}land and {name}{i}town." for i in range(n)]


def node(id, type, name="", func=None):
    n = {"id": id, "type": type, "name": name}
    if func:
        n["func"] = func
    return n

def flow(path, nodes):
    json.dump(nodes, open(path, "w", encoding="utf-8"))

OWN_FUNC = 'const s = msg.req.headers["x-secret"];\nif (s !== env.get("WEBHOOK_SECRET")) {\n  msg.statusCode = 401;\n  return [null, msg];\n}\nreturn [msg, null];'
LESSON_FUNC = 'let recent = flow.get("recentEventIds") || [];\nconst id = msg.req.headers["x-event-id"];\nif (recent.includes(id)) { return [null, msg]; }\nrecent.unshift(id);\nflow.set("recentEventIds", recent.slice(0, 50));\nreturn [msg, null];'

def build(root):
    sub, old, tmpl = (os.path.join(root, x) for x in ("submissions", "lastyear", "template"))
    for d in (sub, old, tmpl):
        os.makedirs(d)
    t_img = png(TEMPLATE_IMG)
    docx(os.path.join(tmpl, "template.docx"), [TEMPLATE_TXT], [t_img])
    common = [TEMPLATE_TXT]
    anna = ["My smart beehive weighs the honey every night and sends the weight to the beekeeper.",
            "The scale under the hive measures to within ten grams how much honey has been added.",
            "When the weight suddenly drops the beekeeper gets a notification about a possible swarm."]
    S = lambda n: os.path.join(sub, f"{n}_1_2_report.docx")
    docx(S("anna"), common + anna, [png(DIAGRAM), t_img], links=[("https://miro.com/app/board/uXjVAAAA=/", "my board")],
         creator="Anna Peeters", modified_by="Anna Peeters")
    docx(S("bram"), common + unique("bram"), [png(DIAGRAM, fmt="JPEG", resize=(800, 600)), t_img],     # rescaled copy
         links=[("https://www.example.org/sensor?utm_source=chatgpt.com", "source")],
         creator="python-docx", modified_by="Bram Jacobs", created="2026-10-22T20:41:00Z")
    docx(S("carl"), common + unique("carl", 1) + [anna[0].replace("every night", "each night"), anna[1]],
         [t_img], links=[("https://miro.com/app/board/uXjVZZZZ=/", "board")],
         creator="python-docx", modified_by="Someone Else", created="2026-10-22T21:05:00Z")
    docx(S("dirk"), common + unique("dirk"), [t_img], links=[("https://miro.com/app/board/uXjVZZZZ=/", "board"),
                                                       ("https://petlibro.com/feeder", "www.petsafe.com")],
         creator="Dirk", modified_by="Dirk", minutes=2, words=500)
    for n in ("eva", "fien", "gert"):
        docx(S(n), common + unique(n, 2) + SHARED_Q, [t_img], creator=n, modified_by=n)
    for n in ("hans", "ines", "jan"):
        docx(S(n), common + unique(n), [t_img], creator=n, modified_by=n)
    docx(os.path.join(sub, "hans_1_2_feedback.txt"), [])                                     # skipped via exclude
    # Node-RED: dirk imported anna's flow (same ids, same function), bram wrote his by hand,
    # the lesson function in 5 students is course material, jan's export is empty
    copied = [node("3f9a1c22b7e0d451", "tab"), node("8c41e7a0d29b3f16", "http in"),
              node("b2d07f5e91ac4e38", "function", "Secret check", OWN_FUNC), node("e6a93b1d0f7c2a59", "http response")]
    flow(os.path.join(sub, "anna_1_3_flows.json"), copied)
    flow(os.path.join(sub, "dirk_1_3_flows.json"), [dict(n, x=n.get("x", 0) + 40) for n in copied])
    flow(os.path.join(sub, "bram_1_3_flows.json"), [node(f"a5f100000000000{i}", "function", f"f{i}", "return msg;") for i in range(5)])
    for k, n in enumerate(("eva", "fien", "gert", "hans", "ines")):
        flow(os.path.join(sub, f"{n}_1_3_flows.json"), [node(f"{k}1c8e2f09ab7d4{k}6", "function", "Dedup", LESSON_FUNC)])
    open(os.path.join(sub, "jan_1_3_flows.json"), "w").close()
    docx(os.path.join(old, "zoe_9_9_report.docx"), [anna[2], "Last year Zoe also wrote some other sentences about her own beehive."],
         creator="Zoe", modified_by="Zoe")
    return sub, old, os.path.join(tmpl, "template.docx")


# ---------- unit checks ----------

big = io.BytesIO()
Image.open(io.BytesIO(png([("r", [60, 22 + i * 70, 340, 72 + i * 70], c) for i, c in enumerate(["blue", "green", "orange", "purple"])])
                      )).resize((800, 600)).save(big, "JPEG")
a, b = fc.signature(png(DIAGRAM)), fc.signature(big.getvalue())
c = fc.signature(png([("e", [50, 50, 350, 250], "black")]))
assert fc.distance(a, b) <= fc.IMG_DIST, fc.distance(a, b)
assert fc.distance(a, c) > fc.IMG_DIST, fc.distance(a, c)
assert fc.signature(png([], size=(50, 40))) is None

x = fc.sentences("Which features are available depends on the chosen model and the price.")
y = fc.sentences("Which features are available depends on the chosen model and price.")
z = fc.sentences("The cat gets its portion of food at fixed times of the day.")
assert len(fc.shared_sentences(x, y)) == 1 and not fc.shared_sentences(x, z)
assert fc.sentences("Step 2 – What do you see as a user of the app") == fc.sentences("step 2 - what do you see as a user of the app")

assert fc.miro_board("https://miro.com/app/board/uXjVEX_oBas=/?share_link_id=1") == "uXjVEX_oBas="
assert fc.domain("www.PetSafe.com/x") == "petsafe.com"
f = fc.link_flags([("https://miro.com/welcomeonboard/abc", ""), ("https://a.com/?utm_source=chatgpt.com", "source"),
                   ("https://petlibro.com", "www.petsafe.com"), ("https://petsafe.com/x", "petsafe.com")])
assert len(f) == 3 and any("AI chat" in x for x in f) and any("invite" in x for x in f) and any("petlibro" in x for x in f), f

f = fc.student_flags("jansenlotte", {"creator": "python-docx", "lastModifiedBy": "Kim Vos"}, {"the teacher"})
assert any("python-docx" in s for s in f) and any("Kim Vos" in s for s in f), f
assert fc.student_flags("maesbart", {"creator": "The Teacher", "lastModifiedBy": "Maes Bart"}, {"the teacher"}) == []
assert fc.student_flags("maesbart", {"Author": "Course 1 - Assignment 4 template"}, set()) == []
assert fc.tool_family({"Producer": "Skia/PDF m158 Google Docs Renderer"}) == "Google Docs"
assert fc.tool_family({}) == "unknown"

if fc.shutil.which("tesseract"):   # OCR: text in a screenshot is read
    from PIL import ImageFont
    im = Image.new("RGB", (900, 200), "white")
    font = ImageFont.load_default(size=40)
    ImageDraw.Draw(im).text((20, 60), "Message ChatGPT   Copy code", fill="black", font=font)
    buf = io.BytesIO(); im.save(buf, "PNG")
    assert "chatgpt" in fc.ocr(buf.getvalue()).lower(), fc.ocr(buf.getvalue())


# ---------- full run on the fake class ----------

with tempfile.TemporaryDirectory() as root:
    sub, old, tmpl = build(root)
    hist = {"bram": [{"assignment": "earlier", "date": "2026-09-01", "tool": "Word", "names": ["Bram Jacobs"]}]}
    res = fc.analyse(sub, tmpl, None, ["the teacher"], old=[old], use_ocr=False, use_embed=False,
                     use_dolos=False, history=hist, exclude=["*_feedback.txt"])
    P = {frozenset((p["a"], p["b"])): p for p in res["pairs"]}
    why = lambda a, b: fc.why(P[frozenset((a, b))]) if frozenset((a, b)) in P else ""
    flags = lambda s: " | ".join(res["data"][s]["flags"])

    assert "matching image" in why("anna", "bram"), why("anna", "bram")
    assert "both made with python-docx" in why("bram", "carl") and "24 min apart" in why("bram", "carl"), why("bram", "carl")
    assert "shared sentence" in why("anna", "carl"), why("anna", "carl")
    assert "same Miro board" in why("carl", "dirk"), why("carl", "dirk")
    assert why("anna", "lastyear/zoe").startswith("earlier year:"), why("anna", "lastyear/zoe")
    assert any(set(c["students"]) == {"eva", "fien", "gert"} and c["kind"] == "sentence" for c in res["clusters"]), res["clusters"]
    for a, b in (("hans", "ines"), ("hans", "jan"), ("ines", "jan"), ("anna", "hans")):
        assert frozenset((a, b)) not in P, (a, b, why(a, b))       # template text and image don't count
    assert res["common"]["images"] == 0 and res["common"]["sentences"] == 0   # template already removed via --template
    assert res["data"]["hans"]["files"] == ["hans_1_2_report.docx", "hans_1_3_flows.json"]   # --exclude works
    assert "4 identical Node-RED node ids" in why("anna", "dirk") and "identical function node(s): Secret check" in why("anna", "dirk"), why("anna", "dirk")
    assert "not random" in flags("bram") and "not random" not in flags("anna"), (flags("bram"), flags("anna"))
    assert len(res["data"]["anna"]["code"]) == 1 and not res["data"]["jan"]["code"]   # function code goes to Dolos
    assert not fc.handmade_ids({"3f9a1c22b7e0d451", "8c41e7a0.d29b3f", "b2d07f5e91ac4e38"})
    assert fc.handmade_ids({"ms_tab", "ms_secret", "ms_dedup"})

    assert "AI chat" in flags("bram"), flags("bram")
    assert "different tool than before: now python-docx" in flags("bram"), flags("bram")
    assert "Someone Else" in flags("carl"), flags("carl")
    assert "petlibro.com" in flags("dirk") and "pasted" in flags("dirk"), flags("dirk")
    assert not flags("hans") and not flags("anna"), (flags("hans"), flags("anna"))
    assert [h["tool"] for h in hist["bram"]] == ["Word", "python-docx"]         # history updated

    # without --template: whatever a large part of the class shares is dropped as course material
    res2 = fc.analyse(sub, None, None, [], use_ocr=False, use_embed=False, use_dolos=False)
    assert res2["common"]["images"] >= 1 and res2["common"]["sentences"] >= 1, res2["common"]
    assert not any({"hans", "ines"} == {p["a"], p["b"]} for p in res2["pairs"])

    out = os.path.join(root, "r.html")
    fc.report(res, out, "test")
    fc.to_json(res, os.path.join(root, "r.json"))
    j = json.load(open(os.path.join(root, "r.json"), encoding="utf-8"))
    assert j["pairs"] and j["clusters"] and "bram" in j["students"]
    page = open(out, encoding="utf-8").read()
    assert "Suspicious pairs" in page and "anna" in page and "data:image/jpeg;base64" in page

print("ok")
