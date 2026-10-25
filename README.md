# fraudecheck

A local tool for teachers that looks for copied or AI-generated work in a batch of student submissions
(a Canvas export: PDF, Word, images, code). It combines image matching, shared and paraphrased sentences,
[JPlag](https://github.com/jplag/JPlag) (text), [Dolos](https://github.com/dodona-edu/dolos) (code),
document metadata and traces in links, and writes one HTML report plus JSON.

Everything runs on your own machine: no student work is sent to an external service.

## What it finds

**Per pair of students**

| Signal | How | Score |
|---|---|---|
| Same image | md5 or 16×16 grayscale (also rescaled/JPEG), including images from PDFs embedded in a docx | 50 |
| Same Miro board | board id from the links | 40 |
| Shared sentences | word 3-grams, also when lightly reworded | 5/sentence (max 30) |
| Paraphrased sentences | sentence embeddings (fastembed, multilingual), from 2 sentences | 4/sentence (max 20) |
| JPlag | text mode, with the template as base code | 50 × % |
| Dolos | code, similarity across the whole project, only outliers relative to the class | 60 × % |
| Same suspicious generator | python-docx, headless browser, ... (+10 if created within 60 min) | 15 |

**Per student** (column *Notable*)

- link copied from an AI chat (`utm_source=chatgpt.com`, `chatgpt.com/`, `claude.ai/`, ...)
- Miro link is a team invite instead of a board link
- the visible link points to a different site than the actual link
- made with a suspicious generator; *last modified by* is someone else
- many words in little editing time (Word keeps track of this: pasted?)
- an image looks like a screenshot of an AI chat (OCR)
- a different tool or name in the metadata than in earlier assignments (history)

**Clusters**: the same piece (image or sentence) in 3+ students.
Anything that appears in ≥ 15% of the class (min. 4 students) is treated as course material and ignored;
the report shows how much was dropped that way.

## Installation

```
pip install -r requirements.txt
```

You also need **poppler** (`pdftotext`, `pdfimages`, `pdfinfo`, `pdftohtml`), e.g. `scoop install poppler`,
`brew install poppler` or `apt install poppler-utils`.

Optional (skipped with a message if missing):

| Component | Installation |
|---|---|
| JPlag (text) | jar from the [releases](https://github.com/jplag/JPlag/releases) in `~/tools/jplag/` or via `--jplag`, + Java 21+ |
| Dolos (code) | `npm install -g @dodona/dolos` |
| OCR | Tesseract with the `nld` and `eng` language data (change `OCR_LANGS` for other languages) |
| Paraphrases | `fastembed` (in requirements.txt); the first run downloads a ~220 MB model from Hugging Face |

## Usage

```
python fraudecheck.py "<folder with submissions>" \
    --template "<blank assignment or template>" \
    --ignore-author "Teacher Name" \
    --ignore-image school-logo.png \
    --old "<submissions from last year>" \
    -o report.html
```

| Option | |
|---|---|
| `--template` | its text and images are ignored |
| `--old` | folder with submissions from an earlier year (repeatable); only new ↔ old is compared |
| `--ignore-author` | author name that is expected (e.g. the teacher who made the template) |
| `--ignore-image` | image to ignore, e.g. the logo from the report template |
| `--exclude` | file name pattern to skip, e.g. `'*_feedback.txt'` (repeatable) |
| `--jplag` | path to the JPlag jar (default: newest `~/tools/jplag/jplag-*.jar`) |
| `--history` | history file (default `~/.fraudecheck/history.json`); `--no-history` to turn it off |
| `--no-ocr`, `--no-embed` | turn off OCR or embeddings (faster) |

File names follow Canvas: `<student>_<id>_<id>_<original>`; everything before the first `_` is the student.
Unpacked zips (folders) are read recursively; `venv`, `__pycache__`, `.git` and `node_modules` are skipped.

**Output**

- `report.html`: suspicious pairs, clusters, code (Dolos), metadata per student
- `report.json`: the same results, machine-readable
- `report_files/`: `jplag.jplag` and `dolos-<ext>/`; the report shows how to open them
  (`java -jar jplag.jar -M VIEW ...`, `dolos serve ...`)

## Privacy

- Submissions, reports and the history stay on your own computer. The report contains student names and
  images: don't share or commit it (`.gitignore` excludes `*.html` and `*.json`).
- The history (`~/.fraudecheck/history.json`) stores, per student, the assignment name, the tool used and the
  names found in the document metadata. Use `--no-history` if you don't want that.
- Only the embedding model is downloaded (once); nothing is uploaded.

## Interpreting the results

These are **signals, not proof**. An editing time of 0 minutes also happens when a student copies the document
into a new file; two students writing about the same product will naturally write similar sentences.
Always check the originals and talk to the student before drawing conclusions.

## Test

```
python test_fraudecheck.py
```

Builds a fake class (10 students + 1 from an earlier year) and checks every signal. JPlag, Dolos and embeddings
are off in the test; OCR is tested if Tesseract is installed.

## License

MIT, see [LICENSE](LICENSE).
