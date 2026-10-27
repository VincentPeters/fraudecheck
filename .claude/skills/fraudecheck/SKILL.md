---
name: fraudecheck
description: Use when checking a batch of student submissions (a Canvas export of PDF, Word, images, code, Node-RED flows) for copying, shared or paraphrased text, reused images, or AI-generated work. Runs the local fraudecheck tool and summarises its report. Triggers on "/fraudecheck", "check for plagiarism", "did students copy", "fraude check", "check these submissions".
---

# Fraudecheck

Runs [fraudecheck](https://github.com/VincentPeters/fraudecheck) on a folder of
submissions and turns its report into a short, cautious summary for the teacher.

**Core principle: signals, not proof.** Never call a student a cheater. Report
what matched and why, and point at the originals to check.

**Privacy: everything stays local.** Never send submission content, names or
report contents to any external service, web search or MCP tool. Never write the
report into a git repo.

## Steps

### 1. Locate the tool

Use `~/Development/work/fraudecheck/fraudecheck.py`. If it isn't there, clone it:

```bash
git clone https://github.com/VincentPeters/fraudecheck.git ~/Development/work/fraudecheck
pip install -r ~/Development/work/fraudecheck/requirements.txt
```

poppler (`pdftotext`) is required; JPlag, Dolos (`dolos`) and Tesseract are
optional — the tool logs which ones it found and skips the rest. Don't install
missing optional tools unless asked; mention what was skipped.

### 2. Work out the arguments

From the user's request and the folder itself (state assumptions, don't ask
when a default is sensible):

| Argument | When |
|---|---|
| `<folder>` | the unpacked Canvas export (`<student>_<id>_<id>_<file>`) — required |
| `--template <file>` | a blank assignment/template sits next to or inside the folder |
| `--old <folder>` | user mentions last year's submissions (repeatable) |
| `--ignore-author "<name>"` | the teacher's name, if known — otherwise the template author shows up as "edited by someone else" |
| `--ignore-image <img>` | a school/report logo |
| `--exclude '<pattern>'` | e.g. Canvas feedback files |
| `--no-embed` / `--no-ocr` | user wants a quick run |
| `--no-history` | user doesn't want `~/.fraudecheck/history.json` updated |

If the folder is a `.zip`, unpack it to a sibling folder first.

Write the report beside the submissions folder, named after it:
`-o "<parent>/fraudecheck-<folder name>.html"`.

### 3. Run

```bash
python ~/Development/work/fraudecheck/fraudecheck.py "<folder>" [options] -o "<report>.html"
```

The first run with embeddings downloads a ~220 MB model; large classes can take
several minutes — run in the background with a generous timeout. On failure,
show the actual stderr.

### 4. Summarise

Read `<report>.json` (`pairs`, `clusters`, `code`, `students[*].flags`) and
report:

- the top pairs (score ≥ ~40) with their `reasons` in plain words
- clusters of 3+ students
- students with flags (AI-chat links, pasted text, AI-chat screenshots, …)
- what was skipped (JPlag/Dolos/OCR not found, `common` = filtered course material)
- the path to the HTML report, which has the images and side-by-side evidence

Keep it short and neutral. End with the reminder that these are leads to check
against the originals and discuss with the student, not conclusions.
