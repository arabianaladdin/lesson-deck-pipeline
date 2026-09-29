# Lesson Deck Pipeline

A small toolkit that builds one-hour tutoring lesson decks (`.pptx`) from
exam-board specifications and real past papers — then QA-checks the result
before a human ever sees it.

I built this to run my own GCSE/IGCSE tutoring business: given a student, an
exam board, a subject and a topic, the pipeline produces a fully-styled,
QA-clean slide deck with the actual spec pages and past-paper questions
(screenshotted, never retyped) worked into the teaching.

This repo holds the **build pipeline only** — the code that turns structured
content into a deck and checks it. The exam-board PDFs, spec/paper screenshots
and student data that feed it are private (copyright and privacy reasons) and
live outside this repo.

## How it fits together

```
spec_capture.py  →  crop real spec/past-paper pages into labelled PNGs
      ↓
  build.js        →  a lesson's own script; calls deck.js to assemble slides
      ↓
   deck.js         →  slide builders (title, spec, concept, worked example,
                       practice, answers, plenary, homework) + style.js
      ↓
  <lesson>.pptx  +  <lesson>.manifest.json   (every text box's geometry)
      ↓
  qa_check.py      →  reads the manifest, fails the build on overflow,
                       overlap, off-slide content, or a leaked mark-scheme link
```

`Deck.build()` calls `qa_check.py` itself and throws if it fails — a deck that
doesn't pass QA is never written off as finished.

## What's in this repo

```
toolkit/
  deck.js          Slide builders + the Deck class (Node, pptxgenjs)
  style.js          Shared layout grid, colours, type scale
  spec_capture.py  PDF → cropped, citation-tagged PNG (PyMuPDF)
  qa_check.py      Structural QA on the built deck's manifest
  package.json
```

A real lesson's `build.js` (not included — it references private spec/paper
screenshots) looks roughly like this:

```js
const { Deck } = require('./toolkit/deck');

const deck = new Deck({ topic, student, board, subject, specCode });

deck.titleSlide();
deck.roadmapSlide([{ title: 'Section 1', focus: '...' }, /* ... */]);
deck.specSlide({ imagePath, imgW, imgH, anchors, notes, title, eyebrow, source });
deck.conceptSlide({ title, eyebrow, lead, bullets, table, callout });
deck.paperSlide({ title, eyebrow, images, citation });   // real exam question
deck.paperSlide({ title, eyebrow, images, markScheme: true }); // its mark scheme
deck.practiceSlide({ title, instruction, questions });
deck.answersSlide({ title, answers });
deck.plenarySlide({ recap, exitQuestions });
deck.homeworkSlide({ source, items, dueBy });

deck.build('out/lesson.pptx');   // writes .pptx + .manifest.json, runs QA
```

## Design decisions worth noting

- **Nothing exam-board is retyped.** The spec and every past-paper question or
  mark scheme is a real screenshot, captured and cropped by `spec_capture.py`,
  never hand-typed content presented as if it were the board's.
- **QA reads geometry, not pixels.** Every text box (and any shape a builder
  draws behind text) is recorded as it's placed. `qa_check.py` then checks
  that recorded geometry against the slide grid — it catches a box overflowing
  its bounds or two boxes overlapping without ever rendering the slide.
- **Provenance lives in speaker notes, not on the slide.** A citation or a
  spec reference is metadata for the teacher, not something a student needs
  to read on-screen.
- **A build fails loudly.** `qa_check.py` also flags a mark-scheme link that
  reached a slide, and a `.pptx` that was hand-edited after its own build —
  both are treated as errors, not warnings.

## Stack

- **Node.js** (`pptxgenjs`) for slide assembly
- **Python** (`PyMuPDF` / `fitz`, `Pillow`) for PDF capture and QA
- No LibreOffice, poppler or other native PDF renderer — everything runs on a
  plain Node + Python install

## Running it

```bash
cd toolkit && npm install
python3 -m pip install --user pymupdf pillow
```

`spec_capture.py` and `qa_check.py` are standalone CLIs (see `--help` on
each); `deck.js` is consumed by a lesson's own `build.js`.

---

*The lesson content this pipeline produces — spec screenshots, past-paper
questions, student names — is not in this repo. This is the code, not the
curriculum.*
