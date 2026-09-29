#!/usr/bin/env python3
"""
qa_check.py - geometric QA for generated lesson decks.

Reads the <deck>.manifest.json emitted by deck.js and flags the failures that
actually bite AI-generated pptx: text overflowing its box, content running off
the bottom of the slide, and boxes overlapping each other.

  python3 qa_check.py path/to/deck.manifest.json

Exit 0 = clean, 1 = errors found. Warnings alone do not fail.

Note: this measures geometry against a text-metrics estimate; it does not
render the slides. Calibri is approximated at 0.47em average advance width,
which is close enough to catch real overflow with a little headroom.
"""
import json, os, re, sys, argparse
from collections import defaultdict

AVG_CHAR_EM = 0.47      # average advance width of Calibri, in em
LINE_EM     = 1.22      # line height as a multiple of font size
ERR_RATIO   = 1.15      # estimated/actual height above this = error. PowerPoint
                        # boxes spill rather than clip, so allow real headroom
                        # before calling it a failure.
WARN_RATIO  = 0.98      # ...above this = warning (tight, worth a look)
MIN_OVERLAP = 0.015     # sq inches; ignore hairline touching

# Decorative or self-sizing kinds we don't height-check.
SKIP_FIT = {"image", "table"}


def est_height(text, w_in, fs_pt, bullet=False):
    if not text or not fs_pt:
        return 0.0
    char_w = AVG_CHAR_EM * fs_pt / 72.0
    usable = max(w_in - (0.22 if bullet else 0.0), 0.3)
    cpl = max(int(usable / char_w), 8)
    lines = 0
    # explicit line breaks in bulleted content are already flattened to spaces,
    # so estimate per logical chunk instead
    for chunk in text.split("\n"):
        chunk = chunk.strip()
        if not chunk:
            continue
        lines += max(1, -(-len(chunk) // cpl))
    lines = max(lines, 1)
    return lines * LINE_EM * fs_pt / 72.0


def contains(outer, inner, tol=0.01):
    return (inner["x"] >= outer["x"] - tol and inner["y"] >= outer["y"] - tol
            and inner["x"] + inner["w"] <= outer["x"] + outer["w"] + tol
            and inner["y"] + inner["h"] <= outer["y"] + outer["h"] + tol)


def rect_overlap(a, b):
    x = min(a["x"] + a["w"], b["x"] + b["w"]) - max(a["x"], b["x"])
    y = min(a["y"] + a["h"], b["y"] + b["h"]) - max(a["y"], b["y"])
    return x * y if x > 0 and y > 0 else 0.0


# A mark-scheme link on a slide is a privacy failure, not a cosmetic one: the
# student sees the deck. Question papers are fine, mark schemes never are.
MS_URL = re.compile(r"https?://\S*(?:%20|[-_/.])?(?:ms|mark[-_%20 ]*scheme)s?(?:%20)?[-_. ]*\.pdf", re.I)
MS_WORD = re.compile(r"https?://\S*mark[-_%20 ]*scheme", re.I)

# Keys every lesson's meta.json must carry. Without these a deck has no record
# of which spec version and which papers it was built from.
META_REQUIRED = [
    "topic", "board", "qualification", "subject", "code", "specRefs",
    "specPdf", "specPages", "specVersion", "specUrl", "tier", "slides",
    "minutes", "examples", "homework", "created", "usedBy",
]


def check_privacy(m):
    out = []
    for b in m["boxes"]:
        t = b.get("text") or ""
        if MS_URL.search(t) or MS_WORD.search(t):
            out.append(f"slide {b['slide']} [{b['kind']}] shows a mark-scheme link - "
                       f"students see this deck. Move it to the internal Notion row.")
    return out


def check_meta(manifest_path, slides):
    """meta.json sits beside the manifest and must stay true to the build."""
    d = os.path.dirname(os.path.abspath(manifest_path))
    mp = os.path.join(d, "meta.json")
    if not os.path.exists(mp):
        return [f"no meta.json beside {os.path.basename(manifest_path)}"], []
    try:
        meta = json.load(open(mp))
    except ValueError as e:
        return [f"meta.json is not valid JSON ({e})"], []
    errs, warns = [], []
    missing = [k for k in META_REQUIRED if k not in meta]
    if missing:
        errs.append("meta.json is missing required key(s): " + ", ".join(missing))
    if "slides" in meta and meta["slides"] != slides:
        errs.append(f"meta.json says slides: {meta['slides']}, the deck has {slides}")
    if not meta.get("usedBy"):
        warns.append("meta.json usedBy is empty - no student is recorded on this deck")
    return errs, warns


def check_artefact(manifest_path):
    """The .pptx is a build artefact. If it is newer than the manifest that was
    written alongside it, someone edited the deck by hand."""
    mf = os.path.abspath(manifest_path)
    pptx = mf[:-len(".manifest.json")] + ".pptx"
    d = os.path.dirname(mf)
    build = os.path.join(d, "build.js")
    errs = []
    if os.path.exists(pptx) and os.path.getmtime(pptx) > os.path.getmtime(mf) + 5:
        errs.append(f"{os.path.basename(pptx)} is newer than its manifest - the .pptx "
                    f"looks hand-edited. Edit build.js and rebuild.")
    if os.path.exists(build) and os.path.getmtime(build) > os.path.getmtime(mf) + 5:
        errs.append("build.js has changed since this deck was built - rebuild before handing over.")
    return errs


def check(path):
    with open(path) as fh:
        m = json.load(fh)
    SW, SH = m["slideW"], m["slideH"]
    bottom = m["grid"]["bodyBottom"]
    by_slide = defaultdict(list)
    for b in m["boxes"]:
        by_slide[b["slide"]].append(b)

    errors, warns = [], []
    for sno in sorted(by_slide):
        boxes = by_slide[sno]
        for b in boxes:
            tag = f"slide {sno} [{b['kind']}]"
            # off-slide
            if b["x"] < -0.01 or b["y"] < -0.01 or b["x"] + b["w"] > SW + 0.01 \
               or b["y"] + b["h"] > SH + 0.01:
                errors.append(f"{tag} extends off the slide "
                              f"(x{b['x']:.2f} y{b['y']:.2f} w{b['w']:.2f} h{b['h']:.2f})")
            # past the content baseline
            elif b["y"] + b["h"] > bottom + 0.12 and b["kind"] not in ("footer",):
                # An error, not a warning: crossing the baseline is how content
                # ends up sitting on top of the footer.
                errors.append(f"{tag} runs past the content baseline "
                              f"(ends {b['y'] + b['h']:.2f}\" vs {bottom}\") - "
                              f"shorten it or move it to its own slide")
            # text fit
            if b["kind"] not in SKIP_FIT and b["fontSize"] and b["text"]:
                need = est_height(b["text"], b["w"], b["fontSize"], b["kind"] == "bullets")
                ratio = need / b["h"] if b["h"] else 99
                if ratio > ERR_RATIO:
                    errors.append(f"{tag} text overflows: needs ~{need:.2f}\" in "
                                  f"{b['h']:.2f}\" ({ratio:.0%}) - \"{b['text'][:58]}...\"")
                elif ratio > WARN_RATIO:
                    warns.append(f"{tag} text is tight ({ratio:.0%}) - "
                                 f"\"{b['text'][:48]}...\"")
        # overlaps
        for i in range(len(boxes)):
            for j in range(i + 1, len(boxes)):
                a, c = boxes[i], boxes[j]
                # A box sitting wholly inside another is deliberate - text on a
                # panel, a number in a band. Only partial overlaps are faults.
                if contains(a, c) or contains(c, a):
                    continue
                ov = rect_overlap(a, c)
                if ov > MIN_OVERLAP:
                    errors.append(f"slide {sno} [{a['kind']}] overlaps [{c['kind']}] "
                                  f"by {ov:.2f} sq in")
    return m, errors, warns


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("manifest")
    ap.add_argument("--strict", action="store_true", help="treat warnings as failures")
    a = ap.parse_args()
    m, errors, warns = check(a.manifest)

    name = m["meta"].get("topic", a.manifest)
    slides = max((b["slide"] for b in m["boxes"]), default=0)
    errors += check_privacy(m)
    me, mw = check_meta(a.manifest, slides)
    errors += me
    warns += mw
    errors += check_artefact(a.manifest)
    print(f"QA: {name} - {slides} slides, {len(m['boxes'])} measured boxes")
    for w in warns:
        print(f"  WARN  {w}")
    for e in errors:
        print(f"  ERROR {e}")
    if not errors and not warns:
        print("  clean - no overflow, no overlap, nothing off-slide")
    elif not errors:
        print(f"  passed with {len(warns)} warning(s)")
    else:
        print(f"  FAILED with {len(errors)} error(s)")
    sys.exit(1 if errors or (a.strict and warns) else 0)


if __name__ == "__main__":
    main()
