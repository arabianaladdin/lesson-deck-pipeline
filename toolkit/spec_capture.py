#!/usr/bin/env python3
"""
spec_capture.py - download an exam-board specification PDF, find the pages that
cover a topic, and capture a REAL image of them (never a retyped version).

Uses PyMuPDF, so no pdftoppm / poppler / LibreOffice needed.

  # 1. what's in the spec, and where?
  python3 spec_capture.py find spec.pdf "periodic table" "group trends"

  # 2. capture page 34, auto-cropped to the matching section
  python3 spec_capture.py capture spec.pdf out/spec_p34.png --page 34 \
      --crop-to "periodic table" --board "Cambridge IGCSE" --code "0620" \
      --title "Chemistry" --version "for examination 2027-2029"

Writes <out>.png and <out>.citation.json next to it.
"""
import argparse, json, os, re, subprocess, sys

try:
    import fitz  # PyMuPDF
except ImportError:
    sys.exit("PyMuPDF missing.  python3 -m pip install --user pymupdf")

DPI = 200


def fetch(url, dest):
    """Download a spec PDF. Exam boards often reject non-browser user agents."""
    os.makedirs(os.path.dirname(dest) or ".", exist_ok=True)
    ua = ("Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 "
          "(KHTML, like Gecko) Chrome/125.0 Safari/537.36")
    subprocess.run(["curl", "-fL", "--compressed", "-A", ua, "-o", dest, url], check=True)
    with open(dest, "rb") as fh:
        if fh.read(5) != b"%PDF-":
            sys.exit(f"{dest} is not a PDF - the board probably gated it behind a page. "
                     "Use a headless browser to reach the real PDF link.")
    return dest


def cmd_find(args):
    doc = fitz.open(args.pdf)
    hits = {}
    for pno in range(len(doc)):
        page = doc[pno]
        text = page.get_text("text")
        low = text.lower()
        score = sum(low.count(t.lower()) for t in args.terms)
        if score:
            # a short context window around the first match, to eyeball relevance
            first = min((low.find(t.lower()) for t in args.terms if t.lower() in low))
            ctx = re.sub(r"\s+", " ", text[max(0, first - 90):first + 260]).strip()
            hits[pno + 1] = (score, ctx)
    if not hits:
        print("No pages matched.")
        return
    for pno, (score, ctx) in sorted(hits.items(), key=lambda kv: -kv[1][0])[:args.top]:
        print(f"\n--- page {pno}  (score {score}) ---\n{ctx}")


def detect_version(doc):
    """Best-effort pull of the 'for examination in ...' line from the cover."""
    head = "\n".join(doc[p].get_text("text") for p in range(min(3, len(doc))))
    for pat in (r"[Ff]or examination(?: in)? [^\n]{0,60}",
                r"Version\s*[\d.]+[^\n]{0,40}",
                r"First teaching\s*[^\n]{0,40}"):
        m = re.search(pat, head)
        if m:
            return re.sub(r"\s+", " ", m.group(0)).strip()
    return ""


def cmd_capture(args):
    doc = fitz.open(args.pdf)
    pno = args.page - 1
    if not 0 <= pno < len(doc):
        sys.exit(f"page {args.page} out of range (PDF has {len(doc)})")
    page = doc[pno]
    clip = None

    if args.crop_to or args.end_before:
        # Mark-scheme pages are often rotated landscape. search_for can return
        # coordinates in unrotated space, which fall outside page.rect and
        # produce an empty crop. Map anything out-of-page into display space.
        def norm(rs):
            out = []
            for r in rs:
                if not page.rect.contains(r):
                    r = r * page.rotation_matrix
                out.append(r)
            return out

        rects = []
        for term in (args.crop_to or []):
            rects += norm(page.search_for(term))
        if rects:
            top = min(r.y0 for r in rects)
            bot = max(r.y1 for r in rects)
            # --end-before stops the crop just above a later phrase, so a
            # question or mark-scheme row ends cleanly at the next part.
            if args.end_before:
                stops = []
                for term in args.end_before:
                    stops += [r.y0 for r in norm(page.search_for(term)) if r.y0 > top]
                if stops:
                    bot = min(stops)
                else:
                    print(f"warn: --end-before {args.end_before} not found below the "
                          "start anchor; using the full match extent", file=sys.stderr)
            pad_top, pad_bot = args.pad_top, args.pad_bottom
            pr = page.rect
            clip = fitz.Rect(pr.x0 + args.pad_x,
                             max(pr.y0, top - pad_top),
                             pr.x1 - args.pad_x,
                             min(pr.y1, bot + pad_bot))
        else:
            print(f"warn: none of {args.crop_to} found on page {args.page}; "
                  "capturing the whole page", file=sys.stderr)

    anchors = {}
    if args.anchor:
        box = clip if clip is not None else page.rect
        for spec in args.anchor:
            if "=" not in spec:
                sys.exit(f"--anchor needs label=phrase, got {spec!r}")
            label, phrase = spec.split("=", 1)
            found = page.search_for(phrase)
            if not found:
                print(f"warn: anchor {label!r} phrase {phrase!r} not found on page "
                      f"{args.page}", file=sys.stderr)
                continue
            r = found[0]
            if not page.rect.contains(r):
                r = r * page.rotation_matrix
            anchors[label] = {
                "x": round((r.x0 - box.x0) / box.width, 4),
                "y": round(((r.y0 + r.y1) / 2 - box.y0) / box.height, 4),
                "x_end": round((r.x1 - box.x0) / box.width, 4),
                "phrase": phrase,
            }

    if clip is not None and (clip.height < 5 or clip.width < 5):
        sys.exit(f"crop came out empty ({clip.width:.0f}x{clip.height:.0f}pt) - "
                 "check the --crop-to / --end-before anchors against the page text")
    pix = page.get_pixmap(dpi=DPI, clip=clip)
    os.makedirs(os.path.dirname(args.out) or ".", exist_ok=True)
    pix.save(args.out)

    cite = {
        "board": args.board, "code": args.code, "title": args.title,
        "page": args.page, "version": args.version or detect_version(doc),
        "source_pdf": os.path.basename(args.pdf), "source_url": args.url or "",
        "width_px": pix.width, "height_px": pix.height,
        "cropped": clip is not None,
        "anchors": anchors,
    }
    bits = [b for b in (cite["board"], cite["title"], cite["code"]) if b]
    cite["line"] = (f"Source: {' '.join(bits)} syllabus, page {args.page}"
                    + (f", {cite['version']}" if cite["version"] else "")
                    + ". Reproduced for teaching purposes.")
    with open(args.out + ".citation.json", "w") as fh:
        json.dump(cite, fh, indent=2)
    print(json.dumps(cite, indent=2))


def main():
    ap = argparse.ArgumentParser(description=__doc__,
                                 formatter_class=argparse.RawDescriptionHelpFormatter)
    sub = ap.add_subparsers(dest="cmd", required=True)

    f = sub.add_parser("find", help="list spec pages matching topic terms")
    f.add_argument("pdf"); f.add_argument("terms", nargs="+")
    f.add_argument("--top", type=int, default=6)
    f.set_defaults(func=cmd_find)

    d = sub.add_parser("fetch", help="download a spec PDF")
    d.add_argument("url"); d.add_argument("dest")
    d.set_defaults(func=lambda a: print(fetch(a.url, a.dest)))

    c = sub.add_parser("capture", help="render a spec page to PNG")
    c.add_argument("pdf"); c.add_argument("out")
    c.add_argument("--page", type=int, required=True)
    c.add_argument("--crop-to", nargs="*", default=None,
                   help="crop tightly around these phrases on the page")
    c.add_argument("--anchor", nargs="*", default=None, metavar="LABEL=PHRASE",
                   help="record where PHRASE sits inside the image (0-1 fractions), "
                        "so a slide can draw an arrow to it")
    c.add_argument("--end-before", nargs="*", default=None,
                   help="end the crop just above this phrase (e.g. the next question part)")
    c.add_argument("--pad-top", type=float, default=34)
    c.add_argument("--pad-bottom", type=float, default=34)
    c.add_argument("--pad-x", type=float, default=6)
    c.add_argument("--board", default=""); c.add_argument("--code", default="")
    c.add_argument("--title", default=""); c.add_argument("--version", default="")
    c.add_argument("--url", default="")
    c.set_defaults(func=cmd_capture)

    a = ap.parse_args()
    a.func(a)


if __name__ == "__main__":
    main()
