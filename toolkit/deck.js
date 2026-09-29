/**
 * deck.js — slide builders for one-hour tutoring lessons.
 *
 * Every builder returns the slide and records the geometry of each text box it
 * places into a manifest. build() writes both the .pptx and <deck>.manifest.json;
 * qa_check.py reads the manifest to detect overflow and overlap.
 */
const PptxGenJS = require('pptxgenjs');
const fs = require('fs');
const { spawnSync } = require('child_process');
const path = require('path');
const { COLOR, FONT, SLIDE, GRID, TYPE } = require('./style');

class Deck {
  constructor(meta) {
    this.meta = meta;                 // {topic, student, board, subject, specCode}
    this.pptx = new PptxGenJS();
    this.pptx.defineLayout({ name: 'LESSON', width: SLIDE.w, height: SLIDE.h });
    this.pptx.layout = 'LESSON';
    this.pptx.author = 'Kevin — Tutoring';
    this.pptx.title = `${meta.topic} — ${meta.student}`;
    this.manifest = [];               // [{slide, kind, x,y,w,h, text, fontSize, lines}]
    this.slideNo = 0;
  }

  _record(kind, o, text, fontSize) {
    this.manifest.push({
      slide: this.slideNo, kind,
      x: o.x, y: o.y, w: o.w, h: o.h,
      fontSize: fontSize || 0,
      text: typeof text === 'string' ? text : '',
    });
  }

  // Text box that is registered with the QA manifest.
  _text(slide, content, opts, kind = 'text') {
    slide.addText(content, opts);
    let flat;
    if (Array.isArray(content)) {
      flat = '';
      content.forEach((c, i) => {
        const t = typeof c === 'string' ? c : (c.text || '');
        const brk = typeof c === 'object' && c.options && c.options.breakLine;
        flat += (i && brk ? '\n' : '') + t;
      });
    } else { flat = content; }
    // kind === false: the caller records a truer box itself (e.g. a callout,
    // whose drawn panel is bigger than the text inside it).
    if (kind !== false) this._record(kind, opts, flat, opts.fontSize);
  }

  _newSlide() {
    this.slideNo += 1;
    const s = this.pptx.addSlide();
    s.background = { color: COLOR.paper };
    return s;
  }

  /**
   * The footer is the slide number and nothing else. The old "topic · section"
   * line added no information a student needed and gave content something to
   * collide with at the bottom of the slide. The `label` argument is accepted
   * and ignored so existing builders keep working.
   *
   * Recorded to the manifest, so qa_check.py can see an overlap with it.
   */
  _footer(slide, _label) {
    this._text(slide, String(this.slideNo), {
      ...TYPE.footer, x: SLIDE.w - GRID.marginX - 0.6, y: GRID.footerY, w: 0.6, h: 0.28, align: 'right',
    }, 'footer');
  }

  // Slide title + short gold rule. Deliberately a rule, not a full-height stripe.
  _title(slide, title, eyebrow) {
    if (eyebrow) {
      this._text(slide, eyebrow.toUpperCase(), {
        x: GRID.marginX, y: GRID.titleY - 0.30, w: GRID.contentW, h: 0.26,
        fontFace: FONT.head, fontSize: 12, bold: true, color: COLOR.gold, charSpacing: 1.6,
      }, 'eyebrow');
    }
    this._text(slide, title, {
      ...TYPE.slideTitle, x: GRID.marginX, y: GRID.titleY, w: GRID.contentW, h: GRID.titleH,
      valign: 'top',
    }, 'title');
    slide.addShape(this.pptx.ShapeType.rect, {
      x: GRID.marginX, y: GRID.ruleY, w: 1.15, h: 0.045, fill: { color: COLOR.gold }, line: { width: 0 },
    });
  }

  /* ---------------------------------------------------------------- slides */

  titleSlide() {
    const s = this._newSlide();
    s.background = { color: COLOR.navy };
    const m = this.meta;
    this._text(s, m.topic, {
      ...TYPE.deckTitle, x: GRID.marginX, y: 2.30, w: GRID.contentW * 0.82, h: 1.60, valign: 'bottom',
    }, 'title');
    s.addShape(this.pptx.ShapeType.rect, {
      x: GRID.marginX, y: 4.10, w: 1.6, h: 0.05, fill: { color: COLOR.gold }, line: { width: 0 },
    });
    this._text(s, `${m.board} ${m.subject}${m.specCode ? '  ·  ' + m.specCode : ''}`, {
      fontFace: FONT.head, fontSize: 19, color: COLOR.gold,
      x: GRID.marginX, y: 4.36, w: GRID.contentW * 0.82, h: 0.42,
    });
    this._text(s, `Prepared for ${m.student}${m.date ? '   ·   ' + m.date : ''}`, {
      fontFace: FONT.body, fontSize: 15, color: 'B9C6D4',
      x: GRID.marginX, y: 4.84, w: GRID.contentW * 0.82, h: 0.38,
    });
    return s;
  }

  /**
   * The lesson plan. No minutes: Kevin does not want per-section timings on the
   * slide, ever. A `mins` field on an item is ignored, and there is no option
   * to switch them back on. Plan the hour, don't print it.
   */
  roadmapSlide(items) {
    const s = this._newSlide();
    this._title(s, 'What we\u2019re doing today', 'Roadmap');
    const rows = [[
      { text: '', options: { fill: { color: COLOR.navy } } },
      { text: 'Section', options: { ...TYPE.tableHead, fill: { color: COLOR.navy } } },
      { text: 'Focus', options: { ...TYPE.tableHead, fill: { color: COLOR.navy } } },
    ]];
    items.forEach((it, i) => {
      const band = i % 2 ? COLOR.navySoft : COLOR.paper;
      rows.push([
        { text: String(i + 1), options: { ...TYPE.tableCell, bold: true, color: COLOR.gold, fill: { color: band }, align: 'center' } },
        { text: it.title, options: { ...TYPE.tableCell, bold: true, color: COLOR.navy, fill: { color: band } } },
        { text: it.focus || '', options: { ...TYPE.tableCell, fill: { color: band } } },
      ]);
    });
    const opts = {
      x: GRID.marginX, y: GRID.bodyTop, w: GRID.contentW,
      colW: [0.55, 3.7, GRID.contentW - 0.55 - 3.7],
      border: { type: 'solid', color: COLOR.line, pt: 0.75 },
      rowH: 0.44, valign: 'middle', margin: 0.09,
    };
    const geo = { x: opts.x, y: opts.y, w: opts.w, h: rows.length * 0.44 };
    this._record('table', geo, '', TYPE.tableCell.fontSize);
    s.addTable(rows, opts);
    this._footer(s, 'Roadmap');
    return s;
  }

  /**
   * The specification page, with short explanations arrowed to the exact lines
   * they refer to.
   *
   * `notes` entries are {text, anchor}, where anchor is a label recorded by
   * spec_capture.py --anchor. Source details go in the speaker notes, not on
   * the slide face.
   */
  specSlide({ imagePath, imgW, imgH, anchors, notes, title, eyebrow, source }) {
    const s = this._newSlide();
    this._title(s, title, eyebrow || 'Specification');

    const top = GRID.bodyTop;
    const availH = GRID.bodyBottom - top;
    const maxImgW = GRID.contentW * 0.46;
    const ar = imgW / imgH;
    let ih = availH, iw = ih * ar;
    if (iw > maxImgW) { iw = maxImgW; ih = iw / ar; }
    const ix = GRID.marginX, iy = top + (availH - ih) / 2;

    s.addShape(this.pptx.ShapeType.rect, {
      x: ix - 0.04, y: iy - 0.04, w: iw + 0.08, h: ih + 0.08,
      fill: { color: COLOR.paper }, line: { color: COLOR.line, width: 1 },
    });
    s.addImage({ path: imagePath, x: ix, y: iy, w: iw, h: ih });
    this._record('image', { x: ix, y: iy, w: iw, h: ih }, '', 0);

    // Note cards down the right, ordered to follow the page top-to-bottom so
    // the arrows do not cross each other.
    const cx = ix + iw + 1.05;
    const cw = SLIDE.w - GRID.marginX - cx;
    const ordered = [...notes].sort((a, b) => {
      const ya = anchors[a.anchor] ? anchors[a.anchor].y : 0;
      const yb = anchors[b.anchor] ? anchors[b.anchor].y : 0;
      return ya - yb;
    });
    const gap = 0.20;
    const ch = (availH - gap * (ordered.length - 1)) / ordered.length;

    ordered.forEach((n, i) => {
      const cy = top + i * (ch + gap);
      s.addShape(this.pptx.ShapeType.rect, {
        x: cx, y: cy, w: cw, h: ch, fill: { color: COLOR.navySoft }, line: { width: 0 },
      });
      s.addShape(this.pptx.ShapeType.rect, {
        x: cx, y: cy, w: 0.05, h: ch, fill: { color: COLOR.gold }, line: { width: 0 },
      });
      this._text(s, n.text, {
        ...TYPE.bodySm, x: cx + 0.22, y: cy + 0.10, w: cw - 0.44, h: ch - 0.20,
        valign: 'middle', lineSpacingMultiple: 1.12,
      }, 'note');

      const a = anchors[n.anchor];
      if (!a) return;
      const tx = ix + Math.min(a.x_end + 0.012, 0.995) * iw;
      const ty = iy + a.y * ih;
      const sx = cx - 0.10, sy = cy + ch / 2;
      s.addShape(this.pptx.ShapeType.line, {
        x: Math.min(sx, tx), y: Math.min(sy, ty),
        w: Math.abs(sx - tx), h: Math.abs(sy - ty),
        flipH: tx < sx, flipV: ty < sy,
        line: { color: COLOR.gold, width: 1.5, endArrowType: 'triangle' },
      });
      s.addShape(this.pptx.ShapeType.ellipse, {
        x: tx - 0.045, y: ty - 0.045, w: 0.09, h: 0.09,
        fill: { color: COLOR.gold }, line: { width: 0 },
      });
    });

    if (source) s.addNotes(source);
    this._footer(s, 'Specification');
    return s;
  }


  /**
   * Concept-teaching slide. Any of lead / bullets / table / callout may be used.
   */
  conceptSlide({ title, eyebrow, lead, bullets, table, callout, twoCol, footerLabel }) {
    const s = this._newSlide();
    this._title(s, title, eyebrow || 'Concept');
    let y = GRID.bodyTop;
    const W = GRID.contentW;

    if (lead) {
      const h = 0.34 * Math.max(1, Math.ceil(lead.length / 95));
      this._text(s, lead, { ...TYPE.subhead, x: GRID.marginX, y, w: W, h, valign: 'top' }, 'lead');
      y += h + 0.24;
    }

    if (bullets && bullets.length) {
      const cols = twoCol ? 2 : 1;
      const colW = cols === 2 ? (W - 0.5) / 2 : W;
      const per = Math.ceil(bullets.length / cols);
      const lineH = 0.44;
      let maxH = 0;
      for (let c = 0; c < cols; c++) {
        const chunk = bullets.slice(c * per, (c + 1) * per);
        if (!chunk.length) continue;
        const h = chunk.reduce((a, b) =>
          a + lineH * Math.max(1, Math.ceil(String(b).length / (cols === 2 ? 48 : 100))), 0);
        maxH = Math.max(maxH, h);
        this._text(s, chunk.map(t => ({ text: t, options: { bullet: { code: '2022' }, breakLine: true } })), {
          ...TYPE.body, x: GRID.marginX + c * (colW + 0.5), y, w: colW, h, valign: 'top', lineSpacingMultiple: 1.15,
        }, 'bullets');
      }
      y += maxH + 0.24;
    }

    if (table) {
      const rows = [table.head.map(hd => ({ text: hd, options: { ...TYPE.tableHead, fill: { color: COLOR.navy } } }))];
      table.rows.forEach((r, i) => {
        const band = i % 2 ? COLOR.navySoft : COLOR.paper;
        rows.push(r.map((c, ci) => ({
          text: String(c),
          options: { ...TYPE.tableCell, fill: { color: band }, bold: ci === 0, color: ci === 0 ? COLOR.navy : COLOR.ink },
        })));
      });
      const rowH = table.rowH || 0.42;
      const opts = {
        x: GRID.marginX, y, w: W, colW: table.colW,
        border: { type: 'solid', color: COLOR.line, pt: 0.75 },
        rowH, valign: 'middle', margin: 0.09,
      };
      this._record('table', { x: opts.x, y: opts.y, w: opts.w, h: rows.length * rowH }, '', TYPE.tableCell.fontSize);
      s.addTable(rows, opts);
      y += rows.length * rowH + 0.22;
    }

    if (callout) {
      // Clamp the panel to the space left above the content baseline. Without
      // this a long callout runs over the footer, which is what it used to do.
      const want = 0.30 + 0.30 * Math.max(1, Math.ceil(callout.text.length / 105));
      const h = Math.min(want, GRID.bodyBottom - y);
      s.addShape(this.pptx.ShapeType.rect, {
        x: GRID.marginX, y, w: W, h, fill: { color: callout.warn ? 'FBEDED' : COLOR.goldSoft }, line: { width: 0 },
      });
      s.addShape(this.pptx.ShapeType.rect, {
        x: GRID.marginX, y, w: 0.06, h, fill: { color: callout.warn ? COLOR.red : COLOR.gold }, line: { width: 0 },
      });
      this._text(s, [
        { text: (callout.label || (callout.warn ? 'Watch out' : 'Key idea')) + '  ', options: { bold: true, color: callout.warn ? COLOR.red : COLOR.navy } },
        { text: callout.text, options: { color: COLOR.ink } },
      ], { ...TYPE.bodySm, x: GRID.marginX + 0.24, y: y + 0.13, w: W - 0.48, h: h - 0.26, valign: 'top' }, 'callout-text');
      // Also record the drawn panel. The text box alone is not what collides
      // with the footer — the panel around it is, and it is taller.
      this._record('callout', { x: GRID.marginX, y, w: W, h }, '', 0);
      y += h + 0.2;
    }
    this._footer(s, footerLabel || 'Teaching');
    return s;
  }

  /**
   * A real exam question, and optionally its official mark scheme, as
   * screenshots from the paper itself. Students see it came from a real paper
   * and learn to read a mark scheme in its actual form.
   *
   * Images are unedited crops of the PDF. Nothing is retyped.
   */
  paperSlide({ title, eyebrow, images, citation, footerLabel }) {
    const s = this._newSlide();
    this._title(s, title, eyebrow);
    const W = GRID.contentW;
    const top = GRID.bodyTop + 0.24;
    const avail = GRID.bodyBottom - top - 0.40;
    const gap = 0.34, labelH = 0.24;

    // Give each image height in proportion to how tall it wants to be at full
    // width, so a short mark-scheme strip doesn't get the same band as a tall
    // question. Then scale everything to fit the available height.
    const want = images.map(im => Math.min(W / (im.w / im.h), avail));
    const totalLabels = images.length * labelH + (images.length - 1) * gap;
    const scale = Math.min(1, (avail - totalLabels) / want.reduce((a, b) => a + b, 0));

    let y = top;
    images.forEach(im => {
      const h = want[images.indexOf(im)] * scale;
      const w = h * (im.w / im.h);
      const x = GRID.marginX + (W - w) / 2;
      this._text(s, im.label.toUpperCase(), {
        fontFace: FONT.head, fontSize: 10, bold: true,
        color: im.markScheme ? COLOR.green : COLOR.muted, charSpacing: 1.2,
        x: GRID.marginX, y, w: W, h: labelH,
      }, 'eyebrow');
      y += labelH;
      s.addShape(this.pptx.ShapeType.rect, {
        x: x - 0.04, y: y - 0.04, w: w + 0.08, h: h + 0.08,
        fill: { color: COLOR.paper },
        line: { color: im.markScheme ? COLOR.green : COLOR.line, width: 1 },
      });
      s.addImage({ path: im.path, x, y, w, h });
      this._record('image', { x, y, w, h }, '', 0);
      y += h + gap;
    });

    // Provenance goes in the speaker notes, never on the slide. A caption under
    // an exam-board screenshot reads as filler; Kevin has rejected it twice.
    if (citation) s.addNotes(citation);
    this._footer(s, footerLabel || 'Past paper');
    return s;
  }

  /** Practice questions — no answers on the slide, for live use. */
  practiceSlide({ n, title, instruction, questions }) {
    const s = this._newSlide();
    this._title(s, title || `Your turn ${n || ''}`.trim(), 'Practice');
    let y = GRID.bodyTop;
    if (instruction) {
      this._text(s, instruction, { ...TYPE.subhead, x: GRID.marginX, y, w: GRID.contentW, h: 0.36 }, 'lead');
      y += 0.5;
    }
    const avail = GRID.bodyBottom - y - 0.1;
    const each = Math.min(1.5, avail / questions.length);
    questions.forEach((q, i) => {
      const qy = y + i * each;
      this._text(s, String(i + 1), {
        fontFace: FONT.head, fontSize: 22, bold: true, color: COLOR.gold,
        x: GRID.marginX, y: qy, w: 0.45, h: 0.4,
      }, 'qnum');
      this._text(s, q.text, {
        ...TYPE.body, x: GRID.marginX + 0.5, y: qy, w: GRID.contentW - 0.5 - 1.1, h: each - 0.12,
        valign: 'top', lineSpacingMultiple: 1.1,
      }, 'question');
      this._text(s, `[${q.marks}]`, {
        ...TYPE.bodySm, bold: true, color: COLOR.muted,
        x: SLIDE.w - GRID.marginX - 1.0, y: qy, w: 1.0, h: 0.32, align: 'right',
      });
      if (i < questions.length - 1) {
        s.addShape(this.pptx.ShapeType.rect, {
          x: GRID.marginX + 0.5, y: qy + each - 0.10, w: GRID.contentW - 0.5, h: 0.012,
          fill: { color: COLOR.line }, line: { width: 0 },
        });
      }
    });
    this._footer(s, 'Practice — no answers shown');
    return s;
  }

  /**
   * The answers to the practice set immediately before it. Rule 5: every
   * practice set is followed by its answers slide, so Kevin can reveal them
   * live rather than improvising.
   *
   * For a practice question taken from a real paper, do NOT use this — show the
   * official mark scheme with paperSlide({markScheme: true}) instead. This is
   * only for questions written for the lesson.
   *
   * answers: [{q?, answer, marks}] — `q` defaults to its 1-based position.
   */
  answersSlide({ title, note, answers, colW }) {
    return this.conceptSlide({
      title: title || 'Your turn — model answers',
      eyebrow: 'Answers',
      footerLabel: 'Answers',
      table: {
        head: ['Q', 'Answer the examiner is looking for', 'Marks'],
        colW: colW || [0.6, GRID.contentW - 0.6 - 0.91, 0.91],
        rows: answers.map((a, i) => [String(a.q || i + 1), a.answer, String(a.marks)]),
        rowH: 0.92,
      },
      callout: note ? { text: note } : undefined,
    });
  }

  plenarySlide({ recap, exitQuestions }) {
    const s = this._newSlide();
    this._title(s, 'Recap — what should stick', 'Plenary');
    const colW = (GRID.contentW - 0.5) / 2;
    this._text(s, 'The three things from today', { ...TYPE.subhead, x: GRID.marginX, y: GRID.bodyTop, w: colW, h: 0.34 }, 'lead');
    this._text(s, recap.map(t => ({ text: t, options: { bullet: { code: '2022' }, breakLine: true } })), {
      ...TYPE.body, x: GRID.marginX, y: GRID.bodyTop + 0.46, w: colW, h: 3.4, valign: 'top', lineSpacingMultiple: 1.2,
    }, 'bullets');
    const ax = GRID.marginX + colW + 0.5;
    s.addShape(this.pptx.ShapeType.rect, {
      x: ax, y: GRID.bodyTop, w: colW, h: 3.86, fill: { color: COLOR.navySoft }, line: { width: 0 },
    });
    this._text(s, 'Exit questions — answer out loud', { ...TYPE.subhead, x: ax + 0.24, y: GRID.bodyTop + 0.2, w: colW - 0.48, h: 0.34 }, 'lead');
    this._text(s, exitQuestions.map((t, i) => ({ text: `${i + 1}.  ${t}`, options: { breakLine: true } })), {
      ...TYPE.bodySm, x: ax + 0.24, y: GRID.bodyTop + 0.66, w: colW - 0.48, h: 3.0, valign: 'top', lineSpacingMultiple: 1.3,
    }, 'bullets');
    this._footer(s, 'Plenary');
    return s;
  }

  /**
   * Homework — the final slide. Links only; no files are downloaded.
   * Pass `items: [{label, url, note}]`. Mark-scheme links must NEVER be listed
   * here — the student sees this slide. They belong in the homework tracker.
   */
  homeworkSlide({ source, items, notes, dueBy, title, url }) {
    // back-compat: a single {title,url} still works
    if (!items) items = [{ label: title, url }];
    const s = this._newSlide();
    this._title(s, 'Homework', 'Before next session');
    const W = GRID.contentW;
    let y = GRID.bodyTop;

    if (source) {
      this._text(s, `Source: ${source}`, {
        ...TYPE.bodySm, color: COLOR.muted, x: GRID.marginX, y, w: W, h: 0.3,
      });
      y += 0.42;
    }

    items.forEach((it, i) => {
      // Worksheet URLs are often long and %20-escaped. Give them a smaller face
      // and a second line rather than letting them run out of the card.
      const linkFs = it.url.length > 76 ? 10 : 12;
      const linkH = it.url.length > 76 ? 0.46 : 0.30;
      const grow = linkH - 0.30;
      const h = (it.note ? 1.16 : 0.92) + grow;
      s.addShape(this.pptx.ShapeType.rect, {
        x: GRID.marginX, y, w: W, h, fill: { color: COLOR.navySoft }, line: { width: 0 },
      });
      s.addShape(this.pptx.ShapeType.rect, {
        x: GRID.marginX, y, w: 0.06, h, fill: { color: COLOR.gold }, line: { width: 0 },
      });
      this._text(s, it.label, {
        ...TYPE.subhead, fontSize: 18, x: GRID.marginX + 0.28, y: y + 0.11, w: W - 0.56, h: 0.34,
      }, 'lead');
      this._text(s, [{
        text: it.url,
        options: { hyperlink: { url: it.url, tooltip: it.label }, color: COLOR.navyMid, underline: true },
      }], {
        fontFace: FONT.body, fontSize: linkFs, x: GRID.marginX + 0.28, y: y + 0.47,
        w: W - 0.56, h: linkH, valign: 'top',
      }, 'link');
      if (it.note) {
        this._text(s, it.note, {
          ...TYPE.caption, x: GRID.marginX + 0.28, y: y + 0.79 + grow, w: W - 0.56, h: 0.30, valign: 'top',
        }, 'citation');
      }
      y += h + 0.16;
    });

    if (notes && notes.length) {
      y += 0.06;
      const h = Math.min(GRID.bodyBottom - y - (dueBy ? 0.46 : 0), 1.5);
      this._text(s, notes.map(t => ({ text: t, options: { bullet: { code: '2022' }, breakLine: true } })), {
        ...TYPE.bodySm, x: GRID.marginX, y, w: W, h, valign: 'top', lineSpacingMultiple: 1.15,
      }, 'bullets');
      y += h + 0.08;
    }
    if (dueBy) {
      this._text(s, `Due: ${dueBy}`, {
        ...TYPE.subhead, fontSize: 17, color: COLOR.gold,
        x: GRID.marginX, y: Math.min(y, GRID.bodyBottom - 0.4), w: W, h: 0.36,
      }, 'lead');
    }
    this._footer(s, 'Homework');
    return s;
  }

  /* ----------------------------------------------------------------- write */
  async build(outPath) {
    fs.mkdirSync(path.dirname(outPath), { recursive: true });
    await this.pptx.writeFile({ fileName: outPath });
    const mf = outPath.replace(/\.pptx$/, '.manifest.json');
    fs.writeFileSync(mf, JSON.stringify({
      meta: this.meta, slideW: SLIDE.w, slideH: SLIDE.h,
      grid: { bodyBottom: GRID.bodyBottom, marginX: GRID.marginX },
      boxes: this.manifest,
    }, null, 2));
    // QA is a non-negotiable, so it runs here rather than relying on someone
    // remembering the command. A failing deck is never written off as finished.
    const qa = spawnSync('python3', [path.join(__dirname, 'qa_check.py'), mf], { encoding: 'utf8' });
    process.stdout.write(qa.stdout || '');
    process.stderr.write(qa.stderr || '');
    if (qa.status !== 0) {
      throw new Error(`qa_check.py failed for ${mf} — fix build.js and rebuild.`);
    }

    return { pptx: outPath, manifest: mf, slides: this.slideNo, qa: 'clean' };
  }
}

module.exports = { Deck, COLOR, GRID, SLIDE, TYPE };
