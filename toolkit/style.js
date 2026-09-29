/**
 * style.js — the single source of truth for lesson-deck visual design.
 * Every deck imports this. Change it here and every future deck inherits it.
 *
 * Design brief: navy/gold, no accent-stripe clichés, real tables (never images
 * of tables), clean typography, generous whitespace.
 */

const COLOR = {
  navy:      '0F2A4A',  // primary — titles, table headers
  navyMid:   '1B3B63',  // secondary navy — subheads
  navySoft:  'E8EDF3',  // pale navy — table banding, panel fills
  gold:      'C8A24A',  // accent — rules, numerals, emphasis
  goldSoft:  'F6EFDC',  // pale gold — callout fills
  ink:       '1A1A1A',  // body text
  muted:     '5A6B7C',  // captions, citations, footers
  paper:     'FFFFFF',
  line:      'C9D2DC',  // hairlines, table borders
  green:     '2E6B4F',  // "correct" / mark-scheme ticks
  red:       'A33A3A',  // misconception warnings
};

const FONT = { head: 'Calibri', body: 'Calibri', mono: 'Consolas' };

// 16:9 at 13.333 x 7.5 inches
const SLIDE = { w: 13.333, h: 7.5 };

// Layout grid — all builders use these so slides align to a common rhythm.
const GRID = {
  marginX:   0.62,
  titleY:    0.42,
  titleH:    0.55,
  ruleY:     1.06,   // gold rule sits under the slide title
  bodyTop:   1.34,   // first line of content
  bodyBottom:6.86,   // content must not cross this
  footerY:   6.98,
  get contentW() { return SLIDE.w - this.marginX * 2; },
  get bodyH()    { return this.bodyBottom - this.bodyTop; },
};

const TYPE = {
  slideTitle: { fontFace: FONT.head, fontSize: 30, bold: true,  color: COLOR.navy },
  deckTitle:  { fontFace: FONT.head, fontSize: 46, bold: true,  color: COLOR.paper },
  subhead:    { fontFace: FONT.head, fontSize: 19, bold: true,  color: COLOR.navyMid },
  body:       { fontFace: FONT.body, fontSize: 17, color: COLOR.ink },
  bodySm:     { fontFace: FONT.body, fontSize: 15, color: COLOR.ink },
  caption:    { fontFace: FONT.body, fontSize: 11, color: COLOR.muted, italic: true },
  footer:     { fontFace: FONT.body, fontSize: 10, color: COLOR.muted },
  tableHead:  { fontFace: FONT.head, fontSize: 14, bold: true, color: COLOR.paper },
  tableCell:  { fontFace: FONT.body, fontSize: 14, color: COLOR.ink },
};

module.exports = { COLOR, FONT, SLIDE, GRID, TYPE };
