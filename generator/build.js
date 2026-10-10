// Builds one PowerPoint deck per topic in slides.json (extracted from the
// "Qualitative Research – 12 Topics, 25 Slides" workbook).
//
//   NODE_PATH=$PWD/node_modules node build.js [outDir] [applyThemePath]
//
// Every deck shares the same layouts and design system; each topic gets its
// own accent colour through the deck theme.

const fs = require("fs");
const path = require("path");
const pptxgen = require("pptxgenjs");
const React = require("react");
const ReactDOMServer = require("react-dom/server");
const sharp = require("sharp");
const Fa = require("react-icons/fa");

const OUT_DIR = path.resolve(process.argv[2] || path.join(__dirname, "..", "decks"));
const APPLY_THEME = process.argv[3];
const { applyTheme } = APPLY_THEME ? require(APPLY_THEME) : { applyTheme: null };
const DECKS = JSON.parse(fs.readFileSync(path.join(__dirname, "slides.json"), "utf8"));

// ---------- palette ----------
const ACCENTS = {
  "01": ["6D2E46", "FaRegEye"],        // Phenomenology – berry
  "02": ["2C5F2D", "FaSeedling"],      // Grounded theory – forest
  "03": ["A2462F", "FaGlobeAmericas"], // Ethnography – terracotta
  "04": ["1E2761", "FaSearchPlus"],    // Case study – navy
  "05": ["5B3A82", "FaBookOpen"],      // Narrative inquiry – violet
  "06": ["0B6E79", "FaClipboardList"], // Qualitative descriptive – teal
  "07": ["9A4317", "FaHandsHelping"],  // Participatory action – burnt orange
  "08": ["4A6526", "FaSyncAlt"],       // Action research – olive
  "09": ["2F5671", "FaLightbulb"],     // Interpretive description – slate blue
  "10": ["8C2F39", "FaComments"],      // Discourse analysis – crimson
  "11": ["6E5326", "FaLandmark"],      // Historical research – bronze
  "12": ["3E4A61", "FaUserEdit"],      // Autoethnography – ink
};

const mix = (hex, target, t) => {
  const a = parseInt(hex, 16), b = parseInt(target, 16);
  const ch = (s) => Math.round(((a >> s) & 255) * (1 - t) + ((b >> s) & 255) * t);
  return [16, 8, 0].map((s) => ch(s).toString(16).padStart(2, "0")).join("").toUpperCase();
};
const tint = (hex, t) => mix(hex, "FFFFFF", t);
const shade = (hex, t) => mix(hex, "000000", t);

function makeTheme(name, acc) {
  return {
    name,
    headFontFace: "Cambria",
    bodyFontFace: "Calibri",
    colors: {
      dk1: "1F2328", lt1: "FFFFFF", dk2: "59606B", lt2: tint(acc, 0.9),
      accent1: acc, accent2: shade(acc, 0.28), accent3: tint(acc, 0.72),
      accent4: tint(acc, 0.8), accent5: "8A929E", accent6: "D9DDE3",
      hlink: acc, folHlink: shade(acc, 0.28),
    },
  };
}

// ---------- icons ----------
const iconCache = new Map();
async function icon(name, hex) {
  const key = name + hex;
  if (!iconCache.has(key)) {
    const svg = ReactDOMServer.renderToStaticMarkup(
      React.createElement(Fa[name], { color: "#" + hex, size: 256 })
    );
    const buf = await sharp(Buffer.from(svg)).resize(256, 256).png().toBuffer();
    iconCache.set(key, "image/png;base64," + buf.toString("base64"));
  }
  return iconCache.get(key);
}

// Picks an icon for a content slide from its title (first match wins).
const ICON_RULES = [
  [/reference|reading/, "FaBook"],
  [/objective/, "FaBullseye"],
  [/worked|illustrative|example/, "FaPuzzlePiece"],
  [/ethic|consent|sensitive/, "FaShieldAlt"],
  [/strength|limitation|error/, "FaBalanceScale"],
  [/quality|credib|rigou?r|trustworth|sufficien|saturation/, "FaCheckCircle"],
  [/report|writing|present/, "FaFileAlt"],
  [/reflexiv|position|bracket|insider|self/, "FaUserCircle"],
  [/question|problem/, "FaQuestionCircle"],
  [/interview|conversation|focus group/, "FaMicrophone"],
  [/memo|fieldnote|journal|note/, "FaStickyNote"],
  [/sampl|participant|recruit|informant|partner|people/, "FaUsers"],
  [/access|gatekeep|entry/, "FaDoorOpen"],
  [/field|setting|site|boundar|bound/, "FaMapMarkedAlt"],
  [/archive|source|document|artifact|material|corpus/, "FaArchive"],
  [/cycle|iterat|concurrent|loop|spiral/, "FaSyncAlt"],
  [/time|period|chronolog|tempor/, "FaClock"],
  [/theor|integrat|model|relational|axial/, "FaProjectDiagram"],
  [/cod|categor|theme|pattern|meaning unit/, "FaTags"],
  [/analy|compar|interpret/, "FaSearch"],
  [/data|transcri|preparation/, "FaDatabase"],
  [/narrative|story|stories|discours|language|talk|text/, "FaComments"],
  [/root|origin|histor|philosoph|foundation/, "FaLandmark"],
  [/tradition|form|orientation|approach|type|design/, "FaLayerGroup"],
  [/fit|when|choos|select/, "FaCompass"],
  [/action|change|improv|plan/, "FaRoute"],
  [/thick|descri/, "FaPenFancy"],
  [/definition|purpose|what/, "FaLightbulb"],
];
const pickIcon = (title) => {
  const t = title.toLowerCase();
  for (const [re, name] of ICON_RULES) if (re.test(t)) return name;
  return "FaRegLightbulb";
};

// ---------- layout choice for content slides ----------
const LAYOUT_RULES = [
  ["table", /table/],
  ["feature", /example|card|hypothetical|excerpt|scenario|wording/],
  ["check", /checklist|prompts|guide|questions|criteria|template|appraisal|journal|log|outline/],
  ["flow", /diagram|workflow|timeline|cycle|loop|process|map|tracker|continuum|model|path|schedule|-to-/],
  ["cards", /matrix|table|comparison|grid|versus|forms/],
];
const ROTATION = ["rows", "cards", "lead", "flow", "check", "feature"];
function chooseLayout(s, prev) {
  const v = s.visual.toLowerCase();
  const t = s.title.toLowerCase();
  let pick = "rows";
  if (/worked|illustrative/.test(t)) pick = "feature";
  else if (/definition|purpose|when .* fit/.test(t)) pick = "lead";
  else for (const [name, re] of LAYOUT_RULES) if (re.test(v)) { pick = name; break; }
  if (pick === prev && pick === "table") pick = "cards";
  else if (pick === prev) pick = ROTATION[(ROTATION.indexOf(pick) + 1) % ROTATION.length];
  return pick;
}

// ---------- geometry (LAYOUT_WIDE 13.333 x 7.5) ----------
const W = 13.333, M = 0.6, CW = W - 2 * M; // content width 12.133
const TOP = 1.8, BOTTOM = 6.55;
const pad2 = (n) => String(n).padStart(2, "0");
const shadow = () => ({ type: "outer", color: "000000", opacity: 0.12, blur: 8, offset: 2, angle: 90 });

async function buildDeck(deck) {
  const num = deck.sheet.slice(0, 2);
  const [acc, topicIcon] = ACCENTS[num];
  const topic = deck.slides[0].title;
  const THEME = makeTheme(`Qualitative Research – ${topic}`, acc);

  const pres = new pptxgen();
  pres.layout = "LAYOUT_WIDE";
  pres.title = topic;
  pres.subject = "Qualitative research methods";
  pres.theme = { headFontFace: THEME.headFontFace, bodyFontFace: THEME.bodyFontFace };
  const C = pres.SchemeColor;
  const footer = `Qualitative Research Methods  ·  Topic ${num}  ·  ${topic}`;

  pres.defineSlideMaster({
    title: "Title",
    background: { color: C.accent1 },
    objects: [
      { placeholder: { options: { name: "kicker", type: "body", x: M, y: 1.3, w: 7.6, h: 0.4, fontSize: 13, bold: true, charSpacing: 3, color: C.accent3, margin: 0 }, text: "" } },
      { placeholder: { options: { name: "title", type: "title", x: M, y: 1.85, w: 7.6, h: 1.85, fontSize: 54, bold: true, color: C.background1, align: "left", valign: "bottom", margin: 0 }, text: "" } },
      { placeholder: { options: { name: "subtitle", type: "body", x: M, y: 3.95, w: 7.2, h: 1.0, fontSize: 22, color: C.background1, valign: "top", margin: 0 }, text: "" } },
      { placeholder: { options: { name: "detail", type: "body", x: M, y: 5.05, w: 7.2, h: 0.6, fontSize: 16, color: C.accent3, valign: "top", margin: 0 }, text: "" } },
    ],
  });
  const contentMaster = (name, bg, titleColor, footColor) => pres.defineSlideMaster({
    title: name,
    background: { color: bg },
    margin: [0.5, M, 0.6, M],
    objects: [
      { placeholder: { options: { name: "title", type: "title", x: 1.55, y: 0.45, w: W - 1.55 - M, h: 0.8, fontSize: 32, bold: true, color: titleColor, align: "left", valign: "middle", margin: 0 }, text: "" } },
      { text: { text: footer, options: { x: M, y: 6.92, w: 9, h: 0.3, fontSize: 11, color: footColor, margin: 0 } } },
    ],
    slideNumber: { x: W - M - 0.8, y: 6.92, w: 0.8, h: 0.3, fontSize: 11, color: footColor, align: "right", margin: 0 },
  });
  contentMaster("Content", C.background1, C.text1, C.text2);
  contentMaster("Dark content", C.accent1, C.background1, C.accent3);

  const iconWhite = (n) => icon(n, "FFFFFF");
  const iconAcc = (n) => icon(n, acc);

  // Accent circle with a white icon in it.
  async function iconCircle(slide, name, x, y, d, fill = C.accent1, white = true) {
    slide.addShape(pres.shapes.OVAL, { x, y, w: d, h: d, fill: { color: fill }, line: { type: "none" }, objectName: "Icon circle" });
    const pad = d * 0.25;
    slide.addImage({ data: white ? await iconWhite(name) : await iconAcc(name), x: x + pad, y: y + pad, w: d - 2 * pad, h: d - 2 * pad, altText: "", objectName: "Icon" });
  }
  // Number badge.
  function numCircle(slide, n, x, y, d, fill = C.accent1, color = C.background1, size = 16) {
    slide.addText(pad2(n), { shape: pres.shapes.OVAL, x, y, w: d, h: d, fill: { color: fill }, line: { type: "none" }, color, fontSize: size, bold: true, align: "center", valign: "middle", margin: 0, isTextBox: true, objectName: `Number ${n}` });
  }
  function body(slide, text, opts) {
    slide.addText(text, { fontSize: 18, color: C.text1, valign: "top", margin: 0, isTextBox: true, paraSpaceAfter: 4, ...opts });
  }
  async function header(slide, s) {
    await iconCircle(slide, pickIcon(s.title), M, 0.47, 0.76);
    slide.addText(s.title, { placeholder: "title" });
  }

  // ----- content layouts (three talking points each) -----
  const L = {
    async cards(slide, pts) {
      const gap = 0.35, w = (CW - gap * (pts.length - 1)) / pts.length, h = 3.7, y0 = TOP + 0.45;
      pts.forEach((p, i) => {
        const x = M + i * (w + gap);
        slide.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y: y0, w, h, rectRadius: 0.12, fill: { color: C.background2 }, line: { type: "none" }, objectName: `Card ${i + 1}` });
        numCircle(slide, i + 1, x + 0.4, y0 + 0.4, 0.75);
        body(slide, p, { x: x + 0.4, y: y0 + 1.45, w: w - 0.8, h: h - 1.7, fontSize: 22 });
      });
    },
    async table(slide, pts, s) {
      const hdr = (t) => ({ text: t, options: { bold: true, color: C.background1, fill: { color: C.accent1 }, fontSize: 14, charSpacing: 2 } });
      const rows = [[hdr("NO."), hdr("KEY POINT")]];
      pts.forEach((p, i) => rows.push([
        { text: pad2(i + 1), options: { bold: true, color: C.accent1, fontSize: 22, fontFace: "Cambria", align: "center" } },
        { text: p, options: { color: C.text1, fontSize: 20 } },
      ].map((c) => ({ ...c, options: { ...c.options, fill: { color: i % 2 ? C.background1 : C.background2 } } }))));
      const rowH = 1.2;
      slide.addTable(rows, { x: M, y: TOP + 0.15, w: CW, colW: [1.3, CW - 1.3], rowH: [0.55, rowH, rowH, rowH], valign: "middle", margin: [0.08, 0.3, 0.08, 0.3], border: { type: "solid", pt: 1, color: THEME.colors.accent6 }, objectName: `${s.title} table` });
    },
    async flow(slide, pts) {
      const gap = 0.35, w = (CW - gap * (pts.length - 1)) / pts.length, d = 1.0, cy = TOP + 0.55, ph = 2.6;
      // connector behind the circles
      slide.addShape(pres.shapes.LINE, { x: M + d, y: cy + d / 2, w: (pts.length - 1) * (w + gap) - d, h: 0, line: { color: C.accent3, width: 2.5, endArrowType: "triangle" }, objectName: "Flow connector" });
      pts.forEach((p, i) => {
        const x = M + i * (w + gap);
        if (i > 0) slide.addShape(pres.shapes.OVAL, { x: x - 0.12, y: cy - 0.12, w: d + 0.24, h: d + 0.24, fill: { color: C.background1 }, line: { type: "none" }, objectName: "Flow gap" });
        numCircle(slide, i + 1, x, cy, d, C.accent1, C.background1, 24);
        slide.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y: cy + d + 0.4, w, h: ph, rectRadius: 0.12, fill: { color: C.background2 }, line: { type: "none" }, objectName: `Step ${i + 1} panel` });
        body(slide, p, { x: x + 0.35, y: cy + d + 0.4, w: w - 0.7, h: ph, fontSize: 22, valign: "middle", paraSpaceAfter: 0 });
      });
    },
    async rows(slide, pts, s) {
      const rowH = 1.4, textW = 7.0;
      pts.forEach((p, i) => {
        const y = TOP + 0.15 + i * (rowH + 0.15);
        numCircle(slide, i + 1, M, y + 0.05, 0.62);
        body(slide, p, { x: M + 0.95, y: y + 0.05, w: textW, h: rowH, fontSize: 22 });
      });
      const px = M + 0.95 + textW + 0.5, pw = W - M - px;
      slide.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: px, y: TOP + 0.15, w: pw, h: BOTTOM - TOP - 0.15, rectRadius: 0.15, fill: { color: C.background2 }, line: { type: "none" }, objectName: "Visual panel" });
      const d = 2.1;
      await iconCircle(slide, pickIcon(s.title), px + (pw - d) / 2, TOP + 0.15 + (BOTTOM - TOP - 0.15 - d) / 2, d);
    },
    async feature(slide, pts, s) {
      const pw = 5.0, ph = BOTTOM - TOP - 0.1;
      const hypo = /hypothetical/i.test(pts[0]);
      const label = hypo ? "HYPOTHETICAL EXAMPLE" : /example|illustrative/i.test(s.title) ? "ILLUSTRATIVE EXAMPLE" : "KEY IDEA";
      const lead = hypo ? pts[0].replace(/^Hypothetical (example|excerpt|observation|path|account|accounts)?:?\s*/i, "") : pts[0];
      const leadText = lead.charAt(0).toUpperCase() + lead.slice(1);
      slide.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: M, y: TOP + 0.1, w: pw, h: ph, rectRadius: 0.15, fill: { color: C.accent1 }, line: { type: "none" }, shadow: shadow(), objectName: "Feature panel" });
      slide.addImage({ data: await icon("FaQuoteLeft", THEME.colors.accent3), x: M + 0.45, y: TOP + 0.5, w: 0.55, h: 0.55, altText: "", objectName: "Quote mark" });
      slide.addText(label, { x: M + 1.2, y: TOP + 0.55, w: pw - 1.6, h: 0.45, fontSize: 12, bold: true, charSpacing: 3, color: C.accent3, valign: "middle", margin: 0, isTextBox: true });
      slide.addText(leadText, { x: M + 0.45, y: TOP + 1.35, w: pw - 0.9, h: ph - 1.65, fontFace: "Cambria", italic: true, fontSize: 28, color: C.background1, valign: "top", margin: 0, isTextBox: true });
      const rx = M + pw + 0.55, rw = W - M - rx, rh = 1.85;
      pts.slice(1).forEach((p, i) => {
        const y = TOP + 0.1 + i * (rh + 0.4) + 0.2;
        numCircle(slide, i + 2, rx, y, 0.62);
        body(slide, p, { x: rx + 0.95, y: y + 0.05, w: rw - 0.95, h: rh, fontSize: 22 });
      });
    },
    async check(slide, pts) {
      const rh = 1.25, gap = 0.3;
      for (let i = 0; i < pts.length; i++) {
        const y = TOP + 0.2 + i * (rh + gap);
        slide.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: M, y, w: CW, h: rh, rectRadius: 0.12, fill: { color: C.background2 }, line: { type: "none" }, objectName: `Check row ${i + 1}` });
        await iconCircle(slide, "FaCheck", M + 0.35, y + (rh - 0.62) / 2, 0.62);
        body(slide, pts[i], { x: M + 1.3, y, w: CW - 1.7, h: rh, fontSize: 22, valign: "middle", paraSpaceAfter: 0 });
      }
    },
    async lead(slide, pts, s) {
      const th = 1.75;
      slide.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: M, y: TOP + 0.1, w: CW, h: th, rectRadius: 0.15, fill: { color: C.accent1 }, line: { type: "none" }, shadow: shadow(), objectName: "Lead panel" });
      await iconCircle(slide, pickIcon(s.title), M + 0.45, TOP + 0.1 + (th - 0.95) / 2, 0.95, C.accent2);
      slide.addText(pts[0], { x: M + 1.75, y: TOP + 0.1, w: CW - 2.2, h: th, fontFace: "Cambria", fontSize: 26, color: C.background1, valign: "middle", margin: 0, isTextBox: true });
      const rest = pts.slice(1), gap = 0.35, w = (CW - gap * (rest.length - 1)) / rest.length;
      const y = TOP + 0.1 + th + 0.35, h = BOTTOM - y;
      rest.forEach((p, i) => {
        const x = M + i * (w + gap);
        slide.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w, h, rectRadius: 0.12, fill: { color: C.background2 }, line: { type: "none" }, objectName: `Point ${i + 2}` });
        numCircle(slide, i + 2, x + 0.35, y + (h - 0.62) / 2, 0.62);
        body(slide, p, { x: x + 1.25, y, w: w - 1.6, h, fontSize: 20, valign: "middle", paraSpaceAfter: 0 });
      });
    },
  };

  const notes = (s) => s.points.join("\n");
  const sectionOf = (n) => (n <= 2 ? "Introduction" : n <= 23 ? "Core content" : "Wrap-up");
  let currentSection = null, prevLayout = null;

  for (const s of deck.slides) {
    const sectionTitle = sectionOf(s.n);
    if (sectionTitle !== currentSection) { pres.addSection({ title: sectionTitle }); currentSection = sectionTitle; }

    if (s.n === 1) {
      const slide = pres.addSlide({ masterName: "Title", sectionTitle });
      slide.addText(`QUALITATIVE RESEARCH METHODS  ·  TOPIC ${num} OF 12`, { placeholder: "kicker" });
      slide.addText(s.title, { placeholder: "title" });
      slide.addText(s.points[0], { placeholder: "subtitle" });
      if (s.points[1]) slide.addText(s.points[1], { placeholder: "detail" });
      const d = 3.7, x = 8.75, y = (7.5 - d) / 2;
      slide.addShape(pres.shapes.OVAL, { x: x - 0.35, y: y - 0.35, w: d + 0.7, h: d + 0.7, fill: { color: C.accent2, transparency: 55 }, line: { type: "none" }, objectName: "Halo" });
      await iconCircle(slide, topicIcon, x, y, d, C.accent2);
      slide.addNotes(notes(s));
      continue;
    }

    if (s.n === 24) {
      const slide = pres.addSlide({ masterName: "Dark content", sectionTitle });
      await iconCircle(slide, "FaFlagCheckered" in Fa ? "FaFlagCheckered" : "FaCheckCircle", M, 0.47, 0.76, C.accent2);
      slide.addText(s.title, { placeholder: "title" });
      const plain = s.points.filter((p) => !/^discuss/i.test(p));
      const prompt = s.points.find((p) => /^discuss/i.test(p));
      const rowH = 1.05;
      plain.forEach((p, i) => {
        const y = TOP + 0.2 + i * (rowH + 0.25);
        numCircle(slide, i + 1, M, y + 0.1, 0.65, C.background1, C.accent1);
        body(slide, p, { x: M + 1.0, y: y + 0.1, w: CW - 1.0, h: rowH, fontSize: 22, color: C.background1 });
      });
      if (prompt) {
        const y = TOP + 0.2 + plain.length * (rowH + 0.25) + 0.15, h = BOTTOM - y;
        slide.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: M, y, w: CW, h, rectRadius: 0.15, fill: { color: C.accent2 }, line: { type: "none" }, objectName: "Discussion prompt" });
        await iconCircle(slide, "FaComments", M + 0.35, y + (h - 0.85) / 2, 0.85, C.background1, false);
        slide.addText([
          { text: "DISCUSSION PROMPT", options: { fontSize: 12, bold: true, charSpacing: 3, color: C.accent3, breakLine: true } },
          { text: prompt, options: { fontSize: 22, fontFace: "Cambria", italic: true, color: C.background1 } },
        ], { x: M + 1.55, y, w: CW - 1.95, h, valign: "middle", margin: 0, isTextBox: true, paraSpaceBefore: 4 });
      }
      slide.addNotes(notes(s));
      prevLayout = null;
      continue;
    }

    const slide = pres.addSlide({ masterName: "Content", sectionTitle });
    await header(slide, s);

    if (s.n === 2) {
      const icons = ["FaLightbulb", "FaClipboardList", "FaCheckCircle"];
      const gap = 0.35, w = (CW - gap * 2) / 3, h = BOTTOM - TOP - 0.1;
      for (let i = 0; i < s.points.length; i++) {
        const x = M + i * (w + gap);
        slide.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y: TOP + 0.1, w, h, rectRadius: 0.12, fill: { color: C.background1 }, line: { color: C.accent6, width: 1 }, shadow: shadow(), objectName: `Objective ${i + 1}` });
        await iconCircle(slide, icons[i], x + 0.4, TOP + 0.5, 0.95);
        slide.addText(`OBJECTIVE ${pad2(i + 1)}`, { x: x + 0.4, y: TOP + 1.7, w: w - 0.8, h: 0.35, fontSize: 12, bold: true, charSpacing: 3, color: C.accent1, margin: 0, isTextBox: true });
        body(slide, s.points[i], { x: x + 0.4, y: TOP + 2.15, w: w - 0.8, h: h - 2.4, fontSize: 20 });
      }
    } else if (s.n === 25) {
      const refs = s.points.map((p) => {
        const k = p.search(/https?:\/\//);
        return k < 0 ? { cite: p, url: null } : { cite: p.slice(0, k).trim().replace(/[.:]\s*$/, ".").replace(/Publisher overview\.$/, "Publisher overview"), url: p.slice(k).trim() };
      });
      const gap = 0.35, h = Math.min(1.7, (BOTTOM - TOP - 0.2 - gap * (refs.length - 1)) / refs.length);
      for (let i = 0; i < refs.length; i++) {
        const r = refs[i], y = TOP + 0.2 + i * (h + gap);
        slide.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: M, y, w: CW, h, rectRadius: 0.12, fill: { color: C.background2 }, line: { type: "none" }, objectName: `Reference ${i + 1}` });
        await iconCircle(slide, "FaBook", M + 0.35, y + (h - 0.7) / 2, 0.7);
        const runs = [{ text: r.cite, options: { fontSize: 20, color: C.text1, breakLine: !!r.url } }];
        if (r.url) runs.push({ text: r.url, options: { fontSize: 14, color: C.accent1, hyperlink: { url: r.url, tooltip: r.url } } });
        slide.addText(runs, { x: M + 1.35, y, w: CW - 1.75, h, valign: "middle", margin: 0, isTextBox: true, paraSpaceAfter: 6 });
      }
    } else {
      const layout = chooseLayout(s, prevLayout);
      await L[layout](slide, s.points, s);
      prevLayout = layout;
    }
    slide.addNotes(notes(s));
  }

  const file = path.join(OUT_DIR, `${num}_${topic.replace(/[^A-Za-z0-9]+/g, "_").replace(/_+$/, "")}.pptx`);
  await pres.writeFile({ fileName: file });
  if (applyTheme) await applyTheme(file, THEME);
  return file;
}

(async () => {
  fs.mkdirSync(OUT_DIR, { recursive: true });
  const only = process.env.ONLY;
  for (const d of DECKS) {
    if (only && !only.split(",").includes(d.sheet.slice(0, 2))) continue;
    console.log(await buildDeck(d));
  }
})().catch((e) => { console.error(e); process.exit(1); });
