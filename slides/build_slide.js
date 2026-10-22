// Builds "מסלול המפגש" as a single 16:9 slide. Animations are injected afterwards by add_animation.py.
const pptxgen = require("pptxgenjs");

const PURPLE = "7E3A94", ORANGE = "F9A93B", DARK = "4A2560", GRAY = "6B6B6B";
const LILAC = "F2EAF6", PEACH = "FAF0E3", TRACK = "E6DAEE";
const FONT = "Arial";

const pres = new pptxgen();
pres.layout = "LAYOUT_16x9"; // 10 x 5.625
pres.title = "מסלול המפגש";
const slide = pres.addSlide();
slide.background = { color: "FFFFFF" };

const he = { fontFace: FONT, rtlMode: true, lang: "he-IL", align: "right", isTextBox: true };

// Logo
slide.addShape(pres.shapes.DIAMOND, { objectName: "LogoPurple", x: 0.35, y: 0.28, w: 0.45, h: 0.45, fill: { color: PURPLE }, line: { type: "none" } });
slide.addShape(pres.shapes.DIAMOND, { objectName: "LogoOrange", x: 0.62, y: 0.47, w: 0.28, h: 0.28, fill: { color: ORANGE }, line: { type: "none" } });

// Title + subtitle
slide.addText("מסלול המפגש", { ...he, objectName: "Title", x: 3.5, y: 0.3, w: 6.0, h: 0.6, fontSize: 32, color: PURPLE, margin: 0, valign: "middle" });
slide.addText("שלושה שלבים, מהתהליך אל התוצר", { ...he, objectName: "Subtitle", x: 3.5, y: 0.92, w: 6.0, h: 0.35, fontSize: 14, color: "8A8A8A", margin: 0, valign: "middle" });

// Road: faint full track, then dashed segments drawn right -> left
const ROAD_Y = 1.95, START_X = 9.5, END_X = 0.6;
const STATIONS = [8.125, 5.205, 2.285]; // circle centers = card centers
slide.addShape(pres.shapes.LINE, { objectName: "Track", x: END_X, y: ROAD_Y, w: START_X - END_X, h: 0, line: { color: TRACK, width: 7 } });
slide.addShape(pres.shapes.OVAL, { objectName: "StartDot", x: START_X - 0.09, y: ROAD_Y - 0.09, w: 0.18, h: 0.18, fill: { color: "FFFFFF" }, line: { color: PURPLE, width: 2 } });

const stops = [START_X, ...STATIONS, END_X];
for (let i = 0; i < 4; i++) {
  const right = stops[i], left = stops[i + 1];
  const opts = { objectName: `Seg${i + 1}`, x: left, y: ROAD_Y, w: right - left, h: 0, line: { color: PURPLE, width: 2.5, dashType: "dash" } };
  if (i === 3) opts.line.beginArrowType = "triangle";
  slide.addShape(pres.shapes.LINE, opts);
}

// Travelling markers (one per segment; each rests under the next station circle)
const M = 0.26;
for (let i = 0; i < 4; i++) {
  slide.addShape(pres.shapes.DIAMOND, { objectName: `Marker${i + 1}`, x: stops[i] - M / 2, y: ROAD_Y - M / 2, w: M, h: M, fill: { color: ORANGE }, line: { color: "FFFFFF", width: 1.5 } });
}

// Cards + station circles
const CARD_W = 2.75, CARD_Y = 2.45, CARD_H = 1.9, D = 0.62;
const cards = [
  { n: "1", circle: PURPLE, fill: LILAC, title: "איך מילה נבנית?", body: "חומסקי, פיאז'ה וויגוצקי: מה הילד מביא, מה הוא בונה ומה הסביבה מתווכת" },
  { n: "2", circle: ORANGE, fill: PEACH, title: "מה זה לדעת מילה?", body: "ידיעה כרצף, היבטי הידיעה, קולטני ויצרני, רוחב ועומק, המילון כרשת והשורש העברי" },
  { n: "3", circle: PURPLE, fill: LILAC, title: "למה זה משנה?", body: "הצצה לקשר עם הקריאה ולפערים בין ילדים, שבהם נעסוק במפגשים הבאים" },
];
cards.forEach((c, i) => {
  const cx = STATIONS[i];
  slide.addText([
    { text: c.title, options: { fontSize: 19, bold: true, color: DARK, breakLine: true, paraSpaceAfter: 8 } },
    { text: c.body, options: { fontSize: 14, color: GRAY } },
  ], { ...he, objectName: `Card${i + 1}`, shape: pres.shapes.ROUNDED_RECTANGLE, rectRadius: 0.08, x: cx - CARD_W / 2, y: CARD_Y, w: CARD_W, h: CARD_H,
       fill: { color: c.fill }, line: { type: "none" }, valign: "top", margin: [14, 16, 10, 16], lineSpacingMultiple: 1.05 });
  slide.addText(c.n, { fontFace: FONT, isTextBox: false, objectName: `Circle${i + 1}`, shape: pres.shapes.OVAL, x: cx - D / 2, y: ROAD_Y - D / 2, w: D, h: D,
       fill: { color: c.circle }, line: { color: "FFFFFF", width: 2.5 }, fontSize: 18, bold: true, color: "FFFFFF", align: "center", valign: "middle", margin: 0 });
});

// Footer
slide.addText([
  { text: "בסוף המפגש: ", options: { color: PURPLE, bold: true } },
  { text: "משימה בפורום וכלי לשימוש בכיתה, בדיקת הבנה", options: { color: "444444" } },
], { ...he, objectName: "Footer", x: 0.6, y: 4.65, w: 8.9, h: 0.4, fontSize: 15, margin: 0, valign: "middle" });

pres.writeFile({ fileName: process.argv[2] || "meeting_path_raw.pptx" }).then(f => console.log("wrote", f));
