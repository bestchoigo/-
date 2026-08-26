const pptx = require("pptxgenjs");
const p = new pptx();
p.layout = "LAYOUT_WIDE";              // 13.333 x 7.5
p.author = "TS자산관리센터";
p.company = "TS자산관리센터";
p.title = "경영인 정기보험 활용 법인 비용처리·목돈마련 전략";

const W = 13.333, H = 7.5;
const NAVY_D = "0E2340", NAVY = "1B3D6B", NAVY_M = "2B4E7E";
const GOLD = "C9A227", GOLD_L = "E8D8A4";
const BG = "F4F6F9", CARD = "FFFFFF";
const TXT = "1C2430", MUTED = "6E7A8A", LIGHT = "B9C6D8", GRAY = "A9B4C2";
const F = "Malgun Gothic";

const sh = () => ({ type: "outer", color: "0E2340", blur: 10, offset: 2, angle: 90, opacity: 0.10 });

/* ─────────────────────────── SLIDE 1 : COVER ─────────────────────────── */
const s1 = p.addSlide();
s1.background = { color: NAVY_D };

// motif: gold circles
s1.addShape(p.ShapeType.ellipse, { x: 9.9, y: -1.5, w: 5.4, h: 5.4, fill: { color: NAVY, transparency: 55 } });
s1.addShape(p.ShapeType.ellipse, { x: 11.3, y: 4.4, w: 3.2, h: 3.2, fill: { color: NAVY_M, transparency: 65 } });
s1.addShape(p.ShapeType.ellipse, { x: 11.55, y: 0.72, w: 0.62, h: 0.62, fill: { color: GOLD } });

s1.addText("CEO TAX BRIEF", {
  x: 0.9, y: 1.25, w: 6, h: 0.3, isTextBox: true, margin: 0,
  fontFace: "Arial", fontSize: 11, bold: true, color: GOLD, charSpacing: 3
});

s1.addText([
  { text: "쌓이는 이익, 사라지는 세금", options: { color: "FFFFFF", breakLine: true } },
  { text: "대표님의 목돈으로 바꾸는 법", options: { color: GOLD_L } }
], {
  x: 0.88, y: 1.82, w: 9.6, h: 1.9, isTextBox: true, margin: 0,
  fontFace: F, fontSize: 40, bold: true, lineSpacing: 52
});

s1.addText("경영인 정기보험을 활용한 법인 비용처리 · 목돈마련 전략", {
  x: 0.9, y: 3.82, w: 9.2, h: 0.42, isTextBox: true, margin: 0,
  fontFace: F, fontSize: 17, color: LIGHT
});

// 3 chips
const chips = [
  ["01", "납입보험료 비용처리"],
  ["02", "법인세 최대 26.4% 절감"],
  ["03", "퇴직금 · 유고 재원 확보"]
];
chips.forEach(([n, t], i) => {
  const x = 0.9 + i * 3.62;
  s1.addShape(p.ShapeType.roundRect, { x, y: 4.72, w: 3.32, h: 1.0, rectRadius: 0.08, fill: { color: NAVY, transparency: 25 }, shadow: sh() });
  s1.addShape(p.ShapeType.ellipse, { x: x + 0.26, y: 5.02, w: 0.42, h: 0.42, fill: { color: GOLD } });
  s1.addText(n, { x: x + 0.26, y: 5.02, w: 0.42, h: 0.42, isTextBox: true, margin: 0, align: "center", valign: "middle", fontFace: "Arial", fontSize: 11, bold: true, color: NAVY_D });
  s1.addText(t, { x: x + 0.82, y: 4.98, w: 2.36, h: 0.5, isTextBox: true, margin: 0, valign: "middle", fontFace: F, fontSize: 13.5, bold: true, color: "FFFFFF" });
});

s1.addText("작성기관 : TS자산관리센터", {
  x: 0.9, y: 6.42, w: 6, h: 0.32, isTextBox: true, margin: 0,
  fontFace: F, fontSize: 11.5, bold: true, color: GOLD
});
s1.addText("법인 CEO를 위한 절세 · 자산관리 컨설팅   |   2026. 08", {
  x: 0.9, y: 6.76, w: 8, h: 0.3, isTextBox: true, margin: 0,
  fontFace: F, fontSize: 10, color: "7E8DA1"
});
s1.addNotes("표지. 경영인 정기보험을 활용한 법인 비용처리 및 목돈마련 전략 제안 - TS자산관리센터");

/* ───────────────────── SLIDE 2 : 문제 → 해법 ───────────────────── */
const s2 = p.addSlide();
s2.background = { color: BG };

s2.addText("이익은 남는데, 쓸 방법이 없습니다", {
  x: 0.7, y: 0.48, w: 9.5, h: 0.6, isTextBox: true, margin: 0,
  fontFace: F, fontSize: 32, bold: true, color: NAVY_D
});
s2.addText("대표님 회사의 세 가지 현실, 그리고 한 장의 계약으로 푸는 해법", {
  x: 0.72, y: 1.10, w: 9.5, h: 0.34, isTextBox: true, margin: 0,
  fontFace: F, fontSize: 14, color: MUTED
});

// LEFT : 현실
s2.addShape(p.ShapeType.ellipse, { x: 0.7, y: 1.75, w: 0.34, h: 0.34, fill: { color: "B3261E" } });
s2.addText("!", { x: 0.7, y: 1.75, w: 0.34, h: 0.34, isTextBox: true, margin: 0, align: "center", valign: "middle", fontFace: "Arial", fontSize: 13, bold: true, color: "FFFFFF" });
s2.addText("대표님의 현실", { x: 1.14, y: 1.74, w: 3.6, h: 0.36, isTextBox: true, margin: 0, valign: "middle", fontFace: F, fontSize: 17, bold: true, color: NAVY_D });

const pains = [
  ["현금은 쌓이는데, 넣을 비용이 없다", "이익잉여금이 늘수록 법인세도, 주식가치(상속·증여세)도 함께 올라갑니다."],
  ["급여 · 상여 · 배당으로 빼면 최고 49.5%", "어렵게 남긴 이익의 절반이 개인 소득세로 사라집니다."],
  ["대표 유고 = 회사의 존폐 위기", "대출 상환, 가지급금, 유족 보상까지 모두 회사가 떠안습니다."]
];
pains.forEach(([t, d], i) => {
  const y = 2.30 + i * 1.20;
  s2.addShape(p.ShapeType.roundRect, { x: 0.7, y, w: 5.55, h: 1.02, rectRadius: 0.06, fill: { color: "E9EDF3" } });
  s2.addText(t, { x: 0.98, y: y + 0.13, w: 5.05, h: 0.32, isTextBox: true, margin: 0, fontFace: F, fontSize: 14.5, bold: true, color: "B3261E" });
  s2.addText(d, { x: 0.98, y: y + 0.47, w: 5.05, h: 0.46, isTextBox: true, margin: 0, fontFace: F, fontSize: 11.5, color: "56616F", lineSpacing: 16 });
});

// RIGHT : 해법
s2.addShape(p.ShapeType.ellipse, { x: 7.05, y: 1.75, w: 0.34, h: 0.34, fill: { color: GOLD } });
s2.addText("✓", { x: 7.05, y: 1.75, w: 0.34, h: 0.34, isTextBox: true, margin: 0, align: "center", valign: "middle", fontFace: "Arial", fontSize: 12, bold: true, color: NAVY_D });
s2.addText("경영인 정기보험 한 장으로", { x: 7.49, y: 1.74, w: 5.2, h: 0.36, isTextBox: true, margin: 0, valign: "middle", fontFace: F, fontSize: 17, bold: true, color: NAVY_D });

const sols = [
  ["1", "비용처리", "납입보험료를 손금으로 처리해 과세표준을 즉시 낮춥니다.  연 1억 납입 시 법인세 약 2,090만 원 절감."],
  ["2", "목돈마련", "해지환급금이 회사의 자산으로 쌓입니다.  퇴직금 · 투자 · 비상자금 재원으로 활용."],
  ["3", "CEO 보장", "유고 시 사망보험금이 법인으로 지급되어 대출 상환과 경영권을 지킵니다."]
];
sols.forEach(([n, t, d], i) => {
  const y = 2.30 + i * 1.20;
  s2.addShape(p.ShapeType.roundRect, { x: 7.05, y, w: 5.58, h: 1.02, rectRadius: 0.06, fill: { color: CARD }, shadow: sh() });
  s2.addShape(p.ShapeType.ellipse, { x: 7.32, y: y + 0.30, w: 0.42, h: 0.42, fill: { color: NAVY_D } });
  s2.addText(n, { x: 7.32, y: y + 0.30, w: 0.42, h: 0.42, isTextBox: true, margin: 0, align: "center", valign: "middle", fontFace: "Arial", fontSize: 13, bold: true, color: GOLD });
  s2.addText(t, { x: 7.90, y: y + 0.11, w: 4.5, h: 0.30, isTextBox: true, margin: 0, fontFace: F, fontSize: 14.5, bold: true, color: NAVY_D });
  s2.addText(d, { x: 7.90, y: y + 0.43, w: 4.55, h: 0.50, isTextBox: true, margin: 0, fontFace: F, fontSize: 11, color: "56616F", lineSpacing: 15 });
});

// bottom band
s2.addShape(p.ShapeType.roundRect, { x: 0.7, y: 6.00, w: 11.93, h: 0.72, rectRadius: 0.07, fill: { color: NAVY_D } });
s2.addText([
  { text: "비용처리 · 목돈마련 · CEO보장", options: { color: GOLD_L, bold: true } },
  { text: "  —  세 가지를 한 계약으로 동시에 해결합니다.", options: { color: "FFFFFF", bold: true } }
], { x: 1.0, y: 6.00, w: 11.3, h: 0.72, isTextBox: true, margin: 0, valign: "middle", fontFace: F, fontSize: 16 });

s2.addText("※ 만기환급금이 없는 보장성보험의 기간경과분 보험료 기준(법인세법 시행령 제19조 제9호). 계약 구조에 따라 해지환급금 상당액의 자산계상 여부가 달라질 수 있어 가입 전 검토가 필요합니다.", {
  x: 0.72, y: 6.86, w: 11.9, h: 0.32, isTextBox: true, margin: 0, fontFace: F, fontSize: 9, color: "8A94A2"
});
s2.addNotes("현실 3가지(비용부족/고율 소득세/CEO 리스크) → 경영인 정기보험 3대 효과(비용처리·목돈마련·CEO보장)");

/* ───────────────────── SLIDE 3 : 숫자 + 회수 전략 ───────────────────── */
const s3 = p.addSlide();
s3.background = { color: BG };

s3.addText("연 1억 원 납입, 10년 뒤의 숫자", {
  x: 0.7, y: 0.48, w: 9.5, h: 0.6, isTextBox: true, margin: 0,
  fontFace: F, fontSize: 32, bold: true, color: NAVY_D
});
s3.addText("낸 돈보다 더 많이 돌아옵니다. 10년간의 대표이사 보장은 덤입니다.", {
  x: 0.72, y: 1.10, w: 9.5, h: 0.34, isTextBox: true, margin: 0,
  fontFace: F, fontSize: 14, color: MUTED
});

// LEFT : chart card
s3.addShape(p.ShapeType.roundRect, { x: 0.7, y: 1.66, w: 6.28, h: 3.92, rectRadius: 0.07, fill: { color: CARD }, shadow: sh() });
s3.addText("납입 대비 회수 재원 (단위 : 억 원)", { x: 1.0, y: 1.86, w: 5.7, h: 0.3, isTextBox: true, margin: 0, fontFace: F, fontSize: 12.5, bold: true, color: NAVY_D });

s3.addChart(p.ChartType.bar, [{
  name: "금액",
  labels: ["납입 누계", "해지환급금", "법인세 절감", "실질 회수"],
  values: [10.0, 9.0, 2.09, 11.09]
}], {
  x: 0.85, y: 2.22, w: 5.98, h: 2.62,
  barDir: "col", barGapWidthPct: 60,
  chartColors: [GRAY, NAVY_M, GOLD, NAVY_D],
  showValue: true, dataLabelPosition: "outEnd", dataLabelFormatCode: "0.0",
  dataLabelFontFace: "Arial", dataLabelFontSize: 12, dataLabelFontBold: true, dataLabelColor: TXT,
  showLegend: false, showTitle: false,
  catAxisLabelColor: "56616F", catAxisLabelFontFace: F, catAxisLabelFontSize: 11,
  valAxisHidden: true, valGridLine: { style: "none" }, catGridLine: { style: "none" },
  valAxisMaxVal: 13, valAxisMinVal: 0,
  catAxisLineShow: false, valAxisLineShow: false
});

s3.addText([
  { text: "10억 납입  →  11.09억 회수", options: { color: NAVY_D, bold: true } },
  { text: "     순증 +1.09억", options: { color: GOLD, bold: true } }
], { x: 1.0, y: 4.92, w: 5.7, h: 0.42, isTextBox: true, margin: 0, valign: "middle", fontFace: F, fontSize: 15 });

// RIGHT : 3 steps
s3.addText("회수 전략 3단계", { x: 7.32, y: 1.68, w: 5.3, h: 0.34, isTextBox: true, margin: 0, fontFace: F, fontSize: 17, bold: true, color: NAVY_D });

const steps = [
  ["STEP 1", "재직 중", "해지환급금은 회사의 유동자산. 자금이 필요하면 계약대출로 즉시 활용합니다."],
  ["STEP 2", "퇴직 시", "계약을 대표 개인 명의로 이전해 퇴직금으로 지급. 퇴직소득 분류과세로 세부담을 크게 낮춥니다."],
  ["STEP 3", "유고 시", "사망보험금이 법인에 지급되어 유족 보상 · 대출 상환 · 가업승계 재원이 됩니다."]
];
steps.forEach(([s, t, d], i) => {
  const y = 2.18 + i * 1.15;
  s3.addShape(p.ShapeType.roundRect, { x: 7.32, y, w: 5.31, h: 1.0, rectRadius: 0.06, fill: { color: CARD }, shadow: sh() });
  s3.addShape(p.ShapeType.ellipse, { x: 7.58, y: y + 0.29, w: 0.42, h: 0.42, fill: { color: GOLD } });
  s3.addText(String(i + 1), { x: 7.58, y: y + 0.29, w: 0.42, h: 0.42, isTextBox: true, margin: 0, align: "center", valign: "middle", fontFace: "Arial", fontSize: 13, bold: true, color: NAVY_D });
  s3.addText([
    { text: s + "   ", options: { color: GOLD, bold: true, fontSize: 9.5, fontFace: "Arial" } },
    { text: t, options: { color: NAVY_D, bold: true, fontSize: 14, fontFace: F } }
  ], { x: 8.16, y: y + 0.10, w: 4.3, h: 0.30, isTextBox: true, margin: 0, valign: "middle" });
  s3.addText(d, { x: 8.16, y: y + 0.42, w: 4.3, h: 0.50, isTextBox: true, margin: 0, fontFace: F, fontSize: 10.5, color: "56616F", lineSpacing: 14 });
});

// CTA band
s3.addShape(p.ShapeType.roundRect, { x: 0.7, y: 5.72, w: 11.93, h: 0.94, rectRadius: 0.07, fill: { color: NAVY_D } });
s3.addText("지금 필요한 것은 단 두 가지", { x: 1.0, y: 5.86, w: 3.6, h: 0.3, isTextBox: true, margin: 0, fontFace: F, fontSize: 12, bold: true, color: GOLD });
s3.addText("① 정관 · 임원 퇴직금 지급규정 정비    ② 회사 이익 규모에 맞는 적정 납입액 산정", {
  x: 1.0, y: 6.18, w: 7.6, h: 0.34, isTextBox: true, margin: 0, fontFace: F, fontSize: 14, bold: true, color: "FFFFFF"
});
s3.addText("TS자산관리센터  |  법인 절세 무료 진단", {
  x: 8.7, y: 5.98, w: 3.7, h: 0.42, isTextBox: true, margin: 0, align: "right", valign: "middle", fontFace: F, fontSize: 13, bold: true, color: GOLD_L
});

s3.addText("※ 예시 : 연 1억 원 10년 납입 가정. 법인세 절감액은 과세표준 2억~200억 원 구간(지방소득세 포함 20.9%) 기준. 해지환급금 90%는 예시이며 실제 환급률·보장금액은 보험사·상품·피보험자 조건에 따라 달라집니다.", {
  x: 0.72, y: 6.78, w: 11.9, h: 0.42, isTextBox: true, margin: 0, fontFace: F, fontSize: 9, color: "8A94A2", lineSpacing: 12
});
s3.addNotes("연 1억 10년 납입 시뮬레이션과 재직중/퇴직/유고 3단계 회수 전략, 그리고 실행 2단계 CTA");

p.writeFile({ fileName: "/home/user/-/deck/경영인정기보험_법인절세전략_TS자산관리센터.pptx" })
 .then(f => console.log("saved:", f));
