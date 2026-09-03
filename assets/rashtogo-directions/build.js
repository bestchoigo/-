const pptxgen = require("pptxgenjs");

const INK = "241A20";
const BERRY = "6D2E46";
const ROSE = "C4708C";
const TINT = "F7EFF2";
const TEXT = "2C2229";
const MUTED = "7B6A72";
const WHITE = "FFFFFF";

const KR = "맑은 고딕";

const sh = (o = {}) =>
  Object.assign({ type: "outer", color: "000000", blur: 14, offset: 4, angle: 90, opacity: 0.18 }, o);

const pres = new pptxgen();
pres.layout = "LAYOUT_WIDE"; // 13.33 x 7.5
pres.author = "RASHTOGO";
pres.title = "래쉬투고 찾아오시는 길";

/* ─────────────────────────── SLIDE 1 · 메인 안내 ─────────────────────────── */
const s1 = pres.addSlide();
s1.background = { color: INK };

s1.addImage({
  path: "images/hero.jpg",
  x: 7.13, y: 0, w: 6.2, h: 7.5,
  sizing: { type: "cover", w: 6.2, h: 7.5 },
});

s1.addText("RASHTOGO", {
  x: 0.9, y: 0.85, w: 5.6, h: 0.35, isTextBox: true, margin: 0,
  fontFace: "Arial", fontSize: 14, bold: true, color: ROSE, charSpacing: 6,
});
s1.addText("래쉬투고", {
  x: 0.9, y: 1.24, w: 5.9, h: 0.85, isTextBox: true, margin: 0,
  fontFace: KR, fontSize: 40, bold: true, color: WHITE,
});
s1.addText("찾아오시는 길", {
  x: 0.9, y: 2.06, w: 5.9, h: 0.95, isTextBox: true, margin: 0,
  fontFace: KR, fontSize: 46, bold: true, color: WHITE,
});

// 지하철 배지
s1.addShape(pres.ShapeType.roundRect, {
  x: 0.9, y: 3.18, w: 5.1, h: 0.72, rectRadius: 0.34,
  fill: { color: BERRY }, line: { type: "none" }, shadow: sh({ blur: 12, offset: 3, opacity: 0.35 }),
});
s1.addText("지하철 5호선 마곡역 3번 출구 · 도보 1분", {
  x: 0.9, y: 3.18, w: 5.1, h: 0.72, isTextBox: true, align: "center", valign: "middle",
  fontFace: KR, fontSize: 17, bold: true, color: WHITE,
});

// 주소 블록
s1.addText("ADDRESS", {
  x: 0.9, y: 4.28, w: 5.9, h: 0.3, isTextBox: true, margin: 0,
  fontFace: "Arial", fontSize: 11, bold: true, color: ROSE, charSpacing: 4,
});
s1.addText("서울특별시 강서구 마곡동 773", {
  x: 0.9, y: 4.62, w: 5.9, h: 0.42, isTextBox: true, margin: 0,
  fontFace: KR, fontSize: 20, bold: true, color: WHITE,
});
s1.addText("힐스테이트 에코 마곡역  340호", {
  x: 0.9, y: 5.06, w: 5.9, h: 0.42, isTextBox: true, margin: 0,
  fontFace: KR, fontSize: 20, bold: true, color: WHITE,
});
s1.addText("상가 / 오피스 엘리베이터 이용 · 3층", {
  x: 0.9, y: 5.56, w: 5.9, h: 0.36, isTextBox: true, margin: 0,
  fontFace: KR, fontSize: 15, color: "C9B9C0",
});

s1.addText("마곡역 3번 출구에서 나와 직진하면 바로 보이는 건물입니다.", {
  x: 0.9, y: 6.32, w: 5.9, h: 0.4, isTextBox: true, margin: 0,
  fontFace: KR, fontSize: 13, italic: true, color: ROSE,
});
s1.addNotes("래쉬투고 찾아오시는 길 안내. 주소: 서울특별시 강서구 마곡동 773 힐스테이트 에코 마곡역 340호. 마곡역 3번 출구 도보 1분.");

/* ─────────────────────────── SLIDE 2 · 3 STEP 경로 ─────────────────────────── */
const s2 = pres.addSlide();
s2.background = { color: WHITE };

s2.addText("마곡역에서 오시는 길", {
  x: 0.6, y: 0.55, w: 8.5, h: 0.7, isTextBox: true, margin: 0,
  fontFace: KR, fontSize: 36, bold: true, color: TEXT,
});
s2.addText("3번 출구 기준 · 도보 1분", {
  x: 0.6, y: 1.22, w: 8.5, h: 0.4, isTextBox: true, margin: 0,
  fontFace: KR, fontSize: 15, color: MUTED,
});
s2.addShape(pres.ShapeType.roundRect, {
  x: 10.55, y: 0.62, w: 2.18, h: 0.62, rectRadius: 0.3,
  fill: { color: TINT }, line: { type: "none" },
});
s2.addText("총 3 STEP", {
  x: 10.55, y: 0.62, w: 2.18, h: 0.62, isTextBox: true, align: "center", valign: "middle",
  fontFace: KR, fontSize: 15, bold: true, color: BERRY,
});

const steps = [
  {
    img: "images/card1.jpg",
    head: "마곡역 3번 출구",
    body: "출구로 나와 GS25 · 올렌즈(OLENS)가\n있는 상가 방향으로 직진합니다.",
  },
  {
    img: "images/card2.jpg",
    head: "GATE 1 입구 진입",
    body: "올렌즈와 나이스공인중개사무소 사이\n‘GATE 1’ 통로로 들어갑니다.",
  },
  {
    img: "images/card3.jpg",
    head: "GATE 4 엘리베이터",
    body: "통로 안쪽 ‘상가 / 오피스’ GATE 4에서\n엘리베이터를 타고 3층으로 갑니다.",
  },
];

const CW = 3.83, GAP = 0.32, X0 = 0.6, CY = 1.92, CH = 4.68;
steps.forEach((st, i) => {
  const x = X0 + i * (CW + GAP);
  s2.addShape(pres.ShapeType.roundRect, {
    x, y: CY, w: CW, h: CH, rectRadius: 0.06,
    fill: { color: TINT }, line: { type: "none" }, shadow: sh({ blur: 16, offset: 3, opacity: 0.14 }),
  });
  s2.addImage({
    path: st.img, x: x + 0.24, y: CY + 0.24, w: CW - 0.48, h: 2.05,
    sizing: { type: "cover", w: CW - 0.48, h: 2.05 },
  });
  s2.addShape(pres.ShapeType.ellipse, {
    x: x + 0.24, y: CY + 2.45, w: 0.6, h: 0.6,
    fill: { color: BERRY }, line: { type: "none" },
  });
  s2.addText(`${i + 1}`, {
    x: x + 0.24, y: CY + 2.45, w: 0.6, h: 0.6, isTextBox: true, align: "center", valign: "middle",
    fontFace: "Arial", fontSize: 22, bold: true, color: WHITE,
  });
  s2.addText(st.head, {
    x: x + 0.24, y: CY + 3.2, w: CW - 0.48, h: 0.45, isTextBox: true, margin: 0,
    fontFace: KR, fontSize: 19, bold: true, color: TEXT,
  });
  s2.addText(st.body, {
    x: x + 0.24, y: CY + 3.7, w: CW - 0.48, h: 0.78, isTextBox: true, margin: 0,
    fontFace: KR, fontSize: 13, color: MUTED, lineSpacingMultiple: 1.25,
  });
});
s2.addNotes("STEP 1 마곡역 3번 출구 → STEP 2 GATE 1 통로 진입 → STEP 3 GATE 4 상가/오피스 엘리베이터 탑승 후 3층.");

/* ─────────────────────────── SLIDE 3 · 도착 안내 ─────────────────────────── */
const s3 = pres.addSlide();
s3.background = { color: WHITE };

s3.addText("도착 · 3층 340호", {
  x: 0.6, y: 0.55, w: 7.4, h: 0.7, isTextBox: true, margin: 0,
  fontFace: KR, fontSize: 36, bold: true, color: TEXT,
});
s3.addText("엘리베이터 홀에서 3층 버튼을 눌러 주세요.", {
  x: 0.6, y: 1.22, w: 7.4, h: 0.4, isTextBox: true, margin: 0,
  fontFace: KR, fontSize: 15, color: MUTED,
});

// 큰 호수 콜아웃
s3.addShape(pres.ShapeType.roundRect, {
  x: 0.6, y: 1.95, w: 5.2, h: 2.15, rectRadius: 0.06,
  fill: { color: INK }, line: { type: "none" }, shadow: sh({ blur: 16, offset: 4, opacity: 0.25 }),
});
s3.addText([
  { text: "340", options: { fontFace: "Arial", fontSize: 60, bold: true, color: WHITE } },
  { text: "호", options: { fontFace: KR, fontSize: 30, bold: true, color: ROSE } },
], {
  x: 0.95, y: 2.08, w: 4.5, h: 1.2, isTextBox: true, margin: 0, valign: "middle",
});
s3.addText("힐스테이트 에코 마곡역 상가 3층", {
  x: 0.98, y: 3.28, w: 4.5, h: 0.4, isTextBox: true, margin: 0,
  fontFace: KR, fontSize: 14, color: "C9B9C0",
});
s3.addText("엘리베이터 하차 후 층 안내판을 확인해 주세요.", {
  x: 0.6, y: 4.22, w: 5.2, h: 0.35, isTextBox: true, margin: 0,
  fontFace: KR, fontSize: 12, color: MUTED,
});

// 체크포인트
const checks = [
  ["1층 랜드마크", "올렌즈(OLENS) · GS25 · VIVA NAPOLI"],
  ["건물 입구", "GATE 1 통로 → 안쪽 GATE 4"],
  ["엘리베이터", "상가 / 오피스 전용 엘리베이터 홀"],
];
checks.forEach(([label, val], i) => {
  const y = 4.82 + i * 0.82;
  s3.addShape(pres.ShapeType.ellipse, {
    x: 0.6, y: y + 0.06, w: 0.44, h: 0.44,
    fill: { color: TINT }, line: { type: "none" },
  });
  s3.addShape(pres.ShapeType.ellipse, {
    x: 0.74, y: y + 0.2, w: 0.16, h: 0.16,
    fill: { color: BERRY }, line: { type: "none" },
  });
  s3.addText(label, {
    x: 1.2, y: y, w: 1.7, h: 0.3, isTextBox: true, margin: 0,
    fontFace: KR, fontSize: 12, bold: true, color: BERRY,
  });
  s3.addText(val, {
    x: 1.2, y: y + 0.28, w: 4.6, h: 0.34, isTextBox: true, margin: 0,
    fontFace: KR, fontSize: 14, color: TEXT,
  });
});

s3.addImage({
  path: "images/tall2.jpg", x: 6.25, y: 1.95, w: 3.3, h: 3.3,
  sizing: { type: "cover", w: 3.3, h: 3.3 },
});
s3.addText("GATE 1 통로 입구", {
  x: 6.25, y: 5.34, w: 3.3, h: 0.32, isTextBox: true, margin: 0,
  fontFace: KR, fontSize: 12, color: MUTED,
});
s3.addImage({
  path: "images/card3.jpg", x: 9.83, y: 1.95, w: 2.9, h: 3.3,
  sizing: { type: "cover", w: 2.9, h: 3.3 },
});
s3.addText("GATE 4 엘리베이터 홀", {
  x: 9.83, y: 5.34, w: 2.9, h: 0.32, isTextBox: true, margin: 0,
  fontFace: KR, fontSize: 12, color: MUTED,
});

s3.addShape(pres.ShapeType.roundRect, {
  x: 6.25, y: 5.95, w: 6.48, h: 1.0, rectRadius: 0.06,
  fill: { color: TINT }, line: { type: "none" },
});
s3.addText("서울특별시 강서구 마곡동 773 힐스테이트 에코 마곡역 340호", {
  x: 6.55, y: 6.06, w: 5.9, h: 0.36, isTextBox: true, margin: 0,
  fontFace: KR, fontSize: 14, bold: true, color: TEXT,
});
s3.addText("지하철 5호선 마곡역 3번 출구에서 도보 1분", {
  x: 6.55, y: 6.44, w: 5.9, h: 0.36, isTextBox: true, margin: 0,
  fontFace: KR, fontSize: 13, color: MUTED,
});
s3.addNotes("도착 안내: 상가/오피스 엘리베이터로 3층 이동, 340호. 주소 서울특별시 강서구 마곡동 773 힐스테이트 에코 마곡역.");

pres.writeFile({ fileName: "래쉬투고_찾아오시는길.pptx" }).then((f) => console.log("saved:", f));
