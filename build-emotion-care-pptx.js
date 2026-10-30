const pptxgen = require("pptxgenjs");
const { applyTheme } = require(process.env.APPLY_THEME || "./apply_theme.js");

const OUT = process.argv[2];
const FONT = "Malgun Gothic";

const THEME = {
  name: "Emotion Care",
  headFontFace: FONT,
  bodyFontFace: FONT,
  colors: {
    dk1: "1F2937", lt1: "FFFFFF", dk2: "6B7280", lt2: "E5F4EC",
    accent1: "16A36A", accent2: "1E90FF", accent3: "00A4A6",
    accent4: "DC143C", accent5: "F07020", accent6: "E3008C",
    hlink: "1E90FF", folHlink: "3455DB",
  },
};

// MakeCode block colors (fixed meaning, so plain hex)
const K = {
  husky: "12BB18", basic: "1E90FF", logic: "00A4A6", vari: "DC143C",
  music: "E63022", led: "5C2D91", radio: "E3008C", func: "3455DB", input: "C724B1",
};
const ROW = "F1F5FA";

const pres = new pptxgen();
pres.layout = "LAYOUT_WIDE"; // 13.333 x 7.5
pres.theme = { headFontFace: FONT, bodyFontFace: FONT };
pres.title = "AI 표정 감지로 만드는 독거노인 안심케어";
const C = pres.SchemeColor;

// ---------- layouts ----------
pres.defineSlideMaster({
  title: "COVER",
  background: { color: "E5F4EC" },
  objects: [
    { placeholder: { options: { name: "kicker", type: "body", x: 0.8, y: 1.3, w: 11.5, h: 0.45, fontSize: 18, bold: true, color: C.accent1, margin: 0 }, text: "" } },
    { placeholder: { options: { name: "title", type: "title", x: 0.8, y: 1.85, w: 11.5, h: 1.9, fontSize: 48, bold: true, color: C.text1, valign: "top", align: "left", margin: 0 }, text: "" } },
  ],
});
pres.defineSlideMaster({
  title: "CONTENT",
  background: { color: "FFFFFF" },
  objects: [
    { placeholder: { options: { name: "title", type: "title", x: 0.6, y: 0.82, w: 12.1, h: 0.65, fontSize: 32, bold: true, color: C.text1, margin: 0, valign: "middle", align: "left" }, text: "" } },
    { placeholder: { options: { name: "sub", type: "body", x: 0.6, y: 1.48, w: 8.4, h: 0.4, fontSize: 16, bold: true, color: C.accent1, margin: 0, valign: "middle" }, text: "" } },
  ],
  slideNumber: { x: 12.4, y: 7.0, w: 0.5, h: 0.3, fontSize: 10, color: "9CA3AF", align: "right" },
});

function contentSlide(section, step, title, sub) {
  const s = pres.addSlide({ masterName: "CONTENT", sectionTitle: section });
  const w = 0.3 + [...step].reduce((a, ch) => a + (ch.charCodeAt(0) > 0x1100 ? 0.17 : 0.09), 0);
  s.addText(step, {
    shape: pres.shapes.ROUNDED_RECTANGLE, rectRadius: 0.15, x: 0.6, y: 0.4, w, h: 0.32,
    fill: { color: C.accent1 }, color: C.background1, fontSize: 12, bold: true, align: "center", margin: 0,
    isTextBox: true, objectName: "step",
  });
  s.addText(title, { placeholder: "title" });
  if (sub) s.addText(sub, { placeholder: "sub" });
  return s;
}

function card(s, x, y, w, h, heading, items, opt = {}) {
  const fs = opt.fs || 15;
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w, h, rectRadius: 0.12, fill: { color: C.background2 }, line: { type: "none" }, objectName: "card" });
  s.addText(heading, { x: x + 0.25, y: y + 0.18, w: w - 0.5, h: 0.4, bold: true, color: C.accent1, fontSize: 18, margin: 0, valign: "middle", isTextBox: true, objectName: "card heading" });
  let cy = y + 0.7;
  const tx = x + 0.5, tW = w - 0.75;
  const lineH = (fs * 1.3) / 72;
  for (const it of items) {
    const parts = (Array.isArray(it) ? it : [it]).map((p) => (typeof p === "string" ? { text: p, options: {} } : p));
    const len = parts.reduce((a, p) => a + tw(p.text) * (fs / 11), 0);
    const lines = Math.max(1, Math.ceil(len / (tW * 0.95)));
    const ih = lines * lineH + 0.04;
    s.addShape(pres.shapes.OVAL, { x: x + 0.28, y: cy + lineH / 2 - 0.045, w: 0.09, h: 0.09, fill: { color: C.accent1 }, line: { type: "none" }, objectName: "card dot" });
    s.addText(parts.map((p) => ({ text: p.text, options: { ...p.options } })), {
      x: tx, y: cy, w: tW, h: ih, fontSize: fs, color: C.text1, margin: 0, valign: "top", lineSpacingMultiple: 1.05, isTextBox: true, objectName: "card item",
    });
    cy += ih + 0.1;
  }
}
const B = (t) => ({ text: t, options: { bold: true } });

// ---------- MakeCode block renderer ----------
const BH = 0.28, GAP = 0.03, BAR = 0.2, FS = 11;
function tw(str) {
  let w = 0;
  for (const ch of str) w += ch.charCodeAt(0) > 0x1100 ? 0.158 : (ch === " " ? 0.045 : 0.088);
  return w;
}
// segment: "text" | ["text", "v"|"h"|"k"|"f"]  (variable / husky reporter / grey reporter / function)
const PILL = { v: "EF4B6B", h: "12BB18", k: "C9CED6", f: "5B76F0" };
function runsOf(segs) {
  return segs.map((sg) => {
    if (typeof sg === "string") return { text: sg, options: {} };
    const [t, kind] = sg;
    return { text: ` ${t} `, options: { highlight: PILL[kind], color: kind === "k" ? "1F2937" : "FFFFFF" } };
  });
}
function segWidth(segs) {
  return segs.reduce((a, sg) => a + (typeof sg === "string" ? tw(sg) : tw(` ${sg[0]} `) + 0.04), 0) * 1.06 + 0.24;
}
const b = (color, ...segs) => ({ t: "b", color, segs });
const c = (color, hd, body, mid, body2, mid2, body3) => ({ t: "c", color, hd, body, mid, body2, mid2, body3 });

function layoutBlocks(node, x, y, out) {
  if (node.t === "b") {
    out.push({ kind: "label", x, y, w: segWidth(node.segs), h: BH, color: node.color, segs: node.segs });
    return y + BH + GAP;
  }
  const bar = { kind: "bar", x, y: y + BH / 2, w: BAR, h: 0, color: node.color };
  out.push(bar);
  out.push({ kind: "label", x, y, w: segWidth(node.hd), h: BH, color: node.color, segs: node.hd });
  y += BH + GAP;
  const bodies = [[node.body, null], [node.body2, node.mid], [node.body3, node.mid2]];
  for (const [body, mid] of bodies) {
    if (!body) continue;
    if (mid) {
      out.push({ kind: "label", x, y, w: segWidth(mid), h: BH, color: node.color, segs: mid });
      y += BH + GAP;
    }
    for (const ch of body) y = layoutBlocks(ch, x + BAR, y, out);
  }
  out.push({ kind: "foot", x, y, w: 1.1, h: 0.13, color: node.color });
  bar.h = y + 0.06 - bar.y;
  return y + 0.13 + GAP;
}
function drawBlocks(s, node, x, y) {
  const out = [];
  const end = layoutBlocks(node, x, y, out);
  for (const o of out) {
    if (o.kind === "bar") s.addShape(pres.shapes.RECTANGLE, { x: o.x, y: o.y, w: o.w, h: o.h, fill: { color: o.color }, line: { type: "none" }, objectName: "block bar" });
    else if (o.kind === "foot") s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: o.x, y: o.y, w: o.w, h: o.h, rectRadius: 0.04, fill: { color: o.color }, line: { type: "none" }, objectName: "block foot" });
    else s.addText(runsOf(o.segs), {
      shape: pres.shapes.ROUNDED_RECTANGLE, rectRadius: 0.05, x: o.x, y: o.y, w: o.w, h: o.h,
      fill: { color: o.color }, color: "FFFFFF", bold: true, fontSize: FS, margin: 5, valign: "middle",
      isTextBox: true, objectName: "block",
    });
  }
  return end;
}
function legend(s, items, y = 1.5) {
  let x = 12.75;
  for (let i = items.length - 1; i >= 0; i--) {
    const [name, col] = items[i];
    const w = tw(name) * 0.8 + 0.1;
    x -= w;
    s.addText(name, { x, y, w, h: 0.25, fontSize: 10, color: C.text1, margin: 0, valign: "middle", isTextBox: true });
    x -= 0.24;
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y: y + 0.04, w: 0.17, h: 0.17, rectRadius: 0.03, fill: { color: col }, line: { type: "none" } });
    x -= 0.2;
  }
}
function notes(s, x, y, w, lines) {
  s.addText(lines.map((l, i) => ({ text: l, options: { breakLine: i < lines.length - 1 } })), {
    x, y, w, h: lines.length * 0.3, fontSize: 12, italic: true, color: C.text2, margin: 0, valign: "top", isTextBox: true, objectName: "notes",
  });
}

// =========================================================
// 1. 표지
pres.addSection({ title: "표지" });
{
  const s = pres.addSlide({ masterName: "COVER", sectionTitle: "표지" });
  s.addText("micro:bit × HUSKYLENS 2 · AI 표정 인식", { placeholder: "kicker" });
  s.addText("AI 표정 감지로 만드는\n독거노인 안심케어", { placeholder: "title" });
  s.addText("슬픈 표정이 오래 이어지거나 얼굴이 한동안 보이지 않으면 보호자에게 알려요", {
    x: 0.8, y: 3.85, w: 11.5, h: 0.45, fontSize: 18, color: C.text1, margin: 0, isTextBox: true,
  });
  const steps = ["① 문제 상황", "② 자연어 묶음", "③ 블록 + 타이머"];
  steps.forEach((t, i) => {
    const x = 0.8 + i * 3.2;
    s.addText(t, {
      shape: pres.shapes.ROUNDED_RECTANGLE, rectRadius: 0.25, x, y: 4.75, w: 2.5, h: 0.55,
      fill: { color: C.background1 }, line: { color: C.accent1, width: 2 }, color: C.accent1,
      fontSize: 17, bold: true, align: "center", valign: "middle", margin: 0, isTextBox: true, objectName: "flow step",
    });
    if (i < 2) s.addText("→", { x: x + 2.55, y: 4.75, w: 0.6, h: 0.55, fontSize: 22, color: C.text2, align: "center", valign: "middle", margin: 0, isTextBox: true });
  });
  s.addText("준비물: micro:bit 2대(어르신용·보호자용), HUSKYLENS 2, 확장 보드 · MakeCode", {
    x: 0.8, y: 5.9, w: 11.5, h: 0.4, fontSize: 15, color: C.text2, margin: 0, isTextBox: true,
  });
}

// 2. 문제 상황
pres.addSection({ title: "① 문제 상황" });
{
  const s = contentSlide("① 문제 상황", "① 문제 상황", "혼자 계신 할머니, 오늘은 괜찮으실까?", "문제를 구체적인 장면으로 바꿔 봅시다");
  const R = (t) => ({ text: t, options: { bold: true, color: C.accent4 } });
  const bubbles = [
    ["김 할머니는 혼자 사십니다. 멀리 사는 손녀는 매일 전화를 드리지만, 할머니는 ", R("“괜찮다”"), "고만 하십니다."],
    ["그런데 요즘 할머니가 거실에 앉아 ", R("오랫동안 슬픈 표정"), "으로 계시거나, ", R("하루 종일 거실에 나오지 않는"), " 날이 있다고 합니다."],
    ["손녀가 매 순간 지켜볼 수는 없습니다. 하지만 ", R("표정과 움직임의 변화"), "를 알 수 있다면 더 빨리 연락하거나 찾아갈 수 있습니다."],
  ];
  bubbles.forEach((parts, i) => {
    const y = 2.1 + i * 1.08;
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.6, y, w: 6.9, h: 0.95, rectRadius: 0.1, fill: { color: ROW }, line: { type: "none" } });
    s.addText(parts.map((p) => (typeof p === "string" ? { text: p, options: {} } : p)), {
      x: 0.85, y, w: 6.4, h: 0.95, fontSize: 15, color: C.text1, valign: "middle", margin: 0, isTextBox: true,
    });
  });
  s.addText("❓ 카메라가 표정을 보고, 이상한 상태가 일정 시간 이상 이어질 때만 알려 줄 수 있을까?", {
    shape: pres.shapes.ROUNDED_RECTANGLE, rectRadius: 0.1, x: 0.6, y: 5.4, w: 6.9, h: 0.95,
    fill: { color: C.background1 }, line: { color: C.accent1, width: 2, dashType: "dash" },
    fontSize: 16, bold: true, color: C.accent1, margin: [0, 14, 0, 14], valign: "middle", isTextBox: true,
  });
  card(s, 7.9, 2.1, 4.85, 4.25, "문제 분해하기", [
    [B("무엇을 감지? "), "얼굴이 보이는지, 어떤 표정인지"],
    [B("언제 알릴까? "), "슬픈 표정이 잠깐이 아니라 계속될 때, 얼굴이 오랫동안 보이지 않을 때"],
    [B("누구에게? "), "보호자 micro:bit로 라디오 전송"],
    [B("필요한 도구 "), "“얼마나 오래?” → ", B("타이머")],
  ]);
  s.addText("※ 개인정보 보호: 영상은 저장하지 않고 표정 결과(번호)만 사용합니다.", {
    x: 0.6, y: 6.55, w: 12, h: 0.3, fontSize: 12, italic: true, color: C.text2, margin: 0, isTextBox: true,
  });
}

// 3. 해결 아이디어
{
  const s = contentSlide("① 문제 상황", "① 문제 상황 → 해결 아이디어", "입력 → 처리 → 출력으로 정리하기", "허스키렌즈2의 “표정 인식(Emotion Recognition)” 기능을 사용합니다");
  const boxes = [
    ["입력", "허스키렌즈2가 얼굴을 찾고 표정 번호(ID)를 알려 줌", K.husky],
    ["처리 (판단 + 타이머)", "슬픈 표정이 몇 초째인지, 얼굴이 안 보인 지 몇 초째인지 계산", K.logic],
    ["출력", "LED 아이콘·소리, 보호자에게 라디오 “SAD” / “CHECK”", K.radio],
  ];
  boxes.forEach(([h, p, col], i) => {
    const x = 0.6 + i * 4.15;
    s.addText([
      { text: h, options: { bold: true, fontSize: 18, breakLine: true, paraSpaceAfter: 4 } },
      { text: p, options: { fontSize: 14 } },
    ], {
      shape: pres.shapes.ROUNDED_RECTANGLE, rectRadius: 0.12, x, y: 2.05, w: 3.7, h: 1.2,
      fill: { color: col }, color: "FFFFFF", margin: [0, 14, 0, 14], valign: "middle", isTextBox: true, objectName: "ipo box",
    });
    if (i < 2) s.addText("→", { x: x + 3.7, y: 2.05, w: 0.45, h: 1.2, fontSize: 24, color: C.text2, align: "center", valign: "middle", margin: 0, isTextBox: true });
  });
  const hdr = (t) => ({ text: t, options: { bold: true, color: "FFFFFF", fill: { color: THEME.colors.accent1 } } });
  const rows = [
    ["웃음 · 보통 표정", "아니요", "슬픔 타이머 리셋", "웃는 얼굴"],
    ["슬픈 표정 (잠깐)", "아니요", "10초 미만", "—"],
    ["슬픈 표정 (계속)", "예", "10초 이상", "슬픈 얼굴 · 소리 · “SAD”"],
    ["얼굴 안 보임 (계속)", "예", "30초 이상", "“?” · “CHECK”"],
  ];
  s.addTable([
    ["상황", "바로 알릴까?", "타이머 조건", "출력"].map(hdr),
    ...rows.map((r, i) => r.map((t, j) => ({ text: t, options: { bold: j === 1 && t === "예", fill: { color: i % 2 ? ROW : "FFFFFF" } } }))),
  ], {
    x: 0.6, y: 3.6, w: 7.4, colW: [1.95, 1.3, 1.6, 2.55], rowH: 0.5, fontSize: 13, color: C.text1,
    valign: "middle", margin: 0.08, border: { type: "none" },
  });
  card(s, 8.35, 3.6, 4.4, 2.9, "왜 타이머가 필요할까?", [
    "사람은 잠깐 찡그리거나 고개를 돌리기도 합니다.",
    ["순간 표정만 보고 알리면 ", B("잘못된 알림"), "이 많아집니다."],
    ["“", B("얼마나 오래"), " 계속되었나”를 재면 믿을 수 있는 알림이 됩니다."],
  ], { fs: 14 });
}

// 4. 타이머 원리
pres.addSection({ title: "② 자연어 묶음" });
{
  const s = contentSlide("② 자연어 묶음", "② 자연어 묶음 준비 · 타이머 원리", "MakeCode 타이머: “시작 시각”을 기억하기", "「실행 시간(ms)」은 micro:bit가 켜진 뒤 흐른 시간입니다 (1초 = 1000ms)");
  s.addText([
    { text: "경과 시간 = ", options: {} },
    { text: "실행 시간(ms)", options: { color: C.accent4 } },
    { text: " − ", options: {} },
    { text: "슬픔시작", options: { color: C.accent4 } },
    { text: "   →   10000보다 크면 알림", options: {} },
  ], {
    shape: pres.shapes.ROUNDED_RECTANGLE, rectRadius: 0.12, x: 0.6, y: 2.05, w: 12.15, h: 0.65,
    fill: { color: ROW }, fontSize: 20, bold: true, color: C.text1, align: "center", valign: "middle", margin: 0, isTextBox: true,
  });
  // timeline
  const X0 = 0.8, W = 11.8, ty = 3.55;
  const px = (f) => X0 + W * f;
  s.addShape(pres.shapes.RECTANGLE, { x: X0, y: ty + 0.12, w: W, h: 0.04, fill: { color: "CBD5E1" }, line: { type: "none" } });
  const segs = [[0.06, 0.16, K.husky, "보통 표정"], [0.22, 0.09, "F59E0B", "슬픔"], [0.31, 0.10, K.husky, "보통"], [0.41, 0.46, K.vari, "슬픈 표정 계속 ……"]];
  for (const [f, w, col, t] of segs) {
    s.addText(t, { shape: pres.shapes.ROUNDED_RECTANGLE, rectRadius: 0.06, x: px(f), y: ty, w: W * w, h: 0.28, fill: { color: col }, color: "FFFFFF", bold: true, fontSize: 12, align: "center", valign: "middle", margin: 0, isTextBox: true });
  }
  const marks = [[0.22, "슬픔시작 = 실행시간", "4초 만에 끝 → 리셋(0)", C.text1], [0.41, "슬픔시작 = 실행시간", "다시 타이머 시작", C.text1], [0.73, "경과 > 10초 → 알림!", "“SAD” 전송", C.accent4]];
  for (const [f, top, bot, col] of marks) {
    s.addShape(pres.shapes.LINE, { x: px(f), y: ty - 0.6, w: 0, h: 1.4, line: { color: col === C.accent4 ? THEME.colors.accent4 : THEME.colors.dk1, width: 1.5 } });
    s.addText(top, { x: px(f) + 0.06, y: ty - 0.62, w: 2.4, h: 0.28, fontSize: 12, bold: true, color: col, margin: 0, isTextBox: true });
    s.addText(bot, { x: px(f) + 0.06, y: ty + 0.5, w: 2.4, h: 0.28, fontSize: 11, color: C.text2, margin: 0, isTextBox: true });
  }
  card(s, 0.6, 4.5, 5.9, 2.5, "세 가지 규칙만 기억하세요", [
    [B("시작 — "), "슬픔이 처음 보이면(슬픔시작 = 0일 때) 지금 시각을 저장"],
    [B("계산 — "), "슬픔이 계속되면 지금 − 슬픔시작으로 경과 시간 계산"],
    [B("리셋 — "), "다른 표정이 나오면 슬픔시작 = 0으로 되돌림"],
  ], { fs: 14 });
  card(s, 6.85, 4.5, 5.9, 2.5, "같은 방법으로 “부재 타이머”", [
    ["얼굴이 보일 때마다 ", B("마지막확인 = 지금 시각")],
    ["얼굴이 안 보이면 ", B("지금 − 마지막확인"), "이 30초를 넘는지 확인"],
    "수업에서는 10초·30초, 실제로는 수십 분~몇 시간으로 늘려요",
  ], { fs: 14 });
}

// 5. 자연어 묶음
{
  const s = contentSlide("② 자연어 묶음", "② 자연어 묶음", "자연어 알고리즘을 4개의 묶음으로");
  const TAG = { 준비: "7B8798", 반복: "10A810", 입력: "1E90FF", 판단: "00A4A6", 저장: "DC143C", 출력: "F07020" };
  const data = [
    ["G", "묶음 ① 준비 — 시작하면"],
    ["01", "준비", 0, "I2C 시작, “표정 인식” 선택, 라디오 그룹 1"],
    ["02", "저장", 0, "슬픔시작 = 0, 마지막확인 = 지금 시각"],
    ["G", "묶음 ② 반복 — 무한반복 실행"],
    ["03", "반복", 0, "계속 반복한다"],
    ["04", "입력", 1, "표정 인식 결과를 가져온다"],
    ["05", "판단", 1, "얼굴이 보이면 → 〔표정 확인〕 묶음 실행"],
    ["06", "판단", 1, "얼굴이 없으면 → 〔부재 확인〕 묶음 실행"],
    ["G", "묶음 ③ 표정 확인 — 슬픔 타이머"],
    ["07", "저장", 0, "마지막확인 = 지금 시각, e = 가운데 얼굴의 표정 ID"],
    ["08", "판단", 0, "e가 슬픔이면"],
    ["09", "저장", 1, "슬픔시작이 0이면 슬픔시작 = 지금 시각 (타이머 시작)"],
    ["10", "판단", 1, "지금 − 슬픔시작 > 10초 이면"],
    ["11", "출력", 2, "슬픈 아이콘, 소리, 라디오 “SAD”"],
    ["12", "저장", 0, "아니면(웃음·보통) 슬픔시작 = 0 (리셋), 웃는 아이콘"],
    ["G", "묶음 ④ 부재 확인 — 부재 타이머"],
    ["13", "판단", 0, "지금 − 마지막확인 > 30초 이면"],
    ["14", "출력", 1, "“?” 표시, 라디오 “CHECK”"],
    ["15", "출력", 0, "아니면 LED 지우기"],
  ];
  let y = 1.6, odd = true;
  const RW = 7.9;
  for (const r of data) {
    if (r[0] === "G") {
      s.addText(r[1], { x: 0.65, y: y + 0.03, w: RW, h: 0.22, fontSize: 11, bold: true, color: C.accent1, margin: 0, valign: "bottom", isTextBox: true, objectName: "group" });
      y += 0.26;
      odd = true;
      continue;
    }
    const [no, tag, ind, text] = r;
    const h = 0.265;
    if (odd) s.addShape(pres.shapes.RECTANGLE, { x: 0.6, y, w: RW, h, fill: { color: ROW }, line: { type: "none" } });
    odd = !odd;
    s.addText(no, { x: 0.68, y, w: 0.35, h, fontSize: 10, color: "94A3B8", margin: 0, valign: "middle", isTextBox: true });
    const tx = 1.05 + ind * 0.6;
    s.addText(tag, { shape: pres.shapes.RECTANGLE, x: tx, y: y + 0.03, w: 0.55, h: h - 0.06, fill: { color: TAG[tag] }, color: "FFFFFF", fontSize: 10, bold: true, align: "center", valign: "middle", margin: 0, isTextBox: true });
    s.addText(text, { x: tx + 0.67, y, w: RW - (tx + 0.67 - 0.6), h, fontSize: 14, color: C.text1, margin: 0, valign: "middle", isTextBox: true });
    y += h;
  }
  card(s, 8.85, 1.6, 3.9, 3.6, "핵심 아이디어", [
    ["긴 알고리즘은 ", B("역할별 묶음"), "으로 나누면 읽기 쉽습니다. 묶음 ③·④는 MakeCode의 ", B("함수"), "가 됩니다."],
    ["슬픔 표정 ID는 먼저 ", B("수 출력 e"), "로 직접 확인하세요. 블록의 6은 예시입니다."],
    "기준 10초·30초도 예시입니다. 측정 → 기준 정하기 → 판단!",
  ], { fs: 14 });
}

// 6. 블록 ①
pres.addSection({ title: "③ 블록 + 타이머" });
{
  const s = contentSlide("③ 블록 + 타이머", "③ 블록 + 타이머 · 1", "MakeCode 블록 ① 준비와 반복", "묶음 ① → 시작하면, 묶음 ② → 무한반복 실행");
  legend(s, [["huskylens2", K.husky], ["기본", K.basic], ["논리", K.logic], ["변수", K.vari], ["라디오", K.radio], ["함수", K.func]]);
  const end = drawBlocks(s, c(K.basic, ["시작하면"], [
    b(K.husky, "initialize I2C until success"),
    b(K.husky, "switch algorithm  emotion recognition"),
    b(K.radio, "라디오 그룹을 1 로 설정"),
    b(K.vari, ["슬픔시작", "v"], " 에 0 저장"),
    b(K.vari, ["마지막확인", "v"], " 에 ", ["실행 시간(ms)", "k"], " 저장"),
  ]), 0.6, 2.1);
  notes(s, 0.6, end + 0.3, 4.3, [
    "※ 1단계: 수 출력 e로 표정마다 ID를 확인",
    "※ 실행 시간(ms)은 ‘입력’ 메뉴 아래쪽에 있어요",
    "※ 마지막확인을 지금 시각으로 시작해야",
    "    켜자마자 “CHECK”가 나가지 않아요",
  ]);
  drawBlocks(s, c(K.basic, ["무한반복 실행"], [
    b(K.husky, "get emotion recognition result"),
    c(K.logic, ["만약 ", ["available emotion recognition", "h"], " 이면"],
      [b(K.func, "함수 호출 ", ["표정확인", "f"])],
      ["아니면"],
      [b(K.func, "함수 호출 ", ["부재확인", "f"])]),
    b(K.basic, "일시중지 200 (ms)"),
  ]), 4.95, 2.1);
  card(s, 9.0, 2.1, 3.75, 3.4, "묶음 → 함수", [
    ["무한반복 안이 짧아져서 ", B("전체 흐름"), "이 한눈에 보여요."],
    "함수는 ‘고급 → 함수 → 함수 만들기’에서 만들어요.",
    "200ms 쉬기: 허스키렌즈가 결과를 갱신할 시간을 줍니다.",
  ], { fs: 14 });
}

// 7. 블록 ②
{
  const s = contentSlide("③ 블록 + 타이머", "③ 블록 + 타이머 · 2", "MakeCode 블록 ② 두 개의 타이머 함수", "묶음 ③ → 함수 표정확인,  묶음 ④ → 함수 부재확인");
  legend(s, [["huskylens2", K.husky], ["기본", K.basic], ["논리", K.logic], ["변수", K.vari], ["음악", K.music], ["라디오", K.radio]]);
  drawBlocks(s, c(K.func, ["함수 표정확인"], [
    b(K.vari, ["마지막확인", "v"], " 에 ", ["실행 시간(ms)", "k"], " 저장"),
    b(K.vari, ["e", "v"], " 에 ", ["emotion nearest to center ID", "h"], " 저장"),
    c(K.logic, ["만약 ", ["e", "v"], " = 6 이면"], [
      c(K.logic, ["만약 ", ["슬픔시작", "v"], " = 0 이면"], [
        b(K.vari, ["슬픔시작", "v"], " 에 ", ["실행 시간(ms)", "k"], " 저장"),
      ]),
      c(K.logic, ["만약 ", ["실행 시간(ms)", "k"], " − ", ["슬픔시작", "v"], " > 10000 이면"], [
        b(K.basic, "아이콘 출력  Sad"),
        b(K.music, "멜로디 재생  power down"),
        b(K.radio, "라디오 전송: 문자열 \"SAD\""),
      ]),
    ], ["아니면"], [
      b(K.vari, ["슬픔시작", "v"], " 에 0 저장"),
      b(K.basic, "아이콘 출력  Happy"),
    ]),
  ]), 0.6, 2.05);
  const end = drawBlocks(s, c(K.func, ["함수 부재확인"], [
    c(K.logic, ["만약 ", ["실행 시간(ms)", "k"], " − ", ["마지막확인", "v"], " > 30000 이면"], [
      b(K.basic, "문자열 출력 \"?\""),
      b(K.radio, "라디오 전송: 문자열 \"CHECK\""),
    ], ["아니면"], [
      b(K.basic, "LED 스크린 지우기"),
    ]),
  ]), 6.9, 2.05);
  notes(s, 6.9, end + 0.35, 5.85, [
    "※ 6은 예시 ID — 내 허스키렌즈2에서 슬픈 표정의 ID로 바꾸세요",
    "※ 10000ms = 10초, 30000ms = 30초 (수업용 기준값)",
    "※ 다른 표정이 나오면 슬픔시작 = 0 → 타이머 리셋",
    "※ 알림이 계속 반복되면? 알림 후 슬픔시작 = 0으로 바꿔 보기",
  ]);
}

// 8. 보호자 + 테스트
{
  const s = contentSlide("③ 블록 + 타이머", "③ 블록 + 타이머 · 3", "보호자 micro:bit와 테스트 계획", "두 번째 micro:bit는 라디오로 받은 문자열에 따라 알려 줍니다");
  const end = drawBlocks(s, c(K.radio, ["라디오 수신하면 ", ["receivedString", "v"]], [
    c(K.logic, ["만약 ", ["receivedString", "v"], " = \"SAD\" 이면"], [
      b(K.basic, "아이콘 출력  Sad"),
      b(K.music, "멜로디 재생  ba ding"),
    ], ["아니면서 만약 ", ["receivedString", "v"], " = \"CHECK\" 이면"], [
      b(K.basic, "문자열 출력 \"?\""),
      b(K.music, "음 재생  중간 C 1 박자"),
    ]),
  ]), 0.6, 2.1);
  notes(s, 0.6, end + 0.3, 4.6, ["※ 두 micro:bit의 라디오 그룹 번호를 같게(1) 맞추세요"]);
  card(s, 5.6, 2.1, 7.15, 2.45, "테스트 체크리스트", [
    "웃는 얼굴 → 웃는 아이콘, 알림 없음",
    "슬픈 표정 5초 → 다시 웃기 → 알림 없음 (타이머 리셋 확인)",
    "슬픈 표정 10초 이상 → 보호자 쪽 Sad + 소리",
    "카메라 앞에서 30초 이상 사라지기 → 보호자 쪽 “?”, 다시 나타나면 멈춤",
  ], { fs: 14 });
  card(s, 5.6, 4.8, 7.15, 1.95, "더 생각해 보기", [
    "슬픔 말고 화남·놀람도 함께 감지하려면?",
    "밤 시간에는 부재 알림을 끄려면 무엇이 더 필요할까?",
    "카메라만 믿어도 될까? 버튼 “괜찮아요” 응답을 더한다면?",
  ], { fs: 14 });
}

(async () => {
  await pres.writeFile({ fileName: OUT });
  await applyTheme(OUT, THEME);
  console.log("wrote", OUT);
})();
