// 柑橘病蟲害辨識 — 口頭報告簡報產生器（v12.3，09-15 週會）
// 數據來源：docs/v12.3_報告_口頭簡報.md（09-14 含 662 補量版）、docs/v12.3_報告_階段性進度週報.md、
//          docs/v12.3_結果_三平台Benchmark整合.md、docs/v12.3_結果_天璣8300即時辨識門檻.md、
//          docs/v12.3_結果_Snapdragon8Gen2即時辨識門檻.md、docs/v5.7_結果_兩臂訓練評估.md §3、
//          docs/v5.7_報告_交付與研究限制.md §1、docs/v5.7_結果_資料集分析與驗收.md §2.2、docs/v5.6_結果_資料集分析與驗收.md §1–2
// 樣式沿用 tools/build_deck.js（v12.1）；上週的產生器不動，以便重建上週的簡報。
const pptxgen = require("pptxgenjs");

// ---- 調色盤：柑橘園（深葉綠主導 + 果橙點綴），與 v12.1 相同 ------------------
const C = {
  dark:   "0F2E22",  // 深葉綠（暗底）
  dark2:  "17402F",  // 暗底上的卡片
  mid:    "2C5F45",  // 中綠
  tint:   "EAF2EC",  // 淡綠卡片底
  tint2:  "F6F9F7",  // 更淡
  rose:   "F7E7E2",  // 淡陶紅卡片底（負向）
  peach:  "FBEEDF",  // 淡橙卡片底（重點）
  accent: "E8963C",  // 果橙（唯一銳利重點色）
  clay:   "B3453A",  // 陶紅（負向 / 未通過）
  white:  "FFFFFF",
  ink:    "1A2620",  // 正文
  muted:  "5E6F66",  // 次要文字
  line:   "C9DCD0",
  sage:   "A9C4B4",  // 暗底上的次要文字
};
const F = "Microsoft JhengHei";

const pres = new pptxgen();
pres.layout = "LAYOUT_WIDE";           // 13.333 x 7.5 吋
pres.author = "Claude Code";
pres.title = "柑橘病蟲害辨識 階段性進度報告 v12.3";

const W = 13.333, H = 7.5, M = 0.7;    // 邊界 0.7"
const CW = W - 2 * M;                  // 可用寬度 11.933"

// ---- 共用元件（與 v12.1 相同）---------------------------------------------
function darkBg(s) { s.background = { color: C.dark }; }
function lightBg(s) { s.background = { color: C.white }; }

function title(s, text, sub) {
  s.addText(text, {
    x: M, y: 0.45, w: CW, h: 0.72, isTextBox: true, margin: 0,
    fontFace: F, fontSize: 32, bold: true, color: C.dark, align: "left",
  });
  if (sub) {
    s.addText(sub, {
      x: M, y: 1.19, w: CW, h: 0.42, isTextBox: true, margin: 0,
      fontFace: F, fontSize: 15, color: C.muted, align: "left",
    });
  }
}

function kicker(s, text, dark) {
  s.addText(text, {
    x: M, y: 0.45, w: CW, h: 0.4, isTextBox: true, margin: 0,
    fontFace: F, fontSize: 14, color: C.accent, charSpacing: 2,
  });
}

function badge(s, n, x, y, d, fill, txtColor) {
  s.addShape(pres.ShapeType.ellipse, { x, y, w: d, h: d, fill: { color: fill || C.mid } });
  s.addText(String(n), {
    x, y, w: d, h: d, isTextBox: true, margin: 0,
    fontFace: F, fontSize: Math.round(d * 26), bold: true,
    color: txtColor || C.white, align: "center", valign: "middle",
  });
}

function card(s, x, y, w, h, fill) {
  s.addShape(pres.ShapeType.roundRect, {
    x, y, w, h, rectRadius: 0.1,
    fill: { color: fill || C.tint }, line: { color: fill || C.tint, width: 0 },
  });
}

function pageNote(s, txt) {
  s.addText(txt, {
    x: M, y: H - 0.62, w: CW, h: 0.32, isTextBox: true, margin: 0,
    fontFace: F, fontSize: 10, color: C.muted, align: "left",
  });
}

const tblBase = {
  fontFace: F, fontSize: 12, color: C.ink, border: { type: "solid", color: C.line, pt: 0.5 },
  valign: "middle", autoPage: false,
};
function hdr(t) {
  return { text: t, options: { bold: true, color: C.white, fill: { color: C.mid }, fontSize: 12 } };
}
function cell(t, o) { return { text: t, options: o || {} }; }

// 圖表共用的安靜外框（每次呼叫都回傳新物件，pptxgenjs 會就地改寫選項）
function quietChart(extra) {
  return Object.assign({
    catAxisLabelColor: C.ink, catAxisLabelFontSize: 12, catAxisLabelFontFace: F,
    valAxisLabelColor: C.muted, valAxisLabelFontSize: 10, valAxisLabelFontFace: F,
    valGridLine: { color: "E4EDE7", size: 1 }, catGridLine: { style: "none" },
    showLegend: true, legendPos: "t", legendFontSize: 12, legendColor: C.ink, legendFontFace: F,
    showTitle: false,
  }, extra);
}

// =====================================================================
// 1. 封面
// =====================================================================
{
  const s = pres.addSlide(); darkBg(s);
  s.addText("柑橘病蟲害辨識　115 資工四A", {
    x: M, y: 1.35, w: CW, h: 0.38, isTextBox: true, margin: 0,
    fontFace: F, fontSize: 15, color: C.accent, charSpacing: 2,
  });
  s.addText("階段性進度報告", {
    x: M, y: 1.78, w: CW, h: 1.0, isTextBox: true, margin: 0,
    fontFace: F, fontSize: 46, bold: true, color: C.white,
  });
  s.addText("v12.1 → v12.3　·　2026-09-15　週二下午進度會議", {
    x: M, y: 2.82, w: CW, h: 0.4, isTextBox: true, margin: 0,
    fontFace: F, fontSize: 16, color: C.sage,
  });

  card(s, M, 3.62, CW, 2.0, C.dark2);
  s.addText("這一週把還能改進的每一條路，都用事先寫好的標準量完了。", {
    x: M + 0.45, y: 3.86, w: CW - 0.9, h: 0.5, isTextBox: true, margin: 0,
    fontFace: F, fontSize: 20, bold: true, color: C.white,
  });
  s.addText("標註的問題找到根因並修好、交付權重改用 v5.7 資料集重訓、兩台高階裝置上即時辨識成立——但都只到 320 解析度；不新拍照片的前提下，工程手段已經用盡。", {
    x: M + 0.45, y: 4.46, w: CW - 0.9, h: 0.95, isTextBox: true, margin: 0,
    fontFace: F, fontSize: 18, bold: true, color: C.accent, valign: "top",
  });

  s.addText("正文約 8 分鐘　·　備答另計", {
    x: M, y: 6.05, w: CW, h: 0.35, isTextBox: true, margin: 0,
    fontFace: F, fontSize: 12, color: "7E9788",
  });
  s.addNotes("【約 30 秒】開場先講這句話，再進三個數字。全場骨架：標註修好 → 交付改用 v5.7 資料集重訓（但模型沒變好）→ 即時辨識變成分裝置 → 三個要當場定的決策。");
}

// =====================================================================
// 2. 三個數字
// =====================================================================
{
  const s = pres.addSlide(); lightBg(s);
  title(s, "開場：三個數字", "這一週的結論可以壓縮成這三個");

  const items = [
    { n: "5 → 0", size: 54, unit: " 張", lab: "薊馬葉害兩人各標同樣 20 張，框數不一致的照片數",
      so: "改成一張葉子一個框後分歧消失\n兩人畫框重疊度 0.746 → 0.899\n類別保留、全類重標", col: C.mid },
    { n: "−0.002", size: 54, lab: "交付權重改用 v5.7 資料集訓練後，標註沒變的八類平均 AP50 變化",
      so: "換資料集是讓數字可信\n模型並沒有變好", col: C.clay },
    { n: "38/30/16", size: 40, unit: " FPS", lab: "320 解析度連續跑 5 分鐘後\n平板、手機、662",
      so: "兩台高階機成立、662 不成立\n下一個解析度連續跑都不行", col: C.accent },
  ];
  const gap = 0.42, cw = (CW - 2 * gap) / 3;
  items.forEach((it, i) => {
    const x = M + i * (cw + gap);
    card(s, x, 1.85, cw, 3.8, C.tint);
    s.addText([
      { text: it.n, options: { fontSize: it.size, bold: true, color: it.col } },
      { text: it.unit || "", options: { fontSize: 20, bold: true, color: it.col } },
    ], {
      x: x + 0.3, y: 2.1, w: cw - 0.6, h: 1.0, isTextBox: true, margin: 0,
      fontFace: F, align: "left", valign: "bottom",
    });
    s.addText(it.lab, {
      x: x + 0.3, y: 3.2, w: cw - 0.6, h: 0.85, isTextBox: true, margin: 0,
      fontFace: F, fontSize: 14, color: C.ink, valign: "top",
    });
    s.addShape(pres.ShapeType.rect, { x: x + 0.3, y: 4.15, w: cw - 0.6, h: 0.012, fill: { color: C.line } });
    s.addText(it.so, {
      x: x + 0.3, y: 4.3, w: cw - 0.6, h: 1.1, isTextBox: true, margin: 0,
      fontFace: F, fontSize: 14, bold: true, color: C.mid, valign: "top",
    });
  });
  s.addNotes("【約 45 秒】三個數字講完就往下走，不要在這裡展開。5 → 0 是標註（兩人框數不一致的照片數。刻意不放 0.899：那是畫框重疊度、不是準確度，放成大數字容易被誤讀）、−0.002 是模型、38/30/16 是部署，後面各有一頁。");
}

// =====================================================================
// 3. 進度：四條線
// =====================================================================
{
  const s = pres.addSlide(); lightBg(s);
  title(s, "進度：四條線的收斂狀態", "與報告庫的 09-13 週報表 6-1、成果總覽 SUMMARY.md 同一套定義");

  // v = 結論（加粗），p = 依據（小字）
  const rows = [
    { k: "模型端", v: "連續四輪、十一種改動全部無效",
      p: "v9 六臂消融零臂通過 2σ 門檻；v12s 放大 3.84 倍參數只換到 +0.003" },
    { k: "訓練協定", v: "有效，但修的是可報告的評估精度",
      p: "固定輪數並報 last.pt，valid 可與 test 併計，逐類 ±2SE 最差值 0.140 → 0.099" },
    { k: "部署端", v: "662 不採用；兩台高階裝置在 320 成立——從「不採用」變成「分裝置」",
      p: "662 延遲與精度沒有交集；高階裝置 fp32@320 連續跑 38.4 / 29.8 FPS，下一個解析度被持續負載擋下", hot: true },
    { k: "資料端", v: "外部影像無效；重標有效但未讓模型變好",
      p: "外部影像 +0.021；重標後八類平均 −0.002；工作包 A、B 都已結案" },
  ];
  rows.forEach((r, i) => {
    const y = 1.8 + i * 1.1;
    card(s, M, y, CW, 0.96, r.hot ? C.tint : C.tint2);
    badge(s, i + 1, M + 0.3, y + 0.21, 0.54, r.hot ? C.accent : C.mid);
    s.addText(r.k, {
      x: M + 1.05, y: y + 0.12, w: 1.6, h: 0.38, isTextBox: true, margin: 0,
      fontFace: F, fontSize: 18, bold: true, color: C.dark,
    });
    s.addText(r.v, {
      x: M + 2.65, y: y + 0.11, w: CW - 2.95, h: 0.38, isTextBox: true, margin: 0,
      fontFace: F, fontSize: 15, bold: true, color: r.hot ? C.accent : C.mid,
    });
    s.addText(r.p, {
      x: M + 1.05, y: y + 0.54, w: CW - 1.35, h: 0.34, isTextBox: true, margin: 0,
      fontFace: F, fontSize: 12.5, color: C.muted,
    });
  });

  s.addText("四條線都被預先登記的判準關閉；兩個有效的手段修的是「數字能不能被相信」，沒有一個讓模型變好。", {
    x: M, y: 6.3, w: CW, h: 0.45, isTextBox: true, margin: 0,
    fontFace: F, fontSize: 14, color: C.ink,
  });
  s.addNotes("【約 45 秒】四條線用報告庫週報與 SUMMARY.md 的同一套定義，教授看過總覽再看投影片會對得上。每條只念加粗的結論；部署端是本週的結構性變化，特別點出來。");
}

// =====================================================================
// 4. 本期回答的七個問題（併成五列）
// =====================================================================
{
  const s = pres.addSlide(); lightBg(s);
  title(s, "本期回答的七個問題", "兩輪標註、兩台高階裝置各併成一列　·　每一個的標準都寫在量測之前，沒有一個是看完結果才決定怎麼算");

  const good = { bold: true, color: C.mid }, bad = { bold: true, color: C.clay }, hot = { bold: true, color: C.accent };
  s.addTable([
    [hdr("問題"), hdr("事先寫好的標準"), hdr("結果"), hdr("所以")],
    [cell("662 上要不要做即時辨識", { bold: true }),
     cell("每張 ≤ 40 ms（25 FPS）且精度 ≥ 0.764 且沒有一類掉超過 0.10"),
     cell("速度全過、精度全不過", bad), cell("662 只做拍照辨識")],
    [cell("薊馬葉害的框能不能畫得一致", { bold: true }),
     cell("兩人重疊度中位數 ≥ 0.85"),
     cell("0.746 → 改規則後 0.899", good), cell("類別保留，全類重標")],
    [cell("外部潛葉蛾照片有沒有用", { bold: true }),
     cell("潛葉蛾 AP50 要多 0.04 以上"),
     cell("+0.021", bad), cell("無效，不交付")],
    [cell("重標能不能解除薊馬葉害的矛盾", { bold: true }),
     cell("valid 與 test 差距 < 0.165"),
     cell("−0.202 → +0.121", good), cell("三輪不收斂的問題解除")],
    [cell("高階裝置上能不能做即時辨識", { bold: true }),
     cell("前面三條，再加「連續跑也要 ≤ 40 ms」；兩台同一份標準"),
     cell("兩台都是 320 過（38／30 FPS）；下一個解析度連續跑不過", hot), cell("只到 320")],
  ], {
    x: M, y: 1.85, w: CW, colW: [3.0, 3.65, 3.15, 2.133], rowH: 0.66,
    ...tblBase, fontSize: 12.5, align: "left",
  });

  card(s, M, 6.0, CW, 0.72, C.tint2);
  s.addText([
    { text: "判準比量測早：", options: { bold: true, color: C.dark } },
    { text: "途中補定的細節、662 的補量協定，也都在對應量測之前 commit——時間可以從 git 紀錄查到（備答有時間線）。", options: { color: C.ink } },
  ], { x: M + 0.35, y: 6.13, w: CW - 0.7, h: 0.46, isTextBox: true, margin: 0, fontFace: F, fontSize: 13.5 });
  s.addNotes("【約 1 分 15 秒】逐列念「問題 → 標準 → 結果」，不要解釋細節。強調最後一欄都是照事先寫好的後果走。被質疑時翻備答的時間線。");
}

// =====================================================================
// 5. 為什麼 0.861 不是進步
// =====================================================================
{
  const s = pres.addSlide(); lightBg(s);
  title(s, "多講一分鐘：為什麼 0.861 不是進步", "新規則讓薊馬葉害的框變大 1.36 倍，而 AP50 只要重疊 ≥ 0.5 就算對——框愈大愈容易算對");

  s.addChart(pres.ChartType.bar, [
    { name: "v11.5（v5.6 資料集，舊標註）", labels: ["九類 test mAP50", "八類平均 AP50（標註沒變）", "薊馬葉害 test AP50"],
      values: [0.80959, 0.85268, 0.4515] },
    { name: "v11.5（v5.7 資料集，新標註）", labels: ["九類 test mAP50", "八類平均 AP50（標註沒變）", "薊馬葉害 test AP50"],
      values: [0.86118, 0.85082, 0.8453] },
  ], quietChart({
    x: M, y: 1.8, w: 6.7, h: 4.15,
    barDir: "col", barGrouping: "clustered", barGapWidthPct: 55,
    chartColors: [C.line, C.mid],
    showValue: true, dataLabelPosition: "outEnd", dataLabelColor: C.ink,
    dataLabelFontSize: 11, dataLabelFontFace: F, dataLabelFormatCode: "0.000",
    catAxisLabelFontSize: 11,
    valAxisMinVal: 0, valAxisMaxVal: 1.0,
  }));

  const rx = M + 7.05, rw = CW - 7.05;
  card(s, rx, 1.8, rw, 1.25, C.tint2);
  s.addText("看起來", { x: rx + 0.3, y: 1.92, w: rw - 0.6, h: 0.3, isTextBox: true, margin: 0, fontFace: F, fontSize: 12, bold: true, color: C.muted });
  s.addText([
    { text: "0.810 → 0.861", options: { fontSize: 24, bold: true, color: C.ink } },
    { text: "　進步 0.05", options: { fontSize: 14, color: C.muted } },
  ], { x: rx + 0.3, y: 2.24, w: rw - 0.6, h: 0.6, isTextBox: true, margin: 0, fontFace: F, valign: "middle" });

  card(s, rx, 3.2, rw, 1.25, C.rose);
  s.addText("但其中 +0.044 是薊馬葉害一類貢獻的", { x: rx + 0.3, y: 3.32, w: rw - 0.6, h: 0.36, isTextBox: true, margin: 0, fontFace: F, fontSize: 14, bold: true, color: C.clay });
  s.addText("它自己 test 升了 +0.394，除以 9 類就是 +0.044。框變大、更容易命中，不是模型變好。", {
    x: rx + 0.3, y: 3.7, w: rw - 0.6, h: 0.52, isTextBox: true, margin: 0, fontFace: F, fontSize: 12.5, color: C.ink, valign: "top",
  });

  card(s, rx, 4.6, rw, 1.35, C.tint);
  s.addText("拿掉它，只看標註沒變的八類", { x: rx + 0.3, y: 4.72, w: rw - 0.6, h: 0.32, isTextBox: true, margin: 0, fontFace: F, fontSize: 12, bold: true, color: C.muted });
  s.addText([
    { text: "−0.002", options: { fontSize: 30, bold: true, color: C.mid } },
    { text: "　0.8527 → 0.8508，等於沒變", options: { fontSize: 14, color: C.ink } },
  ], { x: rx + 0.3, y: 5.08, w: rw - 0.6, h: 0.7, isTextBox: true, margin: 0, fontFace: F, valign: "middle" });

  s.addText("要報 0.861 時，一定要同時說「新標註約定下、不能跟舊數字比」。", {
    x: M, y: 6.2, w: CW, h: 0.5, isTextBox: true, margin: 0,
    fontFace: F, fontSize: 17, bold: true, color: C.accent,
  });
  s.addNotes("【約 1 分鐘】看圖：左邊九類變高、中間八類沒動、右邊薊馬葉害跳上去。那一跳是框定義換掉，不是模型能力。模型兩邊都是 v11.5，只換了資料集；換的理由是數字有解釋力，不是分數變高。");
}

// =====================================================================
// 6. 部署：三台同一套方法
// =====================================================================
{
  const s = pres.addSlide(); lightBg(s);
  title(s, "部署：三台裝置、同一套方法", "320 解析度的 FPS　·　門檻 25 FPS（每張 40 ms）　·　相機預覽最多每秒 30 張");

  s.addChart(pres.ChartType.bar, [
    { name: "短時間（3 輪最快）", labels: ["Snapdragon 662", "天璣 8300 平板", "8 Gen 2 手機"], values: [14.9, 66.6, 32.0] },
    { name: "連續跑 5 分鐘（最慢一段）", labels: ["Snapdragon 662", "天璣 8300 平板", "8 Gen 2 手機"], values: [15.8, 38.4, 29.8] },
  ], quietChart({
    x: M, y: 1.8, w: 7.1, h: 4.2,
    barDir: "col", barGrouping: "clustered", barGapWidthPct: 50,
    chartColors: [C.line, C.mid],
    showValue: true, dataLabelPosition: "outEnd", dataLabelColor: C.ink,
    dataLabelFontSize: 12, dataLabelFontFace: F, dataLabelFormatCode: "0.0",
    valAxisMinVal: 0, valAxisMaxVal: 70, valAxisMajorUnit: 10,
  }));

  const rx = M + 7.45, rw = CW - 7.45;
  const v = [
    { d: "Snapdragon 662（入門手機）", r: "不成立", c: C.clay, f: C.rose, t: "本來就只有 15–16 FPS，連續跑也不變" },
    { d: "天璣 8300 平板", r: "成立，只到 320", c: C.mid, f: C.tint, t: "416 短時間 37 FPS，連續跑剩 24；512 剩 16" },
    { d: "Snapdragon 8 Gen 2 手機", r: "成立，只到 320", c: C.mid, f: C.tint, t: "416 短時間 28 FPS 過線，連續跑剩 19" },
  ];
  v.forEach((it, i) => {
    const y = 1.8 + i * 1.42;
    card(s, rx, y, rw, 1.27, it.f);
    s.addText(it.d, { x: rx + 0.3, y: y + 0.12, w: rw - 0.6, h: 0.3, isTextBox: true, margin: 0, fontFace: F, fontSize: 12, bold: true, color: C.muted });
    s.addText(it.r, { x: rx + 0.3, y: y + 0.42, w: rw - 0.6, h: 0.42, isTextBox: true, margin: 0, fontFace: F, fontSize: 20, bold: true, color: it.c });
    s.addText(it.t, { x: rx + 0.3, y: y + 0.86, w: rw - 0.6, h: 0.3, isTextBox: true, margin: 0, fontFace: F, fontSize: 12, color: C.ink });
  });

  s.addText([
    { text: "兩台高階機都只到 320。", options: { bold: true, color: C.dark } },
    { text: "662 與這兩台之間的中階裝置沒有量——界線在哪裡，現在不知道。", options: { color: C.ink } },
  ], { x: M, y: 6.25, w: CW, h: 0.45, isTextBox: true, margin: 0, fontFace: F, fontSize: 14.5 });
  pageNote(s, "FPS = 1000 ÷ 每張毫秒數，是模型本身的上限，不含相機取像與畫框　·　逐解析度、GPU、NNAPI 見備答");
  s.addNotes("【約 1 分鐘】看右邊三張卡就好：入門機不行、兩台高階機只到 320。圖的重點是灰色（短時間）和綠色（連續跑）的落差——平板落差最大。下一頁講為什麼。");
}

// =====================================================================
// 7. 為什麼要量「連續跑」
// =====================================================================
{
  const s = pres.addSlide(); lightBg(s);
  title(s, "多講一分鐘：為什麼要量「連續跑」", "連續 10 段、每段約 30 秒的平均延遲（ms）——越高越慢，40 ms 以上就不及格　·　平板 416 含 2 ms 後處理保留");

  const segs = ["1", "2", "3", "4", "5", "6", "7", "8", "9", "10"];
  s.addChart(pres.ChartType.line, [
    { name: "平板 416", labels: segs, values: [31.64, 34.01, 35.83, 36.93, 37.72, 38.97, 39.22, 39.78, 42.63, 41.63] },
    { name: "手機 416", labels: segs, values: [38.64, 35.81, 48.67, 49.35, 48.65, 49.09, 48.78, 51.68, 48.67, 48.74] },
    { name: "662 320", labels: segs, values: [62.31, 62.21, 63.32, 62.90, 62.47, 64.54, 63.22, 63.09, 63.18, 63.08] },
    { name: "門檻 40 ms", labels: segs, values: [40, 40, 40, 40, 40, 40, 40, 40, 40, 40] },
  ], quietChart({
    x: M, y: 1.8, w: 7.35, h: 4.2,
    chartColors: [C.accent, C.mid, "8FA79A", C.clay],
    lineSize: 2.5, lineDataSymbol: "circle", lineDataSymbolSize: 6,
    valAxisMinVal: 20, valAxisMaxVal: 70, valAxisMajorUnit: 10,
    catAxisTitle: "段", showCatAxisTitle: true, catAxisTitleColor: C.muted, catAxisTitleFontSize: 11, catAxisTitleFontFace: F,
  }));

  const rx = M + 7.7, rw = CW - 7.7;
  const v = [
    { d: "平板 · 416 解析度", n: "37 → 24 FPS", c: C.accent, t: "頭 30 秒表面 34 → 47 °C，第 9 段起超線" },
    { d: "手機 · 416 解析度", n: "28 → 19 FPS", c: C.mid, t: "沒有過熱，第 3 段起整段掉到另一個檔位" },
    { d: "662 · 320 解析度", n: "15 → 16 FPS", c: C.muted, t: "連續 10 段完全沒變慢，只是本來就太慢" },
  ];
  v.forEach((it, i) => {
    const y = 1.8 + i * 1.42;
    card(s, rx, y, rw, 1.27, C.tint2);
    s.addText(it.d, { x: rx + 0.28, y: y + 0.12, w: rw - 0.56, h: 0.28, isTextBox: true, margin: 0, fontFace: F, fontSize: 12, bold: true, color: C.muted });
    s.addText(it.n, { x: rx + 0.28, y: y + 0.4, w: rw - 0.56, h: 0.44, isTextBox: true, margin: 0, fontFace: F, fontSize: 22, bold: true, color: it.c });
    s.addText(it.t, { x: rx + 0.28, y: y + 0.86, w: rw - 0.56, h: 0.3, isTextBox: true, margin: 0, fontFace: F, fontSize: 11.5, color: C.ink });
  });

  card(s, M, 6.18, CW, 0.78, C.peach);
  s.addText([
    { text: "跑得越快的晶片，連續跑掉得越多：", options: { bold: true, color: C.accent } },
    { text: "平板 1.6–1.7 倍、手機 1.05–1.4 倍、662 不掉。只看短時間的數字，會以為高階機的下一個解析度也可以做即時辨識。", options: { color: C.ink } },
  ], { x: M + 0.3, y: 6.24, w: CW - 0.6, h: 0.66, isTextBox: true, margin: 0, fontFace: F, fontSize: 13.5, valign: "middle" });
  s.addNotes("【約 1 分 15 秒】指著圖：橘線（平板 416）一路往上爬、第 9 段越過紅線；綠線（手機 416）第 3 段突然跳上去；灰線（662）一直平、而且本來就在紅線上面。平板 416 是 09-15 補量的（原本因精度表寫錯而漏量，見備答）；平板 512 更明顯，連續跑剩 16 FPS。換成畫面感受：相機每秒 30 張，16 FPS 大約每兩張只辨識一張。「掉得越多」的原因（算力高 → 發熱大 → 先降頻）是推測，沒量功耗。");
}

// =====================================================================
// 8. 決策一
// =====================================================================
{
  const s = pres.addSlide(); darkBg(s);
  kicker(s, "要在這場會議上決定的事");
  badge(s, "1", M, 1.1, 0.72, C.accent, C.dark);
  s.addText("App 要不要依裝置等級開放即時辨識？", {
    x: M + 0.95, y: 1.14, w: CW - 0.95, h: 0.52, isTextBox: true, margin: 0,
    fontFace: F, fontSize: 32, bold: true, color: C.white,
  });

  const cols = [
    { h: "做", t: "高階機開即時辨識（320），其他只做拍照辨識。", need: "還需要：決定怎麼界定「高階」——662 與兩台高階機之間的中階裝置沒有量，界線在哪裡不知道。" },
    { h: "不做", t: "全部只做拍照辨識。", need: "不需要再量：三台結果寫進報告，當作實測上限。" },
  ];
  const gap = 0.4, cw = (CW - gap) / 2;
  cols.forEach((c, i) => {
    const x = M + i * (cw + gap);
    card(s, x, 2.3, cw, 2.75, C.dark2);
    s.addText(c.h, { x: x + 0.35, y: 2.5, w: cw - 0.7, h: 0.55, isTextBox: true, margin: 0, fontFace: F, fontSize: 26, bold: true, color: C.accent });
    s.addText(c.t, { x: x + 0.35, y: 3.12, w: cw - 0.7, h: 0.7, isTextBox: true, margin: 0, fontFace: F, fontSize: 16, bold: true, color: C.white, valign: "top" });
    s.addText(c.need, { x: x + 0.35, y: 3.9, w: cw - 0.7, h: 1.0, isTextBox: true, margin: 0, fontFace: F, fontSize: 14, color: C.sage, valign: "top" });
  });

  s.addText("這是產品決策——技術面三台都量清楚了。", {
    x: M, y: 5.5, w: CW, h: 0.55, isTextBox: true, margin: 0,
    fontFace: F, fontSize: 20, bold: true, color: C.accent,
  });
  s.addNotes("【約 45 秒】把球丟回去。若選「做」，下週的工作是定分級方式；若選「不做」，部署線就此收尾。");
}

// =====================================================================
// 9. 決策二與三
// =====================================================================
{
  const s = pres.addSlide(); lightBg(s);
  kicker(s, "要在這場會議上決定的事");
  s.addText("決策 2 與決策 3", {
    x: M, y: 0.9, w: CW, h: 0.65, isTextBox: true, margin: 0,
    fontFace: F, fontSize: 32, bold: true, color: C.dark,
  });

  const gap = 0.4, cw = (CW - gap) / 2, y0 = 1.85, ch = 4.3;
  // 決策 2
  card(s, M, y0, cw, ch, C.tint);
  badge(s, "2", M + 0.35, y0 + 0.32, 0.6, C.accent, C.dark);
  s.addText("拍照模式可以接受多久出結果？", { x: M + 1.1, y: y0 + 0.36, w: cw - 1.4, h: 0.5, isTextBox: true, margin: 0, fontFace: F, fontSize: 19, bold: true, color: C.dark });
  s.addText("有了這個上限，才能在三個組合之間選：", { x: M + 0.35, y: y0 + 1.2, w: cw - 0.7, h: 0.4, isTextBox: true, margin: 0, fontFace: F, fontSize: 14, color: C.ink });
  ["fp32@320", "fp32@640", "w8a32@640"].forEach((t, i) => {
    const bx = M + 0.35 + i * 1.8;
    card(s, bx, y0 + 1.72, 1.62, 0.56, C.white);
    s.addText(t, { x: bx, y: y0 + 1.72, w: 1.62, h: 0.56, isTextBox: true, margin: 0, fontFace: F, fontSize: 14, bold: true, color: C.mid, align: "center", valign: "middle" });
  });
  s.addText("選定之後只需要用交付權重（v5.7 資料集訓練）重新匯出那一個組合、重評精度。", { x: M + 0.35, y: y0 + 2.6, w: cw - 0.7, h: 0.8, isTextBox: true, margin: 0, fontFace: F, fontSize: 14, color: C.ink, valign: "top" });
  s.addText("需要：團隊給「按下快門後多久要出結果」的上限", { x: M + 0.35, y: y0 + 3.55, w: cw - 0.7, h: 0.5, isTextBox: true, margin: 0, fontFace: F, fontSize: 14, bold: true, color: C.accent });

  // 決策 3
  const x3 = M + cw + gap;
  card(s, x3, y0, cw, ch, C.tint2);
  badge(s, "3", x3 + 0.35, y0 + 0.32, 0.6, C.accent, C.dark);
  s.addText("剩下的時間要不要繼續追模型精度？", { x: x3 + 1.1, y: y0 + 0.36, w: cw - 1.4, h: 0.5, isTextBox: true, margin: 0, fontFace: F, fontSize: 19, bold: true, color: C.dark });
  s.addText([
    { text: "工程手段已經用盡：改架構、放大模型、調匯出參數、加外部照片，全部量過都沒用", options: { bullet: true, breakLine: true } },
    { text: "剩下有依據的只有兩件：量介殼蟲的框一致性（兩人各半小時），或補拍同域照片（約 353 張）", options: { bullet: true, breakLine: true } },
    { text: "不追的話，剩下的時間用來收尾文件與報告", options: { bullet: true } },
  ], {
    x: x3 + 0.35, y: y0 + 1.15, w: cw - 0.7, h: 2.9, isTextBox: true, margin: 0,
    fontFace: F, fontSize: 14, color: C.ink, paraSpaceAfter: 10, valign: "top",
  });
  s.addNotes("【約 30 秒】兩個都是一句話的問題。決策 2 只需要一個秒數；決策 3 若選「追」，最便宜的是介殼蟲一致性，兩人各半小時。");
}

// =====================================================================
// 10. 下週
// =====================================================================
{
  const s = pres.addSlide(); lightBg(s);
  title(s, "下週", "依決策結果");

  const items = [
    { h: "決策 1 選「做」", t: "決定裝置分級的方式；需要的話量一支中階機，用同一份判準" },
    { h: "決策 2 有上限", t: "選定拍照模式的組合，以交付權重重新匯出並重評精度" },
    { h: "決策 3", t: "選「追」→ 發派介殼蟲的一致性測試\n選「不追」→ 收尾文件" },
  ];
  const gap = 0.35, cw = (CW - 2 * gap) / 3;
  items.forEach((it, i) => {
    const x = M + i * (cw + gap);
    card(s, x, 1.85, cw, 1.95, C.tint);
    badge(s, i + 1, x + 0.3, 2.07, 0.46, C.mid);
    s.addText(it.h, { x: x + 0.9, y: 2.08, w: cw - 1.15, h: 0.44, isTextBox: true, margin: 0, fontFace: F, fontSize: 16, bold: true, color: C.mid, valign: "middle" });
    s.addText(it.t, { x: x + 0.3, y: 2.7, w: cw - 0.6, h: 0.95, isTextBox: true, margin: 0, fontFace: F, fontSize: 13.5, color: C.ink, valign: "top" });
  });

  card(s, M, 4.0, CW, 1.0, C.tint2);
  s.addText([
    { text: "已完成：", options: { bold: true, color: C.mid } },
    { text: "天璣 8300、Snapdragon 8 Gen 2、三台整合三篇實機測試報告已推上報告庫；報告庫的 09-13 週報已補入高階裝置，並新增給教授與評審看的成果總覽 SUMMARY.md。", options: { color: C.ink } },
  ], { x: M + 0.35, y: 4.1, w: CW - 0.7, h: 0.8, isTextBox: true, margin: 0, fontFace: F, fontSize: 14, valign: "middle" });

  card(s, M, 5.15, CW, 1.75, C.rose);
  s.addText("已經有結論、不要再花時間的", { x: M + 0.35, y: 5.28, w: CW - 0.7, h: 0.4, isTextBox: true, margin: 0, fontFace: F, fontSize: 16, bold: true, color: C.clay });
  s.addText("模型架構與損失改動　·　放大模型　·　官方匯出參數　·　662 上的即時辨識　·　NNAPI\n影像合成與生成　·　自動標註（色彩門檻找框、以現有模型預標）　·　外部影像（除非拿得到同域照片）\n刪除薊馬葉害　·　只憑短時間量測下即時辨識的結論", {
    x: M + 0.35, y: 5.72, w: CW - 0.7, h: 1.05, isTextBox: true, margin: 0, fontFace: F, fontSize: 13, color: C.ink, valign: "top",
  });
  s.addNotes("【約 15 秒】這頁講完正文結束，進備答。總長約 8 分鐘。");
}

// =====================================================================
// 11. 備答分隔頁
// =====================================================================
{
  const s = pres.addSlide(); darkBg(s);
  s.addText("備答", {
    x: M, y: 2.6, w: CW, h: 1.0, isTextBox: true, margin: 0,
    fontFace: F, fontSize: 48, bold: true, color: C.white,
  });
  s.addText("以下幾頁是被問到才翻的資料　·　交付模型與資料集、三台數據總表、GPU 與 NNAPI、判準時間線、常見提問", {
    x: M, y: 3.68, w: CW, h: 0.5, isTextBox: true, margin: 0,
    fontFace: F, fontSize: 17, color: C.accent,
  });
  s.addNotes("正文到這裡結束。");
}

// =====================================================================
// 11a. 備答：交付模型 v11.5（v5.7 資料集）的規格與整體指標
// =====================================================================
{
  const s = pres.addSlide(); lightBg(s);
  title(s, "交付模型 v11.5（v5.7 資料集）：規格與指標", "v5.7 是資料集版本，訓練碼與架構都是 v11.5（run v5.7_v11_5）　·　valid、test 各 401 張");

  const lw = 6.3, gap = 0.3, rw = CW - lw - gap;
  const K = { bold: true, color: C.dark };
  s.addTable([
    [hdr("項目"), hdr("內容")],
    [cell("權重", K), cell("last.pt（5.59 MB）")],
    [cell("架構", K), cell("官方 yolo26-p2（scale n、end2end、NMS-free），未自訂修改")],
    [cell("類別", K), cell("9 類：4 種病害、4 種害蟲、薊馬葉害")],
    [cell("資料集", K), cell("v5.7：train 7,264 / valid 401 / test 401 張")],
    [cell("訓練", K), cell("固定 70 輪、patience 0、MuSGD、imgsz 640、batch 20")],
    [cell("耗時", K), cell("Kaggle GPU 約 3.1 小時（每輪約 161 秒，seed 0）")],
    [cell("環境", K), cell("ultralytics 8.4.121（版本釘死）")],
  ], { x: M, y: 1.8, w: lw, colW: [1.0, lw - 1.0], rowH: 0.5, ...tblBase, fontSize: 12, align: "left" });

  const B = { bold: true, color: C.mid };
  const na = { color: C.muted };
  s.addTable([
    [hdr("指標"), hdr("valid"), hdr("test")],
    [cell("mAP50", { bold: true, align: "left" }), cell("0.826", {}), cell("0.861", B)],
    [cell("mAP50-95", { bold: true, align: "left" }), cell("0.626", {}), cell("0.641", B)],
    [cell("Precision", { bold: true, align: "left" }), cell("0.852", {}), cell("0.909", B)],
    [cell("Recall", { bold: true, align: "left" }), cell("0.821", {}), cell("0.812", B)],
    [cell("Detection Jaccard", { bold: true, align: "left" }), cell("—", na), cell("0.727", B)],
    [cell("TP / FP / FN", { bold: true, align: "left" }), cell("—", na), cell("675 / 118 / 135", B)],
    [cell("八類平均 AP50（標註未變動）", { bold: true, align: "left" }), cell("—", na), cell("0.851", B)],
  ], { x: M + lw + gap, y: 1.8, w: rw, colW: [2.55, 1.2, rw - 3.75], rowH: 0.5, ...tblBase, fontSize: 12, align: "center" });

  card(s, M, 6.0, CW, 0.78, C.rose);
  s.addText([
    { text: "口徑：", options: { bold: true, color: C.clay } },
    { text: "九類 test mAP50 0.861 不可與上一版（v5.6 資料集）的 0.810 相比（薊馬葉害換了框定義）；標註沒變的八類平均 0.853 → 0.851，等於沒變。手機上的 TFLite 仍是 v11.5（v5.6 資料集）權重匯出。", options: { color: C.ink } },
  ], { x: M + 0.3, y: 6.06, w: CW - 0.6, h: 0.66, isTextBox: true, margin: 0, fontFace: F, fontSize: 12.5, valign: "middle" });
  pageNote(s, "Precision／Recall 取自 ultralytics val()（conf 0.001）；Jaccard 與 TP／FP／FN 取自混淆矩陣（conf 0.25），兩者不在同一工作點　·　來源 docs/v5.7_報告_交付與研究限制.md §1");
  s.addNotes("被問「模型現在幾分」時翻這頁。先講口徑那一行，再報數字：test mAP50 0.861，但不能跟上週的 0.810 比。");
}

// =====================================================================
// 11b. 備答：交付模型 v11.5（v5.7 資料集）的九類逐類 AP50
// =====================================================================
{
  const s = pres.addSlide(); lightBg(s);
  title(s, "交付模型 v11.5（v5.7 資料集）：九類逐類 AP50", "valid＋test 併計（pooled，802 張）；九類 ±2SE 全部 ≤ 0.10　·　依 pooled 由高到低");
  const weak = new Set(["Citrus_Leaf_Miner（潛葉蛾）", "Thrips（薊馬）", "Scale_Insect（介殼蟲）"]);
  const rows = [
    ["Sooty_Mold（煤煙病）", "0.9950", "0.9950", "0.9950", "0.0210", "33 / 32", ""],
    ["Black_Spot（黑點病）", "0.9603", "0.9454", "0.9528", "0.0304", "25 / 25", ""],
    ["Oily_Spot（油斑病）", "0.9430", "0.9533", "0.9482", "0.0156", "24 / 24", ""],
    ["Aphid（蚜蟲）", "0.8974", "0.9615", "0.9297", "0.0341", "75 / 76", ""],
    ["Canker（潰瘍病）", "0.9282", "0.8809", "0.9045", "0.0431", "24 / 24", "框最多的一類（train 6,489 框）"],
    ["Thrips_Damage（薊馬葉害）", "0.7243", "0.8453", "0.7848", "0.0825", "40 / 40", "新約定（一葉一框），不可跨版比"],
    ["Citrus_Leaf_Miner（潛葉蛾）", "0.7931", "0.6875", "0.7403", "0.0990", "45 / 45", "±2SE 最寬"],
    ["Thrips（薊馬）", "0.6453", "0.7626", "0.7040", "0.0969", "60 / 60", "蟲體來源含圖鑑、微距特寫"],
    ["Scale_Insect（介殼蟲）", "0.5452", "0.7190", "0.6321", "0.0490", "35 / 35", "最弱；框一致性從未量過"],
  ];
  s.addTable([
    [hdr("類別"), hdr("valid"), hdr("test"), hdr("pooled"), hdr("±2SE"), hdr("評估圖 v / t"), hdr("備註")],
    ...rows.map((r) => [
      cell(r[0], { bold: true, color: C.dark, align: "left" }),
      cell(r[1]), cell(r[2]),
      cell(r[3], { bold: true, color: weak.has(r[0]) ? C.clay : C.mid }),
      cell(r[4]), cell(r[5]),
      cell(r[6], { color: C.muted, align: "left", fontSize: 11 }),
    ]),
  ], { x: M, y: 1.75, w: CW, colW: [3.0, 1.0, 1.0, 1.05, 0.95, 1.3, 3.633], rowH: 0.44, ...tblBase, fontSize: 12, align: "center" });

  s.addText("數字取自 last.pt（valid 未參與任何決策，才能與 test 併計）　·　逐類雜訊地板 ±0.04、run 間全距 0.021，小於這個幅度不解讀", {
    x: M, y: 6.3, w: CW, h: 0.6, isTextBox: true, margin: 0, fontFace: F, fontSize: 12.5, color: C.ink, valign: "top",
  });
  s.addNotes("被問「哪一類最差」「某一類幾分」時翻這頁。紅字是 pooled 最低的三類；介殼蟲最弱，而且它的框一致性從未量過，是精度線唯一還沒走過的便宜路。");
}

// =====================================================================
// 11c. 備答：資料集 v5.7 的組成
// =====================================================================
{
  const s = pres.addSlide(); lightBg(s);
  title(s, "資料集 v5.7：組成", "與 v5.6 只差薊馬葉害 208 張重標（一張葉子一個框，245 → 215 框）；其餘八類逐位元相同");
  const T = { bold: true, color: C.dark };
  const fmt = (n) => n.toLocaleString("en-US");
  const data = [
    ["Oily_Spot（油斑病）", 238, 190, 570, 24, 24, 759, 24, 24],
    ["Canker（潰瘍病）", 241, 193, 579, 24, 24, 6489, 167, 182],
    ["Sooty_Mold（煤煙病）", 326, 261, 783, 33, 32, 1036, 33, 32],
    ["Black_Spot（黑點病）", 250, 200, 600, 25, 25, 801, 25, 25],
    ["Scale_Insect（介殼蟲）", 251, 181, 543, 35, 35, 3427, 187, 135],
    ["Citrus_Leaf_Miner（潛葉蛾）", 201, 111, 333, 45, 45, 652, 54, 64],
    ["Thrips（薊馬）", 554, 434, 493, 60, 60, 1317, 115, 97],
    ["Aphid（蚜蟲）", 806, 655, 545, 75, 76, 3284, 216, 212],
    ["Thrips_Damage（薊馬葉害）", 208, 128, 145, 40, 40, 281, 42, 41],
    ["Background（負樣本）", 400, 320, 0, 40, 40, 0, 0, 0],
  ];
  s.addTable([
    [hdr("類別"), hdr("原始圖數"), hdr("train 原始"), hdr("train 增強"), hdr("valid 圖"), hdr("test 圖"), hdr("train 框"), hdr("valid 框"), hdr("test 框")],
    ...data.map((r) => [
      cell(r[0], { bold: true, color: C.dark, align: "left" }),
      ...r.slice(1).map((v, j) => cell(fmt(v), r[0].startsWith("Thrips_Damage") && j >= 5 ? { bold: true, color: C.accent } : {})),
    ]),
    [cell("合計", T), ...[3475, 2673, 4591, 401, 401, 18046, 863, 812].map((v) => cell(fmt(v), T))],
  ], { x: M, y: 1.72, w: CW, colW: [2.65, 1.15, 1.15, 1.15, 0.95, 0.95, 1.3, 1.25, 1.383], rowH: 0.37, ...tblBase, fontSize: 11.5, align: "center" });

  s.addText("原始圖數為不重複的原始影像；train 圖 = 原始 ＋ 增強（增強只做在 train，各類佔 45–75%）　·　跨 split 近重複 0 組\n九類 pooled ±2SE ≤ 0.10（最差潛葉蛾 0.099）　·　Canker 佔 train 框 36%（類別不平衡沿襲 v5.5）　·　橘字為本版重標後的薊馬葉害框數", {
    x: M, y: 6.33, w: CW, h: 0.6, isTextBox: true, margin: 0, fontFace: F, fontSize: 12, color: C.ink, valign: "top",
  });
  s.addNotes("被問「資料集多大」「每類幾張」時翻這頁。重點：評估集每類只有 24–76 張，所以逐類分數有 ±2SE；潛葉蛾原始影像最少（201 張）。");
}
// =====================================================================
// 12. 備答：三台數據總表
// =====================================================================
{
  const s = pres.addSlide(); lightBg(s);
  title(s, "三台 Benchmark 總表（FPS）", "同 8 個 fp32 檔案、同一套量測方法　·　延遲與 mAP50 都來自 v11.5（v5.6 資料集）權重的匯出檔，交付權重尚未重新匯出");

  const nm = { color: C.muted, italic: true };
  const B = { bold: true };
  const G = { bold: true, color: C.mid };
  s.addTable([
    [hdr("組合（未註明者 end2end 開）"), hdr("662 短時間"), hdr("662 連續跑"), hdr("平板 短時間"), hdr("平板 連續跑"), hdr("手機 短時間"), hdr("手機 連續跑"), hdr("mAP50")],
    [cell("fp32@320", B), cell("14.9"), cell("15.8"), cell("66.6", G), cell("38.4", G), cell("32.0", G), cell("29.8", G), cell("0.780")],
    [cell("fp32@416", B), cell("9.1"), cell("沒量", nm), cell("39.0"), cell("沒量", nm), cell("28.2"), cell("19.3", { color: C.clay, bold: true }), cell("0.800")],
    [cell("fp32@416 · end2end 關", B), cell("8.8"), cell("沒量", nm), cell("37.1"), cell("23.5", { color: C.clay, bold: true }), cell("20.1"), cell("沒量", nm), cell("0.815")],
    [cell("fp32@512", B), cell("5.8"), cell("沒量", nm), cell("25.4"), cell("15.6", { color: C.clay, bold: true }), cell("18.9"), cell("沒量", nm), cell("0.814")],
    [cell("fp32@640", B), cell("3.7"), cell("沒量", nm), cell("15.9"), cell("沒量", nm), cell("11.8"), cell("沒量", nm), cell("0.816")],
    [cell("fp32@640 · end2end 關 · GPU", B), cell("3.5"), cell("沒量", nm), cell("29.6"), cell("沒量", nm), cell("30.7"), cell("沒量", nm), cell("0.826")],
    [cell("判定", B), cell("不成立", { bold: true, color: C.clay, colspan: 2 }), cell("成立，只到 320", { bold: true, color: C.mid, colspan: 2 }), cell("成立，只到 320", { bold: true, color: C.mid, colspan: 2 }), cell("")],
  ], {
    x: M, y: 1.85, w: CW, colW: [3.0, 1.28, 1.28, 1.28, 1.28, 1.28, 1.28, 1.253], rowH: 0.43,
    ...tblBase, fontSize: 12.5, align: "center",
  });

  const gap = 0.35, cw = (CW - gap) / 2;
  card(s, M, 5.45, cw, 1.1, C.tint);
  s.addText([
    { text: "平板比 662 快 4.3–4.5 倍，", options: { bold: true, color: C.dark } },
    { text: "各解析度一致，差距是純算力。", options: { color: C.ink } },
  ], { x: M + 0.3, y: 5.55, w: cw - 0.6, h: 0.9, isTextBox: true, margin: 0, fontFace: F, fontSize: 13.5, valign: "middle" });
  card(s, M + cw + gap, 5.45, cw, 1.1, C.peach);
  s.addText([
    { text: "手機只快 2.1–3.2 倍，320 最吃虧；", options: { bold: true, color: C.accent } },
    { text: "閒置時 CPU 頻率上限就只有 50–59%，原因未查明。", options: { color: C.ink } },
  ], { x: M + cw + gap + 0.3, y: 5.55, w: cw - 0.6, h: 0.9, isTextBox: true, margin: 0, fontFace: F, fontSize: 13.5, valign: "middle" });

  pageNote(s, "FPS = 1000 ÷ 有效延遲（end2end 關另加 2 ms）；短時間取 3 輪最快、連續跑取最後 3 段最慢　·　「沒量」是依計畫只量指定組合，不是失敗");
  s.addNotes("被問到「其他解析度呢」「手機為什麼慢」時翻這頁。紅字是短時間過線、連續跑不過的三個組合（平板 416 end2end 關是 09-15 補量）。mAP50 依匯出報告逐變體列出，end2end 關的精度比開著高約 0.01。");
}

// =====================================================================
// 13. 備答：GPU 與 NNAPI
// =====================================================================
{
  const s = pres.addSlide(); lightBg(s);
  title(s, "GPU 與 NNAPI：只有高階機的 GPU 有用", "fp32@640、end2end 關：GPU 延遲 ÷ CPU 延遲（小於 1 才是變快）　·　各只量 1 輪，不構成判定");

  s.addChart(pres.ChartType.bar, [
    { name: "GPU ÷ CPU", labels: ["Snapdragon 662", "天璣 8300 平板", "8 Gen 2 手機"], values: [1.05, 0.52, 0.37] },
  ], quietChart({
    x: M, y: 1.8, w: 5.9, h: 4.4,
    barDir: "col", barGapWidthPct: 70,
    chartColors: [C.mid],
    showLegend: false,
    showValue: true, dataLabelPosition: "outEnd", dataLabelColor: C.ink,
    dataLabelFontSize: 13, dataLabelFontFace: F, dataLabelFormatCode: "0.00",
    valAxisMinVal: 0, valAxisMaxVal: 1.2, valAxisMajorUnit: 0.2,
  }));

  const rx = M + 6.25, rw = CW - 6.25;
  const rows = [
    { h: "end2end 決定 GPU 能不能整張接手", t: "關掉時 511/511 個節點全交給 GPU；開著只有 54/559，偵測頭的算子不被支援。三台一致。", f: C.tint2 },
    { h: "640 在兩台高階機約 31 ms", t: "平板 31.8、手機 30.6 ms，是唯一「精度最高又在 40 ms 內」的組合；662 整張交給 GPU 仍是 282 ms，沒變快。", f: C.tint },
    { h: "但還不能用", t: "只量 1 輪、沒有連續跑、GPU 管線的後處理成本沒量——要另外登記一份計畫再量。", f: C.peach },
    { h: "NNAPI 三台都不用", t: "662 與手機只有 Android 內建的參考實作、沒生效；平板交給 MediaTek 加速器反而慢 1.7–5.4 倍。", f: C.rose },
  ];
  rows.forEach((r, i) => {
    const y = 1.8 + i * 1.13;
    card(s, rx, y, rw, 1.02, r.f);
    s.addText(r.h, { x: rx + 0.3, y: y + 0.1, w: rw - 0.6, h: 0.34, isTextBox: true, margin: 0, fontFace: F, fontSize: 14, bold: true, color: C.dark });
    s.addText(r.t, { x: rx + 0.3, y: y + 0.44, w: rw - 0.6, h: 0.52, isTextBox: true, margin: 0, fontFace: F, fontSize: 11.5, color: C.ink, valign: "top" });
  });
  s.addNotes("被問「GPU 不是比較快嗎」時翻這頁。一句話：高階機上有機會，但要另開一份計畫連同連續跑一起量；662 上 GPU 沒幫助。");
}

// =====================================================================
// 14. 備答：判準比量測早（時間線）
// =====================================================================
{
  const s = pres.addSlide(); lightBg(s);
  title(s, "判準都比量測早", "每一條都可以從 git 紀錄或登記文件查到時間");

  const rows = [
    ["2026-08", "標註一致性門檻：兩人中位 IoU ≥ 0.85", "v11 計畫書", "09-12 才量工作包 A"],
    ["09-11 上午", "v12.2 即時辨識採用門檻登記", "66747bd", "同日下午才量 662"],
    ["09-12 23:44", "薊馬葉害「一張葉子一個框」重測規則", "2914698", "23:51、23:53 才有重測結果"],
    ["09-13 凌晨", "外部影像增益 0.04、矛盾門檻 0.165（驗收文件）", "fc729a9", "之後才訓練兩臂"],
    ["09-13 18:40", "v12.3 高階裝置門檻，新增「連續跑也要 ≤ 40 ms」", "e15e916", "18:43 才開始量平板"],
    ["09-13 晚", "320 取 e2e1、兩組之間冷卻看表面溫度", "384db65、3e620fd", "各自在對應那組開跑之前"],
    ["09-13 22:26 前", "溫度讀取改為不分廠牌", "7f6df28", "之後才開始量手機"],
    ["09-13 23:49", "662 補量協定登記", "b465196", "同一分鐘才開始補量"],
    ["09-15 00:00", "平板補量 416 end2end 關（精度按變體更正後最準的組合）", "f2f76d3", "00:01 才開始量"],
  ];
  rows.forEach((r, i) => {
    const y = 1.72 + i * 0.54;
    card(s, M, y, CW, 0.47, i % 2 === 0 ? C.tint2 : C.tint);
    badge(s, i + 1, M + 0.2, y + 0.05, 0.37, C.mid);
    s.addText(r[0], { x: M + 0.85, y: y, w: 1.8, h: 0.47, isTextBox: true, margin: 0, fontFace: F, fontSize: 13, bold: true, color: C.dark, valign: "middle" });
    s.addText(r[1], { x: M + 2.7, y: y, w: 4.9, h: 0.47, isTextBox: true, margin: 0, fontFace: F, fontSize: 13, color: C.ink, valign: "middle" });
    s.addText(r[2], { x: M + 7.65, y: y, w: 1.8, h: 0.47, isTextBox: true, margin: 0, fontFace: /^[0-9a-f]{7}/.test(r[2]) ? "Consolas" : F, fontSize: 12, color: C.muted, valign: "middle" });
    s.addText(r[3], { x: M + 9.5, y: y, w: CW - 9.7, h: 0.47, isTextBox: true, margin: 0, fontFace: F, fontSize: 12.5, bold: true, color: C.mid, valign: "middle" });
  });
  s.addText("量完之後，判準一條都沒有改。途中補定的都是登記文件沒寫到的細節，而且都寫在結果文件裡。", {
    x: M, y: 6.6, w: CW, h: 0.4, isTextBox: true, margin: 0, fontFace: F, fontSize: 13.5, bold: true, color: C.accent,
  });
  s.addNotes("被問「是不是看完結果才訂標準」時翻這頁，挑一兩列念 commit 時間就夠。");
}

// =====================================================================
// 15–16. 備答：常見提問（兩頁）
// =====================================================================
const QA = [
  ["0.861 比 0.810 高，是不是進步了？", "不是。框定義換了，框變大就容易算對；標註沒變的八類平均是 −0.002，等於沒變"],
  ["0.899 是模型的準確度嗎？", "不是。那是兩個人各自畫框的重疊度（IoU 中位數），用來判斷標註規則畫不畫得一致，和模型無關"],
  ["那為什麼要換成 v5.7 資料集？", "模型還是 v11.5，只換資料集。舊的薊馬葉害框彼此矛盾，那個類別的數字量到的是標註噪聲；換了之後數字才有解釋力"],
  ["為什麼不直接刪掉薊馬葉害？", "它是唯一符合田間場景的薊馬訊號，刪掉數字也不會變好；改規則重測一次就通過了"],
  ["外部照片為什麼沒用？", "那批是摘下的葉子放白紙上拍，跟田間照差很多，而且 124 張只來自約 20–25 片葉子；加資料的潛葉蛾和沒加資料的油斑病動得一樣多"],
  ["兩台高階機都過了，是不是可以做即時辨識？", "只到 320，而且各只有一台；中階裝置沒有量，界線在哪裡不知道"],
  ["手機 30 FPS 比平板慢，為什麼也算過？", "標準是 40 ms（25 FPS）；手機連續跑最慢的一段是 33.5 ms，還有 6.5 ms 餘裕"],
  ["8 Gen 2 比天璣 8300 強，手機為何較慢？", "這支手機閒置時 CPU 頻率上限就被壓在一半左右，來源沒查到；照出廠狀態量、不改系統設定，結論只代表這支手機"],
  ["短時間已經過線，為什麼算不行？", "即時辨識是一直在跑的；平板 30 秒內就降頻，手機跑到第 3 段就掉檔，連續跑都超過 40 ms"],
  ["662 連續跑會不會更慢？", "不會。補量時連續跑 10 段都在 63 ms 左右、完全沒變慢；只是本來就只有 16 FPS，離門檻太遠"],
  ["662 的舊數字只量一輪，可信嗎？", "用跟另外兩台一樣的方法（3 輪交錯）重量過了，8 個檔案差距都在 ±3.1% 以內，結論不變"],
  ["GPU 跑 640 只要約 31 ms，為什麼不用？", "只在兩台高階機上成立，662 整張交給 GPU 也沒變快；而且只量一輪、沒連續跑、沒量後處理，要另外登記計畫再量"],
  ["平板 416 上次沒量，現在呢？", "09-15 補量：它是短時間過線中最準的（0.815），當時精度表寫錯才漏掉；照原規則補量，第 9 段起超過 40 ms，平板仍只到 320"],
  ["還有辦法讓模型更準嗎？", "不拍照片的前提下沒有了；剩下量介殼蟲的標註一致性，或補拍照片"],
  ["這些判準、數字可信嗎？", "判準都比量測早 commit（見前一頁）；09-15 凌晨也用原始輸出把數字全核對過一次，更正處見週報 §5.1"],
];
[QA.slice(0, 7), QA.slice(7)].forEach((part, pi) => {
  const s = pres.addSlide(); lightBg(s);
  title(s, "備答：可能會被問到的（" + (pi + 1) + "／2）", pi === 0 ? "標註、模型與部署結論" : "裝置差異、662 補量、GPU 與判準");
  s.addTable([
    [hdr("問題"), hdr("一句話回答")],
    ...part.map((q) => [cell(q[0], { bold: true, color: C.dark }), cell(q[1])]),
  ], {
    x: M, y: pi === 0 ? 1.8 : 1.7, w: CW, colW: [4.1, 7.833], rowH: pi === 0 ? 0.64 : 0.58,
    ...tblBase, fontSize: 12, align: "left",
  });
  s.addNotes("這頁不用講，被問到再看。");
});

// =====================================================================
// 17. 結尾
// =====================================================================
{
  const s = pres.addSlide(); darkBg(s);
  s.addText("需要當場決定的三件事", {
    x: M, y: 1.2, w: CW, h: 0.7, isTextBox: true, margin: 0,
    fontFace: F, fontSize: 34, bold: true, color: C.white,
  });

  const d = [
    ["1", "App 要不要依裝置等級開放即時辨識？", "做 → 高階機開 320，要先定分級方式　·　不做 → 全部拍照辨識"],
    ["2", "拍照模式可以接受多久出結果？", "有了上限才能在 fp32@320、fp32@640、w8a32@640 之間選"],
    ["3", "剩下的時間要不要繼續追模型精度？", "追 → 介殼蟲一致性或補拍照片　·　不追 → 收尾文件"],
  ];
  d.forEach((it, i) => {
    const y = 2.3 + i * 1.3;
    card(s, M, y, CW, 1.12, C.dark2);
    badge(s, it[0], M + 0.35, y + 0.28, 0.56, C.accent, C.dark);
    s.addText(it[1], { x: M + 1.15, y: y + 0.15, w: CW - 1.5, h: 0.45, isTextBox: true, margin: 0, fontFace: F, fontSize: 20, bold: true, color: C.white });
    s.addText(it[2], { x: M + 1.15, y: y + 0.62, w: CW - 1.5, h: 0.38, isTextBox: true, margin: 0, fontFace: F, fontSize: 14, color: C.sage });
  });

  s.addText("技術面三台都量清楚了。", {
    x: M, y: 6.35, w: CW, h: 0.5, isTextBox: true, margin: 0,
    fontFace: F, fontSize: 20, bold: true, color: C.accent,
  });
  s.addNotes("收尾把三個決策再念一次，然後開放討論。");
}

const OUT = process.argv[2] || "deck.pptx";
pres.writeFile({ fileName: OUT }).then(() => console.log("written:", OUT));
