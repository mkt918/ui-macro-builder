/* 割り勘アプリのプリント（A4・おもて／うら＝1ステップ1枚、全8ページ）を作る。
 *   node handouts/build-warikan-print.js   →   handouts/warikan-print.html
 * 仕様の表は src/tasks/textbook-tasks.js の課題データ（ORIGINAL_TASKS）から取る。
 * 流れ図・ミッションの絵・解説・穴うめのプログラムは、このファイルに書いてある。
 * 長さは mm のつもりで書いた SVG（viewBox の 1 = 1mm）。
 */
const fs = require("fs");
const path = require("path");
const { ORIGINAL_TASKS } = require(path.join(__dirname, "../src/tasks/textbook-tasks.js"));

const esc = (s) => String(s).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
const taskOf = (id) => ORIGINAL_TASKS.find((t) => t.id === id);

/* ======================= SVG の部品（単位 mm） ======================= */
const MONO = "font-family=\"Consolas,'BIZ UDGothic','MS Gothic',monospace\"";
const tw = (t, fs) => [...String(t)].reduce((a, c) => a + (c.charCodeAt(0) > 0x7f ? fs : fs * 0.56), 0);
const T = (x, y, s, o = {}) =>
  `<text x="${x}" y="${y}" font-size="${o.fs || 3.3}"${o.b ? ' font-weight="700"' : ""}${o.a ? ` text-anchor="${o.a}"` : ""}` +
  `${o.f ? ` fill="${o.f}"` : ""}${o.mono ? " " + MONO : ""}>${esc(s)}</text>`;
const R = (x, y, w, h, o = {}) =>
  `<rect x="${x}" y="${y}" width="${w}" height="${h}" rx="${o.rx === undefined ? 1.5 : o.rx}" fill="${o.fill || "#fff"}" ` +
  `stroke="${o.stroke || "#222"}" stroke-width="${o.sw || 0.35}"${o.dash ? ' stroke-dasharray="1.4 0.9"' : ""}/>`;
const tag = (x, y, w, s) => R(x, y, w, 5, { fill: "#111", rx: 1 }) + T(x + w / 2, y + 3.6, s, { a: "middle", f: "#fff", b: 1, fs: 3 });
const bigArrow = (x, y, w, h) =>
  `<polygon points="${x},${y + h * 0.3} ${x + w * 0.55},${y + h * 0.3} ${x + w * 0.55},${y} ${x + w},${y + h / 2} ${x + w * 0.55},${y + h} ${x + w * 0.55},${y + h * 0.7} ${x},${y + h * 0.7}" fill="#cfcfcf" stroke="#222" stroke-width="0.35"/>`;
const svgWrap = (w, h, inner, label, cls) => `<svg class="${cls || "mimg"}" viewBox="0 0 ${w} ${h}" role="img" aria-label="${esc(label)}">${inner}</svg>`;

// ユーザーフォームの絵。rows: [{cap, val, mark}]、mark は「追加した枠」（点線でかこむ）
function formG(x, y, w, rows, o = {}) {
  const rh = o.rh || 7.4;
  const bh = rh - 1.8; // 枠の高さ
  const top = 5.2;
  const h = top + rows.length * rh + (o.button ? 10 : 2);
  let g = R(x, y, w, h, { fill: "#f0f0f0", sw: 0.5, rx: 0.8 });
  g += R(x, y, w, top, { fill: "#d6d6d6", sw: 0.5, rx: 0.8 });
  g += T(x + 1.8, y + 3.7, "UserForm1", { fs: 2.7 }) + T(x + w - 2.6, y + 4.1, "×", { fs: 3.6, b: 1, a: "middle" });
  rows.forEach((r, i) => {
    const ry = y + top + 1.4 + i * rh;
    if (r.mark) g += R(x + 0.8, ry - 0.9, w - 1.6, rh - 0.4, { fill: "none", dash: 1, sw: 0.55, stroke: "#000", rx: 1 });
    g += T(x + 3, ry + bh * 0.75, r.cap, { fs: 3.1, b: 1 });
    g += R(x + w * 0.43, ry, w * 0.54, bh, { sw: 0.5, rx: 0.3 });
    if (r.val) g += T(x + w * 0.7, ry + bh * 0.72, r.val, { fs: 3.3, a: "middle", b: 1 });
  });
  if (o.button) {
    const by = y + top + rows.length * rh + 2.2;
    g += R(x + w / 2 - 11, by, 22, 6, { fill: "#f7f7f7", sw: 0.45, rx: 1 }) + T(x + w / 2, by + 4.2, "計算する", { fs: 2.9, a: "middle" });
  }
  return { svg: g, h, rowY: (i) => y + top + 1.4 + i * rh + bh / 2 };
}
const callout = (x1, y, x2, name) =>
  `<line x1="${x1}" y1="${y}" x2="${x2 - 1}" y2="${y}" stroke="#222" stroke-width="0.35"/>` + T(x2, y + 1, name, { fs: 2.8, b: 1, mono: 1 });

/* ======================= ミッションの絵 ======================= */
const appBox = (x, y, w, h) =>
  R(x, y, w, h, { fill: "#ececec", rx: 3, sw: 0.5 }) +
  T(x + w / 2, y + h * 0.42, "💻", { fs: 8, a: "middle" }) +
  T(x + w / 2, y + h * 0.42 + 6, "割り勘", { fs: 3.5, a: "middle", b: 1 }) +
  T(x + w / 2, y + h * 0.42 + 10.4, "アプリ", { fs: 3.5, a: "middle", b: 1 });
const inCards = (rows) => rows.map((s, i) => R(0, 6 + i * 6.6, 60, 5.9) + T(2.5, 10.2 + i * 6.6, s, { fs: 3.3 })).join("");

const missionImg = {
  warikan1: () =>
    svgWrap(
      172,
      33,
      tag(0, 0, 13, "入力") + inCards(["🧃 ジュース　300円", "🍕 ピザ　1200円", "🍰 ケーキ　900円", "👥 人数　3人"]) +
        bigArrow(62, 12.5, 9, 8) + appBox(73, 6, 30, 26) + bigArrow(105, 12.5, 9, 8) +
        tag(116, 0, 13, "出力") +
        R(116, 6, 56, 12, { dash: 1, sw: 0.5 }) + T(144, 13.8, "合計は ？円", { fs: 5.2, a: "middle", b: 1 }) +
        R(116, 20, 56, 12, { dash: 1, sw: 0.5 }) + T(144, 27.8, "1人 ？円", { fs: 5.2, a: "middle", b: 1 }),
      "ミッションの絵。ジュース300円、ピザ1200円、ケーキ900円、3人を入力すると、合計と1人あたりの金額が出る"
    ),
  warikan2: () =>
    svgWrap(
      172,
      33,
      tag(0, 0, 13, "入力") + inCards(["💰 合計金額　1000円", "👥 人数　3人"]) +
        T(2.5, 27, "🙂🙂🙂", { fs: 7 }) + T(26, 27, "3人で分けると…？", { fs: 3.4, b: 1 }) +
        bigArrow(62, 12.5, 9, 8) + appBox(73, 6, 30, 26) + bigArrow(105, 12.5, 9, 8) +
        tag(116, 0, 13, "出力") +
        R(116, 6, 56, 12, { dash: 1, sw: 0.5 }) + T(144, 13.8, "1人 ？円", { fs: 5.2, a: "middle", b: 1 }) +
        R(116, 20, 56, 12, { dash: 1, sw: 0.5 }) + T(144, 27.8, "あまり ？円", { fs: 5.2, a: "middle", b: 1 }),
      "ミッションの絵。合計金額1000円、人数3人を入力すると、1人あたりの金額とあまりが出る"
    ),
  warikan3: () => {
    // 前：InputBox で聞いて MsgBox で答える（小さな窓2つ）／後：フォームの TextBox に書いて Label に出す
    const dlg = (y, h, title, body) =>
      R(2, y, 56, h, { rx: 0.6, sw: 0.5 }) + R(2, y, 56, 4.6, { fill: "#c9d3e2", rx: 0.6, sw: 0.5 }) +
      T(4, y + 3.3, title, { fs: 2.5 }) + T(55.5, y + 3.6, "×", { fs: 3, b: 1, a: "middle" }) + body;
    const f = formG(86, 8, 62, [
      { cap: "合計金額", val: "4580" },
      { cap: "人数", val: "6" },
      { cap: "1人あたり", val: "763円" },
      { cap: "あまり", val: "2円" },
    ], { rh: 6.9 });
    return svgWrap(
      172,
      f.h + 10,
      tag(0, 0, 66, "ステップ2：InputBox と MsgBox（小さな窓）") +
        dlg(7.5, 18, "Microsoft Excel", T(5, 15.4, "合計金額を入力してください", { fs: 2.6 }) + R(5, 17.6, 36, 4.4, { sw: 0.4, rx: 0.3 }) + T(7, 20.7, "4580", { fs: 2.8 })) +
        dlg(28, 14.5, "Microsoft Excel", T(30, 36.4, "1人763円です", { fs: 3.6, a: "middle", b: 1 }) + R(23, 37.4, 14, 3.8, { fill: "#f4f4f4", sw: 0.4, rx: 0.8 })) +
        bigArrow(62, 17, 20, 10) + T(72, 15, "かわる！", { fs: 3.8, a: "middle", b: 1 }) +
        tag(86, 0, 62, "ステップ3：フォームの TextBox と Label") + f.svg +
        callout(149, f.rowY(0), 151, "TextBox1") + callout(149, f.rowY(1), 151, "TextBox2") +
        callout(149, f.rowY(2), 151, "Label3") + callout(149, f.rowY(3), 151, "Label4"),
      "ミッションの絵。InputBox と MsgBox の小さな窓で入力と表示をしていたのが、ユーザーフォームの TextBox1、TextBox2 に書いて、Label3、Label4 に出すように変わる"
    );
  },
  warikan4: () => {
    const rowsOld = [
      { cap: "合計金額", val: "4580" },
      { cap: "人数", val: "6" },
      { cap: "1人あたり", val: "763円" },
      { cap: "あまり", val: "2円" },
    ];
    const f1 = formG(0, 8, 56, rowsOld, { rh: 6.2 });
    const f2 = formG(78, 8, 62, rowsOld.concat([{ cap: "幹事", val: "765円", mark: true }, { cap: "ほかの人", val: "763円", mark: true }]), { rh: 6.2 });
    return svgWrap(
      172,
      f2.h + 10,
      tag(0, 0, 36, "ステップ3のフォーム") + f1.svg +
        T(66, 22, "足す！", { fs: 3.8, a: "middle", b: 1 }) + bigArrow(58, 24, 16, 9) +
        tag(78, 0, 56, "新しいフォーム（点線が追加）") + f2.svg +
        callout(141, f2.rowY(4), 143, "←追加 Label7,8") + callout(141, f2.rowY(5), 143, "←追加 Label9,10"),
      "ミッションの絵。ステップ3のフォームの下に、幹事とほかの人の Label を足したフォーム"
    );
  },
};

/* ---- 回転ずしアプリのミッションの絵 ---- */
// 回転ずしのフォームの絵。rows: [{btn, lab, sub}]、wari があれば下に「割り勘」の行（追加＝点線）
function sushiG(x, y, w, o) {
  const top = 5.2;
  const rh = 6.8;
  const h = top + 1.2 + o.rows.length * rh + 9 + (o.wari ? rh + 2.4 : 0);
  let g = R(x, y, w, h, { fill: "#f0f0f0", sw: 0.5, rx: 0.8 }) + R(x, y, w, top, { fill: "#d6d6d6", sw: 0.5, rx: 0.8 });
  g += T(x + 1.8, y + 3.7, "UserForm1", { fs: 2.7 }) + T(x + w - 2.6, y + 4.1, "×", { fs: 3.6, b: 1, a: "middle" });
  const btn = (bx, by, s) => R(bx, by, 28, 5.4, { fill: "#f7f7f7", rx: 1, sw: 0.4 }) + T(bx + 14, by + 3.9, s, { fs: 2.7, a: "middle" });
  const lab = (lx, ly, lw, s) => R(lx, ly, lw, 5.4, { sw: 0.4, rx: 0.3 }) + (s ? T(lx + lw / 2, ly + 3.9, s, { fs: 2.8, a: "middle", b: 1 }) : "");
  o.rows.forEach((r, i) => {
    const ry = y + top + 1.2 + i * rh;
    g += btn(x + 2, ry, r.btn) + lab(x + 34, ry, 38, r.lab) + lab(x + 77, ry, 39, r.sub);
  });
  const ty = y + top + 1.2 + o.rows.length * rh + 1.4;
  g += lab(x + 34, ty, 82, o.total);
  if (o.wari) {
    const wy = ty + 7.2;
    g += R(x + 0.8, wy - 1, w - 1.6, rh + 0.2, { fill: "none", dash: 1, sw: 0.55, stroke: "#000", rx: 1 });
    g += btn(x + 2, wy, "割り勘") + lab(x + 34, wy, 38, o.wari[0]) + lab(x + 77, wy, 39, o.wari[1]);
  }
  return { svg: g, h, rowY: (i) => y + top + 1.2 + i * rh + 2.7, totalY: ty + 2.7, wariY: ty + 7.2 + 2.7 };
}

missionImg.sushi1 = () => {
  const f = sushiG(0, 0, 118, {
    rows: [
      { btn: "100円皿", lab: "100円皿2枚", sub: "小計200円" },
      { btn: "120円皿", lab: "120円皿1枚", sub: "小計120円" },
      { btn: "150円皿", lab: "", sub: "" },
      { btn: "200円皿", lab: "", sub: "" },
    ],
    total: "合計320円",
  });
  return svgWrap(
    172,
    f.h + 1.5,
    f.svg +
      T(146, 11, "🍣", { fs: 11, a: "middle" }) +
      T(122, 20, "ボタンを押すたびに", { fs: 3.2, b: 1 }) +
      T(124, 25.5, "・お皿の枚数がふえる", { fs: 3.1 }) +
      T(124, 30.5, "・小計が出る", { fs: 3.1 }) +
      T(124, 35.5, "・合計がふえる", { fs: 3.1 }),
    "ミッションの絵。100円皿、120円皿、150円皿、200円皿のボタンがあり、押すと枚数、小計、合計が表示される"
  );
};

missionImg.sushi2 = () => {
  const f = sushiG(0, 0, 118, {
    rows: [
      { btn: "100円皿", lab: "100円皿2枚", sub: "小計200円" },
      { btn: "120円皿", lab: "120円皿1枚", sub: "小計120円" },
      { btn: "150円皿", lab: "", sub: "" },
      { btn: "200円皿", lab: "", sub: "" },
    ],
    total: "合計320円",
    wari: ["人数3人", "1人あたり106円"],
  });
  return svgWrap(
    172,
    f.h + 1.5,
    f.svg +
      T(122, f.wariY - 4.2, "追加！", { fs: 3.6, b: 1 }) +
      callout(119, f.wariY, 122, "") +
      T(122, f.wariY - 0.2, "CommandButton5", { fs: 2.7, b: 1, mono: 1 }) +
      T(122, f.wariY + 3.4, "Label10・Label11", { fs: 2.7, b: 1, mono: 1 }) +
      T(146, 11, "🍣", { fs: 11, a: "middle" }) +
      T(122, 22, "合計は、そのまま使うよ", { fs: 3.1 }),
    "ミッションの絵。回転ずしのフォームの下に、割り勘ボタンと、人数、1人あたりの金額を出す枠を足す"
  );
};

/* ======================= 流れ図の部品 ======================= */
// 記号は教科書の描き方：はじめ・おわり＝角丸、入出力＝平行四辺形、計算＝長方形
const FC = { BX: 4, BW: 132, W: 146, FS: 3.1, GAP: 3.0, TH: 7, BH: 8.2, SK: 4 };
function segs(x, y, list) {
  let out = "";
  list.forEach((s) => {
    if (s.t !== undefined) {
      out += T(x, y, s.t, { fs: FC.FS, f: "#555" });
      x += tw(s.t, FC.FS) + 0.6;
    } else {
      out += `<line x1="${x}" y1="${y + 1}" x2="${x + s.w}" y2="${y + 1}" stroke="#222" stroke-width="0.3"/>`;
      if (s.fill) out += T(x + 0.8, y - 0.1, s.fill, { fs: FC.FS, f: "#1f4e8c", b: 1 });
      x += s.w + 0.8;
    }
  });
  return out;
}
const fT = (t) => ({ k: "term", t });
// 前のステップと同じ部分は、1つの灰色の箱にまとめる（書かなくてよい）
const fSame = (t) => ({ k: "same", t });
// 流れ図の箱。引数が文字なら青い「お手本」が書いてあり、null なら書きこむ空欄。
// 空欄は、1つの流れ図に3つぐらい（確認程度）にする。
const fIn = (word, v) => ({ k: "io", s: [{ t: "ことば" }, { w: 70, fill: word }, { t: "→変数" }, { w: 20, fill: v }, { t: "を入力" }] });
const fMsg = (word) => ({ k: "io", s: [{ t: "ことば" }, { w: 90, fill: word }, { t: "を表示" }] });
const fTextIn = (cap, tb, v) => ({ k: "proc", s: [{ t: "見出し" }, { w: 28, fill: cap }, { t: "の" }, { t: "TextBox" }, { w: 8, fill: tb }, { t: "→変数" }, { w: 22, fill: v }] });
const fLabel = (word, num, add) => ({ k: "io", add, s: [{ w: 62, fill: word }, { t: "を" }, { t: "Label" }, { w: 12, fill: num }, { t: "に表示" }] });
const fCalc = (a, op, b, c, add) => ({ k: "proc", add, s: [{ w: 26, fill: a }, { t: op }, { w: 26, fill: b }, { t: "→" }, { w: 26, fill: c }] });
const fSum3 = (a, b, c, d) => ({ k: "proc", s: [{ w: 26, fill: a }, { t: "＋" }, { w: 26, fill: b }, { t: "＋" }, { w: 26, fill: c }, { t: "→" }, { w: 22, fill: d }] });

function flowchart(items) {
  const { BX, BW, W, FS, GAP, TH, BH, SK } = FC;
  const cx = BX + BW / 2;
  let y = 1;
  let g = "";
  items.forEach((it, i) => {
    const h = it.k === "term" ? TH : BH;
    const sw = it.add ? 0.9 : 0.5;
    const dash = it.add ? ' stroke-dasharray="2 1"' : "";
    if (it.k === "same") {
      g += `<rect x="${BX}" y="${y}" width="${BW}" height="${h}" rx="1.5" fill="#ececec" stroke="#777" stroke-width="0.45" stroke-dasharray="1.6 1"/>`;
      g += T(cx, y + h / 2 + 1.1, it.t, { fs: FS, a: "middle", f: "#444" });
    } else if (it.k === "term") {
      g += `<rect x="${cx - 14}" y="${y}" width="28" height="${h}" rx="${h / 2}" fill="#fff" stroke="#222" stroke-width="0.5"/>`;
      g += T(cx, y + h / 2 + 1.1, it.t, { fs: FS, a: "middle", b: 1 });
    } else {
      g +=
        it.k === "io"
          ? `<polygon points="${BX + SK},${y} ${BX + BW},${y} ${BX + BW - SK},${y + h} ${BX},${y + h}" fill="#fff" stroke="#222" stroke-width="${sw}"${dash}/>`
          : `<rect x="${BX}" y="${y}" width="${BW}" height="${h}" fill="#fff" stroke="#222" stroke-width="${sw}"${dash}/>`;
      g += segs(BX + (it.k === "io" ? SK + 2.5 : 3), y + h / 2 + 1.1, it.s);
      if (it.add) g += R(BX + BW + 1, y + h / 2 - 2.5, 8, 5, { fill: "#111", rx: 1 }) + T(BX + BW + 5, y + h / 2 + 1.1, "追加", { fs: 2.9, a: "middle", f: "#fff", b: 1 });
    }
    if (i < items.length - 1) {
      g += `<line x1="${cx}" y1="${y + h}" x2="${cx}" y2="${y + h + GAP - 0.4}" stroke="#222" stroke-width="0.45"/>`;
      g += `<polygon points="${cx - 0.9},${y + h + GAP - 1.2} ${cx + 0.9},${y + h + GAP - 1.2} ${cx},${y + h + GAP}" fill="#222"/>`;
    }
    y += h + GAP;
  });
  return svgWrap(W, y - GAP + 1, g, "流れ図を書く欄。はじめ、入力、計算、表示、おわりの記号がならんでいる", "flow");
}

const legend = (extra) => {
  const icon = (inner) => `<svg class="lg-svg" viewBox="0 0 14 7">${inner}</svg>`;
  return `<div class="legend">
    <div class="lg-t">記号のやくそく</div>
    <div class="lg-i">${icon(R(1, 1, 12, 5, { rx: 2.5, sw: 0.5 }))}<span>はじめ・おわり</span></div>
    <div class="lg-i">${icon('<polygon points="3,1 13,1 11,6 1,6" fill="#fff" stroke="#222" stroke-width="0.5"/>')}<span>入力・表示</span></div>
    <div class="lg-i">${icon(R(1, 1, 12, 5, { rx: 0, sw: 0.5 }))}<span>計算</span></div>
    <div class="lg-n"><b>ことば</b>＝画面に出る文字<br>（InputBox や MsgBox のメッセージ）</div>
    <div class="lg-n"><b>変数</b>＝数を入れる箱の名前</div>
    ${extra || ""}
    <div class="lg-n blue">青い字はお手本</div>
  </div>`;
};

/* ======================= 文章の部品 ======================= */
const blankHtml = (s) =>
  esc(s).replace(/\[\[(s|m|xl)?\]\]/g, (_, k) => `<span class="blank ${k || ""}"></span>`);

function prog(lines) {
  return (
    '<div class="codebox">' +
    lines
      .map((l0) => {
        let l = l0;
        const neu = l.startsWith("★");
        if (neu) l = l.slice(1);
        const [body, note] = l.split("||");
        const ind = (body.match(/^ */)[0].length / 4) * 2;
        return `<div class="ln${neu ? " new" : ""}" style="padding-left:${ind}em">${blankHtml(body.trim())}${note ? `<span class="lnote">← ${esc(note)}</span>` : ""}</div>`;
      })
      .join("") +
    "</div>"
  );
}

function lessonHtml(items) {
  return (
    '<div class="box lesson">' +
    items
      .map(
        (it) =>
          `<div class="item"><span class="num">${esc(it.n)}</span><div>${it.h ? `<div class="t">${esc(it.h)}</div>` : ""}` +
          (it.p ? `<p>${it.p}</p>` : "") +
          (it.code ? `<span class="snip">${esc(it.code)}</span>` : "") +
          (it.extra || "") +
          "</div></div>"
      )
      .join("") +
    "</div>"
  );
}

function specHtml(sp) {
  return `<table class="spec">
    <tr><th>入力データ</th><td>${esc(sp.input)}</td><th>出力データ</th><td>${esc(sp.output)}</td></tr>
    <tr><th>処理条件</th><td colspan="3"><ul>${sp.conditions.map((c) => `<li>${esc(c)}</li>`).join("")}</ul></td></tr>
    <tr><th>入力例</th><td>${esc(sp.inputExample)}</td><th>出力例</th><td>${esc(sp.outputExample)}</td></tr>
  </table>`;
}

function tryHtml(head, rows) {
  return `<table class="try"><tr>${head.map((h) => `<th>${esc(h)}</th>`).join("")}<th>できた？</th></tr>${rows
    .map((r) => `<tr>${r.map((c) => `<td>${esc(c)}</td>`).join("")}<td class="chk">☐</td></tr>`)
    .join("")}</table>`;
}

const ROAD = ["① ミッション", "② 仕様", "③ 流れ図", "④ 読む", "⑤ プログラム", "⑥ 動かす"];
const road = (on) => `<ol class="road">${ROAD.map((r, i) => `<li${on.includes(i) ? ' class="on"' : ""}>${r}</li>`).join("")}</ol>`;

/* ======================= ステップの中身 ======================= */
const WARIKAN_STEPS = [
  {
    id: "warikan1",
    lead: "文化祭の打ち上げで、ジュース・ピザ・ケーキを注文しました。",
    main: "3つの値段と人数を入力すると、<u>合計</u>と<u>1人あたりの金額</u>を表示するプログラムを作ろう！",
    flow: [
      fT("はじめ"),
      fIn("ジュースの値段を入力してください", "Juice"),
      fIn("ピザの値段を入力してください", "Pizza"),
      fIn(null, "Cake"),
      fIn("人数を入力してください", null),
      fSum3("Juice", "Pizza", "Cake", "Gokei"),
      fMsg("合計は Gokei 円です"),
      fCalc("Gokei", "÷", "Ninzu", "Hitori"),
      fMsg(null),
      fT("おわり"),
    ],
    legendExtra: "",
    lesson: [
      { n: "1", h: "値段を聞いて、変数（数を入れる箱）にしまう", p: "InputBox で聞いた答えは「文字」です。<b>Val( )</b> をつけると「数」になります。<span class=\"note\">（Long は大きめの整数が入る型）</span>", code: 'Dim Juice As Long\nJuice = Val(InputBox("ジュースの値段を入力してください"))' },
      { n: "2", h: "たし算をして、合計を出す", p: "たし算は「<b>+</b>」です。答えは変数 Gokei（合計）にしまいます。", code: 'Gokei = Juice + Pizza + Cake\nMsgBox ("合計は" & Gokei & "円です")' },
      { n: "3", h: "人数でわり算をする", p: "わり算は「<b>/</b>」です。", code: 'Hitori = Gokei / Ninzu\nMsgBox ("1人" & Hitori & "円です")' },
    ],
    note: "流れ図で書いた「ことば」を InputBox のかっこの中に、「変数」を = の左に書き写そう。",
    program: [
      "Private Sub CommandButton1_Click()",
      "    Dim Juice As Long",
      "    Dim Pizza As Long",
      "    Dim Cake As Long",
      "    Dim Ninzu As Long",
      "    Dim Gokei As Long",
      "    Dim Hitori As Long",
      '    Juice = Val(InputBox("ジュースの値段を入力してください"))',
      '    Pizza = Val(InputBox("[[xl]]"))',
      '    Cake = Val(InputBox("[[xl]]"))',
      '    [[m]] = Val(InputBox("[[xl]]"))',
      "    Gokei = Juice + [[m]] + [[m]]",
      '    MsgBox ("合計は" & Gokei & "円です")',
      "    Hitori = [[m]] / [[m]]",
      '    MsgBox ("1人" & [[m]] & "円です")',
      "End Sub",
    ],
    tryHead: ["ジュース", "ピザ", "ケーキ", "人数", "合計", "1人あたり"],
    tryRows: [["300", "1200", "900", "3", "2400", "800"], ["500", "900", "100", "5", "1500", "300"]],
    foot: "※ 1000円を3人でわると 333.33…円。割り切れないときは、次のステップで考えます。",
  },
  {
    id: "warikan2",
    lead: "1000円を3人で分けると、1人333円で1円あまります。お金に小数はありません。",
    main: "合計金額と人数を入力すると、<u>1人あたりの金額</u>と<u>あまり</u>を表示するプログラムを作ろう！",
    flow: [
      fT("はじめ"),
      fIn("合計金額を入力してください", "Gokei"),
      fIn("人数を入力してください", null),
      fCalc("Gokei", "¥", "Ninzu", "Hitori"),
      fCalc("Gokei", "Mod", "Ninzu", null),
      fMsg("1人 Hitori 円です"),
      fMsg(null),
      fT("おわり"),
    ],
    legendExtra: '<div class="lg-n"><b>¥</b>＝わり算の答えの整数の部分<br><b>Mod</b>＝わり算のあまり</div>',
    lesson: [
      { n: "1", h: "割り切れないとき", p: "1000円を3人で分けると、<b>1人333円</b>で、<b>1円あまり</b>ます。このあまりまで考えるのが「割り勘」です。<br>たしかめ算：<b>333 × 3 + 1 = 1000</b> になれば、正しく分けられています。" },
      { n: "2", h: "この2つの記号を覚えよう", extra: '<table class="ops"><tr><th>¥</th><td>わり算の答えの<b>整数の部分</b>（分けた金額）</td><td><code>1000 ¥ 3</code> → <b>333</b></td></tr><tr><th>Mod</th><td>わり算の<b>あまり</b></td><td><code>1000 Mod 3</code> → <b>1</b></td></tr></table><p class="note">¥ はキーボードの ¥ キーで打てます。</p>' },
      { n: "3", h: "プログラムにすると", code: "Hitori = Gokei ¥ Ninzu\nAmari = Gokei Mod Ninzu" },
      { n: "!", h: "ちゅうい：「/」と「¥」はちがう", p: "「/」は四捨五入されて、合計とずれることがあります。割り勘では「¥」を使おう。" },
    ],
    note: "流れ図の ¥ と Mod を見ながら、＿＿に ¥ か Mod を書こう。",
    program: [
      "Private Sub CommandButton1_Click()",
      "    Dim Gokei As Long",
      "    Dim Ninzu As Long",
      "    Dim Hitori As Long",
      "    Dim Amari As Long",
      '    Gokei = Val(InputBox("合計金額を入力してください"))',
      '    [[m]] = Val(InputBox("[[xl]]"))',
      "    Hitori = Gokei [[s]] Ninzu",
      "    Amari = Gokei [[s]] Ninzu",
      '    MsgBox ("1人" & [[m]] & "円です")',
      '    MsgBox ("あまりは" & [[m]] & "円です")',
      "End Sub",
    ],
    tryHead: ["合計金額", "人数", "1人あたり", "あまり", "たしかめ算（1人あたり × 人数 ＋ あまり）"],
    tryRows: [["1000", "3", "333", "1", "333 × 3 + 1 = 1000"], ["3000", "3", "1000", "0", "1000 × 3 + 0 = 3000"], ["4580", "6", "763", "2", "763 × 6 + 2 = 4580"], ["1100", "3", "366", "2", "366 × 3 + 2 = 1100"]],
    foot: "※ 1100円・3人のとき、「/」で書くと 367円になってしまいます。「¥」で 366円になれば OK！",
  },
  {
    id: "warikan3",
    lead: "ステップ2では、InputBox で答え、MsgBox で表示しました（どちらも小さな窓）。",
    main: "ステップ3では、<u>フォームの TextBox に書いて</u>、<u>フォームの Label に出す</u>、本物の「割り勘アプリ」にしよう！",
    band: "InputBox → TextBox　／　MsgBox → Label　に、かわるよ！",
    flow: [
      fT("はじめ"),
      fTextIn("合計金額", "1", null),
      fTextIn("人数", null, "Ninzu"),
      fCalc("Gokei", "¥", "Ninzu", "Hitori"),
      fCalc("Gokei", "Mod", "Ninzu", "Amari"),
      fLabel("Hitori と 円", "3"),
      fLabel(null, "4"),
      fT("おわり"),
    ],
    legendExtra: '<div class="lg-n"><b>TextBox</b>から読むときは<br>「見出し ○○ の TextBox1 → 変数」</div><div class="lg-n"><b>Label</b>に出すときは<br>「○○を Label3 に表示」</div>',
    lesson: [
      { n: "1", h: "入力する場所が、かわる", p: "ステップ2では InputBox に答えました。ステップ3では、フォームの <b>TextBox</b> に書きます。中身は「TextBox1.Text」。文字なので <b>Val( )</b> で数にします。", code: 'ステップ2：Gokei = Val(InputBox("合計金額を入力してください"))\nステップ3：Gokei = Val(TextBox1.Text)' },
      { n: "2", h: "答えを出す場所も、かわる", p: "MsgBox ではなく、フォームの <b>Label</b> に出します。<b>Label名.Caption = 入れたいもの</b>。「円」は & でつなぎます。<br>Label3 は上の枠（1人あたり）、Label4 は下の枠（あまり）です。", code: 'ステップ2：MsgBox ("1人" & Hitori & "円です")\nステップ3：Label3.Caption = Hitori & "円"' },
      { n: "3", h: "ほかはステップ2と同じ", p: "InputBox と MsgBox は使いません。Dim・¥・Mod は、ステップ2と同じです。" },
    ],
    note: "ステップ2から、入力の2行と、いちばん下の2行が新しくなったよ。",
    program: [
      "Private Sub CommandButton1_Click()",
      "    Dim Gokei As Long",
      "    Dim Ninzu As Long",
      "    Dim Hitori As Long",
      "    Dim Amari As Long",
      "★    Gokei = Val(TextBox[[s]].Text)||InputBox をやめて TextBox に！",
      "★    Ninzu = Val(TextBox[[s]].Text)||InputBox をやめて TextBox に！",
      "    Hitori = Gokei ¥ Ninzu",
      "    Amari = Gokei Mod Ninzu",
      '★    Label[[s]].Caption = [[m]] & "円"||MsgBox をやめて Label に！',
      '★    Label[[s]].Caption = [[m]] & "円"||MsgBox をやめて Label に！',
      "End Sub",
    ],
    tryHead: ["TextBox1（合計金額）", "TextBox2（人数）", "Label3（1人あたり）", "Label4（あまり）"],
    tryRows: [["3000", "3", "1000円", "0円"], ["1000", "3", "333円", "1円"], ["4580", "6", "763円", "2円"], ["500", "7", "71円", "3円"]],
    foot: "※ 小さな窓が出たら、InputBox か MsgBox を消し忘れています。",
  },
  {
    id: "warikan4",
    lead: "あまったお金は、幹事が引き受けます。",
    main: "ステップ3のフォームの下に、<u>「幹事」と「ほかの人」の金額を出す枠（Label）</u>を足そう！",
    band: "★ 新しく足すもの：Label を4つ（Label7・8・9・10）",
    flow: [
      fT("はじめ"),
      fSame("ステップ3と同じ：TextBox から読む → ¥ と Mod の計算"),
      fCalc("Hitori", "＋", null, "Kanji", true),
      fSame("ステップ3と同じ：Label3・Label4 に表示"),
      fLabel(null, "8", true),
      fLabel("Hitori と 円", null, true),
      fT("おわり"),
    ],
    legendExtra: '<div class="lg-n"><b>灰色の箱</b>…ステップ3と同じ（書かなくてOK）</div><div class="lg-n"><b>追加</b>…ステップ3から<br>新しく足すところ。ここだけ書く</div>',
    lesson: [
      { n: "1", h: "幹事の金額を考える", p: "4580円を6人で分けると、1人763円であまりは2円。ほかの5人が763円ずつ出すなら、幹事は <b>763 + 2 = 765円</b> 出せば、ぴったり4580円になります。<br>たしかめ算：<b>765 + 763 × 5 = 4580</b>" },
      { n: "2", h: "フォームに足す Label（ここが新しい！）", p: "ボタンを下にずらして、あいたところに Label を4つ足します。<br><b>Label7</b>＝見出し（Caption を「幹事」）　<b>Label8</b>＝幹事の金額を出す枠<br><b>Label9</b>＝見出し（Caption を「ほかの人」）　<b>Label10</b>＝ほかの人の金額を出す枠<br><span class=\"note\">Label は作った順に Label7・8・9・10 という名前になります。</span>" },
    ],
    note: "ステップ3の続きに、「追加」の行を足そう。",
    program: [
      "Private Sub CommandButton1_Click()",
      "    Dim Gokei As Long",
      "    Dim Ninzu As Long",
      "    Dim Hitori As Long",
      "    Dim Amari As Long",
      "★    Dim [[m]] As Long||追加（幹事の金額を入れる箱）",
      "    Gokei = Val(TextBox1.Text)",
      "    Ninzu = Val(TextBox2.Text)",
      "    Hitori = Gokei ¥ Ninzu",
      "    Amari = Gokei Mod Ninzu",
      "★    [[m]] = Hitori + [[m]]||追加（幹事 ＝ 1人あたり ＋ あまり）",
      '    Label3.Caption = Hitori & "円"',
      '    Label4.Caption = Amari & "円"',
      '★    Label[[s]].Caption = [[m]] & "円"||追加（幹事）',
      '★    Label[[s]].Caption = [[m]] & "円"||追加（ほかの人）',
      "End Sub",
    ],
    tryHead: ["TextBox1（合計金額）", "TextBox2（人数）", "Label3（1人あたり）", "Label4（あまり）", "Label8（幹事）", "Label10（ほかの人）"],
    tryRows: [["3000", "3", "1000円", "0円", "1000円", "1000円"], ["1000", "3", "333円", "1円", "334円", "333円"], ["4580", "6", "763円", "2円", "765円", "763円"], ["1100", "3", "366円", "2円", "368円", "366円"]],
    foot: "🎉 ここまでできたら、割り勘アプリの完成です！",
  },
];

/* ======================= ページを組む ======================= */
// 4桁番号（年・組・番号を数字4つで）と名前。各ステップの最後（うらの右下）に置く
const whoTable =
  '<table class="who"><tr><th>4桁番号</th><td class="digits"><i></i><i></i><i></i><i></i></td><th>名前</th><td class="nm"></td></tr></table>';
// 題名は横幅いっぱいの1行。ページ番号は右上
const headHtml = (t, tail, pg) => `<div class="head">
    <div class="hrow"><span class="series">${t.seriesIcon || "🍕"} ${esc(t.series)}${tail}</span><span class="pg">${pg}</span></div>
    <div class="title"><span class="step-tag">${esc(t.step)}</span>${esc(t.title.replace(/^\S+\s*(ステップ\d|発展)\s*/, ""))}</div>
  </div>`;

function front(st, idx, total) {
  const t = taskOf(st.id);
  const kind = t.kind === "guide" ? "解説つき" : "自分で考えよう";
  return `<section class="page">
  ${headHtml(t, `　［${kind}］　<b>おもて：考える</b>`, `${idx * 2 + 1} / ${total}`)}
  ${road([0, 1, 2])}
  <div class="mission2">
    <div class="mhead">🎯 ① ミッション　<span>ここが一番大事！ どんなアプリを作るのか、イメージしよう</span></div>
    <div class="mbody">
      <p class="mlead">${st.lead}</p>
      <p class="mmain">${st.main}</p>
      ${st.band ? `<div class="mband">${st.band}</div>` : ""}
      ${missionImg[st.id]()}
    </div>
  </div>
  <div class="sec">
    <h2>② 仕様を確かめよう</h2>
    ${specHtml(t.spec)}
  </div>
  <div class="sec">
    <h2>③ 流れ図を書こう　<span class="h2note">${st.flowNote || "平行四辺形には、<b>画面に出す「ことば」</b>と<b>「変数名」</b>を必ず書く"}</span></h2>
    <div class="flowrow">${flowchart(st.flow)}${legend(st.legendExtra)}</div>
  </div>
  <div class="footer"><span>流れ図が書けたら、うら（作る）へ →</span></div>
</section>`;
}

function back(st, idx, total) {
  const t = taskOf(st.id);
  return `<section class="page">
  ${headHtml(t, "　<b>うら：作る</b>", `${idx * 2 + 2} / ${total}`)}
  ${road([3, 4, 5])}
  <div class="sec">
    <h2>④ まず読もう</h2>
    ${lessonHtml(st.lesson)}
  </div>
  <div class="sec">
    <h2>⑤ プログラムを完成させよう　<span class="h2note">${esc(st.note)}</span></h2>
    ${prog(st.program)}
    ${st.after || ""}
  </div>
  <div class="sec">
    <h2>⑥ 動かしてたしかめよう</h2>
    ${tryHtml(st.tryHead, st.tryRows)}
  </div>
  <div class="footer"><span>${esc(st.foot)}</span>${whoTable}</div>
</section>`;
}

/* ======================= CSS ======================= */
const CSS = `
  /* A4 たて・両面印刷用。1ステップ＝おもて（考える）＋うら（作る）。外部ファイルなしで印刷できる */
  :root { --ink: #111; --sub: #555; --line: #222; --soft: #f1f1f1; }
  @page { size: A4 portrait; margin: 9mm 13mm; }
  * { box-sizing: border-box; }
  html { background: #d8d8d8; }
  body { margin: 0; color: var(--ink); background: #fff; font-family: "BIZ UDPGothic", "Yu Gothic", "Meiryo", sans-serif; font-size: 9.5pt; line-height: 1.35; }
  svg { font-family: "BIZ UDPGothic", "Yu Gothic", "Meiryo", "Segoe UI Emoji", sans-serif; }
  .page { padding: 0; page-break-after: always; break-after: page; }
  .page:last-child { page-break-after: auto; break-after: auto; }
  @media screen {
    body { background: transparent; }
    .page { width: 210mm; min-height: 297mm; margin: 12px auto; padding: 9mm 13mm; background: #fff; box-shadow: 0 1px 6px rgba(0,0,0,.25); }
  }
  @media print { html, body { background: #fff; } .page { width: auto; min-height: 0; margin: 0; padding: 0; box-shadow: none; } }

  .head { border-bottom: 2.5px solid var(--line); padding-bottom: 3px; }
  .hrow { display: flex; justify-content: space-between; align-items: baseline; }
  .pg { font-size: 9pt; color: var(--sub); }
  .series { font-size: 9pt; color: var(--sub); }
  .title { width: 100%; white-space: nowrap; font-size: 17pt; font-weight: 700; line-height: 1.25; }
  .step-tag { display: inline-block; margin-right: 6px; padding: 0 9px; border: 2px solid var(--line); border-radius: 6px; font-size: 12pt; vertical-align: 1px; }
  /* 4桁番号・名前の表（うらの右下） */
  table.who { border-collapse: collapse; flex: none; color: var(--ink); }
  table.who th, table.who td { border: 1.5px solid var(--line); height: 8mm; padding: 0 5px; font-size: 9pt; }
  table.who th { background: var(--soft); font-weight: 700; white-space: nowrap; }
  table.who td.digits { padding: 0; width: 30mm; }
  table.who td.digits i { display: inline-block; width: 7.4mm; height: 7.6mm; border-right: 1px dotted #777; vertical-align: top; }
  table.who td.digits i:last-child { border-right: none; }
  table.who td.nm { width: 44mm; }

  /* 流れ（①〜⑥） */
  ol.road { list-style: none; display: flex; gap: 0; margin: 4px 0 0; padding: 0; font-size: 8.5pt; }
  ol.road li { flex: 1; text-align: center; padding: 1px 2px; border: 1.2px solid #999; color: #777; margin-left: -1px; }
  ol.road li.on { background: #111; color: #fff; border-color: #111; font-weight: 700; }

  /* ①〜⑥ のかたまりは、枠で囲む */
  .sec { margin-top: 4px; border: 1.8px solid var(--line); border-radius: 8px; padding: 0 8px 3px; }
  .sec .box { border: none; padding: 0; margin-top: 0; }
  .sec h2 { margin-top: 3px; }
  h2 { margin: 7px 0 3px; font-size: 10.5pt; padding-left: 7px; border-left: 5px solid var(--line); line-height: 1.3; }
  .h2note { font-size: 8.8pt; font-weight: 400; color: var(--sub); }

  /* ミッション（いちばん目立たせる） */
  .mission2 { margin-top: 6px; border: 3px solid #000; border-radius: 10px; overflow: hidden; }
  .mhead { background: #000; color: #fff; padding: 1px 10px; font-size: 11.5pt; font-weight: 700; }
  .mhead span { font-size: 8.5pt; font-weight: 400; margin-left: 6px; }
  .mbody { padding: 4px 10px 6px; }
  .mlead { margin: 0; font-size: 10pt; }
  .mmain { margin: 1px 0 2px; font-size: 12pt; font-weight: 700; line-height: 1.35; }
  .mmain u { text-decoration: none; background: linear-gradient(transparent 60%, #d7d7d7 60%); }
  .mband { margin: 3px 0 4px; padding: 2px 8px; border: 2px dashed #000; border-radius: 6px; font-size: 11.5pt; font-weight: 700; text-align: center; background: #f4f4f4; }
  .mimg { display: block; width: 100%; height: auto; }

  .box { border: 1.5px solid var(--line); border-radius: 6px; padding: 3px 10px 4px; margin-top: 3px; }
  .note { font-size: 9pt; color: var(--sub); }
  p { margin: 2px 0; }
  code { font-family: Consolas, "BIZ UDGothic", "MS Gothic", monospace; font-size: 9.5pt; }

  /* 解説 */
  .lesson { display: grid; gap: 3px; }
  .lesson .item { display: grid; grid-template-columns: 22px 1fr; gap: 0 6px; align-items: start; }
  .num { display: inline-block; width: 21px; height: 21px; line-height: 19px; text-align: center; border: 1.5px solid var(--line); border-radius: 50%; font-weight: 700; font-size: 10pt; }
  .item .t { font-weight: 700; }
  .snip { display: block; margin: 2px 0 0; padding: 1px 8px; border-left: 4px solid var(--line); background: var(--soft); white-space: pre-wrap; font-family: Consolas, "BIZ UDGothic", "MS Gothic", monospace; font-size: 9.5pt; line-height: 1.4; }
  .ops { width: 100%; border-collapse: collapse; margin-top: 3px; }
  .ops th, .ops td { border: 1.2px solid var(--line); padding: 1px 8px; text-align: left; font-weight: 400; }
  .ops th { background: var(--soft); font-weight: 700; width: 14%; }

  /* 仕様 */
  table.spec { width: 100%; border-collapse: collapse; margin-top: 3px; }
  table.spec th, table.spec td { border: 1.5px solid var(--line); font-size: 9pt; padding: 1px 7px; vertical-align: top; text-align: left; }
  table.spec th { width: 19mm; background: var(--soft); font-weight: 700; white-space: nowrap; }
  table.spec ul { margin: 0; padding-left: 16px; }
  table.spec li { margin: 0; }

  /* 流れ図 */
  .flowrow { display: flex; gap: 2mm; align-items: flex-start; margin-top: 3px; }
  svg.flow { flex: none; width: 146mm; height: auto; }
  .legend { flex: 1; border: 1.5px solid var(--line); border-radius: 6px; padding: 3px 6px; font-size: 8.5pt; line-height: 1.3; }
  .lg-t { font-weight: 700; margin-bottom: 2px; }
  .lg-i { display: flex; align-items: center; gap: 4px; margin-bottom: 1px; }
  .lg-svg { width: 11mm; height: auto; flex: none; }
  .lg-n { margin-top: 3px; }
  .lg-n.blue { color: #1f4e8c; font-weight: 700; }

  /* プログラム */
  .codebox { margin-top: 3px; border: 1.5px solid var(--line); border-radius: 6px; padding: 3px 10px 4px; font-family: Consolas, "BIZ UDGothic", "MS Gothic", monospace; font-size: 9.5pt; line-height: 1.55; }
  .ln.new { background: var(--soft); border-left: 4px solid #000; margin-left: -6px; padding-left: 6px; }
  .lnote { margin-left: 10px; font-family: "BIZ UDPGothic", "Yu Gothic", "Meiryo", sans-serif; font-size: 8.5pt; font-weight: 700; }
  .blank { display: inline-block; min-width: 3.2em; height: 1.05em; border-bottom: 1.5px solid var(--line); vertical-align: -1px; }
  .blank.s { min-width: 2.2em; }
  .blank.m { min-width: 4.4em; }
  .blank.xl { min-width: 15em; }

  /* 動かして確かめる表 */
  table.try { width: 100%; border-collapse: collapse; margin-top: 3px; text-align: center; }
  table.try th, table.try td { border: 1.5px solid var(--line); padding: 1px 6px; }
  table.try th { background: var(--soft); font-weight: 700; }
  table.try td.chk { width: 18mm; }
  table.ref td, table.ref th { padding: 0 5px; font-size: 8.8pt; }

  .footer { margin-top: 5px; display: flex; justify-content: space-between; align-items: flex-end; gap: 8px; font-size: 9pt; color: var(--sub); }
`;

/* ---- 回転ずしアプリの流れ図の部品とステップ ---- */
// 「A ＋ B → C」の計算の箱（fill = お手本に書いておく文字）
const fAdd = (a, b, c) => ({ k: "proc", s: [{ w: 26, fill: a }, { t: "＋" }, { w: 14, fill: b }, { t: "→" }, { w: 26, fill: c }] });
const fLabelSay = (word, num) => ({ k: "io", s: [{ t: "ことば" }, { w: 76, fill: word }, { t: "を" }, { t: "Label" }, { w: 8, fill: num }, { t: "に表示" }] });

const SUSHI_STEPS = [
  {
    id: "sushi1",
    lead: "回転ずしで食べたお皿の数を、アプリで数えます。",
    main: "100円・120円・150円・200円のお皿のボタンを押すと、<u>枚数・小計・合計</u>が出るお会計アプリを作ろう！",
    flowNote: "100円皿のボタン（CommandButton1）を押したときの流れ。ほかのお皿も、数字と Label の名前をかえて同じ流れ",
    flow: [
      fT("はじめ"),
      fAdd("Mai100", "1", "Mai100"),
      fAdd("Gokei", "100", null),
      fLabelSay("100円皿 Mai100 枚", "1"),
      fLabelSay(null, "5"),
      fLabelSay("合計 Gokei 円", null),
      fT("おわり"),
    ],
    legendExtra: '<div class="lg-n"><b>Sub の外</b>の Dim（枚数の箱）は、流れ図には書かない</div><div class="lg-n"><b>ことば</b>は、たとえば<br>「100円皿 2 枚」のように書く</div>',
    lesson: [
      { n: "1", h: "押すたびに数がふえる箱", p: "Sub の<b>外</b>に Dim を書くと、ボタンを押し終わっても数が消えません。1ずつふやすには「Mai100 = Mai100 + 1」と書きます。", code: "Dim Mai100 As Long\nPrivate Sub CommandButton1_Click()\n    Mai100 = Mai100 + 1\nEnd Sub" },
      { n: "2", h: "合計の箱と、小計を Label に表示する", p: "合計の箱 Gokei には 100 をたします。小計は「枚数 × 1枚の値段」。Label には <b>Label名.Caption = 入れたいもの</b> で入れ、「&」で文字と数をつなげます。", code: 'Gokei = Gokei + 100\nLabel5.Caption = "小計" & Mai100 * 100 & "円"' },
    ],
    note: "1・2で書き方をつかんだら、150円皿・200円皿は下の表を見て自分で書こう（合計はどれも Label9）",
    program: [
      "Dim Mai100 As Long",
      "Dim Mai120 As Long",
      "Dim Mai150 As Long",
      "Dim Mai200 As Long",
      "Dim Gokei As Long",
      "Private Sub CommandButton1_Click()",
      "    Mai100 = Mai100 + 1",
      "    Gokei = Gokei + [[m]]",
      '    Label1.Caption = "100円皿" & Mai100 & "[[s]]"',
      '    Label5.Caption = "小計" & Mai100 * [[m]] & "円"',
      '    Label9.Caption = "合計" & [[m]] & "円"',
      "End Sub",
      "Private Sub CommandButton2_Click()",
      "    [[m]] = Mai120 + 1",
      "    Gokei = Gokei + [[m]]",
      '    Label[[s]].Caption = "120円皿" & [[m]] & "枚"',
      '    Label[[s]].Caption = "小計" & [[m]] * [[m]] & "円"',
      '    Label9.Caption = "合計" & Gokei & "円"',
      "End Sub",
    ],
    after: `<table class="try ref"><tr><th>お皿</th><th>ボタン</th><th>枚数の箱</th><th>1枚の値段</th><th>枚数の Label</th><th>小計の Label</th></tr>
      <tr><td>100円皿</td><td>CommandButton1</td><td>Mai100</td><td>100</td><td>Label1</td><td>Label5</td></tr>
      <tr><td>120円皿</td><td>CommandButton2</td><td>Mai120</td><td>120</td><td>Label2</td><td>Label6</td></tr>
      <tr><td>150円皿</td><td>CommandButton3</td><td>Mai150</td><td>150</td><td>Label3</td><td>Label7</td></tr>
      <tr><td>200円皿</td><td>CommandButton4</td><td>Mai200</td><td>200</td><td>Label4</td><td>Label8</td></tr></table>
`,
    tryHead: ["押すボタン（上から順に）", "枚数の Label", "小計の Label", "Label9（合計）"],
    tryRows: [
      ["100円皿", "Label1：100円皿1枚", "Label5：小計100円", "合計100円"],
      ["100円皿", "Label1：100円皿2枚", "Label5：小計200円", "合計200円"],
      ["120円皿", "Label2：120円皿1枚", "Label6：小計120円", "合計320円"],
      ["200円皿", "Label4：200円皿1枚", "Label8：小計200円", "合計520円"],
    ],
    foot: "※ 枚数が1から数え直しになるときは、Dim を Sub の中に書いていないか見直そう。",
  },
  {
    id: "sushi2",
    lead: "みんなで食べ終わったら、割り勘にしたい！",
    main: "ステップ1のアプリに「割り勘」ボタンを足して、<u>人数と1人あたりの金額</u>を出そう！",
    band: "★ 新しく足すもの：ボタン1つ（CommandButton5）と、Label を2つ（Label10・Label11）",
    flowNote: "割り勘ボタン（CommandButton5）を押したときの流れ。お皿のボタンは、ステップ1のまま",
    flow: [
      fT("はじめ"),
      fIn("人数を入力してください", null),
      fCalc("Gokei", "¥", "Ninzu", null),
      fLabelSay("人数 Ninzu 人", "10"),
      fLabelSay(null, "11"),
      fT("おわり"),
    ],
    legendExtra: '<div class="lg-n"><b>合計</b>の箱 Gokei は、お皿のボタンで作ったものをそのまま使う</div>',
    lesson: [
      { n: "1", h: "フォームに足すもの（ここが新しい！）", p: "<b>CommandButton5</b>＝Caption を「割り勘」にする<br><b>Label10</b>＝人数を出す枠　<b>Label11</b>＝1人あたりの金額を出す枠" },
      { n: "2", h: "合計の箱は、そのまま使える", p: "Gokei は Sub の外で用意してあるので、割り勘ボタンの中でも使えます。" },
      { n: "3", h: "割り勘の計算", p: "割り勘アプリで使った「¥」（整数の割り算）です。", code: "Hitori = Gokei ¥ Ninzu" },
    ],
    note: "お皿の Sub と Dim は、ステップ1のまま。割り勘ボタンの Sub を足そう。",
    program: [
      "（お皿の Dim と、CommandButton1〜4 の Sub は、ステップ1のまま）",
      "★Private Sub CommandButton5_Click()||追加",
      "★    Dim Ninzu As Long",
      "★    Dim Hitori As Long",
      '★    [[m]] = Val(InputBox("[[xl]]"))',
      "★    Hitori = [[m]] [[s]] [[m]]",
      '★    Label10.Caption = "人数" & [[m]] & "人"',
      '★    Label11.Caption = "1人あたり" & [[m]] & "円"',
      "★End Sub",
    ],
    tryHead: ["お皿のボタンを押す", "Label9（合計）", "割り勘で入力する人数", "Label10", "Label11"],
    tryRows: [
      ["100円皿・120円皿・100円皿", "合計320円", "3", "人数3人", "1人あたり106円"],
      ["さらに 200円皿", "合計520円", "4", "人数4人", "1人あたり130円"],
    ],
    foot: "🎉 ここまでできたら、回転ずしアプリの完成です！",
  },
];

/* ======================= 書き出し ======================= */
function build(steps, outName, docTitle) {
  const total = steps.length * 2;
  const html = `<!doctype html>
<html lang="ja">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>${docTitle}</title>
<style>${CSS}</style>
</head>
<body>
${steps.map((st, i) => front(st, i, total) + "\n" + back(st, i, total)).join("\n")}
</body>
</html>
`;
  fs.writeFileSync(path.join(__dirname, outName), html);
  console.log(outName + " を作りました（" + total + " ページ）");
}

build(WARIKAN_STEPS, "warikan-print.html", "割り勘アプリをつくろう（プリント）");
build(SUSHI_STEPS, "sushi-print.html", "回転ずしアプリをつくろう（プリント）");
