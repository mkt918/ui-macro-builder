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
  const top = 5.2;
  const h = top + rows.length * rh + (o.button ? 10 : 2);
  let g = R(x, y, w, h, { fill: "#f0f0f0", sw: 0.5, rx: 0.8 });
  g += R(x, y, w, top, { fill: "#d6d6d6", sw: 0.5, rx: 0.8 });
  g += T(x + 1.8, y + 3.7, "UserForm1", { fs: 2.7 }) + T(x + w - 2.6, y + 4.1, "×", { fs: 3.6, b: 1, a: "middle" });
  rows.forEach((r, i) => {
    const ry = y + top + 1.4 + i * rh;
    if (r.mark) g += R(x + 0.8, ry - 0.9, w - 1.6, rh - 0.4, { fill: "none", dash: 1, sw: 0.55, stroke: "#000", rx: 1 });
    g += T(x + 3, ry + 4.2, r.cap, { fs: 3.1, b: 1 });
    g += R(x + w * 0.43, ry, w * 0.54, 5.6, { sw: 0.5, rx: 0.3 });
    if (r.val) g += T(x + w * 0.7, ry + 4, r.val, { fs: 3.3, a: "middle", b: 1 });
  });
  if (o.button) {
    const by = y + top + rows.length * rh + 2.2;
    g += R(x + w / 2 - 11, by, 22, 6, { fill: "#f7f7f7", sw: 0.45, rx: 1 }) + T(x + w / 2, by + 4.2, "計算する", { fs: 2.9, a: "middle" });
  }
  return { svg: g, h, rowY: (i) => y + top + 1.4 + i * rh + 2.8 };
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
    const f = formG(84, 8, 62, [{ cap: "1人あたり", val: "333円" }, { cap: "あまり", val: "1円" }], { button: true, rh: 8 });
    return svgWrap(
      172,
      42,
      tag(0, 0, 62, "ステップ2：MsgBox（べつの小さな窓）") +
        R(2, 8, 54, 30, { rx: 0.6, sw: 0.5 }) + R(2, 8, 54, 5.5, { fill: "#c9d3e2", rx: 0.6, sw: 0.5 }) +
        T(4, 11.9, "Microsoft Excel", { fs: 2.8 }) + T(53.5, 12.3, "×", { fs: 3.4, b: 1, a: "middle" }) +
        T(29, 24, "1人333円です", { fs: 4.2, a: "middle", b: 1 }) +
        R(22, 30, 14, 5.5, { fill: "#f4f4f4", sw: 0.45, rx: 0.8 }) + T(29, 34, "OK", { fs: 3, a: "middle" }) +
        T(70, 17, "かわる！", { fs: 3.8, a: "middle", b: 1 }) + bigArrow(60, 19, 20, 10) +
        tag(84, 0, 66, "ステップ3：フォームの Label に表示") + f.svg +
        callout(147, f.rowY(0), 150.5, "Label3") + callout(147, f.rowY(1), 150.5, "Label4"),
      "ミッションの絵。MsgBoxの小さな窓に出していた答えを、ユーザーフォームの Label3 と Label4 に出すように変わる"
    );
  },
  warikan4: () => {
    const rowsOld = [{ cap: "1人あたり", val: "763円" }, { cap: "あまり", val: "2円" }];
    const f1 = formG(0, 8, 56, rowsOld, {});
    const f2 = formG(78, 8, 62, rowsOld.concat([{ cap: "幹事", val: "765円", mark: true }, { cap: "ほかの人", val: "763円", mark: true }]), {});
    return svgWrap(
      172,
      46,
      tag(0, 0, 36, "ステップ3のフォーム") + f1.svg +
        T(66, 22, "足す！", { fs: 3.8, a: "middle", b: 1 }) + bigArrow(58, 24, 16, 9) +
        tag(78, 0, 56, "新しいフォーム（点線が追加）") + f2.svg +
        callout(141, f2.rowY(2), 143, "←追加 Label5,6") + callout(141, f2.rowY(3), 143, "←追加 Label7,8"),
      "ミッションの絵。ステップ3のフォームの下に、幹事とほかの人の Label を足したフォーム"
    );
  },
};

/* ======================= 流れ図の部品 ======================= */
// 記号は教科書の描き方：はじめ・おわり＝角丸、入出力＝平行四辺形、計算＝長方形
const FC = { BX: 4, BW: 132, W: 146, FS: 3.1, GAP: 3.2, TH: 7.2, BH: 8.6, SK: 4 };
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
const fIn = (fill) => ({
  k: "io",
  s: [{ t: "ことば" }, { w: 70, fill: fill && fill[0] }, { t: "→変数" }, { w: 20, fill: fill && fill[1] }, { t: "を入力" }],
});
const fMsg = () => ({ k: "io", s: [{ t: "ことば" }, { w: 90 }, { t: "を表示" }] });
const fLabel = (add) => ({ k: "io", add, s: [{ w: 62 }, { t: "を" }, { t: "Label" }, { w: 12 }, { t: "に表示" }] });
const fCalc = (a, op, b, add) => ({
  k: "proc",
  add,
  s: [{ w: 26 }, op ? { t: op } : { w: 10 }, { w: 26 }].concat(b === "sum" ? [{ t: "＋" }, { w: 26 }] : []).concat([{ t: "→" }, { w: 26 }]),
});
const fSum3 = () => ({
  k: "proc",
  s: [{ w: 26 }, { t: "＋" }, { w: 26 }, { t: "＋" }, { w: 26 }, { t: "→" }, { w: 22 }],
});

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
const STEPS = [
  {
    id: "warikan1",
    lead: "文化祭の打ち上げで、ジュース・ピザ・ケーキを注文しました。",
    main: "3つの値段と人数を入力すると、<u>合計</u>と<u>1人あたりの金額</u>を表示するプログラムを作ろう！",
    flow: [fT("はじめ"), fIn(["ジュースの値段を入力してください", "Juice"]), fIn(), fIn(), fIn(), fSum3(), fMsg(), fCalc(0, "÷", 0), fMsg(), fT("おわり")],
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
    flow: [fT("はじめ"), fIn(["合計金額を入力してください", "Gokei"]), fIn(), { k: "proc", s: [{ w: 26 }, { t: "¥" }, { w: 26 }, { t: "→" }, { w: 26 }] }, { k: "proc", s: [{ w: 26 }, { t: "Mod" }, { w: 26 }, { t: "→" }, { w: 26 }] }, fMsg(), fMsg(), fT("おわり")],
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
    lead: "ステップ2では、答えを MsgBox（べつの小さな窓）に出しました。",
    main: "ステップ3では、答えを<u>フォームの Label</u> に出す、本物の「割り勘アプリ」にしよう！",
    band: "MsgBox  →  Label に表示するように、かわるよ！",
    flow: [fT("はじめ"), fIn(), fIn(), fCalc(0, 0, 0), fCalc(0, 0, 0), fLabel(), fLabel(), fT("おわり")],
    legendExtra: '<div class="lg-n"><b>Label</b>に出すときは<br>「○○を Label3 に表示」<br>のように書く</div>',
    lesson: [
      { n: "1", h: "答えを出す場所が、かわる", p: "ステップ2では MsgBox に出しました。ステップ3では、フォームの上の <b>Label</b> に出します。", code: 'ステップ2：MsgBox ("1人" & Hitori & "円です")\nステップ3：Label3.Caption = Hitori & "円"' },
      { n: "2", h: "Label に文字を入れる書き方", p: "<b>Label名.Caption = 入れたいもの</b> と書きます。「円」は & でつなぎます。<br>Label3 は上の枠（1人あたり）、Label4 は下の枠（あまり）です。", code: 'Label4.Caption = Amari & "円"' },
      { n: "3", h: "MsgBox は使わない", p: "MsgBox を書くと、小さな窓が出てしまいます。Dim・InputBox・¥・Mod はステップ2と同じです。" },
    ],
    note: "ステップ2から、いちばん下の2行だけが新しくなったよ。",
    program: [
      "Private Sub CommandButton1_Click()",
      "    Dim Gokei As Long",
      "    Dim Ninzu As Long",
      "    Dim Hitori As Long",
      "    Dim Amari As Long",
      '    Gokei = Val(InputBox("合計金額を入力してください"))',
      '    Ninzu = Val(InputBox("人数を入力してください"))',
      "    Hitori = Gokei ¥ Ninzu",
      "    Amari = Gokei Mod Ninzu",
      '★    Label[[s]].Caption = [[m]] & "円"||MsgBox をやめて Label に！',
      '★    Label[[s]].Caption = [[m]] & "円"||MsgBox をやめて Label に！',
      "End Sub",
    ],
    tryHead: ["合計金額", "人数", "Label3（1人あたり）", "Label4（あまり）"],
    tryRows: [["3000", "3", "1000円", "0円"], ["1000", "3", "333円", "1円"], ["4580", "6", "763円", "2円"], ["500", "7", "71円", "3円"]],
    foot: "※ MsgBox の窓が出たら、MsgBox を消し忘れています。",
  },
  {
    id: "warikan4",
    lead: "あまったお金は、幹事が引き受けます。",
    main: "ステップ3のフォームの下に、<u>「幹事」と「ほかの人」の金額を出す枠（Label）</u>を足そう！",
    band: "★ 新しく足すもの：Label を4つ（Label5・6・7・8）",
    flow: [fT("はじめ"), fSame("ステップ3と同じ：入力2つ → ¥ と Mod の計算"), fCalc(0, 0, 0, true), fSame("ステップ3と同じ：Label3・Label4 に表示"), fLabel(true), fLabel(true), fT("おわり")],
    legendExtra: '<div class="lg-n"><b>灰色の箱</b>…ステップ3と同じ（書かなくてOK）</div><div class="lg-n"><b>追加</b>…ステップ3から<br>新しく足すところ。ここだけ書く</div>',
    lesson: [
      { n: "1", h: "幹事の金額を考える", p: "4580円を6人で分けると、1人763円であまりは2円。ほかの5人が763円ずつ出すなら、幹事は <b>763 + 2 = 765円</b> 出せば、ぴったり4580円になります。<br>たしかめ算：<b>765 + 763 × 5 = 4580</b>" },
      { n: "2", h: "フォームに足す Label（ここが新しい！）", p: "ボタンを下にずらして、あいたところに Label を4つ足します。<br><b>Label5</b>＝見出し（Caption を「幹事」）　<b>Label6</b>＝幹事の金額を出す枠<br><b>Label7</b>＝見出し（Caption を「ほかの人」）　<b>Label8</b>＝ほかの人の金額を出す枠<br><span class=\"note\">Label は作った順に Label5・6・7・8 という名前になります。</span>" },
    ],
    note: "ステップ3の続きに、「追加」の行を足そう。",
    program: [
      "Private Sub CommandButton1_Click()",
      "    Dim Gokei As Long",
      "    Dim Ninzu As Long",
      "    Dim Hitori As Long",
      "    Dim Amari As Long",
      "★    Dim [[m]] As Long||追加（幹事の金額を入れる箱）",
      '    Gokei = Val(InputBox("合計金額を入力してください"))',
      '    Ninzu = Val(InputBox("人数を入力してください"))',
      "    Hitori = Gokei ¥ Ninzu",
      "    Amari = Gokei Mod Ninzu",
      "★    [[m]] = Hitori + [[m]]||追加（幹事 ＝ 1人あたり ＋ あまり）",
      '    Label3.Caption = Hitori & "円"',
      '    Label4.Caption = Amari & "円"',
      '★    Label[[s]].Caption = [[m]] & "円"||追加（幹事）',
      '★    Label[[s]].Caption = [[m]] & "円"||追加（ほかの人）',
      "End Sub",
    ],
    tryHead: ["合計金額", "人数", "Label3（1人あたり）", "Label4（あまり）", "Label6（幹事）", "Label8（ほかの人）"],
    tryRows: [["3000", "3", "1000円", "0円", "1000円", "1000円"], ["1000", "3", "333円", "1円", "334円", "333円"], ["4580", "6", "763円", "2円", "765円", "763円"], ["1100", "3", "366円", "2円", "368円", "366円"]],
    foot: "🎉 ここまでできたら、割り勘アプリの完成です！",
  },
];

/* ======================= ページを組む ======================= */
const nameBox = '<div class="name">　　年<span></span>組<span></span>番　名前<span class="long"></span></div>';

function front(st, idx) {
  const t = taskOf(st.id);
  const kind = t.kind === "guide" ? "解説つき" : "自分で考えよう";
  return `<section class="page">
  <div class="head">
    <div>
      <div class="series">${t.seriesIcon || "🍕"} ${esc(t.series)}　［${kind}］　<b>おもて：考える</b></div>
      <div class="title"><span class="step-tag">${esc(t.step)}</span>${esc(t.title.replace(/^割り勘アプリ\s*(ステップ\d|発展)\s*/, ""))}</div>
    </div>
    ${nameBox}
  </div>
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
  <h2>② 仕様を確かめよう</h2>
  ${specHtml(t.spec)}
  <h2>③ 流れ図を書こう　<span class="h2note">平行四辺形には、<b>画面に出す「ことば」</b>と<b>「変数名」</b>を必ず書く</span></h2>
  <div class="flowrow">${flowchart(st.flow)}${legend(st.legendExtra)}</div>
  <div class="footer"><span>流れ図が書けたら、うら（作る）へ →</span><span>${idx * 2 + 1} / 8</span></div>
</section>`;
}

function back(st, idx) {
  const t = taskOf(st.id);
  return `<section class="page">
  <div class="head">
    <div>
      <div class="series">${t.seriesIcon || "🍕"} ${esc(t.series)}　<b>うら：作る</b></div>
      <div class="title"><span class="step-tag">${esc(t.step)}</span>${esc(t.title.replace(/^割り勘アプリ\s*(ステップ\d|発展)\s*/, ""))}</div>
    </div>
    ${nameBox}
  </div>
  ${road([3, 4, 5])}
  <h2>④ まず読もう</h2>
  ${lessonHtml(st.lesson)}
  <h2>⑤ プログラムを完成させよう　<span class="h2note">${esc(st.note)}</span></h2>
  ${prog(st.program)}
  <h2>⑥ 動かしてたしかめよう</h2>
  ${tryHtml(st.tryHead, st.tryRows)}
  <div class="footer"><span>${esc(st.foot)}</span><span>${idx * 2 + 2} / 8</span></div>
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

  .head { display: flex; justify-content: space-between; align-items: flex-end; gap: 10px; border-bottom: 2.5px solid var(--line); padding-bottom: 3px; }
  .series { font-size: 9pt; color: var(--sub); }
  .title { font-size: 15pt; font-weight: 700; line-height: 1.25; }
  .step-tag { display: inline-block; margin-right: 6px; padding: 0 9px; border: 2px solid var(--line); border-radius: 6px; font-size: 12pt; vertical-align: 1px; }
  .name { font-size: 9.5pt; white-space: nowrap; }
  .name span { display: inline-block; min-width: 20mm; border-bottom: 1px solid var(--line); margin: 0 3px 0 2px; }
  .name span.long { min-width: 36mm; }

  /* 流れ（①〜⑥） */
  ol.road { list-style: none; display: flex; gap: 0; margin: 4px 0 0; padding: 0; font-size: 8.5pt; }
  ol.road li { flex: 1; text-align: center; padding: 1px 2px; border: 1.2px solid #999; color: #777; margin-left: -1px; }
  ol.road li.on { background: #111; color: #fff; border-color: #111; font-weight: 700; }

  h2 { margin: 7px 0 3px; font-size: 10.5pt; padding-left: 7px; border-left: 5px solid var(--line); line-height: 1.3; }
  .h2note { font-size: 8.8pt; font-weight: 400; color: var(--sub); }

  /* ミッション（いちばん目立たせる） */
  .mission2 { margin-top: 6px; border: 3px solid #000; border-radius: 10px; overflow: hidden; }
  .mhead { background: #000; color: #fff; padding: 2px 10px; font-size: 12pt; font-weight: 700; }
  .mhead span { font-size: 8.5pt; font-weight: 400; margin-left: 6px; }
  .mbody { padding: 4px 10px 6px; }
  .mlead { margin: 0; font-size: 10pt; }
  .mmain { margin: 2px 0 3px; font-size: 12.5pt; font-weight: 700; line-height: 1.35; }
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
  table.spec th, table.spec td { border: 1.5px solid var(--line); padding: 2px 7px; vertical-align: top; text-align: left; }
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
  .codebox { margin-top: 3px; border: 1.5px solid var(--line); border-radius: 6px; padding: 3px 10px 4px; font-family: Consolas, "BIZ UDGothic", "MS Gothic", monospace; font-size: 9.5pt; line-height: 1.75; }
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

  .footer { margin-top: 4px; display: flex; justify-content: space-between; font-size: 9pt; color: var(--sub); }
`;

const html = `<!doctype html>
<html lang="ja">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>割り勘アプリをつくろう（プリント）</title>
<style>${CSS}</style>
</head>
<body>
${STEPS.map((st, i) => front(st, i) + "\n" + back(st, i)).join("\n")}
</body>
</html>
`;

fs.writeFileSync(path.join(__dirname, "warikan-print.html"), html);
console.log("warikan-print.html を作りました（" + STEPS.length * 2 + " ページ）");
