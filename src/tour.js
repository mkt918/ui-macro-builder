/* ===== はじめてガイド =====
 * 初めて開いたとき、吹き出しで画面の使い方を順番に案内する。
 * 「❓ 使い方」からもう一度見られる。見終わったら localStorage に印を付ける。
 *   UmbTour.start(force)
 */
(function () {
  "use strict";

  const LS_DONE = "umb_tour_done";

  // target: 吹き出しを付ける場所（CSS セレクタ）。before: その案内を出す前にすること
  const STEPS = [
    {
      target: "#right-panel-form .form-stage",
      title: "ようこそ！ これは「ユーザーフォーム」です",
      text: "フォームは、ボタンやテキストボックスを並べて作る、自分だけの画面です。いまはボタンが1つ置いてあります。",
      before: () => document.getElementById("right-tab-form").click(),
    },
    {
      target: "#form-toolbar",
      title: "① フォームを作る",
      text: "ここのボタンを押すと、フォームに部品が置かれます。部品をクリックすると、下の欄で名前や表示する文字を変えられます。ドラッグで動かすこともできます。",
    },
    {
      target: "#blockly-area",
      title: "② 命令を組む",
      text: "「🖱 CommandButton1 をクリックしたとき」の中に入れたブロックが、フォームのボタンを押したときに動きます。",
      before: () => document.getElementById("tab-blocks").click(),
    },
    {
      target: ".blocklyToolboxDiv",
      title: "ブロックはここから",
      text: "左の なかま をクリックして、ブロックをドラッグしよう。まずは「💬 入出力」の「📢 メッセージを表示」を、クリックしたときの中に入れてみよう。",
    },
    {
      target: "#run-btn",
      title: "③ ▶ 実行",
      text: "▶ 実行 を押すとフォームが開きます。フォームのボタンを押すと、組んだ命令が動きます。",
    },
    {
      target: "#run-dock",
      title: "動きを見よう",
      text: "変数（値を入れる箱）や、いまどこを動かしているかが、ここに出ます。⏱ のスライダーで、前の時点にもどって見返せます。",
    },
    {
      target: "#tab-code",
      title: "本物の VBA も見られる",
      text: "「🤖 VBAコード」を押すと、ブロックが本物の VBA のプログラムになっているのが見られます。ここに直接書いても動きます。困ったら「❓ 使い方」を見てね。",
    },
  ];

  let idx = 0;
  let layer = null;

  function done() {
    try {
      localStorage.setItem(LS_DONE, "1");
    } catch (e) {
      /* 保存できなくても困らない */
    }
    if (layer) layer.remove();
    layer = null;
    window.removeEventListener("resize", place);
  }

  function place() {
    if (!layer) return;
    const step = STEPS[idx];
    const target = document.querySelector(step.target);
    const ring = layer.querySelector(".tour-ring");
    const bubble = layer.querySelector(".tour-bubble");
    if (!target) {
      ring.style.display = "none";
      bubble.style.left = "50%";
      bubble.style.top = "30%";
      bubble.style.transform = "translateX(-50%)";
      return;
    }
    const r = target.getBoundingClientRect();
    ring.style.display = "block";
    ring.style.left = r.left - 6 + "px";
    ring.style.top = r.top - 6 + "px";
    ring.style.width = r.width + 12 + "px";
    ring.style.height = r.height + 12 + "px";
    bubble.style.transform = "";
    const bw = Math.min(360, window.innerWidth - 24);
    bubble.style.width = bw + "px";
    // 下に置けなければ上、左右は画面に収める
    const bh = bubble.offsetHeight || 160;
    let top = r.bottom + 14;
    if (top + bh > window.innerHeight - 8) top = Math.max(8, r.top - bh - 14);
    if (top < 8) top = Math.min(window.innerHeight - bh - 8, r.top + 20);
    let left = r.left + r.width / 2 - bw / 2;
    left = Math.max(12, Math.min(window.innerWidth - bw - 12, left));
    bubble.style.left = left + "px";
    bubble.style.top = top + "px";
  }

  function show() {
    const step = STEPS[idx];
    if (step.before) step.before();
    layer.querySelector(".tour-count").textContent = `${idx + 1} / ${STEPS.length}`;
    layer.querySelector(".tour-title").textContent = step.title;
    layer.querySelector(".tour-text").textContent = step.text;
    layer.querySelector(".tour-prev").disabled = idx === 0;
    layer.querySelector(".tour-next").textContent = idx === STEPS.length - 1 ? "はじめる！" : "次へ ▶";
    setTimeout(place, 60);
  }

  function start(force) {
    let seen = false;
    try {
      seen = localStorage.getItem(LS_DONE) === "1";
    } catch (e) {
      seen = false;
    }
    if (seen && !force) return;
    if (layer) layer.remove();
    idx = 0;
    layer = document.createElement("div");
    layer.className = "tour-layer";
    layer.innerHTML = `
      <div class="tour-ring"></div>
      <div class="tour-bubble" role="dialog" aria-live="polite">
        <div class="tour-top"><span class="tour-count"></span><button class="tour-skip" type="button">✕ とじる</button></div>
        <div class="tour-title"></div>
        <div class="tour-text"></div>
        <div class="tour-actions">
          <button class="tour-prev" type="button">◀ もどる</button>
          <button class="tour-next" type="button">次へ ▶</button>
        </div>
      </div>`;
    document.body.appendChild(layer);
    layer.querySelector(".tour-skip").addEventListener("click", done);
    layer.querySelector(".tour-prev").addEventListener("click", () => {
      if (idx > 0) {
        idx--;
        show();
      }
    });
    layer.querySelector(".tour-next").addEventListener("click", () => {
      if (idx < STEPS.length - 1) {
        idx++;
        show();
      } else done();
    });
    window.addEventListener("resize", place);
    show();
  }

  window.UmbTour = { start };
})();
