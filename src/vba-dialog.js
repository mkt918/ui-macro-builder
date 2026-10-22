/* ===== InputBox / MsgBox のダイアログ =====
 * 教科書に載っている Windows の標準ダイアログと同じ見た目で出す。
 *   VbaDialog.msgBox({ prompt, title, buttons }) → Promise<押したボタンの値>（vbOK=1, vbYes=6 など）
 *   VbaDialog.inputBox({ prompt, title, default }) → Promise<入力した文字（キャンセルは null）>
 */
(function () {
  "use strict";

  const BUTTON_SETS = {
    0: [["OK", 1]],
    1: [["OK", 1], ["キャンセル", 2]],
    2: [["中止(A)", 3], ["再試行(R)", 4], ["無視(I)", 5]],
    3: [["はい(Y)", 6], ["いいえ(N)", 7], ["キャンセル", 2]],
    4: [["はい(Y)", 6], ["いいえ(N)", 7]],
    5: [["再試行(R)", 4], ["キャンセル", 2]],
  };
  const ICONS = { 16: ["✖", "critical"], 32: ["?", "question"], 48: ["!", "exclamation"], 64: ["i", "information"] };

  let openCount = 0;

  function make(tag, cls, text) {
    const el = document.createElement(tag);
    if (cls) el.className = cls;
    if (text !== undefined) el.textContent = text;
    return el;
  }

  // 共通の枠（背景・タイトルバー・×）
  function frame(title) {
    const backdrop = make("div", "vba-dlg-backdrop");
    const box = make("div", "vba-dlg");
    box.setAttribute("role", "dialog");
    box.setAttribute("aria-modal", "true");
    const bar = make("div", "vba-dlg-titlebar");
    bar.appendChild(make("span", "vba-dlg-title", title || "Microsoft Excel"));
    const close = make("button", "vba-dlg-close", "✕");
    close.type = "button";
    close.title = "閉じる";
    bar.appendChild(close);
    box.appendChild(bar);
    backdrop.appendChild(box);
    return { backdrop, box, close };
  }

  function show(backdrop) {
    openCount++;
    document.body.appendChild(backdrop);
  }
  function hide(backdrop) {
    openCount = Math.max(0, openCount - 1);
    backdrop.remove();
  }

  function msgBox(opts) {
    opts = opts || {};
    const buttons = Number(opts.buttons) || 0;
    const set = BUTTON_SETS[buttons & 7] || BUTTON_SETS[0];
    const icon = ICONS[buttons & 0x70];
    return new Promise((resolve) => {
      const { backdrop, box, close } = frame(opts.title);
      const body = make("div", "vba-dlg-body msg");
      if (icon) {
        const ic = make("div", `vba-dlg-icon ${icon[1]}`, icon[0]);
        body.appendChild(ic);
      }
      body.appendChild(make("div", "vba-dlg-prompt", String(opts.prompt === undefined ? "" : opts.prompt).replace(/\r\n|\r/g, "\n")));
      box.appendChild(body);
      const footer = make("div", "vba-dlg-footer");
      const btns = set.map(([label, value], idx) => {
        const b = make("button", "vba-dlg-btn" + (idx === 0 ? " default" : ""), label);
        b.type = "button";
        b.addEventListener("click", () => done(value));
        footer.appendChild(b);
        return b;
      });
      box.appendChild(footer);
      // × と Esc：キャンセルがあればキャンセル、OK だけなら OK。はい／いいえだけのときは閉じられない（Windows と同じ）
      const cancelValue = set.some(([, v]) => v === 2) ? 2 : set.length === 1 ? 1 : null;
      if (cancelValue === null) close.disabled = true;
      close.addEventListener("click", () => cancelValue !== null && done(cancelValue));
      const onKey = (e) => {
        if (e.key === "Escape" && cancelValue !== null) {
          e.preventDefault();
          done(cancelValue);
        } else if (e.key === "Enter") {
          e.preventDefault();
          const focused = btns.indexOf(document.activeElement);
          done(set[focused >= 0 ? focused : 0][1]);
        }
      };
      document.addEventListener("keydown", onKey, true);
      function done(value) {
        document.removeEventListener("keydown", onKey, true);
        hide(backdrop);
        resolve(value);
      }
      show(backdrop);
      btns[0].focus();
    });
  }

  function inputBox(opts) {
    opts = opts || {};
    return new Promise((resolve) => {
      const { backdrop, box, close } = frame(opts.title);
      box.classList.add("input");
      const body = make("div", "vba-dlg-body input");
      body.appendChild(make("div", "vba-dlg-prompt", String(opts.prompt === undefined ? "" : opts.prompt).replace(/\r\n|\r/g, "\n")));
      const side = make("div", "vba-dlg-side");
      const ok = make("button", "vba-dlg-btn default", "OK");
      ok.type = "button";
      const cancel = make("button", "vba-dlg-btn", "キャンセル");
      cancel.type = "button";
      side.appendChild(ok);
      side.appendChild(cancel);
      body.appendChild(side);
      box.appendChild(body);
      const input = make("input", "vba-dlg-input");
      input.type = "text";
      input.value = opts.default === undefined ? "" : String(opts.default);
      input.setAttribute("aria-label", opts.prompt || "入力");
      box.appendChild(input);
      const help = make("div", "vba-dlg-help", "答えを入れて「OK」を押そう（Enter でも OK）");
      box.appendChild(help);
      ok.addEventListener("click", () => done(input.value));
      cancel.addEventListener("click", () => done(null));
      close.addEventListener("click", () => done(null));
      const onKey = (e) => {
        if (e.key === "Escape") {
          e.preventDefault();
          done(null);
        } else if (e.key === "Enter" && !e.isComposing) {
          // 日本語入力の変換確定の Enter では閉じない
          e.preventDefault();
          done(input.value);
        }
      };
      document.addEventListener("keydown", onKey, true);
      function done(value) {
        document.removeEventListener("keydown", onKey, true);
        hide(backdrop);
        resolve(value);
      }
      show(backdrop);
      input.focus();
      input.select();
    });
  }

  window.VbaDialog = {
    msgBox,
    inputBox,
    isOpen: () => openCount > 0,
  };
})();
