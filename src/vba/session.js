/* ===== 実行の管理（session） =====
 * 1回分の実行（フォームを表示してから閉じるまで、またはマクロ1回）をまとめて管理する。
 * DOM に依存しないので、ブラウザと Node の両方で使える。
 *
 * - start(手続き名)  : 手続きを実行する。MsgBox / InputBox のところで止まり、
 *                      { status: "waiting", request } を返す。resume(答え) で続きを実行する
 * - showForm()       : フォームを表示して UserForm_Initialize を実行する
 * - click(部品名)    : ボタンを押したことにして、CommandButton1_Click などを実行する。
 *                      実行した分の step は、タイムラインの「区間（segment）」として後ろに足していく
 * - setText(部品名, 文字) : TextBox に生徒が入力したことを記録する（巻き戻したときに表示を合わせるため）
 * - modelAt(i)       : i 番目の step を実行した後の状態（-1 は実行前）
 *
 * status: "idle"（次の操作待ち）/ "waiting"（ダイアログの答え待ち）/ "running"（時間で区切って実行中）
 *         / "ended"（End・Unload Me でフォームが閉じた）/ "error"
 */
(function (root) {
  "use strict";

  const isNode = typeof module !== "undefined" && module.exports;
  const Ast = isNode ? require("./ast-parser.js") : root.VbaAst;
  const Rt = isNode ? require("./runtime.js") : root.VbaRuntime;

  class ZenkakuError extends Error {
    constructor(list) {
      super("全角の文字があります。");
      this.kind = "zenkaku";
      this.zenkaku = list;
      const first = list[0];
      this.line = first.line;
      this.col = first.col;
      const chars = Array.from(new Set(list.map((z) => (z.ch === "　" ? "全角スペース" : z.ch)))).join(" ");
      this.friendly = `命令の中に、全角の文字（${chars}）が ${list.length} か所あります。`;
      this.hint = "日本語入力をオフにして打ち直すか、「半角に直す」ボタンを押そう（\" \" で囲んだ文字の中はそのままで大丈夫）。";
    }
  }

  // ソースを解析する。文法エラー・全角の文字があれば例外を投げる
  function compile(src) {
    const { ast, zenkaku } = Ast.parseProgram(src);
    if (zenkaku.length) throw new ZenkakuError(zenkaku);
    return ast;
  }

  class Session {
    // opts: { form, files, initialCells, seed, ast }
    constructor(src, opts) {
      opts = opts || {};
      this.ast = opts.ast || compile(src);
      this.rt = new Rt.Runtime(this.ast, opts);
      this.store = new Rt.StepStore(this.rt.initial, this.rt.steps);
      this.status = "idle";
      this.gen = null;
      this.request = null;
      this.error = null;
      this.segments = []; // { label, start, end }（start〜end-1 番目の step がその区間）
      this.timeSliceMs = opts.timeSliceMs || 0; // 0 なら一気に最後まで実行する
    }

    get steps() {
      return this.rt.steps;
    }
    get state() {
      return this.rt.state;
    }

    procNames() {
      return this.ast.procs.map((p) => p.name);
    }
    hasProc(name) {
      return this.rt.procs.has(String(name).toLowerCase());
    }
    // ▶実行で選べるマクロ（引数の無い Sub。イベント用の手続きは除く）
    macroNames() {
      const controls = Object.keys(this.rt.state.controls).map((n) => n.toLowerCase());
      return this.ast.procs
        .filter((p) => p.kind === "Sub" && p.params.length === 0)
        .filter((p) => {
          const k = p.name.toLowerCase();
          if (k.startsWith("userform_")) return false;
          return !controls.some((c) => k.startsWith(c + "_"));
        })
        .map((p) => p.name);
    }

    busy() {
      return this.status === "waiting" || this.status === "running";
    }

    start(procName, label) {
      if (this.busy()) throw new Error("実行中です");
      this.gen = this.rt.runEntry(procName);
      this.segments.push({ label: label || procName, start: this.steps.length, end: this.steps.length });
      this.status = "running";
      return this.resume(undefined);
    }

    // 答え（InputBox なら文字列・キャンセルは null、MsgBox なら押したボタンの値）を渡して続ける
    resume(answer) {
      if (!this.gen) return { status: this.status };
      const t0 = Date.now();
      let input = answer;
      try {
        for (;;) {
          const r = this.gen.next(input);
          input = undefined;
          if (r.done) {
            this.gen = null;
            this.request = null;
            this.closeSegment();
            this.status = r.value && r.value.ended ? "ended" : "idle";
            return { status: this.status };
          }
          const req = r.value;
          if (req.type === "tick") {
            if (this.timeSliceMs && Date.now() - t0 > this.timeSliceMs) {
              this.status = "running";
              this.closeSegment();
              return { status: "running" };
            }
            continue;
          }
          this.request = req;
          this.status = "waiting";
          this.closeSegment();
          return { status: "waiting", request: req };
        }
      } catch (e) {
        this.gen = null;
        this.request = null;
        this.closeSegment();
        if (e && (e.kind === "runtime" || e.kind === "compile" || e.kind === "limit" || e.kind === "syntax")) {
          this.status = "error";
          this.error = e;
          return { status: "error", error: e };
        }
        throw e; // ツール自体の不具合
      }
    }

    closeSegment() {
      const seg = this.segments[this.segments.length - 1];
      if (seg) seg.end = this.steps.length;
    }

    // フォームを表示する（UserForm_Initialize があれば実行する）
    showForm() {
      if (this.busy()) throw new Error("実行中です");
      this.rt.set(["form", "visible"], true);
      this.segments.push({ label: "フォームを表示", start: this.steps.length, end: this.steps.length });
      this.rt.record("form", this.rt.formName, "🪟 フォームを表示", { line: null }, { userAction: true });
      this.closeSegment();
      if (this.hasProc("UserForm_Initialize")) return this.start("UserForm_Initialize", "UserForm_Initialize");
      return { status: "idle" };
    }

    // 生徒がフォームの × を押した
    closeForm() {
      if (this.busy()) return { status: this.status };
      this.rt.set(["form", "visible"], false);
      this.rt.record("form", this.rt.formName, "🪟 × でフォームを閉じた", { line: null }, { userAction: true });
      this.status = "ended";
      return { status: this.status };
    }

    // 生徒が TextBox に文字を入力した
    setText(ctlName, text) {
      const name = this.rt.findControl(ctlName);
      if (!name) return;
      if (this.rt.state.controls[name].Text === text) return;
      this.rt.set(["controls", name, "Text"], String(text));
      this.rt.record("user", name, `⌨️ ${name} に「${text}」と入力`, { line: null }, { userAction: true });
    }

    // 生徒がボタンを押した（イベントの手続きが無ければ何もしない）
    click(ctlName, eventName) {
      const ev = eventName || "Click";
      const name = this.rt.findControl(ctlName) || ctlName;
      const proc = `${name}_${ev}`;
      if (!this.hasProc(proc)) return { status: this.status, noHandler: true };
      return this.start(proc, proc);
    }

    modelAt(index) {
      return this.store.modelAt(index);
    }
    finalModel() {
      return this.store.modelCopy(this.steps.length - 1);
    }
  }

  // ---------- テスト・課題判定用：台本どおりに自動で実行する ----------
  // actions: [{ run: "Macro" } | { showForm: true } | { text: { TextBox1: "1" } } | { click: "CommandButton1" }]
  // 各 action の inputs: InputBox に順に答える文字列（null はキャンセル）
  function runScript(src, opts) {
    opts = opts || {};
    const session = new Session(src, opts);
    const results = [];
    const answerAll = (res, inputs) => {
      while (res.status === "waiting") {
        const q = res.request;
        if (q.type === "inputbox") {
          if (!inputs.length) return Object.assign({}, res, { status: "waiting" });
          res = session.resume(inputs.shift());
        } else {
          res = session.resume(undefined);
        }
      }
      return res;
    };
    for (const act of opts.actions || []) {
      const inputs = (act.inputs || []).slice();
      let res;
      const before = session.steps.length;
      if (act.text) {
        Object.keys(act.text).forEach((k) => session.setText(k, act.text[k]));
        res = { status: session.status };
      }
      if (act.run) res = answerAll(session.start(act.run), inputs);
      else if (act.showForm) res = answerAll(session.showForm(), inputs);
      else if (act.click) res = answerAll(session.click(act.click), inputs);
      const msgsNow = session.state.msgs.slice();
      results.push({
        status: res ? res.status : session.status,
        error: session.error,
        newMsgs: msgsNow.slice(countMsgs(session, before)),
        stepsAfter: session.steps.length, // この操作が終わった時点の状態は session.modelAt(stepsAfter - 1)
      });
      if (session.status === "error" || session.status === "ended") break;
    }
    return { session, results, state: session.state, error: session.error };
  }
  function countMsgs(session, stepIndex) {
    if (stepIndex <= 0) return 0;
    const m = session.modelAt(stepIndex - 1);
    return (m.msgs || []).length;
  }

  const api = { Session, compile, runScript, ZenkakuError };
  if (isNode) module.exports = api;
  else root.VbaSession = api;
})(typeof window !== "undefined" ? window : globalThis);
