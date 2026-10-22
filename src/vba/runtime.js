/* ===== VBA 実行エンジン（runtime） =====
 * AST を実行する。DOM に依存しないので、ブラウザと Node の両方で使える。
 *
 * - ジェネレータ関数で書いてあり、MsgBox / InputBox のところで一時停止（yield）する。
 *   呼び出し側（session.js）が答えを渡して再開する。こうすると「質問 → 表示」の順番が本物どおりになる
 * - 状態（セル・変数・配列・フォームの部品・MsgBox の履歴・ファイル）が変わるたびに step を記録する。
 *   step には差分（ops）だけを入れ、KEYFRAME_EVERY ごとに全体の写し（kf）も入れる（メモリ節約）
 * - 計算の決まりは VBA どおり。ただし演算子の優先順位は教科書の表に合わせる（ast-parser.js 参照）
 * - エラーは VBA と同じ文（message）に加えて、初学者向けのやさしい説明（friendly / hint）を持つ
 */
(function (root) {
  "use strict";

  const MAX_STEPS = 20000; // 記録する step の上限
  const MAX_INSTR = 2000000; // 実行する命令数の上限（無限ループ対策）
  const TICK_EVERY = 5000; // この命令数ごとに一度 yield して、画面が固まらないようにする
  const KEYFRAME_EVERY = 200;
  const WORKBOOK_PATH = "C:\\Users\\student\\Documents";

  const TYPE_JA = {
    Integer: "整数型", Long: "長整数型", Single: "単精度浮動小数点数型", Double: "倍精度浮動小数点数型",
    String: "文字列型", Variant: "バリアント型", Boolean: "ブール型", Date: "日付型", Currency: "通貨型",
    Byte: "バイト型", Object: "オブジェクト型",
  };

  // ---------- エラー ----------
  const ERR_TEXT = {
    5: "プロシージャの呼び出し、または引数が不正です。",
    6: "オーバーフローしました。",
    9: "インデックスが有効範囲にありません。",
    11: "0 で除算しました。",
    13: "型が一致しません。",
    52: "ファイル名または番号が不正です。",
    53: "ファイルが見つかりません。",
    54: "ファイル モードが不正です。",
    55: "ファイルは既に開かれています。",
    62: "ファイルにこれ以上データがありません。",
    424: "オブジェクトが必要です。",
    438: "オブジェクトは、このプロパティまたはメソッドをサポートしていません。",
    1004: "アプリケーション定義またはオブジェクト定義のエラーです。",
  };
  const ERR_FRIENDLY = {
    5: ["関数や命令に渡した値が正しくありません。", "かっこの中に入れた値を確かめよう。"],
    6: ["数が大きすぎて（または小さすぎて）、この型の箱に入りません。", "整数型（Integer）に入るのは -32768〜32767 までです。大きな数を入れるなら Long 型にしよう。"],
    9: ["配列の番号（添字）が、用意した範囲の外です。", "Dim で用意した数（例: Dim Umei(5) なら 0〜5）を確かめよう。"],
    11: ["0 で割り算をしようとしました。", "割る数が 0 にならないか確かめよう。"],
    13: ["数が必要なところに、数ではない文字が入っています（または、その逆です）。", "空っぽの入力や文字を計算に使っていないか確かめよう。文字を数に変えるときは Val を使おう。"],
    52: ["その番号のファイルは開かれていません。", "先に Open ～ As #1 でファイルを開こう。番号が Open と同じか確かめよう。"],
    53: ["そのファイルが見つかりません。", "📁ファイルのタブに、同じ名前のファイルがあるか確かめよう。"],
    54: ["そのファイルは、この使い方では開いていません。", "読むときは For Input、書くときは For Output で開こう。"],
    55: ["その番号のファイルは、もう開いています。", "同じ番号で2回 Open していないか確かめよう。使い終わったら Close しよう。"],
    62: ["ファイルの最後までもう読み終わっています。", "Do While Not EOF(1) のように、EOF で終わりを確かめながら読もう。"],
    424: ["ここには部品やセルなどの「もの」が必要です。", "名前のつづりを確かめよう。"],
    438: ["その部品には、そのプロパティや命令はありません。", "つづりを確かめよう（例: Caption、Text、Value）。"],
    1004: ["セルの指定が正しくありません。", "セルの番地（例: \"A1\"）や、Cells の行・列（1以上）を確かめよう。"],
  };

  class VbaRuntimeError extends Error {
    constructor(code, node, extra) {
      const msg = (extra && extra.message) || `実行時エラー '${code}':\n${ERR_TEXT[code] || "エラーが発生しました。"}`;
      super(msg);
      this.kind = "runtime";
      this.code = code;
      this.line = node ? node.line : null;
      this.col = node ? node.col : null;
      const f = ERR_FRIENDLY[code] || [];
      this.friendly = f[0] || "";
      this.hint = f[1] || "";
      if (extra) {
        if (extra.friendly) this.friendly = extra.friendly;
        if (extra.hint) this.hint = extra.hint;
        if (extra.kind) this.kind = extra.kind;
      }
    }
  }
  function rtError(code, node, extra) {
    return new VbaRuntimeError(code, node, extra);
  }
  function compileError(message, node, friendly, hint) {
    return new VbaRuntimeError(0, node, { message: `コンパイル エラー:\n${message}`, friendly, hint, kind: "compile" });
  }

  class ExitSignal {
    constructor(what) {
      this.what = what;
    }
  }
  class EndSignal {}

  // ---------- 値の変換 ----------
  const OBJ = Symbol("vbaObject");
  function isObj(v) {
    return v !== null && typeof v === "object" && v[OBJ];
  }
  function mkObj(kind, props) {
    return Object.assign({ [OBJ]: true, kind }, props);
  }

  function roundHalfEven(x) {
    const f = Math.floor(x);
    const d = x - f;
    if (d > 0.5) return f + 1;
    if (d < 0.5) return f;
    return f % 2 === 0 ? f : f + 1;
  }
  function roundHalfAway(x, digits) {
    const m = Math.pow(10, digits || 0);
    const v = Math.abs(x) * m;
    const r = Math.round(v + 1e-9) / m;
    return x < 0 ? -r : r;
  }

  function normalizeNumText(s) {
    // 全角の数字・記号を半角に（日本語の VBA は全角数字も数として読む）
    return s.replace(/[０-９．－＋，]/g, (c) => String.fromCharCode(c.charCodeAt(0) - 0xfee0)).replace(/\u3000/g, " ");
  }

  function parseNumber(s) {
    let t = normalizeNumText(String(s)).trim();
    if (t === "") return null;
    if (/^&H[0-9A-F]+$/i.test(t)) return parseInt(t.slice(2), 16);
    if (/^[+-]?\d{1,3}(,\d{3})+(\.\d*)?$/.test(t)) t = t.replace(/,/g, "");
    if (/^[+-]?(\d+\.?\d*|\.\d+)([eEdD][+-]?\d+)?$/.test(t)) return Number(t.replace(/[dD]/, "e"));
    if (/^true$/i.test(t)) return -1;
    if (/^false$/i.test(t)) return 0;
    return null;
  }

  function numToStr(n) {
    if (!isFinite(n)) return String(n);
    if (Number.isInteger(n) && Math.abs(n) < 1e15) return String(n);
    let s = String(Number(n.toPrecision(15)));
    if (/e/.test(s)) {
      s = s.replace(/e([+-])(\d)$/, "E$10$2").replace(/e/, "E");
    }
    return s;
  }

  function toStr(v) {
    if (v === undefined || v === null) return "";
    if (typeof v === "string") return v;
    if (typeof v === "number") return numToStr(v);
    if (typeof v === "boolean") return v ? "True" : "False";
    return String(v);
  }

  function dispVal(v) {
    if (v === undefined) return "（空）";
    if (typeof v === "string") return `"${v}"`;
    return toStr(v);
  }

  const TYPE_DEFAULT = { Integer: 0, Long: 0, Single: 0, Double: 0, Currency: 0, Byte: 0, String: "", Boolean: false, Date: 0, Variant: undefined, Object: undefined };
  function defaultFor(type) {
    return Object.prototype.hasOwnProperty.call(TYPE_DEFAULT, type) ? TYPE_DEFAULT[type] : undefined;
  }

  // ---------- 実行エンジン ----------
  class Runtime {
    constructor(ast, opts) {
      opts = opts || {};
      this.ast = ast;
      this.src = ast.src || "";
      this.opts = opts;
      this.optionExplicit = !!ast.optionExplicit;
      this.optionBase = ast.optionBase || 0;
      this.procs = new Map();
      ast.procs.forEach((p) => this.procs.set(p.name.toLowerCase(), p));
      this.frames = [];
      this.moduleVars = new Map();
      this.instr = 0;
      this.files = {}; // ファイル番号 -> ハンドル
      this.inputFiles = Object.assign({}, opts.files || {}); // 読み込み用の仮想ファイル（名前 -> 本文）
      this.inputCounts = {};
      this.seed = opts.seed !== undefined ? opts.seed >>> 0 : 12345;
      this.unloaded = false;

      const form = opts.form || null;
      const controls = {};
      if (form && form.controls) {
        form.controls.forEach((c) => {
          controls[c.name] = Object.assign({ type: c.type, Visible: true, Enabled: true }, defaultControlProps(c.type), c.props || {});
          if (c.type === "ListBox" && !Array.isArray(controls[c.name].List)) controls[c.name].List = [];
        });
      }
      const initialCells = opts.initialCells ? JSON.parse(JSON.stringify(opts.initialCells)) : {};
      this.state = {
        cells: initialCells,
        sheet: "Sheet1",
        sheets: ["Sheet1"],
        vars: {},
        varTypes: {},
        arrays: {},
        controls,
        form: { name: form ? form.name || "UserForm1" : "UserForm1", caption: form ? form.caption || "UserForm1" : "UserForm1", visible: false },
        msgs: [],
        debug: [],
        files: {},
        proc: null,
      };
      this.formName = this.state.form.name;
      this.steps = [];
      this.pending = [];
      this.recording = false;
      this.initModule();
      this.initial = structuredClone(this.state);
      this.recording = true;
    }

    // ---------- 記録 ----------
    // 記録用の状態（state）は、実行中の変数の実体とオブジェクトを共有しないよう、必ず写しを入れる
    set(path, value) {
      const isObject = value !== null && typeof value === "object";
      setPath(this.state, path, isObject ? structuredClone(value) : value);
      if (this.recording) this.pending.push([path, isObject ? structuredClone(value) : value]);
    }
    record(scope, key, desc, node, extra) {
      if (this.steps.length >= MAX_STEPS) {
        throw rtError(0, node, {
          message: `処理が多すぎます（${MAX_STEPS} ステップ）。`,
          friendly: `処理の回数が多すぎるので、${MAX_STEPS} ステップで止めました（くり返しが終わらない「無限ループ」かもしれません）。`,
          hint: "くり返しの回数を少なくするか、くり返しがいつか終わる条件になっているか確かめよう。",
          kind: "limit",
        });
      }
      const n = node || this.curNode;
      const step = {
        scope,
        key,
        desc,
        line: n ? n.line : null,
        proc: this.state.proc,
        ops: this.pending,
      };
      if (extra) Object.assign(step, extra);
      this.pending = [];
      this.steps.push(step);
      if (this.steps.length % KEYFRAME_EVERY === 0) step.kf = structuredClone(this.state);
      return step;
    }

    // ---------- 変数 ----------
    initModule() {
      const prevRec = this.recording;
      this.recording = false;
      for (const d of this.ast.decls) {
        if (d.type === "Dim") {
          for (const v of d.vars) this.declare(this.moduleVars, v, d, "module");
        } else if (d.type === "Const") {
          for (const it of d.items) {
            const value = runSync(this.evalValue(it.value));
            this.moduleVars.set(it.name.toLowerCase(), { name: it.name, type: it.varType || "Variant", value: it.varType ? coerce(it.varType, value, it.value) : value, isConst: true, scope: "module" });
          }
        }
      }
      this.syncVisible();
      this.recording = prevRec;
    }

    declare(map, v, node, scope) {
      const key = v.name.toLowerCase();
      if (map.has(key)) {
        throw compileError("同じ適用範囲内で宣言が重複しています。", { line: v.line, col: v.col },
          `「${v.name}」という箱（変数）が、すでに用意されています。`,
          "同じ名前の Dim が2つないか確かめよう。");
      }
      const type = v.varType || "Variant";
      let entry;
      if (v.dims) {
        entry = { name: v.name, type, isArray: true, dims: [], data: {}, scope };
        if (v.dims.length) entry.dims = v.dims.map((dm) => this.evalBound(dm, v));
      } else {
        entry = { name: v.name, type, value: defaultFor(type), scope };
      }
      map.set(key, entry);
      return entry;
    }

    evalBound(dm, at) {
      const lo = dm.lower ? toNum(runSync(this.evalValue(dm.lower)), dm.lower) : this.optionBase;
      const hi = toNum(runSync(this.evalValue(dm.upper)), dm.upper);
      if (hi < lo) throw rtError(9, at);
      return [Math.round(lo), Math.round(hi)];
    }

    frame() {
      return this.frames[this.frames.length - 1] || null;
    }

    lookupVar(name) {
      const key = name.toLowerCase();
      const f = this.frame();
      if (f && f.vars.has(key)) return f.vars.get(key);
      if (this.moduleVars.has(key)) return this.moduleVars.get(key);
      // まだ実行していない Dim（後ろの行に書いた Dim）でも、同じ手続きの中なら使える
      if (f && f.proc && f.proc._dimIndex && f.proc._dimIndex.has(key)) {
        const { v, node } = f.proc._dimIndex.get(key);
        return this.declare(f.vars, v, node, "local");
      }
      return null;
    }

    // 実体（real）が、いま画面に見えている名前の一覧（ByRef で別名になっていることがある）
    aliasesOf(real) {
      const names = [];
      const f = this.frame();
      const shadow = new Set();
      if (f) {
        f.vars.forEach((fe, k) => {
          shadow.add(k);
          if (fe === real || fe.ref === real) names.push(fe.name);
        });
      }
      const mk = real.name.toLowerCase();
      if (this.moduleVars.get(mk) === real && !shadow.has(mk)) names.push(real.name);
      return names;
    }
    frameHas(e) {
      const f = this.frame();
      return !!f && f.vars.get(e.name.toLowerCase()) === e;
    }

    syncVisible() {
      const vars = {};
      const varTypes = {};
      const arrays = {};
      const add = (e) => {
        const real = e.ref || e;
        if (real.isArray) {
          arrays[e.name] = { type: real.type, dims: real.dims, data: real.data };
          delete vars[e.name];
        } else if (real.elemRef) {
          vars[e.name] = real.elemRef.get();
          varTypes[e.name] = real.type;
        } else {
          vars[e.name] = real.value;
          varTypes[e.name] = real.isConst ? `${real.type}（定数）` : real.type;
          delete arrays[e.name];
        }
      };
      this.moduleVars.forEach(add);
      const f = this.frame();
      if (f) f.vars.forEach(add);
      this.set(["vars"], vars);
      this.set(["varTypes"], varTypes);
      this.set(["arrays"], arrays);
    }

    readEntry(e) {
      if (e.ref) return this.readEntry(e.ref);
      if (e.elemRef) return e.elemRef.get();
      return e.value;
    }

    writeEntry(e, value, node) {
      if (e.ref) {
        const rv = this.writeEntry(e.ref, value, node);
        if (this.frameHas(e)) this.set(["vars", e.name], rv);
        return rv;
      }
      if (e.isConst) {
        throw compileError("定数には値を代入できません。", node, `「${e.name}」は Const で作った定数なので、値を変えられません。`, "値を変えたいときは Dim で変数を作ろう。");
      }
      if (e.isArray) {
        throw compileError("配列には値を代入できません。", node, `「${e.name}」は配列なので、「${e.name}(1) = …」のように番号を付けて入れよう。`);
      }
      const v = e.type === "Variant" || e.type === "Object" ? derefForVariant(value, node) : coerce(e.type, value, node);
      if (e.elemRef) {
        e.elemRef.set(v, node);
        if (this.frameHas(e)) this.set(["vars", e.name], v);
        return v;
      }
      e.value = v;
      for (const n of this.aliasesOf(e)) this.set(["vars", n], v);
      return v;
    }

    implicitDeclare(name, node) {
      if (this.optionExplicit) throw undefinedVar(name, node);
      const f = this.frame();
      const map = f ? f.vars : this.moduleVars;
      const entry = { name, type: "Variant", value: undefined, scope: f ? "local" : "module", implicit: true };
      map.set(name.toLowerCase(), entry);
      this.set(["vars", name], undefined);
      this.set(["varTypes", name], "Variant");
      return entry;
    }

    // ---------- 配列 ----------
    arrIndexKey(entry, idxValues, node) {
      if (!entry.dims.length) {
        throw rtError(9, node, { friendly: `配列「${entry.name}」の大きさがまだ決まっていません。`, hint: "ReDim で大きさを決めてから使おう。" });
      }
      if (idxValues.length !== entry.dims.length) {
        throw rtError(9, node, { friendly: `配列「${entry.name}」は ${entry.dims.length} 次元なので、番号を ${entry.dims.length} 個書こう。` });
      }
      const parts = idxValues.map((iv, k) => {
        const n = roundHalfEven(toNum(iv, node));
        const [lo, hi] = entry.dims[k];
        if (n < lo || n > hi) {
          throw rtError(9, node, {
            friendly: `配列「${entry.name}」の番号 ${n} は、用意した範囲（${lo}〜${hi}）の外です。`,
            hint: `Dim ${entry.name}(${hi}) のように用意した数を確かめよう。`,
          });
        }
        return n;
      });
      return parts.join(",");
    }
    arrGet(entry, key) {
      return Object.prototype.hasOwnProperty.call(entry.data, key) ? entry.data[key] : defaultFor(entry.type);
    }
    arrSet(entry, key, value, node) {
      const v = entry.type === "Variant" ? derefForVariant(value, node) : coerce(entry.type, value, node);
      entry.data[key] = v;
      for (const n of this.aliasesOf(entry)) this.set(["arrays", n, "data", key], v);
      return v;
    }

    // ---------- 命令数のカウント ----------
    countInstr(node) {
      this.instr++;
      if (this.instr > MAX_INSTR) {
        throw rtError(0, node, {
          message: "処理が終わりません。",
          friendly: "くり返しがいつまでも終わらないようです（無限ループかもしれません）。",
          hint: "くり返しの条件が、いつか変わるようになっているか確かめよう。",
          kind: "limit",
        });
      }
      return this.instr % TICK_EVERY === 0;
    }

    // ---------- 手続きの呼び出し ----------
    *runEntry(name) {
      const proc = this.procs.get(String(name).toLowerCase());
      if (!proc) {
        throw compileError("Sub または Function が定義されていません。", null, `「${name}」という手続きが見つかりません。`);
      }
      try {
        yield* this.callProc(proc, [], null);
      } catch (e) {
        if (e instanceof EndSignal) return { ended: true };
        throw e;
      }
      return { ended: this.unloaded };
    }

    *callProc(proc, argNodes, callNode) {
      if (this.frames.length > 200) {
        throw rtError(28, callNode, { message: "実行時エラー '28':\nスタック領域が不足しています。", friendly: "手続きが自分自身を何度も呼び出して、止まらなくなりました。", hint: "呼び出しが終わる条件があるか確かめよう。" });
      }
      if (!proc._dimIndex) proc._dimIndex = indexDims(proc.body);
      argNodes = argNodes || [];
      const params = proc.params;
      const realArgs = argNodes.filter((a) => a !== null);
      const required = params.filter((p) => !p.optional).length;
      if (argNodes.length > params.length || realArgs.length < required) {
        throw compileError("引数の数が一致していません。", callNode, `「${proc.name}」に渡す値の数が合っていません（${params.length} 個必要です）。`);
      }
      const vars = new Map();
      for (let k = 0; k < params.length; k++) {
        const p = params[k];
        const argNode = argNodes[k] || null;
        const key = p.name.toLowerCase();
        if (!argNode) {
          const dv = p.defaultValue ? yield* this.evalValue(p.defaultValue) : defaultFor(p.varType);
          vars.set(key, { name: p.name, type: p.varType, value: p.varType === "Variant" ? dv : coerce(p.varType, dv, p), scope: "local" });
          continue;
        }
        const ref = p.byRef ? yield* this.refFor(argNode) : null;
        if (ref && ref.entry) {
          if (p.isArray && !ref.entry.isArray && !(ref.entry.ref && ref.entry.ref.isArray)) {
            throw rtError(13, argNode, { friendly: `「${proc.name}」には配列を渡そう。` });
          }
          vars.set(key, { name: p.name, ref: ref.entry.ref || ref.entry, type: p.varType });
        } else if (ref && ref.elem) {
          vars.set(key, { name: p.name, type: p.varType, elemRef: ref.elem });
        } else {
          const v = yield* this.evalValue(argNode);
          vars.set(key, { name: p.name, type: p.varType, value: p.varType === "Variant" ? v : coerce(p.varType, v, argNode), scope: "local" });
        }
      }
      if (proc.kind === "Function") {
        const rt = proc.returnType || "Variant";
        vars.set(proc.name.toLowerCase(), { name: proc.name, type: rt, value: defaultFor(rt), scope: "local", isReturn: true });
      }
      const callerProc = this.state.proc;
      this.frames.push({ proc, vars });
      this.set(["proc"], proc.name);
      this.syncVisible();
      this.record("proc", proc.name, `▶ ${proc.name} を始める`, proc, { procEvent: "start" });
      let result;
      try {
        yield* this.execBody(proc.body);
      } catch (e) {
        if (!(e instanceof ExitSignal && (e.what === "Sub" || e.what === "Function"))) {
          // End・エラーのときは、エラーの場所が分かるようにフレームだけ外して、そのまま外へ伝える
          this.frames.pop();
          throw e;
        }
      }
      if (proc.kind === "Function") result = vars.get(proc.name.toLowerCase()).value;
      // 呼び出し元の変数表示に戻してから「終わる」を記録する（変数の戻しが次の step に混ざらないように）。
      // いちばん外側の手続きが終わったときは、結果を見られるよう、最後の変数・配列の表示を残しておく
      this.frames.pop();
      this.set(["proc"], callerProc);
      if (this.frames.length) this.syncVisible();
      this.curNode = { line: proc.endLine, col: 1 };
      this.record("proc", proc.name, `■ ${proc.name} を終わる`, this.curNode, { procEvent: "end", proc: proc.name });
      return result;
    }

    // ByRef で渡すための参照。変数そのもの・配列の要素なら参照を返し、それ以外（式）は null
    *refFor(node) {
      if (node.type === "Ident") {
        const e = this.lookupVar(node.name);
        if (e) return { entry: e };
        if (!this.isKnownName(node.name)) return { entry: this.implicitDeclare(node.name, node) };
        return null;
      }
      if (node.type === "Call" && node.callee.type === "Ident") {
        const e = this.lookupVar(node.callee.name);
        const real = e && (e.ref || e);
        if (real && real.isArray) {
          const idx = [];
          for (const a of node.args) idx.push(yield* this.evalValue(a));
          const key = this.arrIndexKey(real, idx, node);
          const self = this;
          return {
            elem: {
              get: () => self.arrGet(real, key),
              set: (v, n) => self.arrSet(real, key, v, n),
            },
          };
        }
      }
      return null;
    }

    isKnownName(name) {
      const k = name.toLowerCase();
      return (
        this.procs.has(k) ||
        BUILTIN_CONST.has(k) ||
        BUILTIN_FUNCS.has(k) ||
        OBJECT_NAMES.has(k) ||
        this.findControl(name) !== null ||
        k === "me" ||
        k === this.formName.toLowerCase()
      );
    }

    // ---------- 文の実行 ----------
    *execBody(stmts) {
      for (const s of stmts) yield* this.exec(s);
    }

    *exec(s) {
      this.curNode = s;
      if (this.countInstr(s)) yield { type: "tick" };
      switch (s.type) {
        case "Dim": {
          const f = this.frame();
          for (const v of s.vars) {
            const key = v.name.toLowerCase();
            if (f && f.vars.has(key)) {
              // 後ろの行で先に使われて、すでに用意してあった（Dim の巻き上げ）
              continue;
            }
            const entry = this.declare(f ? f.vars : this.moduleVars, v, s, "local");
            if (entry.isArray) {
              this.set(["arrays", entry.name], { type: entry.type, dims: entry.dims, data: {} });
            } else {
              this.set(["vars", entry.name], entry.value);
              this.set(["varTypes", entry.name], entry.type);
            }
            const shape = entry.isArray && entry.dims.length ? `(${entry.dims.map((d) => `${d[0]}〜${d[1]}`).join(", ")})` : "";
            this.record("var", entry.name, `📦 ${entry.name}${shape}（${TYPE_JA[entry.type] || entry.type}）の箱を用意`, s, { declare: true });
          }
          return;
        }
        case "Const": {
          const f = this.frame();
          for (const it of s.items) {
            const value = yield* this.evalValue(it.value);
            const v = it.varType ? coerce(it.varType, value, it.value) : value;
            (f ? f.vars : this.moduleVars).set(it.name.toLowerCase(), { name: it.name, type: it.varType || "Variant", value: v, isConst: true, scope: "local" });
            this.set(["vars", it.name], v);
            this.set(["varTypes", it.name], `${it.varType || "Variant"}（定数）`);
            this.record("var", it.name, `📌 定数 ${it.name} = ${dispVal(v)}`, s);
          }
          return;
        }
        case "ReDim": {
          for (const v of s.vars) {
            let e = this.lookupVar(v.name);
            if (!e) {
              const f = this.frame();
              e = { name: v.name, type: v.varType || "Variant", isArray: true, dims: [], data: {}, scope: "local" };
              (f ? f.vars : this.moduleVars).set(v.name.toLowerCase(), e);
            }
            const real = e.ref || e;
            if (!real.isArray) {
              real.isArray = true;
              real.data = {};
              delete real.value;
            }
            const dims = [];
            for (const dm of v.dims) {
              const lo = dm.lower ? toNum(yield* this.evalValue(dm.lower), dm.lower) : this.optionBase;
              const hi = toNum(yield* this.evalValue(dm.upper), dm.upper);
              dims.push([Math.round(lo), Math.round(hi)]);
            }
            const old = real.data;
            real.dims = dims;
            real.data = {};
            if (s.preserve) {
              Object.keys(old).forEach((k) => {
                const parts = k.split(",").map(Number);
                if (parts.length === dims.length && parts.every((p, i) => p >= dims[i][0] && p <= dims[i][1])) real.data[k] = old[k];
              });
            }
            for (const n of this.aliasesOf(real)) this.set(["arrays", n], { type: real.type, dims, data: real.data });
            this.record("array", real.name, `📦 配列 ${real.name} の大きさを (${dims.map((d) => `${d[0]}〜${d[1]}`).join(", ")}) にした`, s);
          }
          return;
        }
        case "Assign": {
          // 右辺がセルや部品なら、その値（既定のプロパティ）を入れる。配列はそのまま渡す
          const raw = yield* this.eval(s.value);
          const value = isObj(raw) && raw.kind !== "array" ? derefValue(raw, s.value, this) : raw;
          this.pendingDesc = null;
          yield* this.assign(s.target, value, s);
          const d = this.pendingDesc || { scope: "var", key: "", desc: `${this.text(s)}` };
          this.record(d.scope, d.key, d.desc, s, d.extra);
          return;
        }
        case "CallStmt": {
          this.pendingDesc = null;
          yield* this.callStatement(s);
          if (this.pending.length || this.pendingDesc) {
            const d = this.pendingDesc || { scope: "call", key: "", desc: this.text(s) };
            this.record(d.scope, d.key, d.desc, s, d.extra);
          }
          return;
        }
        case "If": {
          for (const br of s.branches) {
            const c = toBool(yield* this.evalValue(br.cond), br.cond);
            this.record("branch", br.line, `🔀 ${this.text(br.cond)} → ${c ? "はい（YES）" : "いいえ（NO）"}`, { line: br.line }, { branch: c });
            if (c) {
              yield* this.execBody(br.body);
              return;
            }
          }
          if (s.elseBody) yield* this.execBody(s.elseBody);
          return;
        }
        case "Select": {
          const v = yield* this.evalValue(s.expr);
          for (const c of s.cases) {
            let hit = false;
            for (const t of c.tests) {
              if (t.kind === "value") hit = compare(v, yield* this.evalValue(t.expr), t.expr) === 0;
              else if (t.kind === "range") {
                const a = yield* this.evalValue(t.from);
                const b = yield* this.evalValue(t.to);
                hit = compare(v, a, t.from) >= 0 && compare(v, b, t.to) <= 0;
              } else hit = compareOp(t.op, v, yield* this.evalValue(t.expr), t.expr);
              if (hit) break;
            }
            if (hit) {
              this.record("branch", c.line, `🔀 ${this.text(s.expr)}（${dispVal(v)}）は Case ${c.tests.map((t) => this.caseText(t)).join(", ")} に当てはまる`, { line: c.line }, { branch: true });
              yield* this.execBody(c.body);
              return;
            }
          }
          if (s.elseBody) {
            this.record("branch", s.line, `🔀 ${this.text(s.expr)}（${dispVal(v)}）はどの Case にも当てはまらない → Case Else`, s, { branch: false });
            yield* this.execBody(s.elseBody);
          } else {
            this.record("branch", s.line, `🔀 ${this.text(s.expr)}（${dispVal(v)}）はどの Case にも当てはまらない`, s, { branch: false });
          }
          return;
        }
        case "For":
          yield* this.execFor(s);
          return;
        case "Do":
          yield* this.execDo(s);
          return;
        case "Exit":
          this.record("exit", s.what, `↩ Exit ${s.what}（${s.what === "For" ? "For のくり返し" : s.what === "Do" ? "Do のくり返し" : "手続き"}をぬける）`, s);
          throw new ExitSignal(s.what);
        case "End":
          this.set(["form", "visible"], false);
          this.record("end", "End", "⏹ End でプログラムを終わる", s);
          throw new EndSignal();
        case "Open":
          yield* this.execOpen(s);
          return;
        case "Close":
          yield* this.execClose(s);
          return;
        case "InputFile":
          yield* this.execInputFile(s);
          return;
        case "LineInput":
          yield* this.execLineInput(s);
          return;
        case "PrintFile":
        case "WriteFile":
          yield* this.execPrintFile(s);
          return;
        default:
          throw compileError("この命令は使えません。", s, "この命令はこのツールではまだ使えません。");
      }
    }

    caseText(t) {
      if (t.kind === "value") return this.text(t.expr);
      if (t.kind === "range") return `${this.text(t.from)} To ${this.text(t.to)}`;
      return `Is ${t.op} ${this.text(t.expr)}`;
    }

    *execFor(s) {
      let entry = this.lookupVar(s.var.name);
      if (!entry) {
        if (this.findControl(s.var.name)) throw compileError("For のカウンタには変数を使おう。", s.var);
        entry = this.implicitDeclare(s.var.name, s.var);
      }
      const start = toNum(yield* this.evalValue(s.from), s.from);
      const end = toNum(yield* this.evalValue(s.to), s.to);
      const step = s.step ? toNum(yield* this.evalValue(s.step), s.step) : 1;
      const name = entry.name;
      this.writeEntry(entry, start, s.var);
      this.record("var", name, `🔁 For: ${name} = ${toStr(this.readEntry(entry))}`, s, { loop: "start" });
      for (;;) {
        const cur = toNum(this.readEntry(entry), s.var);
        const done = step >= 0 ? cur > end : cur < end;
        if (done) {
          this.record("branch", s.line, `🔁 ${name} = ${toStr(cur)} は ${toStr(end)} を${step >= 0 ? "こえた" : "下回った"}ので、くり返しを終わる`, s, { branch: false, loop: "end" });
          return;
        }
        if (this.countInstr(s)) yield { type: "tick" };
        try {
          yield* this.execBody(s.body);
        } catch (e) {
          if (e instanceof ExitSignal && e.what === "For") return;
          throw e;
        }
        this.curNode = { line: s.endLine, col: 1 };
        const nextVal = toNum(this.readEntry(entry), s.var) + step;
        this.writeEntry(entry, nextVal, s.var);
        this.record("var", name, `🔁 Next: ${name} = ${toStr(this.readEntry(entry))}`, this.curNode, { loop: "next" });
      }
    }

    *execDo(s) {
      const check = function* (cond, line) {
        const c = toBool(yield* this.evalValue(cond.expr), cond.expr);
        const go = cond.kind === "While" ? c : !c;
        this.record("branch", line, `🔁 ${cond.kind} ${this.text(cond.expr)} → ${go ? "くり返す" : "くり返しを終わる"}`, { line }, { branch: go, loop: go ? "continue" : "end" });
        return go;
      }.bind(this);
      for (;;) {
        if (this.countInstr(s)) yield { type: "tick" };
        if (s.pre && !(yield* check(s.pre, s.line))) return;
        try {
          yield* this.execBody(s.body);
        } catch (e) {
          if (e instanceof ExitSignal && e.what === "Do") return;
          throw e;
        }
        if (s.post) {
          this.curNode = { line: s.endLine, col: 1 };
          if (!(yield* check(s.post, s.endLine))) return;
        }
      }
    }

    text(node) {
      return node ? this.src.slice(node.start, node.end) : "";
    }

    // ---------- 代入 ----------
    *assign(target, value, stmt) {
      if (target.type === "Ident") {
        const e = this.lookupVar(target.name);
        if (e) {
          const v = this.writeEntry(e, value, stmt);
          this.pendingDesc = { scope: "var", key: e.name, desc: `「${e.name}」に ${dispVal(v)} を入れた` };
          return;
        }
        const ctl = this.findControl(target.name);
        if (ctl) {
          this.setControlProp(ctl, defaultPropName(this.state.controls[ctl].type), value, stmt);
          return;
        }
        if (this.procs.has(target.name.toLowerCase())) {
          throw compileError("関数呼び出しが代入の左辺にあります。", target, `「${target.name}」は手続きの名前なので、値は入れられません。`);
        }
        const ne = this.implicitDeclare(target.name, target);
        const v = this.writeEntry(ne, value, stmt);
        this.pendingDesc = { scope: "var", key: ne.name, desc: `「${ne.name}」に ${dispVal(v)} を入れた` };
        return;
      }
      if (target.type === "Call") {
        if (target.callee.type === "Ident") {
          const e = this.lookupVar(target.callee.name);
          const real = e && (e.ref || e);
          if (real && real.isArray) {
            const idx = [];
            for (const a of target.args) idx.push(yield* this.evalValue(a));
            const key = this.arrIndexKey(real, idx, target);
            const v = this.arrSet(real, key, value, stmt);
            const shown = this.aliasesOf(real)[0] || real.name;
            this.pendingDesc = { scope: "array", key: `${shown}(${key})`, desc: `「${shown}(${key})」に ${dispVal(v)} を入れた`, extra: { arrayName: shown, index: key } };
            return;
          }
          if (real && !real.isArray && real.type !== "Variant") {
            throw compileError("配列が必要です。", target, `「${real.name}」は配列ではないので、「${real.name}(…)」とは書けません。`, `配列にするなら Dim ${real.name}(5) As ${real.type} のように用意しよう。`);
          }
        }
        const obj = yield* this.eval(target);
        yield* this.setDefault(obj, value, stmt, target);
        return;
      }
      if (target.type === "Member") {
        const obj = yield* this.eval(target.object);
        yield* this.setProp(obj, target.name, value, stmt, target);
        return;
      }
      throw compileError("代入の左側が正しくありません。", target, "「=」の左には、値を入れる先（変数・部品のプロパティ・セル）を書こう。");
    }

    *setDefault(obj, value, stmt, node) {
      if (isObj(obj)) {
        if (obj.kind === "range") return this.setRangeValue(obj, value, stmt);
        if (obj.kind === "control") return this.setControlProp(obj.name, defaultPropName(this.state.controls[obj.name].type), value, stmt);
      }
      throw compileError("代入の左側が正しくありません。", node, "「=」の左には、値を入れる先（変数・部品のプロパティ・セル）を書こう。");
    }

    // ---------- 式の評価 ----------
    // evalValue: 計算に使う値（セル・部品は既定のプロパティの値に変える）
    *evalValue(node) {
      const v = yield* this.eval(node);
      return derefValue(v, node, this);
    }

    *eval(node) {
      switch (node.type) {
        case "Num":
          return node.value;
        case "Str":
          return node.value;
        case "Paren":
          return yield* this.evalValue(node.expr);
        case "Unary": {
          const v = yield* this.evalValue(node.operand);
          if (node.op === "-") return -toNum(v, node.operand);
          if (node.op === "+") return toNum(v, node.operand);
          if (typeof v === "boolean") return !v;
          return ~roundHalfEven(toNum(v, node.operand));
        }
        case "Bin": {
          const a = yield* this.evalValue(node.left);
          const b = yield* this.evalValue(node.right);
          return binop(node.op, a, b, node);
        }
        case "Ident":
          return yield* this.evalIdent(node);
        case "Member": {
          const obj = yield* this.eval(node.object);
          return yield* this.getProp(obj, node.name, node, null);
        }
        case "Call":
          return yield* this.evalCall(node);
        default:
          throw compileError("この式は使えません。", node);
      }
    }

    *evalIdent(node) {
      const name = node.name;
      const k = name.toLowerCase();
      const e = this.lookupVar(name);
      if (e) {
        const real = e.ref || e;
        if (real.isArray) return mkObj("array", { entry: real });
        return this.readEntry(e);
      }
      const ctl = this.findControl(name);
      if (ctl) return mkObj("control", { name: ctl });
      if (k === "me" || k === this.formName.toLowerCase()) return mkObj("form", {});
      const proc = this.procs.get(k);
      if (proc) {
        if (proc.kind !== "Function") {
          throw compileError("関数またはプロシージャではなく、変数が必要です。", node, `「${name}」は Sub なので、値として使えません。`, "値を返すなら Function にしよう。");
        }
        return yield* this.callProc(proc, [], node);
      }
      if (BUILTIN_CONST.has(k)) return BUILTIN_CONST.get(k);
      if (BUILTIN_FUNCS.has(k)) return yield* this.callBuiltin(k, [], node);
      if (OBJECT_NAMES.has(k)) return this.objectByName(k, node);
      if (this.optionExplicit) throw undefinedVar(name, node);
      return undefined; // まだ値を入れていない変数（Empty）
    }

    objectByName(k, node) {
      switch (k) {
        case "worksheets":
        case "sheets":
          return mkObj("sheets", {});
        case "activesheet":
          return mkObj("sheet", { name: this.state.sheet });
        case "thisworkbook":
        case "activeworkbook":
          return mkObj("workbook", {});
        case "application":
          return mkObj("application", {});
        case "worksheetfunction":
          return mkObj("wsf", {});
        case "debug":
          return mkObj("debug", {});
        case "cells":
          return mkObj("cellsAll", {});
        case "selection":
          return mkObj("range", { addrs: ["A1"] });
        case "controls":
          return mkObj("controls", {});
        default:
          throw compileError("この名前は使えません。", node);
      }
    }

    *evalArgs(args) {
      const out = [];
      for (const a of args || []) out.push(a === null ? undefined : yield* this.evalValue(a));
      return out;
    }

    *evalCall(node) {
      const callee = node.callee;
      if (callee.type === "Ident") {
        const name = callee.name;
        const k = name.toLowerCase();
        const e = this.lookupVar(name);
        if (e) {
          const real = e.ref || e;
          if (real.isArray) {
            const idx = yield* this.evalArgs(node.args);
            return this.arrGet(real, this.arrIndexKey(real, idx, node));
          }
          if (real.type === "Variant" && isObj(real.value) && real.value.kind === "array") {
            const idx = yield* this.evalArgs(node.args);
            const ent = real.value.entry;
            return this.arrGet(ent, this.arrIndexKey(ent, idx, node));
          }
          throw compileError("配列が必要です。", node, `「${name}」は配列ではないので、「${name}(…)」とは書けません。`);
        }
        const proc = this.procs.get(k);
        if (proc) {
          if (proc.kind !== "Function") {
            throw compileError("関数またはプロシージャではなく、変数が必要です。", node, `「${name}」は Sub なので、値として使えません。`, "Sub を呼ぶときは「Call " + name + "」と書こう。");
          }
          return yield* this.callProc(proc, node.args, node);
        }
        if (BUILTIN_FUNCS.has(k)) return yield* this.callBuiltin(k, node.args, node);
        if (k === "range") return yield* this.makeRange(node.args, node, null);
        if (k === "cells") return yield* this.makeCells(node.args, node);
        if (k === "rows" || k === "columns") {
          const [n] = yield* this.evalArgs(node.args);
          return mkObj(k === "rows" ? "rows" : "columns", { index: k === "rows" ? roundHalfEven(toNum(n, node)) : colToNum(n, node) });
        }
        if (k === "worksheets" || k === "sheets") {
          const [n] = yield* this.evalArgs(node.args);
          return this.sheetByKey(n, node);
        }
        if (k === "controls") {
          const [n] = yield* this.evalArgs(node.args);
          return this.controlByName(toStr(n), node);
        }
        if (this.findControl(name)) {
          throw rtError(438, node, { friendly: `「${name}」は部品なので、「${name}(…)」とは書けません。`, hint: `「${name}.Caption」のように、プロパティを付けよう。` });
        }
        if (this.optionExplicit) {
          throw compileError("Sub または Function が定義されていません。", callee, `「${name}」という命令（関数）が見つかりません。`, "つづりを確かめよう。");
        }
        throw compileError("Sub または Function が定義されていません。", callee, `「${name}」という命令（関数）が見つかりません。`, "つづりを確かめよう。配列なら、先に Dim で用意しよう。");
      }
      if (callee.type === "Member") {
        const obj = yield* this.eval(callee.object);
        return yield* this.getProp(obj, callee.name, node, node.args);
      }
      // a(1)(2) のような形
      const obj = yield* this.eval(callee);
      if (isObj(obj) && obj.kind === "array") {
        const idx = yield* this.evalArgs(node.args);
        return this.arrGet(obj.entry, this.arrIndexKey(obj.entry, idx, node));
      }
      throw compileError("この書き方は使えません。", node);
    }

    // ---------- 命令としての呼び出し ----------
    *callStatement(s) {
      const callee = s.callee;
      // Call Foo(1, 2) / Foo(1) のように、callee 自体が Call ノードになっている形
      if (callee.type === "Call" && s.args === null) {
        const inner = callee.callee;
        if (inner.type === "Ident") {
          const proc = this.procs.get(inner.name.toLowerCase());
          if (proc) {
            yield* this.callProc(proc, callee.args, s);
            return;
          }
          const k = inner.name.toLowerCase();
          if (BUILTIN_FUNCS.has(k)) {
            yield* this.callBuiltin(k, callee.args, callee);
            return;
          }
          if (k === "unload") {
            this.unloadForm(s);
            return;
          }
        }
        if (inner.type === "Member") {
          const obj = yield* this.eval(inner.object);
          yield* this.invoke(obj, inner.name, callee.args, s);
          return;
        }
        yield* this.eval(callee);
        return;
      }
      const args = s.args || [];
      if (callee.type === "Ident") {
        const k = callee.name.toLowerCase();
        const proc = this.procs.get(k);
        if (proc) {
          yield* this.callProc(proc, args, s);
          return;
        }
        if (k === "unload") {
          this.unloadForm(s);
          return;
        }
        if (k === "randomize") {
          if (args.length && args[0]) this.seed = (toNum(yield* this.evalValue(args[0]), s) * 1000) >>> 0;
          return;
        }
        if (BUILTIN_FUNCS.has(k)) {
          yield* this.callBuiltin(k, args, s);
          return;
        }
        const e = this.lookupVar(callee.name);
        if (e || this.findControl(callee.name)) {
          throw compileError("Sub または Function が定義されていません。", callee, `「${callee.name}」だけでは命令になりません。`, `値を入れるなら「${callee.name} = …」と書こう。`);
        }
        throw compileError("Sub または Function が定義されていません。", callee, `「${callee.name}」という命令が見つかりません。`, "つづりを確かめよう。");
      }
      if (callee.type === "Member") {
        const obj = yield* this.eval(callee.object);
        yield* this.invoke(obj, callee.name, args, s);
        return;
      }
      throw compileError("この行は命令として読めません。", s);
    }

    unloadForm(s) {
      this.unloaded = true;
      this.set(["form", "visible"], false);
      this.pendingDesc = { scope: "form", key: this.formName, desc: "🪟 フォームを閉じる（Unload Me）" };
    }

    // メソッドの呼び出し（obj.Name 引数...）
    *invoke(obj, name, argNodes, s) {
      yield* this.getProp(obj, name, s, argNodes || [], true);
    }

    // ---------- オブジェクトのプロパティ ----------
    // asStatement: 命令として呼ばれた（戻り値を使わない）
    *getProp(obj, name, node, argNodes, asStatement) {
      const k = name.toLowerCase();
      if (!isObj(obj)) {
        throw rtError(424, node, { friendly: `「.${name}」の前にあるものは、部品やセルではありません。`, hint: "名前のつづりを確かめよう。" });
      }
      switch (obj.kind) {
        case "range":
          return yield* this.rangeProp(obj, k, name, node, argNodes, asStatement);
        case "interior":
        case "font":
        case "borders":
          if (argNodes && argNodes.length) throw rtError(438, node);
          return this.formatGet(obj, k, node);
        case "rows":
        case "columns":
          if (k === "clearcontents" || k === "clear" || k === "delete") {
            this.clearLine(obj, node);
            return undefined;
          }
          throw rtError(438, node);
        case "sheets": {
          if (k === "add") {
            let n = this.state.sheets.length + 1;
            while (this.state.sheets.includes("Sheet" + n)) n++;
            const nm = "Sheet" + n;
            this.set(["sheets"], this.state.sheets.concat([nm]));
            this.set(["sheet"], nm);
            this.pendingDesc = { scope: "sheet", key: nm, desc: `シート「${nm}」を追加` };
            return mkObj("sheet", { name: nm });
          }
          if (k === "count") return this.state.sheets.length;
          if (k === "item") {
            const [n] = yield* this.evalArgs(argNodes);
            return this.sheetByKey(n, node);
          }
          throw rtError(438, node);
        }
        case "sheet": {
          if (k === "select" || k === "activate") {
            this.set(["sheet"], obj.name);
            this.pendingDesc = { scope: "sheet", key: obj.name, desc: `シート「${obj.name}」に切り替え` };
            return undefined;
          }
          if (k === "name") return obj.name;
          if (k === "range") return yield* this.makeRange(argNodes, node, null);
          if (k === "cells") return argNodes && argNodes.length ? yield* this.makeCells(argNodes, node) : mkObj("cellsAll", {});
          throw rtError(438, node);
        }
        case "cellsAll":
          if (k === "clearcontents" || k === "clear") {
            Object.keys(this.state.cells).forEach((a) => {
              if (this.state.cells[a].value !== undefined) this.set(["cells", a, "value"], undefined);
            });
            this.pendingDesc = { scope: "cell", key: "A1", desc: "すべてのセルをクリア" };
            return undefined;
          }
          throw rtError(438, node);
        case "workbook":
          if (k === "path") return WORKBOOK_PATH;
          if (k === "name") return "Book1.xlsm";
          throw rtError(438, node);
        case "application":
          if (k === "worksheetfunction") return mkObj("wsf", {});
          if (k === "screenupdating" || k === "displayalerts") return true;
          throw rtError(438, node);
        case "wsf": {
          // Sum(Range("A1:A3")) のようにセル範囲を渡せるよう、引数は値に変えずに渡す
          const args = [];
          for (const a of argNodes || []) args.push(a === null ? undefined : yield* this.eval(a));
          return wsFunction(k, args, node, this);
        }
        case "debug":
          if (k === "print") {
            const args = yield* this.evalArgs(argNodes);
            const text = args.map(toStr).join(" ");
            this.set(["debug", this.state.debug.length], text);
            this.pendingDesc = { scope: "debug", key: "", desc: `🐞 Debug.Print ${text}` };
            return undefined;
          }
          throw rtError(438, node);
        case "control":
          return yield* this.controlProp(obj.name, k, name, node, argNodes, asStatement);
        case "ctlfont": {
          const c = this.state.controls[obj.name];
          if (k === "size") return c.FontSize;
          if (k === "bold") return !!c.FontBold;
          if (k === "italic") return !!c.FontItalic;
          if (k === "name") return c.FontName || "MS UI Gothic";
          throw rtError(438, node);
        }
        case "form": {
          const ctl = this.findControl(name);
          if (ctl) {
            const ref = mkObj("control", { name: ctl });
            if (argNodes && argNodes.length) throw rtError(438, node);
            return ref;
          }
          if (k === "controls") {
            if (argNodes && argNodes.length) {
              const [n] = yield* this.evalArgs(argNodes);
              return this.controlByName(toStr(n), node);
            }
            return mkObj("controls", {});
          }
          if (k === "caption") return this.state.form.caption;
          if (k === "show") {
            this.set(["form", "visible"], true);
            this.pendingDesc = { scope: "form", key: this.formName, desc: "🪟 フォームを表示" };
            const init = this.procs.get("userform_initialize");
            if (init) yield* this.callProc(init, [], node);
            return undefined;
          }
          if (k === "hide") {
            this.set(["form", "visible"], false);
            this.pendingDesc = { scope: "form", key: this.formName, desc: "🪟 フォームを隠す" };
            return undefined;
          }
          throw rtError(438, node, { friendly: `フォームに「${name}」という部品やプロパティはありません。`, hint: "部品の名前のつづりを確かめよう。" });
        }
        case "controls": {
          if (k === "item") {
            const [n] = yield* this.evalArgs(argNodes);
            return this.controlByName(toStr(n), node);
          }
          if (k === "count") return Object.keys(this.state.controls).length;
          throw rtError(438, node);
        }
        case "array":
          throw compileError("配列にはプロパティがありません。", node);
        default:
          throw rtError(438, node);
      }
    }

    // obj.Name = value
    *setProp(obj, name, value, stmt, node) {
      const k = name.toLowerCase();
      if (!isObj(obj)) {
        throw rtError(424, node, { friendly: `「.${name}」の前にあるものは、部品やセルではありません。`, hint: "名前のつづりを確かめよう。" });
      }
      switch (obj.kind) {
        case "range":
          if (k === "value" || k === "value2" || k === "formula" || k === "text") return this.setRangeValue(obj, value, stmt);
          throw rtError(438, node, { friendly: `セルには「${name}」に値を入れることはできません。`, hint: "セルの値は .Value に入れよう。" });
        case "interior":
        case "font":
        case "borders":
          return this.formatSet(obj, k, value, stmt, node);
        case "sheet": {
          if (k === "name") {
            const nm = toStr(derefValue(value, node, this));
            const sheets = this.state.sheets.map((s) => (s === obj.name ? nm : s));
            this.set(["sheets"], sheets);
            if (this.state.sheet === obj.name) this.set(["sheet"], nm);
            // Worksheets.Add.Name = "x" は「シートを追加」としてまとめて説明する
            const added = this.pendingDesc && this.pendingDesc.scope === "sheet" && this.pendingDesc.key === obj.name;
            this.pendingDesc = { scope: "sheet", key: nm, desc: added ? `シート「${nm}」を追加` : `シートの名前を「${nm}」にした` };
            return;
          }
          throw rtError(438, node);
        }
        case "application":
          return; // ScreenUpdating などは無視する
        case "control": {
          const c = this.state.controls[obj.name];
          const prop = canonicalProp(c.type, k);
          if (!prop) throw unknownProp(obj.name, name, node);
          return this.setControlProp(obj.name, prop, value, stmt);
        }
        case "ctlfont": {
          const map = { size: "FontSize", bold: "FontBold", italic: "FontItalic", name: "FontName" };
          if (!map[k]) throw rtError(438, node);
          return this.setControlProp(obj.name, map[k], value, stmt);
        }
        case "form":
          if (k === "caption") {
            const v = toStr(derefValue(value, node, this));
            this.set(["form", "caption"], v);
            this.pendingDesc = { scope: "form", key: this.formName, desc: `🪟 フォームのタイトルを「${v}」にした` };
            return;
          }
          throw rtError(438, node);
        default:
          throw rtError(438, node);
      }
    }

    // ---------- フォームの部品 ----------
    findControl(name) {
      const k = String(name).toLowerCase();
      for (const n of Object.keys(this.state.controls)) if (n.toLowerCase() === k) return n;
      return null;
    }
    controlByName(name, node) {
      const n = this.findControl(name);
      if (!n) {
        throw rtError(13, node, { message: "実行時エラー '-2147024809':\n指定されたオブジェクトが見つかりません。", friendly: `「${name}」という名前の部品は、フォームにありません。`, hint: "部品の名前（プロパティの Name）を確かめよう。" });
      }
      return mkObj("control", { name: n });
    }

    *controlProp(ctlName, k, rawName, node, argNodes, asStatement) {
      const c = this.state.controls[ctlName];
      if (c.type === "ListBox") {
        if (k === "additem") {
          const [v] = yield* this.evalArgs(argNodes);
          const list = (c.List || []).concat([toStr(v)]);
          this.set(["controls", ctlName, "List"], list);
          this.pendingDesc = { scope: "control", key: ctlName, desc: `${ctlName} に「${toStr(v)}」を追加` };
          return undefined;
        }
        if (k === "clear") {
          this.set(["controls", ctlName, "List"], []);
          this.pendingDesc = { scope: "control", key: ctlName, desc: `${ctlName} を空にした` };
          return undefined;
        }
        if (k === "listcount") return (c.List || []).length;
        if (k === "list") {
          const args = yield* this.evalArgs(argNodes);
          if (!args.length) return mkObj("array", { entry: { name: "List", dims: [[0, (c.List || []).length - 1]], data: Object.assign({}, c.List || []), type: "Variant" } });
          return (c.List || [])[roundHalfEven(toNum(args[0], node))];
        }
      }
      if (k === "setfocus") return undefined;
      if (k === "font") return mkObj("ctlfont", { name: ctlName });
      if (argNodes && argNodes.length) throw rtError(438, node);
      const prop = canonicalProp(c.type, k);
      if (!prop) throw unknownProp(ctlName, rawName, node);
      if (asStatement) throw compileError("Sub または Function が定義されていません。", node, `「${ctlName}.${rawName}」だけでは命令になりません。`, `値を入れるなら「${ctlName}.${rawName} = …」と書こう。`);
      if (prop === "Value" && c.type === "TextBox") return c.Text;
      if (prop === "Value" && c.type === "ListBox") return c.ListIndex >= 0 ? (c.List || [])[c.ListIndex] : undefined;
      return c[prop];
    }

    setControlProp(ctlName, prop, value, stmt) {
      const c = this.state.controls[ctlName];
      let v = derefValue(value, stmt, this);
      if (prop === "Value" && c.type === "TextBox") prop = "Text";
      if (["Caption", "Text", "FontName"].includes(prop)) v = toStr(v);
      else if (["Visible", "Enabled", "FontBold", "FontItalic", "Locked"].includes(prop)) v = toBool(v, stmt);
      else if (["FontSize", "Left", "Top", "Width", "Height", "ForeColor", "BackColor", "BorderStyle", "SpecialEffect", "TextAlign", "ListIndex"].includes(prop)) v = toNum(v, stmt);
      this.set(["controls", ctlName, prop], v);
      let desc;
      if (prop === "Caption") desc = `${ctlName} に「${v}」を表示`;
      else if (prop === "Text") desc = `${ctlName} の文字を「${v}」にした`;
      else desc = `${ctlName}.${prop} を ${dispVal(v)} にした`;
      this.pendingDesc = { scope: "control", key: ctlName, prop, desc };
    }

    // ---------- セル ----------
    *makeRange(argNodes, node) {
      const args = [];
      for (const a of argNodes || []) args.push(yield* this.eval(a));
      if (args.length === 1) {
        const v = derefValue(args[0], node, this);
        return mkObj("range", { addrs: parseRangeAddr(toStr(v), node) });
      }
      if (args.length === 2 && isObj(args[0]) && isObj(args[1]) && args[0].kind === "range" && args[1].kind === "range") {
        const a = parseAddr(args[0].addrs[0]);
        const b = parseAddr(args[1].addrs[0]);
        return mkObj("range", { addrs: rectAddrs(a, b) });
      }
      if (args.length === 2) {
        const a = parseAddr(toStr(derefValue(args[0], node, this)));
        const b = parseAddr(toStr(derefValue(args[1], node, this)));
        if (!a || !b) throw rtError(1004, node);
        return mkObj("range", { addrs: rectAddrs(a, b) });
      }
      throw rtError(1004, node);
    }

    *makeCells(argNodes, node) {
      const [r, c] = yield* this.evalArgs(argNodes);
      const row = roundHalfEven(toNum(r, node));
      const col = colToNum(c, node);
      if (row < 1 || col < 1) {
        throw rtError(1004, node, { friendly: `Cells(${row}, ${col}) というセルはありません。`, hint: "Cells(行, 列) の行と列は、どちらも 1 以上にしよう。" });
      }
      return mkObj("range", { addrs: [colLetters(col) + row] });
    }

    cellGet(addr) {
      const c = this.state.cells[addr];
      return c && c.value !== undefined && c.value !== "" ? c.value : undefined;
    }

    setRangeValue(rng, value, stmt) {
      let v = derefValue(value, stmt, this);
      if (typeof v === "string") {
        const n = v.trim() === "" ? null : parseNumber(v);
        if (n !== null && !/^&H/i.test(v.trim())) v = n;
      }
      if (typeof v === "boolean") v = v ? "TRUE" : "FALSE";
      for (const a of rng.addrs) {
        if (!this.state.cells[a]) this.set(["cells", a], {});
        this.set(["cells", a, "value"], v === undefined ? undefined : v);
      }
      const a0 = rng.addrs[0];
      this.pendingDesc = { scope: "cell", key: a0, desc: `${rangeLabel(rng)} に「${toStr(v)}」を入力` };
    }

    *rangeProp(rng, k, rawName, node, argNodes, asStatement) {
      const a0 = rng.addrs[0];
      switch (k) {
        case "value":
        case "value2":
        case "formula":
          return this.cellGet(a0);
        case "text":
          return toStr(this.cellGet(a0));
        case "interior":
          return mkObj("interior", { range: rng });
        case "font":
          return mkObj("font", { range: rng });
        case "borders":
          return mkObj("borders", { range: rng });
        case "row":
          return parseAddr(a0).row;
        case "column":
          return parseAddr(a0).col;
        case "address":
          return "$" + a0.replace(/(\d+)/, "$$$1");
        case "count":
          return rng.addrs.length;
        case "offset": {
          const [r, c] = yield* this.evalArgs(argNodes);
          const p = parseAddr(a0);
          const row = p.row + roundHalfEven(toNum(r || 0, node));
          const col = p.col + roundHalfEven(toNum(c || 0, node));
          if (row < 1 || col < 1) throw rtError(1004, node);
          return mkObj("range", { addrs: [colLetters(col) + row] });
        }
        case "clearcontents":
        case "clear":
          for (const a of rng.addrs) {
            if (this.state.cells[a]) {
              if (k === "clear") this.set(["cells", a], undefined);
              else this.set(["cells", a, "value"], undefined);
            }
          }
          this.pendingDesc = { scope: "cell", key: a0, desc: `${rangeLabel(rng)} をクリア` };
          return undefined;
        case "copy": {
          const args = [];
          for (const a of argNodes || []) args.push(yield* this.eval(a));
          const dest = args[0];
          if (!isObj(dest) || dest.kind !== "range") return undefined;
          const src0 = parseAddr(a0);
          const dst0 = parseAddr(dest.addrs[0]);
          for (const a of rng.addrs) {
            const p = parseAddr(a);
            const to = colLetters(dst0.col + (p.col - src0.col)) + (dst0.row + (p.row - src0.row));
            if (!this.state.cells[to]) this.set(["cells", to], {});
            this.set(["cells", to, "value"], this.cellGet(a));
          }
          this.pendingDesc = { scope: "cell", key: dest.addrs[0], desc: `${rangeLabel(rng)} を ${dest.addrs[0]} にコピー` };
          return undefined;
        }
        case "select":
        case "activate":
          this.pendingDesc = { scope: "cell", key: a0, desc: `${rangeLabel(rng)} を選択` };
          return undefined;
        default:
          throw rtError(438, node, { friendly: `セルに「${rawName}」というプロパティはありません。`, hint: "セルの値は .Value で読み書きしよう。" });
      }
    }

    formatGet(obj, k, node) {
      const c = this.state.cells[obj.range.addrs[0]] || {};
      if (obj.kind === "interior" && k === "color") return c.bg ? hexToVbaColor(c.bg) : 16777215;
      if (obj.kind === "font" && k === "bold") return !!c.bold;
      if (obj.kind === "font" && k === "size") return c.fontSize || 11;
      if (obj.kind === "font" && k === "color") return c.color ? hexToVbaColor(c.color) : 0;
      if (obj.kind === "borders" && k === "linestyle") return c.border ? 1 : -4142;
      throw rtError(438, node);
    }

    formatSet(obj, k, value, stmt, node) {
      const v = derefValue(value, node, this);
      const rng = obj.range;
      let field;
      let fv;
      let what;
      if (obj.kind === "interior" && (k === "color" || k === "colorindex")) {
        field = "bg";
        fv = k === "color" ? vbaColorToHex(toNum(v, node)) : colorIndexToHex(toNum(v, node));
        what = "背景色を変更";
      } else if (obj.kind === "font" && k === "bold") {
        field = "bold";
        fv = toBool(v, node) || undefined;
        what = fv ? "太字に" : "太字をやめた";
      } else if (obj.kind === "font" && k === "size") {
        field = "fontSize";
        fv = toNum(v, node);
        what = `文字サイズを ${fv} に`;
      } else if (obj.kind === "font" && k === "color") {
        field = "color";
        fv = vbaColorToHex(toNum(v, node));
        what = "文字の色を変更";
      } else if (obj.kind === "borders" && (k === "linestyle" || k === "weight")) {
        field = "border";
        const n = toNum(v, node);
        fv = n === -4142 || n === 0 ? undefined : true;
        what = fv ? "罫線を引いた" : "罫線を消した";
      } else {
        throw rtError(438, node);
      }
      for (const a of rng.addrs) {
        if (!this.state.cells[a]) this.set(["cells", a], {});
        this.set(["cells", a, field], fv);
      }
      this.pendingDesc = { scope: "cell", key: rng.addrs[0], desc: `${rangeLabel(rng)} の${what}` };
    }

    clearLine(obj, node) {
      const isRow = obj.kind === "rows";
      Object.keys(this.state.cells).forEach((a) => {
        const p = parseAddr(a);
        if (p && (isRow ? p.row === obj.index : p.col === obj.index) && this.state.cells[a].value !== undefined) {
          this.set(["cells", a, "value"], undefined);
        }
      });
      const key = isRow ? "A" + obj.index : colLetters(obj.index) + "1";
      this.pendingDesc = { scope: "cell", key, desc: isRow ? `${obj.index} 行目をクリア` : `${colLetters(obj.index)} 列をクリア` };
    }

    sheetByKey(n, node) {
      if (typeof n === "number") {
        const name = this.state.sheets[roundHalfEven(n) - 1];
        if (!name) throw rtError(9, node, { friendly: `${n} 番目のシートはありません。` });
        return mkObj("sheet", { name });
      }
      const s = toStr(n);
      const found = this.state.sheets.find((x) => x.toLowerCase() === s.toLowerCase());
      if (!found) throw rtError(9, node, { friendly: `「${s}」という名前のシートはありません。`, hint: "シートの名前を確かめよう。" });
      return mkObj("sheet", { name: found });
    }

    // ---------- 組み込み関数 ----------
    *callBuiltin(k, argNodes, node) {
      if (k === "msgbox") return yield* this.doMsgBox(argNodes, node);
      if (k === "inputbox") return yield* this.doInputBox(argNodes, node);
      if (k === "ubound" || k === "lbound") {
        const arrV = argNodes[0] ? yield* this.eval(argNodes[0]) : undefined;
        const dimN = argNodes[1] ? toNum(yield* this.evalValue(argNodes[1]), node) : 1;
        if (!isObj(arrV) || arrV.kind !== "array") throw rtError(13, node, { friendly: `${k === "ubound" ? "UBound" : "LBound"} には配列の名前を書こう。` });
        const d = arrV.entry.dims[dimN - 1];
        if (!d) throw rtError(9, node);
        return k === "ubound" ? d[1] : d[0];
      }
      if (k === "isarray") {
        const v = argNodes[0] ? yield* this.eval(argNodes[0]) : undefined;
        return isObj(v) && v.kind === "array";
      }
      if (k === "eof") {
        const [n] = yield* this.evalArgs(argNodes);
        const h = this.fileHandle(n, node);
        if (h.mode !== "Input") throw rtError(54, node);
        return atEOF(h);
      }
      if (k === "freefile") {
        let n = 1;
        while (this.files[n]) n++;
        return n;
      }
      if (k === "rnd") {
        this.seed = (Math.imul(this.seed, 1103515245) + 12345) >>> 0;
        return (this.seed >>> 8) / 16777216;
      }
      if (k === "iif") {
        const [c, a, b] = yield* this.evalArgs(argNodes);
        return toBool(c, node) ? a : b;
      }
      const args = yield* this.evalArgs(argNodes);
      return builtinFunc(k, args, node);
    }

    *doMsgBox(argNodes, node) {
      const [prompt, buttons, title] = yield* this.evalArgs(argNodes);
      const text = toStr(prompt);
      const btn = buttons === undefined ? 0 : toNum(buttons, node);
      this.set(["msgs", this.state.msgs.length], text);
      this.record("msg", text, `📢 「${text}」を表示`, node, { title: title === undefined ? "Microsoft Excel" : toStr(title), buttons: btn });
      const answer = yield { type: "msgbox", prompt: text, title: title === undefined ? "Microsoft Excel" : toStr(title), buttons: btn, line: node.line };
      if (answer !== undefined && answer !== null) return answer;
      const kind = btn & 7;
      if (kind === 4 || kind === 3) return 6; // vbYes
      return 1; // vbOK
    }

    *doInputBox(argNodes, node) {
      const [prompt, title, def] = yield* this.evalArgs(argNodes);
      const base = `${node.line}:${node.col}`;
      const n = this.inputCounts[base] || 0;
      this.inputCounts[base] = n + 1;
      const key = `${base}#${n}`;
      const req = {
        type: "inputbox",
        prompt: toStr(prompt),
        title: title === undefined ? "Microsoft Excel" : toStr(title),
        default: def === undefined ? "" : toStr(def),
        key,
        line: node.line,
      };
      const answer = yield req;
      const text = answer === null || answer === undefined ? "" : String(answer);
      this.record("input", req.prompt, answer === null ? `⌨️ 「${req.prompt}」でキャンセル（"" になる）` : `⌨️ 「${req.prompt}」に「${text}」と入力`, node, { answer: text });
      return text;
    }

    // ---------- ファイル ----------
    fileHandle(n, node) {
      const num = roundHalfEven(toNum(n, node));
      const h = this.files[num];
      if (!h) throw rtError(52, node, { friendly: `#${num} の番号のファイルは開かれていません。` });
      return h;
    }

    *execOpen(s) {
      const path = toStr(yield* this.evalValue(s.path));
      const num = roundHalfEven(toNum(yield* this.evalValue(s.fileNum), s.fileNum));
      if (num < 1 || num > 511) throw rtError(52, s);
      if (this.files[num]) throw rtError(55, s);
      const base = path.split(/[\\/¥]/).pop();
      if (s.mode === "Input") {
        const found = Object.keys(this.inputFiles).find((f) => f.toLowerCase() === base.toLowerCase());
        const outFound = Object.keys(this.state.files).find((f) => f.toLowerCase() === base.toLowerCase());
        if (!found && !outFound) {
          throw rtError(53, s, { friendly: `「${base}」というファイルが見つかりません。`, hint: "📁ファイルのタブに、同じ名前のファイルがあるか確かめよう。" });
        }
        const text = found ? this.inputFiles[found] : this.state.files[outFound];
        this.files[num] = { name: found || outFound, mode: "Input", text: String(text).replace(/^\uFEFF/, "").replace(/\r\n?/g, "\n"), pos: 0 };
      } else {
        const prev = s.mode === "Append" ? this.state.files[base] || this.inputFiles[base] || "" : "";
        this.files[num] = { name: base, mode: s.mode, out: prev };
      }
      this.record("file", base, `📂 ${base} を開く（${s.mode === "Input" ? "読む" : "書く"}・#${num}）`, s);
    }

    *execClose(s) {
      const nums = [];
      if (!s.fileNums.length) nums.push(...Object.keys(this.files).map(Number));
      else for (const f of s.fileNums) nums.push(roundHalfEven(toNum(yield* this.evalValue(f), f)));
      const names = [];
      for (const n of nums) {
        const h = this.files[n];
        if (!h) continue;
        if (h.mode !== "Input") this.set(["files", h.name], h.out);
        names.push(h.name);
        delete this.files[n];
      }
      this.record("file", names.join(", "), `📁 ${names.join(", ") || "ファイル"} を閉じる`, s);
    }

    *execInputFile(s) {
      const h = this.fileHandle(yield* this.evalValue(s.fileNum), s);
      if (h.mode !== "Input") throw rtError(54, s);
      const got = [];
      const names = [];
      this.pendingDesc = null;
      for (const t of s.targets) {
        if (atEOF(h)) throw rtError(62, s);
        const field = readField(h);
        yield* this.assign(t, field, s);
        got.push(`${this.text(t)} ← ${dispVal(field)}`);
        if (t.type === "Ident") names.push((this.lookupVar(t.name) || { name: t.name }).name);
      }
      this.record("file", h.name, `📄 ${h.name} から読みこむ: ${got.join("、")}`, s, { names });
    }

    *execLineInput(s) {
      const h = this.fileHandle(yield* this.evalValue(s.fileNum), s);
      if (h.mode !== "Input") throw rtError(54, s);
      if (h.pos >= h.text.length) throw rtError(62, s);
      let end = h.text.indexOf("\n", h.pos);
      if (end < 0) end = h.text.length;
      const lineText = h.text.slice(h.pos, end);
      h.pos = end + 1;
      yield* this.assign(s.target, lineText, s);
      this.record("file", h.name, `📄 ${h.name} から1行読みこむ: ${dispVal(lineText)}`, s);
    }

    *execPrintFile(s) {
      const h = this.fileHandle(yield* this.evalValue(s.fileNum), s);
      if (h.mode === "Input") throw rtError(54, s);
      let out = "";
      let trailing = false;
      for (let k = 0; k < s.items.length; k++) {
        const it = s.items[k];
        const v = yield* this.evalValue(it.expr);
        if (s.type === "WriteFile") {
          out += typeof v === "string" ? `"${v}"` : toStr(v);
          if (k < s.items.length - 1) out += ",";
        } else {
          out += toStr(v);
          if (it.sep === ",") out += "\t";
        }
        trailing = !!it.sep;
      }
      if (!trailing || s.type === "WriteFile") out += "\r\n";
      h.out += out;
      this.record("file", h.name, `✏️ ${h.name} に書きこむ: ${out.replace(/\r\n$/, "")}`, s);
    }
  }

  // ---------- 補助関数 ----------
  function setPath(obj, path, value) {
    let o = obj;
    for (let i = 0; i < path.length - 1; i++) {
      const k = path[i];
      if (o[k] === undefined || o[k] === null || typeof o[k] !== "object") o[k] = {};
      o = o[k];
    }
    const last = path[path.length - 1];
    if (path.length === 0) return;
    if (value === undefined) delete o[last];
    else o[last] = value;
  }

  // ジェネレータを最後まで動かす（一時停止しない式の評価用）
  function runSync(gen) {
    let r = gen.next();
    while (!r.done) {
      if (r.value && r.value.type !== "tick") {
        throw compileError("ここでは MsgBox や InputBox は使えません。", null, "宣言（Dim・Const）の中では MsgBox や InputBox は使えません。");
      }
      r = gen.next();
    }
    return r.value;
  }

  function indexDims(body) {
    const map = new Map();
    const walk = (stmts) => {
      for (const s of stmts || []) {
        if (s.type === "Dim") s.vars.forEach((v) => map.set(v.name.toLowerCase(), { v, node: s }));
        if (s.type === "If") {
          s.branches.forEach((b) => walk(b.body));
          walk(s.elseBody);
        }
        if (s.type === "Select") {
          s.cases.forEach((c) => walk(c.body));
          walk(s.elseBody);
        }
        if (s.type === "For" || s.type === "Do") walk(s.body);
      }
    };
    walk(body);
    return map;
  }

  function undefinedVar(name, node) {
    return compileError("変数が定義されていません。", node,
      `「${name}」という箱（変数）が、まだ用意されていません。`,
      `使う前に「Dim ${name} As Integer」のように Dim で用意しよう。Dim の行と、つづりが同じか（大文字・小文字以外）も確かめよう。`);
  }

  function unknownProp(ctlName, prop, node) {
    return rtError(438, node, {
      friendly: `${ctlName} には「${prop}」というプロパティはありません。`,
      hint: "つづりを確かめよう（ラベルは Caption、テキストボックスは Text）。",
    });
  }

  function defaultControlProps(type) {
    switch (type) {
      case "Label":
        return { Caption: "" };
      case "TextBox":
        return { Text: "" };
      case "CommandButton":
        return { Caption: "" };
      case "ListBox":
        return { List: [], ListIndex: -1 };
      default:
        return {};
    }
  }
  function defaultPropName(type) {
    if (type === "TextBox") return "Text";
    if (type === "ListBox") return "Value";
    return "Caption";
  }
  const CONTROL_PROPS = [
    "Caption", "Text", "Value", "Visible", "Enabled", "ForeColor", "BackColor", "BorderStyle", "SpecialEffect",
    "TextAlign", "Left", "Top", "Width", "Height", "FontSize", "FontBold", "FontItalic", "FontName", "Name",
    "ListIndex", "Locked", "WordWrap", "MultiLine", "Tag",
  ];
  function canonicalProp(type, k) {
    const found = CONTROL_PROPS.find((p) => p.toLowerCase() === k);
    if (!found) return null;
    if (found === "Text" && type !== "TextBox" && type !== "ListBox") return null;
    if (found === "Caption" && type === "TextBox") return null;
    return found;
  }

  function derefValue(v, node, rt) {
    if (!isObj(v)) return v;
    if (v.kind === "range") return rt.cellGet(v.addrs[0]);
    if (v.kind === "control") {
      const c = rt.state.controls[v.name];
      if (c.type === "TextBox") return c.Text;
      if (c.type === "ListBox") return c.ListIndex >= 0 ? (c.List || [])[c.ListIndex] : undefined;
      return c.Caption;
    }
    if (v.kind === "array") {
      throw rtError(13, node, { friendly: "配列全体は、そのまま計算や表示には使えません。", hint: "Umei(1) のように番号を付けよう。" });
    }
    if (v.kind === "sheet") return v.name;
    throw rtError(13, node, { friendly: "部品やセルなどの「もの」を、値として使おうとしています。", hint: "Label1.Caption のように、プロパティを付けよう。" });
  }

  function derefForVariant(v, node) {
    if (isObj(v) && v.kind === "array") return v;
    if (isObj(v)) throw rtError(13, node);
    return v;
  }

  function toNum(v, node) {
    if (typeof v === "number") return v;
    if (typeof v === "boolean") return v ? -1 : 0;
    if (v === undefined || v === null) return 0;
    if (typeof v === "string") {
      const n = parseNumber(v);
      if (n === null) {
        throw rtError(13, node, {
          friendly: v === "" ? "空っぽ（何も入っていない）の文字を、数として使おうとしました。" : `「${v}」は数ではないので、計算に使えません。`,
          hint: v === "" ? "入力が空っぽでないか確かめよう。Val を使うと、空っぽは 0 になります。" : "数を入れたか確かめよう。文字の先頭の数だけ使うなら Val を使おう。",
        });
      }
      return n;
    }
    throw rtError(13, node);
  }

  function toBool(v, node) {
    if (typeof v === "boolean") return v;
    if (typeof v === "number") return v !== 0;
    if (v === undefined || v === null) return false;
    if (typeof v === "string") {
      if (/^true$/i.test(v.trim())) return true;
      if (/^false$/i.test(v.trim())) return false;
      return toNum(v, node) !== 0;
    }
    throw rtError(13, node);
  }

  const INT_RANGE = { Integer: [-32768, 32767], Long: [-2147483648, 2147483647], Byte: [0, 255] };
  function coerce(type, v, node) {
    if (isObj(v)) v = undefined;
    switch (type) {
      case "Integer":
      case "Long":
      case "Byte": {
        const n = roundHalfEven(toNum(v, node));
        const [lo, hi] = INT_RANGE[type];
        if (n < lo || n > hi) {
          throw rtError(6, node, {
            friendly: `${toStr(n)} は、${TYPE_JA[type]}（${type}）の箱に入りきりません（${lo}〜${hi} まで）。`,
            hint: type === "Integer" ? "大きな数を入れるなら、Dim を As Long にしよう。" : "入れる数を確かめよう。",
          });
        }
        return n;
      }
      case "Single":
        return Number(Math.fround(toNum(v, node)).toPrecision(7));
      case "Double":
      case "Currency":
        return toNum(v, node);
      case "String":
        return toStr(v);
      case "Boolean":
        return toBool(v, node);
      case "Date":
        return v;
      default:
        return v;
    }
  }

  function compare(a, b, node) {
    if (typeof a === "boolean") a = a ? -1 : 0;
    if (typeof b === "boolean") b = b ? -1 : 0;
    if (a === undefined && b === undefined) return 0;
    if (typeof a === "string" && typeof b === "string") return a < b ? -1 : a > b ? 1 : 0;
    if (a === undefined) a = typeof b === "string" ? "" : 0;
    if (b === undefined) b = typeof a === "string" ? "" : 0;
    if (typeof a === "string" && typeof b === "string") return a < b ? -1 : a > b ? 1 : 0;
    if (typeof a === "string" || typeof b === "string") {
      const sa = typeof a === "string" ? a : b;
      if (parseNumber(sa) === null) {
        throw rtError(13, node, {
          friendly: `数と文字（"${sa}"）を比べようとしました。`,
          hint: "数どうし、文字どうしで比べよう。入力を数にするときは Val を使おう。",
        });
      }
    }
    const x = toNum(a, node);
    const y = toNum(b, node);
    return x < y ? -1 : x > y ? 1 : 0;
  }
  function compareOp(op, a, b, node) {
    const c = compare(a, b, node);
    switch (op) {
      case "=":
        return c === 0;
      case "<>":
        return c !== 0;
      case "<":
        return c < 0;
      case ">":
        return c > 0;
      case "<=":
        return c <= 0;
      case ">=":
        return c >= 0;
    }
    return false;
  }

  function binop(op, a, b, node) {
    switch (op) {
      case "+":
        if (typeof a === "string" && typeof b === "string") return a + b;
        if (a === undefined && typeof b === "string") return b;
        if (b === undefined && typeof a === "string") return a;
        return toNum(a, node.left) + toNum(b, node.right);
      case "-":
        return toNum(a, node.left) - toNum(b, node.right);
      case "*":
        return toNum(a, node.left) * toNum(b, node.right);
      case "/": {
        const y = toNum(b, node.right);
        const x = toNum(a, node.left);
        if (y === 0) throw rtError(11, node);
        return x / y;
      }
      case "\\": {
        const x = roundHalfEven(toNum(a, node.left));
        const y = roundHalfEven(toNum(b, node.right));
        if (y === 0) throw rtError(11, node);
        return Math.trunc(x / y);
      }
      case "Mod": {
        const x = roundHalfEven(toNum(a, node.left));
        const y = roundHalfEven(toNum(b, node.right));
        if (y === 0) throw rtError(11, node);
        const r = x % y;
        return r === 0 ? 0 : r;
      }
      case "^":
        return Math.pow(toNum(a, node.left), toNum(b, node.right));
      case "&":
        return toStr(a) + toStr(b);
      case "=":
      case "<>":
      case "<":
      case ">":
      case "<=":
      case ">=":
        return compareOp(op, a, b, node);
      case "And":
      case "Or":
      case "Xor":
        if (typeof a === "boolean" && typeof b === "boolean") {
          if (op === "And") return a && b;
          if (op === "Or") return a || b;
          return a !== b;
        }
        {
          const x = typeof a === "boolean" ? (a ? -1 : 0) : roundHalfEven(toNum(a, node.left));
          const y = typeof b === "boolean" ? (b ? -1 : 0) : roundHalfEven(toNum(b, node.right));
          if (op === "And") return x & y;
          if (op === "Or") return x | y;
          return x ^ y;
        }
      case "Like": {
        const pat = toStr(b)
          .replace(/[.+^${}()|\\]/g, "\\$&")
          .replace(/\*/g, ".*")
          .replace(/\?/g, ".")
          .replace(/#/g, "\\d");
        return new RegExp("^" + pat + "$").test(toStr(a));
      }
      case "Is":
        return a === b;
    }
    throw compileError("この演算子は使えません。", node);
  }

  // ---------- セル番地 ----------
  function colLetters(n) {
    let s = "";
    while (n > 0) {
      const m = (n - 1) % 26;
      s = String.fromCharCode(65 + m) + s;
      n = Math.floor((n - 1) / 26);
    }
    return s;
  }
  function lettersToNum(s) {
    let n = 0;
    for (const ch of s.toUpperCase()) n = n * 26 + (ch.charCodeAt(0) - 64);
    return n;
  }
  function colToNum(c, node) {
    if (typeof c === "string" && /^[A-Za-z]+$/.test(c.trim())) return lettersToNum(c.trim());
    return roundHalfEven(toNum(c, node));
  }
  function parseAddr(addr) {
    const m = String(addr).trim().match(/^\$?([A-Za-z]+)\$?(\d+)$/);
    if (!m) return null;
    return { col: lettersToNum(m[1]), row: parseInt(m[2], 10) };
  }
  function rectAddrs(a, b) {
    const out = [];
    const r1 = Math.min(a.row, b.row);
    const r2 = Math.max(a.row, b.row);
    const c1 = Math.min(a.col, b.col);
    const c2 = Math.max(a.col, b.col);
    for (let r = r1; r <= r2; r++) for (let c = c1; c <= c2; c++) out.push(colLetters(c) + r);
    return out;
  }
  function parseRangeAddr(text, node) {
    const parts = text.split(",");
    const out = [];
    for (const part of parts) {
      const [x, y] = part.split(":");
      const a = parseAddr(x);
      if (!a || a.row < 1) {
        throw rtError(1004, node, { friendly: `「${text}」はセルの番地として正しくありません。`, hint: "セルの番地は \"A1\" や \"B3\" のように、列の英字と行の数字で書こう。" });
      }
      if (y !== undefined) {
        const b = parseAddr(y);
        if (!b) throw rtError(1004, node, { friendly: `「${text}」はセルの番地として正しくありません。` });
        out.push(...rectAddrs(a, b));
      } else out.push(colLetters(a.col) + a.row);
    }
    return out;
  }
  function rangeLabel(rng) {
    if (rng.addrs.length === 1) return rng.addrs[0];
    return `${rng.addrs[0]}:${rng.addrs[rng.addrs.length - 1]}`;
  }
  function vbaColorToHex(n) {
    n = Math.round(n) & 0xffffff;
    const r = n & 255;
    const g = (n >> 8) & 255;
    const b = (n >> 16) & 255;
    return "#" + [r, g, b].map((x) => x.toString(16).padStart(2, "0")).join("");
  }
  function hexToVbaColor(hex) {
    const m = String(hex).match(/^#?([0-9a-f]{2})([0-9a-f]{2})([0-9a-f]{2})$/i);
    if (!m) return 16777215;
    return parseInt(m[1], 16) + parseInt(m[2], 16) * 256 + parseInt(m[3], 16) * 65536;
  }
  const COLOR_INDEX = { 1: "#000000", 2: "#ffffff", 3: "#ff0000", 4: "#00ff00", 5: "#0000ff", 6: "#ffff00", 7: "#ff00ff", 8: "#00ffff" };
  function colorIndexToHex(n) {
    return COLOR_INDEX[n] || "#ffffff";
  }

  // ---------- ファイルの読み込み ----------
  function atEOF(h) {
    return !/\S/.test(h.text.slice(h.pos).replace(/\u3000/g, "x"));
  }
  function readField(h) {
    const t = h.text;
    while (h.pos < t.length && (t[h.pos] === " " || t[h.pos] === "\t" || t[h.pos] === "\n")) h.pos++;
    let value;
    let quoted = false;
    if (t[h.pos] === '"') {
      quoted = true;
      const end = t.indexOf('"', h.pos + 1);
      const stop = end < 0 ? t.length : end;
      value = t.slice(h.pos + 1, stop);
      h.pos = stop + 1;
      while (h.pos < t.length && (t[h.pos] === " " || t[h.pos] === "\t")) h.pos++;
    } else {
      let end = h.pos;
      while (end < t.length && t[end] !== "," && t[end] !== "\n") end++;
      value = t.slice(h.pos, end).replace(/[ \t]+$/, "");
      h.pos = end;
    }
    if (t[h.pos] === "," || t[h.pos] === "\n") h.pos++;
    if (!quoted) {
      const n = value === "" ? null : parseNumber(value);
      if (n !== null && /^[\s+\-0-9.,eE]+$/.test(value)) return n;
    }
    return value;
  }

  // ---------- 組み込み関数（値だけで計算できるもの） ----------
  const BUILTIN_CONST = new Map([
    ["true", true], ["false", false], ["empty", undefined], ["null", null], ["nothing", undefined],
    ["vbcrlf", "\r\n"], ["vbnewline", "\r\n"], ["vbcr", "\r"], ["vblf", "\n"], ["vbtab", "\t"], ["vbnullstring", ""],
    ["vbokonly", 0], ["vbokcancel", 1], ["vbabortretryignore", 2], ["vbyesnocancel", 3], ["vbyesno", 4], ["vbretrycancel", 5],
    ["vbcritical", 16], ["vbquestion", 32], ["vbexclamation", 48], ["vbinformation", 64],
    ["vbok", 1], ["vbcancel", 2], ["vbabort", 3], ["vbretry", 4], ["vbignore", 5], ["vbyes", 6], ["vbno", 7],
    ["vbblack", 0], ["vbred", 255], ["vbgreen", 65280], ["vbyellow", 65535], ["vbblue", 16711680], ["vbmagenta", 16711935], ["vbcyan", 16776960], ["vbwhite", 16777215],
    ["xlup", -4162], ["xldown", -4121], ["xltoleft", -4159], ["xltoright", -4161], ["xlnone", -4142], ["xlcontinuous", 1],
    ["fmtextalignleft", 1], ["fmtextaligncenter", 2], ["fmtextalignright", 3],
  ]);
  const OBJECT_NAMES = new Set(["worksheets", "sheets", "activesheet", "thisworkbook", "activeworkbook", "application", "worksheetfunction", "debug", "cells", "selection", "controls"]);
  const BUILTIN_FUNCS = new Set([
    "msgbox", "inputbox", "val", "str", "cstr", "cint", "clng", "csng", "cdbl", "cbool", "cbyte", "ccur", "int", "fix", "abs", "sgn",
    "sqr", "exp", "log", "sin", "cos", "tan", "atn", "round", "rnd", "len", "left", "right", "mid", "trim", "ltrim", "rtrim",
    "ucase", "lcase", "instr", "instrrev", "replace", "space", "string", "strreverse", "chr", "chrw", "asc", "ascw", "format",
    "isnumeric", "isempty", "isnull", "isarray", "ubound", "lbound", "date", "now", "time", "year", "month", "day", "hour", "minute",
    "second", "weekday", "timer", "rgb", "eof", "freefile", "hex", "oct", "iif", "typename", "randomize", "doevents",
  ]);

  function builtinFunc(k, a, node) {
    const num = (i) => toNum(a[i], node);
    const str = (i) => toStr(a[i]);
    switch (k) {
      case "val":
        return vbaVal(str(0));
      case "str": {
        const n = num(0);
        return (n >= 0 ? " " : "") + numToStr(n);
      }
      case "cstr":
        return toStr(a[0]);
      case "cint":
        return coerce("Integer", a[0], node);
      case "clng":
        return coerce("Long", a[0], node);
      case "cbyte":
        return coerce("Byte", a[0], node);
      case "csng":
        return coerce("Single", a[0], node);
      case "cdbl":
      case "ccur":
        return num(0);
      case "cbool":
        return toBool(a[0], node);
      case "int":
        return Math.floor(num(0));
      case "fix":
        return Math.trunc(num(0));
      case "abs":
        return Math.abs(num(0));
      case "sgn":
        return Math.sign(num(0));
      case "sqr": {
        const n = num(0);
        if (n < 0) throw rtError(5, node, { friendly: "マイナスの数の平方根（Sqr）は計算できません。" });
        return Math.sqrt(n);
      }
      case "exp":
        return Math.exp(num(0));
      case "log": {
        const n = num(0);
        if (n <= 0) throw rtError(5, node);
        return Math.log(n);
      }
      case "sin":
        return Math.sin(num(0));
      case "cos":
        return Math.cos(num(0));
      case "tan":
        return Math.tan(num(0));
      case "atn":
        return Math.atan(num(0));
      case "round": {
        const d = a[1] === undefined ? 0 : num(1);
        const m = Math.pow(10, d);
        return roundHalfEven(num(0) * m) / m;
      }
      case "len":
        return a[0] === undefined ? 0 : str(0).length;
      case "left":
        return str(0).slice(0, Math.max(0, num(1)));
      case "right": {
        const s = str(0);
        const n = Math.max(0, num(1));
        return n === 0 ? "" : s.slice(-n);
      }
      case "mid": {
        const s = str(0);
        const start = num(1);
        if (start < 1) throw rtError(5, node, { friendly: "Mid の2番目（何文字目から）は 1 以上にしよう。" });
        return a[2] === undefined ? s.slice(start - 1) : s.substr(start - 1, Math.max(0, num(2)));
      }
      case "trim":
        return str(0).replace(/^[ \u3000]+|[ \u3000]+$/g, "");
      case "ltrim":
        return str(0).replace(/^[ \u3000]+/, "");
      case "rtrim":
        return str(0).replace(/[ \u3000]+$/, "");
      case "ucase":
        return str(0).toUpperCase();
      case "lcase":
        return str(0).toLowerCase();
      case "instr": {
        let start = 1;
        let s;
        let f;
        if (a.length >= 3 && typeof a[0] === "number") {
          start = num(0);
          s = str(1);
          f = str(2);
        } else {
          s = str(0);
          f = str(1);
        }
        return s.indexOf(f, start - 1) + 1;
      }
      case "instrrev":
        return str(0).lastIndexOf(str(1)) + 1;
      case "replace":
        return str(0).split(str(1)).join(str(2));
      case "space":
        return " ".repeat(Math.max(0, num(0)));
      case "string":
        return (typeof a[1] === "number" ? String.fromCharCode(a[1]) : str(1).charAt(0)).repeat(Math.max(0, num(0)));
      case "strreverse":
        return str(0).split("").reverse().join("");
      case "chr":
      case "chrw":
        return String.fromCharCode(num(0));
      case "asc":
      case "ascw": {
        const s = str(0);
        if (!s) throw rtError(5, node, { friendly: "空っぽの文字には Asc を使えません。" });
        return s.charCodeAt(0);
      }
      case "format":
        return vbaFormat(a[0], a[1] === undefined ? undefined : str(1));
      case "isnumeric":
        return typeof a[0] === "number" || typeof a[0] === "boolean" || a[0] === undefined || (typeof a[0] === "string" && parseNumber(a[0]) !== null);
      case "isempty":
        return a[0] === undefined;
      case "isnull":
        return a[0] === null;
      case "date": {
        const d = new Date();
        return `${d.getFullYear()}/${String(d.getMonth() + 1).padStart(2, "0")}/${String(d.getDate()).padStart(2, "0")}`;
      }
      case "now": {
        const d = new Date();
        return `${d.getFullYear()}/${String(d.getMonth() + 1).padStart(2, "0")}/${String(d.getDate()).padStart(2, "0")} ${d.toTimeString().slice(0, 8)}`;
      }
      case "time":
        return new Date().toTimeString().slice(0, 8);
      case "year":
      case "month":
      case "day":
      case "hour":
      case "minute":
      case "second":
      case "weekday": {
        const d = a[0] === undefined ? new Date() : new Date(str(0).replace(/\//g, "-"));
        if (isNaN(d)) throw rtError(13, node);
        return { year: d.getFullYear(), month: d.getMonth() + 1, day: d.getDate(), hour: d.getHours(), minute: d.getMinutes(), second: d.getSeconds(), weekday: d.getDay() + 1 }[k];
      }
      case "timer": {
        const d = new Date();
        return d.getHours() * 3600 + d.getMinutes() * 60 + d.getSeconds() + d.getMilliseconds() / 1000;
      }
      case "rgb":
        return (num(0) & 255) + (num(1) & 255) * 256 + (num(2) & 255) * 65536;
      case "hex":
        return Math.round(num(0)).toString(16).toUpperCase();
      case "oct":
        return Math.round(num(0)).toString(8);
      case "typename": {
        const v = a[0];
        if (v === undefined) return "Empty";
        if (typeof v === "string") return "String";
        if (typeof v === "boolean") return "Boolean";
        if (typeof v === "number") return Number.isInteger(v) && Math.abs(v) <= 32767 ? "Integer" : "Double";
        return "Object";
      }
      case "doevents":
        return 0;
    }
    throw compileError("この関数は使えません。", node);
  }

  function wsFunction(k, rawArgs, node, rt) {
    const flat = [];
    rawArgs.forEach((v) => {
      if (isObj(v) && v.kind === "range") v.addrs.forEach((a) => flat.push(rt.cellGet(a)));
      else flat.push(isObj(v) ? derefValue(v, node, rt) : v);
    });
    const args = rawArgs.map((v) => (isObj(v) ? derefValue(v, node, rt) : v));
    const nums = flat.filter((v) => typeof v === "number");
    switch (k) {
      case "round":
        return roundHalfAway(toNum(args[0], node), toNum(args[1] || 0, node));
      case "roundup": {
        const m = Math.pow(10, toNum(args[1] || 0, node));
        const x = toNum(args[0], node);
        return (x < 0 ? -Math.ceil(-x * m - 1e-9) : Math.ceil(x * m - 1e-9)) / m;
      }
      case "rounddown": {
        const m = Math.pow(10, toNum(args[1] || 0, node));
        return Math.trunc(toNum(args[0], node) * m + 1e-9 * Math.sign(toNum(args[0], node))) / m;
      }
      case "sum":
        return nums.reduce((s, x) => s + x, 0);
      case "max":
        return nums.length ? Math.max(...nums) : 0;
      case "min":
        return nums.length ? Math.min(...nums) : 0;
      case "average":
        if (!nums.length) throw rtError(1004, node);
        return nums.reduce((s, x) => s + x, 0) / nums.length;
      case "count":
        return nums.length;
      case "counta":
        return flat.filter((v) => v !== undefined && v !== "").length;
    }
    throw rtError(438, node, { friendly: `WorksheetFunction.${k} はこのツールではまだ使えません。` });
  }

  // 教科書 p.93：先頭から順に数値に変換し、数字以外の文字が見つかったらそこで止める。数字が無ければ 0
  function vbaVal(s) {
    const t = normalizeNumText(s).replace(/[ \t\r\n]/g, "");
    if (/^&H[0-9A-F]+/i.test(t)) return parseInt(t.slice(2).match(/^[0-9A-F]+/i)[0], 16);
    const m = t.match(/^[+-]?(\d+\.?\d*|\.\d+)([eEdD][+-]?\d+)?/);
    if (!m) return 0;
    return Number(m[0].replace(/[dD]/, "e"));
  }

  // Format（よく使う形だけ）
  function vbaFormat(v, fmt) {
    if (fmt === undefined || fmt === "") return toStr(v);
    const named = fmt.toLowerCase();
    let n = typeof v === "number" ? v : typeof v === "string" ? parseNumber(v) : v === undefined ? 0 : typeof v === "boolean" ? (v ? -1 : 0) : null;
    if (named === "general number") return n === null ? toStr(v) : toStr(n);
    if (n === null) return toStr(v);
    if (named === "currency") return (n < 0 ? "-" : "") + "￥" + groupDigits(roundHalfAway(Math.abs(n), 0).toFixed(0));
    if (named === "fixed") return roundHalfAway(n, 2).toFixed(2);
    if (named === "standard") return (n < 0 ? "-" : "") + groupDigits(roundHalfAway(Math.abs(n), 2).toFixed(2));
    if (named === "percent") return roundHalfAway(n * 100, 2).toFixed(2) + "%";
    const sections = fmt.split(";");
    let pat = sections[0];
    let neg = false;
    if (n < 0) {
      if (sections[1] !== undefined) pat = sections[1];
      else neg = true;
      n = Math.abs(n);
    }
    // 先頭・末尾の飾り文字（"円" や ¥ など）を取り出す
    const m = pat.match(/^([^#0,.%]*)([#0,.%]*)(.*)$/);
    if (!m) return toStr(v);
    const strip = (s) => s.replace(/"([^"]*)"/g, "$1").replace(/\\(.)/g, "$1");
    const prefix = strip(m[1]);
    let body = m[2];
    const suffix = strip(m[3]);
    let percent = false;
    if (body.includes("%")) {
      percent = true;
      body = body.replace(/%/g, "");
      n *= 100;
    }
    const [intPat, decPat = ""] = body.split(".");
    const decimals = decPat.replace(/[^#0]/g, "").length;
    const minDec = (decPat.match(/0/g) || []).length;
    const minInt = (intPat.match(/0/g) || []).length;
    const group = intPat.includes(",");
    let s = roundHalfAway(n, decimals).toFixed(decimals);
    let [ip, dp = ""] = s.split(".");
    if (ip === "0" && minInt === 0) ip = "";
    while (ip.length < minInt) ip = "0" + ip;
    if (group) ip = groupDigits(ip);
    while (dp.length > minDec && dp.endsWith("0")) dp = dp.slice(0, -1);
    let out = ip + (decPat.length ? (dp.length ? "." + dp : decPat.includes("0") ? "." : "") : "");
    if (out.endsWith(".")) out = out.slice(0, -1);
    if (out === "") out = "";
    return (neg ? "-" : "") + prefix + out + (percent ? "%" : "") + suffix;
  }
  function groupDigits(s) {
    return s.replace(/\B(?=(\d{3})+(?!\d))/g, ",");
  }

  // ---------- step の再生（差分から各時点の状態を作る） ----------
  class StepStore {
    constructor(initial, steps) {
      this.initial = initial;
      this.steps = steps;
      this.cacheIndex = -1;
      this.cache = null;
    }
    // index: -1 = 実行前の状態、n = n 番目の step を実行した後の状態
    modelAt(index) {
      if (index < 0 || !this.steps.length) {
        if (this.cacheIndex !== -1 || !this.cache) {
          this.cache = structuredClone(this.initial);
          this.cacheIndex = -1;
        }
        return this.cache;
      }
      index = Math.min(index, this.steps.length - 1);
      if (this.cache && this.cacheIndex <= index && index - this.cacheIndex <= KEYFRAME_EVERY) {
        for (let i = this.cacheIndex + 1; i <= index; i++) applyOps(this.cache, this.steps[i].ops);
        this.cacheIndex = index;
        return this.cache;
      }
      let base = -1;
      for (let i = index; i >= 0; i--) {
        if (this.steps[i].kf) {
          base = i;
          break;
        }
      }
      this.cache = structuredClone(base >= 0 ? this.steps[base].kf : this.initial);
      for (let i = base + 1; i <= index; i++) applyOps(this.cache, this.steps[i].ops);
      this.cacheIndex = index;
      return this.cache;
    }
    // 書き換えてもよい写しを返す
    modelCopy(index) {
      return structuredClone(this.modelAt(index));
    }
  }
  function applyOps(state, ops) {
    for (const [path, value] of ops) setPath(state, path, value === undefined ? undefined : structuredClone(value));
  }

  const api = {
    Runtime, StepStore, VbaRuntimeError, ExitSignal, EndSignal,
    toStr, toNum, numToStr, vbaVal, vbaFormat, roundHalfEven, TYPE_JA, MAX_STEPS, MAX_INSTR,
  };
  if (typeof module !== "undefined" && module.exports) module.exports = api;
  else root.VbaRuntime = api;
})(typeof window !== "undefined" ? window : globalThis);
