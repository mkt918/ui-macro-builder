/* ===== 仮想 Excel プレビュー =====
 * 実行エンジン（src/vba/session.js）が記録した step を、HTML テーブルに描いて再生する。
 *
 * ExcelView : 描画・再生（アニメーション）・タイムライン・セルの直接入力
 *
 * cursor は「何個の step を実行し終えたか」。0 = 実行前、n = n 個目の step の後の状態。
 * 各時点の状態は session.store.modelAt(i) で取り出す（差分から復元される）。
 */

const GRID_ROWS = 26; // 最小表示行数
const GRID_COLS = 26; // 最小表示列数（Z列まで）
const MAX_GRID_ROWS = 100; // 安全上限
const MAX_GRID_COLS = 26; // Z列
const MAX_ARRAY_DISPLAY = 30; // 配列ビジュアライザの表示上限

const TYPE_JA_PV = {
  Integer: "整数型", Long: "長整数型", Single: "単精度浮動小数点数型", Double: "倍精度浮動小数点数型",
  String: "文字列型", Variant: "バリアント型", Boolean: "ブール型", Date: "日付型", Currency: "通貨型",
  Byte: "バイト型", Object: "オブジェクト型",
};

function colLetter(n) {
  // 1 -> A
  return String.fromCharCode(64 + n);
}
function colNum(letter) {
  return letter.toUpperCase().charCodeAt(0) - 64;
}
// HTMLエスケープ（ビジュアライザ描画用）
function escapeHtmlPv(s) {
  return String(s).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
}
// "A1" -> {col:1, row:1}
function parseAddr(addr) {
  const m = String(addr).match(/^([A-Za-z]+)(\d+)$/);
  if (!m) return null;
  return { col: colNum(m[1]), row: parseInt(m[2], 10), letter: m[1].toUpperCase() };
}
// 変数・配列の値を表示用の文字に
function showVal(v) {
  if (v === undefined || v === null) return "";
  if (typeof v === "boolean") return v ? "True" : "False";
  if (typeof v === "object") return "…";
  return String(v);
}
function typeJa(t) {
  if (!t) return "";
  const base = String(t).replace(/（定数）$/, "");
  return (TYPE_JA_PV[base] || base) + (/（定数）$/.test(t) ? "・定数" : "");
}

/* ---------- ビュー ---------- */
class ExcelView {
  constructor(els) {
    this.table = els.table;
    this.refEl = els.ref;
    this.formulaEl = els.formula;
    this.statusEl = els.status;
    this.arrayEl = els.array; // 配列ビジュアライザのコンテナ
    this.varEl = els.vars || null; // 変数ウォッチのコンテナ
    this.msgOverlay = els.msgbox || null; // MsgBox オーバーレイ（タイムラインで見返すとき用）
    this.tabsEl = els.tabs; // シートタブのコンテナ
    this.onEdit = els.onEdit || null; // セル編集時のコールバック
    this.onPlayState = els.onPlayState || null; // 再生状態変更コールバック
    this.onStepChange = els.onStepChange || null; // カーソル移動時（タイムライン同期用）
    this.onStepApplied = els.onStepApplied || null; // step を表示したとき（行・ブロックのハイライト用）
    // 再生中・1コマで MsgBox の step に来たとき（ダイアログを出す）。done() で再生を続ける
    this.onMsgStep = els.onMsgStep || null;
    // 再生中に記録の最後まで来たとき。InputBox などの答え待ちなら、答えのあと resume() を呼んで true を返す
    this.onReachEnd = els.onReachEnd || null;
    this.selected = "A1"; // キーボード操作の選択セル
    if (this.msgOverlay) {
      const ok = this.msgOverlay.querySelector(".msgbox-ok");
      if (ok) ok.addEventListener("click", () => this.hideMsgBox());
    }
    this.session = null;
    this.store = null;
    this.steps = [];
    this.cursor = 0;
    this.timer = null;
    this.speed = 500;
    this.rows = GRID_ROWS;
    this.cols = GRID_COLS;
    this.initialCells = {}; // 生徒が直接入力した初期データ
    this.editing = false; // 編集中フラグ（描画でclobberしない用）
    this.playing = false; // アニメーション中は編集不可
    this.idleText = "セルに値を入力したり、ブロックを組み立てよう";
    this.buildGrid(GRID_ROWS, GRID_COLS);
    this.bindEditing();
  }

  buildGrid(rows, cols) {
    this.rows = rows;
    this.cols = cols;
    let html = "<tr><th class='corner'></th>";
    for (let c = 1; c <= cols; c++) html += `<th>${colLetter(c)}</th>`;
    html += "</tr>";
    for (let r = 1; r <= rows; r++) {
      html += `<tr><td class='row-num'>${r}</td>`;
      for (let c = 1; c <= cols; c++) {
        const addr = colLetter(c) + r;
        html += `<td id='cell-${addr}' data-addr='${addr}' class='editable'></td>`;
      }
      html += "</tr>";
    }
    this.table.innerHTML = html;
  }

  // セル編集（イベント委譲でテーブル全体に1度だけ設定）
  bindEditing() {
    this.table.setAttribute("tabindex", "0"); // キーボード操作を受け取るため

    this.table.addEventListener("dblclick", (e) => {
      const td = e.target.closest("td.editable");
      if (!td || this.playing) return;
      this.beginEdit(td);
    });

    // シングルクリックで選択
    this.table.addEventListener("click", (e) => {
      const td = e.target.closest("td.editable");
      if (!td || this.playing || this.editing) return;
      const addr = td.dataset.addr;
      this.selectCell(addr);

      // Blocklyのテキストフィールドが編集中なら、そこにセルアドレスを入力
      const active = document.activeElement;
      if (active && active.classList.contains("blocklyHtmlInput")) {
        active.value = addr;
        active.dispatchEvent(new Event("input", { bubbles: true }));
        active.dispatchEvent(new KeyboardEvent("keydown", { key: "Enter", bubbles: true }));
      } else {
        this.table.focus({ preventScroll: true });
      }
    });

    // Excel ライクなキーボード操作
    this.table.addEventListener("keydown", (e) => {
      if (this.editing || this.playing) return; // 編集中・再生中は無視
      const addr = this.selected;
      if (!addr) return;
      const td = document.getElementById("cell-" + addr);

      switch (e.key) {
        case "ArrowUp":
          e.preventDefault();
          this.moveSelection(-1, 0);
          return;
        case "ArrowDown":
          e.preventDefault();
          this.moveSelection(1, 0);
          return;
        case "ArrowLeft":
          e.preventDefault();
          this.moveSelection(0, -1);
          return;
        case "ArrowRight":
          e.preventDefault();
          this.moveSelection(0, 1);
          return;
        case "Tab":
          e.preventDefault();
          this.moveSelection(0, e.shiftKey ? -1 : 1);
          return;
        case "Enter":
        case "F2":
          e.preventDefault();
          if (td) this.beginEdit(td);
          return;
        case "Delete":
        case "Backspace":
          e.preventDefault();
          this.setInitialCell(addr, "");
          if (this.onEdit) this.onEdit();
          return;
        default:
          break;
      }
      // 文字キーを押したらそのまま入力開始（Excel と同じ挙動）
      if (e.key.length === 1 && !e.ctrlKey && !e.metaKey && !e.altKey) {
        e.preventDefault();
        if (td) this.beginEdit(td, e.key);
      }
    });
  }

  // セルを選択（キーボード操作・クリック共通）
  selectCell(addr) {
    const td = document.getElementById("cell-" + addr);
    if (!td) return;
    const prev = this.selected && document.getElementById("cell-" + this.selected);
    if (prev) prev.classList.remove("selected");
    this.selected = addr;
    td.classList.add("selected");
    td.scrollIntoView({ block: "nearest", inline: "nearest" });
    this.refEl.textContent = addr;
    // いま表示しているモデルの値を数式バーに出す
    const model = this.currentModel();
    const c = (model.cells || {})[addr];
    this.formulaEl.textContent = c && c.value !== undefined ? c.value : "";
  }

  // 選択セルを相対移動（グリッド外には出ない）
  moveSelection(dr, dc) {
    const p = parseAddr(this.selected || "A1");
    if (!p) return;
    const row = Math.max(1, Math.min(this.rows, p.row + dr));
    const col = Math.max(1, Math.min(this.cols, p.col + dc));
    this.selectCell(colLetter(col) + row);
  }

  // initialChar: 文字キーで編集開始したときの1文字目（Excel と同じ挙動）
  beginEdit(td, initialChar) {
    const addr = td.dataset.addr;
    this.editing = true;
    this.selectCell(addr);
    td.contentEditable = "true";
    td.classList.add("editing");
    // 現在の初期値を表示
    const cur = this.initialCells[addr];
    const prevText = cur && cur.value !== undefined ? String(cur.value) : "";
    td.textContent = initialChar !== undefined ? initialChar : prevText;
    td.focus();
    // 文字入力で始めたときは末尾にキャレット、それ以外は全選択
    const range = document.createRange();
    range.selectNodeContents(td);
    if (initialChar !== undefined) range.collapse(false);
    const sel = window.getSelection();
    sel.removeAllRanges();
    sel.addRange(range);

    let moveAfter = null; // 確定後の移動方向 [dr, dc]
    const commit = () => {
      td.contentEditable = "false";
      td.classList.remove("editing");
      this.editing = false;
      const raw = td.textContent.trim();
      this.setInitialCell(addr, raw);
      td.removeEventListener("blur", onBlur);
      td.removeEventListener("keydown", onKey);
      if (this.onEdit) this.onEdit();
      // 確定後もキーボードで続けて操作できるようにする
      this.table.focus({ preventScroll: true });
      if (moveAfter) this.moveSelection(moveAfter[0], moveAfter[1]);
    };
    const onBlur = () => commit();
    const onKey = (e) => {
      if (e.key === "Enter") {
        e.preventDefault();
        moveAfter = [1, 0]; // 確定して下へ
        td.blur();
      } else if (e.key === "Tab") {
        e.preventDefault();
        moveAfter = [0, e.shiftKey ? -1 : 1]; // 確定して横へ
        td.blur();
      } else if (e.key === "Escape") {
        e.preventDefault();
        td.textContent = prevText;
        td.blur();
      }
    };
    td.addEventListener("blur", onBlur);
    td.addEventListener("keydown", onKey);
  }

  // 初期データを設定（数値ならNumber、空なら値だけ削除して書式は残す）
  setInitialCell(addr, raw) {
    const cur = this.initialCells[addr] || {};
    if (raw === "") {
      delete cur.value;
      // 書式も何も無くなったらセルごと削除
      if (Object.keys(cur).length === 0) delete this.initialCells[addr];
      else this.initialCells[addr] = cur;
      return;
    }
    const num = Number(raw);
    cur.value = raw !== "" && !isNaN(num) ? num : raw;
    this.initialCells[addr] = cur;
  }

  setInitialCells(obj) {
    this.initialCells = obj ? JSON.parse(JSON.stringify(obj)) : {};
  }
  getInitialCells() {
    return this.initialCells;
  }

  // ---------- 各時点の状態 ----------
  modelAt(index) {
    if (this.store) return this.store.modelAt(index);
    return this.initialModel();
  }
  // 実行前の状態（実行エンジンが無いとき＝文法エラー中などは、直接入力したセルだけ）
  initialModel() {
    if (this.store) return this.store.modelAt(-1);
    return { cells: this.initialCells, arrays: {}, vars: {}, varTypes: {}, sheet: "Sheet1", sheets: ["Sheet1"], msgs: [], controls: {} };
  }
  // いま画面に映しているモデル（cursor は「次に実行するステップ」を指す）
  currentModel() {
    return this.modelAt(this.cursor - 1);
  }

  // いまの表示内容を「実行開始時の状態」として焼き付ける。
  // 途中結果を土台に次のアルゴリズムを組み立てられるようにするための機能。
  bakeCurrentAsInitial() {
    const model = this.currentModel();
    const next = {};
    for (const addr in model.cells) {
      const c = model.cells[addr];
      if (!c) continue;
      const keep = {};
      if (c.value !== undefined && c.value !== "") keep.value = c.value;
      if (c.bg) keep.bg = c.bg;
      if (c.bold) keep.bold = true;
      if (c.fontSize) keep.fontSize = c.fontSize;
      if (c.border) keep.border = true;
      if (Object.keys(keep).length) next[addr] = keep;
    }
    this.initialCells = next;
    return Object.keys(next).length;
  }

  // 初期データを全消去
  clearInitial() {
    this.initialCells = {};
  }

  // 初期状態と比べて変化したセルのアドレス集合（結果ハイライト用）
  diffFromInitial(model) {
    const diff = new Set();
    const base = this.initialCells || {};
    const cells = (model && model.cells) || {};
    const sig = (c) =>
      !c
        ? ""
        : [
            c.value !== undefined ? c.value : "",
            c.bg || "",
            c.bold ? "b" : "",
            c.fontSize || "",
            c.border ? "r" : "",
          ].join("");
    const addrs = new Set([...Object.keys(base), ...Object.keys(cells)]);
    addrs.forEach((addr) => {
      if (sig(base[addr]) !== sig(cells[addr])) diff.add(addr);
    });
    return diff;
  }

  // 記録に出てくるセルと初期データから、必要なグリッドの大きさを決めて作り直す
  fitGrid() {
    let maxRow = GRID_ROWS;
    let maxCol = GRID_COLS;
    const bump = (addr) => {
      const p = parseAddr(addr);
      if (!p) return;
      if (p.row > maxRow) maxRow = p.row;
      if (p.col > maxCol) maxCol = p.col;
    };
    for (const step of this.steps) {
      for (const [path] of step.ops) if (path[0] === "cells" && path[1]) bump(path[1]);
    }
    for (const addr in this.initialCells) bump(addr);
    maxRow = Math.min(maxRow, MAX_GRID_ROWS);
    maxCol = Math.min(maxCol, MAX_GRID_COLS);
    if (maxRow !== this.rows || maxCol !== this.cols) {
      this.buildGrid(maxRow, maxCol);
    }
  }

  // 配列の箱の、最近の書きかわり：この step で変わった箱（strong）と、直前の2つの step で変わった箱（soft）。
  // 並べ替えの「交換」（2つの箱の中身が入れかわる）の様子を目で追うために使う
  recentArrayChanges(idx) {
    const strong = new Set();
    const soft = new Set();
    const collect = (i, into) => {
      const st = this.steps[i];
      if (!st) return;
      for (const [path] of st.ops) {
        if (path[0] === "arrays" && path[2] === "data" && path[1] !== undefined) into.add(`${path[1]}|${path[3]}`);
      }
    };
    collect(idx, strong);
    collect(idx - 1, soft);
    collect(idx - 2, soft);
    strong.forEach((k) => soft.delete(k));
    return { strong, soft };
  }

  // モデルを丸ごと描画
  // diffSet: 初期状態から変化したセル（結果表示時にハイライトする）
  // recent: 配列の箱の最近の書きかわり（recentArrayChanges の結果）
  renderModel(model, active, changed, diffSet, recent) {
    const cells = model.cells || {};
    const base = this.initialCells || {};
    for (let r = 1; r <= this.rows; r++) {
      for (let c = 1; c <= this.cols; c++) {
        const addr = colLetter(c) + r;
        const td = document.getElementById("cell-" + addr);
        if (!td) continue;
        if (td.classList.contains("editing")) continue; // 編集中はclobberしない
        const cell = cells[addr] || {};
        td.textContent = cell.value !== undefined ? cell.value : "";
        td.style.background = cell.bg || "";
        td.style.fontWeight = cell.bold ? "bold" : "";
        td.style.fontSize = cell.fontSize ? cell.fontSize + "px" : "";
        // 罫線は inset box-shadow で描く（active-cell の outline と共存させるため）
        td.style.boxShadow = cell.border ? "inset 0 0 0 1.5px #555" : "";
        td.style.color = cell.color || (cell.bg && cell.bg !== "#ffffff" ? "#fff" : "");
        td.classList.remove("active-cell", "changed", "result-diff");
        if (active && active.scope === "cell" && addr === active.key) td.classList.add("active-cell");
        if (changed && changed.scope === "cell" && addr === changed.key) td.classList.add("changed");
        if (diffSet && diffSet.has(addr)) td.classList.add("result-diff");
        // 実行の出発点（初期データ）を持つセルに印を付ける
        td.classList.toggle("has-initial", !!base[addr]);
        // 選択中セル
        td.classList.toggle("selected", addr === this.selected);
      }
    }
    this.renderArray(model.arrays || {}, active, recent);
    this.renderVars(model.vars || {}, model.varTypes || {}, active);
    this.renderTabs(model.sheets || ["Sheet1"], model.sheet || "Sheet1");
  }

  // 変数（教科書の絵と同じ「名前の付いた箱」）
  renderVars(vars, types, active) {
    if (!this.varEl) return;
    const names = Object.keys(vars).concat(Object.keys(types).filter((n) => !(n in vars)));
    if (names.length === 0) {
      this.varEl.classList.remove("show");
      this.varEl.innerHTML = "";
      return;
    }
    this.varEl.classList.add("show");
    let html = '<span class="var-label" title="変数は、値を入れておく「箱」です">📦 変数（箱）</span><div class="var-boxes">';
    names.forEach((n) => {
      const isActive = active && (active.scope === "var" || active.scope === "file") && (active.key === n || (active.names && active.names.includes(n)));
      const v = vars[n];
      const empty = v === undefined || v === "";
      const t = types[n];
      html += `<div class="var-box ${isActive ? "active" : ""}" title="${escapeHtmlPv(n)}（${escapeHtmlPv(typeJa(t))}）">
        <div class="var-box-val ${empty ? "empty" : ""}">${empty ? "&nbsp;" : escapeHtmlPv(showVal(v))}</div>
        <div class="var-box-name">${escapeHtmlPv(n)}</div>
        ${t ? `<div class="var-box-type">${escapeHtmlPv(typeJa(t))}</div>` : ""}
      </div>`;
    });
    html += "</div>";
    this.varEl.innerHTML = html;
  }

  // 配列（教科書の絵と同じく、添字の付いた箱を横に並べる）
  renderArray(arrays, active, recent) {
    if (!this.arrayEl) return;
    const rc = recent || { strong: new Set(), soft: new Set() };
    const names = Object.keys(arrays);
    if (names.length === 0) {
      this.arrayEl.classList.remove("show");
      this.arrayEl.innerHTML = "";
      return;
    }
    this.arrayEl.classList.add("show");
    let html = "";
    names.forEach((name) => {
      const a = arrays[name];
      const data = a.data || {};
      const dims = a.dims || [];
      const activeKey = active && active.scope === "array" && active.arrayName === name ? active.index : null;
      html += `<div class="array-row"><span class="array-label">📦 配列 ${escapeHtmlPv(name)}<small>（${escapeHtmlPv(typeJa(a.type))}）</small></span>`;
      if (dims.length === 0) {
        html += '<span class="array-note">大きさがまだ決まっていません（ReDim で決めよう）</span></div>';
        return;
      }
      if (dims.length === 1) {
        const [lo, hi] = dims[0];
        const last = Math.min(hi, lo + MAX_ARRAY_DISPLAY - 1);
        html += '<div class="array-cells">';
        for (let i = lo; i <= last; i++) {
          const key = String(i);
          const has = Object.prototype.hasOwnProperty.call(data, key);
          const cls = ["array-cell"];
          if (!has) cls.push("unused");
          if (activeKey === key) cls.push("active-cell");
          if (rc.strong.has(`${name}|${key}`)) cls.push("changed-now");
          else if (rc.soft.has(`${name}|${key}`)) cls.push("changed-prev");
          html += `<div class="${cls.join(" ")}">
            <div class="array-val">${has ? escapeHtmlPv(showVal(data[key])) : "&nbsp;"}</div>
            <div class="array-idx">${i}</div>
          </div>`;
        }
        if (hi > last) html += `<div class="array-cell array-more">… +${hi - last}</div>`;
        html += "</div></div>";
        return;
      }
      // 2次元：表で表示（行が1つ目の添字、列が2つ目の添字）
      const [[r0, r1], [c0, c1]] = dims;
      const rLast = Math.min(r1, r0 + 19);
      const cLast = Math.min(c1, c0 + 14);
      html += '<table class="array-2d"><tr><th></th>';
      for (let c = c0; c <= cLast; c++) html += `<th>${c}</th>`;
      html += "</tr>";
      for (let r = r0; r <= rLast; r++) {
        html += `<tr><th>${r}</th>`;
        for (let c = c0; c <= cLast; c++) {
          const key = `${r},${c}`;
          const has = Object.prototype.hasOwnProperty.call(data, key);
          const ch = rc.strong.has(`${name}|${key}`) ? "changed-now" : rc.soft.has(`${name}|${key}`) ? "changed-prev" : "";
          html += `<td class="${has ? "" : "unused"} ${activeKey === key ? "active-cell" : ""} ${ch}">${has ? escapeHtmlPv(showVal(data[key])) : ""}</td>`;
        }
        html += "</tr>";
      }
      html += "</table></div>";
    });
    this.arrayEl.innerHTML = html;
  }

  // MsgBox オーバーレイ（タイムラインで見返すとき用。再生中・1コマのときはダイアログを使う）
  showMsgBox(text, title) {
    if (!this.msgOverlay) return;
    this.msgOverlay.querySelector(".msgbox-body").textContent = text;
    const t = this.msgOverlay.querySelector(".msgbox-title");
    if (t) t.textContent = title || "Microsoft Excel";
    this.msgOverlay.hidden = false;
  }
  hideMsgBox() {
    if (this.msgOverlay) this.msgOverlay.hidden = true;
  }

  // シートタブ
  renderTabs(sheets, current) {
    if (!this.tabsEl) return;
    this.tabsEl.innerHTML = sheets
      .map((s) => `<div class="sheet-tab ${s === current ? "active" : ""}">${escapeHtmlPv(s)}</div>`)
      .join("");
  }

  // 実行（session）を読み込む。session が null のときは、直接入力したセルだけを表示する
  loadSession(session, statusText) {
    this.stop();
    this.session = session;
    this.store = session ? session.store : null;
    this.steps = session ? session.steps : [];
    this.cursor = 0;
    this.hideMsgBox();
    this.fitGrid();
    this.renderModel(this.initialModel(), null, null);
    this.refEl.textContent = "A1";
    this.formulaEl.textContent = "";
    this.statusEl.textContent = statusText || (this.steps.length ? `準備完了（全 ${this.steps.length} ステップ）` : this.idleText);
    if (this.onStepApplied) this.onStepApplied(null);
    this.notifyStepChange();
  }

  // session の記録が後ろに増えたとき（InputBox に答えた・ボタンを押したなど）
  stepsAppended() {
    this.fitGrid();
    this.notifyStepChange();
  }

  // 実行がすべて終わっていて、最後の step を表示しているか
  isFinished(idx) {
    const s = this.session;
    // フォームで実行中は、命令が終わってもボタンを押すのを待っているので「おわり」ではない
    if (this.liveForm && s && s.status === "idle") return false;
    const done = !s || (s.status !== "waiting" && s.status !== "running");
    return done && idx === this.steps.length - 1;
  }

  // interactive: 再生・1コマで進めたとき（MsgBox はダイアログで出す）
  applyStep(idx, interactive) {
    const step = this.steps[idx];
    if (!step) return;
    const model = this.modelAt(idx);
    const marker = { scope: step.scope, key: step.key, arrayName: step.arrayName, index: step.index, names: step.names };
    // 最後のステップ＝実行結果。初期状態から変わったセルをまとめて光らせる
    const isLast = this.isFinished(idx);
    const diffSet = isLast ? this.diffFromInitial(model) : null;
    this.renderModel(model, marker, marker, diffSet, this.recentArrayChanges(idx));
    if (step.scope === "msg" && !interactive) {
      this.showMsgBox(step.key, step.title);
    } else {
      this.hideMsgBox();
    }
    // 数式バー（左の欄に「どこ」、右の欄に「何」）
    if (step.scope === "cell") {
      this.refEl.textContent = step.key;
      const cell = model.cells[step.key];
      this.formulaEl.textContent = cell && cell.value !== undefined ? cell.value : "";
      const td = document.getElementById("cell-" + step.key);
      if (td) td.scrollIntoView({ block: "nearest", inline: "nearest" });
    } else if (step.scope === "array") {
      this.refEl.textContent = step.key;
      const a = model.arrays && model.arrays[step.arrayName];
      this.formulaEl.textContent = a && a.data ? showVal(a.data[step.index]) : "";
    } else if (step.scope === "var") {
      this.refEl.textContent = step.key;
      this.formulaEl.textContent = model.vars ? showVal(model.vars[step.key]) : "";
    } else if (step.scope === "msg") {
      this.refEl.textContent = "MsgBox";
      this.formulaEl.textContent = step.key;
    } else if (step.scope === "control" || step.scope === "user") {
      this.refEl.textContent = step.key;
      const c = model.controls && model.controls[step.key];
      this.formulaEl.textContent = c ? showVal(c.Caption !== undefined ? c.Caption : c.Text) : "";
    } else {
      this.refEl.textContent = step.proc || model.sheet || "";
      this.formulaEl.textContent = "";
    }
    const waitingClick = this.liveForm && this.session && this.session.status === "idle" && idx === this.steps.length - 1;
    if (isLast) {
      this.statusEl.textContent =
        `✅ 実行おわり（全 ${this.steps.length} ステップ）｜ ${diffSet.size} 個のセルが変化 — ${step.desc}`;
    } else if (waitingClick) {
      this.statusEl.textContent = `🖱 フォームのボタンを押してみよう（全 ${this.steps.length} ステップ）— ${step.desc}`;
    } else {
      this.statusEl.textContent = `ステップ ${idx + 1} / ${this.steps.length}: ${step.desc}`;
    }
    if (this.onStepApplied) this.onStepApplied(step);
  }

  // タイムライン（スライダー・カウンタ）へ現在位置を通知
  notifyStepChange() {
    if (this.onStepChange) this.onStepChange(this.cursor, this.steps.length);
  }

  // 1つ進める。done: MsgBox のダイアログを閉じたあとに呼ばれる
  stepForward(interactive, done) {
    if (this.cursor >= this.steps.length) return false;
    const step = this.steps[this.cursor];
    this.applyStep(this.cursor, interactive);
    this.cursor++;
    this.notifyStepChange();
    if (interactive && step.scope === "msg" && this.onMsgStep) {
      this.onMsgStep(step, done || (() => {}));
      return "dialog";
    }
    return true;
  }

  stepBackward() {
    if (this.cursor <= 0) return false;
    this.goToStep(this.cursor - 1);
    return true;
  }

  // cursor: 0 = 実行前の状態, n = n ステップ実行後の状態
  goToStep(cursor) {
    this.stop();
    const n = Math.max(0, Math.min(cursor, this.steps.length));
    this.cursor = n;
    if (n === 0) {
      this.hideMsgBox();
      this.renderModel(this.initialModel(), null, null);
      this.refEl.textContent = "A1";
      this.formulaEl.textContent = "";
      this.statusEl.textContent = this.steps.length
        ? `⏱ 実行前の状態（全 ${this.steps.length} ステップ）`
        : this.idleText;
      if (this.onStepApplied) this.onStepApplied(null);
    } else {
      this.applyStep(n - 1, false);
    }
    this.notifyStepChange();
  }

  // 最後まで一気に進めて結果を見る
  goToEnd() {
    if (!this.steps.length) return;
    this.goToStep(this.steps.length);
  }

  // 再生。restart: 最初から（最後まで見終わっているときも最初から）
  play(restart) {
    this.stop();
    if (restart || this.cursor >= this.steps.length) {
      const waiting = this.session && this.session.status === "waiting";
      if (restart || !waiting) this.cursor = 0;
    }
    this.playing = true;
    this.table.classList.add("playing");
    if (this.onPlayState) this.onPlayState(true);
    this.tick();
  }

  tick() {
    if (!this.playing) return;
    this.timer = null;
    if (this.cursor >= this.steps.length) {
      // 記録の最後まで来た。答え待ち（InputBox など）なら、答えのあと続きを再生する
      const handled = this.onReachEnd ? this.onReachEnd(() => this.continuePlay()) : false;
      if (!handled) this.stop();
      return;
    }
    const r = this.stepForward(true, () => this.continuePlay());
    if (r === "dialog") return; // MsgBox のダイアログを閉じるまで待つ
    this.timer = setTimeout(() => this.tick(), this.speed);
  }

  continuePlay() {
    if (!this.playing) return;
    this.stepsAppended();
    if (this.timer) clearTimeout(this.timer);
    this.timer = setTimeout(() => this.tick(), Math.min(this.speed, 250));
  }

  stop() {
    if (this.timer) {
      clearTimeout(this.timer);
      this.timer = null;
    }
    const wasPlaying = this.playing;
    this.playing = false;
    this.table.classList.remove("playing");
    if (wasPlaying && this.onPlayState) this.onPlayState(false);
  }

  reset() {
    // 実行開始時の状態（初期データ）にきっちり戻す
    this.goToStep(0);
    this.statusEl.textContent = this.steps.length
      ? `準備完了（全 ${this.steps.length} ステップ）`
      : this.idleText;
  }

  setSpeed(ms) {
    this.speed = ms;
  }

  setZoom(percent) {
    const scale = percent / 100;
    this.table.style.transform = `scale(${scale})`;
    this.table.style.transformOrigin = "top left";
    const gridWrap = this.table.closest(".excel-grid-wrap");
    if (gridWrap) {
      gridWrap.style.height = `calc(100% * ${scale})`;
    }
  }

  // 最終結果のモデル（クリア判定用・書き換えてよい写し）。
  // 古い課題の check(model) のために model.arr（配列 arr の中身）も付ける
  finalModel() {
    const m = this.store ? this.store.modelCopy(this.steps.length - 1) : JSON.parse(JSON.stringify(this.initialModel()));
    m.arr = m.arrays && m.arrays.arr ? m.arrays.arr.data || {} : {};
    return m;
  }
}
