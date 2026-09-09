/* ===== 仮想 Excel プレビュー =====
 * Blockly のブロックツリーを「解釈実行」して、
 * セルがどう変化するかをステップごとに記録し、アニメーション表示する。
 *
 * ExcelModel  : セルの値・書式を保持するデータモデル
 * Interpreter : ブロックを walk して「ステップ列」を生成
 * ExcelView   : HTML テーブルへの描画とアニメーション
 */

const GRID_ROWS = 26; // 最小表示行数（Z行まで）
const GRID_COLS = 26; // 最小表示列数（Z列まで）
const MAX_GRID_ROWS = 26; // 安全上限（Z行）
const MAX_GRID_COLS = 26; // Z列
const MAX_STEPS = 5000; // ステップ爆発ガード
const MAX_ARRAY_DISPLAY = 30; // 配列ビジュアライザの表示上限

const COLOR_HEX = {
  RED: "#e74c3c",
  GREEN: "#217346",
  BLUE: "#3498db",
  YELLOW: "#f1c40f",
  WHITE: "#ffffff",
};

function colLetter(n) {
  // 1 -> A
  return String.fromCharCode(64 + n);
}
function colNum(letter) {
  return letter.toUpperCase().charCodeAt(0) - 64;
}
// FieldVariable から変数名（表示名）を取り出す
function varNameOf(block) {
  const f = block.getField("VAR");
  return f ? f.getText() : "x";
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

/* ---------- モデル ---------- */
class ExcelModel {
  constructor() {
    this.cells = {}; // "A1" -> { value, bg, bold }
    this.arr = {}; // 配列: index -> value
    this.sheet = "Sheet1"; // 現在のシート名
    this.sheets = ["Sheet1"]; // 存在するシート一覧
  }
  key(col, row) {
    return colLetter(col) + row;
  }
  set(addr, value) {
    const c = (this.cells[addr] = this.cells[addr] || {});
    c.value = value;
  }
  get(addr) {
    const c = this.cells[addr];
    return c && c.value !== undefined ? c.value : "";
  }
  setBg(addr, color) {
    const c = (this.cells[addr] = this.cells[addr] || {});
    c.bg = color;
  }
  setBold(addr) {
    const c = (this.cells[addr] = this.cells[addr] || {});
    c.bold = true;
  }
  clearContents(addr) {
    if (this.cells[addr]) this.cells[addr].value = "";
  }
  setArr(idx, value) {
    this.arr[idx] = value;
  }
  getArr(idx) {
    return this.arr[idx] !== undefined ? this.arr[idx] : "";
  }
  addSheet(name) {
    if (!this.sheets.includes(name)) this.sheets.push(name);
  }
  selectSheet(name) {
    this.sheet = name;
  }
}

/* ---------- インタプリタ ----------
 * ブロックを実行して steps[] を作る。
 * 各 step = { addr, type:'value'|'bg'|'bold'|'clear', model:<snapshot>, desc }
 */
class Interpreter {
  constructor(initialCells, opts) {
    opts = opts || {};
    // InputBox の答えキャッシュ（block.id -> 値）。
    // interactive=true（▶実行時）のときだけ prompt() で質問し、
    // 編集中の再構築では仮値 "?" を使ってダイアログを出さない。
    this.inputCache = opts.inputCache || {};
    this.interactive = !!opts.interactive;
    this.model = new ExcelModel();
    // 実行開始時の状態（値＋書式）を種付け。
    // ここが「毎回同じ出発点」になるので、実行の再現性はこの初期データで決まる。
    if (initialCells) {
      for (const addr in initialCells) {
        const c = initialCells[addr];
        if (!c) continue;
        if (c.value !== undefined && c.value !== "") this.model.set(addr, c.value);
        if (c.bg) this.model.setBg(addr, c.bg);
        if (c.bold) this.model.setBold(addr);
        if (c.fontSize || c.border) {
          const cell = (this.model.cells[addr] = this.model.cells[addr] || {});
          if (c.fontSize) cell.fontSize = c.fontSize;
          if (c.border) cell.border = true;
        }
      }
    }
    this.steps = [];
    this.i = 0; // ループカウンタ
    this.vars = {}; // 変数ストア: name -> value
  }

  snapshot() {
    return {
      cells: JSON.parse(JSON.stringify(this.model.cells)),
      arr: JSON.parse(JSON.stringify(this.model.arr)),
      vars: JSON.parse(JSON.stringify(this.vars)),
      sheet: this.model.sheet,
      sheets: this.model.sheets.slice(),
    };
  }

  // scope: 'cell' | 'array' | 'sheet'
  record(scope, key, desc) {
    if (this.steps.length >= MAX_STEPS) {
      const err = new Error("STEP_LIMIT");
      err.code = "STEP_LIMIT";
      throw err;
    }
    this.steps.push({ scope, key, model: this.snapshot(), desc });
  }

  // セルアドレスの {i} 置換
  resolveAddr(addr) {
    if (/\{i\}/.test(addr)) {
      // "C{i}" 形式
      return addr.replace(/\{i\}/g, this.i);
    }
    return addr;
  }

  run(topBlock) {
    let block = topBlock;
    while (block) {
      this.execBlock(block);
      block = block.getNextBlock();
    }
  }

  execBlock(block) {
    switch (block.type) {
      case "cell_set_value": {
        const addr = this.resolveAddr(block.getFieldValue("CELL"));
        const val = this.evalValue(block.getInputTargetBlock("VALUE"));
        this.model.set(addr, val);
        this.record("cell", addr, `${addr} に「${val}」を入力`);
        break;
      }
      case "cell_copy": {
        const from = this.resolveAddr(block.getFieldValue("FROM"));
        const to = this.resolveAddr(block.getFieldValue("TO"));
        this.model.set(to, this.model.get(from));
        this.record("cell", to, `${from} を ${to} にコピー`);
        break;
      }
      case "cell_clear": {
        const addr = this.resolveAddr(block.getFieldValue("CELL"));
        this.model.clearContents(addr);
        this.record("cell", addr, `${addr} をクリア`);
        break;
      }
      case "array_set": {
        const idx = this.evalValue(block.getInputTargetBlock("INDEX"));
        const val = this.evalValue(block.getInputTargetBlock("VALUE"));
        this.model.setArr(idx, val);
        this.record("array", idx, `配列の ${idx} 番目に「${val}」を入れる`);
        break;
      }
      case "sheet_add": {
        const name = block.getFieldValue("NAME");
        this.model.addSheet(name);
        this.record("sheet", name, `シート「${name}」を追加`);
        break;
      }
      case "sheet_select": {
        const name = block.getFieldValue("NAME");
        this.model.selectSheet(name);
        this.record("sheet", name, `シート「${name}」に切り替え`);
        break;
      }
      case "loop_repeat": {
        const times = Number(this.evalValue(block.getInputTargetBlock("TIMES"))) || 0;
        const saved = this.i;
        for (let k = 1; k <= times; k++) {
          this.i = k;
          this.run(block.getInputTargetBlock("DO"));
          if (this.steps.length >= MAX_STEPS) break;
        }
        this.i = saved;
        break;
      }
      case "loop_range": {
        const start = Number(this.evalValue(block.getInputTargetBlock("START")));
        const end = Number(this.evalValue(block.getInputTargetBlock("END")));
        const saved = this.i;
        for (let k = start; k <= end; k++) {
          this.i = k;
          this.run(block.getInputTargetBlock("DO"));
          if (this.steps.length >= MAX_STEPS) break;
        }
        this.i = saved;
        break;
      }
      case "loop_for_step": {
        const start = Number(this.evalValue(block.getInputTargetBlock("START")));
        const end = Number(this.evalValue(block.getInputTargetBlock("END")));
        const step = Number(this.evalValue(block.getInputTargetBlock("STEP"))) || 1;
        const saved = this.i;
        if (step > 0) {
          for (let k = start; k <= end; k += step) {
            this.i = k;
            this.run(block.getInputTargetBlock("DO"));
            if (this.steps.length >= MAX_STEPS) break;
          }
        } else if (step < 0) {
          for (let k = start; k >= end; k += step) {
            this.i = k;
            this.run(block.getInputTargetBlock("DO"));
            if (this.steps.length >= MAX_STEPS) break;
          }
        }
        this.i = saved;
        break;
      }
      case "loop_while": {
        const saved = this.i;
        let guard = 0;
        while (this.evalValue(block.getInputTargetBlock("CONDITION"))) {
          if (++guard > MAX_STEPS || this.steps.length >= MAX_STEPS) break;
          this.run(block.getInputTargetBlock("DO"));
        }
        this.i = saved;
        break;
      }
      case "loop_do_until": {
        const saved = this.i;
        let guard = 0;
        while (!this.evalValue(block.getInputTargetBlock("CONDITION"))) {
          if (++guard > MAX_STEPS || this.steps.length >= MAX_STEPS) break;
          this.run(block.getInputTargetBlock("DO"));
        }
        this.i = saved;
        break;
      }
      case "io_msgbox": {
        const msg = String(this.evalValue(block.getInputTargetBlock("MSG")));
        this.record("msg", msg, `📢 「${msg}」を表示`);
        break;
      }
      case "range_border": {
        const r1 = Number(this.evalValue(block.getInputTargetBlock("R1"))) || 1;
        const c1 = Number(this.evalValue(block.getInputTargetBlock("C1"))) || 1;
        const r2 = Number(this.evalValue(block.getInputTargetBlock("R2"))) || 1;
        const c2 = Number(this.evalValue(block.getInputTargetBlock("C2"))) || 1;
        for (let r = r1; r <= r2; r++) {
          for (let c = c1; c <= c2; c++) {
            const addr = colLetter(c) + r;
            const cell = (this.model.cells[addr] = this.model.cells[addr] || {});
            cell.border = true;
          }
        }
        this.record("cell", colLetter(c1) + r1, `行${r1}列${c1}〜行${r2}列${c2} 罫線`);
        break;
      }
      case "range_select": {
        const r1 = Number(this.evalValue(block.getInputTargetBlock("R1"))) || 1;
        const c1 = Number(this.evalValue(block.getInputTargetBlock("C1"))) || 1;
        const r2 = Number(this.evalValue(block.getInputTargetBlock("R2"))) || 1;
        const c2 = Number(this.evalValue(block.getInputTargetBlock("C2"))) || 1;
        this.record("cell", colLetter(c1) + r1, `行${r1}列${c1}〜行${r2}列${c2} を選択`);
        break;
      }
      case "cond_if": {
        const cond = this.evalValue(block.getInputTargetBlock("CONDITION"));
        if (cond) {
          this.run(block.getInputTargetBlock("THEN"));
        } else {
          this.run(block.getInputTargetBlock("ELSE"));
        }
        break;
      }
      case "fmt_bgcolor": {
        const addr = this.resolveAddr(block.getFieldValue("CELL"));
        const color = block.getFieldValue("COLOR");
        this.model.setBg(addr, COLOR_HEX[color]);
        this.record("cell", addr, `${addr} の背景色を変更`);
        break;
      }
      case "fmt_bold": {
        const addr = this.resolveAddr(block.getFieldValue("CELL"));
        this.model.setBold(addr);
        this.record("cell", addr, `${addr} を太字に`);
        break;
      }
      case "fmt_fontsize": {
        const addr = this.resolveAddr(block.getFieldValue("CELL"));
        const size = Number(this.evalValue(block.getInputTargetBlock("SIZE"))) || 14;
        const c = (this.model.cells[addr] = this.model.cells[addr] || {});
        c.fontSize = size;
        this.record("cell", addr, `${addr} の文字サイズを ${size} に`);
        break;
      }
      case "cell_clear_row": {
        const row = Number(this.evalValue(block.getInputTargetBlock("ROW"))) || 1;
        for (let col = 1; col <= 30; col++) {
          const addr = String.fromCharCode(64 + col) + row;
          this.model.clearContents(addr);
        }
        this.record("cell", "A" + row, `${row} 行目をクリア`);
        break;
      }
      case "var_set": {
        const name = varNameOf(block);
        const val = this.evalValue(block.getInputTargetBlock("VALUE"));
        this.vars[name] = val;
        this.record("var", name, `変数「${name}」に「${val}」を入れる`);
        break;
      }
      case "var_change": {
        const name = varNameOf(block);
        const delta = Number(this.evalValue(block.getInputTargetBlock("DELTA"))) || 0;
        const cur = Number(this.vars[name]) || 0;
        this.vars[name] = cur + delta;
        this.record("var", name, `変数「${name}」を ${delta} ふやす（→ ${this.vars[name]}）`);
        break;
      }
      case "cells_set_value": {
        const row = Number(this.evalValue(block.getInputTargetBlock("ROW"))) || 1;
        const col = Number(this.evalValue(block.getInputTargetBlock("COL"))) || 1;
        const addr = colLetter(col) + row;
        const val = this.evalValue(block.getInputTargetBlock("VALUE"));
        this.model.set(addr, val);
        this.record("cell", addr, `Cells(${row},${col}) に「${val}」を入力`);
        break;
      }
      case "cells_clear": {
        const row = Number(this.evalValue(block.getInputTargetBlock("ROW"))) || 1;
        const col = Number(this.evalValue(block.getInputTargetBlock("COL"))) || 1;
        const addr = colLetter(col) + row;
        this.model.clearContents(addr);
        this.record("cell", addr, `Cells(${row},${col}) をクリア`);
        break;
      }
      case "cells_bgcolor": {
        const row = Number(this.evalValue(block.getInputTargetBlock("ROW"))) || 1;
        const col = Number(this.evalValue(block.getInputTargetBlock("COL"))) || 1;
        const addr = colLetter(col) + row;
        const color = block.getFieldValue("COLOR");
        this.model.setBg(addr, COLOR_HEX[color]);
        this.record("cell", addr, `Cells(${row},${col}) 背景色変更`);
        break;
      }
      case "cells_bold": {
        const row = Number(this.evalValue(block.getInputTargetBlock("ROW"))) || 1;
        const col = Number(this.evalValue(block.getInputTargetBlock("COL"))) || 1;
        const addr = colLetter(col) + row;
        this.model.setBold(addr);
        this.record("cell", addr, `Cells(${row},${col}) を太字に`);
        break;
      }
    }
  }

  // 値ブロックを評価
  evalValue(block) {
    if (!block) return "";
    switch (block.type) {
      case "value_number":
        return Number(block.getFieldValue("NUM"));
      case "value_text":
        return block.getFieldValue("TEXT");
      case "loop_index":
        return this.i;
      case "cell_get_value":
        return this.model.get(this.resolveAddr(block.getFieldValue("CELL")));
      case "array_get": {
        const idx = this.evalValue(block.getInputTargetBlock("INDEX"));
        return this.model.getArr(idx);
      }
      case "value_math": {
        const a = Number(this.evalValue(block.getInputTargetBlock("A"))) || 0;
        const b = Number(this.evalValue(block.getInputTargetBlock("B"))) || 0;
        const op = block.getFieldValue("OP");
        if (op === "+") return a + b;
        if (op === "-") return a - b;
        if (op === "*") return a * b;
        if (op === "mod") return b ? a % b : 0;
        return b ? a / b : 0;
      }
      case "cond_compare": {
        const a = this.evalValue(block.getInputTargetBlock("A"));
        const b = this.evalValue(block.getInputTargetBlock("B"));
        const op = block.getFieldValue("OP");
        switch (op) {
          case "=": return a == b;
          case "<>": return a != b;
          case "<": return a < b;
          case "<=": return a <= b;
          case ">": return a > b;
          case ">=": return a >= b;
        }
        return false;
      }
      case "cond_is_even": {
        const num = Number(this.evalValue(block.getInputTargetBlock("NUM")));
        return num % 2 === 0;
      }
      case "var_get": {
        const name = varNameOf(block);
        return this.vars[name] !== undefined ? this.vars[name] : 0;
      }
      case "cells_get_value": {
        const row = Number(this.evalValue(block.getInputTargetBlock("ROW"))) || 1;
        const col = Number(this.evalValue(block.getInputTargetBlock("COL"))) || 1;
        return this.model.get(colLetter(col) + row);
      }
      case "io_inputbox": {
        const key = block.id;
        if (this.inputCache[key] !== undefined) return this.inputCache[key];
        if (!this.interactive) return "?"; // 編集中はダイアログを出さない
        const promptText = block.getFieldValue("PROMPT");
        const raw = window.prompt(promptText);
        const str = raw === null ? "" : raw;
        const num = Number(str);
        const val = str.trim() !== "" && !isNaN(num) ? num : str;
        this.inputCache[key] = val; // 同じ実行中は1回の答えを使い回す
        return val;
      }
      case "text_concat": {
        const a = this.evalValue(block.getInputTargetBlock("A"));
        const b = this.evalValue(block.getInputTargetBlock("B"));
        return String(a) + String(b);
      }
      case "value_date": {
        return new Date().toLocaleDateString("ja-JP");
      }
      case "math_round": {
        const val = Number(this.evalValue(block.getInputTargetBlock("VALUE"))) || 0;
        const digits = Number(this.evalValue(block.getInputTargetBlock("DIGITS"))) || 0;
        const factor = Math.pow(10, digits);
        return Math.round(val * factor) / factor;
      }
    }
    return "";
  }
}

/* ---------- ビュー ---------- */
const EMPTY_MODEL = { cells: {}, arr: {}, vars: {}, sheet: "Sheet1", sheets: ["Sheet1"] };

class ExcelView {
  constructor(els) {
    this.table = els.table;
    this.refEl = els.ref;
    this.formulaEl = els.formula;
    this.statusEl = els.status;
    this.arrayEl = els.array; // 配列ビジュアライザのコンテナ
    this.varEl = els.vars || null; // 変数ウォッチのコンテナ
    this.msgOverlay = els.msgbox || null; // MsgBox オーバーレイ
    this.tabsEl = els.tabs; // シートタブのコンテナ
    this.onEdit = els.onEdit || null; // セル編集時のコールバック
    this.onPlayState = els.onPlayState || null; // 再生状態変更コールバック
    this.onStepChange = els.onStepChange || null; // カーソル移動時（タイムライン同期用）
    this.selected = "A1"; // キーボード操作の選択セル
    if (this.msgOverlay) {
      const ok = this.msgOverlay.querySelector(".msgbox-ok");
      if (ok) ok.addEventListener("click", () => this.hideMsgBox());
    }
    this.steps = [];
    this.cursor = 0;
    this.timer = null;
    this.speed = 500;
    this.rows = GRID_ROWS;
    this.cols = GRID_COLS;
    this.initialCells = {}; // 生徒が直接入力した初期データ
    this.editing = false; // 編集中フラグ（描画でclobberしない用）
    this.playing = false; // アニメーション中は編集不可
    this._workspace = null; // Blockly workspace への参照（常時変数表示用）
    this.buildGrid(GRID_ROWS, GRID_COLS);
    this.bindEditing();
  }

  // workspace をバインドして変数パネルを常時更新
  bindWorkspace(workspace) {
    this._workspace = workspace;
    workspace.addChangeListener((e) => {
      if (
        e.type === Blockly.Events.VAR_CREATE ||
        e.type === Blockly.Events.VAR_DELETE ||
        e.type === Blockly.Events.VAR_RENAME
      ) {
        this.renderVars({}, null);
      }
    });
    this.renderVars({}, null);
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
  // 初期データを1つのモデルにして表示用に
  initialModel() {
    return { cells: this.initialCells, arr: {}, vars: {}, sheet: "Sheet1", sheets: ["Sheet1"] };
  }

  // いま画面に映しているモデル（cursor は「次に実行するステップ」を指す）
  currentModel() {
    if (this.cursor > 0 && this.steps[this.cursor - 1]) {
      return this.steps[this.cursor - 1].model;
    }
    return this.initialModel();
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
          ].join("");
    const addrs = new Set([...Object.keys(base), ...Object.keys(cells)]);
    addrs.forEach((addr) => {
      if (sig(base[addr]) !== sig(cells[addr])) diff.add(addr);
    });
    return diff;
  }

  // 全ステップ＋初期データから必要なグリッドサイズを算出して再構築
  fitGrid(steps) {
    let maxRow = GRID_ROWS;
    let maxCol = GRID_COLS;
    const bump = (addr) => {
      const p = parseAddr(addr);
      if (!p) return;
      if (p.row > maxRow) maxRow = p.row;
      if (p.col > maxCol) maxCol = p.col;
    };
    for (const step of steps) {
      for (const addr in step.model.cells) bump(addr);
    }
    for (const addr in this.initialCells) bump(addr);
    maxRow = Math.min(maxRow, MAX_GRID_ROWS);
    maxCol = Math.min(maxCol, MAX_GRID_COLS);
    if (maxRow !== this.rows || maxCol !== this.cols) {
      this.buildGrid(maxRow, maxCol);
    }
  }

  // モデルを丸ごと描画
  // diffSet: 初期状態から変化したセル（結果表示時にハイライトする）
  renderModel(model, active, changed, diffSet) {
    const cells = model.cells || {};
    const base = this.initialCells || {};
    // セル
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
        td.style.color = cell.bg && cell.bg !== "#ffffff" ? "#fff" : "";
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
    this.renderArray(model.arr || {}, active);
    this.renderVars(model.vars || {}, active);
    this.renderTabs(model.sheets || ["Sheet1"], model.sheet || "Sheet1");
  }

  // 変数ウォッチ（📦 変数の名前と今の値を常時表示）
  renderVars(vars, active) {
    if (!this.varEl) return;
    const ws = this._workspace;
    // ワークスペースに定義された変数名を取得
    const declaredNames = ws
      ? ws.getAllVariables().map((v) => v.name)
      : Object.keys(vars);
    if (declaredNames.length === 0) {
      this.varEl.classList.remove("show");
      this.varEl.innerHTML = "";
      return;
    }
    this.varEl.classList.add("show");
    let html = '<span class="var-label">📦 変数</span><div class="var-chips">';
    declaredNames.forEach((n) => {
      const isActive = active && active.scope === "var" && active.key === n;
      const val = vars[n] !== undefined ? vars[n] : "—";
      html += `<div class="var-chip ${isActive ? "active" : ""}">
        <span class="var-name">${escapeHtmlPv(n)}</span>
        <span class="var-val">${escapeHtmlPv(val)}</span>
      </div>`;
    });
    html += "</div>";
    this.varEl.innerHTML = html;
  }

  // MsgBox オーバーレイ
  showMsgBox(text) {
    if (!this.msgOverlay) return;
    this.msgOverlay.querySelector(".msgbox-body").textContent = text;
    this.msgOverlay.hidden = false;
  }
  hideMsgBox() {
    if (this.msgOverlay) this.msgOverlay.hidden = true;
  }

  // 配列ビジュアライザ（実行中のみ表示）
  renderArray(arr, active) {
    if (!this.arrayEl) return;
    const keys = Object.keys(arr).map(Number).sort((a, b) => a - b);
    if (keys.length === 0) {
      this.arrayEl.classList.remove("show");
      this.arrayEl.innerHTML = "";
      return;
    }
    this.arrayEl.classList.add("show");
    const max = Math.max(...keys);
    const shown = Math.min(max, MAX_ARRAY_DISPLAY); // 表示上限ガード
    let html = '<span class="array-label">📦 配列 arr</span><div class="array-cells">';
    for (let i = 1; i <= shown; i++) {
      const v = arr[i] !== undefined ? arr[i] : "";
      const isActive = active && active.scope === "array" && Number(active.key) === i;
      html += `<div class="array-cell ${isActive ? "active-cell" : ""}">
        <div class="array-idx">${i}</div>
        <div class="array-val">${v}</div>
      </div>`;
    }
    if (max > shown) {
      html += `<div class="array-cell array-more">… +${max - shown}</div>`;
    }
    html += "</div>";
    this.arrayEl.innerHTML = html;
  }

  // シートタブ
  renderTabs(sheets, current) {
    if (!this.tabsEl) return;
    this.tabsEl.innerHTML = sheets
      .map((s) => `<div class="sheet-tab ${s === current ? "active" : ""}">${s}</div>`)
      .join("");
  }

  load(steps, limitHit) {
    this.stop();
    this.steps = steps;
    this.cursor = 0;
    this.hideMsgBox();
    this.fitGrid(steps);
    // アイドル状態では生徒が入力した初期データを表示
    this.renderModel(this.initialModel(), null, null);
    this.refEl.textContent = "A1";
    this.formulaEl.textContent = "";
    if (limitHit) {
      this.statusEl.textContent =
        `⚠️ 処理が多すぎます（${MAX_STEPS}ステップで打ち切り）。くり返し回数を減らしてください`;
    } else {
      this.statusEl.textContent = steps.length
        ? `準備完了（全 ${steps.length} ステップ）`
        : "セルに値を入力したり、ブロックを組み立てよう";
    }
    this.notifyStepChange();
  }

  applyStep(idx) {
    const step = this.steps[idx];
    if (!step) return;
    const marker = { scope: step.scope, key: step.key };
    // 最後のステップ＝実行結果。初期状態から変わったセルをまとめて光らせる
    const isLast = idx === this.steps.length - 1;
    const diffSet = isLast ? this.diffFromInitial(step.model) : null;
    this.renderModel(step.model, marker, marker, diffSet);
    // MsgBox ステップのときだけオーバーレイ表示
    if (step.scope === "msg") {
      this.showMsgBox(step.key);
    } else {
      this.hideMsgBox();
    }
    // フォーミュラバー表示
    if (step.scope === "cell") {
      this.refEl.textContent = step.key;
      const cell = step.model.cells[step.key];
      this.formulaEl.textContent = cell && cell.value !== undefined ? cell.value : "";
      const td = document.getElementById("cell-" + step.key);
      if (td) td.scrollIntoView({ block: "nearest", inline: "nearest" });
    } else if (step.scope === "array") {
      this.refEl.textContent = `arr(${step.key})`;
      this.formulaEl.textContent = step.model.arr[step.key];
    } else if (step.scope === "var") {
      this.refEl.textContent = step.key;
      this.formulaEl.textContent = step.model.vars ? step.model.vars[step.key] : "";
    } else if (step.scope === "msg") {
      this.refEl.textContent = "MsgBox";
      this.formulaEl.textContent = step.key;
    } else {
      this.refEl.textContent = step.model.sheet;
      this.formulaEl.textContent = "";
    }
    if (isLast) {
      this.statusEl.textContent =
        `✅ 実行おわり（全 ${this.steps.length} ステップ）｜ ${diffSet.size} 個のセルが変化 — ${step.desc}`;
    } else {
      this.statusEl.textContent = `ステップ ${idx + 1} / ${this.steps.length}: ${step.desc}`;
    }
  }

  // タイムライン（スライダー・カウンタ）へ現在位置を通知
  notifyStepChange() {
    if (this.onStepChange) this.onStepChange(this.cursor, this.steps.length);
  }

  stepForward() {
    if (this.cursor >= this.steps.length) {
      // 最終ステップの結果メッセージ（変化したセル数）を消さずに残す
      return false;
    }
    this.applyStep(this.cursor);
    this.cursor++;
    this.notifyStepChange();
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
        : "セルに値を入力したり、ブロックを組み立てよう";
    } else {
      this.applyStep(n - 1);
    }
    this.notifyStepChange();
  }

  // 最後まで一気に進めて結果を見る
  goToEnd() {
    if (!this.steps.length) return;
    this.goToStep(this.steps.length);
  }

  play() {
    this.stop();
    if (this.cursor >= this.steps.length) this.cursor = 0;
    this.playing = true;
    this.table.classList.add("playing");
    if (this.onPlayState) this.onPlayState(true);
    const tick = () => {
      if (!this.stepForward()) {
        this.stop();
        return;
      }
      this.timer = setTimeout(tick, this.speed);
    };
    tick();
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
      : "セルに値を入力したり、ブロックを組み立てよう";
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

  // 最終結果のモデル（クリア判定用）。ブロックが無くても初期データを返す
  finalModel() {
    return this.steps.length
      ? this.steps[this.steps.length - 1].model
      : this.initialModel();
  }
}

// ワークスペースからステップ列を生成
// 戻り値: { steps, limitHit }
// opts: { inputCache, interactive } — InputBox の質問制御（▶実行時のみ interactive）
function buildSteps(workspace, initialCells, opts) {
  const interp = new Interpreter(initialCells, opts);
  let limitHit = false;
  try {
    const topBlocks = workspace.getTopBlocks(true);
    for (const block of topBlocks) {
      // 出力ブロック（値ブロック単体）は実行対象外。文ブロックの先頭のみ実行
      if (block.outputConnection) continue;
      interp.run(block);
    }
  } catch (e) {
    if (e && e.code === "STEP_LIMIT") {
      limitHit = true; // それまでのステップは保持
    } else {
      throw e;
    }
  }
  return { steps: interp.steps, limitHit };
}
