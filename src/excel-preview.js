/* ===== 仮想 Excel プレビュー =====
 * Blockly のブロックツリーを「解釈実行」して、
 * セルがどう変化するかをステップごとに記録し、アニメーション表示する。
 *
 * ExcelModel  : セルの値・書式を保持するデータモデル
 * Interpreter : ブロックを walk して「ステップ列」を生成
 * ExcelView   : HTML テーブルへの描画とアニメーション
 */

const GRID_ROWS = 30; // 最小表示行数
const GRID_COLS = 30; // 最小表示列数 A..AD
const MAX_GRID_ROWS = 100; // 安全上限
const MAX_GRID_COLS = 30; // A..AD
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
  constructor(initialCells) {
    this.model = new ExcelModel();
    // 生徒が仮想Excelに直接入力した初期データを種付け
    if (initialCells) {
      for (const addr in initialCells) {
        const c = initialCells[addr];
        if (c && c.value !== undefined && c.value !== "") {
          this.model.set(addr, c.value);
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
        const times = Number(block.getFieldValue("TIMES"));
        const saved = this.i;
        for (let k = 1; k <= times; k++) {
          this.i = k;
          this.run(block.getInputTargetBlock("DO"));
        }
        this.i = saved;
        break;
      }
      case "loop_range": {
        const start = Number(block.getFieldValue("START"));
        const end = Number(block.getFieldValue("END"));
        const saved = this.i;
        for (let k = start; k <= end; k++) {
          this.i = k;
          this.run(block.getInputTargetBlock("DO"));
        }
        this.i = saved;
        break;
      }
      case "loop_for_step": {
        const start = Number(block.getFieldValue("START"));
        const end = Number(block.getFieldValue("END"));
        const step = Number(block.getFieldValue("STEP")) || 1;
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
        const size = Number(block.getFieldValue("SIZE"));
        const c = (this.model.cells[addr] = this.model.cells[addr] || {});
        c.fontSize = size;
        this.record("cell", addr, `${addr} の文字サイズを ${size} に`);
        break;
      }
      case "cell_clear_row": {
        const row = Number(block.getFieldValue("ROW"));
        for (let col = 1; col <= 30; col++) {
          const addr = String.fromCharCode(64 + col) + row;
          this.model.clearContents(addr);
        }
        this.record("cell", "A" + row, `${row} 行目をクリア`);
        break;
      }
      case "var_set": {
        const name = block.getFieldValue("NAME");
        const val = this.evalValue(block.getInputTargetBlock("VALUE"));
        this.vars[name] = val;
        this.record("var", name, `変数「${name}」に「${val}」を入れる`);
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
        const name = block.getFieldValue("NAME");
        return this.vars[name] !== undefined ? this.vars[name] : 0;
      }
      case "cells_get_value": {
        const row = Number(this.evalValue(block.getInputTargetBlock("ROW"))) || 1;
        const col = Number(this.evalValue(block.getInputTargetBlock("COL"))) || 1;
        return this.model.get(colLetter(col) + row);
      }
    }
    return "";
  }
}

/* ---------- ビュー ---------- */
const EMPTY_MODEL = { cells: {}, arr: {}, sheet: "Sheet1", sheets: ["Sheet1"] };

class ExcelView {
  constructor(els) {
    this.table = els.table;
    this.refEl = els.ref;
    this.formulaEl = els.formula;
    this.statusEl = els.status;
    this.arrayEl = els.array; // 配列ビジュアライザのコンテナ
    this.tabsEl = els.tabs; // シートタブのコンテナ
    this.onEdit = els.onEdit || null; // セル編集時のコールバック
    this.steps = [];
    this.cursor = 0;
    this.timer = null;
    this.speed = 500;
    this.rows = GRID_ROWS;
    this.cols = GRID_COLS;
    this.initialCells = {}; // 生徒が直接入力した初期データ
    this.editing = false; // 編集中フラグ（描画でclobberしない用）
    this.playing = false; // アニメーション中は編集不可
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
    this.table.addEventListener("dblclick", (e) => {
      const td = e.target.closest("td.editable");
      if (!td || this.playing) return;
      this.beginEdit(td);
    });
    // シングルクリックでも選択表示
    this.table.addEventListener("click", (e) => {
      const td = e.target.closest("td.editable");
      if (!td || this.playing || this.editing) return;
      const addr = td.dataset.addr;
      this.refEl.textContent = addr;
      const c = this.initialCells[addr];
      this.formulaEl.textContent = c && c.value !== undefined ? c.value : "";

      // Blocklyのテキストフィールドが編集中なら、そこにセルアドレスを入力
      const active = document.activeElement;
      if (active && active.classList.contains("blocklyHtmlInput")) {
        active.value = addr;
        active.dispatchEvent(new Event("input", { bubbles: true }));
        active.dispatchEvent(new KeyboardEvent("keydown", { key: "Enter", bubbles: true }));
      }
    });
  }

  beginEdit(td) {
    const addr = td.dataset.addr;
    this.editing = true;
    td.contentEditable = "true";
    td.classList.add("editing");
    // 現在の初期値を表示
    const cur = this.initialCells[addr];
    td.textContent = cur && cur.value !== undefined ? cur.value : "";
    td.focus();
    // テキスト全選択
    const range = document.createRange();
    range.selectNodeContents(td);
    const sel = window.getSelection();
    sel.removeAllRanges();
    sel.addRange(range);

    const commit = () => {
      td.contentEditable = "false";
      td.classList.remove("editing");
      this.editing = false;
      const raw = td.textContent.trim();
      this.setInitialCell(addr, raw);
      td.removeEventListener("blur", onBlur);
      td.removeEventListener("keydown", onKey);
      if (this.onEdit) this.onEdit();
    };
    const onBlur = () => commit();
    const onKey = (e) => {
      if (e.key === "Enter") {
        e.preventDefault();
        td.blur();
      } else if (e.key === "Escape") {
        e.preventDefault();
        td.textContent = cur && cur.value !== undefined ? cur.value : "";
        td.blur();
      }
    };
    td.addEventListener("blur", onBlur);
    td.addEventListener("keydown", onKey);
  }

  // 初期データを設定（数値ならNumber、空なら削除）
  setInitialCell(addr, raw) {
    if (raw === "") {
      delete this.initialCells[addr];
      return;
    }
    const num = Number(raw);
    const value = raw !== "" && !isNaN(num) ? num : raw;
    this.initialCells[addr] = { value };
  }

  setInitialCells(obj) {
    this.initialCells = obj ? JSON.parse(JSON.stringify(obj)) : {};
  }
  getInitialCells() {
    return this.initialCells;
  }
  // 初期データを1つのモデルにして表示用に
  initialModel() {
    return { cells: this.initialCells, arr: {}, sheet: "Sheet1", sheets: ["Sheet1"] };
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
  renderModel(model, active, changed) {
    const cells = model.cells || {};
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
        td.style.color = cell.bg && cell.bg !== "#ffffff" ? "#fff" : "#000";
        td.classList.remove("active-cell", "changed");
        if (active && active.scope === "cell" && addr === active.key) td.classList.add("active-cell");
        if (changed && changed.scope === "cell" && addr === changed.key) td.classList.add("changed");
      }
    }
    this.renderArray(model.arr || {}, active);
    this.renderTabs(model.sheets || ["Sheet1"], model.sheet || "Sheet1");
  }

  // 配列ビジュアライザ
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
  }

  applyStep(idx) {
    const step = this.steps[idx];
    if (!step) return;
    const marker = { scope: step.scope, key: step.key };
    this.renderModel(step.model, marker, marker);
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
    } else {
      this.refEl.textContent = step.model.sheet;
      this.formulaEl.textContent = "";
    }
    this.statusEl.textContent = `ステップ ${idx + 1} / ${this.steps.length}: ${step.desc}`;
  }

  stepForward() {
    if (this.cursor >= this.steps.length) {
      this.statusEl.textContent = "✅ 完了！ ↺ リセットで最初から";
      return false;
    }
    this.applyStep(this.cursor);
    this.cursor++;
    return true;
  }

  play() {
    this.stop();
    if (this.cursor >= this.steps.length) this.cursor = 0;
    this.playing = true;
    this.table.classList.add("playing");
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
    this.playing = false;
    this.table.classList.remove("playing");
  }

  reset() {
    this.stop();
    this.cursor = 0;
    // リセットで初期データに戻す
    this.renderModel(this.initialModel(), null, null);
    this.refEl.textContent = "A1";
    this.formulaEl.textContent = "";
    this.statusEl.textContent = this.steps.length
      ? `準備完了（全 ${this.steps.length} ステップ）`
      : "セルに値を入力したり、ブロックを組み立てよう";
  }

  setSpeed(ms) {
    this.speed = ms;
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
function buildSteps(workspace, initialCells) {
  const interp = new Interpreter(initialCells);
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
