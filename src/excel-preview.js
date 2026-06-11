/* ===== 仮想 Excel プレビュー =====
 * Blockly のブロックツリーを「解釈実行」して、
 * セルがどう変化するかをステップごとに記録し、アニメーション表示する。
 *
 * ExcelModel  : セルの値・書式を保持するデータモデル
 * Interpreter : ブロックを walk して「ステップ列」を生成
 * ExcelView   : HTML テーブルへの描画とアニメーション
 */

const GRID_ROWS = 12;
const GRID_COLS = 8; // A..H

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
  constructor() {
    this.model = new ExcelModel();
    this.steps = [];
    this.i = 0; // ループカウンタ
  }

  snapshot() {
    return {
      cells: JSON.parse(JSON.stringify(this.model.cells)),
      arr: JSON.parse(JSON.stringify(this.model.arr)),
      sheet: this.model.sheet,
      sheets: this.model.sheets.slice(),
    };
  }

  // scope: 'cell' | 'array' | 'sheet'
  record(scope, key, desc) {
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
        return op === "+" ? a + b : op === "-" ? a - b : op === "*" ? a * b : b ? a / b : 0;
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
    this.steps = [];
    this.cursor = 0;
    this.timer = null;
    this.speed = 500;
    this.buildGrid();
  }

  buildGrid() {
    let html = "<tr><th class='corner'></th>";
    for (let c = 1; c <= GRID_COLS; c++) html += `<th>${colLetter(c)}</th>`;
    html += "</tr>";
    for (let r = 1; r <= GRID_ROWS; r++) {
      html += `<tr><td class='row-num'>${r}</td>`;
      for (let c = 1; c <= GRID_COLS; c++) {
        html += `<td id='cell-${colLetter(c)}${r}'></td>`;
      }
      html += "</tr>";
    }
    this.table.innerHTML = html;
  }

  // モデルを丸ごと描画
  renderModel(model, active, changed) {
    const cells = model.cells || {};
    // セル
    for (let r = 1; r <= GRID_ROWS; r++) {
      for (let c = 1; c <= GRID_COLS; c++) {
        const addr = colLetter(c) + r;
        const td = document.getElementById("cell-" + addr);
        if (!td) continue;
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
    let html = '<span class="array-label">📦 配列 arr</span><div class="array-cells">';
    for (let i = 1; i <= max; i++) {
      const v = arr[i] !== undefined ? arr[i] : "";
      const isActive = active && active.scope === "array" && Number(active.key) === i;
      html += `<div class="array-cell ${isActive ? "active-cell" : ""}">
        <div class="array-idx">${i}</div>
        <div class="array-val">${v}</div>
      </div>`;
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

  load(steps) {
    this.stop();
    this.steps = steps;
    this.cursor = 0;
    this.renderModel(EMPTY_MODEL, null, null);
    this.refEl.textContent = "A1";
    this.formulaEl.textContent = "";
    this.statusEl.textContent = steps.length
      ? `準備完了（全 ${steps.length} ステップ）`
      : "ブロックがありません";
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
  }

  reset() {
    this.stop();
    this.cursor = 0;
    this.renderModel(EMPTY_MODEL, null, null);
    this.refEl.textContent = "A1";
    this.formulaEl.textContent = "";
    this.statusEl.textContent = this.steps.length
      ? `準備完了（全 ${this.steps.length} ステップ）`
      : "ブロックがありません";
  }

  setSpeed(ms) {
    this.speed = ms;
  }

  // 最終結果のモデル（クリア判定用）
  finalModel() {
    return this.steps.length ? this.steps[this.steps.length - 1].model : EMPTY_MODEL;
  }
}

// ワークスペースからステップ列を生成
function buildSteps(workspace) {
  const interp = new Interpreter();
  const topBlocks = workspace.getTopBlocks(true);
  for (const block of topBlocks) {
    // 出力ブロック（値ブロック単体）は実行対象外。文ブロックの先頭のみ実行
    if (block.outputConnection) continue;
    interp.run(block);
  }
  return interp.steps;
}
