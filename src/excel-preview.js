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
    return JSON.parse(JSON.stringify(this.model.cells));
  }

  record(addr, type, desc) {
    this.steps.push({ addr, type, model: this.snapshot(), desc });
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
        this.record(addr, "value", `${addr} に「${val}」を入力`);
        break;
      }
      case "cell_copy": {
        const from = this.resolveAddr(block.getFieldValue("FROM"));
        const to = this.resolveAddr(block.getFieldValue("TO"));
        this.model.set(to, this.model.get(from));
        this.record(to, "value", `${from} を ${to} にコピー`);
        break;
      }
      case "cell_clear": {
        const addr = this.resolveAddr(block.getFieldValue("CELL"));
        this.model.clearContents(addr);
        this.record(addr, "clear", `${addr} をクリア`);
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
        this.record(addr, "bg", `${addr} の背景色を変更`);
        break;
      }
      case "fmt_bold": {
        const addr = this.resolveAddr(block.getFieldValue("CELL"));
        this.model.setBold(addr);
        this.record(addr, "bold", `${addr} を太字に`);
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
class ExcelView {
  constructor(tableEl, refEl, formulaEl, statusEl) {
    this.table = tableEl;
    this.refEl = refEl;
    this.formulaEl = formulaEl;
    this.statusEl = statusEl;
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
  renderModel(cells, activeAddr, changedAddr) {
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
        if (addr === activeAddr) td.classList.add("active-cell");
        if (addr === changedAddr) td.classList.add("changed");
      }
    }
  }

  load(steps) {
    this.stop();
    this.steps = steps;
    this.cursor = 0;
    this.renderModel({}, null, null);
    this.refEl.textContent = "A1";
    this.formulaEl.textContent = "";
    this.statusEl.textContent = steps.length
      ? `準備完了（全 ${steps.length} ステップ）`
      : "ブロックがありません";
  }

  applyStep(idx) {
    const step = this.steps[idx];
    if (!step) return;
    this.renderModel(step.model, step.addr, step.addr);
    this.refEl.textContent = step.addr;
    const val = step.model[step.addr];
    this.formulaEl.textContent = val && val.value !== undefined ? val.value : "";
    this.statusEl.textContent = `ステップ ${idx + 1} / ${this.steps.length}: ${step.desc}`;
    // ハイライトしたセルへスクロール
    const td = document.getElementById("cell-" + step.addr);
    if (td) td.scrollIntoView({ block: "nearest", inline: "nearest" });
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
    this.renderModel({}, null, null);
    this.refEl.textContent = "A1";
    this.formulaEl.textContent = "";
    this.statusEl.textContent = this.steps.length
      ? `準備完了（全 ${this.steps.length} ステップ）`
      : "ブロックがありません";
  }

  setSpeed(ms) {
    this.speed = ms;
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
