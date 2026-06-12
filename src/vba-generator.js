/* ===== VBA コード生成器 =====
 * Blockly ワークスペースのブロックツリーを VBA コードに変換する。
 * Blockly 標準の CodeGenerator を継承して実装。
 */

const vbaGenerator = new Blockly.Generator("VBA");

// 演算子の優先順位（簡易: すべて括弧で囲むので 0 でOK）
vbaGenerator.ORDER_ATOMIC = 0;
vbaGenerator.ORDER_NONE = 99;

// 色名 → RGB
const COLOR_RGB = {
  RED: "RGB(231, 76, 60)",
  GREEN: "RGB(33, 115, 70)",
  BLUE: "RGB(52, 152, 219)",
  YELLOW: "RGB(241, 196, 15)",
  WHITE: "RGB(255, 255, 255)",
};

// インデント設定
vbaGenerator.INDENT = "    ";

vbaGenerator.scrub_ = function (block, code, thisOnly) {
  const next = block.nextConnection && block.nextConnection.targetBlock();
  let nextCode = "";
  if (next && !thisOnly) {
    nextCode = "\n" + vbaGenerator.blockToCode(next);
  }
  return code + nextCode;
};

// ----- セルアドレスを VBA の Range(...) 式に -----
// "A1"   -> Range("A1")
// "A{i}" -> Range("A" & i)   (動的な行番号)
function rangeExpr(addr) {
  if (/\{i\}/.test(addr)) {
    const col = addr.replace(/\{i\}/g, "");
    return `Range("${col}" & i)`;
  }
  return `Range("${addr}")`;
}

// ===== セル操作 =====

vbaGenerator.forBlock["cell_set_value"] = function (block) {
  const cell = block.getFieldValue("CELL");
  const value = vbaGenerator.valueToCode(block, "VALUE", vbaGenerator.ORDER_NONE) || '""';
  return `${rangeExpr(cell)}.Value = ${value}`;
};

vbaGenerator.forBlock["cell_get_value"] = function (block) {
  const cell = block.getFieldValue("CELL");
  return [`${rangeExpr(cell)}.Value`, vbaGenerator.ORDER_ATOMIC];
};

vbaGenerator.forBlock["cell_copy"] = function (block) {
  const from = block.getFieldValue("FROM");
  const to = block.getFieldValue("TO");
  return `${rangeExpr(from)}.Copy ${rangeExpr(to)}`;
};

vbaGenerator.forBlock["cell_clear"] = function (block) {
  const cell = block.getFieldValue("CELL");
  return `${rangeExpr(cell)}.ClearContents`;
};

vbaGenerator.forBlock["cell_clear_row"] = function (block) {
  const row = vbaGenerator.valueToCode(block, "ROW", vbaGenerator.ORDER_NONE) || "1";
  return `Rows(${row}).ClearContents`;
};

// ===== 繰り返し =====

vbaGenerator.forBlock["loop_repeat"] = function (block) {
  const times = vbaGenerator.valueToCode(block, "TIMES", vbaGenerator.ORDER_NONE) || "1";
  const body = vbaGenerator.statementToCode(block, "DO");
  return `For i = 1 To ${times}\n${body}\nNext i`;
};

vbaGenerator.forBlock["loop_range"] = function (block) {
  const start = vbaGenerator.valueToCode(block, "START", vbaGenerator.ORDER_NONE) || "1";
  const end = vbaGenerator.valueToCode(block, "END", vbaGenerator.ORDER_NONE) || "1";
  const body = vbaGenerator.statementToCode(block, "DO");
  return `For i = ${start} To ${end}\n${body}\nNext i`;
};

vbaGenerator.forBlock["loop_for_step"] = function (block) {
  const start = vbaGenerator.valueToCode(block, "START", vbaGenerator.ORDER_NONE) || "1";
  const end = vbaGenerator.valueToCode(block, "END", vbaGenerator.ORDER_NONE) || "1";
  const step = vbaGenerator.valueToCode(block, "STEP", vbaGenerator.ORDER_NONE) || "1";
  const body = vbaGenerator.statementToCode(block, "DO");
  return `For i = ${start} To ${end} Step ${step}\n${body}\nNext i`;
};

vbaGenerator.forBlock["loop_while"] = function (block) {
  const cond = vbaGenerator.valueToCode(block, "CONDITION", vbaGenerator.ORDER_NONE) || "False";
  const body = vbaGenerator.statementToCode(block, "DO");
  return `Do While ${cond}\n${body}\nLoop`;
};

vbaGenerator.forBlock["loop_index"] = function () {
  return ["i", vbaGenerator.ORDER_ATOMIC];
};

// ===== 条件分岐 =====

vbaGenerator.forBlock["cond_if"] = function (block) {
  const cond = vbaGenerator.valueToCode(block, "CONDITION", vbaGenerator.ORDER_NONE) || "False";
  const thenCode = vbaGenerator.statementToCode(block, "THEN");
  const elseCode = vbaGenerator.statementToCode(block, "ELSE");
  let code = `If ${cond} Then\n${thenCode}`;
  if (elseCode.trim()) {
    code += `\nElse\n${elseCode}`;
  }
  code += `\nEnd If`;
  return code;
};

vbaGenerator.forBlock["cond_compare"] = function (block) {
  const a = vbaGenerator.valueToCode(block, "A", vbaGenerator.ORDER_NONE) || "0";
  const op = block.getFieldValue("OP");
  const b = vbaGenerator.valueToCode(block, "B", vbaGenerator.ORDER_NONE) || "0";
  return [`${a} ${op} ${b}`, vbaGenerator.ORDER_ATOMIC];
};

vbaGenerator.forBlock["cond_is_even"] = function (block) {
  const num = vbaGenerator.valueToCode(block, "NUM", vbaGenerator.ORDER_NONE) || "0";
  return [`(${num} Mod 2 = 0)`, vbaGenerator.ORDER_ATOMIC];
};

// ===== 書式設定 =====

vbaGenerator.forBlock["fmt_bgcolor"] = function (block) {
  const cell = block.getFieldValue("CELL");
  const color = block.getFieldValue("COLOR");
  return `${rangeExpr(cell)}.Interior.Color = ${COLOR_RGB[color]}`;
};

vbaGenerator.forBlock["fmt_bold"] = function (block) {
  const cell = block.getFieldValue("CELL");
  return `${rangeExpr(cell)}.Font.Bold = True`;
};

vbaGenerator.forBlock["fmt_fontsize"] = function (block) {
  const cell = block.getFieldValue("CELL");
  const size = vbaGenerator.valueToCode(block, "SIZE", vbaGenerator.ORDER_NONE) || "14";
  return `${rangeExpr(cell)}.Font.Size = ${size}`;
};

// ===== Cells（行・列を変数で指定）=====

vbaGenerator.forBlock["cells_set_value"] = function (block) {
  const row = vbaGenerator.valueToCode(block, "ROW", vbaGenerator.ORDER_NONE) || "1";
  const col = vbaGenerator.valueToCode(block, "COL", vbaGenerator.ORDER_NONE) || "1";
  const val = vbaGenerator.valueToCode(block, "VALUE", vbaGenerator.ORDER_NONE) || '""';
  return `Cells(${row}, ${col}).Value = ${val}`;
};

vbaGenerator.forBlock["cells_get_value"] = function (block) {
  const row = vbaGenerator.valueToCode(block, "ROW", vbaGenerator.ORDER_NONE) || "1";
  const col = vbaGenerator.valueToCode(block, "COL", vbaGenerator.ORDER_NONE) || "1";
  return [`Cells(${row}, ${col}).Value`, vbaGenerator.ORDER_ATOMIC];
};

vbaGenerator.forBlock["cells_clear"] = function (block) {
  const row = vbaGenerator.valueToCode(block, "ROW", vbaGenerator.ORDER_NONE) || "1";
  const col = vbaGenerator.valueToCode(block, "COL", vbaGenerator.ORDER_NONE) || "1";
  return `Cells(${row}, ${col}).ClearContents`;
};

vbaGenerator.forBlock["cells_bgcolor"] = function (block) {
  const row = vbaGenerator.valueToCode(block, "ROW", vbaGenerator.ORDER_NONE) || "1";
  const col = vbaGenerator.valueToCode(block, "COL", vbaGenerator.ORDER_NONE) || "1";
  const color = block.getFieldValue("COLOR");
  return `Cells(${row}, ${col}).Interior.Color = ${COLOR_RGB[color]}`;
};

vbaGenerator.forBlock["cells_bold"] = function (block) {
  const row = vbaGenerator.valueToCode(block, "ROW", vbaGenerator.ORDER_NONE) || "1";
  const col = vbaGenerator.valueToCode(block, "COL", vbaGenerator.ORDER_NONE) || "1";
  return `Cells(${row}, ${col}).Font.Bold = True`;
};

vbaGenerator.forBlock["range_border"] = function (block) {
  const r1 = vbaGenerator.valueToCode(block, "R1", vbaGenerator.ORDER_NONE) || "1";
  const c1 = vbaGenerator.valueToCode(block, "C1", vbaGenerator.ORDER_NONE) || "1";
  const r2 = vbaGenerator.valueToCode(block, "R2", vbaGenerator.ORDER_NONE) || "1";
  const c2 = vbaGenerator.valueToCode(block, "C2", vbaGenerator.ORDER_NONE) || "1";
  return `Range(Cells(${r1}, ${c1}), Cells(${r2}, ${c2})).Borders.LineStyle = 1`;
};

vbaGenerator.forBlock["range_select"] = function (block) {
  const r1 = vbaGenerator.valueToCode(block, "R1", vbaGenerator.ORDER_NONE) || "1";
  const c1 = vbaGenerator.valueToCode(block, "C1", vbaGenerator.ORDER_NONE) || "1";
  const r2 = vbaGenerator.valueToCode(block, "R2", vbaGenerator.ORDER_NONE) || "1";
  const c2 = vbaGenerator.valueToCode(block, "C2", vbaGenerator.ORDER_NONE) || "1";
  return `Range(Cells(${r1}, ${c1}), Cells(${r2}, ${c2})).Select`;
};

// ===== 入出力 =====

vbaGenerator.forBlock["io_inputbox"] = function (block) {
  const prompt = block.getFieldValue("PROMPT");
  return [`InputBox("${prompt}")`, vbaGenerator.ORDER_ATOMIC];
};

vbaGenerator.forBlock["io_msgbox"] = function (block) {
  const msg = vbaGenerator.valueToCode(block, "MSG", vbaGenerator.ORDER_NONE) || '""';
  return `MsgBox ${msg}`;
};

vbaGenerator.forBlock["text_concat"] = function (block) {
  const a = vbaGenerator.valueToCode(block, "A", vbaGenerator.ORDER_NONE) || '""';
  const b = vbaGenerator.valueToCode(block, "B", vbaGenerator.ORDER_NONE) || '""';
  return [`${a} & ${b}`, vbaGenerator.ORDER_ATOMIC];
};

vbaGenerator.forBlock["value_date"] = function () {
  return ["Date", vbaGenerator.ORDER_ATOMIC];
};

vbaGenerator.forBlock["math_round"] = function (block) {
  const val = vbaGenerator.valueToCode(block, "VALUE", vbaGenerator.ORDER_NONE) || "0";
  const digits = vbaGenerator.valueToCode(block, "DIGITS", vbaGenerator.ORDER_NONE) || "0";
  return [`WorksheetFunction.Round(${val}, ${digits})`, vbaGenerator.ORDER_ATOMIC];
};

vbaGenerator.forBlock["loop_do_until"] = function (block) {
  const cond = vbaGenerator.valueToCode(block, "CONDITION", vbaGenerator.ORDER_NONE) || "False";
  const body = vbaGenerator.statementToCode(block, "DO");
  return `Do Until ${cond}\n${body}\nLoop`;
};

// ===== 変数 =====

// FieldVariable から変数名（表示名）を取り出す
function varName(block) {
  const f = block.getField("VAR");
  return f ? f.getText() : "x";
}

vbaGenerator.forBlock["var_set"] = function (block) {
  const name = varName(block);
  const val = vbaGenerator.valueToCode(block, "VALUE", vbaGenerator.ORDER_NONE) || "0";
  return `${name} = ${val}`;
};

vbaGenerator.forBlock["var_get"] = function (block) {
  return [varName(block), vbaGenerator.ORDER_ATOMIC];
};

vbaGenerator.forBlock["var_change"] = function (block) {
  const name = varName(block);
  const delta = vbaGenerator.valueToCode(block, "DELTA", vbaGenerator.ORDER_NONE) || "1";
  return `${name} = ${name} + ${delta}`;
};

// ===== 配列 =====

vbaGenerator.forBlock["array_set"] = function (block) {
  const idx = vbaGenerator.valueToCode(block, "INDEX", vbaGenerator.ORDER_NONE) || "1";
  const val = vbaGenerator.valueToCode(block, "VALUE", vbaGenerator.ORDER_NONE) || "0";
  return `arr(${idx}) = ${val}`;
};

vbaGenerator.forBlock["array_get"] = function (block) {
  const idx = vbaGenerator.valueToCode(block, "INDEX", vbaGenerator.ORDER_NONE) || "1";
  return [`arr(${idx})`, vbaGenerator.ORDER_ATOMIC];
};

// ===== シート操作 =====

vbaGenerator.forBlock["sheet_add"] = function (block) {
  const name = block.getFieldValue("NAME");
  return `Worksheets.Add.Name = "${name}"`;
};

vbaGenerator.forBlock["sheet_select"] = function (block) {
  const name = block.getFieldValue("NAME");
  return `Worksheets("${name}").Select`;
};

// ===== 値ブロック =====

vbaGenerator.forBlock["value_number"] = function (block) {
  return [String(block.getFieldValue("NUM")), vbaGenerator.ORDER_ATOMIC];
};

vbaGenerator.forBlock["value_text"] = function (block) {
  const text = block.getFieldValue("TEXT");
  return [`"${text}"`, vbaGenerator.ORDER_ATOMIC];
};

vbaGenerator.forBlock["value_math"] = function (block) {
  const a = vbaGenerator.valueToCode(block, "A", vbaGenerator.ORDER_NONE) || "0";
  const op = block.getFieldValue("OP");
  const b = vbaGenerator.valueToCode(block, "B", vbaGenerator.ORDER_NONE) || "0";
  const vbaOp = op === "mod" ? "Mod" : op;
  return [`(${a} ${vbaOp} ${b})`, vbaGenerator.ORDER_ATOMIC];
};

// ===== ワークスペース全体 → Sub MyMacro() でラップ =====
// 実際にブロックで使われている変数だけ収集（未使用変数はDimしない）
function collectVarNames(workspace) {
  const names = new Set();
  workspace.getAllBlocks(false).forEach((b) => {
    if (b.type === "var_set" || b.type === "var_get" || b.type === "var_change") {
      const f = b.getField("VAR");
      if (f) names.add(f.getText());
    }
  });
  return names;
}

function generateVBA(workspace) {
  let body = vbaGenerator.workspaceToCode(workspace);
  if (!body.trim()) {
    return "Sub MyMacro()\n    ' ブロックを組み立てるとここにコードが表示されます\nEnd Sub";
  }
  const usesI = /\bi\b/.test(body);
  const usesArr = /\barr\(/.test(body);
  const varNames = collectVarNames(workspace);
  const indented = body
    .split("\n")
    .map((line) => (line.trim() ? "    " + line : line))
    .join("\n");
  let header = "Sub MyMacro()\n";
  if (usesI) header += "    Dim i As Integer\n";
  if (usesArr) header += "    Dim arr(1 To 100) As Variant\n";
  varNames.forEach((n) => {
    header += `    Dim ${n} As Variant\n`;
  });
  return header + indented + "\nEnd Sub";
}
