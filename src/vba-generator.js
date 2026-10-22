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

// どこにもつながっていない値ブロックは、コードに出さない（出すと文として読めずエラーになる）
vbaGenerator.scrubNakedValue = function () {
  return "";
};

// 文字列リテラル（" は "" にする）
function vbaStr(s) {
  return '"' + String(s === undefined || s === null ? "" : s).replace(/"/g, '""') + '"';
}

// ----- くり返しのカウンタ -----
// 入れ子の深さに応じて i → j → k … を使う（内側のくり返しが外側の i を書き換えないように）。
// 「カウンタ」ブロックとセル番地の {i} は、一番内側のくり返しのカウンタを指す
const LOOP_COUNTERS = ["i", "j", "k", "l", "m"];
vbaGenerator._loopDepth = 0;
vbaGenerator._usedCounters = new Set();
function currentCounter() {
  const name = LOOP_COUNTERS[Math.max(0, vbaGenerator._loopDepth - 1)] || "i";
  vbaGenerator._usedCounters.add(name);
  return name;
}
// くり返しの中身を1段深いカウンタで生成する。戻り値: [カウンタ名, 中身のコード]
function loopBody(block, inputName) {
  vbaGenerator._loopDepth++;
  try {
    const name = currentCounter();
    return [name, vbaGenerator.statementToCode(block, inputName)];
  } finally {
    vbaGenerator._loopDepth--;
  }
}

// ----- セルアドレスを VBA の Range(...) 式に -----
// "A1"   -> Range("A1")
// "A{i}" -> Range("A" & i)   (動的な行番号。入れ子なら一番内側のカウンタ)
function rangeExpr(addr) {
  if (/\{i\}/.test(addr)) {
    const col = addr.replace(/\{i\}/g, "");
    return `Range(${vbaStr(col)} & ${currentCounter()})`;
  }
  return `Range(${vbaStr(addr)})`;
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
  const [c, body] = loopBody(block, "DO");
  return `For ${c} = 1 To ${times}\n${body}\nNext ${c}`;
};

vbaGenerator.forBlock["loop_range"] = function (block) {
  const start = vbaGenerator.valueToCode(block, "START", vbaGenerator.ORDER_NONE) || "1";
  const end = vbaGenerator.valueToCode(block, "END", vbaGenerator.ORDER_NONE) || "1";
  const [c, body] = loopBody(block, "DO");
  return `For ${c} = ${start} To ${end}\n${body}\nNext ${c}`;
};

vbaGenerator.forBlock["loop_for_step"] = function (block) {
  const start = vbaGenerator.valueToCode(block, "START", vbaGenerator.ORDER_NONE) || "1";
  const end = vbaGenerator.valueToCode(block, "END", vbaGenerator.ORDER_NONE) || "1";
  const step = vbaGenerator.valueToCode(block, "STEP", vbaGenerator.ORDER_NONE) || "1";
  const [c, body] = loopBody(block, "DO");
  return `For ${c} = ${start} To ${end} Step ${step}\n${body}\nNext ${c}`;
};

vbaGenerator.forBlock["loop_while"] = function (block) {
  const cond = vbaGenerator.valueToCode(block, "CONDITION", vbaGenerator.ORDER_NONE) || "False";
  const body = vbaGenerator.statementToCode(block, "DO");
  return `Do While ${cond}\n${body}\nLoop`;
};

vbaGenerator.forBlock["loop_index"] = function () {
  return [currentCounter(), vbaGenerator.ORDER_ATOMIC];
};

// ===== 条件分岐 =====

vbaGenerator.forBlock["cond_if"] = function (block) {
  const cond = vbaGenerator.valueToCode(block, "CONDITION", vbaGenerator.ORDER_NONE) || "False";
  const thenCode = vbaGenerator.statementToCode(block, "THEN");
  const elseCode = vbaGenerator.statementToCode(block, "ELSE");
  let code = `If ${cond} Then\n${thenCode}`;
  for (let i = 0; block.getInput("ELSEIF_COND_" + i); i++) {
    const c = vbaGenerator.valueToCode(block, "ELSEIF_COND_" + i, vbaGenerator.ORDER_NONE) || "False";
    code += `\nElseIf ${c} Then\n${vbaGenerator.statementToCode(block, "ELSEIF_DO_" + i)}`;
  }
  if (elseCode.trim()) {
    code += `\nElse\n${elseCode}`;
  }
  code += `\nEnd If`;
  return code;
};

// Select Case：Case の行は1段、中身は2段字下げする
vbaGenerator.forBlock["cond_select"] = function (block) {
  const subject = vbaGenerator.valueToCode(block, "SUBJECT", vbaGenerator.ORDER_NONE) || "0";
  const deeper = (code) =>
    code
      .split("\n")
      .map((l) => (l.trim() ? vbaGenerator.INDENT + l : l))
      .join("\n");
  let code = `Select Case ${subject}`;
  for (let i = 0; block.getInput("CASE_DO_" + i); i++) {
    const tests = String(block.getFieldValue("CASE_" + i) || "").trim() || "0";
    code += `\n${vbaGenerator.INDENT}Case ${tests}\n${deeper(vbaGenerator.statementToCode(block, "CASE_DO_" + i))}`;
  }
  const elseCode = vbaGenerator.statementToCode(block, "ELSE");
  if (elseCode.trim()) code += `\n${vbaGenerator.INDENT}Case Else\n${deeper(elseCode)}`;
  return code + "\nEnd Select";
};

vbaGenerator.forBlock["cond_logic"] = function (block) {
  const a = vbaGenerator.valueToCode(block, "A", vbaGenerator.ORDER_NONE) || "False";
  const b = vbaGenerator.valueToCode(block, "B", vbaGenerator.ORDER_NONE) || "False";
  return [`(${a} ${block.getFieldValue("OP")} ${b})`, vbaGenerator.ORDER_ATOMIC];
};

vbaGenerator.forBlock["cond_not"] = function (block) {
  const a = vbaGenerator.valueToCode(block, "A", vbaGenerator.ORDER_NONE) || "False";
  return [`Not (${a})`, vbaGenerator.ORDER_ATOMIC];
};

vbaGenerator.forBlock["value_bool"] = function (block) {
  return [block.getFieldValue("BOOL"), vbaGenerator.ORDER_ATOMIC];
};

vbaGenerator.forBlock["prog_exit_sub"] = function () {
  return "Exit Sub";
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
  const title = block.getFieldValue("TITLE") || "";
  const def = block.getFieldValue("DEFAULT") || "";
  if (!title && !def) return [`InputBox(${vbaStr(prompt)})`, vbaGenerator.ORDER_ATOMIC];
  // はじめの答えが数なら、教科書のように数のまま書く
  const defCode = !def ? '""' : /^-?\d+(\.\d+)?$/.test(def) ? def : vbaStr(def);
  return [`InputBox(${vbaStr(prompt)}, ${vbaStr(title)}, ${defCode})`, vbaGenerator.ORDER_ATOMIC];
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
  return `Worksheets.Add.Name = ${vbaStr(name)}`;
};

vbaGenerator.forBlock["sheet_select"] = function (block) {
  const name = block.getFieldValue("NAME");
  return `Worksheets(${vbaStr(name)}).Select`;
};

// ===== ユーザーフォーム =====
// 手続き（Sub）を作るブロックは、本体のコードだけを返す。
// Sub の見出し・Dim・End Sub は generateVBAWithMap がまとめて付ける
const PROC_BLOCK_TYPES = new Set(["form_button_click", "form_initialize", "proc_sub"]);
function procHeaderOf(block) {
  if (block.type === "form_initialize") return "Private Sub UserForm_Initialize()";
  if (block.type === "proc_sub") return `Sub ${block.getFieldValue("NAME")}()`;
  return `Private Sub ${block.getFieldValue("CONTROL")}_Click()`;
}
vbaGenerator.forBlock["proc_sub"] = function (block) {
  return vbaGenerator.statementToCode(block, "DO");
};
vbaGenerator.forBlock["proc_call"] = function (block) {
  return `Call ${block.getFieldValue("NAME")}`;
};
vbaGenerator.forBlock["var_dim"] = function (block) {
  return `Dim ${varName(block)} As ${block.getFieldValue("TYPE")}`;
};
vbaGenerator.forBlock["math_format"] = function (block) {
  const v = vbaGenerator.valueToCode(block, "VALUE", vbaGenerator.ORDER_NONE) || "0";
  return [`Format(${v}, ${vbaStr(block.getFieldValue("FORMAT"))})`, vbaGenerator.ORDER_ATOMIC];
};
vbaGenerator.forBlock["value_newline"] = function () {
  return ["vbCrLf", vbaGenerator.ORDER_ATOMIC];
};
vbaGenerator.forBlock["form_button_click"] = function (block) {
  return vbaGenerator.statementToCode(block, "DO");
};
vbaGenerator.forBlock["form_initialize"] = function (block) {
  return vbaGenerator.statementToCode(block, "DO");
};

vbaGenerator.forBlock["ctl_set_prop"] = function (block) {
  const val = vbaGenerator.valueToCode(block, "VALUE", vbaGenerator.ORDER_NONE) || '""';
  return `${block.getFieldValue("CONTROL")}.${block.getFieldValue("PROP")} = ${val}`;
};
vbaGenerator.forBlock["ctl_get_prop"] = function (block) {
  return [`${block.getFieldValue("CONTROL")}.${block.getFieldValue("PROP")}`, vbaGenerator.ORDER_ATOMIC];
};
vbaGenerator.forBlock["form_unload"] = function () {
  return "Unload Me";
};
vbaGenerator.forBlock["prog_end"] = function () {
  return "End";
};
vbaGenerator.forBlock["math_val"] = function (block) {
  const v = vbaGenerator.valueToCode(block, "VALUE", vbaGenerator.ORDER_NONE) || '""';
  return [`Val(${v})`, vbaGenerator.ORDER_ATOMIC];
};

// ===== 値ブロック =====

vbaGenerator.forBlock["value_number"] = function (block) {
  return [String(block.getFieldValue("NUM")), vbaGenerator.ORDER_ATOMIC];
};

vbaGenerator.forBlock["value_text"] = function (block) {
  const text = block.getFieldValue("TEXT");
  return [vbaStr(text), vbaGenerator.ORDER_ATOMIC];
};

vbaGenerator.forBlock["value_math"] = function (block) {
  const a = vbaGenerator.valueToCode(block, "A", vbaGenerator.ORDER_NONE) || "0";
  const op = block.getFieldValue("OP");
  const b = vbaGenerator.valueToCode(block, "B", vbaGenerator.ORDER_NONE) || "0";
  const vbaOp = op === "mod" ? "Mod" : op;
  return [`(${a} ${vbaOp} ${b})`, vbaGenerator.ORDER_ATOMIC];
};

// ===== ワークスペース全体 → 手続き（Sub）ごとにまとめる =====
// 実際にブロックで使われている変数だけ収集（未使用変数はDimしない）
function collectVarNames(blocks) {
  const names = new Set();
  blocks.forEach((b) => {
    if (b.type === "var_set" || b.type === "var_get" || b.type === "var_change") {
      const f = b.getField("VAR");
      if (f) names.add(f.getText());
    }
  });
  return names;
}

// 生成コードの各行がどのブロックから作られたかを調べるための目印。
// STATEMENT_PREFIX で文ブロックの前に「'@@'ブロックID'」という行を入れ、あとで取り除く
const LINE_MARK_RE = /^\s*'@@'(.*)'\s*$/;

// 目印の行を取り除き、各行がどのブロックのものかを調べる。戻り値: { lines, ids }
function splitMarkedCode(raw) {
  const lines = [];
  const ids = [];
  let pendingId = null;
  raw.split("\n").forEach((line) => {
    const m = line.match(LINE_MARK_RE);
    if (m) {
      pendingId = m[1];
      return;
    }
    lines.push(line);
    if (line.trim()) {
      ids.push(pendingId);
      pendingId = null;
    } else {
      ids.push(null);
    }
  });
  // Next / End If / Else / Loop の行は、同じ深さで直前に始まったブロックのものとする
  for (let k = 0; k < lines.length; k++) {
    if (ids[k] || !lines[k].trim()) continue;
    const indent = lines[k].match(/^\s*/)[0].length;
    for (let p = k - 1; p >= 0; p--) {
      if (!lines[p].trim()) continue;
      const pi = lines[p].match(/^\s*/)[0].length;
      if (pi === indent && ids[p]) {
        ids[k] = ids[p];
        break;
      }
      if (pi < indent) break;
    }
  }
  while (lines.length && !lines[lines.length - 1].trim()) {
    lines.pop();
    ids.pop();
  }
  return { lines, ids };
}

// 1つの手続き分のコードを作る。stacks: 本体にするブロック（つながりの先頭）の配列
function buildProc(header, stacks, ownerBlock, descendants) {
  vbaGenerator._loopDepth = 0;
  vbaGenerator._usedCounters = new Set();
  let raw = "";
  if (ownerBlock) {
    raw = vbaGenerator.blockToCode(ownerBlock);
    if (Array.isArray(raw)) raw = raw[0];
    // 手続きブロック自身の目印（見出しの行用）を取り出す
    const first = raw.split("\n")[0].match(LINE_MARK_RE);
    if (first) raw = raw.split("\n").slice(1).join("\n");
    // 本体は statementToCode で1段字下げされているので、そのまま使う
    var headerId = first ? first[1] : ownerBlock.id;
  } else {
    raw = stacks
      .map((b) => {
        const c = vbaGenerator.blockToCode(b);
        return Array.isArray(c) ? "" : c;
      })
      .filter((c) => c.trim())
      .map((c) => c.split("\n").map((l) => (l.trim() && !LINE_MARK_RE.test(l) ? "    " + l : l)).join("\n"))
      .join("\n");
  }
  const { lines: bodyLines, ids: bodyIds } = splitMarkedCode(raw);
  const body = bodyLines.join("\n");
  const head = [header];
  // 予約名（くり返しのカウンタ i・j・k / 配列 arr）と同じ名前の変数があっても
  // Dim を二重に出力しない（「同じ適用範囲内で宣言が重複しています」エラー対策）
  const declared = new Set();
  LOOP_COUNTERS.forEach((c) => {
    if (vbaGenerator._usedCounters.has(c)) {
      head.push(`    Dim ${c} As Integer`);
      declared.add(c);
    }
  });
  if (/\barr\(/.test(body)) {
    head.push("    Dim arr(1 To 100) As Variant");
    declared.add("arr");
  }
  // 「📦 箱を用意する（Dim）」ブロックで用意した変数は、自動で Dim しない
  descendants.forEach((b) => {
    if (b.type === "var_dim") declared.add(varName(b));
  });
  collectVarNames(descendants).forEach((n) => {
    if (declared.has(n)) return;
    head.push(`    Dim ${n} As Variant`);
    declared.add(n);
  });
  const lines = head.concat(bodyLines.length ? bodyLines : [], ["End Sub"]);
  const ids = head.map((_, k) => (k === 0 && ownerBlock ? headerId : null)).concat(bodyIds, [ownerBlock ? headerId : null]);
  return { lines, ids };
}

// 戻り値: { code, lineMap }。lineMap は「コードの行番号（1始まり）→ ブロックID」
// ・「〇〇をクリックしたとき」などの手続きブロック → Private Sub 〇〇_Click() 〜 End Sub
// ・どの手続きブロックにも入っていないブロック → Sub MyMacro() 〜 End Sub
function generateVBAWithMap(workspace) {
  vbaGenerator.init(workspace);
  vbaGenerator.STATEMENT_PREFIX = "'@@%1\n";
  const parts = [];
  try {
    const tops = workspace.getTopBlocks(true);
    const procBlocks = tops.filter((b) => PROC_BLOCK_TYPES.has(b.type));
    const loose = tops.filter((b) => !PROC_BLOCK_TYPES.has(b.type) && !b.outputConnection);
    if (loose.length) {
      const desc = [];
      loose.forEach((b) => desc.push(...b.getDescendants(false)));
      parts.push(buildProc("Sub MyMacro()", loose, null, desc));
    }
    procBlocks.forEach((b) => parts.push(buildProc(procHeaderOf(b), null, b, b.getDescendants(false))));
  } finally {
    vbaGenerator.STATEMENT_PREFIX = null;
  }
  if (!parts.length) {
    return { code: "Sub MyMacro()\n    ' ブロックを組み立てるとここにコードが表示されます\nEnd Sub", lineMap: {} };
  }
  // フォームの手続きがあるときは、教科書と同じく先頭に Option Explicit を書く
  const hasForm = parts.length > 1 || workspace.getTopBlocks(false).some((b) => PROC_BLOCK_TYPES.has(b.type));
  const lines = hasForm ? ["Option Explicit"] : [];
  const ids = hasForm ? [null] : [];
  parts.forEach((p, k) => {
    if (k > 0 || hasForm) {
      lines.push("");
      ids.push(null);
    }
    lines.push(...p.lines);
    ids.push(...p.ids);
  });
  const lineMap = {};
  ids.forEach((id, k) => {
    if (id) lineMap[k + 1] = id;
  });
  return { code: lines.join("\n"), lineMap };
}

function generateVBA(workspace) {
  return generateVBAWithMap(workspace).code;
}
