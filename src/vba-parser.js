/* ===== VBA コード → ブロック 変換パーサー =====
 * 手打ちされた VBA コードを解析し、Blockly のワークスペース JSON
 * （Blockly.serialization.workspaces.load に渡せる形）を組み立てる。
 *
 * 対応範囲は意図的に「このアプリのブロック語彙で表現できる VBA」に限定している。
 * Select Case / With / 独自 Function など対応外の構文に出会ったら
 * VbaParseError を投げる（行番号つき）。呼び出し側はそれを捕まえて
 * エラー表示し、ワークスペースは直前の状態を維持する。
 *
 * Blockly に依存しないので Node 単体でテストできる（変数IDの解決だけ
 * resolveVariables() で workspace を受け取って行う）。
 */

class VbaParseError extends Error {
  constructor(message, line) {
    super(message);
    this.name = "VbaParseError";
    this.line = line; // 1始まりの行番号（分からなければ null）
  }
}

// ===== 色: vba-generator.js の COLOR_RGB と対になる逆引きテーブル =====
const COLOR_RGB_TRIPLES = {
  RED: [231, 76, 60],
  GREEN: [33, 115, 70],
  BLUE: [52, 152, 219],
  YELLOW: [241, 196, 15],
  WHITE: [255, 255, 255],
};
function rgbToColorName(inner, line) {
  const nums = inner.split(",").map((s) => parseInt(s.trim(), 10));
  if (nums.length !== 3 || nums.some((n) => isNaN(n))) {
    throw new VbaParseError(`RGB(...) の書き方が読み取れません: RGB(${inner})`, line);
  }
  for (const name in COLOR_RGB_TRIPLES) {
    const t = COLOR_RGB_TRIPLES[name];
    if (t[0] === nums[0] && t[1] === nums[1] && t[2] === nums[2]) return name;
  }
  throw new VbaParseError(
    `この色（RGB(${inner})）は色ブロックに無い色です。🟥🟩🟦🟨⬜ のどれかを使ってください`,
    line
  );
}

const HIRAGANA_RE = /[぀-ゟ]/;

// ===== 字句解析（式の中身だけを対象。文レベルは行ごとの正規表現で処理） =====
function tokenize(str, line) {
  const tokens = [];
  let i = 0;
  const n = str.length;
  while (i < n) {
    const c = str[i];
    if (/\s/.test(c)) {
      i++;
      continue;
    }
    if (c === '"') {
      let j = i + 1;
      let out = "";
      while (j < n && str[j] !== '"') {
        out += str[j];
        j++;
      }
      if (j >= n) throw new VbaParseError('文字列の " が閉じていません', line);
      tokens.push({ type: "STR", value: out });
      i = j + 1;
      continue;
    }
    if (/[0-9]/.test(c) || (c === "-" && /[0-9]/.test(str[i + 1] || "") && isUnaryPos(tokens))) {
      let j = i + (c === "-" ? 1 : 0);
      let start = i;
      j = start + (c === "-" ? 1 : 0);
      while (j < n && /[0-9.]/.test(str[j])) j++;
      tokens.push({ type: "NUM", value: Number(str.slice(start, j)) });
      i = j;
      continue;
    }
    if (/[A-Za-z_]/.test(c)) {
      let j = i;
      while (j < n && /[A-Za-z0-9_]/.test(str[j])) j++;
      const word = str.slice(i, j);
      tokens.push({ type: "IDENT", value: word });
      i = j;
      continue;
    }
    if (c === "<" && str[i + 1] === ">") {
      tokens.push({ type: "OP", value: "<>" });
      i += 2;
      continue;
    }
    if (c === "<" && str[i + 1] === "=") {
      tokens.push({ type: "OP", value: "<=" });
      i += 2;
      continue;
    }
    if (c === ">" && str[i + 1] === "=") {
      tokens.push({ type: "OP", value: ">=" });
      i += 2;
      continue;
    }
    if ("+-*/&=<>(),.".includes(c)) {
      tokens.push({ type: "OP", value: c });
      i++;
      continue;
    }
    throw new VbaParseError(`読み取れない文字です: "${c}"`, line);
  }
  return tokens;
}
// マイナスを「単項」とみなす位置か（式の先頭、'(' の直後、演算子/カンマの直後）
function isUnaryPos(tokens) {
  if (tokens.length === 0) return true;
  const last = tokens[tokens.length - 1];
  return last.type === "OP" && last.value !== ")";
}

// ===== 式パーサー（トークン列 → 値ブロック仕様） =====
class ExprParser {
  constructor(tokens, line) {
    this.tokens = tokens;
    this.pos = 0;
    this.line = line;
  }
  peek() {
    return this.tokens[this.pos];
  }
  next() {
    return this.tokens[this.pos++];
  }
  expectOp(v) {
    const t = this.next();
    if (!t || t.type !== "OP" || t.value !== v) {
      throw new VbaParseError(`「${v}」が来るはずです`, this.line);
    }
    return t;
  }
  atEnd() {
    return this.pos >= this.tokens.length;
  }

  // 比較を含む条件式（If / Do While / Do Until 用）
  parseCondition() {
    const a = this.parseConcat();
    const t = this.peek();
    const cmpOps = ["=", "<>", "<", ">", "<=", ">="];
    if (t && t.type === "OP" && cmpOps.includes(t.value)) {
      this.next();
      const b = this.parseConcat();
      if (!this.atEnd()) throw new VbaParseError("条件の書き方が読み取れません", this.line);
      return { type: "cond_compare", fields: { OP: t.value }, inputs: { A: { block: a }, B: { block: b } } };
    }
    throw new VbaParseError(
      "条件には比較（= や < など）が必要です。例: GOKEI < 100",
      this.line
    );
  }

  // 一般の値式（比較なし）。呼び出し後に atEnd() でないと呼び出し側が判断する
  parseExpr() {
    const v = this.parseConcat();
    if (!this.atEnd()) throw new VbaParseError("式の書き方が読み取れません", this.line);
    return v;
  }

  parseConcat() {
    let left = this.parseAdd();
    while (this.peek() && this.peek().type === "OP" && this.peek().value === "&") {
      this.next();
      const right = this.parseAdd();
      left = { type: "text_concat", inputs: { A: { block: left }, B: { block: right } } };
    }
    return left;
  }
  parseAdd() {
    let left = this.parseMul();
    while (this.peek() && this.peek().type === "OP" && (this.peek().value === "+" || this.peek().value === "-")) {
      const op = this.next().value;
      const right = this.parseMul();
      left = { type: "value_math", fields: { OP: op }, inputs: { A: { block: left }, B: { block: right } } };
    }
    return left;
  }
  parseMul() {
    let left = this.parsePrimary();
    while (
      this.peek() &&
      ((this.peek().type === "OP" && (this.peek().value === "*" || this.peek().value === "/")) ||
        (this.peek().type === "IDENT" && this.peek().value === "Mod"))
    ) {
      const t = this.next();
      const op = t.type === "IDENT" ? "mod" : t.value;
      const right = this.parsePrimary();
      left = { type: "value_math", fields: { OP: op }, inputs: { A: { block: left }, B: { block: right } } };
    }
    return left;
  }

  parsePrimary() {
    const t = this.peek();
    if (!t) throw new VbaParseError("式が足りません", this.line);

    if (t.type === "NUM") {
      this.next();
      return { type: "value_number", fields: { NUM: t.value } };
    }
    if (t.type === "STR") {
      this.next();
      return { type: "value_text", fields: { TEXT: t.value } };
    }
    if (t.type === "OP" && t.value === "(") {
      this.next();
      const inner = this.parseConcat();
      this.expectOp(")");
      return inner;
    }
    if (t.type === "IDENT") {
      return this.parseIdentChain();
    }
    throw new VbaParseError(`ここに来るはずのないものがあります: "${t.value}"`, this.line);
  }

  parseIdentChain() {
    const name = this.next().value;

    if (name === "Range") {
      this.expectOp("(");
      const addr = this.parseAddrArg();
      this.expectOp(")");
      this.expectDot("Value");
      return { type: "cell_get_value", fields: { CELL: addr } };
    }
    if (name === "Cells") {
      this.expectOp("(");
      const row = this.parseConcat();
      this.expectOp(",");
      const col = this.parseConcat();
      this.expectOp(")");
      this.expectDot("Value");
      return { type: "cells_get_value", inputs: { ROW: { block: row }, COL: { block: col } } };
    }
    if (name === "arr") {
      this.expectOp("(");
      const index = this.parseConcat();
      this.expectOp(")");
      return { type: "array_get", inputs: { INDEX: { block: index } } };
    }
    if (name === "InputBox") {
      this.expectOp("(");
      const t = this.next();
      if (!t || t.type !== "STR") throw new VbaParseError('InputBox の中身は "文字" にしてください', this.line);
      this.expectOp(")");
      return { type: "io_inputbox", fields: { PROMPT: t.value } };
    }
    if (name === "WorksheetFunction") {
      this.expectOp(".");
      const fn = this.next();
      if (!fn || fn.type !== "IDENT" || fn.value !== "Round") {
        throw new VbaParseError("WorksheetFunction.Round 以外は対応していません", this.line);
      }
      this.expectOp("(");
      const val = this.parseConcat();
      this.expectOp(",");
      const digits = this.parseConcat();
      this.expectOp(")");
      return { type: "math_round", inputs: { VALUE: { block: val }, DIGITS: { block: digits } } };
    }
    if (name === "Date") {
      return { type: "value_date" };
    }
    if (name === "i") {
      return { type: "loop_index" };
    }
    if (name === "True" || name === "False") {
      throw new VbaParseError("真偽値（True/False）を値として使うブロックはありません", this.line);
    }
    // 予約語なのに専用ルールに引っかからなかった＝対応外の使い方
    const RESERVED = [
      "For", "To", "Step", "Next", "If", "Then", "Else", "End", "Do", "While", "Until", "Loop",
      "Dim", "As", "Sub", "MsgBox", "Range", "Cells", "Rows", "Worksheets", "WorksheetFunction",
      "Round", "Select", "ClearContents", "Borders", "LineStyle", "Copy", "Add", "Name",
      "Interior", "Color", "Font", "Bold", "Size", "Value", "RGB",
    ];
    if (RESERVED.includes(name)) {
      throw new VbaParseError(`「${name}」の使い方が対応している形と違います`, this.line);
    }
    if (HIRAGANA_RE.test(name)) {
      throw new VbaParseError(`変数名にひらがなは使えません: ${name}`, this.line);
    }
    return { type: "var_get", fields: { VAR: { __var: name } } };
  }

  expectDot(word) {
    this.expectOp(".");
    const t = this.next();
    if (!t || t.type !== "IDENT" || t.value !== word) {
      throw new VbaParseError(`「.${word}」が来るはずです`, this.line);
    }
  }

  // Range( ... ) の中身：STR（"A1"）または STR & i（"A" & i）
  parseAddrArg() {
    const t = this.next();
    if (!t || t.type !== "STR") {
      throw new VbaParseError('セル番地は "A1" のように書いてください', this.line);
    }
    if (this.peek() && this.peek().type === "OP" && this.peek().value === "&") {
      this.next();
      const idTok = this.next();
      if (!idTok || idTok.type !== "IDENT" || idTok.value !== "i") {
        throw new VbaParseError('動的なセル番地は "A" & i の形だけ対応しています', this.line);
      }
      return t.value + "{i}";
    }
    return t.value;
  }
}

function parseExprString(text, line) {
  const tokens = tokenize(text, line);
  return new ExprParser(tokens, line).parseExpr();
}
function parseConditionString(text, line) {
  const tokens = tokenize(text, line);
  return new ExprParser(tokens, line).parseCondition();
}

// ===== Range(...) の中身だけを文字列のまま取り出す（フィールド用。ブロックにはしない） =====
function parseRangeAddrArg(rawInner, line) {
  const t = rawInner.trim();
  let m = t.match(/^"([^"]*)"$/);
  if (m) return m[1];
  m = t.match(/^"([A-Za-z]+)"\s*&\s*i$/);
  if (m) return m[1] + "{i}";
  throw new VbaParseError(`セル番地の書き方が読み取れません: Range(${t})`, line);
}

// ===== 行の前処理 =====
// コメント（' 以降。文字列の中の ' は無視）を取り除き、行番号付きの非空行リストを作る
function preprocessLines(source) {
  const rawLines = source.split(/\r\n|\r|\n/);
  const out = [];
  rawLines.forEach((raw, idx) => {
    let line = raw;
    let inStr = false;
    for (let i = 0; i < line.length; i++) {
      if (line[i] === '"') inStr = !inStr;
      else if (line[i] === "'" && !inStr) {
        line = line.slice(0, i);
        break;
      }
    }
    const trimmed = line.trim();
    if (trimmed) out.push({ text: trimmed, line: idx + 1 });
  });
  return out;
}

// Sub ヘッダー・End Sub・Dim 行を取り除く
function stripWrapper(lines) {
  let out = lines.slice();
  if (out.length && /^Sub\s+\w+\s*\(\s*\)\s*$/i.test(out[0].text)) out = out.slice(1);
  if (out.length && /^End\s+Sub$/i.test(out[out.length - 1].text)) out = out.slice(0, -1);
  out = out.filter((l) => !/^Dim\s+\S+/i.test(l.text));
  return out;
}

// ===== 文のパーサー本体（行を消費するカーソル） =====
class LineCursor {
  constructor(lines) {
    this.lines = lines;
    this.pos = 0;
  }
  atEnd() {
    return this.pos >= this.lines.length;
  }
  peek() {
    return this.lines[this.pos];
  }
  next() {
    return this.lines[this.pos++];
  }
}

const STMT_STARTERS = {
  FOR: /^For\s+/i,
  DO: /^Do\s+/i,
  IF: /^If\s+/i,
  MSGBOX: /^MsgBox\s+/i,
};

function normalizeKeywordLine(text) {
  return text.trim();
}

// stopWords: このカーソル位置の行が一致したら止まる（消費しない）
function parseStatements(cursor, stopWords) {
  let head = null;
  let tail = null;
  while (!cursor.atEnd()) {
    const line = normalizeKeywordLine(cursor.peek().text);
    if (stopWords.some((w) => new RegExp("^" + w + "\\b", "i").test(line) || line.toLowerCase() === w.toLowerCase())) {
      break;
    }
    const block = parseOneStatement(cursor);
    if (!head) head = block;
    else tail.next = { block };
    tail = block;
  }
  return head;
}

function parseOneStatement(cursor) {
  const { text, line } = cursor.peek();
  if (STMT_STARTERS.FOR.test(text)) return parseFor(cursor);
  if (STMT_STARTERS.DO.test(text)) return parseDo(cursor);
  if (STMT_STARTERS.IF.test(text)) return parseIf(cursor);
  if (STMT_STARTERS.MSGBOX.test(text)) return parseMsgBox(cursor);
  return parseSimpleStatement(cursor);
}

function parseFor(cursor) {
  const { text, line } = cursor.next();
  const m = text.match(/^For\s+i\s*=\s*(.+?)\s+To\s+(.+?)(?:\s+Step\s+(.+))?$/i);
  if (!m) {
    throw new VbaParseError(
      '"For i = 開始 To 終わり" の形にしてください（例: For i = 1 To 5）',
      line
    );
  }
  const start = parseExprString(m[1], line);
  const end = parseExprString(m[2], line);
  const step = m[3] ? parseExprString(m[3], line) : null;
  const body = parseStatements(cursor, ["Next"]);
  if (cursor.atEnd()) throw new VbaParseError('"Next" が見つかりません（For に対応する終わり）', line);
  const closer = cursor.next();
  if (!/^Next(\s+i)?$/i.test(closer.text)) {
    throw new VbaParseError('"Next" の書き方が違います', closer.line);
  }
  const inputs = { START: { block: start }, END: { block: end } };
  if (body) inputs.DO = { block: body };
  if (step) {
    inputs.STEP = { block: step };
    return { type: "loop_for_step", inputs };
  }
  return { type: "loop_range", inputs };
}

function parseDo(cursor) {
  const { text, line } = cursor.next();
  const m = text.match(/^Do\s+(While|Until)\s+(.+)$/i);
  if (!m) {
    throw new VbaParseError('"Do While 条件" か "Do Until 条件" の形にしてください', line);
  }
  const cond = parseConditionString(m[2], line);
  const body = parseStatements(cursor, ["Loop"]);
  if (cursor.atEnd()) throw new VbaParseError('"Loop" が見つかりません（Do に対応する終わり）', line);
  const closer = cursor.next();
  if (!/^Loop$/i.test(closer.text)) {
    throw new VbaParseError('"Loop" の書き方が違います', closer.line);
  }
  const inputs = { CONDITION: { block: cond } };
  if (body) inputs.DO = { block: body };
  const type = /^while$/i.test(m[1]) ? "loop_while" : "loop_do_until";
  return { type, inputs };
}

function parseIf(cursor) {
  const { text, line } = cursor.next();
  const m = text.match(/^If\s+(.+?)\s+Then$/i);
  if (!m) {
    throw new VbaParseError('"If 条件 Then" の形にしてください（Then で終わる1行）', line);
  }
  const cond = parseConditionString(m[1], line);
  const thenBody = parseStatements(cursor, ["Else", "End If"]);
  if (cursor.atEnd()) throw new VbaParseError('"End If" が見つかりません（If に対応する終わり）', line);
  let elseBody = null;
  if (/^Else$/i.test(cursor.peek().text)) {
    cursor.next();
    elseBody = parseStatements(cursor, ["End If"]);
    if (cursor.atEnd()) throw new VbaParseError('"End If" が見つかりません（Else に対応する終わり）', line);
  }
  const closer = cursor.next();
  if (!/^End If$/i.test(closer.text)) {
    throw new VbaParseError('"End If" の書き方が違います', closer.line);
  }
  const inputs = { CONDITION: { block: cond } };
  if (thenBody) inputs.THEN = { block: thenBody };
  if (elseBody) inputs.ELSE = { block: elseBody };
  return { type: "cond_if", inputs };
}

function parseMsgBox(cursor) {
  const { text, line } = cursor.next();
  const m = text.match(/^MsgBox\s+(.+)$/i);
  const msg = parseExprString(m[1], line);
  return { type: "io_msgbox", inputs: { MSG: { block: msg } } };
}

// 「簡単な1行の文」たち：代入／メソッド呼び出し系
const SIMPLE_RULES = [
  // シート
  [/^Worksheets\.Add\.Name\s*=\s*"([^"]*)"$/i, (m) => ({ type: "sheet_add", fields: { NAME: m[1] } })],
  [/^Worksheets\("([^"]*)"\)\.Select$/i, (m) => ({ type: "sheet_select", fields: { NAME: m[1] } })],

  // 範囲（Cells指定）の罫線・選択
  [
    /^Range\(Cells\(([^,]+),([^)]+)\),\s*Cells\(([^,]+),([^)]+)\)\)\.Borders\.LineStyle\s*=\s*1$/i,
    (m, line) => ({
      type: "range_border",
      inputs: {
        R1: { block: parseExprString(m[1], line) },
        C1: { block: parseExprString(m[2], line) },
        R2: { block: parseExprString(m[3], line) },
        C2: { block: parseExprString(m[4], line) },
      },
    }),
  ],
  [
    /^Range\(Cells\(([^,]+),([^)]+)\),\s*Cells\(([^,]+),([^)]+)\)\)\.Select$/i,
    (m, line) => ({
      type: "range_select",
      inputs: {
        R1: { block: parseExprString(m[1], line) },
        C1: { block: parseExprString(m[2], line) },
        R2: { block: parseExprString(m[3], line) },
        C2: { block: parseExprString(m[4], line) },
      },
    }),
  ],

  // Rows(...)
  [/^Rows\((.+)\)\.ClearContents$/i, (m, line) => ({ type: "cell_clear_row", inputs: { ROW: { block: parseExprString(m[1], line) } } })],

  // Range(...) 系（アドレスは括弧の中身をあとで parseRangeAddrArg で解決）
  [
    /^Range\((.+?)\)\.Copy\s+Range\((.+?)\)$/i,
    (m, line) => ({
      type: "cell_copy",
      fields: { FROM: parseRangeAddrArg(m[1], line), TO: parseRangeAddrArg(m[2], line) },
    }),
  ],
  [/^Range\((.+?)\)\.ClearContents$/i, (m, line) => ({ type: "cell_clear", fields: { CELL: parseRangeAddrArg(m[1], line) } })],
  [
    /^Range\((.+?)\)\.Interior\.Color\s*=\s*RGB\((.+?)\)$/i,
    (m, line) => ({ type: "fmt_bgcolor", fields: { CELL: parseRangeAddrArg(m[1], line), COLOR: rgbToColorName(m[2], line) } }),
  ],
  [/^Range\((.+?)\)\.Font\.Bold\s*=\s*True$/i, (m, line) => ({ type: "fmt_bold", fields: { CELL: parseRangeAddrArg(m[1], line) } })],
  [
    /^Range\((.+?)\)\.Font\.Size\s*=\s*(.+)$/i,
    (m, line) => ({
      type: "fmt_fontsize",
      fields: { CELL: parseRangeAddrArg(m[1], line) },
      inputs: { SIZE: { block: parseExprString(m[2], line) } },
    }),
  ],
  [
    /^Range\((.+?)\)\.Value\s*=\s*(.+)$/i,
    (m, line) => ({
      type: "cell_set_value",
      fields: { CELL: parseRangeAddrArg(m[1], line) },
      inputs: { VALUE: { block: parseExprString(m[2], line) } },
    }),
  ],

  // Cells(...) 系
  [
    /^Cells\(([^,]+),([^)]+)\)\.Interior\.Color\s*=\s*RGB\((.+?)\)$/i,
    (m, line) => ({
      type: "cells_bgcolor",
      fields: { COLOR: rgbToColorName(m[3], line) },
      inputs: { ROW: { block: parseExprString(m[1], line) }, COL: { block: parseExprString(m[2], line) } },
    }),
  ],
  [
    /^Cells\(([^,]+),([^)]+)\)\.Font\.Bold\s*=\s*True$/i,
    (m, line) => ({
      type: "cells_bold",
      inputs: { ROW: { block: parseExprString(m[1], line) }, COL: { block: parseExprString(m[2], line) } },
    }),
  ],
  [
    /^Cells\(([^,]+),([^)]+)\)\.ClearContents$/i,
    (m, line) => ({
      type: "cells_clear",
      inputs: { ROW: { block: parseExprString(m[1], line) }, COL: { block: parseExprString(m[2], line) } },
    }),
  ],
  [
    /^Cells\(([^,]+),([^)]+)\)\.Value\s*=\s*(.+)$/i,
    (m, line) => ({
      type: "cells_set_value",
      inputs: {
        ROW: { block: parseExprString(m[1], line) },
        COL: { block: parseExprString(m[2], line) },
        VALUE: { block: parseExprString(m[3], line) },
      },
    }),
  ],

  // 配列
  [
    /^arr\((.+?)\)\s*=\s*(.+)$/i,
    (m, line) => ({
      type: "array_set",
      inputs: { INDEX: { block: parseExprString(m[1], line) }, VALUE: { block: parseExprString(m[2], line) } },
    }),
  ],

  // 変数：name = name + delta → 増減ブロック（それ以外の代入より先に判定）
  [
    /^([A-Za-z_]\w*)\s*=\s*\1\s*\+\s*(.+)$/i,
    (m, line) => {
      if (m[1] === "i") throw new VbaParseError("カウンタ i には代入できません", line);
      if (HIRAGANA_RE.test(m[1])) throw new VbaParseError(`変数名にひらがなは使えません: ${m[1]}`, line);
      return { type: "var_change", fields: { VAR: { __var: m[1] } }, inputs: { DELTA: { block: parseExprString(m[2], line) } } };
    },
  ],
  // 変数：一般の代入
  [
    /^([A-Za-z_]\w*)\s*=\s*(.+)$/i,
    (m, line) => {
      if (m[1] === "i") throw new VbaParseError("カウンタ i には代入できません", line);
      if (HIRAGANA_RE.test(m[1])) throw new VbaParseError(`変数名にひらがなは使えません: ${m[1]}`, line);
      return { type: "var_set", fields: { VAR: { __var: m[1] } }, inputs: { VALUE: { block: parseExprString(m[2], line) } } };
    },
  ],
];

function parseSimpleStatement(cursor) {
  const { text, line } = cursor.next();
  for (const [re, build] of SIMPLE_RULES) {
    const m = text.match(re);
    if (m) return build(m, line);
  }
  // 代入っぽい形（左辺 = 右辺）だが SIMPLE_RULES に引っかからなかった場合、
  // ひらがな変数名が原因なら分かりやすいメッセージにする
  const asg = text.match(/^(\S+)\s*=\s*.+$/);
  if (asg && HIRAGANA_RE.test(asg[1])) {
    throw new VbaParseError(`変数名にひらがなは使えません: ${asg[1]}`, line);
  }
  throw new VbaParseError(`対応していない書き方です: "${text}"`, line);
}

// ===== 公開API =====

// source: エディタの全文（Sub〜End Subの外側込みでも中身だけでもOK）
// 戻り値: { root: <先頭ブロック spec or null>, variableNames: string[] }
function parseVBA(source) {
  const lines = stripWrapper(preprocessLines(source));
  const cursor = new LineCursor(lines);
  const root = parseStatements(cursor, []);
  if (!cursor.atEnd()) {
    const { text, line } = cursor.peek();
    throw new VbaParseError(`ここでの書き方が読み取れません: "${text}"`, line);
  }
  const variableNames = new Set();
  if (root) collectVariables(root, variableNames);
  return { root, variableNames: [...variableNames] };
}

function collectVariables(block, set) {
  if (!block || typeof block !== "object") return;
  if (block.fields && block.fields.VAR && block.fields.VAR.__var) {
    set.add(block.fields.VAR.__var);
  }
  if (block.inputs) {
    for (const key in block.inputs) collectVariables(block.inputs[key].block, set);
  }
  if (block.next) collectVariables(block.next.block, set);
}

// {__var:name} プレースホルダを実際の変数IDに解決する（workspace が要る＝ブラウザ側で呼ぶ）
function resolveVariables(block, workspace) {
  if (!block || typeof block !== "object") return;
  if (block.fields && block.fields.VAR && block.fields.VAR.__var) {
    const v = workspace.createVariable(block.fields.VAR.__var);
    block.fields.VAR = { id: v.getId() };
  }
  if (block.inputs) {
    for (const key in block.inputs) resolveVariables(block.inputs[key].block, workspace);
  }
  if (block.next) resolveVariables(block.next.block, workspace);
}

// root ブロック（先頭の1つ）を Blockly.serialization.workspaces.load 用のJSONに包む
function toWorkspaceJson(root) {
  return { blocks: { languageVersion: 0, blocks: root ? [root] : [] } };
}

// Node / ブラウザ両対応のエクスポート
const api = { parseVBA, resolveVariables, toWorkspaceJson, VbaParseError };
if (typeof module !== "undefined" && module.exports) {
  module.exports = api;
}
if (typeof window !== "undefined") {
  window.VbaParser = api;
}
