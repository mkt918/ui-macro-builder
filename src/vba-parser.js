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
      while (j < n) {
        if (str[j] === '"') {
          if (str[j + 1] === '"') {
            out += '"'; // "" は文字列の中の " 1文字
            j += 2;
            continue;
          }
          break;
        }
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
    if ("+-*/&=<>(),.^\\".includes(c)) {
      tokens.push({ type: "OP", value: c });
      i++;
      continue;
    }
    if (c === "¥") {
      tokens.push({ type: "OP", value: "\\" }); // 教科書の ¥（整数のわり算）
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

// かっこの外にあるカンマで文字列を分ける（"3, 4" や "Umei(1), 2" のように、かっこの中のカンマは分けない）
function splitTopLevel(str, sep) {
  const parts = [];
  let depth = 0;
  let inStr = false;
  let cur = "";
  for (const ch of str) {
    if (ch === '"') inStr = !inStr;
    if (!inStr) {
      if (ch === "(") depth++;
      else if (ch === ")") depth--;
      else if (ch === sep && depth === 0) {
        parts.push(cur);
        cur = "";
        continue;
      }
    }
    cur += ch;
  }
  parts.push(cur);
  return parts;
}

const VBA_TYPE_NAMES = { integer: "Integer", long: "Long", single: "Single", double: "Double", string: "String", variant: "Variant" };

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

  // 条件式（If / Do While / Do Until 用）。優先順位は ( ) → 比較 → Not → And → Or
  parseCondition() {
    const c = this.parseCondOr();
    if (!this.atEnd()) throw new VbaParseError("条件の書き方が読み取れません", this.line);
    return c;
  }
  isKeyword(word) {
    const t = this.peek();
    return !!t && t.type === "IDENT" && t.value.toLowerCase() === word;
  }
  parseCondOr() {
    let left = this.parseCondAnd();
    while (this.isKeyword("or")) {
      this.next();
      const right = this.parseCondAnd();
      left = { type: "cond_logic", fields: { OP: "Or" }, inputs: { A: { block: left }, B: { block: right } } };
    }
    return left;
  }
  parseCondAnd() {
    let left = this.parseCondNot();
    while (this.isKeyword("and")) {
      this.next();
      const right = this.parseCondNot();
      left = { type: "cond_logic", fields: { OP: "And" }, inputs: { A: { block: left }, B: { block: right } } };
    }
    return left;
  }
  parseCondNot() {
    if (this.isKeyword("not")) {
      this.next();
      return { type: "cond_not", inputs: { A: { block: this.parseCondNot() } } };
    }
    return this.parseCondCompare();
  }
  parseCondCompare() {
    const cmpOps = ["=", "<>", "<", ">", "<=", ">="];
    const start = this.pos;
    let a;
    try {
      a = this.parseConcat();
    } catch (e) {
      // 「(A > 1)」のように、かっこの中が条件のとき：かっこ全体を条件として読み直す
      const t0 = this.tokens[start];
      if (!(t0 && t0.type === "OP" && t0.value === "(")) throw e;
      this.pos = start + 1;
      const inner = this.parseCondOr();
      this.expectOp(")");
      return inner;
    }
    const t = this.peek();
    if (t && t.type === "OP" && cmpOps.includes(t.value)) {
      this.next();
      const b = this.parseConcat();
      return { type: "cond_compare", fields: { OP: t.value }, inputs: { A: { block: a }, B: { block: b } } };
    }
    if (a.type === "value_bool") return a;
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
  // * / ¥ Mod は同じ順位で左から（教科書 p.93 の表）
  parseMul() {
    let left = this.parsePow();
    while (
      this.peek() &&
      ((this.peek().type === "OP" && ["*", "/", "\\"].includes(this.peek().value)) ||
        (this.peek().type === "IDENT" && this.peek().value === "Mod"))
    ) {
      const t = this.next();
      const op = t.type === "IDENT" ? "mod" : t.value;
      const right = this.parsePow();
      left = { type: "value_math", fields: { OP: op }, inputs: { A: { block: left }, B: { block: right } } };
    }
    return left;
  }
  parsePow() {
    let left = this.parsePrimary();
    while (this.peek() && this.peek().type === "OP" && this.peek().value === "^") {
      this.next();
      const right = this.parsePrimary();
      left = { type: "value_math", fields: { OP: "^" }, inputs: { A: { block: left }, B: { block: right } } };
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

    // Me.Label1.Caption の Me. は読み飛ばす
    if (name === "Me" && this.peek() && this.peek().type === "OP" && this.peek().value === ".") {
      this.next();
      return this.parseIdentChain();
    }
    if (name === "Val") {
      this.expectOp("(");
      const v = this.parseConcat();
      this.expectOp(")");
      return { type: "math_val", inputs: { VALUE: { block: v } } };
    }
    if (name === "Format") {
      this.expectOp("(");
      const v = this.parseConcat();
      this.expectOp(",");
      const f = this.next();
      if (!f || f.type !== "STR") throw new VbaParseError('Format の書式は "###,##0" のように " で囲んでください', this.line);
      this.expectOp(")");
      return { type: "math_format", fields: { FORMAT: f.value }, inputs: { VALUE: { block: v } } };
    }
    if (name === "vbCrLf" || name === "vbNewLine") {
      return { type: "value_newline" };
    }
    if (name === "Int") {
      this.expectOp("(");
      const v = this.parseConcat();
      this.expectOp(")");
      return { type: "math_int", inputs: { VALUE: { block: v } } };
    }
    if (name === "Rnd") {
      if (this.peek() && this.peek().type === "OP" && this.peek().value === "(") {
        this.next();
        this.expectOp(")");
      }
      return { type: "math_rnd" };
    }
    if (name === "UBound" || name === "LBound") {
      this.expectOp("(");
      const a = this.next();
      if (!a || a.type !== "IDENT") throw new VbaParseError(`${name} には配列の名前を書いてください`, this.line);
      this.expectOp(")");
      return { type: "arr_bound", fields: { VAR: { __var: a.value }, KIND: name } };
    }
    // フォームの部品のプロパティ（TextBox1.Text / Label1.Caption）
    const dot = this.peek();
    const prop = this.tokens[this.pos + 1];
    if (dot && dot.type === "OP" && dot.value === "." && prop && prop.type === "IDENT" && /^(Caption|Text|Value)$/.test(prop.value)) {
      this.next();
      this.next();
      return { type: "ctl_get_prop", fields: { CONTROL: name, PROP: prop.value === "Caption" ? "Caption" : "Text" } };
    }

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
      const fields = { PROMPT: t.value };
      // 2番目（窓の題名）・3番目（はじめの答え）は、文字か数のときだけブロックで表せる
      for (const key of ["TITLE", "DEFAULT"]) {
        if (!(this.peek() && this.peek().type === "OP" && this.peek().value === ",")) break;
        this.next();
        const a = this.next();
        if (!a || (a.type !== "STR" && a.type !== "NUM")) {
          throw new VbaParseError("InputBox の2番目・3番目は、文字か数だけブロックで表せます", this.line);
        }
        fields[key] = String(a.value);
      }
      this.expectOp(")");
      return { type: "io_inputbox", fields };
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
    if (COUNTER_NAMES.includes(name)) {
      checkCounterRef(name, this.line);
      return { type: "loop_index" };
    }
    if (name === "True" || name === "False") {
      return { type: "value_bool", fields: { BOOL: name } };
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
    // 名前付き配列の要素：Umei(1) ／ Hyou(1, 2)
    if (this.peek() && this.peek().type === "OP" && this.peek().value === "(") {
      this.next();
      const idx = [this.parseConcat()];
      if (this.peek() && this.peek().type === "OP" && this.peek().value === ",") {
        this.next();
        idx.push(this.parseConcat());
      }
      this.expectOp(")");
      if (idx.length === 1) return { type: "arr_get_n", fields: { VAR: { __var: name } }, inputs: { INDEX: { block: idx[0] } } };
      return { type: "arr_get_2", fields: { VAR: { __var: name } }, inputs: { INDEX1: { block: idx[0] }, INDEX2: { block: idx[1] } } };
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
      const innermost = counterStack.length ? counterStack[counterStack.length - 1] : null;
      if (idTok && idTok.type === "IDENT" && innermost && idTok.value === innermost) return t.value + "{i}";
      if (!idTok || idTok.type !== "IDENT" || !COUNTER_NAMES.includes(idTok.value)) {
        throw new VbaParseError('動的なセル番地は "A" & i の形だけ対応しています', this.line);
      }
      checkCounterRef(idTok.value, this.line);
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
  m = t.match(/^"([A-Za-z]+)"\s*&\s*([A-Za-z_]\w*)$/);
  if (m && counterStack.length && counterStack[counterStack.length - 1] === m[2]) {
    return m[1] + "{i}"; // 一番内側のくり返しのカウンタ
  }
  if (m && COUNTER_NAMES.includes(m[2])) {
    checkCounterRef(m[2], line);
    return m[1] + "{i}";
  }
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
  DO: /^Do(\s+|$)/i,
  IF: /^If\s+/i,
  MSGBOX: /^MsgBox(\s+|\s*\()/i,
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
  if (/^Select\s+Case\s+/i.test(text)) return parseSelect(cursor);
  if (STMT_STARTERS.MSGBOX.test(text)) return parseMsgBox(cursor);
  return parseSimpleStatement(cursor);
}

// ブロックの「カウンタ」は一番内側のくり返しのカウンタを指す。
// 入れ子の深さに応じて i → j → k を使う（vba-generator.js と同じ決まり）
const COUNTER_NAMES = ["i", "j", "k"];
let counterStack = [];

// カウンタ名の参照がブロックで表せるか確かめる（外側のくり返しのカウンタは表せない）
function checkCounterRef(name, line) {
  if (counterStack.length === 0) {
    if (name !== "i") throw new VbaParseError(`${name} はくり返しのカウンタ用の名前です`, line);
    return name;
  }
  if (counterStack[counterStack.length - 1] !== name) {
    throw new VbaParseError(`外側のくり返しのカウンタ（${name}）は、ブロックでは表せません`, line);
  }
  return name;
}

function parseFor(cursor) {
  const { text, line } = cursor.next();
  const m = text.match(/^For\s+([A-Za-z_]\w*)\s*=\s*(.+?)\s+To\s+(.+?)(?:\s+Step\s+(.+))?$/i);
  if (!m) {
    throw new VbaParseError(
      '"For カウンタ = 開始 To 終わり" の形にしてください（例: For i = 1 To 5）',
      line
    );
  }
  const counter = m[1];
  const isAuto = COUNTER_NAMES.includes(counter);
  if (!isAuto) checkVarNameAssignable(counter, line);
  const start = parseExprString(m[2], line);
  const end = parseExprString(m[3], line);
  const step = m[4] ? parseExprString(m[4], line) : null;
  counterStack.push(counter);
  let body;
  try {
    body = parseStatements(cursor, ["Next"]);
  } finally {
    counterStack.pop();
  }
  if (cursor.atEnd()) throw new VbaParseError('"Next" が見つかりません（For に対応する終わり）', line);
  const closer = cursor.next();
  const nm = closer.text.match(/^Next(?:\s+([A-Za-z_]\w*))?$/i);
  if (!nm || (nm[1] && nm[1] !== counter)) {
    throw new VbaParseError('"Next" の書き方が違います', closer.line);
  }
  const inputs = { START: { block: start }, END: { block: end } };
  if (body) inputs.DO = { block: body };
  if (!isAuto) {
    // カウンタに好きな名前を付けた For
    inputs.STEP = step ? { block: step } : { shadow: { type: "value_number", fields: { NUM: 1 } } };
    return { type: "loop_for", fields: { VAR: { __var: counter } }, inputs };
  }
  if (step) {
    inputs.STEP = { block: step };
    return { type: "loop_for_step", inputs };
  }
  return { type: "loop_range", inputs };
}

// Do ... Loop While/Until（後判定）
function parseDoPost(cursor, line) {
  const body = parseStatements(cursor, ["Loop"]);
  if (cursor.atEnd()) throw new VbaParseError('"Loop" が見つかりません（Do に対応する終わり）', line);
  const closer = cursor.next();
  const m = closer.text.match(/^Loop\s+(While|Until)\s+(.+)$/i);
  if (!m) throw new VbaParseError('後判定は "Loop While 条件" か "Loop Until 条件" の形にしてください', closer.line);
  const cond = parseConditionString(m[2], closer.line);
  const inputs = { CONDITION: { block: cond } };
  if (body) inputs.DO = { block: body };
  const kind = /^while$/i.test(m[1]) ? "While" : "Until";
  return { type: "loop_do_post", fields: { KIND: kind }, inputs };
}

function parseDo(cursor) {
  const { text, line } = cursor.next();
  if (/^Do$/i.test(text)) return parseDoPost(cursor, line);
  const m = text.match(/^Do\s+(While|Until)\s+(.+)$/i);
  if (!m) {
    throw new VbaParseError('"Do While 条件" か "Do Until 条件"（または、後判定の "Do" ... "Loop While 条件"）の形にしてください', line);
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
  const thenBody = parseStatements(cursor, ["ElseIf", "Else", "End If"]);
  if (cursor.atEnd()) throw new VbaParseError('"End If" が見つかりません（If に対応する終わり）', line);
  const elseifs = [];
  while (!cursor.atEnd()) {
    const em = cursor.peek().text.match(/^ElseIf\s+(.+?)\s+Then$/i);
    if (!em) break;
    const l = cursor.next().line;
    const c = parseConditionString(em[1], l);
    const b = parseStatements(cursor, ["ElseIf", "Else", "End If"]);
    if (cursor.atEnd()) throw new VbaParseError('"End If" が見つかりません（ElseIf に対応する終わり）', l);
    elseifs.push({ c, b });
  }
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
  elseifs.forEach((e, i) => {
    inputs["ELSEIF_COND_" + i] = { block: e.c };
    if (e.b) inputs["ELSEIF_DO_" + i] = { block: e.b };
  });
  if (elseBody) inputs.ELSE = { block: elseBody };
  const spec = { type: "cond_if", inputs };
  if (elseifs.length) spec.extraState = { count: elseifs.length };
  return spec;
}

// Select Case 式 / Case ... / Case Else / End Select
function parseSelect(cursor) {
  const { text, line } = cursor.next();
  const m = text.match(/^Select\s+Case\s+(.+)$/i);
  if (!m) throw new VbaParseError('"Select Case 式" の形にしてください', line);
  const subject = parseExprString(m[1], line);
  const cases = [];
  let elseBody = null;
  for (;;) {
    if (cursor.atEnd()) throw new VbaParseError('"End Select" が見つかりません（Select Case に対応する終わり）', line);
    const t = cursor.peek();
    if (/^End\s+Select$/i.test(t.text)) {
      cursor.next();
      break;
    }
    const cm = t.text.match(/^Case\s+(.+)$/i);
    if (!cm) throw new VbaParseError('Select Case のすぐ下には "Case 値" を書いてください', t.line);
    cursor.next();
    const body = parseStatements(cursor, ["Case", "End Select"]);
    if (/^Else$/i.test(cm[1].trim())) {
      elseBody = body;
    } else {
      cases.push({ tests: cm[1].trim(), body });
    }
  }
  const inputs = { SUBJECT: { block: subject } };
  const fields = {};
  cases.forEach((c, i) => {
    fields["CASE_" + i] = c.tests;
    if (c.body) inputs["CASE_DO_" + i] = { block: c.body };
  });
  if (elseBody) inputs.ELSE = { block: elseBody };
  return { type: "cond_select", fields, inputs, extraState: { count: cases.length } };
}

function parseMsgBox(cursor) {
  const { text, line } = cursor.next();
  const m = text.match(/^MsgBox\s*(.+)$/i);
  const msg = parseExprString(m[1], line);
  return { type: "io_msgbox", inputs: { MSG: { block: msg } } };
}

// 「簡単な1行の文」たち：代入／メソッド呼び出し系
const SIMPLE_RULES = [
  // フォーム
  [/^Unload\s+Me$/i, () => ({ type: "form_unload" })],
  [/^End$/i, () => ({ type: "prog_end" })],
  [/^Exit\s+Sub$/i, () => ({ type: "prog_exit_sub" })],
  [/^Exit\s+(For|Do)$/i, (m) => ({ type: "loop_exit", fields: { WHAT: /^for$/i.test(m[1]) ? "For" : "Do" } })],
  // 命令のまとまり（Sub）の呼び出し
  [/^Call\s+([A-Za-z_]\w*)(\s*\(\s*\))?$/i, (m) => ({ type: "proc_call", fields: { NAME: m[1] } })],
  // 箱を用意する（Dim 名前 As 型）
  [
    /^Dim\s+([A-Za-z_]\w*)(?:\s+As\s+(Integer|Long|Single|Double|String|Variant))?$/i,
    (m, line) => {
      checkVarNameAssignable(m[1], line);
      const t = (m[2] || "Variant").toLowerCase();
      const TYPE = { integer: "Integer", long: "Long", single: "Single", double: "Double", string: "String", variant: "Variant" }[t];
      return { type: "var_dim", fields: { VAR: { __var: m[1] }, TYPE } };
    },
  ],
  [
    /^(?:Me\.)?([A-Za-z_]\w*)\.(Caption|Text|Value|Visible|Enabled)\s*=\s*(.+)$/,
    (m, line) => ({
      type: "ctl_set_prop",
      fields: { CONTROL: m[1], PROP: m[2] === "Value" ? "Text" : m[2] },
      inputs: { VALUE: { block: parseExprString(m[3], line) } },
    }),
  ],

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

  [/^Randomize$/i, () => ({ type: "rand_init" })],
  // 配列の箱を用意する（Dim Umei(5) As String ／ Dim Hyou(3, 4) As Integer）
  [
    /^Dim\s+([A-Za-z_]\w*)\s*\((.*)\)(?:\s+As\s+(Integer|Long|Single|Double|String|Variant))?$/i,
    (m, line) => {
      checkVarNameAssignable(m[1], line);
      const t = VBA_TYPE_NAMES[(m[3] || "Variant").toLowerCase()];
      return { type: "arr_dim", fields: { VAR: { __var: m[1] }, DIMS: m[2].trim(), TYPE: t } };
    },
  ],
  // 配列の要素に入れる：Umei(1) = "食料品" ／ Hyou(1, 2) = 5
  [
    /^([A-Za-z_]\w*)\s*\((.+)\)\s*=\s*(.+)$/,
    (m, line) => {
      const name = m[1];
      if (/^(Range|Cells|Rows|Worksheets|WorksheetFunction|InputBox|MsgBox|Val|Format|Int)$/i.test(name)) {
        throw new VbaParseError(`対応していない書き方です: "${m[0]}"`, line);
      }
      checkVarNameAssignable(name, line);
      const idx = splitTopLevel(m[2], ",");
      if (idx.length > 2) throw new VbaParseError("3次元以上の配列は、ブロックでは表せません", line);
      const value = { block: parseExprString(m[3], line) };
      if (idx.length === 1) {
        return { type: "arr_set_n", fields: { VAR: { __var: name } }, inputs: { INDEX: { block: parseExprString(idx[0], line) }, VALUE: value } };
      }
      return {
        type: "arr_set_2",
        fields: { VAR: { __var: name } },
        inputs: { INDEX1: { block: parseExprString(idx[0], line) }, INDEX2: { block: parseExprString(idx[1], line) }, VALUE: value },
      };
    },
  ],

  // 変数：name = name + delta → 増減ブロック（それ以外の代入より先に判定）
  [
    /^([A-Za-z_]\w*)\s*=\s*\1\s*\+\s*(.+)$/i,
    (m, line) => {
      checkVarNameAssignable(m[1], line);
      return { type: "var_change", fields: { VAR: { __var: m[1] } }, inputs: { DELTA: { block: parseExprString(m[2], line) } } };
    },
  ],
  // 変数：一般の代入
  [
    /^([A-Za-z_]\w*)\s*=\s*(.+)$/i,
    (m, line) => {
      checkVarNameAssignable(m[1], line);
      return { type: "var_set", fields: { VAR: { __var: m[1] } }, inputs: { VALUE: { block: parseExprString(m[2], line) } } };
    },
  ],
];

// 変数名として代入してよいかチェックする（ひらがな禁止・予約名禁止）。
// 予約名（i, arr）はループカウンタ・配列と衝突し、生成コードで Dim が
// 二重になってしまうため弾く（ブロックエディタ側の制限と揃える）
function checkVarNameAssignable(name, line) {
  if (COUNTER_NAMES.includes(name)) throw new VbaParseError(`カウンタ ${name} には代入できません`, line);
  if (name === "arr") throw new VbaParseError("arr は配列の名前として予約されているため、変数には使えません", line);
  if (HIRAGANA_RE.test(name)) throw new VbaParseError(`変数名にひらがなは使えません: ${name}`, line);
}

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
// 手続き（Sub 〜 End Sub）ごとに行を分ける
function splitProcedures(lines) {
  const procs = [];
  const outside = [];
  let cur = null;
  for (const l of lines) {
    const t = l.text;
    if (/^Option\s+Explicit$/i.test(t)) continue;
    const m = t.match(/^(?:(?:Private|Public)\s+)?Sub\s+([A-Za-z_]\w*)\s*\(\s*\)$/i);
    if (m) {
      if (cur) throw new VbaParseError(`「${cur.name}」の End Sub が見つかりません`, cur.line);
      cur = { name: m[1], line: l.line, body: [] };
      continue;
    }
    if (/^End\s+Sub$/i.test(t)) {
      if (!cur) throw new VbaParseError("End Sub に対応する Sub がありません", l.line);
      procs.push(cur);
      cur = null;
      continue;
    }
    if (cur) cur.body.push(l);
    else outside.push(l);
  }
  if (cur) throw new VbaParseError(`「${cur.name}」の End Sub が見つかりません`, cur.line);
  return { procs, outside };
}

// Dim の行を「📦 箱を用意する」ブロック用の行に直す。
// ブロックが自動で付ける Dim（カウンタ i・j・k、配列 arr、As Variant）は読み飛ばす
function normalizeDims(lines) {
  const out = [];
  for (const l of lines) {
    const m = l.text.match(/^Dim\s+(.+)$/i);
    if (!m) {
      out.push(l);
      continue;
    }
    for (const part of splitTopLevel(m[1], ",")) {
      const p = part.trim().match(/^([A-Za-z_]\w*)(\s*\(.*\))?(?:\s+As\s+(\w+))?$/i);
      if (!p) throw new VbaParseError(`Dim の書き方が読み取れません: "${l.text}"`, l.line);
      const [, name, dims, type] = p;
      if (dims) {
        if (name === "arr") continue; // ブロックが自動で付ける arr の Dim は読み飛ばす
        out.push({ text: `Dim ${name}${dims.trim()}${type ? " As " + type : ""}`, line: l.line });
        continue;
      }
      if (COUNTER_NAMES.includes(name)) continue;
      if (!type || /^Variant$/i.test(type)) continue;
      out.push({ text: `Dim ${name} As ${type}`, line: l.line });
    }
  }
  return out;
}

function parseBody(lines) {
  counterStack = [];
  const body = normalizeDims(lines);
  const cursor = new LineCursor(body);
  const root = parseStatements(cursor, []);
  if (!cursor.atEnd()) {
    const { text, line } = cursor.peek();
    throw new VbaParseError(`ここでの書き方が読み取れません: "${text}"`, line);
  }
  return root;
}

// 戻り値: { root, roots, variableNames }
// roots はワークスペースに並べるブロック（手続きごとに1つ）。root は従来どおり先頭の1つ
function parseVBA(source) {
  counterStack = [];
  const lines = preprocessLines(source);
  const { procs, outside } = splitProcedures(lines);
  const roots = [];
  if (!procs.length) {
    // Sub の中身だけが書かれている
    const root = parseBody(outside);
    if (root) roots.push(root);
  } else {
    const stray = outside.find((l) => !/^Dim\s+\S+/i.test(l.text));
    if (stray) throw new VbaParseError(`Sub の外に命令があります: "${stray.text}"`, stray.line);
    // Sub の外の Dim は「🌐 全体で使う箱」ブロックになる
    for (const l of outside) {
      const m = l.text.match(/^Dim\s+(.+)$/i);
      for (const part of splitTopLevel(m[1], ",")) {
        const p = part.trim().match(/^([A-Za-z_]\w*)\s*(?:\((.*)\))?(?:\s+As\s+(\w+))?$/i);
        if (!p) throw new VbaParseError(`Dim の書き方が読み取れません: "${l.text}"`, l.line);
        checkVarNameAssignable(p[1], l.line);
        const type = VBA_TYPE_NAMES[(p[3] || "Variant").toLowerCase()];
        if (!type) throw new VbaParseError(`型「${p[3]}」は、ブロックでは表せません`, l.line);
        roots.push({ type: "mod_dim", fields: { VAR: { __var: p[1] }, DIMS: (p[2] || "").trim(), TYPE: type } });
      }
    }
    let macroCount = 0;
    for (const p of procs) {
      const body = parseBody(p.body);
      const DO = body ? { DO: { block: body } } : undefined;
      if (/^UserForm_Initialize$/i.test(p.name)) {
        roots.push({ type: "form_initialize", inputs: DO });
        continue;
      }
      const click = p.name.match(/^([A-Za-z_]\w*)_Click$/i);
      if (click && !/^UserForm$/i.test(click[1])) {
        roots.push({ type: "form_button_click", fields: { CONTROL: click[1] }, inputs: DO });
        continue;
      }
      // MyMacro はブロックの「どこにも入っていないブロック」、それ以外の Sub は「📜 命令のまとまり」になる
      if (/^MyMacro$/i.test(p.name)) {
        macroCount++;
        if (macroCount > 1) throw new VbaParseError("MyMacro が2つあります", p.line);
        if (body) roots.push(body);
        continue;
      }
      roots.push({ type: "proc_sub", fields: { NAME: p.name }, inputs: DO });
    }
  }
  const variableNames = new Set();
  roots.forEach((r) => collectVariables(r, variableNames));
  return { root: roots[0] || null, roots, variableNames: [...variableNames] };
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

// {__var:name} プレースホルダを、Blockly が読める変数の指定にする。
// ID ではなく名前で指定する（Blockly.serialization.workspaces.load は読み込む前に変数を消すので、
// 先に createVariable しておいた ID は使えず、名前が i・j… に置きかわってしまうため）
function resolveVariables(block, workspace) {
  if (!block || typeof block !== "object") return;
  if (block.fields && block.fields.VAR && block.fields.VAR.__var) {
    block.fields.VAR = { name: block.fields.VAR.__var };
  }
  if (block.inputs) {
    for (const key in block.inputs) resolveVariables(block.inputs[key].block, workspace);
  }
  if (block.next) resolveVariables(block.next.block, workspace);
}

// ブロック（1つ、または配列）を Blockly.serialization.workspaces.load 用のJSONに包む
function toWorkspaceJson(rootOrRoots) {
  const list = Array.isArray(rootOrRoots) ? rootOrRoots : rootOrRoots ? [rootOrRoots] : [];
  return { blocks: { languageVersion: 0, blocks: list } };
}

// Node / ブラウザ両対応のエクスポート
const api = { parseVBA, resolveVariables, toWorkspaceJson, VbaParseError };
if (typeof module !== "undefined" && module.exports) {
  module.exports = api;
}
if (typeof window !== "undefined") {
  window.VbaParser = api;
}
