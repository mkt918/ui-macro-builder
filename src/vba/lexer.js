/* ===== VBA 字句解析（lexer） =====
 * VBA のソースをトークン列に分ける。DOM に依存しないので、ブラウザと Node の両方で使える。
 *
 * - 文の区切り（改行と ":"）は type "nl" のトークンにする
 * - 行の継続（行末の " _"）は読み飛ばす
 * - コメント（' と、文頭の Rem）は読み飛ばす
 * - 文字列・コメントの外にある全角の記号・数字・英字・スペースは、半角として読み進めつつ
 *   zenkaku リストに記録する（初学者が日本語入力のまま打ってしまうつまずき対策）
 * - 整数除算は "\"・"¥"（U+00A5）・全角の "￥" のどれも受け付ける（教科書は ¥ と表記）
 */
(function (root) {
  "use strict";

  // 全角 → 半角（文字列・コメントの外で使われたら警告する文字）
  const ZENKAKU_SYMBOLS = {
    "　": " ", // 全角スペース
    "＝": "=", "（": "(", "）": ")", "＋": "+", "－": "-", "−": "-", "＊": "*", "／": "/",
    "￥": "\\", "＼": "\\", "＜": "<", "＞": ">", "，": ",", "．": ".", "＆": "&", "＾": "^",
    "＃": "#", "：": ":", "＿": "_", "＂": '"', "”": '"', "“": '"', "’": "'", "‘": "'", "＇": "'",
  };

  function zenkakuToHankaku(ch) {
    if (Object.prototype.hasOwnProperty.call(ZENKAKU_SYMBOLS, ch)) return ZENKAKU_SYMBOLS[ch];
    const code = ch.charCodeAt(0);
    // 全角英数字（Ａ〜Ｚ・ａ〜ｚ・０〜９など U+FF01〜U+FF5E）
    if (code >= 0xff01 && code <= 0xff5e) return String.fromCharCode(code - 0xfee0);
    return null;
  }

  const DQUOTES = new Set(['"', "＂", "”", "“"]);

  function isLetter(ch) {
    return /[\p{L}]/u.test(ch);
  }
  function isIdentPart(ch) {
    return /[\p{L}\p{N}_]/u.test(ch);
  }
  function isDigit(ch) {
    return ch >= "0" && ch <= "9";
  }

  class VbaLexError extends Error {
    constructor(message, line, col, extra) {
      super(message);
      this.kind = "syntax";
      this.line = line;
      this.col = col;
      Object.assign(this, extra || {});
    }
  }

  // 戻り値: { tokens, zenkaku: [{line, col, ch, fix}] }
  function tokenize(src) {
    const tokens = [];
    const zenkaku = [];
    let i = 0;
    let line = 1;
    let lineStart = 0;
    let spaceBefore = true;
    const n = src.length;

    const col = (pos) => pos - lineStart + 1;
    const lastSignificant = () => tokens[tokens.length - 1];
    const atStatementStart = () => {
      const t = lastSignificant();
      return !t || t.type === "nl";
    };
    const push = (type, value, start, end, extra) => {
      const tok = { type, value, line, col: col(start), start, end, spaceBefore };
      if (extra) Object.assign(tok, extra);
      tokens.push(tok);
      spaceBefore = false;
      return tok;
    };
    const newline = (start) => {
      // 連続する改行は1つにまとめる
      const t = lastSignificant();
      if (t && t.type !== "nl") push("nl", "\n", start, start + 1);
      line++;
      lineStart = start + 1;
      spaceBefore = true;
    };
    const noteZenkaku = (pos, ch, fix) => {
      zenkaku.push({ line, col: col(pos), ch, fix, pos });
    };

    while (i < n) {
      let ch = src[i];

      // 改行
      if (ch === "\r") {
        if (src[i + 1] === "\n") i++;
        newline(i);
        i++;
        continue;
      }
      if (ch === "\n") {
        newline(i);
        i++;
        continue;
      }

      // 空白（全角スペースは警告して空白扱い）
      if (ch === " " || ch === "\t") {
        spaceBefore = true;
        i++;
        continue;
      }
      if (ch === "　") {
        noteZenkaku(i, ch, " ");
        spaceBefore = true;
        i++;
        continue;
      }

      // コメント
      if (ch === "'" || ch === "’" || ch === "‘" || ch === "＇") {
        if (ch !== "'") noteZenkaku(i, ch, "'");
        while (i < n && src[i] !== "\n" && src[i] !== "\r") i++;
        continue;
      }

      // 文字列
      if (DQUOTES.has(ch)) {
        if (ch !== '"') noteZenkaku(i, ch, '"');
        const start = i;
        const startLine = line;
        const startCol = col(i);
        const smart = ch !== '"';
        i++;
        let value = "";
        let closed = false;
        while (i < n) {
          const c = src[i];
          if (c === "\n" || c === "\r") break;
          const isClose = c === '"' || (smart && DQUOTES.has(c));
          if (isClose) {
            if (src[i + 1] === '"') {
              value += '"';
              i += 2;
              continue;
            }
            if (c !== '"') noteZenkaku(i, c, '"');
            i++;
            closed = true;
            break;
          }
          value += c;
          i++;
        }
        if (!closed) {
          throw new VbaLexError("文字列の終わりの \" がありません。", startLine, startCol, {
            friendly: "文字列（\" で囲んだ文字）の終わりの \" が見つかりません。",
            hint: "文字列は \"こんにちは\" のように、前と後ろを半角の \" で囲もう。",
          });
        }
        push("str", value, start, i);
        continue;
      }

      // 全角の記号・英数字は半角として読む（警告を記録）
      const fix = zenkakuToHankaku(ch);
      if (fix !== null && ch !== fix) {
        noteZenkaku(i, ch, fix);
        // 全角英字・数字は識別子・数値の一部として続けて読むため、ここで1文字ずつ置き換えて処理を進める
        src = src.slice(0, i) + fix + src.slice(i + 1);
        ch = fix;
        if (ch === '"' || ch === "'") continue; // 置き換え後に文字列・コメントとして読み直す
      }

      // 行の継続 " _"（直後が改行。後ろの空白は許す）
      if (ch === "_" && spaceBefore) {
        let j = i + 1;
        while (j < n && (src[j] === " " || src[j] === "\t")) j++;
        if (j >= n || src[j] === "\n" || src[j] === "\r") {
          i = j;
          if (src[i] === "\r") i++;
          if (src[i] === "\n") i++;
          line++;
          lineStart = i;
          spaceBefore = true;
          continue;
        }
      }

      // 数値
      if (isDigit(ch) || (ch === "." && isDigit(src[i + 1] || ""))) {
        const start = i;
        while (i < n && isDigit(src[i])) i++;
        if (src[i] === "." && isDigit(src[i + 1] || "")) {
          i++;
          while (i < n && isDigit(src[i])) i++;
        } else if (src[i] === "." && !isLetter(src[i + 1] || "")) {
          i++; // "1." のような書き方
        }
        if ((src[i] === "E" || src[i] === "e") && /[+\-]?\d/.test(src.slice(i + 1, i + 3))) {
          i++;
          if (src[i] === "+" || src[i] === "-") i++;
          while (i < n && isDigit(src[i])) i++;
        }
        const text = src.slice(start, i);
        push("num", Number(text), start, i, { text });
        continue;
      }

      // &H（16進）
      if (ch === "&" && (src[i + 1] === "H" || src[i + 1] === "h") && /[0-9A-Fa-f]/.test(src[i + 2] || "")) {
        const start = i;
        i += 2;
        while (i < n && /[0-9A-Fa-f]/.test(src[i])) i++;
        const text = src.slice(start, i);
        if (src[i] === "&") i++;
        push("num", parseInt(text.slice(2), 16), start, i, { text });
        continue;
      }

      // 識別子・キーワード
      if (isLetter(ch) || ch === "_") {
        const start = i;
        i++;
        while (i < n) {
          let c = src[i];
          const f = zenkakuToHankaku(c);
          if (f !== null && c !== f && /[A-Za-z0-9_]/.test(f)) {
            noteZenkaku(i, c, f);
            src = src.slice(0, i) + f + src.slice(i + 1);
            c = f;
          }
          if (!isIdentPart(c)) break;
          i++;
        }
        // 型宣言文字は Mid$ などの "$" だけ読む（"&" を認めると Kin&"円" を誤読するため）
        const name = src.slice(start, i);
        if (src[i] === "$") i++;
        // Rem コメント
        if (name.toLowerCase() === "rem" && atStatementStart()) {
          while (i < n && src[i] !== "\n" && src[i] !== "\r") i++;
          continue;
        }
        push("ident", name, start, i);
        continue;
      }

      // 演算子・区切り
      const two = src.slice(i, i + 2);
      if (two === "<=" || two === ">=" || two === "<>" || two === "=<" || two === "=>") {
        const norm = two === "=<" ? "<=" : two === "=>" ? ">=" : two;
        push("op", norm, i, i + 2);
        i += 2;
        continue;
      }
      if (ch === "¥") {
        push("op", "\\", i, i + 1);
        i++;
        continue;
      }
      if ("+-*/\\^&=<>(),.#".includes(ch)) {
        push("op", ch, i, i + 1);
        i++;
        continue;
      }
      if (ch === ":") {
        // ":" は文の区切り
        const t = lastSignificant();
        if (t && t.type !== "nl") push("nl", ":", i, i + 1);
        spaceBefore = true;
        i++;
        continue;
      }
      if (ch === ";") {
        push("op", ";", i, i + 1);
        i++;
        continue;
      }

      throw new VbaLexError(`使えない文字「${ch}」があります。`, line, col(i), {
        friendly: `「${ch}」はプログラムの中では使えない文字です。`,
        hint: "文字として表示したいときは \" で囲もう。",
      });
    }

    const t = lastSignificant();
    if (t && t.type !== "nl") push("nl", "\n", n, n);
    tokens.push({ type: "eof", value: null, line, col: col(n), start: n, end: n, spaceBefore: true });
    return { tokens, zenkaku };
  }

  // 文字列・コメントの外にある全角文字を半角に直したソースを返す（「半角に直す」ボタン用）
  function fixZenkaku(src) {
    const { zenkaku } = tokenizeLoose(src);
    if (!zenkaku.length) return src;
    let out = src;
    // 後ろから置き換えると位置がずれない
    zenkaku
      .slice()
      .sort((a, b) => b.pos - a.pos)
      .forEach((z) => {
        out = out.slice(0, z.pos) + z.fix + out.slice(z.pos + 1);
      });
    return out;
  }

  // 字句解析エラーが出ても、そこまでに見つかった全角文字を返す
  function tokenizeLoose(src) {
    try {
      return tokenize(src);
    } catch (e) {
      // エラー位置より前の全角文字だけでも返すため、もう一度途中まで調べる
      const zenkaku = [];
      try {
        const lines = src.split(/\r?\n/);
        let offset = 0;
        lines.forEach((ln, idx) => {
          let inStr = false;
          for (let k = 0; k < ln.length; k++) {
            const c = ln[k];
            if (!inStr && (c === "'" )) break;
            if (c === '"') {
              inStr = !inStr;
              continue;
            }
            if (inStr) continue;
            const f = zenkakuToHankaku(c);
            if (f !== null && f !== c) zenkaku.push({ line: idx + 1, col: k + 1, ch: c, fix: f, pos: offset + k });
          }
          offset += ln.length + 1;
        });
      } catch (_) {
        /* 何もしない */
      }
      return { tokens: [], zenkaku, error: e };
    }
  }

  const api = { tokenize, tokenizeLoose, fixZenkaku, VbaLexError, zenkakuToHankaku };
  if (typeof module !== "undefined" && module.exports) module.exports = api;
  else root.VbaLexer = api;
})(typeof window !== "undefined" ? window : globalThis);
