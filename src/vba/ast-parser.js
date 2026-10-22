/* ===== VBA 構文解析（AST を作る） =====
 * lexer のトークン列から AST（構文木）を作る。DOM に依存しない。
 *
 * 演算子の優先順位は教科書（第4章 p.93）の表に合わせる：
 *   ( ) → ^ → 単項マイナス → * / \ Mod（同じ順位は左から）→ + - → & → 比較 → Not → And → Or/Xor
 * （本物の VBA は * / → \ → Mod の順だが、このツールは教科書を優先する）
 *
 * すべてのノードは line / col（1始まり）と start / end（ソース上の位置）を持つ。
 * node.text でソースの該当部分を取り出せる（流れ図・説明文に使う）。
 */
(function (root) {
  "use strict";

  const Lexer =
    typeof module !== "undefined" && module.exports ? require("./lexer.js") : root.VbaLexer;

  class VbaSyntaxError extends Error {
    constructor(message, tok, extra) {
      super(message);
      this.kind = "syntax";
      this.line = tok ? tok.line : null;
      this.col = tok ? tok.col : null;
      Object.assign(this, extra || {});
    }
  }

  const TYPE_NAMES = {
    integer: "Integer", long: "Long", single: "Single", double: "Double", string: "String",
    variant: "Variant", boolean: "Boolean", date: "Date", currency: "Currency", byte: "Byte",
    object: "Object", range: "Object", worksheet: "Object", control: "Object", collection: "Object",
  };

  // 式の先頭に来てはいけない予約語（ここで止めて分かりやすいエラーにする）
  const NOT_EXPR = new Set([
    "then", "else", "elseif", "end", "to", "step", "next", "loop", "wend", "as", "dim", "if", "for",
    "do", "while", "until", "sub", "function", "case", "select", "exit", "call", "option", "private",
    "public", "const", "redim", "open", "close",
  ]);

  // 終わりのキーワードと、対応する始まり（エラー文用）
  const TERMINATOR_INFO = {
    endsub: { text: "End Sub", opener: "Sub" },
    endfunction: { text: "End Function", opener: "Function" },
    endif: { text: "End If", opener: "If" },
    else: { text: "Else", opener: "If" },
    elseif: { text: "ElseIf", opener: "If" },
    endselect: { text: "End Select", opener: "Select Case" },
    case: { text: "Case", opener: "Select Case" },
    next: { text: "Next", opener: "For" },
    loop: { text: "Loop", opener: "Do" },
    wend: { text: "Wend", opener: "While" },
  };

  // 開いている構文 → それを閉じる終わり
  const OPENER_END = {
    Sub: "End Sub", Function: "End Function", If: "End If", "Select Case": "End Select",
    For: "Next", Do: "Loop", While: "Wend",
  };

  class Parser {
    constructor(tokens, src) {
      this.toks = tokens;
      this.src = src;
      this.pos = 0;
      this.ctx = []; // 開いている構文のスタック [{opener, tok, ends:Set}]
    }

    peek(k = 0) {
      return this.toks[Math.min(this.pos + k, this.toks.length - 1)];
    }
    next() {
      const t = this.toks[this.pos];
      if (this.pos < this.toks.length - 1) this.pos++;
      this.lastEnd = t.end;
      return t;
    }
    isKw(tok, word) {
      return tok && tok.type === "ident" && tok.value.toLowerCase() === word;
    }
    peekKw(word, k = 0) {
      return this.isKw(this.peek(k), word);
    }
    isOp(tok, op) {
      return tok && tok.type === "op" && tok.value === op;
    }
    peekOp(op, k = 0) {
      return this.isOp(this.peek(k), op);
    }
    atEOS(k = 0) {
      const t = this.peek(k);
      return t.type === "nl" || t.type === "eof";
    }
    skipNl() {
      while (this.peek().type === "nl") this.next();
    }

    fail(message, tok, friendly, hint) {
      throw new VbaSyntaxError(message, tok || this.peek(), { friendly, hint });
    }

    expectKw(word, friendly, hint) {
      const t = this.peek();
      if (!this.isKw(t, word)) {
        this.fail(`${word} が必要です。`, t, friendly || `ここには「${word}」が必要です。`, hint);
      }
      return this.next();
    }
    expectOp(op, friendly, hint) {
      const t = this.peek();
      if (!this.isOp(t, op)) {
        this.fail(`${op} が必要です。`, t, friendly || `ここには「${op}」が必要です。`, hint);
      }
      return this.next();
    }
    expectIdent(friendly, hint) {
      const t = this.peek();
      if (t.type !== "ident") this.fail("識別子が必要です。", t, friendly || "ここには名前が必要です。", hint);
      return this.next();
    }
    expectEOS() {
      const t = this.peek();
      if (t.type === "nl" || t.type === "eof") return;
      const shown = t.type === "str" ? `"${t.value}"` : String(t.value);
      this.fail("コンパイル エラー: 構文エラー", t,
        `「${shown}」から先が読めません。`,
        "1行に1つの命令を書こう。記号や \" の閉じ忘れがないか確かめよう。");
    }

    node(type, startTok, props) {
      const n = Object.assign({ type, line: startTok.line, col: startTok.col, start: startTok.start }, props);
      n.end = this.lastEnd !== undefined ? this.lastEnd : startTok.end;
      return n;
    }
    text(node) {
      return this.src.slice(node.start, node.end);
    }

    // ---------- 終わりのキーワード ----------
    terminatorAt() {
      const t = this.peek();
      if (t.type !== "ident") return null;
      const w = t.value.toLowerCase();
      if (w === "end") {
        const t2 = this.peek(1);
        if (t2.type !== "ident") return null;
        const w2 = t2.value.toLowerCase();
        if (w2 === "sub") return "endsub";
        if (w2 === "function") return "endfunction";
        if (w2 === "if") return "endif";
        if (w2 === "select") return "endselect";
        return null;
      }
      if (w === "endif") return "endif";
      if (w === "else") return "else";
      if (w === "elseif") return "elseif";
      if (w === "case") return "case";
      if (w === "next") return "next";
      if (w === "loop") return "loop";
      if (w === "wend") return "wend";
      return null;
    }

    // allowed: 終わりのキーワード id の集合
    parseBlock(allowed) {
      const body = [];
      for (;;) {
        this.skipNl();
        const t = this.peek();
        if (t.type === "eof") this.missingEnd(t);
        const term = this.terminatorAt();
        if (term) {
          if (allowed.has(term)) return body;
          // 外側の構文の終わりなら「今の構文の閉じ忘れ」、どこにも無ければ「対応する始まりが無い」
          const outer = this.ctx.some((c) => c.ends.has(term));
          if (outer) this.missingEnd(t);
          const info = TERMINATOR_INFO[term];
          this.fail(`${info.text} に対応する ${info.opener} がありません。`, t,
            `「${info.text}」がありますが、その前に「${info.opener}」がありません。`,
            `「${info.opener}」を書き忘れていないか、または「${info.text}」が余分でないか確かめよう。`);
        }
        body.push(this.parseStatement());
      }
    }

    missingEnd(atTok) {
      const c = this.ctx[this.ctx.length - 1];
      if (!c) this.fail("コンパイル エラー: 構文エラー", atTok);
      const endText = OPENER_END[c.opener];
      throw new VbaSyntaxError(`${c.opener} に対応する ${endText} がありません。`, c.tok, {
        friendly: `${c.tok.line}行目の「${c.opener}」を閉じる「${endText}」が見つかりません。`,
        hint: `「${c.opener}」で始めたら、最後に「${endText}」を書こう。`,
      });
    }

    withCtx(opener, tok, ends, fn) {
      this.ctx.push({ opener, tok, ends: new Set(ends) });
      try {
        return fn();
      } finally {
        this.ctx.pop();
      }
    }

    // ---------- モジュール ----------
    parseModule() {
      const prog = { type: "Program", optionExplicit: false, optionBase: 0, decls: [], procs: [] };
      const procNames = new Map();
      for (;;) {
        this.skipNl();
        const t = this.peek();
        if (t.type === "eof") break;
        if (this.isKw(t, "option")) {
          this.next();
          if (this.peekKw("explicit")) {
            this.next();
            prog.optionExplicit = true;
          } else if (this.peekKw("base")) {
            this.next();
            const n = this.next();
            prog.optionBase = n.type === "num" ? n.value : 0;
          } else if (this.peekKw("compare")) {
            this.next();
            this.next();
          } else {
            this.fail("コンパイル エラー: 構文エラー", this.peek(), "Option のあとには Explicit を書こう。");
          }
          this.expectEOS();
          continue;
        }
        if (this.isProcStart()) {
          const proc = this.parseProc();
          const key = proc.name.toLowerCase();
          if (procNames.has(key)) {
            throw new VbaSyntaxError(`名前が適切ではありません: ${proc.name}`, { line: proc.line, col: proc.col }, {
              friendly: `「${proc.name}」という名前の手続き（Sub / Function）が2つあります。`,
              hint: "同じ名前の Sub は1つだけにしよう。片方の名前を変えるか、消そう。",
            });
          }
          procNames.set(key, proc);
          prog.procs.push(proc);
          continue;
        }
        if ((this.isKw(t, "private") || this.isKw(t, "public")) && this.peekKw("const", 1)) {
          this.next();
          prog.decls.push(this.parseConst("module"));
          this.expectEOS();
          continue;
        }
        if (this.isKw(t, "dim") || this.isKw(t, "private") || this.isKw(t, "public") || this.isKw(t, "static")) {
          prog.decls.push(this.parseDim("module"));
          this.expectEOS();
          continue;
        }
        if (this.isKw(t, "const")) {
          prog.decls.push(this.parseConst("module"));
          this.expectEOS();
          continue;
        }
        const term = this.terminatorAt();
        if (term === "endsub" || term === "endfunction") {
          const info = TERMINATOR_INFO[term];
          this.fail(`${info.text} に対応する ${info.opener} がありません。`, t,
            `「${info.text}」がありますが、その前に「Sub」の行がありません。`,
            "Sub の行を消してしまっていないか確かめよう。");
        }
        this.fail("コンパイル エラー: プロシージャの外では無効です。", t,
          "命令が Sub の外に書かれています。",
          "命令は「Private Sub CommandButton1_Click()」と「End Sub」の間に書こう。");
      }
      return prog;
    }

    isProcStart() {
      let k = 0;
      while (this.peekKw("private", k) || this.peekKw("public", k) || this.peekKw("static", k)) k++;
      return this.peekKw("sub", k) || this.peekKw("function", k);
    }

    parseProc() {
      const startTok = this.peek();
      let isPrivate = false;
      while (this.peekKw("private") || this.peekKw("public") || this.peekKw("static")) {
        if (this.peekKw("private")) isPrivate = true;
        this.next();
      }
      const kindTok = this.next();
      const kind = kindTok.value.toLowerCase() === "sub" ? "Sub" : "Function";
      const nameTok = this.expectIdent(`${kind} のあとには名前が必要です。`, "例: Sub Hello()");
      const params = [];
      if (this.peekOp("(")) {
        this.next();
        if (!this.peekOp(")")) {
          for (;;) {
            params.push(this.parseParam());
            if (this.peekOp(",")) {
              this.next();
              continue;
            }
            break;
          }
        }
        this.expectOp(")", "引数のかっこ「)」が閉じていません。");
      }
      let returnType = null;
      if (kind === "Function" && this.peekKw("as")) {
        this.next();
        returnType = this.parseTypeName();
      }
      this.expectEOS();
      const headerEnd = this.lastEnd;
      const endId = kind === "Sub" ? "endsub" : "endfunction";
      const body = this.withCtx(kind, startTok, [endId], () => this.parseBlock(new Set([endId])));
      const endTok = this.next(); // End
      this.next(); // Sub / Function
      this.expectEOS();
      const proc = this.node("Proc", startTok, {
        kind, name: nameTok.value, isPrivate, params, returnType, body,
        endLine: endTok.line, headerEnd,
      });
      return proc;
    }

    parseParam() {
      const startTok = this.peek();
      let optional = false;
      let byRef = true;
      if (this.peekKw("optional")) {
        this.next();
        optional = true;
      }
      if (this.peekKw("byval")) {
        this.next();
        byRef = false;
      } else if (this.peekKw("byref")) {
        this.next();
      }
      const nameTok = this.expectIdent("引数の名前が必要です。");
      let isArray = false;
      if (this.peekOp("(")) {
        this.next();
        this.expectOp(")");
        isArray = true;
      }
      let varType = "Variant";
      if (this.peekKw("as")) {
        this.next();
        varType = this.parseTypeName();
      }
      let defaultValue = null;
      if (this.peekOp("=")) {
        this.next();
        defaultValue = this.parseExpr();
      }
      return this.node("Param", startTok, { name: nameTok.value, byRef, optional, isArray, varType, defaultValue });
    }

    parseTypeName() {
      if (this.peekKw("new")) this.next();
      const t = this.expectIdent("As のあとには型の名前が必要です。", "例: As Integer / As String");
      const canon = TYPE_NAMES[t.value.toLowerCase()];
      if (!canon) {
        // MSForms.Label のような書き方も Object として扱う
        while (this.peekOp(".")) {
          this.next();
          this.expectIdent();
        }
        return "Object";
      }
      // String * 10（固定長）は長さを読み飛ばす
      if (canon === "String" && this.peekOp("*")) {
        this.next();
        this.next();
      }
      return canon;
    }

    // ---------- 宣言 ----------
    parseDim(scope) {
      const startTok = this.next(); // Dim / Private / Public / Static
      const keyword = startTok.value.toLowerCase();
      const vars = [];
      for (;;) {
        if (this.peekKw("withevents")) this.next();
        const nameTok = this.expectIdent("Dim のあとには変数の名前が必要です。", "例: Dim Tanka As Integer");
        if (NOT_EXPR.has(nameTok.value.toLowerCase())) {
          this.fail("名前が適切ではありません。", nameTok,
            `「${nameTok.value}」は VBA の決まった言葉なので、変数の名前には使えません。`,
            "別の名前（例: Kazu, Goukei）にしよう。");
        }
        let dims = null;
        if (this.peekOp("(")) {
          this.next();
          dims = [];
          if (!this.peekOp(")")) {
            for (;;) {
              let lower = null;
              let upper = this.parseExpr();
              if (this.peekKw("to")) {
                this.next();
                lower = upper;
                upper = this.parseExpr();
              }
              dims.push({ lower, upper });
              if (this.peekOp(",")) {
                this.next();
                continue;
              }
              break;
            }
          }
          this.expectOp(")", "配列のかっこ「)」が閉じていません。");
        }
        let varType = "Variant";
        let explicitType = false;
        if (this.peekKw("as")) {
          this.next();
          varType = this.parseTypeName();
          explicitType = true;
        }
        vars.push({ name: nameTok.value, dims, varType, explicitType, line: nameTok.line, col: nameTok.col });
        if (this.peekOp(",")) {
          this.next();
          continue;
        }
        break;
      }
      return this.node("Dim", startTok, { scope, keyword, isStatic: keyword === "static", vars });
    }

    parseConst(scope) {
      const startTok = this.next();
      const items = [];
      for (;;) {
        const nameTok = this.expectIdent("Const のあとには名前が必要です。", "例: Const Tanka As Integer = 100");
        let varType = null;
        if (this.peekKw("as")) {
          this.next();
          varType = this.parseTypeName();
        }
        this.expectOp("=", "Const には「= 値」が必要です。", "例: Const Tanka As Integer = 100");
        const value = this.parseExpr();
        items.push({ name: nameTok.value, varType, value, line: nameTok.line });
        if (this.peekOp(",")) {
          this.next();
          continue;
        }
        break;
      }
      return this.node("Const", startTok, { scope, items });
    }

    parseReDim() {
      const startTok = this.next();
      let preserve = false;
      if (this.peekKw("preserve")) {
        this.next();
        preserve = true;
      }
      const vars = [];
      for (;;) {
        const nameTok = this.expectIdent("ReDim のあとには配列の名前が必要です。");
        this.expectOp("(", "ReDim では「配列名(大きさ)」のように書こう。");
        const dims = [];
        for (;;) {
          let lower = null;
          let upper = this.parseExpr();
          if (this.peekKw("to")) {
            this.next();
            lower = upper;
            upper = this.parseExpr();
          }
          dims.push({ lower, upper });
          if (this.peekOp(",")) {
            this.next();
            continue;
          }
          break;
        }
        this.expectOp(")");
        let varType = null;
        if (this.peekKw("as")) {
          this.next();
          varType = this.parseTypeName();
        }
        vars.push({ name: nameTok.value, dims, varType });
        if (this.peekOp(",")) {
          this.next();
          continue;
        }
        break;
      }
      return this.node("ReDim", startTok, { preserve, vars });
    }

    // ---------- 文 ----------
    parseStatement() {
      const t = this.peek();
      if (t.type !== "ident") {
        if (this.isOp(t, ".")) {
          this.fail("コンパイル エラー: 構文エラー", t, "With はこのツールではまだ使えません。", "「Label1.Caption = ...」のように、部品の名前から書こう。");
        }
        this.fail("コンパイル エラー: 構文エラー", t,
          "この行は命令として読めません。",
          "命令は「変数 = 値」や「MsgBox \"…\"」のような形で書こう。");
      }
      const w = t.value.toLowerCase();
      let stmt;
      switch (w) {
        case "dim":
        case "static":
          stmt = this.parseDim("local");
          break;
        case "const":
          stmt = this.parseConst("local");
          break;
        case "redim":
          stmt = this.parseReDim();
          break;
        case "if":
          return this.parseIf();
        case "select":
          return this.parseSelect();
        case "for":
          return this.parseFor();
        case "do":
          return this.parseDo();
        case "while":
          return this.parseWhile();
        case "exit": {
          this.next();
          const what = this.peek();
          const kw = what.type === "ident" ? what.value.toLowerCase() : "";
          const map = { sub: "Sub", function: "Function", for: "For", do: "Do" };
          if (!map[kw]) {
            this.fail("コンパイル エラー: 構文エラー", what, "Exit のあとには Sub / For / Do のどれかを書こう。", "例: Exit Sub");
          }
          this.next();
          stmt = this.node("Exit", t, { what: map[kw] });
          if (map[kw] === "For" && !this.ctx.some((c) => c.opener === "For")) {
            this.fail("Exit For は For...Next の外では使えません。", t, "Exit For は For ～ Next の中でだけ使えます。");
          }
          if (map[kw] === "Do" && !this.ctx.some((c) => c.opener === "Do")) {
            this.fail("Exit Do は Do...Loop の外では使えません。", t, "Exit Do は Do ～ Loop の中でだけ使えます。");
          }
          break;
        }
        case "end":
          if (this.atEOS(1)) {
            this.next();
            stmt = this.node("End", t, {});
          } else {
            const t2 = this.peek(1);
            this.fail("コンパイル エラー: 構文エラー", t2,
              `「End ${t2.value}」は使えない書き方です。`,
              "終わりは「End If」「End Sub」「End Select」のように書こう。");
          }
          break;
        case "call": {
          this.next();
          const callee = this.parseChain({ stopAtSpacedParen: false });
          stmt = this.node("CallStmt", t, { callee, args: null, explicitCall: true });
          break;
        }
        case "open":
          stmt = this.parseOpen();
          break;
        case "close":
          stmt = this.parseClose();
          break;
        case "input":
          if (this.peekOp("#", 1)) {
            stmt = this.parseInputFile();
            break;
          }
          stmt = this.parseAssignOrCall();
          break;
        case "line":
          if (this.peekKw("input", 1)) {
            stmt = this.parseLineInput();
            break;
          }
          stmt = this.parseAssignOrCall();
          break;
        case "print":
        case "write":
          if (this.peekOp("#", 1)) {
            stmt = this.parsePrintFile(w === "write");
            break;
          }
          stmt = this.parseAssignOrCall();
          break;
        case "let":
        case "set":
          this.next();
          stmt = this.parseAssignOrCall(true);
          break;
        case "on":
          this.fail("このツールでは On Error はまだ使えません。", t, "On Error はこのツールではまだ使えません。", "エラーが起きないように、If で先に確かめよう。");
          break;
        case "goto":
        case "gosub":
          this.fail("このツールでは GoTo はまだ使えません。", t, "GoTo はこのツールでは使えません。", "If や Do ～ Loop を使って書き直そう。");
          break;
        case "with":
          this.fail("このツールでは With はまだ使えません。", t, "With はこのツールではまだ使えません。", "「Label1.Caption = ...」のように、部品の名前から書こう。");
          break;
        case "sub":
        case "function":
        case "private":
        case "public":
        case "option":
          this.fail("コンパイル エラー: 構文エラー", t,
            "Sub の中に、別の Sub や Option が書かれています。",
            "前の Sub の最後に「End Sub」を書き忘れていないか確かめよう。");
          break;
        default:
          stmt = this.parseAssignOrCall();
      }
      this.expectEOS();
      return stmt;
    }

    // 1行 If の中など、":" で区切った文を同じ行で読む
    parseInlineStatements(stopAtElse) {
      const list = [];
      for (;;) {
        if (stopAtElse && this.peekKw("else")) break;
        const t = this.peek();
        if (t.type === "nl" || t.type === "eof") break;
        list.push(this.parseSimpleInline());
        const sep = this.peek();
        if (sep.type === "nl" && sep.value === ":") {
          this.next();
          continue;
        }
        break;
      }
      return list;
    }

    // 1行 If の中で使える文（ブロック構文は除く）
    parseSimpleInline() {
      const t = this.peek();
      const w = t.type === "ident" ? t.value.toLowerCase() : "";
      if (w === "exit") {
        this.next();
        const what = this.next();
        const map = { sub: "Sub", function: "Function", for: "For", do: "Do" };
        const m = map[String(what.value).toLowerCase()];
        if (!m) this.fail("コンパイル エラー: 構文エラー", what, "Exit のあとには Sub / For / Do のどれかを書こう。");
        return this.node("Exit", t, { what: m });
      }
      if (w === "end" && (this.atEOS(1) || this.peekKw("else", 1) || this.peek(1).value === ":")) {
        this.next();
        return this.node("End", t, {});
      }
      if (w === "call") {
        this.next();
        const callee = this.parseChain({ stopAtSpacedParen: false });
        return this.node("CallStmt", t, { callee, args: null, explicitCall: true });
      }
      if (["if", "for", "do", "while", "select", "dim"].includes(w)) {
        this.fail("コンパイル エラー: 構文エラー", t,
          "1行の If の中には、この命令は書けません。",
          "If ～ Then のあとで改行して、End If までの間に書こう。");
      }
      return this.parseAssignOrCall();
    }

    parseAssignOrCall(requireAssign) {
      const startTok = this.peek();
      const pos0 = this.pos;
      const target = this.parseChain({ stopAtSpacedParen: false });
      if (this.peekOp("=")) {
        this.next();
        const value = this.parseExpr();
        return this.node("Assign", startTok, { target, value });
      }
      if (requireAssign) {
        this.fail("コンパイル エラー: 構文エラー", this.peek(), "「=」が必要です。", "例: Set x = Range(\"A1\")");
      }
      const t = this.peek();
      const endOfStmt = t.type === "nl" || t.type === "eof" || this.isKw(t, "else");
      if (endOfStmt) {
        if (target.type === "Num" || target.type === "Str" || target.type === "Paren") {
          this.fail("コンパイル エラー: 構文エラー", startTok, "値だけが書かれていて、何をするかがわかりません。", "例: Kotae = 5 のように「=」で入れる先を書こう。");
        }
        return this.node("CallStmt", startTok, { callee: target, args: null, explicitCall: false });
      }
      // 「名前 引数, 引数」の形（MsgBox "こんにちは" / Range("A1").Copy Range("B1")）
      this.pos = pos0;
      const callee = this.parseChain({ stopAtSpacedParen: true });
      if (callee.type !== "Ident" && callee.type !== "Member") {
        this.fail("コンパイル エラー: 構文エラー", this.peek(),
          "この行は命令として読めません。",
          "「=」を書き忘れていないか確かめよう。例: Kin = Tanka * Suu");
      }
      const args = [];
      for (;;) {
        if (this.peekOp(",")) {
          args.push(null);
          this.next();
          continue;
        }
        args.push(this.parseExpr());
        if (this.peekOp(",")) {
          this.next();
          if (this.atEOS()) args.push(null);
          continue;
        }
        break;
      }
      return this.node("CallStmt", startTok, { callee, args, explicitCall: false });
    }

    parseIf() {
      const ifTok = this.next();
      const cond = this.parseCondition("If");
      this.expectKw("then", "If の条件のあとに「Then」が必要です。", "例: If Omosa > 50 Then");
      if (!this.atEOS()) {
        // 1行 If
        const thenBody = this.parseInlineStatements(true);
        let elseBody = null;
        if (this.peekKw("else")) {
          this.next();
          elseBody = this.parseInlineStatements(false);
        }
        this.expectEOS();
        return this.node("If", ifTok, { branches: [{ cond, body: thenBody, line: ifTok.line }], elseBody, singleLine: true, endLine: ifTok.line });
      }
      return this.withCtx("If", ifTok, ["endif", "else", "elseif"], () => {
        const branches = [];
        let elseBody = null;
        let elseLine = null;
        const body = this.parseBlock(new Set(["elseif", "else", "endif"]));
        branches.push({ cond, body, line: ifTok.line });
        for (;;) {
          const term = this.terminatorAt();
          const t = this.peek();
          if (term === "elseif") {
            if (elseBody) this.fail("ElseIf は Else より前に書こう。", t, "ElseIf は Else より前に書こう。");
            this.next();
            const c = this.parseCondition("ElseIf");
            this.expectKw("then", "ElseIf の条件のあとに「Then」が必要です。");
            this.expectEOS();
            const b = this.parseBlock(new Set(["elseif", "else", "endif"]));
            branches.push({ cond: c, body: b, line: t.line });
            continue;
          }
          if (term === "else") {
            if (elseBody) this.fail("Else が2つあります。", t, "1つの If の中に Else は1つだけ書けます。");
            this.next();
            elseLine = t.line;
            if (!this.atEOS()) this.expectEOS();
            elseBody = this.parseBlock(new Set(["endif", "else", "elseif"]));
            continue;
          }
          // End If
          const endTok = this.next();
          if (endTok.value.toLowerCase() === "end") this.next();
          this.expectEOS();
          return this.node("If", ifTok, { branches, elseBody, elseLine, singleLine: false, endLine: endTok.line });
        }
      });
    }

    parseCondition(where) {
      const t = this.peek();
      if (this.isKw(t, "then") || this.atEOS()) {
        this.fail("コンパイル エラー: 構文エラー", t, `${where} のあとに条件が必要です。`, "例: If Omosa > 50 Then");
      }
      return this.parseExpr();
    }

    parseSelect() {
      const selTok = this.next();
      this.expectKw("case", "Select のあとには「Case」が必要です。", "例: Select Case Bango");
      const expr = this.parseExpr();
      this.expectEOS();
      return this.withCtx("Select Case", selTok, ["endselect", "case"], () => {
        const cases = [];
        let elseBody = null;
        for (;;) {
          this.skipNl();
          const t = this.peek();
          if (t.type === "eof") this.missingEnd(t);
          const term = this.terminatorAt();
          if (term === "endselect") {
            this.next();
            this.next();
            this.expectEOS();
            return this.node("Select", selTok, { expr, cases, elseBody, endLine: t.line });
          }
          if (term !== "case") {
            if (cases.length === 0 && !term) {
              this.fail("コンパイル エラー: 構文エラー", t, "Select Case のすぐ下には「Case 値」を書こう。");
            }
            if (term) {
              const outer = this.ctx.some((c) => c.ends.has(term));
              if (outer) this.missingEnd(t);
            }
            this.fail("コンパイル エラー: 構文エラー", t);
          }
          this.next(); // Case
          if (this.peekKw("else")) {
            this.next();
            this.expectEOS();
            elseBody = this.parseBlock(new Set(["case", "endselect"]));
            continue;
          }
          const tests = [];
          for (;;) {
            const tt = this.peek();
            if (this.peekKw("is")) {
              this.next();
              const opTok = this.next();
              if (opTok.type !== "op" || !["=", "<>", "<", ">", "<=", ">="].includes(opTok.value)) {
                this.fail("コンパイル エラー: 構文エラー", opTok, "Case Is のあとには比較の記号（> や <= など）を書こう。", "例: Case Is >= 80");
              }
              tests.push({ kind: "is", op: opTok.value, expr: this.parseConcat(), line: tt.line });
            } else if (this.peek().type === "op" && ["<", ">", "<=", ">=", "<>", "="].includes(this.peek().value)) {
              const opTok = this.next();
              tests.push({ kind: "is", op: opTok.value, expr: this.parseConcat(), line: tt.line });
            } else {
              const a = this.parseExpr();
              if (this.peekKw("to")) {
                this.next();
                const b = this.parseExpr();
                tests.push({ kind: "range", from: a, to: b, line: tt.line });
              } else {
                tests.push({ kind: "value", expr: a, line: tt.line });
              }
            }
            if (this.peekOp(",")) {
              this.next();
              continue;
            }
            break;
          }
          this.expectEOS();
          const body = this.parseBlock(new Set(["case", "endselect"]));
          cases.push({ tests, body, line: t.line });
        }
      });
    }

    parseFor() {
      const forTok = this.next();
      if (this.peekKw("each")) {
        this.fail("このツールでは For Each はまだ使えません。", this.peek(), "For Each はこのツールではまだ使えません。", "For i = 1 To 10 の形で書こう。");
      }
      const varTok = this.expectIdent("For のあとにはカウンタ変数の名前が必要です。", "例: For i = 1 To 10");
      const varNode = this.node("Ident", varTok, { name: varTok.value });
      this.expectOp("=", "For のカウンタ変数のあとに「=」が必要です。", "例: For i = 1 To 10");
      const from = this.parseExpr();
      this.expectKw("to", "For の始まりの値のあとに「To」が必要です。", "例: For i = 1 To 10");
      const to = this.parseExpr();
      let step = null;
      if (this.peekKw("step")) {
        this.next();
        step = this.parseExpr();
      }
      this.expectEOS();
      return this.withCtx("For", forTok, ["next"], () => {
        const body = this.parseBlock(new Set(["next"]));
        const nextTok = this.next();
        if (this.peek().type === "ident") {
          const v = this.next();
          if (v.value.toLowerCase() !== varTok.value.toLowerCase()) {
            this.fail("Next の制御変数が正しくありません。", v,
              `「Next ${v.value}」は「For ${varTok.value}」と名前が合っていません。`,
              `「Next ${varTok.value}」にしよう。For が入れ子のときは、内側から先に閉じよう。`);
          }
        }
        this.expectEOS();
        return this.node("For", forTok, { var: varNode, from, to, step, body, endLine: nextTok.line });
      });
    }

    parseDo() {
      const doTok = this.next();
      let pre = null;
      if (this.peekKw("while") || this.peekKw("until")) {
        const k = this.next();
        pre = { kind: k.value.toLowerCase() === "while" ? "While" : "Until", expr: this.parseCondition(`Do ${k.value}`) };
      }
      this.expectEOS();
      return this.withCtx("Do", doTok, ["loop"], () => {
        const body = this.parseBlock(new Set(["loop"]));
        const loopTok = this.next();
        let post = null;
        if (this.peekKw("while") || this.peekKw("until")) {
          const k = this.next();
          if (pre) this.fail("コンパイル エラー: 構文エラー", k, "条件は Do の行か Loop の行のどちらか一方に書こう。");
          post = { kind: k.value.toLowerCase() === "while" ? "While" : "Until", expr: this.parseCondition(`Loop ${k.value}`) };
        }
        this.expectEOS();
        return this.node("Do", doTok, { pre, post, body, endLine: loopTok.line });
      });
    }

    parseWhile() {
      const wTok = this.next();
      const cond = this.parseCondition("While");
      this.expectEOS();
      return this.withCtx("While", wTok, ["wend"], () => {
        const body = this.parseBlock(new Set(["wend"]));
        const wendTok = this.next();
        this.expectEOS();
        return this.node("Do", wTok, { pre: { kind: "While", expr: cond }, post: null, body, endLine: wendTok.line, isWend: true });
      });
    }

    parseFileNum() {
      if (this.peekOp("#")) this.next();
      return this.parseExpr();
    }

    parseOpen() {
      const t = this.next();
      const path = this.parseExpr();
      this.expectKw("for", "Open のファイル名のあとに「For Input」などが必要です。", "例: Open ThisWorkbook.Path & \"¥uriage.csv\" For Input As #1");
      const modeTok = this.expectIdent("For のあとには Input / Output / Append のどれかを書こう。");
      const modes = { input: "Input", output: "Output", append: "Append" };
      const mode = modes[modeTok.value.toLowerCase()];
      if (!mode) {
        this.fail("コンパイル エラー: 構文エラー", modeTok, "For のあとには Input（読む）か Output（書く）を書こう。");
      }
      if (this.peekKw("access")) {
        this.next();
        while (this.peek().type === "ident" && !this.peekKw("as")) this.next();
      }
      this.expectKw("as", "Open の最後に「As #1」のようにファイル番号が必要です。", "例: Open \"uriage.csv\" For Input As #1");
      const fileNum = this.parseFileNum();
      return this.node("Open", t, { path, mode, fileNum });
    }

    parseClose() {
      const t = this.next();
      const fileNums = [];
      while (!this.atEOS()) {
        fileNums.push(this.parseFileNum());
        if (this.peekOp(",")) this.next();
        else break;
      }
      return this.node("Close", t, { fileNums });
    }

    parseInputFile() {
      const t = this.next(); // Input
      const fileNum = this.parseFileNum();
      this.expectOp(",", "Input #1 のあとに「,」と変数名が必要です。", "例: Input #1, KouMei, Suu");
      const targets = [];
      for (;;) {
        targets.push(this.parseChain({ stopAtSpacedParen: false }));
        if (this.peekOp(",")) {
          this.next();
          continue;
        }
        break;
      }
      return this.node("InputFile", t, { fileNum, targets });
    }

    parseLineInput() {
      const t = this.next(); // Line
      this.next(); // Input
      const fileNum = this.parseFileNum();
      this.expectOp(",", "Line Input #1 のあとに「,」と変数名が必要です。");
      const target = this.parseChain({ stopAtSpacedParen: false });
      return this.node("LineInput", t, { fileNum, target });
    }

    parsePrintFile(isWrite) {
      const t = this.next();
      const fileNum = this.parseFileNum();
      const items = [];
      if (this.peekOp(",")) {
        this.next();
        while (!this.atEOS()) {
          const expr = this.parseExpr();
          let sep = null;
          if (this.peekOp(";") || this.peekOp(",")) sep = this.next().value;
          items.push({ expr, sep });
          if (!sep) break;
        }
      }
      return this.node(isWrite ? "WriteFile" : "PrintFile", t, { fileNum, items });
    }

    // ---------- 式 ----------
    parseExpr() {
      return this.parseOr();
    }
    bin(op, left, right, opTok) {
      const n = { type: "Bin", op, left, right, line: left.line, col: left.col, start: left.start, end: right.end, opLine: opTok.line };
      return n;
    }
    parseOr() {
      let left = this.parseAnd();
      while (this.peekKw("or") || this.peekKw("xor")) {
        const opTok = this.next();
        const right = this.parseAnd();
        left = this.bin(opTok.value.toLowerCase() === "or" ? "Or" : "Xor", left, right, opTok);
      }
      return left;
    }
    parseAnd() {
      let left = this.parseNot();
      while (this.peekKw("and")) {
        const opTok = this.next();
        const right = this.parseNot();
        left = this.bin("And", left, right, opTok);
      }
      return left;
    }
    parseNot() {
      if (this.peekKw("not")) {
        const t = this.next();
        const operand = this.parseNot();
        return { type: "Unary", op: "Not", operand, line: t.line, col: t.col, start: t.start, end: operand.end };
      }
      return this.parseCompare();
    }
    parseCompare() {
      let left = this.parseConcat();
      for (;;) {
        const t = this.peek();
        if (t.type === "op" && ["=", "<>", "<", ">", "<=", ">="].includes(t.value)) {
          this.next();
          const right = this.parseConcat();
          left = this.bin(t.value, left, right, t);
          continue;
        }
        if (this.isKw(t, "like") || this.isKw(t, "is")) {
          this.next();
          const right = this.parseConcat();
          left = this.bin(t.value.toLowerCase() === "like" ? "Like" : "Is", left, right, t);
          continue;
        }
        return left;
      }
    }
    parseConcat() {
      let left = this.parseAdd();
      while (this.peekOp("&")) {
        const t = this.next();
        const right = this.parseAdd();
        left = this.bin("&", left, right, t);
      }
      return left;
    }
    parseAdd() {
      let left = this.parseMul();
      while (this.peekOp("+") || this.peekOp("-")) {
        const t = this.next();
        const right = this.parseMul();
        left = this.bin(t.value, left, right, t);
      }
      return left;
    }
    parseMul() {
      let left = this.parseUnary();
      while (this.peekOp("*") || this.peekOp("/") || this.peekOp("\\") || this.peekKw("mod")) {
        const t = this.next();
        const op = t.type === "ident" ? "Mod" : t.value;
        const right = this.parseUnary();
        left = this.bin(op, left, right, t);
      }
      return left;
    }
    parseUnary() {
      if (this.peekOp("-") || this.peekOp("+")) {
        const t = this.next();
        const operand = this.parseUnary();
        return { type: "Unary", op: t.value, operand, line: t.line, col: t.col, start: t.start, end: operand.end };
      }
      return this.parsePow();
    }
    parsePow() {
      let left = this.parseChain({ stopAtSpacedParen: false });
      while (this.peekOp("^")) {
        const t = this.next();
        let right;
        if (this.peekOp("-")) {
          const m = this.next();
          const operand = this.parseChain({ stopAtSpacedParen: false });
          right = { type: "Unary", op: "-", operand, line: m.line, col: m.col, start: m.start, end: operand.end };
        } else {
          right = this.parseChain({ stopAtSpacedParen: false });
        }
        left = this.bin("^", left, right, t);
      }
      return left;
    }

    parsePrimary() {
      const t = this.peek();
      if (t.type === "num") {
        this.next();
        return this.node("Num", t, { value: t.value });
      }
      if (t.type === "str") {
        this.next();
        return this.node("Str", t, { value: t.value });
      }
      if (this.isOp(t, "(")) {
        this.next();
        const expr = this.parseExpr();
        this.expectOp(")", "かっこ「(」が閉じていません。", "「(」と「)」の数がそろっているか確かめよう。");
        return this.node("Paren", t, { expr });
      }
      if (this.isOp(t, "#")) {
        // ファイル番号（EOF(#1) など）
        this.next();
        return this.parsePrimary();
      }
      if (t.type === "ident") {
        const w = t.value.toLowerCase();
        if (w === "not") {
          // "a" & Not (x > 1) のように、式の途中に出てくる Not
          this.next();
          const operand = this.parseCompare();
          return { type: "Unary", op: "Not", operand, line: t.line, col: t.col, start: t.start, end: operand.end };
        }
        if (NOT_EXPR.has(w)) {
          this.fail("コンパイル エラー: 構文エラー", t,
            `値や式が必要なところに「${t.value}」があります。`,
            "計算や値（数・\"文字\"・変数）を書き忘れていないか確かめよう。");
        }
        this.next();
        return this.node("Ident", t, { name: t.value });
      }
      if (t.type === "nl" || t.type === "eof") {
        this.fail("コンパイル エラー: 構文エラー", t, "式（値や計算）が途中で終わっています。", "「=」や「+」のあとに、値を書き忘れていないか確かめよう。");
      }
      this.fail("コンパイル エラー: 構文エラー", t,
        `「${t.value}」をここに書くことはできません。`,
        "記号の書き方や、かっこの数を確かめよう。");
    }

    // a / a(1) / a.b / Range("A1").Value / Me.Label1.Caption など
    parseChain(opts) {
      const startTok = this.peek();
      let expr = this.parsePrimary();
      for (;;) {
        const t = this.peek();
        if (this.isOp(t, "(") && !(opts && opts.stopAtSpacedParen && t.spaceBefore)) {
          if (expr.type === "Num" || expr.type === "Str") break;
          this.next();
          const args = [];
          if (!this.peekOp(")")) {
            for (;;) {
              if (this.peekOp(",")) {
                args.push(null);
                this.next();
                continue;
              }
              args.push(this.parseExpr());
              if (this.peekOp(",")) {
                this.next();
                if (this.peekOp(")")) args.push(null);
                continue;
              }
              break;
            }
          }
          this.expectOp(")", "かっこ「(」が閉じていません。", "「(」と「)」の数がそろっているか確かめよう。");
          expr = this.node("Call", startTok, { callee: expr, args });
          continue;
        }
        if (this.isOp(t, ".")) {
          this.next();
          const nameTok = this.peek();
          if (nameTok.type !== "ident") {
            this.fail("コンパイル エラー: 構文エラー", nameTok, "「.」のあとにはプロパティの名前が必要です。", "例: Label1.Caption");
          }
          this.next();
          expr = this.node("Member", startTok, { object: expr, name: nameTok.value });
          continue;
        }
        return expr;
      }
      return expr;
    }
  }

  // 戻り値: { ast, zenkaku }
  function parseProgram(src) {
    let lexed;
    try {
      lexed = Lexer.tokenize(src);
    } catch (e) {
      e.zenkaku = Lexer.tokenizeLoose(src).zenkaku;
      throw e;
    }
    const p = new Parser(lexed.tokens, src);
    try {
      const ast = p.parseModule();
      ast.src = src;
      return { ast, zenkaku: lexed.zenkaku };
    } catch (e) {
      e.zenkaku = lexed.zenkaku;
      throw e;
    }
  }

  function nodeText(src, node) {
    return node ? src.slice(node.start, node.end) : "";
  }

  const api = { parseProgram, nodeText, VbaSyntaxError, TYPE_NAMES };
  if (typeof module !== "undefined" && module.exports) module.exports = api;
  else root.VbaAst = api;
})(typeof window !== "undefined" ? window : globalThis);
