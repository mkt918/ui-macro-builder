/* ===== 流れ図（フローチャート） =====
 * プログラム（AST）から、教科書と同じ書き方の流れ図を SVG で描く。手続きごとに横に並べる。
 *
 *   端子（はじめ／おわり）・処理（Tanka × Suu → Kin）・準備（Initialize の中の代入）・
 *   入出力（Suuを入力／〜を表示）・判断（ひし形と YES / NO）・ループ端（ループ1 〜の間）・定義済み処理（Call）
 *
 * 描き方の決まりは docs/textbook-userform-plan.md の「流れ図の描き方」の表に合わせている。
 *   new FlowChart(container, { onHover(line), onClick(line) })
 *   .render(ast)                  … 描き直す
 *   .highlight(line, proc, branch) … 実行中の記号を光らせる（branch: 判断の結果 true/false）
 */
(function () {
  "use strict";

  const GAP = 20; // 記号と記号のすき間（矢印の長さ）
  const COL_GAP = 34; // 右に分かれる枝との間
  const FONT = 13;
  const LINE_H = 17;

  function esc(s) {
    return String(s).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");
  }
  // 文字の幅の目安（全角は1文字、半角は0.55文字）
  function textUnits(s) {
    let n = 0;
    for (const ch of s) n += /[\x20-\x7e]/.test(ch) ? 0.58 : 1;
    return n;
  }
  // 長い文字は折り返す（1行 maxUnits 文字ぶんまで）
  function wrap(s, maxUnits) {
    const lines = [];
    String(s)
      .split("\n")
      .forEach((part) => {
        let cur = "";
        let u = 0;
        for (const ch of part) {
          const w = /[\x20-\x7e]/.test(ch) ? 0.58 : 1;
          if (u + w > maxUnits && cur) {
            // 英単語の途中で切れないよう、近くに空白があればそこで折り返す
            const sp = cur.lastIndexOf(" ");
            if (/[A-Za-z0-9_]/.test(ch) && sp > cur.length / 2) {
              lines.push(cur.slice(0, sp));
              cur = cur.slice(sp + 1);
              u = textUnits(cur);
            } else {
              lines.push(cur);
              cur = "";
              u = 0;
            }
          }
          cur += ch;
          u += w;
        }
        lines.push(cur);
      });
    return lines;
  }

  // ---------- AST → 流れ図の部品 ----------
  // 部品: { k: "seq", items } / { k: "node", shape, text, line } / { k: "if", ... } / { k: "loop", ... }
  class Builder {
    constructor(src) {
      this.src = src;
      this.loopNo = 0;
      this.inInit = false;
    }
    text(n) {
      return n ? this.src.slice(n.start, n.end) : "";
    }
    // 式を教科書の書き方に（* → ×、/ → ÷、>= → ≧ …）
    pretty(n) {
      let t = this.text(n);
      t = t
        .replace(/\s*\*\s*/g, " × ")
        .replace(/\s*\/\s*/g, " ÷ ")
        .replace(/\s*\\\s*/g, " ¥ ")
        .replace(/>=/g, "≧")
        .replace(/<=/g, "≦")
        .replace(/<>/g, "≠")
        .replace(/\bMe\./g, "");
      return t.replace(/\s+/g, " ").trim();
    }
    // 表示する値から、Label1.Caption & … や 改行（Chr(13) & Chr(10)・vbCrLf）・Format を取り除く
    shownValue(valueNode, targetText) {
      let t = this.text(valueNode).replace(/\bMe\./g, "");
      const tgt = targetText.replace(/\bMe\./g, "").replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
      t = t.replace(new RegExp("^" + tgt + "\\s*&\\s*"), "");
      t = t.replace(/\s*&\s*(Chr\(13\)\s*&\s*Chr\(10\)|Chr\(10\)|Chr\(13\)|vbCrLf|vbNewLine|vbLf)\s*$/i, "");
      t = t.replace(/^Format\((.+?),\s*"[^"]*"\)$/i, "$1");
      t = t.replace(/^Val\((.+)\)$/i, "$1");
      return t.trim();
    }

    seq(stmts) {
      const items = [];
      for (const s of stmts || []) {
        const it = this.stmt(s);
        if (it) items.push(it);
      }
      return { k: "seq", items };
    }

    node(shape, text, s) {
      return { k: "node", shape, text, line: s.line };
    }

    stmt(s) {
      switch (s.type) {
        case "Dim":
        case "Const":
        case "ReDim":
        case "Open":
        case "Close":
          return null; // 教科書の流れ図にも出てこない
        case "Assign":
          return this.assign(s);
        case "CallStmt":
          return this.call(s);
        case "If":
          return this.ifStmt(s);
        case "Select":
          return this.selectStmt(s);
        case "For":
          return this.forStmt(s);
        case "Do":
          return this.doStmt(s);
        case "Exit":
          if (s.what === "Sub" || s.what === "Function") return { k: "node", shape: "terminal", text: "おわり", line: s.line, exit: true };
          return this.node("process", s.what === "For" ? "くり返しをぬける（Exit For）" : "くり返しをぬける（Exit Do）", s);
        case "End":
          return { k: "node", shape: "terminal", text: "おわり", line: s.line, exit: true };
        case "InputFile":
          return this.node("io", `ファイルのデータを ${s.targets.map((t) => this.text(t)).join(", ")} に読みこむ`, s);
        case "LineInput":
          return this.node("io", `ファイルの1行を ${this.text(s.target)} に読みこむ`, s);
        case "PrintFile":
        case "WriteFile":
          return this.node("io", `${s.items.map((i) => this.text(i.expr)).join(", ")} をファイルに書きこむ`, s);
        default:
          return this.node("process", this.text(s), s);
      }
    }

    assign(s) {
      const target = this.text(s.target).replace(/\bMe\./g, "");
      const value = this.text(s.value);
      const ctl = target.match(/^(\w+)\.(Caption|Text|Value)$/i);
      if (ctl && /Caption/i.test(ctl[2])) {
        return this.node("io", `${this.shownValue(s.value, target) || '""'}を${ctl[1]}に表示`, s);
      }
      if (/InputBox\s*\(/i.test(value)) {
        return this.node("io", `${target}を入力`, s);
      }
      const tb = value.replace(/\bMe\./g, "").match(/^(?:Val\()?(\w+)\.(Text|Value)\)?$/i);
      if (tb) return this.node(this.inInit ? "prep" : "process", `${tb[1]} → ${target}`, s);
      const rng = target.match(/^(?:Range\("([A-Z]+\d+)"\)|Cells\((.+)\))\.Value$/i);
      const t = rng ? (rng[1] ? `セル${rng[1]}` : target.replace(/\.Value$/i, "")) : target;
      return this.node(this.inInit ? "prep" : "process", `${this.pretty(s.value)} → ${t}`, s);
    }

    call(s) {
      const callee = s.callee;
      const head = (callee.type === "Call" ? callee.callee : callee);
      const name = head.type === "Ident" ? head.name : head.type === "Member" ? head.name : "";
      const argNodes = callee.type === "Call" && s.args === null ? callee.args : s.args || [];
      const lower = name.toLowerCase();
      if (lower === "msgbox") {
        const a = argNodes[0];
        let t = a ? this.text(a) : '""';
        if (a && a.type === "Paren") t = this.text(a.expr);
        return this.node("io", `${t}を表示`, s);
      }
      if (lower === "unload") return this.node("process", "フォームを閉じる", s);
      // 自分で作った Sub の呼び出しは「定義済み処理」の記号
      if (head.type === "Ident" && this.isUserProc && this.isUserProc(name)) return this.node("predef", name, s);
      return this.node("process", this.text(s), s);
    }

    cond(n) {
      const t = this.text(n).replace(/\bMe\./g, "");
      if (/^Not\s+EOF\s*\(\s*#?\d+\s*\)$/i.test(t)) return "データがある";
      if (/^EOF\s*\(\s*#?\d+\s*\)$/i.test(t)) return "データの終わり";
      return this.pretty(n);
    }

    // 枝の最後が「おわり」（End・Exit Sub）か
    exits(seq) {
      const last = seq.items[seq.items.length - 1];
      if (!last) return false;
      if (last.k === "node") return !!last.exit;
      if (last.k === "if") return last.yes && last.no && this.exits(last.yes) && this.exits(last.no);
      return false;
    }

    ifStmt(s) {
      // ElseIf は「いいえ」の側に次の判断をつなげる（教科書の例3と同じ形）
      const build = (k) => {
        if (k >= s.branches.length) return s.elseBody ? this.seq(s.elseBody) : { k: "seq", items: [] };
        const br = s.branches[k];
        const yes = this.seq(br.body);
        const no = build(k + 1);
        const noSeq = no.k === "seq" ? no : { k: "seq", items: [no] };
        return { k: "if", cond: this.cond(br.cond), line: br.line, yes, no: noSeq };
      };
      const r = build(0);
      return r.k === "if" ? r : null;
    }

    selectStmt(s) {
      const subject = this.pretty(s.expr);
      const testText = (t) => {
        if (t.kind === "value") return `${subject} = ${this.pretty(t.expr)}`;
        if (t.kind === "range") return `${this.pretty(t.from)} ≦ ${subject} ≦ ${this.pretty(t.to)}`;
        return `${subject} ${t.op.replace(">=", "≧").replace("<=", "≦").replace("<>", "≠")} ${this.pretty(t.expr)}`;
      };
      const build = (k) => {
        if (k >= s.cases.length) return s.elseBody ? this.seq(s.elseBody) : { k: "seq", items: [] };
        const c = s.cases[k];
        const no = build(k + 1);
        return { k: "if", cond: c.tests.map(testText).join(" または "), line: c.line, yes: this.seq(c.body), no: no.k === "seq" ? no : { k: "seq", items: [no] } };
      };
      const r = build(0);
      return r.k === "if" ? r : null;
    }

    forStmt(s) {
      const no = ++this.loopNo;
      const step = s.step ? this.pretty(s.step) : "1";
      const head = `ループ${no}\n${this.text(s.var)} を ${this.pretty(s.from)} から ${this.pretty(s.to)} まで${step !== "1" ? `（${step} ずつ）` : ""}`;
      return { k: "loop", start: { text: head, line: s.line }, end: { text: `ループ${no}`, line: s.endLine }, body: this.seq(s.body) };
    }

    doStmt(s) {
      const no = ++this.loopNo;
      const condText = (c) => {
        const t = this.cond(c.expr);
        if (t === "データがある") return c.kind === "While" ? "データがある間" : "データがなくなるまで";
        return c.kind === "While" ? `${t} の間` : `${t} になるまで`;
      };
      const startText = s.pre ? `ループ${no}\n${condText(s.pre)}` : `ループ${no}`;
      const endText = s.post ? `ループ${no}\n${condText(s.post)}` : `ループ${no}`;
      return { k: "loop", start: { text: startText, line: s.line }, end: { text: endText, line: s.endLine }, body: this.seq(s.body) };
    }

    proc(p, procNames) {
      this.loopNo = 0;
      this.inInit = /^UserForm_Initialize$/i.test(p.name);
      this.isUserProc = (n) => procNames.has(n.toLowerCase());
      const body = this.seq(p.body);
      const items = [{ k: "node", shape: "terminal", text: "はじめ", line: p.line }].concat(body.items);
      if (!this.exits(body)) items.push({ k: "node", shape: "terminal", text: "おわり", line: p.endLine });
      return { name: p.name, seq: { k: "seq", items } };
    }
  }

  // ---------- 配置（大きさを測ってから座標を決める） ----------
  // 測った結果: { w, h, cx（中心線の x）, exits（下へ流れが続かない）}
  function measureNode(n) {
    const maxUnits = n.shape === "decision" ? 12 : 15;
    const lines = wrap(n.text, maxUnits);
    const units = Math.max(...lines.map(textUnits));
    let w = Math.max(n.shape === "terminal" ? 110 : 130, Math.ceil(units * FONT) + (n.shape === "io" ? 50 : n.shape === "decision" ? 70 : 34));
    let h = Math.max(n.shape === "decision" ? 56 : 36, lines.length * LINE_H + (n.shape === "decision" ? 30 : 16));
    if (n.shape === "loopStart" || n.shape === "loopEnd") h += 4;
    n.lines = lines;
    return { w, h, cx: w / 2, exits: !!n.exit };
  }

  function measure(item) {
    if (item.k === "node") {
      item.m = measureNode(item);
      return item.m;
    }
    if (item.k === "loop") {
      const sNode = { k: "node", shape: "loopStart", text: item.start.text, line: item.start.line };
      const eNode = { k: "node", shape: "loopEnd", text: item.end.text, line: item.end.line };
      item.seqWrapped = { k: "seq", items: [sNode].concat(item.body.items, [eNode]) };
      item.m = measure(item.seqWrapped);
      item.m.exits = false;
      return item.m;
    }
    if (item.k === "seq") {
      let left = 0;
      let right = 0;
      let h = 0;
      let exits = false;
      item.items.forEach((it, i) => {
        const m = measure(it);
        left = Math.max(left, m.cx);
        right = Math.max(right, m.w - m.cx);
        h += m.h + (i > 0 ? GAP : 0);
        if (m.exits) exits = true;
      });
      const last = item.items[item.items.length - 1];
      item.m = { w: left + right, h, cx: left, exits: !!(last && last.m.exits) };
      return item.m;
    }
    if (item.k === "if") {
      const d = { k: "node", shape: "decision", text: item.cond, line: item.line };
      item.dnode = d;
      const dm = measure(d);
      const Y = measure(item.yes);
      const N = measure(item.no);
      const yesExits = item.yes.items.length > 0 && Y.exits;
      const noExits = item.no.items.length > 0 && N.exits;
      // 「はい」の側が End・Exit Sub で終わり、「いいえ」が空なら、「はい」を横に出して流れを下に続ける（教科書 p.107）
      item.sideYes = yesExits && item.no.items.length === 0;
      const down = item.sideYes ? item.no : item.yes;
      const side = item.sideYes ? item.yes : item.no;
      const Dm = down.m;
      const Sm = side.m;
      const leftExt = Math.max(dm.cx, Dm.cx);
      const mainRight = Math.max(dm.w - dm.cx, Dm.w - Dm.cx);
      const sideOffset = leftExt + mainRight + COL_GAP; // 横の枝の左端
      const sideW = side.items.length ? Sm.w : 20;
      item.layout = { leftExt, mainRight, sideOffset, sideW };
      const branchH = Math.max(down.items.length ? Dm.h : 0, side.items.length ? Sm.h : 0);
      const downEx = down.items.length > 0 && Dm.exits;
      const sideEx = side.items.length > 0 && Sm.exits;
      item.mergeDown = !downEx;
      item.mergeSide = !sideEx;
      const merges = item.mergeDown || item.mergeSide;
      const h = dm.h + GAP + branchH + (merges ? GAP : 0);
      item.m = { w: sideOffset + sideW, h, cx: leftExt, exits: !merges || (yesExits && noExits) };
      return item.m;
    }
    return { w: 0, h: 0, cx: 0, exits: false };
  }

  // ---------- SVG を書く ----------
  class Painter {
    constructor(proc) {
      this.parts = [];
      this.proc = proc;
    }
    line(x1, y1, x2, y2, arrow) {
      this.parts.push(`<line class="fc-edge" x1="${x1}" y1="${y1}" x2="${x2}" y2="${y2}" ${arrow ? 'marker-end="url(#fc-arrow)"' : ""}/>`);
    }
    label(x, y, text, line, dir) {
      this.parts.push(`<text class="fc-yn" data-line="${line}" data-dir="${dir}" x="${x}" y="${y}">${text}</text>`);
    }
    shape(n, x, y) {
      const { w, h } = n.m;
      const attrs = `class="fc-node fc-${n.shape}" data-line="${n.line || ""}" data-proc="${esc(this.proc)}"`;
      let path;
      if (n.shape === "terminal") path = `<rect x="${x}" y="${y}" width="${w}" height="${h}" rx="${h / 2}" ry="${h / 2}"/>`;
      else if (n.shape === "io") path = `<polygon points="${x + 14},${y} ${x + w},${y} ${x + w - 14},${y + h} ${x},${y + h}"/>`;
      else if (n.shape === "decision") path = `<polygon points="${x + w / 2},${y} ${x + w},${y + h / 2} ${x + w / 2},${y + h} ${x},${y + h / 2}"/>`;
      else if (n.shape === "prep") path = `<polygon points="${x + 12},${y} ${x + w - 12},${y} ${x + w},${y + h / 2} ${x + w - 12},${y + h} ${x + 12},${y + h} ${x},${y + h / 2}"/>`;
      else if (n.shape === "loopStart") path = `<polygon points="${x + 12},${y} ${x + w - 12},${y} ${x + w},${y + 12} ${x + w},${y + h} ${x},${y + h} ${x},${y + 12}"/>`;
      else if (n.shape === "loopEnd") path = `<polygon points="${x},${y} ${x + w},${y} ${x + w},${y + h - 12} ${x + w - 12},${y + h} ${x + 12},${y + h} ${x},${y + h - 12}"/>`;
      else if (n.shape === "predef") path = `<rect x="${x}" y="${y}" width="${w}" height="${h}"/><line x1="${x + 9}" y1="${y}" x2="${x + 9}" y2="${y + h}"/><line x1="${x + w - 9}" y1="${y}" x2="${x + w - 9}" y2="${y + h}"/>`;
      else path = `<rect x="${x}" y="${y}" width="${w}" height="${h}"/>`;
      const lines = n.lines || [n.text];
      const top = y + h / 2 - ((lines.length - 1) * LINE_H) / 2 + 4.5;
      const text = lines.map((l, i) => `<tspan x="${x + w / 2}" y="${top + i * LINE_H}">${esc(l)}</tspan>`).join("");
      const tip = `<title>${esc(n.line ? `${n.line}行目` : "")}</title>`;
      this.parts.push(`<g ${attrs}>${tip}${path}<text class="fc-text">${text}</text></g>`);
    }

    // item を (x: 中心線の左端からの基準, y) に描く。戻り値: 下の出口の y（流れが続かなければ null）
    draw(item, x0, y) {
      const m = item.m;
      if (item.k === "node") {
        this.shape(item, x0 + m.cx - m.w / 2, y);
        return item.exit ? null : y + m.h;
      }
      if (item.k === "loop") return this.draw(item.seqWrapped, x0, y);
      if (item.k === "seq") {
        let cy = y;
        let open = true;
        item.items.forEach((it, i) => {
          if (i > 0) {
            if (open) this.line(x0 + m.cx, cy, x0 + m.cx, cy + GAP, true);
            cy += GAP;
          }
          const out = this.draw(it, x0 + m.cx - it.m.cx, cy);
          cy += it.m.h;
          open = out !== null;
        });
        return open ? cy : null;
      }
      if (item.k === "if") {
        const d = item.dnode;
        const L = item.layout;
        const cx = x0 + L.leftExt;
        this.shape(d, cx - d.m.w / 2, y);
        const dBottom = y + d.m.h;
        const dRightX = cx + d.m.w / 2;
        const dMidY = y + d.m.h / 2;
        const down = item.sideYes ? item.no : item.yes;
        const side = item.sideYes ? item.yes : item.no;
        const downLabel = item.sideYes ? "NO" : "YES";
        const sideLabel = item.sideYes ? "YES" : "NO";
        this.label(cx + 6, dBottom + 13, downLabel, d.line, downLabel === "YES" ? "yes" : "no");
        this.label(dRightX + 4, dMidY - 6, sideLabel, d.line, sideLabel === "YES" ? "yes" : "no");
        const by = dBottom + GAP;
        const branchH = item.m.h - d.m.h - GAP - (item.mergeDown || item.mergeSide ? GAP : 0);
        const mergeY = by + branchH + GAP;
        // 下の枝
        let downOut;
        if (down.items.length) {
          this.line(cx, dBottom, cx, by, true);
          downOut = this.draw(down, cx - down.m.cx, by);
        } else {
          downOut = dBottom;
        }
        // 横の枝
        const sx = x0 + L.sideOffset + (side.items.length ? side.m.cx : 10);
        let sideOut;
        if (side.items.length) {
          this.line(dRightX, dMidY, sx, dMidY, false);
          this.line(sx, dMidY, sx, by, true);
          sideOut = this.draw(side, sx - side.m.cx, by);
        } else {
          this.line(dRightX, dMidY, sx, dMidY, false);
          sideOut = dMidY;
        }
        const merges = item.mergeDown || item.mergeSide;
        if (!merges) return null;
        if (item.mergeDown && downOut !== null) this.line(cx, downOut, cx, mergeY, false);
        if (item.mergeSide && sideOut !== null) {
          this.line(sx, sideOut, sx, mergeY, false);
          this.line(sx, mergeY, cx, mergeY, true);
        }
        return mergeY;
      }
      return y;
    }
  }

  class FlowChart {
    constructor(container, opts) {
      this.el = container;
      this.opts = opts || {};
      this.current = null;
      this.el.addEventListener("mouseover", (e) => {
        const g = e.target.closest(".fc-node");
        if (g && this.opts.onHover) this.opts.onHover(Number(g.dataset.line) || null);
      });
      this.el.addEventListener("mouseleave", () => {
        if (this.opts.onHover) this.opts.onHover(null);
      });
      this.el.addEventListener("click", (e) => {
        const g = e.target.closest(".fc-node");
        if (g && this.opts.onClick) this.opts.onClick(Number(g.dataset.line) || null);
      });
    }

    render(ast, message) {
      if (!ast) {
        this.el.innerHTML = `<div class="fc-empty">${esc(message || "プログラムにまちがいがあるので、流れ図を描けません")}</div>`;
        return;
      }
      if (!ast.procs.length) {
        this.el.innerHTML = '<div class="fc-empty">Sub ～ End Sub を書くと、ここに流れ図が出ます</div>';
        return;
      }
      const b = new Builder(ast.src || "");
      const procNames = new Set(ast.procs.map((p) => p.name.toLowerCase()));
      const charts = ast.procs.map((p) => b.proc(p, procNames));
      let html = "";
      charts.forEach((c) => {
        const m = measure(c.seq);
        const pad = 14;
        const W = Math.ceil(m.w + pad * 2);
        const H = Math.ceil(m.h + pad * 2);
        const p = new Painter(c.name);
        p.draw(c.seq, pad, pad);
        html += `<div class="fc-proc"><div class="fc-title">${esc(c.name)}</div>
          <svg class="fc-svg" width="${W}" height="${H}" viewBox="0 0 ${W} ${H}" xmlns="http://www.w3.org/2000/svg">
            <defs><marker id="fc-arrow" viewBox="0 0 10 10" refX="9" refY="5" markerWidth="7" markerHeight="7" orient="auto-start-reverse"><path d="M0,0 L10,5 L0,10 z" class="fc-arrowhead"/></marker></defs>
            ${p.parts.join("")}
          </svg></div>`;
      });
      this.el.innerHTML = `<div class="fc-row">${html}</div><div class="fc-legend">記号の上にマウスを乗せると、その命令の行が光ります</div>`;
      this.highlight(this.current && this.current.line, this.current && this.current.proc, this.current && this.current.branch);
    }

    // 実行中の記号を光らせる
    highlight(line, proc, branch) {
      this.current = { line, proc, branch };
      this.el.querySelectorAll(".fc-node.active, .fc-yn.active").forEach((n) => n.classList.remove("active"));
      if (!line) return;
      const sel = `.fc-node[data-line="${line}"]` + (proc ? `[data-proc="${CSS.escape(proc)}"]` : "");
      const nodes = this.el.querySelectorAll(sel);
      nodes.forEach((n) => n.classList.add("active"));
      if (branch === true || branch === false) {
        this.el.querySelectorAll(`.fc-yn[data-line="${line}"][data-dir="${branch ? "yes" : "no"}"]`).forEach((n) => n.classList.add("active"));
      }
      const first = nodes[0];
      if (first && this.el.offsetParent) {
        const r = first.getBoundingClientRect();
        const box = this.el.getBoundingClientRect();
        if (r.top < box.top || r.bottom > box.bottom) first.scrollIntoView({ block: "nearest", inline: "nearest" });
      }
    }
  }

  window.FlowChart = FlowChart;
})();
