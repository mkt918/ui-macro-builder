/* ===== ユーザーフォーム（デザイナーと実行時の表示） =====
 * 右ペインの「🪟 ユーザーフォーム」タブ。
 *  - デザイン中：部品（ラベル・テキストボックス・ボタン）を置く・動かす・大きさを変える・プロパティを変える
 *  - 実行中　　：実行エンジンが記録したフォームの状態を表示し、テキストボックスへの入力とボタンを受け付ける
 *
 * フォームのデータ（form）:
 *   { name, caption, width, height, controls: [{ name, type, left, top, width, height, props: {...} }] }
 */
(function () {
  "use strict";

  const GRID = 6;
  const DEFAULT_SIZE = { Label: [120, 24], TextBox: [130, 26], CommandButton: [110, 32], ListBox: [150, 110] };
  const TYPE_JA = { Label: "ラベル", TextBox: "テキストボックス", CommandButton: "コマンドボタン", ListBox: "リストボックス" };
  const TYPE_ICON = { Label: "🏷", TextBox: "⌨", CommandButton: "🔘", ListBox: "📋" };
  const HELP = {
    Label: "文字を表示する部品です。プログラムで「Label1.Caption = \"こんにちは\"」とすると、表示が変わります。",
    TextBox: "文字を入力してもらう部品です。入力された文字は「TextBox1.Text」で読めます（計算に使うときは Val(TextBox1.Text) で数にしよう）。",
    CommandButton: "押すと命令が動くボタンです。ダブルクリックすると、押したときの命令（CommandButton1_Click）を書く場所ができます。",
    ListBox: "いくつもの文字を一覧で表示する部品です。「ListBox1.AddItem \"りんご\"」で追加できます。",
  };
  const RESERVED = new Set(["userform1", "me", "userform", "end", "sub", "dim", "if", "for", "next", "do", "loop", "then", "else", "msgbox", "inputbox", "val"]);

  function clone(o) {
    return JSON.parse(JSON.stringify(o));
  }
  function snap(n) {
    return Math.round(n / GRID) * GRID;
  }
  function vbaColorToHex(n) {
    if (n === undefined || n === null || n === "" || isNaN(Number(n))) return null;
    n = Number(n) & 0xffffff;
    const r = n & 255;
    const g = (n >> 8) & 255;
    const b = (n >> 16) & 255;
    return "#" + [r, g, b].map((x) => x.toString(16).padStart(2, "0")).join("");
  }
  function hexToVbaColor(hex) {
    const m = String(hex).match(/^#?([0-9a-f]{2})([0-9a-f]{2})([0-9a-f]{2})$/i);
    if (!m) return undefined;
    return parseInt(m[1], 16) + parseInt(m[2], 16) * 256 + parseInt(m[3], 16) * 65536;
  }
  function esc(s) {
    return String(s === undefined || s === null ? "" : s).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");
  }

  function makeControl(type, name, left, top) {
    const [w, h] = DEFAULT_SIZE[type] || [100, 24];
    const props = { FontSize: 12 };
    if (type === "TextBox") props.Text = "";
    else if (type === "ListBox") props.List = [];
    else props.Caption = name;
    return { name, type, left, top, width: w, height: h, props };
  }

  // 新しく作ったときのフォーム（教科書の最初の形：ボタンが1つ）
  function defaultForm() {
    return { name: "UserForm1", caption: "UserForm1", width: 360, height: 240, controls: [makeControl("CommandButton", "CommandButton1", 126, 168)] };
  }

  // 実行エンジンに渡す形（位置と大きさもプロパティとして渡す）
  function toRuntimeForm(form) {
    return {
      name: form.name,
      caption: form.caption,
      controls: form.controls.map((c) => ({
        name: c.name,
        type: c.type,
        props: Object.assign({}, c.props, { Left: c.left, Top: c.top, Width: c.width, Height: c.height }),
      })),
    };
  }

  // 部品1つ分の見た目（style 属性）
  function controlStyle(type, p, geo) {
    const s = [`left:${geo.left}px`, `top:${geo.top}px`, `width:${geo.width}px`, `height:${geo.height}px`];
    if (p.FontSize) s.push(`font-size:${p.FontSize}px`);
    if (p.FontBold) s.push("font-weight:bold");
    if (p.FontItalic) s.push("font-style:italic");
    const fc = vbaColorToHex(p.ForeColor);
    if (fc) s.push(`color:${fc}`);
    const bc = vbaColorToHex(p.BackColor);
    if (bc) s.push(`background:${bc}`);
    const align = { 1: "left", 2: "center", 3: "right" }[p.TextAlign];
    if (align) {
      s.push(`text-align:${align}`);
      if (type === "Label") s.push(`justify-content:${align === "left" ? "flex-start" : align === "right" ? "flex-end" : "center"}`);
    }
    if (type === "Label") {
      if (Number(p.BorderStyle) === 1) s.push("border:1px solid #000");
      const fx = Number(p.SpecialEffect);
      if (fx === 1) s.push("border:2px outset #f0f0f0");
      else if (fx === 2) s.push("border:2px inset #f0f0f0; background:" + (bc || "#ffffff"));
      else if (fx === 3) s.push("border:2px groove #f0f0f0");
      else if (fx === 6) s.push("border:2px ridge #f0f0f0");
    }
    if (p.Visible === false) s.push("visibility:hidden");
    return s.join(";");
  }

  class UserFormPanel {
    // opts: { stage, props, onChange(form, info), onEditHandler(name), onRunClick(name, texts), onRunClose() }
    constructor(opts) {
      this.opts = opts;
      this.stage = opts.stage;
      this.propsEl = opts.props;
      this.form = defaultForm();
      this.selected = null;
      this.mode = "design";
      this.showNames = false;
      this.history = [];
      this.run = null; // { controls, form, interactive, info }
      this.runTexts = {};
      this.bindStage();
      this.render();
    }

    // ---------- データ ----------
    setForm(form) {
      this.form = form ? clone(form) : defaultForm();
      this.selected = null;
      this.history = [];
      this.render();
    }
    getForm() {
      return clone(this.form);
    }
    runtimeForm() {
      return toRuntimeForm(this.form);
    }
    controlNames() {
      return this.form.controls.map((c) => ({ name: c.name, type: c.type }));
    }
    find(name) {
      return this.form.controls.find((c) => c.name === name) || null;
    }
    pushHistory() {
      this.history.push(JSON.stringify(this.form));
      if (this.history.length > 60) this.history.shift();
    }
    undo() {
      if (this.mode !== "design" || !this.history.length) return;
      const before = this.controlNames().map((c) => c.name).join(",");
      this.form = JSON.parse(this.history.pop());
      if (this.selected && !this.find(this.selected)) this.selected = null;
      this.render();
      const after = this.controlNames().map((c) => c.name).join(",");
      this.changed({ structure: before !== after });
    }
    changed(info) {
      if (this.opts.onChange) this.opts.onChange(this.getForm(), info || {});
    }

    nextName(type) {
      let n = 1;
      while (this.form.controls.some((c) => c.name.toLowerCase() === (type + n).toLowerCase())) n++;
      return type + n;
    }

    // 空いている場所を探して部品を置く（ドラッグしなくても置けるように）
    addControl(type) {
      if (this.mode !== "design") return;
      this.pushHistory();
      const name = this.nextName(type);
      const [w, h] = DEFAULT_SIZE[type];
      const overlaps = (x, y) => this.form.controls.some((c) => x < c.left + c.width + GRID && x + w + GRID > c.left && y < c.top + c.height + GRID && y + h + GRID > c.top);
      let pos = null;
      for (let y = 12; y + h <= this.form.height - 6 && !pos; y += 36) {
        for (let x = 12; x + w <= this.form.width - 6; x += 18) {
          if (!overlaps(x, y)) {
            pos = [x, y];
            break;
          }
        }
      }
      if (!pos) pos = [12 + (this.form.controls.length % 5) * 12, 12 + (this.form.controls.length % 5) * 12];
      this.form.controls.push(makeControl(type, name, snap(pos[0]), snap(pos[1])));
      this.selected = name;
      this.render();
      this.changed({ structure: true, added: name });
    }

    deleteSelected() {
      if (this.mode !== "design" || !this.selected) return;
      this.pushHistory();
      const name = this.selected;
      this.form.controls = this.form.controls.filter((c) => c.name !== name);
      this.selected = null;
      this.render();
      this.changed({ structure: true, deleted: name });
    }

    // ---------- 表示モード ----------
    setDesign() {
      this.mode = "design";
      this.run = null;
      this.render();
    }
    // controls: 実行中の部品の状態（実行エンジンの state.controls）、formState: { visible, caption }
    setRun(controls, formState, interactive, info) {
      this.mode = "run";
      this.run = { controls: controls || {}, form: formState || {}, interactive: !!interactive, info: info || "" };
      this.runTexts = {};
      Object.keys(this.run.controls).forEach((n) => {
        const c = this.run.controls[n];
        if (c.type === "TextBox") this.runTexts[n] = c.Text || "";
      });
      this.render();
    }
    setShowNames(on) {
      this.showNames = !!on;
      this.render();
    }

    // ---------- 描画 ----------
    render() {
      this.renderStage();
      this.renderProps();
    }

    renderStage() {
      const f = this.form;
      if (this.mode === "run") {
        const r = this.run;
        if (!r.form.visible) {
          this.stage.innerHTML = `<div class="uf-closed">🪟 フォームは閉じています<br><small>${esc(r.info || "▶ 実行 を押すと開きます")}</small></div>`;
          return;
        }
      }
      const caption = this.mode === "run" ? this.run.form.caption || f.caption : f.caption;
      let html = `<div class="uf-window ${this.mode}" style="width:${f.width}px">`;
      html += `<div class="uf-titlebar"><span class="uf-title">${esc(caption)}</span><button class="uf-x" title="フォームを閉じる" ${this.mode === "run" && this.run.interactive ? "" : "tabindex=-1"}>✕</button></div>`;
      html += `<div class="uf-client ${this.mode === "design" ? "grid" : ""}" style="height:${f.height}px">`;
      const list = this.mode === "run" ? this.runControlsList() : f.controls.map((c) => ({ c, p: c.props, geo: c }));
      list.forEach(({ c, p, geo }) => {
        html += this.controlHtml(c, p, geo);
      });
      if (this.mode === "design") html += '<div class="uf-resize" title="ドラッグでフォームの大きさを変える"></div>';
      html += "</div></div>";
      if (this.mode === "run" && this.run.info) html = `<div class="uf-run-info">${esc(this.run.info)}</div>` + html;
      this.stage.innerHTML = html;
      if (this.mode === "run") this.bindRunControls();
    }

    // 実行中は、実行エンジンの状態（位置・大きさもプロパティ）から描く
    runControlsList() {
      const out = [];
      const rc = this.run.controls;
      this.form.controls.forEach((d) => {
        const st = rc[d.name];
        if (!st) return;
        const geo = {
          left: st.Left !== undefined ? st.Left : d.left,
          top: st.Top !== undefined ? st.Top : d.top,
          width: st.Width !== undefined ? st.Width : d.width,
          height: st.Height !== undefined ? st.Height : d.height,
        };
        out.push({ c: d, p: st, geo });
      });
      return out;
    }

    controlHtml(c, p, geo) {
      const sel = this.mode === "design" && this.selected === c.name ? " selected" : "";
      const style = controlStyle(c.type, p, geo);
      const callout = this.showNames ? `<span class="uf-callout">${esc(c.name)}</span>` : "";
      const handle = sel ? '<span class="uf-handle" title="ドラッグで大きさを変える"></span>' : "";
      const attrs = `data-name="${esc(c.name)}" style="${style}"`;
      const run = this.mode === "run";
      const can = run && this.run.interactive && p.Enabled !== false;
      if (c.type === "Label") {
        return `<div class="uf-ctl uf-label${sel}" ${attrs}><span class="uf-text">${esc(p.Caption)}</span>${callout}${handle}</div>`;
      }
      if (c.type === "TextBox") {
        if (run) {
          return `<div class="uf-ctl uf-textbox run" ${attrs}><input type="text" data-tb="${esc(c.name)}" value="${esc(this.runTexts[c.name] !== undefined ? this.runTexts[c.name] : p.Text)}" ${can ? "" : "disabled"} aria-label="${esc(c.name)}">${callout}</div>`;
        }
        return `<div class="uf-ctl uf-textbox${sel}" ${attrs}><span class="uf-text">${esc(p.Text)}</span>${callout}${handle}</div>`;
      }
      if (c.type === "CommandButton") {
        if (run) {
          return `<button class="uf-ctl uf-button run" ${attrs} data-btn="${esc(c.name)}" ${can ? "" : "disabled"}>${esc(p.Caption)}${callout}</button>`;
        }
        return `<div class="uf-ctl uf-button${sel}" ${attrs}><span class="uf-text">${esc(p.Caption)}</span>${callout}${handle}</div>`;
      }
      if (c.type === "ListBox") {
        const items = (p.List || []).map((x) => `<div class="uf-li">${esc(x)}</div>`).join("");
        return `<div class="uf-ctl uf-listbox${sel}" ${attrs}>${items}${callout}${handle}</div>`;
      }
      return "";
    }

    bindRunControls() {
      this.stage.querySelectorAll("input[data-tb]").forEach((inp) => {
        inp.addEventListener("input", () => {
          this.runTexts[inp.dataset.tb] = inp.value;
        });
        inp.addEventListener("keydown", (e) => {
          if (e.key === "Enter" && !e.isComposing) e.preventDefault();
        });
      });
      this.stage.querySelectorAll("button[data-btn]").forEach((b) => {
        b.addEventListener("click", () => {
          if (this.opts.onRunClick) this.opts.onRunClick(b.dataset.btn, Object.assign({}, this.runTexts));
        });
      });
      const x = this.stage.querySelector(".uf-x");
      if (x) {
        x.addEventListener("click", () => {
          if (this.run && this.run.interactive && this.opts.onRunClose) this.opts.onRunClose();
        });
      }
    }

    // ---------- デザイン中の操作（選ぶ・動かす・大きさを変える） ----------
    bindStage() {
      this.stage.setAttribute("tabindex", "0");
      this.stage.addEventListener("pointerdown", (e) => {
        if (this.mode !== "design") return;
        const ctl = e.target.closest(".uf-ctl");
        const resize = e.target.closest(".uf-resize");
        const handle = e.target.closest(".uf-handle");
        if (resize) return this.startDrag(e, "form");
        if (ctl) {
          const name = ctl.dataset.name;
          if (this.selected !== name) {
            this.selected = name;
            this.render();
          }
          return this.startDrag(e, handle ? "size" : "move", name);
        }
        if (e.target.closest(".uf-client")) {
          if (this.selected) {
            this.selected = null;
            this.render();
          }
        }
      });
      this.stage.addEventListener("dblclick", (e) => {
        if (this.mode !== "design") return;
        const ctl = e.target.closest(".uf-ctl");
        if (ctl && this.opts.onEditHandler) this.opts.onEditHandler(ctl.dataset.name);
      });
      this.stage.addEventListener("keydown", (e) => {
        if (this.mode !== "design") return;
        if (e.target !== this.stage) return;
        if ((e.key === "Delete" || e.key === "Backspace") && this.selected) {
          e.preventDefault();
          this.deleteSelected();
        } else if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === "z") {
          e.preventDefault();
          this.undo();
        } else if (this.selected && ["ArrowUp", "ArrowDown", "ArrowLeft", "ArrowRight"].includes(e.key)) {
          e.preventDefault();
          const c = this.find(this.selected);
          this.pushHistory();
          const d = e.shiftKey ? 1 : GRID;
          if (e.key === "ArrowUp") c.top = Math.max(0, c.top - d);
          if (e.key === "ArrowDown") c.top = Math.min(this.form.height - c.height, c.top + d);
          if (e.key === "ArrowLeft") c.left = Math.max(0, c.left - d);
          if (e.key === "ArrowRight") c.left = Math.min(this.form.width - c.width, c.left + d);
          this.render();
          this.changed({});
        }
      });
    }

    startDrag(e, kind, name) {
      e.preventDefault();
      this.stage.focus({ preventScroll: true });
      const sx = e.clientX;
      const sy = e.clientY;
      const c = name ? this.find(name) : null;
      const start = c ? { left: c.left, top: c.top, width: c.width, height: c.height } : { width: this.form.width, height: this.form.height };
      const snapshot = JSON.stringify(this.form);
      let moved = false;
      const el = name ? this.stage.querySelector(`.uf-ctl[data-name="${CSS.escape(name)}"]`) : null;
      const win = this.stage.querySelector(".uf-window");
      const client = this.stage.querySelector(".uf-client");
      const onMove = (ev) => {
        const dx = ev.clientX - sx;
        const dy = ev.clientY - sy;
        if (!moved && Math.abs(dx) + Math.abs(dy) < 3) return;
        moved = true;
        if (kind === "move") {
          c.left = Math.max(0, Math.min(this.form.width - c.width, snap(start.left + dx)));
          c.top = Math.max(0, Math.min(this.form.height - c.height, snap(start.top + dy)));
          if (el) {
            el.style.left = c.left + "px";
            el.style.top = c.top + "px";
          }
        } else if (kind === "size") {
          c.width = Math.max(18, Math.min(this.form.width - c.left, snap(start.width + dx)));
          c.height = Math.max(12, Math.min(this.form.height - c.top, snap(start.height + dy)));
          if (el) {
            el.style.width = c.width + "px";
            el.style.height = c.height + "px";
          }
        } else {
          this.form.width = Math.max(160, Math.min(900, snap(start.width + dx)));
          this.form.height = Math.max(100, Math.min(700, snap(start.height + dy)));
          if (win) win.style.width = this.form.width + "px";
          if (client) client.style.height = this.form.height + "px";
        }
      };
      const onUp = () => {
        window.removeEventListener("pointermove", onMove);
        window.removeEventListener("pointerup", onUp);
        if (moved) {
          this.history.push(snapshot);
          this.render();
          this.changed({});
        }
      };
      window.addEventListener("pointermove", onMove);
      window.addEventListener("pointerup", onUp);
    }

    // ---------- プロパティ欄 ----------
    renderProps() {
      const el = this.propsEl;
      if (!el) return;
      if (this.mode === "run") {
        el.innerHTML = '<div class="uf-props-note">▶ 実行中は、プロパティは変えられません（■ 止める と、デザインにもどります）</div>';
        return;
      }
      const c = this.selected ? this.find(this.selected) : null;
      let html = "";
      if (!c) {
        html += `<div class="uf-props-head">🪟 フォーム <b>${esc(this.form.name)}</b></div>`;
        html += '<div class="uf-props-help">上の「部品を置く」のボタンを押すと、フォームに部品が置かれます。部品をクリックすると、ここで名前や文字を変えられます。</div>';
        html += '<div class="uf-props-grid">';
        html += this.field("フォームのタイトル（Caption）", "form:caption", "text", this.form.caption);
        html += this.field("幅（Width）", "form:width", "number", this.form.width);
        html += this.field("高さ（Height）", "form:height", "number", this.form.height);
        html += "</div>";
      } else {
        const p = c.props;
        html += `<div class="uf-props-head">${TYPE_ICON[c.type]} ${TYPE_JA[c.type]} <b>${esc(c.name)}</b><button class="uf-del" data-act="delete" title="この部品を消す（Delete キー）">🗑 消す</button></div>`;
        html += `<div class="uf-props-help">${esc(HELP[c.type] || "")}</div>`;
        html += '<div class="uf-props-grid">';
        html += this.field("名前（Name）", "name", "text", c.name, "プログラムで使う名前です（英字で始める）");
        if (c.type === "TextBox") html += this.field("入力された文字（Text）", "prop:Text", "text", p.Text);
        else if (c.type !== "ListBox") html += this.field("表示する文字（Caption）", "prop:Caption", "text", p.Caption);
        html += this.field("文字の大きさ（Font）", "prop:FontSize", "number", p.FontSize || 12);
        html += this.field("太字（Font）", "prop:FontBold", "checkbox", !!p.FontBold);
        html += this.field("文字の色（ForeColor）", "prop:ForeColor", "color", vbaColorToHex(p.ForeColor) || "#000000");
        html += this.field("背景の色（BackColor）", "prop:BackColor", "color", vbaColorToHex(p.BackColor) || (c.type === "TextBox" ? "#ffffff" : c.type === "CommandButton" ? "#e1e1e1" : "#f0f0f0"));
        if (c.type === "Label" || c.type === "TextBox") {
          html += this.select("文字のそろえ方（TextAlign）", "prop:TextAlign", p.TextAlign || 1, [[1, "左"], [2, "中央"], [3, "右"]]);
        }
        if (c.type === "Label") {
          html += this.select("枠線（BorderStyle）", "prop:BorderStyle", p.BorderStyle || 0, [[0, "なし"], [1, "あり"]]);
          html += this.select("見た目（SpecialEffect）", "prop:SpecialEffect", p.SpecialEffect || 0, [[0, "平ら"], [1, "浮き出し"], [2, "へこみ"], [3, "溝"], [6, "盛り上がり"]]);
        }
        html += this.field("横の位置（Left）", "geo:left", "number", c.left);
        html += this.field("縦の位置（Top）", "geo:top", "number", c.top);
        html += this.field("幅（Width）", "geo:width", "number", c.width);
        html += this.field("高さ（Height）", "geo:height", "number", c.height);
        html += "</div>";
        html += '<div class="uf-name-error" hidden></div>';
      }
      el.innerHTML = html;
      this.bindProps(c);
    }

    field(label, key, type, value, title) {
      const t = title ? ` title="${esc(title)}"` : "";
      if (type === "checkbox") {
        return `<label class="uf-prop"${t}><span>${esc(label)}</span><input type="checkbox" data-key="${key}" ${value ? "checked" : ""}></label>`;
      }
      return `<label class="uf-prop"${t}><span>${esc(label)}</span><input type="${type}" data-key="${key}" value="${esc(value)}"></label>`;
    }
    select(label, key, value, options) {
      const opts = options.map(([v, l]) => `<option value="${v}" ${Number(value) === v ? "selected" : ""}>${esc(l)}</option>`).join("");
      return `<label class="uf-prop"><span>${esc(label)}</span><select data-key="${key}">${opts}</select></label>`;
    }

    bindProps(c) {
      const el = this.propsEl;
      const del = el.querySelector('[data-act="delete"]');
      if (del) del.addEventListener("click", () => this.deleteSelected());
      el.querySelectorAll("[data-key]").forEach((inp) => {
        const key = inp.dataset.key;
        let snapshot = null;
        const apply = (commit) => {
          if (!snapshot) snapshot = JSON.stringify(this.form);
          const v = inp.type === "checkbox" ? inp.checked : inp.value;
          if (key === "name") {
            if (commit) this.renameSelected(v, inp, snapshot);
            return;
          }
          const [kind, prop] = key.split(":");
          if (kind === "form") {
            if (prop === "caption") this.form.caption = v;
            else this.form[prop] = Math.max(prop === "width" ? 160 : 100, Math.min(900, Number(v) || 0));
          } else if (kind === "geo") {
            c[prop] = Math.max(0, Number(v) || 0);
          } else if (prop === "ForeColor" || prop === "BackColor") {
            c.props[prop] = hexToVbaColor(v);
          } else if (prop === "FontSize") {
            c.props.FontSize = Math.max(6, Math.min(72, Number(v) || 12));
          } else if (prop === "FontBold") {
            c.props.FontBold = !!v;
          } else if (["TextAlign", "BorderStyle", "SpecialEffect"].includes(prop)) {
            c.props[prop] = Number(v);
          } else {
            c.props[prop] = v;
          }
          this.renderStage();
          if (commit) {
            this.history.push(snapshot);
            snapshot = null;
            this.changed({});
          }
        };
        if (key === "name") {
          inp.addEventListener("change", () => apply(true));
          inp.addEventListener("keydown", (e) => {
            if (e.key === "Enter") inp.blur();
          });
          return;
        }
        inp.addEventListener("input", () => apply(false));
        inp.addEventListener("change", () => apply(true));
      });
    }

    renameSelected(newName, inp, snapshot) {
      const c = this.find(this.selected);
      if (!c) return;
      const name = String(newName).trim();
      const errEl = this.propsEl.querySelector(".uf-name-error");
      const fail = (msg) => {
        if (errEl) {
          errEl.textContent = "⚠️ " + msg;
          errEl.hidden = false;
        }
        inp.value = c.name;
      };
      if (name === c.name) return;
      if (!/^[A-Za-z][A-Za-z0-9_]{0,39}$/.test(name)) return fail("名前は英字で始めて、英字・数字・_ だけで書こう（例: lblKekka, btnKeisan）");
      if (RESERVED.has(name.toLowerCase())) return fail(`「${name}」は VBA で使う言葉なので、部品の名前にはできません`);
      if (this.form.controls.some((x) => x !== c && x.name.toLowerCase() === name.toLowerCase())) return fail(`「${name}」という部品はもうあります`);
      this.history.push(snapshot);
      const from = c.name;
      c.name = name;
      if ((c.type === "Label" || c.type === "CommandButton") && c.props.Caption === from) c.props.Caption = name;
      this.selected = name;
      this.render();
      this.changed({ structure: true, renamed: { from, to: name } });
    }
  }

  window.UserFormPanel = UserFormPanel;
  window.UserFormUtil = { defaultForm, toRuntimeForm, vbaColorToHex, hexToVbaColor, TYPE_JA };
})();
