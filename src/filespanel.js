/* ===== 📁 ファイル（仮想のファイル置き場） =====
 * プログラムの Open ～ For Input で読めるファイル（CSV など）をここに置く。
 * 本物のパスは使わず、ファイル名だけで照らし合わせる（ThisWorkbook.Path & "¥uriage.csv" も読める）。
 * プログラムが Open ～ For Output で書き出したファイルは、下に（読むだけで）表示する。
 *
 *   new FilesPanel({ list, out, onChange(files), onToast(msg) })
 *   .setFiles(obj) / .getFiles() / .setOutputs(obj) / .addFile(name, text)
 */
(function () {
  "use strict";

  const MAX_FILE = 100 * 1024; // 1ファイルの上限（文字数の目安）
  const MAX_TOTAL = 300 * 1024; // 合計の上限（localStorage の容量対策）

  // 教科書 p.127 の例題 4-3-06 の入力ファイル（購入者名と数量）
  const SAMPLE_URIAGE =
    "高橋　礼,4\r\n山口　桃子,5\r\n奥原　希美,2\r\n松友　美紀,7\r\n永原　和香,8\r\n松本　友子,10\r\n米本　小奈津,23\r\n田中　志保子,5\r\n東野　亜里沙,13\r\n栗原　歩美,20\r\n広田　花,9\r\n";

  function esc(s) {
    return String(s).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");
  }

  // バイト列を文字にする：まず UTF-8 として厳密に読み、だめなら Shift_JIS（Excel の CSV で多い）
  function decodeBytes(buf) {
    try {
      return { text: new TextDecoder("utf-8", { fatal: true }).decode(buf), enc: "UTF-8" };
    } catch (e) {
      try {
        return { text: new TextDecoder("shift_jis").decode(buf), enc: "Shift_JIS" };
      } catch (e2) {
        return { text: new TextDecoder("utf-8").decode(buf), enc: "UTF-8" };
      }
    }
  }

  class FilesPanel {
    constructor(opts) {
      this.opts = opts;
      this.files = {};
      this.outputs = {};
      this.outKey = "";
      this.render();
    }

    getFiles() {
      return JSON.parse(JSON.stringify(this.files));
    }
    setFiles(obj) {
      this.files = obj ? JSON.parse(JSON.stringify(obj)) : {};
      this.render();
    }
    setOutputs(obj) {
      const key = JSON.stringify(obj || {});
      if (key === this.outKey) return;
      this.outKey = key;
      this.outputs = obj ? JSON.parse(key) : {};
      this.renderOut();
    }
    total() {
      return Object.keys(this.files).reduce((n, k) => n + this.files[k].length + k.length, 0);
    }
    toast(msg) {
      if (this.opts.onToast) this.opts.onToast(msg);
    }
    changed() {
      if (this.opts.onChange) this.opts.onChange(this.getFiles());
    }

    uniqueName(name) {
      const dot = name.lastIndexOf(".");
      const base = dot > 0 ? name.slice(0, dot) : name;
      const ext = dot > 0 ? name.slice(dot) : "";
      let n = name;
      let k = 2;
      while (Object.keys(this.files).some((f) => f.toLowerCase() === n.toLowerCase())) n = `${base}${k++}${ext}`;
      return n;
    }

    addFile(name, text) {
      const clean = String(name || "data.csv").replace(/[\\/:*?"<>|]/g, "_").trim() || "data.csv";
      if (text.length > MAX_FILE) {
        this.toast(`「${clean}」は大きすぎて入れられません（${Math.round(MAX_FILE / 1024)}KB まで）`);
        return false;
      }
      if (this.total() + text.length > MAX_TOTAL) {
        this.toast("ファイルが多すぎて、これ以上は入れられません。いらないファイルを消してね");
        return false;
      }
      const n = this.uniqueName(clean);
      this.files[n] = text;
      this.render();
      this.changed();
      return n;
    }

    deleteFile(name) {
      delete this.files[name];
      this.render();
      this.changed();
    }

    renameFile(from, to) {
      const clean = String(to).replace(/[\\/:*?"<>|]/g, "_").trim();
      if (!clean || clean === from) return false;
      if (Object.keys(this.files).some((f) => f !== from && f.toLowerCase() === clean.toLowerCase())) {
        this.toast(`「${clean}」というファイルはもうあります`);
        return false;
      }
      const next = {};
      Object.keys(this.files).forEach((k) => (next[k === from ? clean : k] = this.files[k]));
      this.files = next;
      this.render();
      this.changed();
      return true;
    }

    // PC のファイルを読みこむ
    readFromPc(fileList) {
      Array.from(fileList).forEach((f) => {
        const reader = new FileReader();
        reader.onload = () => {
          const { text, enc } = decodeBytes(reader.result);
          const n = this.addFile(f.name, text);
          if (n) this.toast(`📂 「${n}」を読みこみました（文字コード: ${enc}）`);
        };
        reader.readAsArrayBuffer(f);
      });
    }

    addSample() {
      const n = this.addFile("uriage.csv", SAMPLE_URIAGE);
      if (n) this.toast(`📁 見本の「${n}」を入れました（教科書 例題 4-3-06 の入力ファイル）`);
    }

    // ---------- 描画 ----------
    render() {
      const el = this.opts.list;
      const names = Object.keys(this.files);
      if (!names.length) {
        el.innerHTML =
          '<div class="files-empty">まだファイルがありません。<br>「＋ 新しいファイル」で作るか、「📂 PCのファイルを読みこむ」で CSV を入れよう。<br>教科書の例題 4-3-06 では、見本の <b>uriage.csv</b> を使います。</div>';
        return;
      }
      el.innerHTML = names
        .map((n) => {
          const text = this.files[n];
          const lines = text === "" ? 0 : text.replace(/\r?\n$/, "").split(/\r?\n/).length;
          return `<div class="files-item" data-name="${esc(n)}">
            <div class="files-item-head">
              <span class="files-icon">📄</span>
              <input class="files-name" type="text" value="${esc(n)}" aria-label="ファイル名" title="ファイルの名前（Open で書く名前）">
              <span class="files-meta">${lines} 行</span>
              <button class="files-del" title="このファイルを消す">🗑</button>
            </div>
            <textarea class="files-text" spellcheck="false" aria-label="${esc(n)} の中身" wrap="off">${esc(text)}</textarea>
          </div>`;
        })
        .join("");
      el.querySelectorAll(".files-item").forEach((item) => {
        const name = item.dataset.name;
        item.querySelector(".files-del").addEventListener("click", () => {
          if (confirm(`「${name}」を消しますか？`)) this.deleteFile(name);
        });
        const nameInput = item.querySelector(".files-name");
        nameInput.addEventListener("change", () => {
          if (!this.renameFile(name, nameInput.value)) nameInput.value = name;
        });
        const ta = item.querySelector(".files-text");
        ta.addEventListener("input", () => {
          if (ta.value.length > MAX_FILE) ta.value = ta.value.slice(0, MAX_FILE);
          this.files[name] = ta.value;
          const lines = ta.value === "" ? 0 : ta.value.replace(/\r?\n$/, "").split(/\r?\n/).length;
          item.querySelector(".files-meta").textContent = `${lines} 行`;
          this.changed();
        });
      });
    }

    renderOut() {
      const el = this.opts.out;
      if (!el) return;
      const names = Object.keys(this.outputs);
      if (!names.length) {
        el.hidden = true;
        el.innerHTML = "";
        return;
      }
      el.hidden = false;
      el.innerHTML =
        '<div class="files-out-title">✏️ プログラムが書き出したファイル（読むだけ）</div>' +
        names
          .map(
            (n) => `<div class="files-item out" data-name="${esc(n)}">
          <div class="files-item-head"><span class="files-icon">📝</span><span class="files-name-fixed">${esc(n)}</span>
          <button class="files-keep" title="このファイルを、プログラムから読めるファイルにする">📌 読めるファイルにする</button></div>
          <textarea class="files-text" readonly wrap="off">${esc(this.outputs[n])}</textarea></div>`
          )
          .join("");
      el.querySelectorAll(".files-keep").forEach((b) => {
        b.addEventListener("click", () => {
          const n = b.closest(".files-item").dataset.name;
          const added = this.addFile(n, this.outputs[n]);
          if (added) this.toast(`📌 「${added}」を、読めるファイルに入れました`);
        });
      });
    }
  }

  window.FilesPanel = FilesPanel;
  window.FilesUtil = { SAMPLE_URIAGE, decodeBytes };
})();
