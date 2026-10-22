/* ===== メインアプリ =====
 * Blockly 初期化・課題切り替え・ヒント・コード生成・Excel プレビューを結線する。
 */

(function () {
  let workspace = null;
  let view = null;
  let codeEditor = null; // CodeMirror インスタンス（VBAコード直打ち）
  let codeDirty = false; // コード側が編集されてブロックへ未反映かどうか
  let suppressCodeChange = false; // setValue によるプログラム的な更新中フラグ
  let dedentTimer = null; // 保留中の自動デデント（setTimeout ID）
  let dedentTimerLine = null; // ↑がどの行に対する処理か
  let currentTaskId = null;
  let hasLoadedOnce = false; // 初回ロード時は保存前のブロックが存在しないため保存をスキップ
  let suppressSave = false; // 復元中の保存抑制フラグ（B7）
  let sharedViewMode = false; // 共有リンク閲覧中は保存スロットに書き込まない
  const inputCache = {}; // InputBox の答え（block.id -> 値）。▶実行ごとにクリア
  const savedBlocks = {}; // taskId -> XML文字列（課題ごとのセーブスロット）
  const savedInitial = {}; // taskId -> 仮想Excelの初期データ
  const solved = new Set();

  // ----- 実行エンジン（src/vba/）まわり -----
  let session = null; // いま画面に出している実行（VbaSession.Session）
  let runActive = false; // ▶実行・1コマで始めた実行の途中か（false なら編集中のプレビュー）
  let lineMap = {}; // コードの行番号 → ブロックID（ブロックから作ったコードのとき）
  // 実行するプログラムの元。"blocks" = ブロックから作ったコード / "code" = VBAコードタブで書いたコード
  let codeOwner = "blocks";
  let blocksMatchCode = true; // codeOwner が "code" のとき、ブロックがコードと同じ内容か
  let blocksMismatchReason = ""; // ブロックにできなかった理由
  let previewTimer = null;
  let codeSaveTimer = null;
  let execLine = null; // 実行中として光らせているコードの行（0始まり）
  let errorMarkLine = null; // エラーとして光らせているコードの行（0始まり）
  let zenkakuMarks = []; // 全角文字の波線
  const savedCode = {}; // taskId -> { text, match, reason }（VBAコードタブで書いたコード）
  const savedForms = {}; // taskId -> フォームのデータ（userform.js の form）
  let formPanel = null; // 🪟 ユーザーフォームのデザイナー兼実行画面
  let lastArea = "form"; // 3ステップ表示用：最後にさわった場所（"form" / "program"）
  let flowChart = null; // 🔀 流れ図
  let flowAst = null; // 流れ図にするプログラム（AST）
  let flowMsg = null; // 描けないときの説明
  let flowDirty = true; // 描き直しが必要か（タブが見えているときだけ描く）
  let hoverLine = null; // 流れ図の記号にマウスを乗せて光らせているコードの行（0始まり）
  let runTarget = null; // ▶実行で動かすもの："form"（フォームを開く）/ "macro:名前"
  const LS_FORMS = "umb_forms";
  const LOAD_GROUP = "umb-code-to-blocks"; // コード→ブロック変換で起きたブロックの変化の目印
  // 状態を変える step（これが1つも無ければ「プログラムを実行した結果」とはみなさない）
  const STATE_SCOPES = new Set(["cell", "array", "control", "sheet", "file", "msg", "var"]);

  const LS_BLOCKS = "umb_blocks";
  const LS_SOLVED = "umb_solved";
  const LS_INITIAL = "umb_initial";
  const LS_CODE = "umb_code";
  const FREE_MODE_ID = "__free__"; // クエスト未選択時（フリーモード）の保存キー

  // ----- テーマ（一度だけ定義: B5）-----
  let _themes = null;
  function getThemes() {
    if (_themes) return _themes;
    _themes = {
      // アプリの Catppuccin パレットに合わせる
      dark: Blockly.Theme.defineTheme("umb_dark", {
        base: Blockly.Themes.Classic,
        componentStyles: {
          workspaceBackgroundColour: "#11111b",
          toolboxBackgroundColour: "#181825",
          toolboxForegroundColour: "#cdd6f4",
          flyoutBackgroundColour: "#1e1e2e",
          flyoutForegroundColour: "#cdd6f4",
          scrollbarColour: "#313244",
        },
      }),
      light: Blockly.Theme.defineTheme("umb_light", {
        base: Blockly.Themes.Classic,
        componentStyles: {
          workspaceBackgroundColour: "#f8f9fa",
          toolboxBackgroundColour: "#ffffff",
          toolboxForegroundColour: "#4c4f69",
          flyoutBackgroundColour: "#eff1f5",
          flyoutForegroundColour: "#4c4f69",
          scrollbarColour: "#dce0e8",
        },
      }),
    };
    return _themes;
  }

  // ----- テーマ管理 -----
  function initTheme() {
    const saved = localStorage.getItem("theme") || "dark";
    applyTheme(saved);
  }
  function applyTheme(mode) {
    const themes = workspace ? getThemes() : null;
    if (mode === "light") {
      document.body.classList.add("light-mode");
      document.getElementById("theme-toggle-btn").textContent = "☀️";
      if (workspace) workspace.setTheme(themes.light);
      updateWorkspaceGridColour(false);
    } else {
      document.body.classList.remove("light-mode");
      document.getElementById("theme-toggle-btn").textContent = "🌙";
      if (workspace) workspace.setTheme(themes.dark);
      updateWorkspaceGridColour(true);
    }
    localStorage.setItem("theme", mode);
  }

  // ----- VBAコード直打ちエディタの文字サイズ（好みで変更可能） -----
  const LS_CODE_FONT = "umb_codeFontSize";
  const CODE_FONT_MIN = 12;
  const CODE_FONT_MAX = 50;
  const CODE_FONT_STEP = 2;
  const CODE_FONT_DEFAULT = 20;
  function getCodeFontSize() {
    const saved = Number(localStorage.getItem(LS_CODE_FONT));
    return saved && saved >= CODE_FONT_MIN && saved <= CODE_FONT_MAX ? saved : CODE_FONT_DEFAULT;
  }
  function applyCodeFontSize(size) {
    document.documentElement.style.setProperty("--code-font-size", `${size}px`);
    const label = document.getElementById("code-font-size-value");
    if (label) label.textContent = size;
    if (codeEditor) codeEditor.refresh();
  }
  function initCodeFontSize() {
    applyCodeFontSize(getCodeFontSize());
    const dec = document.getElementById("code-font-dec");
    const inc = document.getElementById("code-font-inc");
    if (dec) {
      dec.addEventListener("click", () => {
        const next = Math.max(CODE_FONT_MIN, getCodeFontSize() - CODE_FONT_STEP);
        localStorage.setItem(LS_CODE_FONT, next);
        applyCodeFontSize(next);
      });
    }
    if (inc) {
      inc.addEventListener("click", () => {
        const next = Math.min(CODE_FONT_MAX, getCodeFontSize() + CODE_FONT_STEP);
        localStorage.setItem(LS_CODE_FONT, next);
        applyCodeFontSize(next);
      });
    }
  }

  // ----- コード直打ちエディタのグリッド線（1文字＝1マスの目安表示） -----
  // インデントが4マスになっているか、全角/半角の文字幅の違いを目で確認できるようにする
  const LS_CODE_GRID = "umb_codeGridLines";
  function applyCodeGrid(enabled) {
    const host = document.getElementById("code-editor");
    if (host) host.classList.toggle("show-grid", enabled);
  }
  function initCodeGrid() {
    const enabled = localStorage.getItem(LS_CODE_GRID) === "1";
    applyCodeGrid(enabled);
    const checkbox = document.getElementById("code-grid-toggle");
    if (checkbox) {
      checkbox.checked = enabled;
      checkbox.addEventListener("change", () => {
        localStorage.setItem(LS_CODE_GRID, checkbox.checked ? "1" : "0");
        applyCodeGrid(checkbox.checked);
      });
    }
  }

  // ----- 仮想Excel表示エリアの幅（ドラッグで調整可能） -----
  const LS_RIGHT_PANE_WIDTH = "umb_rightPaneWidth";
  const RIGHT_PANE_MIN = 360;
  const RIGHT_PANE_MAX = 1200;
  function applyRightPaneWidth(width) {
    document.documentElement.style.setProperty("--right-pane-width", `${width}px`);
  }
  function initRightPaneResize() {
    const resizer = document.getElementById("pane-resizer");
    if (!resizer) return;

    const saved = Number(localStorage.getItem(LS_RIGHT_PANE_WIDTH));
    if (saved && saved >= RIGHT_PANE_MIN && saved <= RIGHT_PANE_MAX) {
      applyRightPaneWidth(saved);
    }

    resizer.addEventListener("pointerdown", (e) => {
      e.preventDefault();
      resizer.classList.add("dragging");
      resizer.setPointerCapture(e.pointerId);

      const onMove = (moveEvent) => {
        // 仮想Excelはビューポート右端に接しているため、マウスX座標から右端までの
        // 距離＝そのまま欲しい幅になる
        const maxWidth = Math.min(RIGHT_PANE_MAX, window.innerWidth - 300);
        const width = Math.round(
          Math.max(RIGHT_PANE_MIN, Math.min(maxWidth, window.innerWidth - moveEvent.clientX))
        );
        applyRightPaneWidth(width);
      };
      const onUp = (upEvent) => {
        resizer.classList.remove("dragging");
        resizer.releasePointerCapture(upEvent.pointerId);
        resizer.removeEventListener("pointermove", onMove);
        resizer.removeEventListener("pointerup", onUp);
        const current = getComputedStyle(document.documentElement).getPropertyValue("--right-pane-width");
        const px = parseInt(current, 10);
        if (px) localStorage.setItem(LS_RIGHT_PANE_WIDTH, px);
      };
      resizer.addEventListener("pointermove", onMove);
      resizer.addEventListener("pointerup", onUp);
    });
  }

  // workspace.setTheme() はグリッド線の色を更新しないため、切替のたびに手動で合わせる
  // （合わせないと、ライト→ダークと往復したときにグリッドがほぼ見えなくなる）
  function updateWorkspaceGridColour(isDark) {
    if (!workspace) return;
    try {
      const grid = workspace.getGrid && workspace.getGrid();
      // Blockly 11.2.1 の Grid クラスには色を変更する公開APIが無いため、
      // 生成済みの SVG 線要素（line1/line2）の stroke 属性を直接書き換える
      const colour = isDark ? "#333" : "#ddd";
      if (grid && grid.line1 && grid.line1.setAttribute) grid.line1.setAttribute("stroke", colour);
      if (grid && grid.line2 && grid.line2.setAttribute) grid.line2.setAttribute("stroke", colour);
    } catch (e) {
      // Blockly のバージョン差で失敗しても致命的ではないので握りつぶす
      console.warn("グリッド色の更新に失敗:", e);
    }
  }

  // ----- 永続化（B4 / F1）-----
  function persist() {
    try {
      localStorage.setItem(LS_BLOCKS, JSON.stringify(savedBlocks));
      localStorage.setItem(LS_SOLVED, JSON.stringify([...solved]));
      localStorage.setItem(LS_INITIAL, JSON.stringify(savedInitial));
      localStorage.setItem(LS_CODE, JSON.stringify(savedCode));
      localStorage.setItem(LS_FORMS, JSON.stringify(savedForms));
    } catch (e) {
      console.warn("保存失敗:", e);
    }
  }
  function loadPersisted() {
    try {
      const b = JSON.parse(localStorage.getItem(LS_BLOCKS) || "{}");
      Object.assign(savedBlocks, b);
      const s = JSON.parse(localStorage.getItem(LS_SOLVED) || "[]");
      s.forEach((id) => solved.add(id));
      const ini = JSON.parse(localStorage.getItem(LS_INITIAL) || "{}");
      Object.assign(savedInitial, ini);
      const code = JSON.parse(localStorage.getItem(LS_CODE) || "{}");
      Object.assign(savedCode, code);
      const forms = JSON.parse(localStorage.getItem(LS_FORMS) || "{}");
      Object.assign(savedForms, forms);
    } catch (e) {
      console.warn("読み込み失敗:", e);
    }
  }

  function currentKey() {
    return currentTaskId === null ? FREE_MODE_ID : currentTaskId;
  }

  // VBAコードタブで書いたコードを保存する（ブロックから作ったコードのときは保存しない）
  function saveCodeState() {
    if (sharedViewMode || !codeEditor) return;
    const key = currentKey();
    if (codeOwner === "code") {
      savedCode[key] = { text: codeEditor.getValue(), match: blocksMatchCode, reason: blocksMismatchReason };
    } else {
      delete savedCode[key];
    }
    persist();
  }

  // 課題を切り替えたとき：保存しておいたコードがあれば戻す
  function restoreCodeState(key) {
    const sc = savedCode[key];
    codeDirty = false;
    if (sc && typeof sc.text === "string") {
      codeOwner = "code";
      blocksMatchCode = !!sc.match;
      blocksMismatchReason = sc.reason || "";
      setEditorCode(sc.text);
    } else {
      codeOwner = "blocks";
      blocksMatchCode = true;
      blocksMismatchReason = "";
    }
    updateCodeSyncStatus();
    updateCodeOwnerOverlay();
  }

  function setEditorCode(text) {
    if (!codeEditor) return;
    if (codeEditor.getValue() === text) return;
    suppressCodeChange = true;
    codeEditor.setValue(text);
    suppressCodeChange = false;
  }

  // ----- 仮想Excelのセル編集時（F: 直接入力）-----
  function onCellEdited() {
    savedInitial[currentKey()] = view.getInitialCells();
    persist();
    // 初期データが変わったので実行し直す
    runPreview();
    updateBaselineLabel();
  }

  // ----- 実行タイムライン（スライダー・カウンタ）の同期 -----
  function updateTimeline(cursor, total) {
    const slider = document.getElementById("tl-slider");
    const count = document.getElementById("tl-count");
    if (slider) {
      slider.max = String(total);
      slider.value = String(cursor);
      slider.disabled = total === 0;
    }
    if (count) count.textContent = `${cursor} / ${total}`;
    document.querySelectorAll(".tl-btn").forEach((b) => (b.disabled = total === 0));
  }

  // ----- 説明とヒントパネルの開閉 -----
  function applyQuestCollapsed(collapsed) {
    const collapseBtn = document.getElementById("quest-collapse-btn");
    document.getElementById("main-layout").classList.toggle("quest-collapsed", collapsed);
    if (collapseBtn) {
      collapseBtn.textContent = collapsed ? "📖 説明とヒント ▸" : "📖 説明とヒント ▾";
    }
    localStorage.setItem("questCollapsed", collapsed ? "1" : "0");
    if (workspace) setTimeout(() => Blockly.svgResize(workspace), 50);
  }

  // ----- ブロックが1つも無いときだけキャンバスに案内を出す -----
  function updateCanvasGuide() {
    const guide = document.getElementById("canvas-guide");
    if (!guide || !workspace) return;
    guide.hidden = workspace.getTopBlocks(false).length > 0;
  }

  // ----- 「＋ 変数を作る」ボタンの新規変数名プロンプト（ひらがな・予約名を検証） -----
  function promptNewVariableName(ws) {
    const raw = window.prompt("新しい変数の名前を入力してください（例: Tanka、Suu、Kin。ひらがな不可、i・j・k・arr は使えません）", "");
    if (raw === null) return; // キャンセル
    const name = raw.trim();
    if (!name) return;
    const errMsg = window.UMB_validateVarName ? window.UMB_validateVarName(name) : null;
    if (errMsg) {
      alert(errMsg);
      return promptNewVariableName(ws); // 直してもう一度
    }
    ws.createVariable(name); // 既に同名があれば Blockly 側でそのまま再利用される
  }

  // ----- 「実行の出発点」表示の更新 -----
  function updateBaselineLabel() {
    const el = document.getElementById("baseline-label");
    if (!el || !view) return;
    const n = Object.keys(view.getInitialCells() || {}).length;
    el.textContent = n ? `📌 出発点: ${n} セル` : "📌 出発点: なし";
  }

  // ===== VBAコード直打ち：入力アシスト（スニペット・自動インデント） =====

  // キーワードを打つと自動で例文候補がポップアップ表示され（CodeMirrorの
  // show-hintアドオン）、選ぶとその場で展開される骨組み集。
  // Tabキーでの直接展開（ポップアップを使わない打ち方）にも同じデータを使う。
  // label: ポップアップに出す見出し / desc: 一言説明
  // cursorLine/cursorCol は挿入後にカーソルを置く位置（0行目基準、baseIndent加算前）
  const CODE_SNIPPETS = {
    for: {
      template: "For i = 1 To 10\n    \nNext i",
      cursorLine: 1, cursorCol: 4,
      label: "For i = ... To ... / Next i",
      desc: "決まった回数くり返す",
    },
    while: {
      template: "Do While \nLoop",
      cursorLine: 0, cursorCol: 9,
      label: "Do While ... / Loop",
      desc: "条件が真の間くり返す",
    },
    until: {
      template: "Do Until \nLoop",
      cursorLine: 0, cursorCol: 9,
      label: "Do Until ... / Loop",
      desc: "条件が真になるまでくり返す",
    },
    if: {
      template: "If  Then\nEnd If",
      cursorLine: 0, cursorCol: 3,
      label: "If ... Then / End If",
      desc: "もし〜なら",
    },
    ifelse: {
      template: "If  Then\nElse\nEnd If",
      cursorLine: 0, cursorCol: 3,
      label: "If ... Then / Else / End If",
      desc: "もし〜なら／そうでなければ",
    },
    dim: {
      template: "Dim 変数名 As Variant",
      cursorLine: 0, cursorCol: 4,
      label: "Dim ... As ...",
      desc: "変数を用意する（このアプリでは自動で付くので、通常は書かなくてOK）",
    },
  };

  // 現在行のインデントを引き継ぎつつスニペットを展開する。
  // 注意: CODE_SNIPPETS のテンプレート文字列自体に、行ごとの相対インデント
  // （本体行は4スペース、閉じキーワード行は0）を埋め込んである。
  // ここでは各行の先頭に baseIndent を足すだけでよい（さらに4スペース加算しない）
  //
  // from/to を省略した場合（Tabキーでの直接展開）は「今まさに打ち終えた
  // キーワード＝key の文字数ぶん」をカーソルから遡って置換対象にする。
  // ヒントポップアップからの選択時は、実際にタイプされていた文字数が
  // key と一致するとは限らない（例: "fo" まで打った時点で "for" を選ぶ等）ため、
  // 呼び出し側が計算した実際の from/to を明示的に渡す
  function insertCodeSnippet(cm, key, from, to) {
    const snip = CODE_SNIPPETS[key];
    if (!snip) return false;
    const cursor = to || cm.getCursor();
    const line = cm.getLine(cursor.line);
    const baseIndent = (line.match(/^(\s*)/) || ["", ""])[1];
    const wordStart = from || { line: cursor.line, ch: cursor.ch - key.length };
    const lines = snip.template.split("\n");
    const indented = lines.map((l, i) => (i === 0 ? l : baseIndent + l));
    cm.replaceRange(indented.join("\n"), wordStart, cursor);
    const targetLine = wordStart.line + snip.cursorLine;
    const targetCh = (snip.cursorLine === 0 ? wordStart.ch : baseIndent.length) + snip.cursorCol;
    cm.setCursor({ line: targetLine, ch: targetCh });
    return true;
  }

  // 今カーソルの直前にある単語（アルファベット列）と、その開始位置を返す
  function currentWordBeforeCursor(cm) {
    const cursor = cm.getCursor();
    const line = cm.getLine(cursor.line);
    const before = line.slice(0, cursor.ch);
    const m = before.match(/([A-Za-z]+)$/);
    if (!m) return null;
    return { word: m[1], from: { line: cursor.line, ch: cursor.ch - m[1].length }, to: cursor };
  }

  // For/Dim 等のキーワードを打つと、CodeMirror の show-hint アドオンで
  // 例文候補をその場にポップアップ表示する（入力アシスト）。
  // 前方一致するキーワードが1つも無ければ null を返し、ポップアップは出さない
  if (typeof CodeMirror !== "undefined" && CodeMirror.registerHelper) {
    CodeMirror.registerHelper("hint", "vbaSnippet", function (cm) {
      const cur = currentWordBeforeCursor(cm);
      if (!cur || !cur.word) return null;
      const wordLower = cur.word.toLowerCase();
      const matches = Object.keys(CODE_SNIPPETS).filter((k) => k.startsWith(wordLower));
      if (matches.length === 0) return null;
      return {
        list: matches.map((key) => {
          const snip = CODE_SNIPPETS[key];
          return {
            // text はプレーンテキスト表示のフォールバック用（render を使うので通常は表示されない）
            text: snip.label,
            displayText: snip.label,
            className: "vba-hint-item",
            render: function (el) {
              const title = document.createElement("div");
              title.className = "vba-hint-title";
              title.textContent = snip.label;
              const desc = document.createElement("div");
              desc.className = "vba-hint-desc";
              desc.textContent = snip.desc;
              el.appendChild(title);
              el.appendChild(desc);
            },
            // 選択されたら insertCodeSnippet で実際に展開する
            // （from/to はポップアップを開いた時点のものではなく、選択時点の
            // 最新カーソル位置を使う。表示中にさらに文字を打ち足した場合に対応するため）
            hint: function (cm2) {
              const latest = currentWordBeforeCursor(cm2) || cur;
              insertCodeSnippet(cm2, key, latest.from, latest.to);
            },
          };
        }),
        from: cur.from,
        to: cur.to,
      };
    });
  }

  // Tab: 直前の単語がスニペットキーワードならそれを展開。それ以外は通常インデント
  function handleCodeTabKey(cm) {
    if (cm.somethingSelected()) {
      cm.execCommand("indentMore");
      return;
    }
    const cursor = cm.getCursor();
    const line = cm.getLine(cursor.line);
    const before = line.slice(0, cursor.ch);
    const m = before.match(/([A-Za-z]+)$/);
    if (m && CODE_SNIPPETS[m[1].toLowerCase()] && m[1].toLowerCase() === m[1]) {
      if (insertCodeSnippet(cm, m[1])) return;
    }
    cm.replaceSelection("    ");
  }

  // Enter: For / Do While・Until / If...Then / Else の次行は自動でインデントを1段深くする
  function handleCodeEnterKey(cm) {
    // 閉じキーワード（Next/Loop/End If/Else）を打った直後、間髪入れずEnterを押すと、
    // change イベント経由の setTimeout(0) デデント（dedentClosingKeywordLine）が
    // まだ実行されていない古いインデント状態を読んでしまい、次の行までインデントが
    // ズレて引き継がれてしまう競合があった。Enter処理の先頭で同期的にデデントを
    // 確定させてから改行することで解消する。
    // 注意: dedentClosingKeywordLine は「呼ばれるたびに無条件で4スペース減らす」
    // 実装で冪等ではない（2回呼ぶと8段階減ってしまう）。保留中の非同期デデントが
    // 後から同じ行にもう一度効いてしまわないよう、ここでタイマーを止めてから呼ぶ
    const enterLine = cm.getCursor().line;
    if (dedentTimer && dedentTimerLine === enterLine) {
      clearTimeout(dedentTimer);
      dedentTimer = null;
    }
    dedentClosingKeywordLine(cm, enterLine);

    const cursor = cm.getCursor(); // 上のデデントで行が短くなっている場合があるので取り直す
    const line = cm.getLine(cursor.line);
    const before = line.slice(0, cursor.ch);
    const trimmed = before.trim();
    const curIndent = (before.match(/^(\s*)/) || ["", ""])[1];
    const opensBlock =
      /^For\s+/i.test(trimmed) ||
      /^Do\s+(While|Until)\s+/i.test(trimmed) ||
      /\bThen$/i.test(trimmed) ||
      /^Else$/i.test(trimmed);
    const newIndent = opensBlock ? curIndent + "    " : curIndent;
    cm.replaceSelection("\n" + newIndent);
  }

  // Next / Loop / End If / Else を打ち終えた行は、その場でインデントを1段浅くする
  function dedentClosingKeywordLine(cm, lineNo) {
    const line = cm.getLine(lineNo);
    if (line === undefined) return;
    const m = line.match(/^(\s*)(Next(?:\s+i)?|Loop|End If|Else)\s*$/i);
    if (!m) return;
    const cur = m[1];
    const desired = cur.length >= 4 ? cur.slice(4) : "";
    if (desired === cur) return;
    const keyword = m[2];
    cm.replaceRange(desired + keyword, { line: lineNo, ch: 0 }, { line: lineNo, ch: cur.length + keyword.length });
  }

  // ----- コードの状態の表示（コードタブの見出しの横） -----
  function updateCodeSyncStatus() {
    const el = document.getElementById("code-sync-status");
    if (!el) return;
    let text = "";
    if (codeDirty) text = "✏️ 入力中…（打ち終わると自動で確かめます）";
    else if (codeOwner === "code" && blocksMatchCode) text = "🤖 このコードで実行中（ブロックにも反映ずみ）";
    else if (codeOwner === "code") text = "🤖 このコードで実行中（ブロックでは表せない書き方があります）";
    el.textContent = text;
    el.hidden = !text;
  }

  // ブロックタブの上にかぶせる案内（VBAコードタブで書いたコードを実行しているとき）
  function updateCodeOwnerOverlay() {
    const ov = document.getElementById("code-owner-overlay");
    if (!ov) return;
    if (codeOwner !== "code") {
      ov.hidden = true;
      return;
    }
    const text = document.getElementById("code-owner-text");
    const toBlocks = document.getElementById("code-owner-to-blocks");
    if (blocksMatchCode) {
      text.textContent =
        "ブロックは、VBAコードと同じ内容になっています。ブロックで続けると、VBAコードはブロックから作り直されます（Dim の書き方などが変わることがあります）。";
      toBlocks.textContent = "🧩 ブロックで続ける";
      toBlocks.classList.remove("danger");
    } else {
      text.textContent =
        "VBAコードに、ブロックでは表せない書き方があるため、ブロックは前のままです" +
        (blocksMismatchReason ? `（${blocksMismatchReason}）` : "") +
        "。いまは VBAコードのまま実行しています。";
      toBlocks.textContent = "🧩 ブロックから作り直す（VBAコードの変更は消えます）";
      toBlocks.classList.add("danger");
    }
    ov.hidden = false;
  }

  // ----- コード → ブロック 変換 -----
  // タブ切替時 / エディタからのフォーカス離脱時に呼ばれる。
  // 書いたコードは書き換えない（実行はコードのまま）。ブロックで表せる形なら、ブロックにも並べる
  function applyCodeToBlocks() {
    if (!codeEditor || !workspace) return;
    if (!codeDirty) return;
    codeDirty = false;
    const text = codeEditor.getValue();
    hideCodeError();

    let result = null;
    try {
      result = window.VbaParser.parseVBA(text);
    } catch (e) {
      blocksMatchCode = false;
      blocksMismatchReason = e.line ? `${e.line}行目: ${e.message}` : e.message;
    }
    if (result) {
      // この変化はコード→ブロック変換によるもの。onWorkspaceChange でコードを作り直さないよう目印を付ける
      Blockly.Events.setGroup(LOAD_GROUP);
      try {
        workspace.clear();
        if (result.roots.length) {
          result.roots.forEach((r) => window.VbaParser.resolveVariables(r, workspace));
          const json = window.VbaParser.toWorkspaceJson(result.roots);
          Blockly.serialization.workspaces.load(json, workspace);
          workspace.cleanUp();
        }
        blocksMatchCode = true;
        blocksMismatchReason = "";
      } catch (e) {
        console.warn("ブロック構築に失敗:", e);
        blocksMatchCode = false;
        blocksMismatchReason = "ブロックを組み立てられませんでした";
      } finally {
        Blockly.Events.setGroup(false);
      }
    }
    lineMap = {};
    saveCodeState();
    updateCodeSyncStatus();
    updateCodeOwnerOverlay();
  }

  // 書いたコードはいつも保存しているので、課題の切り替えなどで消えることはない
  function confirmDiscardDirtyCode() {
    if (codeDirty) applyCodeToBlocks();
    clearTimeout(codeSaveTimer);
    saveCodeState();
    return true;
  }

  // ブロックからコードを作り直す前の確認（VBAコードタブで書いたコードが消えるとき）
  function confirmReplaceCode() {
    if (codeOwner !== "code") return true;
    return confirm("VBAコードタブで書いたプログラムは、ブロックから作ったコードに置きかわります。続けますか？");
  }

  // ブロックを正にする（コードはブロックから作り直す）
  function switchToBlocksOwner() {
    codeOwner = "blocks";
    codeDirty = false;
    blocksMatchCode = true;
    blocksMismatchReason = "";
    saveCodeState();
    updateCodeSyncStatus();
    updateCodeOwnerOverlay();
    onWorkspaceChange();
  }

  function hideCodeError() {
    const banner = document.getElementById("code-error-banner");
    if (banner) banner.hidden = true;
  }

  // ===== プログラムのまちがいの表示（やさしい説明つき） =====
  let shownError = null;
  function errorTitle(err, preview) {
    const at = err.line ? `${err.line}行目` : "";
    if (err.kind === "zenkaku") return `⚠️ 全角の文字があります（${at}）`;
    if (err.kind === "syntax" || err.kind === "compile") return `⚠️ 書き方のまちがい（${at || "コンパイル エラー"}）`;
    if (err.kind === "limit") return `⚠️ 処理が止まりませんでした${at ? `（${at}）` : ""}`;
    return preview ? `⚠️ このまま実行すると ${at} でエラーになります` : `⚠️ 実行時エラー（${at}）`;
  }

  function showProgramError(err, preview) {
    shownError = err;
    const panel = document.getElementById("program-error");
    if (!panel) return;
    document.getElementById("program-error-title").textContent = errorTitle(err, preview);
    document.getElementById("program-error-friendly").textContent = err.friendly || err.message || "";
    const hint = document.getElementById("program-error-hint");
    hint.textContent = err.hint ? `💡 ${err.hint}` : "";
    hint.hidden = !err.hint;
    const vba = document.getElementById("program-error-vba");
    const vbaText = err.kind === "zenkaku" ? "" : String(err.message || "").replace(/\n/g, " ");
    vba.textContent = vbaText ? `Excel の表示: ${vbaText}` : "";
    vba.hidden = !vbaText;
    document.getElementById("program-error-goto").hidden = !err.line;
    document.getElementById("program-error-fix").hidden = err.kind !== "zenkaku";
    panel.classList.toggle("preview", !!preview && err.kind !== "zenkaku" && err.kind !== "syntax" && err.kind !== "compile");
    panel.hidden = false;
    markErrorInCode(err);
  }

  function hideProgramError() {
    shownError = null;
    const panel = document.getElementById("program-error");
    if (panel) panel.hidden = true;
    clearErrorMarks();
  }

  function clearErrorMarks() {
    if (!codeEditor) return;
    if (errorMarkLine !== null) codeEditor.removeLineClass(errorMarkLine, "background", "code-error-line");
    errorMarkLine = null;
    zenkakuMarks.forEach((m) => m.clear());
    zenkakuMarks = [];
  }

  function markErrorInCode(err) {
    clearErrorMarks();
    if (!codeEditor) return;
    if (err.kind === "zenkaku" && err.zenkaku) {
      err.zenkaku.forEach((z) => {
        zenkakuMarks.push(codeEditor.markText({ line: z.line - 1, ch: z.col - 1 }, { line: z.line - 1, ch: z.col }, { className: "cm-zenkaku", title: `全角の「${z.ch}」→ 半角の「${z.fix}」に` }));
      });
    }
    if (err.line && err.line <= codeEditor.lineCount()) {
      errorMarkLine = err.line - 1;
      codeEditor.addLineClass(errorMarkLine, "background", "code-error-line");
    }
  }

  // ===== 実行（src/vba/session.js） =====
  // 実行するプログラム：いつも VBAコードタブの中身（ブロックのときはブロックから作ったコードが入っている）
  function currentSource() {
    return codeEditor ? codeEditor.getValue() : generateVBA(workspace);
  }

  function createSession() {
    try {
      const form = formPanel ? formPanel.runtimeForm() : null;
      return { s: new VbaSession.Session(currentSource(), { initialCells: view.getInitialCells(), form, seed: 20261022 }) };
    } catch (e) {
      if (e && e.kind) return { error: e };
      throw e;
    }
  }

  // フォームで動かすプログラムか（ボタンの _Click や UserForm_Initialize がある）
  function formModeAvailable(s) {
    const names = formPanel ? formPanel.controlNames().map((c) => c.name.toLowerCase()) : [];
    return s.procNames().some((p) => {
      const k = p.toLowerCase();
      return k === "userform_initialize" || names.some((n) => k.startsWith(n + "_"));
    });
  }

  // ▶実行で動かすものの一覧を作り直す（フォームを開く / 引数の無い Sub）
  function updateRunTargets(s) {
    const sel = document.getElementById("run-target");
    const opts = [];
    if (s && formModeAvailable(s)) opts.push(["form", "🪟 フォームを開く"]);
    if (s) s.macroNames().forEach((m) => opts.push(["macro:" + m, `📜 Sub ${m}`]));
    const values = opts.map((o) => o[0]);
    if (!values.includes(runTarget)) runTarget = values[0] || null;
    if (sel) {
      sel.innerHTML = opts.map(([v, l]) => `<option value="${escapeHtml(v)}">${escapeHtml(l)}</option>`).join("");
      if (runTarget) sel.value = runTarget;
      sel.hidden = opts.length <= 1;
    }
  }

  // 選んだもので実行を始める（フォームなら Initialize まで、Sub なら最初の答え待ちまで）
  function startTarget(s) {
    if (runTarget === "form") return s.showForm();
    if (runTarget && runTarget.startsWith("macro:")) return s.start(runTarget.slice(6));
    return null;
  }

  function schedulePreview(ms) {
    clearTimeout(previewTimer);
    previewTimer = setTimeout(runPreview, ms === undefined ? 150 : ms);
  }

  // 編集中のプレビュー：MsgBox はそのまま OK、InputBox は前に答えた値があれば使い、無ければそこで止まる
  function runPreview() {
    clearTimeout(previewTimer);
    if (!view) return;
    if (runActive && (view.playing || (window.VbaDialog && VbaDialog.isOpen()))) {
      showToast("プログラムが変わったので、実行を止めました");
    }
    runActive = false;
    view.stop();
    const { s, error } = createSession();
    if (error) {
      session = null;
      view.loadSession(null, "⚠️ プログラムにまちがいがあります（下の説明を見てね）");
      showProgramError(error, true);
      setFlowSource(null, "プログラムにまちがいがあるので、流れ図を描けません（下の説明を見て直そう）");
      checkQuestClear();
      return;
    }
    session = s;
    setFlowSource(s.ast);
    view.liveForm = false;
    updateRunTargets(s);
    let status = null;
    let res = startTarget(s);
    if (res) {
      res = autoAnswer(res);
      if (res.status === "waiting") status = "⌨️ InputBox の答え待ちのところまで準備しました（▶ 実行すると質問が出ます）";
      else if (runTarget === "form") status = "🪟 ▶ 実行 を押すとフォームが開きます。フォームのボタンを押して動かそう";
    } else if (s.procNames().length) {
      status = "ℹ️ ▶ 実行で動かせるものがありません（ボタンの命令は、フォームにそのボタンを置くと動かせます）";
    }
    view.loadSession(s, status);
    if (s.status === "error") showProgramError(s.error, true);
    else hideProgramError();
    checkQuestClear();
    updateFormPanel();
    updateStepGuide();
  }

  function autoAnswer(res) {
    while (res.status === "waiting") {
      const q = res.request;
      if (q.type === "msgbox") res = session.resume(undefined);
      else if (Object.prototype.hasOwnProperty.call(inputCache, q.key)) res = session.resume(inputCache[q.key]);
      else break;
    }
    return res;
  }

  // ▶実行・1コマ：新しく実行を始める（InputBox は毎回たずね直す）
  function startRun(play) {
    Object.keys(inputCache).forEach((k) => delete inputCache[k]);
    const { s, error } = createSession();
    if (error) {
      showProgramError(error, false);
      return false;
    }
    updateRunTargets(s);
    if (!runTarget) {
      document.getElementById("step-status").textContent = "ℹ️ ▶ 実行で動かせるものがありません（ボタンの命令は、フォームにそのボタンを置くと動かせます）";
      return false;
    }
    session = s;
    runActive = true;
    hideProgramError();
    setFlowSource(s.ast);
    startTarget(s);
    view.liveForm = runTarget === "form";
    view.loadSession(s, "▶ 実行中…");
    if (runTarget === "form") switchRightTab("form");
    if (play) view.play(true);
    updateFormPanel();
    updateStepGuide();
    return true;
  }

  // ===== ユーザーフォーム =====
  // フォームで実行中か
  function formRunning() {
    return !!session && runActive && runTarget === "form";
  }

  // フォームの表示を、いまの実行の状態に合わせる
  function updateFormPanel() {
    if (!formPanel) return;
    const toolbar = document.getElementById("form-toolbar");
    const runBar = document.getElementById("form-run-bar");
    if (!formRunning()) {
      if (formPanel.mode !== "design") formPanel.setDesign();
      toolbar.hidden = false;
      runBar.hidden = true;
      return;
    }
    const m = view.currentModel();
    const atLatest = view.cursor >= view.steps.length;
    const dialogOpen = window.VbaDialog && VbaDialog.isOpen();
    const interactive = atLatest && !view.playing && !dialogOpen && session.status === "idle";
    let info = "";
    if (session.status === "ended") info = "フォームは閉じられました。▶ 実行 でもう一度開けます";
    else if (interactive) info = "🖱 ボタンを押してみよう（テキストボックスには文字を入力できます）";
    else if (view.playing || dialogOpen) info = "⏳ 命令が動いています…";
    else if (!atLatest) info = "⏱ 前の時点を見ています（⏭ 結果 で最新にもどると、また操作できます）";
    formPanel.setRun(m.controls || {}, m.form || {}, interactive, info);
    toolbar.hidden = true;
    runBar.hidden = false;
    document.getElementById("form-run-text").textContent = session.status === "ended" ? "■ 終わりました" : "▶ 実行中";
  }

  // 実行中のフォームでボタンが押された
  function onFormButton(name, texts) {
    if (!formRunning() || session.busy() || view.playing) return;
    Object.keys(texts).forEach((n) => session.setText(n, texts[n]));
    const res = session.click(name);
    view.stepsAppended();
    if (res.noHandler) {
      showToast(`「${name}」を押したときの命令（${name}_Click）がまだありません。🧩 ブロックの「🖱 ${name} をクリックしたとき」を使おう`, 5000);
      view.goToEnd();
      updateFormPanel();
      return;
    }
    view.play(false); // 押したところから続きを再生する
  }

  // 実行中のフォームの × が押された
  function onFormClose() {
    if (!formRunning()) return;
    session.closeForm();
    view.stepsAppended();
    view.goToEnd();
    updateFormPanel();
  }

  // ■ 止める：実行をやめてデザインにもどる
  function stopRun() {
    runActive = false;
    view.liveForm = false;
    runPreview();
  }

  // フォームのデザインが変わった
  function onFormChanged(form, info) {
    if (!sharedViewMode) {
      savedForms[currentKey()] = form;
      persist();
    }
    if (info.renamed) renameControlEverywhere(info.renamed.from, info.renamed.to);
    if (runActive) stopRun();
    else schedulePreview(150);
    updateStepGuide();
  }

  // 部品の名前を変えたら、ブロックとコードの中の名前も変える
  function renameControlEverywhere(from, to) {
    let n = 0;
    workspace.getAllBlocks(false).forEach((b) => {
      if (b.getField("CONTROL") && b.getFieldValue("CONTROL") === from) {
        b.setFieldValue(to, "CONTROL");
        n++;
      }
    });
    if (codeOwner === "code" && codeEditor) {
      const re = new RegExp(`\\b${from}(?=[._])`, "g");
      const text = codeEditor.getValue();
      const next = text.replace(re, to);
      if (next !== text) {
        setEditorCode(next);
        saveCodeState();
        n++;
      }
    }
    if (n) showToast(`プログラムの中の「${from}」も「${to}」に変えました`);
  }

  // 部品をダブルクリック：その部品を押したときの命令を書く場所を用意する
  function onEditHandler(name) {
    const ctl = formPanel.controlNames().find((c) => c.name === name);
    if (!ctl) return;
    if (ctl.type !== "CommandButton") {
      showToast("ダブルクリックで命令を書けるのは、ボタン（CommandButton）です");
      return;
    }
    if (codeOwner === "code") {
      const text = codeEditor.getValue();
      const re = new RegExp(`Sub\\s+${name}_Click\\s*\\(`, "i");
      if (!re.test(text)) {
        const add = `${text.trim() ? "\n\n" : ""}Private Sub ${name}_Click()\n    \nEnd Sub`;
        codeEditor.replaceRange(add, { line: codeEditor.lastLine() });
      }
      document.getElementById("tab-code").click();
      setTimeout(() => {
        const lines = codeEditor.getValue().split("\n");
        const idx = lines.findIndex((l) => re.test(l));
        codeEditor.focus();
        if (idx >= 0) codeEditor.setCursor({ line: idx + 1, ch: (lines[idx + 1] || "").length });
      }, 80);
      return;
    }
    let block = workspace.getTopBlocks(false).find((b) => b.type === "form_button_click" && b.getFieldValue("CONTROL") === name);
    if (!block) {
      block = workspace.newBlock("form_button_click");
      block.setFieldValue(name, "CONTROL");
      block.initSvg();
      block.render();
      const others = workspace.getTopBlocks(false).filter((b) => b !== block);
      const bottom = others.reduce((y, b) => Math.max(y, b.getBoundingRectangle().bottom), 0);
      block.moveBy(40, bottom + 40);
    }
    document.getElementById("tab-blocks").click();
    setTimeout(() => {
      workspace.centerOnBlock(block.id);
      block.select();
    }, 80);
    showToast(`「🖱 ${name} をクリックしたとき」の中に、命令のブロックを入れよう`);
  }

  // ① フォームを作る → ② 命令を組む → ③ ▶ 実行 の、いまの場所を光らせる
  function updateStepGuide() {
    const chips = document.querySelectorAll(".step-chip");
    if (!chips.length) return;
    let cur = lastArea === "form" ? 1 : 2;
    if (runActive) cur = 3;
    chips.forEach((c) => c.classList.toggle("current", Number(c.dataset.step) === cur));
  }

  // InputBox / MsgBox の答えをダイアログで聞く
  function askRequest(req) {
    if (req.type === "inputbox") {
      return VbaDialog.inputBox({ prompt: req.prompt, title: req.title, default: req.default }).then((ans) => {
        inputCache[req.key] = ans;
        return ans;
      });
    }
    return VbaDialog.msgBox({ prompt: req.prompt, title: req.title, buttons: req.buttons });
  }

  // 再生で記録の最後まで来た：答え待ちならダイアログを出し、答えのあと続きを再生する
  function onReachEnd(resume) {
    if (!session || session.status !== "waiting") {
      finishRun();
      return false;
    }
    askRequest(session.request).then((ans) => {
      session.resume(ans);
      resume();
    });
    return true;
  }

  // 再生中・1コマで MsgBox の step に来た：ダイアログを出し、OK のあと続ける
  function onMsgStep(step, done) {
    VbaDialog.msgBox({ prompt: step.key, title: step.title, buttons: step.buttons }).then((result) => {
      const live = session && session.status === "waiting" && session.request && session.request.type === "msgbox" && view.cursor === session.steps.length;
      if (live) {
        session.resume(result);
        view.stepsAppended();
      }
      done();
    });
  }

  // 1コマ進める。最初なら実行を始める。記録の最後で答え待ちなら、質問を出してから進める
  function stepOnce() {
    if (window.VbaDialog && VbaDialog.isOpen()) return;
    view.stop();
    if (!runActive && view.cursor === 0) {
      if (!startRun(false)) return;
    }
    runActive = true;
    const atEnd = () => view.cursor >= view.steps.length;
    if (!atEnd()) {
      view.stepForward(true, () => {});
      if (atEnd() && session && session.status !== "waiting") finishRun();
      return;
    }
    if (session && session.status === "waiting") {
      askRequest(session.request).then((ans) => {
        session.resume(ans);
        view.stepsAppended();
        if (!atEnd()) view.stepForward(true, () => {});
        if (atEnd() && session.status !== "waiting") finishRun();
      });
    }
  }

  // 実行が最後まで終わったとき（フォームのときは、ボタンを押すのを待つ状態）
  function finishRun() {
    if (!session) return;
    if (session.status === "error") showProgramError(session.error, false);
    else if (formRunning() && session.status === "idle") {
      document.getElementById("step-status").textContent = "🖱 フォームのボタンを押してみよう（全 " + session.steps.length + " ステップ）";
    } else if (formRunning() && session.status === "ended") {
      document.getElementById("step-status").textContent = "■ フォームが閉じて、プログラムが終わりました";
    }
    checkQuestClear();
    updateFormPanel();
  }

  // 実行中の行・ブロックを光らせる
  function onStepApplied(step) {
    if (codeEditor) {
      if (execLine !== null) codeEditor.removeLineClass(execLine, "background", "code-exec-line");
      execLine = null;
      if (step && step.line && step.line <= codeEditor.lineCount()) {
        execLine = step.line - 1;
        codeEditor.addLineClass(execLine, "background", "code-exec-line");
        if (document.getElementById("tab-code-panel").classList.contains("active")) {
          codeEditor.scrollIntoView({ line: execLine, ch: 0 }, 80);
        }
      }
    }
    if (workspace) {
      const id = step && step.line && codeOwner === "blocks" ? lineMap[step.line] : null;
      try {
        workspace.highlightBlock(id && workspace.getBlockById(id) ? id : null);
      } catch (e) {
        /* ハイライトできなくても実行には影響しない */
      }
    }
    if (formRunning()) updateFormPanel();
    if (flowChart) flowChart.highlight(step ? step.line : null, step ? step.proc : null, step ? step.branch : undefined);
  }

  // ===== 🔀 流れ図 =====
  function setFlowSource(ast, msg) {
    flowAst = ast;
    flowMsg = msg || null;
    flowDirty = true;
    if (document.getElementById("right-panel-flow").classList.contains("active")) renderFlow();
  }
  function renderFlow() {
    if (!flowChart || !flowDirty) return;
    flowDirty = false;
    try {
      flowChart.render(flowAst, flowMsg);
    } catch (e) {
      console.warn("流れ図を描けませんでした:", e);
      flowChart.render(null, "この流れ図はまだ描けません");
    }
  }
  // 流れ図の記号にマウスを乗せた：その命令の行を光らせる
  function onFlowHover(line) {
    if (!codeEditor) return;
    if (hoverLine !== null) codeEditor.removeLineClass(hoverLine, "background", "code-hover-line");
    hoverLine = null;
    if (line && line <= codeEditor.lineCount()) {
      hoverLine = line - 1;
      codeEditor.addLineClass(hoverLine, "background", "code-hover-line");
    }
  }
  // 流れ図の記号をクリックした：その命令の場所（ブロックかコードの行）へ飛ぶ
  function onFlowClick(line) {
    if (!line) return;
    const id = codeOwner === "blocks" ? lineMap[line] : null;
    if (id && workspace.getBlockById(id)) {
      document.getElementById("tab-blocks").click();
      setTimeout(() => {
        workspace.centerOnBlock(id);
        workspace.getBlockById(id).select();
      }, 60);
      return;
    }
    document.getElementById("tab-code").click();
    setTimeout(() => {
      codeEditor.refresh();
      codeEditor.focus();
      codeEditor.setCursor({ line: line - 1, ch: 0 });
      codeEditor.scrollIntoView({ line: line - 1, ch: 0 }, 80);
    }, 80);
  }

  // ----- 初期化 -----
  window.addEventListener("load", () => {
    initTheme();
    loadPersisted();
    Blockly.setLocale(Blockly.Msg);

    const isDark = !document.body.classList.contains("light-mode");
    const themes = getThemes();

    workspace = Blockly.inject("blockly-area", {
      toolbox: TOOLBOX,
      grid: { spacing: 24, length: 3, colour: isDark ? "#333" : "#ddd", snap: true },
      zoom: { controls: true, wheel: true, startScale: 1.0, maxScale: 2.5, minScale: 0.4 },
      trashcan: true,
      move: { scrollbars: true, drag: true, wheel: true },
      theme: isDark ? themes.dark : themes.light,
      renderer: "zelos",
      media: "vendor/blockly/media/",
    });

    // 📦 変数カテゴリ：日本語ブロック + 「変数を作る」ボタン + 登録済み変数の取り出し
    workspace.registerToolboxCategoryCallback("VARIABLE_JP", function (ws) {
      const items = [
        { kind: "button", text: "＋ 変数を作る", callbackKey: "CREATE_VARIABLE_JP" },
        { kind: "block", type: "var_dim" },
        { kind: "block", type: "var_set" },
        { kind: "block", type: "var_change", inputs: { DELTA: { shadow: { type: "value_number", fields: { NUM: 1 } } } } },
      ];
      const vars = ws.getVariableMap ? ws.getVariableMap().getAllVariables() : ws.getAllVariables();
      vars.forEach((v) => {
        items.push({ kind: "block", type: "var_get", fields: { VAR: { name: v.getName ? v.getName() : v.name, id: v.getId ? v.getId() : v.id_ } } });
      });
      if (!vars.length) {
        items.push({ kind: "block", type: "var_get" });
      }
      return items;
    });
    // 「＋ 変数を作る」ボタン。Blockly標準の createVariableButtonHandler は
    // ひらがな/予約名チェックを経由しないため、自前でプロンプト＋検証する
    workspace.registerButtonCallback("CREATE_VARIABLE_JP", function (button) {
      promptNewVariableName(button.getTargetWorkspace());
    });

    view = new ExcelView({
      table: document.getElementById("excel-table"),
      ref: document.getElementById("cell-ref"),
      formula: document.getElementById("formula-bar"),
      status: document.getElementById("step-status"),
      array: document.getElementById("array-viz"),
      vars: document.getElementById("var-viz"),
      msgbox: document.getElementById("msgbox-overlay"),
      tabs: document.getElementById("excel-tabs"),
      onEdit: onCellEdited, // 仮想Excelのセル編集時
      onPlayState: (playing) => {
        // 実行中は ▶ → ⏸ にトグル
        const btn = document.getElementById("run-btn");
        btn.textContent = playing ? "⏸ 一時停止" : "▶ 実行";
        updateFormPanel();
        updateStepGuide();
      },
      onStepChange: updateTimeline, // タイムラインのスライダー・カウンタを同期
      onStepApplied, // 実行中の行・ブロックを光らせる
      onMsgStep, // 再生中・1コマで MsgBox に来たらダイアログを出す
      onReachEnd, // 記録の最後まで来たら、InputBox などの答えを聞いて続ける
    });

    // 🪟 ユーザーフォーム（デザイナー兼実行画面）
    formPanel = new UserFormPanel({
      stage: document.getElementById("form-stage"),
      props: document.getElementById("form-props"),
      onChange: onFormChanged,
      onEditHandler,
      onRunClick: onFormButton,
      onRunClose: onFormClose,
    });
    // ブロックの「部品の名前」の選択肢は、いまのフォームの部品から作る
    window.UMB_formControls = () => (formPanel ? formPanel.controlNames() : []);

    // 🔀 流れ図
    flowChart = new FlowChart(document.getElementById("flow-body"), { onHover: onFlowHover, onClick: onFlowClick });

    // ブロック変更 → コード再生成 + 実行し直し
    workspace.addChangeListener(onWorkspaceChange);

    // ----- VBAコードエディタ（CodeMirror）初期化 -----
    codeEditor = CodeMirror(document.getElementById("code-editor"), {
      mode: "vbscript",
      lineNumbers: true,
      indentUnit: 4,
      tabSize: 4,
      indentWithTabs: false,
      lineWrapping: false,
      extraKeys: {
        Tab: handleCodeTabKey,
        Enter: handleCodeEnterKey,
        // show-hint アドオン内蔵の Esc ハンドラ（Esc: n.close）が、この環境では
        // なぜか CodeMirror.lookupKey に見つけてもらえず動作しなかったため、
        // extraKeys 側で明示的に閉じる処理を持たせて確実に効くようにする
        Esc: function (cm) {
          if (cm.state.completionActive) {
            cm.state.completionActive.close();
            return;
          }
          return CodeMirror.Pass;
        },
      },
    });
    codeEditor.on("change", (cm, changeObj) => {
      if (suppressCodeChange) return;
      // 生徒がコードを書いた：これからはこのコードを実行する（ブロックからは作り直さない）
      codeDirty = true;
      codeOwner = "code";
      updateCodeSyncStatus();
      hideCodeError();
      clearErrorMarks();
      schedulePreview(700); // 打ち終わって少ししたら、動きを確かめ直す
      clearTimeout(codeSaveTimer);
      codeSaveTimer = setTimeout(saveCodeState, 800);
      if (changeObj.origin !== "setValue") {
        const lineNo = changeObj.to.line;
        // 同じ行に対して前回分の保留中デデントがあれば上書き（二重発火でのインデント
        // 過剰減少を防ぐ。Enterキー側の同期デデントとも dedentTimer 経由で連携する）
        if (dedentTimer) clearTimeout(dedentTimer);
        dedentTimerLine = lineNo;
        dedentTimer = setTimeout(() => {
          dedentClosingKeywordLine(cm, dedentTimerLine);
          dedentTimer = null;
        }, 0);
      }
    });
    codeEditor.on("blur", () => {
      if (codeDirty) applyCodeToBlocks();
    });

    // 入力アシスト: For/Dim等のキーワードを打つと自動で例文候補をポップアップ表示する。
    // 前方一致する候補が無ければ何も出さない（hintヘルパーがnullを返し showHint は無視する）
    if (typeof CodeMirror !== "undefined" && codeEditor.showHint) {
      codeEditor.on("inputRead", (cm, change) => {
        if (change.origin !== "+input") return;
        if (!/[A-Za-z]/.test(change.text[0] || "")) return;
        if (cm.state.completionActive) return; // 既にポップアップ表示中なら出し直さない
        const cur = currentWordBeforeCursor(cm);
        if (!cur) return;
        const wordLower = cur.word.toLowerCase();
        if (!Object.keys(CODE_SNIPPETS).some((k) => k.startsWith(wordLower))) return;
        cm.showHint({ hint: CodeMirror.hint.vbaSnippet, completeSingle: false, alignWithWord: true });
      });
    }

    bindControls();
    loadFreeMode(); // 起動時はクエスト未選択＝フリーモード
    const fromShare = loadFromShareUrl(); // 共有リンクがあれば作品を表示
    // 初めて開いたときは、はじめてガイドを出す（共有リンクで開いた先生には出さない）
    if (!fromShare && window.UmbTour) setTimeout(() => UmbTour.start(false), 600);

    // 他タブでの更新を取り込む（複数タブを開いたとき、後から persist した方が
    // 一方的に上書きしてしまう事故を防ぐ）
    window.addEventListener("storage", onOtherTabStorageChange);
  });

  // ----- 他タブでの localStorage 更新を取り込む -----
  // storage イベントは「変更したタブ以外」でのみ発火する。
  // クリア済みフラグは合算（減らさない）、ブロック/初期データは
  // 「今このタブで編集中でない課題」だけ取り込む（作業中のものを壊さないため）
  function onOtherTabStorageChange(e) {
    if (![LS_BLOCKS, LS_SOLVED, LS_INITIAL, LS_CODE, LS_FORMS].includes(e.key)) return;
    try {
      const otherSolved = JSON.parse(localStorage.getItem(LS_SOLVED) || "[]");
      let newlySolvedCount = 0;
      otherSolved.forEach((id) => {
        if (!solved.has(id)) {
          solved.add(id);
          newlySolvedCount++;
        }
      });

      const otherBlocks = JSON.parse(localStorage.getItem(LS_BLOCKS) || "{}");
      const otherInitial = JSON.parse(localStorage.getItem(LS_INITIAL) || "{}");
      const activeKey = currentTaskId === null ? FREE_MODE_ID : currentTaskId;
      // 実際に内容が変わったものだけ数える（毎回のstorageイベントで
      // 無関係なキーまで「更新した」扱いにしてトーストが乱発しないように）
      let updatedTaskCount = 0;
      Object.keys(otherBlocks).forEach((k) => {
        if (k === activeKey) return;
        if (savedBlocks[k] !== otherBlocks[k]) {
          savedBlocks[k] = otherBlocks[k];
          updatedTaskCount++;
        }
      });
      Object.keys(otherInitial).forEach((k) => {
        if (k !== activeKey) savedInitial[k] = otherInitial[k];
      });
      const otherCode = JSON.parse(localStorage.getItem(LS_CODE) || "{}");
      Object.keys(savedCode).forEach((k) => {
        if (k !== activeKey && !otherCode[k]) delete savedCode[k];
      });
      Object.keys(otherCode).forEach((k) => {
        if (k !== activeKey) savedCode[k] = otherCode[k];
      });
      const otherForms = JSON.parse(localStorage.getItem(LS_FORMS) || "{}");
      Object.keys(otherForms).forEach((k) => {
        if (k !== activeKey) savedForms[k] = otherForms[k];
      });

      if (newlySolvedCount > 0 || updatedTaskCount > 0) {
        updateCurrentQuestDisplay();
        // クエスト一覧モーダルを開いたままだった場合、✓表示を最新化する
        const questModal = document.getElementById("quest-modal");
        if (questModal && !questModal.hidden) buildQuestModal();
        const msg =
          newlySolvedCount > 0
            ? `🔄 別のタブでの更新を取り込みました（クエストクリア +${newlySolvedCount}）`
            : "🔄 別のタブでの更新を取り込みました";
        showToast(msg);
      }
    } catch (err) {
      console.warn("他タブの更新の取り込みに失敗:", err);
    }
  }

  // ----- クエスト一覧モーダル描画 -----
  function buildQuestModal() {
    const list = document.getElementById("quest-list");
    list.innerHTML = "";
    TASKS.forEach((quest) => {
      const item = document.createElement("div");
      item.className = "quest-item";
      if (solved.has(quest.id)) item.classList.add("solved");
      item.innerHTML = `
        <div class="quest-item-left">
          <div class="quest-item-title">${quest.title}</div>
          <div class="quest-item-goal">${quest.goal.split("\n")[0]}</div>
        </div>
        <div class="quest-item-stars">${"★".repeat(quest.difficulty)}</div>
      `;
      item.addEventListener("click", () => {
        // loadTask が false を返す＝コード未反映を理由に中止された場合は
        // モーダルを閉じない（生徒がコードタブに戻って直せるように）
        if (loadTask(quest.id)) {
          document.getElementById("quest-modal").hidden = true;
        }
      });
      list.appendChild(item);
    });
  }

  function updateCurrentQuestDisplay() {
    const task = TASKS.find((t) => t.id === currentTaskId);
    const freeModeBtn = document.getElementById("free-mode-btn");
    if (task) {
      const check = solved.has(task.id) ? "✓ " : "";
      document.getElementById("current-quest").textContent =
        `${check}📍 ${task.title} ${"★".repeat(task.difficulty)}`;
      if (freeModeBtn) freeModeBtn.hidden = false;
    } else {
      document.getElementById("current-quest").textContent = "🎨 フリーモード";
      if (freeModeBtn) freeModeBtn.hidden = true;
    }
  }

  // ----- 課題読み込み -----
  // 戻り値: true = 読み込みを実行した / false = 中止した（呼び出し元はUIを変更しないこと）
  function loadTask(taskId) {
    // コードタブに未反映の変更があれば、先に確認する（無ければ即 true）
    if (!confirmDiscardDirtyCode()) return false;
    // 現在のブロックを保存（共有閲覧中は保存されない）
    if (hasLoadedOnce && workspace) {
      saveCurrentBlocks();
    }
    hasLoadedOnce = true;
    sharedViewMode = false; // 課題を選んだら通常モードに復帰
    currentTaskId = taskId;
    const task = TASKS.find((t) => t.id === taskId);
    if (!task) return false;

    // 課題説明
    document.getElementById("task-title").textContent = task.title;
    document.getElementById("task-difficulty").textContent = "★".repeat(task.difficulty) +
      "☆".repeat(5 - task.difficulty);
    document.getElementById("task-goal").textContent = task.goal;
    renderGoalPreview(task);

    // ヒントUIリセット
    resetHints(task);

    // 仮想Excelの初期データを復元（課題ごと）
    view.setInitialCells(savedInitial[taskId] || {});

    // ブロック復元（無ければ真っ白）。復元中は保存抑制（B7）
    // ※ codeDirty はここに来る前に confirmDiscardDirtyCode() で確実に false 済み
    suppressSave = true;
    workspace.clear();
    const xml = savedBlocks[taskId];
    if (xml) {
      try {
        const dom = Blockly.utils.xml.textToDom(xml);
        Blockly.Xml.domToWorkspace(dom, workspace);
      } catch (e) {
        console.warn("ブロック復元失敗:", e);
      }
    }
    suppressSave = false;
    restoreCodeState(taskId);
    runActive = false;
    formPanel.setForm(savedForms[taskId] || (task.form ? task.form : UserFormUtil.defaultForm()));

    updateCurrentQuestDisplay();
    onWorkspaceChange();
    applyQuestCollapsed(false); // 課題の内容が見えた状態で始める
    return true;
  }

  // ----- フリーモード読み込み（クエスト未選択時の初期状態）-----
  // 戻り値: true = 読み込みを実行した / false = 中止した
  function loadFreeMode() {
    if (!confirmDiscardDirtyCode()) return false;
    if (hasLoadedOnce && workspace) {
      saveCurrentBlocks();
    }
    hasLoadedOnce = true;
    sharedViewMode = false;
    currentTaskId = null;

    document.getElementById("task-title").textContent = "🎨 フリーモード";
    document.getElementById("task-difficulty").textContent = "";
    document.getElementById("task-goal").textContent =
      "クエストを選ばなくても、自由にブロックを組み立てて試せます。\n上の「🎮 クエスト選択」から課題に挑戦することもできます。";
    renderGoalPreview({ goalPreview: null });

    resetHints(null);

    view.setInitialCells(savedInitial[FREE_MODE_ID] || {});

    // ※ codeDirty はここに来る前に confirmDiscardDirtyCode() で確実に false 済み
    suppressSave = true;
    workspace.clear();
    const xml = savedBlocks[FREE_MODE_ID];
    if (xml) {
      try {
        const dom = Blockly.utils.xml.textToDom(xml);
        Blockly.Xml.domToWorkspace(dom, workspace);
      } catch (e) {
        console.warn("ブロック復元失敗:", e);
      }
    }
    suppressSave = false;
    restoreCodeState(FREE_MODE_ID);
    runActive = false;
    formPanel.setForm(savedForms[FREE_MODE_ID] || UserFormUtil.defaultForm());
    // 初めて開いたとき：教科書の最初の形（ボタンを押したときの命令の入れ物）を置いておく
    if (!xml && !savedCode[FREE_MODE_ID] && workspace.getTopBlocks(false).length === 0) {
      const hat = workspace.newBlock("form_button_click");
      hat.setFieldValue("CommandButton1", "CONTROL");
      hat.initSvg();
      hat.render();
      hat.moveBy(40, 40);
    }

    updateCurrentQuestDisplay();
    onWorkspaceChange();
    return true;
  }

  // ----- 完成イメージのミニグリッド描画（F7）-----
  function renderGoalPreview(task) {
    const wrap = document.getElementById("goal-preview");
    const grid = document.getElementById("goal-preview-grid");
    if (!task.goalPreview || task.goalPreview.length === 0) {
      wrap.hidden = true;
      grid.innerHTML = "";
      return;
    }
    wrap.hidden = false;
    // 範囲を算出（最小 A1..C3）
    let maxRow = 3,
      maxCol = 3;
    const byAddr = {};
    task.goalPreview.forEach((g) => {
      byAddr[g.addr] = g;
      const m = g.addr.match(/^([A-Z]+)(\d+)$/);
      if (m) {
        maxRow = Math.max(maxRow, parseInt(m[2], 10));
        maxCol = Math.max(maxCol, m[1].charCodeAt(0) - 64);
      }
    });
    const colL = (n) => String.fromCharCode(64 + n);
    let html = "<tr><th></th>";
    for (let c = 1; c <= maxCol; c++) html += `<th>${colL(c)}</th>`;
    html += "</tr>";
    for (let r = 1; r <= maxRow; r++) {
      html += `<tr><td class='gp-rownum'>${r}</td>`;
      for (let c = 1; c <= maxCol; c++) {
        const g = byAddr[colL(c) + r];
        const bg = g && g.bg ? `background:${g.bg};color:#fff;` : "";
        const val = g && g.value !== undefined ? g.value : "";
        html += `<td style="${bg}">${val}</td>`;
      }
      html += "</tr>";
    }
    grid.innerHTML = html;
  }

  function saveCurrentBlocks() {
    if (suppressSave || sharedViewMode) return;
    try {
      const key = currentTaskId === null ? FREE_MODE_ID : currentTaskId;
      const dom = Blockly.Xml.workspaceToDom(workspace);
      savedBlocks[key] = Blockly.Xml.domToText(dom);
      persist();
    } catch (e) {
      console.warn("ブロック保存失敗:", e);
    }
  }

  // ----- ワークスペース変更時 -----
  function onWorkspaceChange(event) {
    if (event && event.isUiEvent) return;

    // コード→ブロック変換で並べたブロックの変化：コードは生徒が書いたまま（作り直さない）
    if (event && event.group === LOAD_GROUP) {
      saveCurrentBlocks();
      updateCanvasGuide();
      return;
    }

    // ブロックが正のときだけ、ブロックから VBA コードを作ってコードエディタに入れる
    // （VBAコードタブで書いたコードを実行しているときは、書いたコードを消さない）
    if (codeOwner === "blocks") {
      const g = generateVBAWithMap(workspace);
      lineMap = g.lineMap;
      setEditorCode(g.code);
    }

    // エラーチェック（未接続の値ブロックなど簡易検出）
    checkErrors();
    saveCurrentBlocks();

    // 実行し直す（ドラッグ中の細かい変化はまとめる）
    schedulePreview(120);

    updateBaselineLabel();
    updateCanvasGuide();
  }

  // シンタックスハイライトは CodeMirror（vbscript モード）が担当するため、
  // 旧・単一パス正規表現ハイライタ（highlightVBA）は撤去済み。

  // ----- クエストクリア自動判定 -----
  // ----- 「実際に実行可能なプログラムが組まれているか」の判定 -----
  // クエストクリア判定は、これが true のときだけ task.check() を呼ぶ。
  // セルへの直接入力だけ（ブロックを一切組んでいない）でも model の見た目は
  // 正解と同じになりうるため、「ブロックを実行した結果としての model か」を
  // 先に保証する。新しい判定条件（例: 特定ブロックの使用を必須にする等）を
  // 追加したくなったら、この関数に足せば全課題に一括で効く。
  function hasExecutableProgram() {
    return !!session && session.status !== "error" && session.steps.some((st) => STATE_SCOPES.has(st.scope) && !st.declare);
  }

  function checkQuestClear() {
    const task = TASKS.find((t) => t.id === currentTaskId);
    const banner = document.getElementById("clear-banner");
    const nextBtn = document.getElementById("next-quest-btn");
    if (!task || typeof task.check !== "function") {
      banner.classList.remove("show");
      if (nextBtn) nextBtn.hidden = true;
      return;
    }
    let passed = false;
    if (hasExecutableProgram()) {
      try {
        passed = task.check(view.finalModel());
      } catch (e) {
        passed = false;
      }
    }
    if (passed) {
      const firstTime = !solved.has(currentTaskId);
      if (firstTime) {
        solved.add(currentTaskId);
        persist();
        celebrate(); // 🎉 演出（F5）
        updateCurrentQuestDisplay();
      }
      banner.classList.add("show");
      // 次のクエストへボタン（F4）
      if (nextBtn) {
        const idx = TASKS.findIndex((t) => t.id === currentTaskId);
        nextBtn.hidden = idx < 0 || idx >= TASKS.length - 1;
      }
    } else {
      banner.classList.remove("show");
      if (nextBtn) nextBtn.hidden = true;
    }
  }

  // ----- クリア演出（紙吹雪）F5 -----
  function celebrate() {
    const layer = document.getElementById("confetti-layer");
    if (!layer) return;
    const colors = ["#e94560", "#f1c40f", "#1abc9c", "#3498db", "#9b59b6", "#e67e22"];
    for (let i = 0; i < 40; i++) {
      const piece = document.createElement("div");
      piece.className = "confetti";
      piece.style.left = Math.random() * 100 + "%";
      piece.style.background = colors[i % colors.length];
      piece.style.animationDelay = Math.random() * 0.3 + "s";
      piece.style.animationDuration = 1.2 + Math.random() * 0.8 + "s";
      layer.appendChild(piece);
      setTimeout(() => piece.remove(), 2200);
    }
  }

  // ----- トースト通知（他タブでの更新取り込みなど、軽いお知らせ用） -----
  let toastTimer = null;
  function showToast(message, duration) {
    const el = document.getElementById("toast");
    if (!el) return;
    el.textContent = message;
    el.hidden = false;
    // hidden解除直後にクラスを付けないとtransitionが効かないことがあるため次フレームで
    requestAnimationFrame(() => el.classList.add("show"));
    clearTimeout(toastTimer);
    toastTimer = setTimeout(() => {
      el.classList.remove("show");
      setTimeout(() => {
        el.hidden = true;
      }, 300);
    }, duration || 3500);
  }

  function goNextQuest() {
    const idx = TASKS.findIndex((t) => t.id === currentTaskId);
    if (idx >= 0 && idx < TASKS.length - 1) {
      loadTask(TASKS[idx + 1].id);
    }
  }

  // ----- 簡易エラーチェック（ブロックを赤くハイライト） -----
  function checkErrors() {
    const banner = document.getElementById("error-banner");
    const allBlocks = workspace.getAllBlocks(false).filter((b) => !b.isInFlyout);
    let errorCount = 0;

    allBlocks.forEach((b) => {
      let warn = null;
      // 未接続の値ブロック（出力があるのにどこにも刺さっていない）
      if (b.outputConnection && !b.outputConnection.isConnected()) {
        warn = "この値ブロックは差込口につなげる必要があります";
      }
      // 値の差込口が空のブロック
      b.inputList.forEach((input) => {
        if (
          input.connection &&
          input.connection.type === Blockly.INPUT_VALUE &&
          !input.connection.isConnected()
        ) {
          warn = warn || "値が空の差込口があります";
        }
      });
      // カウンタ i / {i} をループの外で使っている（B8）
      if (b.type === "loop_index" && !isInsideLoop(b)) {
        warn = warn || "カウンタ i は繰り返しの中で使います";
      }
      if (usesDynamicCell(b) && !isInsideLoop(b)) {
        warn = warn || "{i} は繰り返しの中で使います";
      }
      b.setWarningText(warn);
      if (warn) errorCount++;
    });

    if (errorCount > 0) {
      banner.textContent = `⚠️ ${errorCount} 個のブロックに問題があります（赤いマークを確認してください）`;
      banner.classList.add("show");
    } else {
      banner.classList.remove("show");
    }
  }

  // ブロックが繰り返し系ブロックの中にあるか
  const LOOP_TYPES = new Set([
    "loop_repeat",
    "loop_range",
    "loop_for_step",
    "loop_for",
    "loop_while",
    "loop_do_until",
    "loop_do_post",
  ]);
  function isInsideLoop(block) {
    let p = block.getSurroundParent();
    while (p) {
      if (LOOP_TYPES.has(p.type)) return true;
      p = p.getSurroundParent();
    }
    return false;
  }
  // セルアドレス系フィールドに {i} を含むか
  function usesDynamicCell(block) {
    return ["CELL", "FROM", "TO"].some((f) => {
      const v = block.getFieldValue && block.getFieldValue(f);
      return v && /\{i\}/.test(v);
    });
  }

  // ----- ヒント -----
  function resetHints(task) {
    const display = document.getElementById("hint-display");
    display.innerHTML = "";
    document.querySelectorAll(".hint-btn").forEach((btn, i) => {
      btn.classList.remove("revealed");
      btn.disabled = i >= (task && task.hints ? task.hints.length : 0);
    });
    document.getElementById("answer-btn").disabled = !(task && task.answer);
  }

  // ----- 右ペイン（📊仮想Excel / 📖リファレンス）のタブ切り替え -----
  const REFERENCE_URL = "https://mkt918.github.io/vba-programming-basics/reference.html";
  function switchRightTab(target) {
    document.querySelectorAll(".right-tab").forEach((t) => {
      t.classList.toggle("active", t.dataset.rightTab === target);
    });
    document.querySelectorAll(".right-tab-panel").forEach((p) => p.classList.remove("active"));
    const panel = document.getElementById(`right-panel-${target}`);
    if (panel) panel.classList.add("active");
    // リファレンスを見ている間は、変数の箱・タイムラインを隠して広く使う
    const rp = document.getElementById("right-pane");
    if (rp) rp.classList.toggle("reference-active", target === "reference");
    if (target === "form") updateFormPanel();
    if (target === "flow") {
      renderFlow();
      const step = view && view.cursor > 0 ? view.steps[view.cursor - 1] : null;
      if (flowChart) flowChart.highlight(step ? step.line : null, step ? step.proc : null, step ? step.branch : undefined);
    }
    updateStepGuide();

    if (target === "reference") {
      const iframe = document.getElementById("reference-iframe");
      const fallback = document.getElementById("reference-fallback");
      if (iframe && !iframe.src) {
        // 初めて開いたときだけ実際に読み込む（無駄な事前アクセスを避ける）
        iframe.src = REFERENCE_URL;
        let loaded = false;
        iframe.addEventListener(
          "load",
          () => {
            loaded = true;
            if (fallback) fallback.hidden = true;
          },
          { once: true }
        );
        // 読み込みに時間がかかっている（あるいは失敗した）場合の案内を出す
        setTimeout(() => {
          if (!loaded && fallback) fallback.hidden = false;
        }, 6000);
      }
    }
  }

  function bindControls() {
    // テーマ切り替え
    document.getElementById("theme-toggle-btn").addEventListener("click", () => {
      const current = localStorage.getItem("theme") || "dark";
      const next = current === "dark" ? "light" : "dark";
      applyTheme(next);
    });

    // コード直打ちエディタの文字サイズ / グリッド線
    initCodeFontSize();
    initCodeGrid();

    // 仮想Excel表示エリアの幅（ドラッグ調整）
    initRightPaneResize();

    // 使い方ガイド
    const helpModal = document.getElementById("help-modal");
    const helpBtn = document.getElementById("help-btn");
    if (helpBtn && helpModal) {
      helpBtn.addEventListener("click", () => (helpModal.hidden = false));
      const helpClose = document.getElementById("help-modal-close");
      if (helpClose) helpClose.addEventListener("click", () => (helpModal.hidden = true));
      helpModal.addEventListener("click", (e) => {
        if (e.target === helpModal) helpModal.hidden = true;
      });
      const tourBtn = document.getElementById("help-tour-btn");
      if (tourBtn) {
        tourBtn.addEventListener("click", () => {
          helpModal.hidden = true;
          if (window.UmbTour) UmbTour.start(true);
        });
      }
    }

    // 右ペイン（📊仮想Excel / 📖リファレンス）のタブ切り替え。
    // モーダルではなくタブにしているのは、コードを見ながら（隠さずに）
    // リファレンスを参照できるようにするため
    document.querySelectorAll(".right-tab").forEach((tab) => {
      tab.addEventListener("click", () => switchRightTab(tab.dataset.rightTab));
    });
    // ヘッダーの「📖 リファレンス」ボタンは右ペインのリファレンスタブへのショートカット
    const referenceBtn = document.getElementById("reference-btn");
    if (referenceBtn) {
      referenceBtn.addEventListener("click", () => switchRightTab("reference"));
    }

    // Esc でモーダルを閉じる
    document.addEventListener("keydown", (e) => {
      if (e.key !== "Escape") return;
      ["help-modal", "quest-modal", "answer-modal"].forEach((id) => {
        const el = document.getElementById(id);
        if (el && !el.hidden) el.hidden = true;
      });
    });

    // クエスト選択ボタン
    buildQuestModal();
    document.getElementById("quest-list-btn").addEventListener("click", () => {
      buildQuestModal();
      document.getElementById("quest-modal").hidden = false;
    });
    document.getElementById("quest-modal-close").addEventListener("click", () => {
      document.getElementById("quest-modal").hidden = true;
    });
    // フリーモードに戻る
    const freeModeBtn = document.getElementById("free-mode-btn");
    if (freeModeBtn) {
      freeModeBtn.addEventListener("click", () => {
        loadFreeMode();
      });
    }

    document.getElementById("quest-modal").addEventListener("click", (e) => {
      if (e.target === document.getElementById("quest-modal")) {
        document.getElementById("quest-modal").hidden = true;
      }
    });

    // ヒントボタン
    document.querySelectorAll(".hint-btn").forEach((btn) => {
      btn.addEventListener("click", () => {
        const idx = Number(btn.dataset.hint);
        const task = TASKS.find((t) => t.id === currentTaskId);
        if (!task || !task.hints[idx]) return;
        if (btn.classList.contains("revealed")) return;
        btn.classList.add("revealed");
        const item = document.createElement("div");
        item.className = "hint-item";
        item.textContent = `💡 ヒント${idx + 1}: ${task.hints[idx]}`;
        document.getElementById("hint-display").appendChild(item);
      });
    });

    // 答えを見る → 確認ダイアログ
    const modal = document.getElementById("answer-modal");
    document.getElementById("answer-btn").addEventListener("click", () => {
      modal.hidden = false;
    });
    document.getElementById("answer-cancel").addEventListener("click", () => {
      modal.hidden = true;
    });
    document.getElementById("answer-confirm").addEventListener("click", () => {
      modal.hidden = true;
      const task = TASKS.find((t) => t.id === currentTaskId);
      if (!task) return;
      const item = document.createElement("div");
      item.className = "hint-item";
      item.style.borderLeftColor = "#6c3483";
      item.innerHTML = `📖 <b>模範解答コード</b><pre style="margin-top:6px;white-space:pre-wrap;font-family:Consolas,monospace;font-size:11px;">${escapeHtml(
        task.answer
      )}</pre>`;
      document.getElementById("hint-display").appendChild(item);

      // 模範解答のブロック配置をワークスペースに読み込む（コードだけでなく組み方も見せる）。
      // コードタブに未反映の変更があれば先に確認する（中止なら文字コードの表示だけで終わる）
      if (task.answerBlocks && confirmReplaceCode()) {
        codeOwner = "blocks";
        codeDirty = false;
        suppressSave = true;
        try {
          workspace.clear();
          Blockly.serialization.workspaces.load(task.answerBlocks, workspace);
          workspace.cleanUp(); // 整列して見やすく
          workspace.scrollCenter();
          document.getElementById("tab-blocks").click(); // ブロックタブで見せる
        } catch (e) {
          console.warn("模範解答ブロックの読み込み失敗:", e);
        }
        suppressSave = false;
        switchToBlocksOwner();
        showToast("📖 模範解答のブロックを表示しました。▶ 実行で動きを確かめよう");
      }
    });

    // 実行コントロール
    // ▶実行：最初から実行して再生する（InputBox は毎回たずねる）。再生中に押すと一時停止、もう一度で続きから
    document.getElementById("run-btn").addEventListener("click", () => {
      if (window.VbaDialog && VbaDialog.isOpen()) return;
      if (view.playing) {
        view.stop(); // 一時停止（カーソル位置は保持）
        return;
      }
      const midRun = runActive && view.cursor > 0 && (view.cursor < view.steps.length || (session && session.status === "waiting"));
      if (midRun) {
        view.play(false); // 一時停止したところから続ける
        return;
      }
      startRun(true);
    });
    // 1コマ：1つずつ進める。答え待ちのところでは質問を出す
    document.getElementById("step-btn").addEventListener("click", () => stepOnce());
    document.getElementById("reset-btn").addEventListener("click", () => {
      view.reset();
      if (session && session.status === "error") hideProgramError();
    });

    // ----- 実行タイムライン：好きな時点へ自由に移動 -----
    const tlSlider = document.getElementById("tl-slider");
    if (tlSlider) {
      tlSlider.addEventListener("input", (e) => view.goToStep(Number(e.target.value)));
    }
    const tlBind = (id, fn) => {
      const el = document.getElementById(id);
      if (el) el.addEventListener("click", fn);
    };
    tlBind("tl-first", () => view.goToStep(0));
    tlBind("tl-prev", () => view.stepBackward());
    tlBind("tl-next", () => stepOnce());
    tlBind("tl-last", () => {
      view.goToEnd();
      if (session && session.status === "error" && view.cursor === view.steps.length) showProgramError(session.error, !runActive);
    });

    // ----- 🪟 ユーザーフォームの道具 -----
    document.querySelectorAll(".form-add-btn").forEach((b) => {
      b.addEventListener("click", () => {
        formPanel.addControl(b.dataset.add);
        lastArea = "form";
        updateStepGuide();
      });
    });
    document.getElementById("form-show-names").addEventListener("change", (e) => formPanel.setShowNames(e.target.checked));
    document.getElementById("form-undo").addEventListener("click", () => formPanel.undo());
    document.getElementById("form-stop").addEventListener("click", () => stopRun());
    document.getElementById("run-target").addEventListener("change", (e) => {
      runTarget = e.target.value;
      runActive = false;
      runPreview();
    });

    // ----- ① フォームを作る → ② 命令を組む → ③ ▶ 実行 -----
    document.querySelectorAll(".step-chip").forEach((chip) => {
      chip.addEventListener("click", () => {
        const step = Number(chip.dataset.step);
        if (step === 1) {
          lastArea = "form";
          switchRightTab("form");
        } else if (step === 2) {
          lastArea = "program";
          if (codeOwner === "code") document.getElementById("tab-code").click();
          else document.getElementById("tab-blocks").click();
        } else {
          document.getElementById("run-btn").click();
        }
        updateStepGuide();
      });
    });
    document.getElementById("right-panel-form").addEventListener("pointerdown", () => {
      if (!runActive) lastArea = "form";
      updateStepGuide();
    });
    document.querySelector(".editor-pane").addEventListener("pointerdown", () => {
      lastArea = "program";
      updateStepGuide();
    });

    // ----- プログラムのまちがいパネル -----
    document.getElementById("program-error-close").addEventListener("click", () => {
      const panel = document.getElementById("program-error");
      if (panel) panel.hidden = true;
    });
    document.getElementById("program-error-goto").addEventListener("click", () => {
      if (!shownError || !shownError.line) return;
      document.getElementById("tab-code").click();
      setTimeout(() => {
        codeEditor.refresh();
        codeEditor.focus();
        codeEditor.setCursor({ line: shownError.line - 1, ch: Math.max(0, (shownError.col || 1) - 1) });
        codeEditor.scrollIntoView({ line: shownError.line - 1, ch: 0 }, 80);
      }, 80);
    });
    document.getElementById("program-error-fix").addEventListener("click", () => {
      if (!codeEditor) return;
      const fixed = VbaLexer.fixZenkaku(codeEditor.getValue());
      codeEditor.setValue(fixed); // 生徒が直したのと同じ扱い（change イベントで実行し直す）
      showToast("🔧 全角の文字を半角に直しました");
    });

    // ----- VBAコードで実行しているときのブロックタブの案内 -----
    document.getElementById("code-owner-to-code").addEventListener("click", () => {
      document.getElementById("tab-code").click();
    });
    document.getElementById("code-owner-to-blocks").addEventListener("click", () => {
      if (!blocksMatchCode && !confirm("VBAコードタブで書いた変更は消えて、いまのブロックからコードを作り直します。よいですか？")) return;
      switchToBlocksOwner();
    });

    // ----- 実行の出発点（初期データ）の管理 -----
    const bakeBtn = document.getElementById("bake-btn");
    if (bakeBtn) {
      bakeBtn.addEventListener("click", () => {
        const n = view.bakeCurrentAsInitial();
        onCellEdited(); // 保存 + ステップ再構築（新しい出発点から作り直される）
        document.getElementById("step-status").textContent =
          `📌 今の表示（${n} セル）を実行の出発点として保存しました`;
      });
    }
    const clearInitialBtn = document.getElementById("clear-initial-btn");
    if (clearInitialBtn) {
      clearInitialBtn.addEventListener("click", () => {
        if (Object.keys(view.getInitialCells() || {}).length === 0) return;
        if (!confirm("実行の出発点のデータを全部消しますか？")) return;
        view.clearInitial();
        onCellEdited();
      });
    }

    document.getElementById("speed-slider").addEventListener("input", (e) => {
      // スライダー右 = 速い になるよう反転
      view.setSpeed(1600 - Number(e.target.value));
    });
    view.setSpeed(1600 - 500);

    // Excel拡大縮小
    const zoomSlider = document.getElementById("excel-zoom-slider");
    if (zoomSlider) {
      zoomSlider.addEventListener("input", (e) => {
        view.setZoom(Number(e.target.value));
      });
      view.setZoom(100);
    }

    // コードコピー
    document.getElementById("copy-btn").addEventListener("click", () => {
      const code = codeEditor ? codeEditor.getValue() : "";
      const btn = document.getElementById("copy-btn");
      // file:// で直接開いた場合など、非セキュアなコンテキストでは
      // navigator.clipboard 自体が存在しないことがある
      if (!navigator.clipboard || !navigator.clipboard.writeText) {
        alert("この環境では自動コピーが使えません。コードを選択して手動でコピーしてください。");
        return;
      }
      navigator.clipboard
        .writeText(code)
        .then(() => {
          btn.textContent = "✅ コピー済み";
          setTimeout(() => (btn.textContent = "📋 コピー"), 1500);
        })
        .catch(() => {
          alert("コピーに失敗しました。コードを選択して手動でコピーしてください。");
        });
    });

    // エディタタブ切り替え
    document.querySelectorAll(".editor-tab").forEach((tab) => {
      tab.addEventListener("click", () => {
        const target = tab.dataset.tab;
        const leavingCodeTab = document.getElementById("tab-code-panel").classList.contains("active");
        // コードタブを離れるときは、書いたコードをブロックにも並べてみる（コードはそのまま実行する）
        if (leavingCodeTab && target !== "code" && codeDirty) applyCodeToBlocks();
        updateCodeOwnerOverlay();
        document.querySelectorAll(".editor-tab").forEach((t) => t.classList.remove("active"));
        document.querySelectorAll(".tab-panel").forEach((p) => p.classList.remove("active"));
        tab.classList.add("active");
        document.getElementById(`tab-${target}-panel`).classList.add("active");
        // ブロックタブに戻ったときBlocklyをリサイズ
        if (target === "blocks") {
          setTimeout(() => Blockly.svgResize(workspace), 50);
        }
        if (target === "code" && codeEditor) {
          setTimeout(() => codeEditor.refresh(), 50); // 非表示中にサイズが取れていないため
        }
      });
    });

    // 説明とヒントの開閉（フリーモード起動時は前回の状態を引き継ぐ）
    const collapseBtn = document.getElementById("quest-collapse-btn");
    collapseBtn.addEventListener("click", () => {
      const collapsed = document
        .getElementById("main-layout")
        .classList.contains("quest-collapsed");
      applyQuestCollapsed(!collapsed);
    });
    applyQuestCollapsed((localStorage.getItem("questCollapsed") || "1") === "1");

    // 全部消す（F2）
    document.getElementById("clear-all-btn").addEventListener("click", () => {
      if (workspace.getTopBlocks(false).length === 0 && codeOwner === "blocks") return;
      const msg = codeOwner === "code"
        ? "組み立てたブロックと、VBAコードタブで書いたプログラムを全部消しますか？"
        : "組み立てたブロックを全部消しますか？";
      if (!confirm(msg)) return;
      workspace.clear();
      switchToBlocksOwner();
    });

    // 次のクエストへ（F4）
    const nextBtn = document.getElementById("next-quest-btn");
    if (nextBtn) nextBtn.addEventListener("click", goNextQuest);

    // 📤 先生に送る（作ったものを見られるリンクをコピーする）
    const shareBtn = document.getElementById("share-btn");
    if (shareBtn) {
      shareBtn.addEventListener("click", () => {
        try {
          if (codeDirty) applyCodeToBlocks();
          const json = JSON.stringify(buildWorkPayload());
          const encoded = encodeURIComponent(btoa(unescape(encodeURIComponent(json))));
          const url = `${location.origin}${location.pathname}#share=${encoded}`;
          const done = () => showToast("📤 リンクをコピーしました。Classroom や Teams などに貼り付けて、先生に送ろう", 5000);
          if (!navigator.clipboard || !navigator.clipboard.writeText) {
            // 自動コピーできない環境では、URLを選んでコピーできるダイアログで代替
            prompt("このリンクをコピーして、先生に送ろう（この環境では自動コピーが使えません）:", url);
            return;
          }
          navigator.clipboard
            .writeText(url)
            .then(done)
            .catch(() => {
              prompt("このリンクをコピーして、先生に送ろう（自動コピーに失敗しました）:", url);
            });
        } catch (e) {
          console.warn("共有リンク生成失敗:", e);
        }
      });
    }

    // 💾 保存 / 📂 開く（PC の再起動でブラウザのデータが消えても、ファイルから戻せるように）
    document.getElementById("file-save-btn").addEventListener("click", saveWorkToFile);
    const fileInput = document.getElementById("file-open-input");
    document.getElementById("file-open-btn").addEventListener("click", () => {
      fileInput.value = "";
      fileInput.click();
    });
    fileInput.addEventListener("change", () => {
      const f = fileInput.files && fileInput.files[0];
      if (f) openWorkFile(f);
    });
  }

  // ===== 作品のまとめ（ファイル保存・先生に送るリンクで共通） =====
  function buildWorkPayload() {
    const dom = Blockly.Xml.workspaceToDom(workspace);
    const payload = {
      app: "UIMacroBuilder",
      v: 2,
      savedAt: new Date().toISOString(),
      task: currentTaskId,
      xml: Blockly.Xml.domToText(dom),
      initial: view.getInitialCells(),
      form: formPanel.getForm(),
    };
    if (codeOwner === "code") payload.code = { text: codeEditor.getValue(), match: blocksMatchCode, reason: blocksMismatchReason };
    return payload;
  }

  // まとめを読み込んで、いまの作業に入れる（shared: 共有リンクで見ているだけ＝保存しない）
  function applyWorkPayload(payload, shared) {
    runActive = false;
    suppressSave = true;
    workspace.clear();
    if (payload.xml) {
      const dom = Blockly.utils.xml.textToDom(payload.xml);
      Blockly.Xml.domToWorkspace(dom, workspace);
    }
    suppressSave = false;
    sharedViewMode = !!shared;
    view.setInitialCells(payload.initial || {});
    formPanel.setForm(payload.form || UserFormUtil.defaultForm());
    if (payload.code && typeof payload.code.text === "string") {
      codeOwner = "code";
      blocksMatchCode = !!payload.code.match;
      blocksMismatchReason = payload.code.reason || "";
      setEditorCode(payload.code.text);
    } else {
      codeOwner = "blocks";
      blocksMatchCode = true;
      blocksMismatchReason = "";
    }
    codeDirty = false;
    if (!shared) {
      const key = currentKey();
      savedInitial[key] = view.getInitialCells();
      savedForms[key] = formPanel.getForm();
      saveCodeState();
    }
    updateCodeSyncStatus();
    updateCodeOwnerOverlay();
    onWorkspaceChange();
  }

  function saveWorkToFile() {
    if (codeDirty) applyCodeToBlocks();
    const d = new Date();
    const pad = (n) => String(n).padStart(2, "0");
    const def = `vba_${d.getFullYear()}${pad(d.getMonth() + 1)}${pad(d.getDate())}_${pad(d.getHours())}${pad(d.getMinutes())}`;
    const raw = prompt("ファイルの名前を入れてください（例: jis4-2-01）", def);
    if (raw === null) return;
    const name = (raw.trim() || def).replace(/[\\/:*?"<>|]/g, "_").replace(/\.vbalearn$/i, "") + ".vbalearn";
    const blob = new Blob([JSON.stringify(buildWorkPayload(), null, 1)], { type: "application/json" });
    const a = document.createElement("a");
    a.href = URL.createObjectURL(blob);
    a.download = name;
    document.body.appendChild(a);
    a.click();
    setTimeout(() => {
      URL.revokeObjectURL(a.href);
      a.remove();
    }, 1000);
    showToast(`💾 「${name}」をダウンロードフォルダに保存しました`, 4500);
  }

  function openWorkFile(file) {
    const reader = new FileReader();
    reader.onload = () => {
      let payload;
      try {
        payload = JSON.parse(String(reader.result));
      } catch (e) {
        alert("このファイルは開けませんでした（💾 保存 で作ったファイルを選んでね）");
        return;
      }
      if (!payload || (payload.app !== "UIMacroBuilder" && !payload.xml)) {
        alert("このファイルは開けませんでした（💾 保存 で作ったファイルを選んでね）");
        return;
      }
      if (!confirm(`「${file.name}」を開きますか？\nいま作っているものは、開いたファイルの内容に置きかわります。`)) return;
      applyWorkPayload(payload, false);
      showToast(`📂 「${file.name}」を開きました`);
    };
    reader.readAsText(file);
  }

  // 共有 URL から作品を復元（先生が生徒のリンクを開いたとき）
  function loadFromShareUrl() {
    const m = location.hash.match(/share=([^&]+)/);
    if (!m) return false;
    try {
      const text = decodeURIComponent(escape(atob(decodeURIComponent(m[1]))));
      let payload;
      if (text.trim().startsWith("{")) payload = JSON.parse(text);
      else payload = { v: 1, xml: text }; // 古い形式（ブロックだけ）
      history.replaceState(null, "", location.pathname); // ハッシュ消去
      applyWorkPayload(payload, true);
      showToast("📤 送られてきた作品を表示しています（この画面での変更は保存されません）", 5000);
      return true;
    } catch (e) {
      console.warn("共有リンク復元失敗:", e);
      return false;
    }
  }

  function escapeHtml(s) {
    return s.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
  }
})();
