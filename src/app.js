/* ===== メインアプリ =====
 * Blockly 初期化・課題切り替え・ヒント・コード生成・Excel プレビューを結線する。
 */

(function () {
  let workspace = null;
  let view = null;
  let codeEditor = null; // CodeMirror インスタンス（VBAコード直打ち）
  let codeDirty = false; // コード側が編集されてブロックへ未反映かどうか
  let suppressCodeChange = false; // setValue によるプログラム的な更新中フラグ
  let currentTaskId = null;
  let hasLoadedOnce = false; // 初回ロード時は保存前のブロックが存在しないため保存をスキップ
  let suppressSave = false; // 復元中の保存抑制フラグ（B7）
  let sharedViewMode = false; // 共有リンク閲覧中は保存スロットに書き込まない
  const inputCache = {}; // InputBox の答え（block.id -> 値）。▶実行ごとにクリア
  const savedBlocks = {}; // taskId -> XML文字列（課題ごとのセーブスロット）
  const savedInitial = {}; // taskId -> 仮想Excelの初期データ
  const solved = new Set();

  const LS_BLOCKS = "umb_blocks";
  const LS_SOLVED = "umb_solved";
  const LS_INITIAL = "umb_initial";
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
    } catch (e) {
      console.warn("読み込み失敗:", e);
    }
  }

  // ----- 仮想Excelのセル編集時（F: 直接入力）-----
  function onCellEdited() {
    const key = currentTaskId === null ? FREE_MODE_ID : currentTaskId;
    savedInitial[key] = view.getInitialCells();
    persist();
    // 初期データが変わったのでステップを作り直す
    onWorkspaceChange();
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
    const raw = window.prompt("新しい変数の名前を入力してください（ひらがな不可、i / arr は予約語）", "");
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

  // for / while / until / if / ifelse と打って Tab で展開できる骨組み。
  // cursorLine/cursorCol は挿入後にカーソルを置く位置（0行目基準、baseIndent加算前）
  const CODE_SNIPPETS = {
    for: { template: "For i = 1 To 10\n    \nNext i", cursorLine: 1, cursorCol: 4 },
    while: { template: "Do While \nLoop", cursorLine: 0, cursorCol: 9 },
    until: { template: "Do Until \nLoop", cursorLine: 0, cursorCol: 9 },
    if: { template: "If  Then\nEnd If", cursorLine: 0, cursorCol: 3 },
    ifelse: { template: "If  Then\nElse\nEnd If", cursorLine: 0, cursorCol: 3 },
  };

  // 現在行のインデントを引き継ぎつつスニペットを展開する
  // 注意: CODE_SNIPPETS のテンプレート文字列自体に、行ごとの相対インデント
  // （本体行は4スペース、閉じキーワード行は0）を埋め込んである。
  // ここでは各行の先頭に baseIndent を足すだけでよい（さらに4スペース加算しない）
  function insertCodeSnippet(cm, key) {
    const snip = CODE_SNIPPETS[key];
    if (!snip) return false;
    const cursor = cm.getCursor();
    const line = cm.getLine(cursor.line);
    const baseIndent = (line.match(/^(\s*)/) || ["", ""])[1];
    const wordStart = { line: cursor.line, ch: cursor.ch - key.length };
    const lines = snip.template.split("\n");
    const indented = lines.map((l, i) => (i === 0 ? l : baseIndent + l));
    cm.replaceRange(indented.join("\n"), wordStart, cursor);
    const targetLine = wordStart.line + snip.cursorLine;
    const targetCh = (snip.cursorLine === 0 ? wordStart.ch : baseIndent.length) + snip.cursorCol;
    cm.setCursor({ line: targetLine, ch: targetCh });
    return true;
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
    const cursor = cm.getCursor();
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

  // ----- コード編集済み状態の表示更新 -----
  function updateCodeSyncStatus() {
    const el = document.getElementById("code-sync-status");
    if (!el) return;
    if (codeDirty) {
      el.textContent = "✏️ 未反映（ブロックタブに戻すと反映）";
      el.hidden = false;
    } else {
      el.hidden = true;
    }
  }

  function hideCodeError() {
    const banner = document.getElementById("code-error-banner");
    if (banner) banner.hidden = true;
    if (codeEditor) {
      for (let i = 0; i < codeEditor.lineCount(); i++) {
        codeEditor.removeLineClass(i, "background", "code-error-line");
      }
    }
  }

  function showCodeError(err) {
    const banner = document.getElementById("code-error-banner");
    if (banner) {
      const lineText = err.line ? `${err.line}行目: ` : "";
      banner.textContent = `⚠️ ${lineText}${err.message}`;
      banner.hidden = false;
    }
    if (codeEditor && err.line) {
      codeEditor.addLineClass(err.line - 1, "background", "code-error-line");
      codeEditor.scrollIntoView({ line: err.line - 1, ch: 0 }, 60);
    }
  }

  // ----- コード → ブロック 変換の実行 -----
  // タブ切替時 / エディタからのフォーカス離脱時に呼ばれる。
  // 解析に失敗した場合はブロックを一切変更せず、エラー表示のみ行う（要件4.2）
  function applyCodeToBlocks() {
    if (!codeEditor || !workspace) return;
    if (!codeDirty) return;
    const text = codeEditor.getValue();
    hideCodeError();

    let result;
    try {
      result = window.VbaParser.parseVBA(text);
    } catch (e) {
      showCodeError(e);
      return; // ブロックは変更しない。codeDirty も維持（コードは消さない）
    }

    suppressSave = true;
    try {
      workspace.clear();
      if (result.root) {
        window.VbaParser.resolveVariables(result.root, workspace);
        const json = window.VbaParser.toWorkspaceJson(result.root);
        Blockly.serialization.workspaces.load(json, workspace);
        workspace.cleanUp();
      }
    } catch (e) {
      suppressSave = false;
      console.warn("ブロック構築に失敗:", e);
      showCodeError({ message: "ブロックを組み立てられませんでした（" + e.message + "）", line: null });
      return;
    }
    suppressSave = false;

    codeDirty = false;
    updateCodeSyncStatus();
    onWorkspaceChange(); // ブロックから正規化されたコードで表示を更新
  }

  // ----- コード編集中の内容を、課題切替・全消去などの前に守る -----
  // 戻り値: true = 続行してよい（反映できた or 生徒が破棄に同意した）
  //         false = 中止すべき（生徒が「戻る」を選んだ）
  // codeDirty のまま呼び出し元で workspace.clear() 等を続けると、
  // 手打ちしたコードが無警告で失われてしまうため、必ずこれを通す
  function confirmDiscardDirtyCode() {
    if (!codeDirty) return true;
    applyCodeToBlocks(); // まず反映を試みる（成功すれば codeDirty は false になる）
    if (!codeDirty) return true;
    const ok = confirm(
      "VBAコードタブに、まだブロックへ反映されていない変更があります。\n" +
        "このまま進めると、その変更は失われます。続けますか？"
    );
    if (ok) {
      codeDirty = false;
      hideCodeError();
      updateCodeSyncStatus(); // 「✏️ 未反映」バッジを消す
    }
    return ok;
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
    });

    // 📦 変数カテゴリ：日本語ブロック + 「変数を作る」ボタン + 登録済み変数の取り出し
    workspace.registerToolboxCategoryCallback("VARIABLE_JP", function (ws) {
      const items = [
        { kind: "button", text: "＋ 変数を作る", callbackKey: "CREATE_VARIABLE_JP" },
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
      },
      onStepChange: updateTimeline, // タイムラインのスライダー・カウンタを同期
    });

    // 変数パネルを常時更新（VAR_CREATE/DELETE/RENAME 時）
    view.bindWorkspace(workspace);

    // ブロック変更 → コード再生成 + ステップ再構築
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
      },
    });
    codeEditor.on("change", (cm, changeObj) => {
      if (suppressCodeChange) return;
      codeDirty = true;
      updateCodeSyncStatus();
      hideCodeError(); // 直前のエラーは打ち直している間は消しておく
      if (changeObj.origin !== "setValue") {
        const lineNo = changeObj.to.line;
        setTimeout(() => dedentClosingKeywordLine(cm, lineNo), 0);
      }
    });
    codeEditor.on("blur", () => {
      if (codeDirty) applyCodeToBlocks();
    });

    bindControls();
    loadFreeMode(); // 起動時はクエスト未選択＝フリーモード
    loadFromShareUrl(); // 共有リンクがあればブロック復元（F8）

    // 他タブでの更新を取り込む（複数タブを開いたとき、後から persist した方が
    // 一方的に上書きしてしまう事故を防ぐ）
    window.addEventListener("storage", onOtherTabStorageChange);
  });

  // ----- 他タブでの localStorage 更新を取り込む -----
  // storage イベントは「変更したタブ以外」でのみ発火する。
  // クリア済みフラグは合算（減らさない）、ブロック/初期データは
  // 「今このタブで編集中でない課題」だけ取り込む（作業中のものを壊さないため）
  function onOtherTabStorageChange(e) {
    if (![LS_BLOCKS, LS_SOLVED, LS_INITIAL].includes(e.key)) return;
    try {
      const otherSolved = JSON.parse(localStorage.getItem(LS_SOLVED) || "[]");
      let solvedChanged = false;
      otherSolved.forEach((id) => {
        if (!solved.has(id)) {
          solved.add(id);
          solvedChanged = true;
        }
      });

      const otherBlocks = JSON.parse(localStorage.getItem(LS_BLOCKS) || "{}");
      const otherInitial = JSON.parse(localStorage.getItem(LS_INITIAL) || "{}");
      const activeKey = currentTaskId === null ? FREE_MODE_ID : currentTaskId;
      Object.keys(otherBlocks).forEach((k) => {
        if (k !== activeKey) savedBlocks[k] = otherBlocks[k];
      });
      Object.keys(otherInitial).forEach((k) => {
        if (k !== activeKey) savedInitial[k] = otherInitial[k];
      });

      if (solvedChanged) updateCurrentQuestDisplay();
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

    // VBA コード生成 → コードエディタへ反映
    // ただしコード側が編集済み（codeDirty）のときは上書きしない
    // （手打ち中のコードを消さないため。ブロックタブに戻った時に反映される）
    const code = generateVBA(workspace);
    if (codeEditor && !codeDirty) {
      // codeDirty=false のときは生徒はコード欄を編集していないので、
      // カーソル位置を気にせずまるごと置き換えてよい
      suppressCodeChange = true;
      codeEditor.setValue(code);
      suppressCodeChange = false;
    }

    // エラーチェック（未接続の値ブロックなど簡易検出）
    checkErrors();

    // ステップ再構築（仮想Excelの初期データを土台にする）
    // 編集中は interactive を付けない＝InputBox はダイアログを出さず仮値で動く
    try {
      const result = buildSteps(workspace, view.getInitialCells(), { inputCache });
      view.load(result.steps, result.limitHit);
      checkQuestClear();
    } catch (e) {
      console.warn("ステップ生成エラー:", e);
    }

    saveCurrentBlocks();

    // 変数パネルを現在のステップの値で更新（常時表示）
    // step は { scope, key, model, desc } なので値は model.vars にある
    const curStep = view.steps[view.cursor - 1];
    const curVars = curStep && curStep.model ? curStep.model.vars || {} : {};
    const marker = curStep ? { scope: curStep.scope, key: curStep.key } : null;
    view.renderVars(curVars, marker);

    updateBaselineLabel();
    updateCanvasGuide();
  }

  // シンタックスハイライトは CodeMirror（vbscript モード）が担当するため、
  // 旧・単一パス正規表現ハイライタ（highlightVBA）は撤去済み。

  // ----- クエストクリア自動判定 -----
  function checkQuestClear() {
    const task = TASKS.find((t) => t.id === currentTaskId);
    const banner = document.getElementById("clear-banner");
    const nextBtn = document.getElementById("next-quest-btn");
    if (!task || typeof task.check !== "function") {
      banner.classList.remove("show");
      if (nextBtn) nextBtn.hidden = true;
      return;
    }
    // セルへの直接入力だけ（ブロックを一切組んでいない）でクリア判定が
    // 通ってしまわないよう、実行可能なステップが1つ以上あることを条件にする
    const hasSteps = view.steps.length > 0;
    let passed = false;
    if (hasSteps) {
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
    "loop_while",
    "loop_do_until",
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

  function bindControls() {
    // テーマ切り替え
    document.getElementById("theme-toggle-btn").addEventListener("click", () => {
      const current = localStorage.getItem("theme") || "dark";
      const next = current === "dark" ? "light" : "dark";
      applyTheme(next);
    });

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
      if (task.answerBlocks && confirmDiscardDirtyCode()) {
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
        onWorkspaceChange();
        document.getElementById("step-status").textContent =
          "📖 模範解答のブロックを表示しました。▶ 実行で動きを確かめよう";
      }
    });

    // 実行コントロール
    // 実行中クリックで一時停止。InputBox があれば最初からの実行時にだけ質問する
    document.getElementById("run-btn").addEventListener("click", () => {
      if (view.playing) {
        view.stop(); // 一時停止（カーソル位置は保持）
        return;
      }
      const fresh = view.cursor === 0 || view.cursor >= view.steps.length;
      const hasInput = workspace.getAllBlocks(false).some((b) => b.type === "io_inputbox");
      if (hasInput && fresh) {
        Object.keys(inputCache).forEach((k) => delete inputCache[k]); // 毎回聞き直す
        try {
          const result = buildSteps(workspace, view.getInitialCells(), {
            inputCache,
            interactive: true,
          });
          view.load(result.steps, result.limitHit);
        } catch (e) {
          console.warn("ステップ生成エラー:", e);
        }
      }
      view.play();
    });
    document.getElementById("step-btn").addEventListener("click", () => {
      view.stop();
      view.stepForward();
    });
    document.getElementById("reset-btn").addEventListener("click", () => view.reset());

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
    tlBind("tl-next", () => {
      view.stop();
      view.stepForward();
    });
    tlBind("tl-last", () => view.goToEnd());

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
        // コードタブを離れるときは、手打ちしたコードをブロックへ反映する。
        // 解析エラーで反映できなかった場合は、エラーが見えるようコードタブに留まる
        if (leavingCodeTab && target !== "code" && codeDirty) {
          applyCodeToBlocks();
          if (codeDirty) return; // まだ未反映＝エラー中。タブは切り替えない
        }
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
      if (workspace.getTopBlocks(false).length === 0) return;
      if (!confirm("組み立てたブロックを全部消しますか？")) return;
      if (!confirmDiscardDirtyCode()) return; // コード未反映があれば先に確認
      workspace.clear();
      onWorkspaceChange();
    });

    // 次のクエストへ（F4）
    const nextBtn = document.getElementById("next-quest-btn");
    if (nextBtn) nextBtn.addEventListener("click", goNextQuest);

    // 共有リンク（F8）
    const shareBtn = document.getElementById("share-btn");
    if (shareBtn) {
      shareBtn.addEventListener("click", () => {
        try {
          const dom = Blockly.Xml.workspaceToDom(workspace);
          const xml = Blockly.Xml.domToText(dom);
          const encoded = encodeURIComponent(btoa(unescape(encodeURIComponent(xml))));
          const url = `${location.origin}${location.pathname}#share=${encoded}`;
          if (!navigator.clipboard || !navigator.clipboard.writeText) {
            // 自動コピーできない環境では、URLを選んでコピーできるダイアログで代替
            prompt("このURLをコピーしてください（この環境では自動コピーが使えません）:", url);
            return;
          }
          navigator.clipboard
            .writeText(url)
            .then(() => {
              shareBtn.textContent = "✅ リンクをコピー";
              setTimeout(() => (shareBtn.textContent = "🔗 共有"), 1500);
            })
            .catch(() => {
              prompt("このURLをコピーしてください（自動コピーに失敗しました）:", url);
            });
        } catch (e) {
          console.warn("共有リンク生成失敗:", e);
        }
      });
    }
  }

  // 共有 URL からブロックを復元（F8）
  function loadFromShareUrl() {
    const m = location.hash.match(/share=([^&]+)/);
    if (!m) return false;
    if (!confirmDiscardDirtyCode()) return false; // コード未反映があれば先に確認
    try {
      const xml = decodeURIComponent(escape(atob(decodeURIComponent(m[1]))));
      suppressSave = true;
      workspace.clear();
      const dom = Blockly.utils.xml.textToDom(xml);
      Blockly.Xml.domToWorkspace(dom, workspace);
      suppressSave = false;
      sharedViewMode = true; // 閲覧モード：課題スロットを上書きしない
      history.replaceState(null, "", location.pathname); // ハッシュ消去
      onWorkspaceChange();
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
