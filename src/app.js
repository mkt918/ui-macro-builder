/* ===== メインアプリ =====
 * Blockly 初期化・課題切り替え・ヒント・コード生成・Excel プレビューを結線する。
 */

(function () {
  let workspace = null;
  let view = null;
  let currentTaskId = null;
  let suppressSave = false; // 復元中の保存抑制フラグ（B7）
  let sharedViewMode = false; // 共有リンク閲覧中は保存スロットに書き込まない
  const inputCache = {}; // InputBox の答え（block.id -> 値）。▶実行ごとにクリア
  const savedBlocks = {}; // taskId -> XML文字列（課題ごとのセーブスロット）
  const savedInitial = {}; // taskId -> 仮想Excelの初期データ
  const solved = new Set();

  const LS_BLOCKS = "umb_blocks";
  const LS_SOLVED = "umb_solved";
  const LS_INITIAL = "umb_initial";

  // ----- テーマ（一度だけ定義: B5）-----
  let _themes = null;
  function getThemes() {
    if (_themes) return _themes;
    _themes = {
      dark: Blockly.Theme.defineTheme("umb_dark", {
        base: Blockly.Themes.Classic,
        componentStyles: {
          workspaceBackgroundColour: "#0d1117",
          toolboxBackgroundColour: "#16213e",
          toolboxForegroundColour: "#e0e0e0",
          flyoutBackgroundColour: "#12182e",
          flyoutForegroundColour: "#e0e0e0",
          scrollbarColour: "#30363d",
        },
      }),
      light: Blockly.Theme.defineTheme("umb_light", {
        base: Blockly.Themes.Classic,
        componentStyles: {
          workspaceBackgroundColour: "#fafafa",
          toolboxBackgroundColour: "#ffffff",
          toolboxForegroundColour: "#333333",
          flyoutBackgroundColour: "#f5f5f5",
          flyoutForegroundColour: "#333333",
          scrollbarColour: "#cccccc",
        },
      }),
    };
    return _themes;
  }

  // ----- レイアウト管理 -----
  function applyLayout(layout) {
    const main = document.getElementById("main-layout");
    if (layout === "B") {
      main.classList.add("layout-b");
    } else {
      main.classList.remove("layout-b");
    }
    document.querySelectorAll(".layout-btn").forEach((btn) => {
      btn.classList.toggle("active", btn.dataset.layout === layout);
    });
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
    } else {
      document.body.classList.remove("light-mode");
      document.getElementById("theme-toggle-btn").textContent = "🌙";
      if (workspace) workspace.setTheme(themes.dark);
    }
    localStorage.setItem("theme", mode);
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
    if (currentTaskId) {
      savedInitial[currentTaskId] = view.getInitialCells();
      persist();
    }
    // 初期データが変わったのでステップを作り直す
    onWorkspaceChange();
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
    workspace.registerButtonCallback("CREATE_VARIABLE_JP", function (button) {
      Blockly.Variables.createVariableButtonHandler(button.getTargetWorkspace());
    });

    view = new ExcelView({
      table: document.getElementById("excel-table"),
      ref: document.getElementById("cell-ref"),
      formula: document.getElementById("formula-bar"),
      status: document.getElementById("step-status"),
      array: document.getElementById("array-viz"),
      tabs: document.getElementById("excel-tabs"),
      onEdit: onCellEdited, // 仮想Excelのセル編集時
    });

    // ブロック変更 → コード再生成 + ステップ再構築
    workspace.addChangeListener(onWorkspaceChange);

    bindControls();
    loadTask(TASKS[0].id);
    loadFromShareUrl(); // 共有リンクがあればブロック復元（F8）
  });

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
        loadTask(quest.id);
        document.getElementById("quest-modal").hidden = true;
      });
      list.appendChild(item);
    });
  }

  function updateCurrentQuestDisplay() {
    const task = TASKS.find((t) => t.id === currentTaskId);
    if (task) {
      const check = solved.has(task.id) ? "✓ " : "";
      document.getElementById("current-quest").textContent =
        `${check}📍 ${task.title} ${"★".repeat(task.difficulty)}`;
    }
  }

  // ----- 課題読み込み -----
  function loadTask(taskId) {
    // 現在のブロックを保存（共有閲覧中は保存されない）
    if (currentTaskId && workspace) {
      saveCurrentBlocks();
    }
    sharedViewMode = false; // 課題を選んだら通常モードに復帰
    currentTaskId = taskId;
    const task = TASKS.find((t) => t.id === taskId);
    if (!task) return;

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
      const dom = Blockly.Xml.workspaceToDom(workspace);
      savedBlocks[currentTaskId] = Blockly.Xml.domToText(dom);
      persist();
    } catch (e) {
      console.warn("ブロック保存失敗:", e);
    }
  }

  // ----- ワークスペース変更時 -----
  function onWorkspaceChange(event) {
    if (event && event.isUiEvent) return;

    // VBA コード生成 + シンタックスハイライト（F6）
    const code = generateVBA(workspace);
    document.querySelector("#code-output code").innerHTML = highlightVBA(code);

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

    if (currentTaskId) saveCurrentBlocks();
  }

  // ----- VBA シンタックスハイライト（F6）-----
  // 単一パスのトークナイザ。挿入したタグを再処理しないので安全。
  function highlightVBA(code) {
    const tokenRe =
      /('[^\n]*)|("[^"]*")|\b(\d+)\b|\b(Sub|End|Dim|As|Integer|Long|Single|String|Variant|For|To|Step|Next|If|Then|Else|And|Or|Not|Mod|True|False|Do|While|Until|Loop)\b|\b(Cells|Range|Rows|Worksheets|WorksheetFunction|Round|MsgBox|InputBox|Date)\b/g;
    let out = "";
    let last = 0;
    let m;
    while ((m = tokenRe.exec(code)) !== null) {
      out += escapeHtml(code.slice(last, m.index));
      if (m[1]) out += `<span class="vba-cm">${escapeHtml(m[1])}</span>`;
      else if (m[2]) out += `<span class="vba-st">${escapeHtml(m[2])}</span>`;
      else if (m[3]) out += `<span class="vba-nm">${escapeHtml(m[3])}</span>`;
      else if (m[4]) out += `<span class="vba-kw">${escapeHtml(m[4])}</span>`;
      else if (m[5]) out += `<span class="vba-fn">${escapeHtml(m[5])}</span>`;
      last = m.index + m[0].length;
    }
    out += escapeHtml(code.slice(last));
    return out;
  }

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
    let passed = false;
    try {
      passed = task.check(view.finalModel());
    } catch (e) {
      passed = false;
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
      btn.disabled = i >= (task.hints ? task.hints.length : 0);
    });
    document.getElementById("answer-btn").disabled = !task.answer;
  }

  function bindControls() {
    // テーマ切り替え
    document.getElementById("theme-toggle-btn").addEventListener("click", () => {
      const current = localStorage.getItem("theme") || "dark";
      const next = current === "dark" ? "light" : "dark";
      applyTheme(next);
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
    });

    // 実行コントロール
    // InputBox ブロックがあれば、実行時にだけ質問して答えでステップを作り直す
    document.getElementById("run-btn").addEventListener("click", () => {
      const hasInput = workspace.getAllBlocks(false).some((b) => b.type === "io_inputbox");
      if (hasInput) {
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
    document.getElementById("speed-slider").addEventListener("input", (e) => {
      // スライダー右 = 速い になるよう反転
      view.setSpeed(1600 - Number(e.target.value));
    });
    view.setSpeed(1600 - 500);

    // コードコピー
    document.getElementById("copy-btn").addEventListener("click", () => {
      const code = document.querySelector("#code-output code").textContent;
      navigator.clipboard.writeText(code).then(() => {
        const btn = document.getElementById("copy-btn");
        btn.textContent = "✅ コピー済み";
        setTimeout(() => (btn.textContent = "📋 コピー"), 1500);
      });
    });

    // エディタタブ切り替え
    document.querySelectorAll(".editor-tab").forEach((tab) => {
      tab.addEventListener("click", () => {
        const target = tab.dataset.tab;
        document.querySelectorAll(".editor-tab").forEach((t) => t.classList.remove("active"));
        document.querySelectorAll(".tab-panel").forEach((p) => p.classList.remove("active"));
        tab.classList.add("active");
        document.getElementById(`tab-${target}-panel`).classList.add("active");
        // ブロックタブに戻ったときBlocklyをリサイズ
        if (target === "blocks") {
          setTimeout(() => Blockly.svgResize(workspace), 50);
        }
      });
    });

    // レイアウト切り替え（A / B）
    const savedLayout = localStorage.getItem("layout") || "A";
    applyLayout(savedLayout);
    document.querySelectorAll(".layout-btn").forEach((btn) => {
      btn.addEventListener("click", () => {
        applyLayout(btn.dataset.layout);
        localStorage.setItem("layout", btn.dataset.layout);
        setTimeout(() => Blockly.svgResize(workspace), 100);
      });
    });

    // 説明とヒントの開閉（既定: 折りたたみ）
    const collapseBtn = document.getElementById("quest-collapse-btn");
    function applyQuestCollapsed(collapsed) {
      document.getElementById("main-layout").classList.toggle("quest-collapsed", collapsed);
      collapseBtn.textContent = collapsed ? "📖 説明とヒント ▸" : "📖 説明とヒント ▾";
      localStorage.setItem("questCollapsed", collapsed ? "1" : "0");
      setTimeout(() => Blockly.svgResize(workspace), 50);
    }
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
      if (confirm("組み立てたブロックを全部消しますか？")) {
        workspace.clear();
        onWorkspaceChange();
      }
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
          navigator.clipboard.writeText(url).then(() => {
            shareBtn.textContent = "✅ リンクをコピー";
            setTimeout(() => (shareBtn.textContent = "🔗 共有"), 1500);
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
