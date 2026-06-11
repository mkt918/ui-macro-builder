/* ===== メインアプリ =====
 * Blockly 初期化・課題切り替え・ヒント・コード生成・Excel プレビューを結線する。
 */

(function () {
  let workspace = null;
  let view = null;
  let currentTaskId = null;
  const savedBlocks = {}; // taskId -> XML文字列（課題ごとのセーブスロット）
  const solved = new Set();

  // ----- テーマ管理 -----
  function initTheme() {
    const saved = localStorage.getItem("theme") || "dark";
    applyTheme(saved);
  }
  function applyTheme(mode) {
    if (mode === "light") {
      document.body.classList.add("light-mode");
      document.getElementById("theme-toggle-btn").textContent = "☀️";
      if (workspace) {
        const lightTheme = Blockly.Theme.defineTheme("light", {
          base: Blockly.Themes.Classic,
          componentStyles: {
            workspaceBackgroundColour: "#fafafa",
            toolboxBackgroundColour: "#ffffff",
            toolboxForegroundColour: "#333333",
            flyoutBackgroundColour: "#f5f5f5",
            flyoutForegroundColour: "#333333",
            scrollbarColour: "#cccccc",
          },
        });
        workspace.setTheme(lightTheme);
      }
    } else {
      document.body.classList.remove("light-mode");
      document.getElementById("theme-toggle-btn").textContent = "🌙";
      if (workspace) {
        const darkTheme = Blockly.Theme.defineTheme("dark", {
          base: Blockly.Themes.Classic,
          componentStyles: {
            workspaceBackgroundColour: "#0d1117",
            toolboxBackgroundColour: "#16213e",
            toolboxForegroundColour: "#e0e0e0",
            flyoutBackgroundColour: "#12182e",
            flyoutForegroundColour: "#e0e0e0",
            scrollbarColour: "#30363d",
          },
        });
        workspace.setTheme(darkTheme);
      }
    }
    localStorage.setItem("theme", mode);
  }

  // ----- 初期化 -----
  window.addEventListener("load", () => {
    initTheme();
    Blockly.setLocale(Blockly.Msg);

    const isDark = !document.body.classList.contains("light-mode");
    const darkTheme = Blockly.Theme.defineTheme("dark", {
      base: Blockly.Themes.Classic,
      componentStyles: {
        workspaceBackgroundColour: "#0d1117",
        toolboxBackgroundColour: "#16213e",
        toolboxForegroundColour: "#e0e0e0",
        flyoutBackgroundColour: "#12182e",
        flyoutForegroundColour: "#e0e0e0",
        scrollbarColour: "#30363d",
      },
    });
    const lightTheme = Blockly.Theme.defineTheme("light", {
      base: Blockly.Themes.Classic,
      componentStyles: {
        workspaceBackgroundColour: "#fafafa",
        toolboxBackgroundColour: "#ffffff",
        toolboxForegroundColour: "#333333",
        flyoutBackgroundColour: "#f5f5f5",
        flyoutForegroundColour: "#333333",
        scrollbarColour: "#cccccc",
      },
    });

    workspace = Blockly.inject("blockly-area", {
      toolbox: TOOLBOX,
      grid: { spacing: 24, length: 3, colour: isDark ? "#222" : "#ddd", snap: true },
      zoom: { controls: true, wheel: true, startScale: 0.95, maxScale: 2, minScale: 0.4 },
      trashcan: true,
      move: { scrollbars: true, drag: true, wheel: true },
      theme: isDark ? darkTheme : lightTheme,
    });

    view = new ExcelView(
      document.getElementById("excel-table"),
      document.getElementById("cell-ref"),
      document.getElementById("formula-bar"),
      document.getElementById("step-status")
    );

    // ブロック変更 → コード再生成 + ステップ再構築
    workspace.addChangeListener(onWorkspaceChange);

    bindControls();
    loadTask(TASKS[0].id);
    buildQuestModal();
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
      document.getElementById("current-quest").textContent = `📍 ${task.title} ${"★".repeat(task.difficulty)}`;
    }
  }

  // ----- 課題読み込み -----
  function loadTask(taskId) {
    // 現在のブロックを保存
    if (currentTaskId && workspace) {
      saveCurrentBlocks();
    }
    currentTaskId = taskId;
    const task = TASKS.find((t) => t.id === taskId);
    if (!task) return;

    // 課題説明
    document.getElementById("task-title").textContent = task.title;
    document.getElementById("task-difficulty").textContent = "★".repeat(task.difficulty) +
      "☆".repeat(5 - task.difficulty);
    document.getElementById("task-goal").textContent = task.goal;

    // ヒントUIリセット
    resetHints(task);

    // ブロック復元（無ければ真っ白）
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

    updateCurrentQuestDisplay();
    onWorkspaceChange();
  }

  function saveCurrentBlocks() {
    try {
      const dom = Blockly.Xml.workspaceToDom(workspace);
      savedBlocks[currentTaskId] = Blockly.Xml.domToText(dom);
    } catch (e) {
      console.warn("ブロック保存失敗:", e);
    }
  }

  // ----- ワークスペース変更時 -----
  function onWorkspaceChange(event) {
    if (event && event.isUiEvent) return;

    // VBA コード生成
    const code = generateVBA(workspace);
    document.querySelector("#code-output code").textContent = code;

    // エラーチェック（未接続の値ブロックなど簡易検出）
    checkErrors();

    // ステップ再構築
    try {
      const steps = buildSteps(workspace);
      view.load(steps);
    } catch (e) {
      console.warn("ステップ生成エラー:", e);
    }

    if (currentTaskId) saveCurrentBlocks();
  }

  // ----- 簡易エラーチェック -----
  function checkErrors() {
    const banner = document.getElementById("error-banner");
    const orphans = workspace
      .getAllBlocks(false)
      .filter((b) => b.outputConnection && !b.outputConnection.isConnected() && !b.isInFlyout);
    if (orphans.length > 0) {
      banner.textContent =
        "⚠️ つながっていない値ブロックがあります。文ブロックの差込口につなげましょう。";
      banner.classList.add("show");
    } else {
      banner.classList.remove("show");
    }
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
    document.getElementById("run-btn").addEventListener("click", () => view.play());
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
  }

  function escapeHtml(s) {
    return s.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
  }
})();
