/* ===== Blockly カスタムブロック定義 =====
 * 小学生でもわかるブロック名・説明
 * 各ブロックは vba-generator.js で VBA に変換される
 */

const CAT_COLORS = {
  cell:  "#2980b9",
  loop:  "#e67e22",
  cond:  "#8e44ad",
  fmt:   "#16a085",
  array: "#27ae60",
  sheet: "#c0392b",
  value: "#2c3e50",
  var:   "#d35400",
};

// ===== セル操作 =====

// セルに値を入れる
Blockly.Blocks["cell_set_value"] = {
  init: function () {
    this.appendValueInput("VALUE")
      .appendField("📝 セル")
      .appendField(new Blockly.FieldTextInput("A1"), "CELL")
      .appendField("の値を");
    this.appendDummyInput().appendField("に変える");
    this.setInputsInline(true);
    this.setPreviousStatement(true, null);
    this.setNextStatement(true, null);
    this.setColour(CAT_COLORS.cell);
    this.setTooltip("セルの中に文字や数字を入れます");
  },
};

// セルの値を読む
Blockly.Blocks["cell_get_value"] = {
  init: function () {
    this.appendDummyInput()
      .appendField("🔍 セル")
      .appendField(new Blockly.FieldTextInput("A1"), "CELL")
      .appendField("の中身");
    this.setOutput(true, null);
    this.setColour(CAT_COLORS.cell);
    this.setTooltip("セルに入っている値を取り出します");
  },
};

// セルをコピー
Blockly.Blocks["cell_copy"] = {
  init: function () {
    this.appendDummyInput()
      .appendField("📋 セル")
      .appendField(new Blockly.FieldTextInput("A1"), "FROM")
      .appendField("の内容を")
      .appendField(new Blockly.FieldTextInput("B1"), "TO")
      .appendField("にコピーする");
    this.setPreviousStatement(true, null);
    this.setNextStatement(true, null);
    this.setColour(CAT_COLORS.cell);
    this.setTooltip("セルの内容を別のセルにコピーします");
  },
};

// セルを消す
Blockly.Blocks["cell_clear"] = {
  init: function () {
    this.appendDummyInput()
      .appendField("🗑 セル")
      .appendField(new Blockly.FieldTextInput("A1"), "CELL")
      .appendField("の中身を消す");
    this.setPreviousStatement(true, null);
    this.setNextStatement(true, null);
    this.setColour(CAT_COLORS.cell);
    this.setTooltip("セルの中身を空にします");
  },
};

// 行全体を消す（新規）
Blockly.Blocks["cell_clear_row"] = {
  init: function () {
    this.appendDummyInput()
      .appendField("🗑")
      .appendField(new Blockly.FieldNumber(1, 1, 1000, 1), "ROW")
      .appendField("行目をすべて消す");
    this.setPreviousStatement(true, null);
    this.setNextStatement(true, null);
    this.setColour(CAT_COLORS.cell);
    this.setTooltip("指定した行のセルをすべて空にします");
  },
};

// ===== 繰り返し =====

Blockly.Blocks["loop_repeat"] = {
  init: function () {
    this.appendDummyInput()
      .appendField("🔁")
      .appendField(new Blockly.FieldNumber(3, 1, 1000, 1), "TIMES")
      .appendField("回くりかえす");
    this.appendStatementInput("DO").appendField("↓ やること");
    this.setPreviousStatement(true, null);
    this.setNextStatement(true, null);
    this.setColour(CAT_COLORS.loop);
    this.setTooltip("中のブロックを指定回数くり返します。何回目かは「かうんた」で使えます");
  },
};

Blockly.Blocks["loop_range"] = {
  init: function () {
    this.appendDummyInput()
      .appendField("🔁")
      .appendField(new Blockly.FieldNumber(1, 1, 1000, 1), "START")
      .appendField("から")
      .appendField(new Blockly.FieldNumber(5, 1, 1000, 1), "END")
      .appendField("まで1ずつくりかえす");
    this.appendStatementInput("DO").appendField("↓ やること");
    this.setPreviousStatement(true, null);
    this.setNextStatement(true, null);
    this.setColour(CAT_COLORS.loop);
    this.setTooltip("開始から終了まで1ずつ増えながらくり返します");
  },
};

Blockly.Blocks["loop_for_step"] = {
  init: function () {
    this.appendDummyInput()
      .appendField("🔁 For")
      .appendField(new Blockly.FieldNumber(1, -1000, 1000, 1), "START")
      .appendField("から")
      .appendField(new Blockly.FieldNumber(10, -1000, 1000, 1), "END")
      .appendField("まで")
      .appendField(new Blockly.FieldNumber(2, -1000, 1000, 1), "STEP")
      .appendField("ずつくりかえす");
    this.appendStatementInput("DO").appendField("↓ やること");
    this.setPreviousStatement(true, null);
    this.setNextStatement(true, null);
    this.setColour(CAT_COLORS.loop);
    this.setTooltip("開始から終了まで指定した数ずつ増やしながらくり返します（負の数で減らすことも可）");
  },
};

Blockly.Blocks["loop_while"] = {
  init: function () {
    this.appendValueInput("CONDITION").appendField("🔁 While（");
    this.appendDummyInput().appendField("の間くりかえす）");
    this.appendStatementInput("DO").appendField("↓ やること");
    this.setInputsInline(false);
    this.setPreviousStatement(true, null);
    this.setNextStatement(true, null);
    this.setColour(CAT_COLORS.loop);
    this.setTooltip("条件が正しい間ずっとくり返します。条件が False になると止まります");
  },
};

Blockly.Blocks["loop_index"] = {
  init: function () {
    this.appendDummyInput().appendField("🔢 かうんた（今何回目？）");
    this.setOutput(true, "Number");
    this.setColour(CAT_COLORS.loop);
    this.setTooltip("今何回目か（いまの番号）を表します");
  },
};

// ===== 条件分岐 =====

Blockly.Blocks["cond_if"] = {
  init: function () {
    this.appendValueInput("CONDITION").appendField("❓ もし");
    this.appendStatementInput("THEN").appendField("なら → やること");
    this.appendStatementInput("ELSE").appendField("ちがうなら → やること");
    this.setPreviousStatement(true, null);
    this.setNextStatement(true, null);
    this.setColour(CAT_COLORS.cond);
    this.setTooltip("条件が合うかどうかで処理を分けます。「ちがうなら」は空でもOK");
  },
};

Blockly.Blocks["cond_compare"] = {
  init: function () {
    this.appendValueInput("A");
    this.appendDummyInput().appendField(
      new Blockly.FieldDropdown([
        ["＝ ひとしい", "="],
        ["≠ ちがう", "<>"],
        ["＜ より小さい", "<"],
        ["≤ 以下", "<="],
        ["＞ より大きい", ">"],
        ["≥ 以上", ">="],
      ]),
      "OP"
    );
    this.appendValueInput("B");
    this.setInputsInline(true);
    this.setOutput(true, "Boolean");
    this.setColour(CAT_COLORS.cond);
    this.setTooltip("2つの数字や文字を比べます");
  },
};

// 偶数判定（新規）
Blockly.Blocks["cond_is_even"] = {
  init: function () {
    this.appendValueInput("NUM").appendField("🔢");
    this.appendDummyInput().appendField("は偶数？");
    this.setInputsInline(true);
    this.setOutput(true, "Boolean");
    this.setColour(CAT_COLORS.cond);
    this.setTooltip("数が偶数（2で割り切れる）かどうかを調べます");
  },
};

// ===== 書式設定 =====

Blockly.Blocks["fmt_bgcolor"] = {
  init: function () {
    this.appendDummyInput()
      .appendField("🎨 セル")
      .appendField(new Blockly.FieldTextInput("A1"), "CELL")
      .appendField("の背景色を")
      .appendField(
        new Blockly.FieldDropdown([
          ["🟥 赤", "RED"],
          ["🟩 緑", "GREEN"],
          ["🟦 青", "BLUE"],
          ["🟨 黄", "YELLOW"],
          ["⬜ 白（消す）", "WHITE"],
        ]),
        "COLOR"
      )
      .appendField("にする");
    this.setPreviousStatement(true, null);
    this.setNextStatement(true, null);
    this.setColour(CAT_COLORS.fmt);
    this.setTooltip("セルの背景の色を変えます");
  },
};

Blockly.Blocks["fmt_bold"] = {
  init: function () {
    this.appendDummyInput()
      .appendField("𝗕 セル")
      .appendField(new Blockly.FieldTextInput("A1"), "CELL")
      .appendField("の文字を太字にする");
    this.setPreviousStatement(true, null);
    this.setNextStatement(true, null);
    this.setColour(CAT_COLORS.fmt);
    this.setTooltip("セルの文字を太字（ふとじ）にします");
  },
};

// 文字サイズ変更（新規）
Blockly.Blocks["fmt_fontsize"] = {
  init: function () {
    this.appendDummyInput()
      .appendField("🔡 セル")
      .appendField(new Blockly.FieldTextInput("A1"), "CELL")
      .appendField("の文字サイズを")
      .appendField(new Blockly.FieldNumber(14, 6, 72, 1), "SIZE")
      .appendField("にする");
    this.setPreviousStatement(true, null);
    this.setNextStatement(true, null);
    this.setColour(CAT_COLORS.fmt);
    this.setTooltip("セルの文字の大きさを変えます");
  },
};

// ===== 配列（リスト）=====

Blockly.Blocks["array_set"] = {
  init: function () {
    this.appendValueInput("INDEX").appendField("📦 リストの");
    this.appendValueInput("VALUE").appendField("番目に");
    this.appendDummyInput().appendField("を入れる");
    this.setInputsInline(true);
    this.setPreviousStatement(true, null);
    this.setNextStatement(true, null);
    this.setColour(CAT_COLORS.array);
    this.setTooltip("リスト（箱の列）の指定した番目に値を入れます");
  },
};

Blockly.Blocks["array_get"] = {
  init: function () {
    this.appendValueInput("INDEX").appendField("📦 リストの");
    this.appendDummyInput().appendField("番目の中身");
    this.setInputsInline(true);
    this.setOutput(true, null);
    this.setColour(CAT_COLORS.array);
    this.setTooltip("リストの指定した番目の値を取り出します");
  },
};

// ===== シート操作 =====

Blockly.Blocks["sheet_add"] = {
  init: function () {
    this.appendDummyInput()
      .appendField("➕ シート")
      .appendField(new Blockly.FieldTextInput("Sheet2"), "NAME")
      .appendField("を追加する");
    this.setPreviousStatement(true, null);
    this.setNextStatement(true, null);
    this.setColour(CAT_COLORS.sheet);
    this.setTooltip("新しいシート（ページ）を追加します");
  },
};

Blockly.Blocks["sheet_select"] = {
  init: function () {
    this.appendDummyInput()
      .appendField("↩ シート")
      .appendField(new Blockly.FieldTextInput("Sheet1"), "NAME")
      .appendField("に切り替える");
    this.setPreviousStatement(true, null);
    this.setNextStatement(true, null);
    this.setColour(CAT_COLORS.sheet);
    this.setTooltip("指定したシートに切り替えます");
  },
};

// ===== Cells（行・列を変数で指定）=====

Blockly.Blocks["cells_set_value"] = {
  init: function () {
    this.appendValueInput("ROW")
      .appendField("📝 Cells（行");
    this.appendValueInput("COL")
      .appendField("列");
    this.appendValueInput("VALUE")
      .appendField("）の値を");
    this.appendDummyInput().appendField("に変える");
    this.setInputsInline(true);
    this.setPreviousStatement(true, null);
    this.setNextStatement(true, null);
    this.setColour(CAT_COLORS.cell);
    this.setTooltip("行番号と列番号で場所を指定してセルに値を入れます（変数も使えます）");
  },
};

Blockly.Blocks["cells_get_value"] = {
  init: function () {
    this.appendValueInput("ROW")
      .appendField("🔍 Cells（行");
    this.appendValueInput("COL")
      .appendField("列");
    this.appendDummyInput().appendField("）の中身");
    this.setInputsInline(true);
    this.setOutput(true, null);
    this.setColour(CAT_COLORS.cell);
    this.setTooltip("行番号と列番号でセルを指定して値を取り出します");
  },
};

Blockly.Blocks["cells_clear"] = {
  init: function () {
    this.appendValueInput("ROW")
      .appendField("🗑 Cells（行");
    this.appendValueInput("COL")
      .appendField("列");
    this.appendDummyInput().appendField("）の中身を消す");
    this.setInputsInline(true);
    this.setPreviousStatement(true, null);
    this.setNextStatement(true, null);
    this.setColour(CAT_COLORS.cell);
    this.setTooltip("行番号と列番号でセルを指定して中身を消します");
  },
};

Blockly.Blocks["cells_bgcolor"] = {
  init: function () {
    this.appendValueInput("ROW")
      .appendField("🎨 Cells（行");
    this.appendValueInput("COL")
      .appendField("列");
    this.appendDummyInput()
      .appendField("）の背景色を")
      .appendField(
        new Blockly.FieldDropdown([
          ["🟥 赤", "RED"],
          ["🟩 緑", "GREEN"],
          ["🟦 青", "BLUE"],
          ["🟨 黄", "YELLOW"],
          ["⬜ 白（消す）", "WHITE"],
        ]),
        "COLOR"
      )
      .appendField("にする");
    this.setInputsInline(true);
    this.setPreviousStatement(true, null);
    this.setNextStatement(true, null);
    this.setColour(CAT_COLORS.fmt);
    this.setTooltip("行番号と列番号でセルを指定して背景色を変えます");
  },
};

Blockly.Blocks["cells_bold"] = {
  init: function () {
    this.appendValueInput("ROW")
      .appendField("𝗕 Cells（行");
    this.appendValueInput("COL")
      .appendField("列");
    this.appendDummyInput().appendField("）を太字にする");
    this.setInputsInline(true);
    this.setPreviousStatement(true, null);
    this.setNextStatement(true, null);
    this.setColour(CAT_COLORS.fmt);
    this.setTooltip("行番号と列番号でセルを指定して太字にします");
  },
};

// ===== 変数 =====

Blockly.Blocks["var_set"] = {
  init: function () {
    this.appendValueInput("VALUE")
      .appendField("📦 変数")
      .appendField(new Blockly.FieldTextInput("ごうけい"), "NAME")
      .appendField("に");
    this.appendDummyInput().appendField("を入れる");
    this.setInputsInline(true);
    this.setPreviousStatement(true, null);
    this.setNextStatement(true, null);
    this.setColour(CAT_COLORS.var);
    this.setTooltip("変数（名前のついた箱）に値を入れます");
  },
};

Blockly.Blocks["var_get"] = {
  init: function () {
    this.appendDummyInput()
      .appendField("📦 変数")
      .appendField(new Blockly.FieldTextInput("ごうけい"), "NAME")
      .appendField("の中身");
    this.setOutput(true, null);
    this.setColour(CAT_COLORS.var);
    this.setTooltip("変数に入っている値を取り出します");
  },
};

// ===== 値ブロック =====

Blockly.Blocks["value_number"] = {
  init: function () {
    this.appendDummyInput()
      .appendField("🔢")
      .appendField(new Blockly.FieldNumber(0), "NUM");
    this.setOutput(true, "Number");
    this.setColour(CAT_COLORS.value);
    this.setTooltip("数字を表すブロックです");
  },
};

Blockly.Blocks["value_text"] = {
  init: function () {
    this.appendDummyInput()
      .appendField("🔤「")
      .appendField(new Blockly.FieldTextInput("テキスト"), "TEXT")
      .appendField("」");
    this.setOutput(true, "String");
    this.setColour(CAT_COLORS.value);
    this.setTooltip("文字（テキスト）を表すブロックです");
  },
};

Blockly.Blocks["value_math"] = {
  init: function () {
    this.appendValueInput("A");
    this.appendDummyInput().appendField(
      new Blockly.FieldDropdown([
        ["＋ たす", "+"],
        ["－ ひく", "-"],
        ["× かける", "*"],
        ["÷ わる", "/"],
        ["÷のあまり", "mod"],
      ]),
      "OP"
    );
    this.appendValueInput("B");
    this.setInputsInline(true);
    this.setOutput(true, "Number");
    this.setColour(CAT_COLORS.value);
    this.setTooltip("計算をします。「÷のあまり」は偶数・奇数の判定に使えます");
  },
};

// ===== ツールボックス定義 =====
const TOOLBOX = {
  kind: "categoryToolbox",
  contents: [
    {
      kind: "category",
      name: "📝 セル操作",
      colour: CAT_COLORS.cell,
      contents: [
        { kind: "block", type: "cell_set_value" },
        { kind: "block", type: "cell_get_value" },
        { kind: "block", type: "cell_copy" },
        { kind: "block", type: "cell_clear" },
        { kind: "block", type: "cell_clear_row" },
        { kind: "block", type: "cells_set_value" },
        { kind: "block", type: "cells_get_value" },
        { kind: "block", type: "cells_clear" },
      ],
    },
    {
      kind: "category",
      name: "🔁 くりかえし",
      colour: CAT_COLORS.loop,
      contents: [
        { kind: "block", type: "loop_repeat" },
        { kind: "block", type: "loop_range" },
        { kind: "block", type: "loop_for_step" },
        { kind: "block", type: "loop_while" },
        { kind: "block", type: "loop_index" },
      ],
    },
    {
      kind: "category",
      name: "❓ もし〜なら",
      colour: CAT_COLORS.cond,
      contents: [
        { kind: "block", type: "cond_if" },
        { kind: "block", type: "cond_compare" },
        { kind: "block", type: "cond_is_even" },
      ],
    },
    {
      kind: "category",
      name: "🎨 色・文字デザイン",
      colour: CAT_COLORS.fmt,
      contents: [
        { kind: "block", type: "fmt_bgcolor" },
        { kind: "block", type: "fmt_bold" },
        { kind: "block", type: "fmt_fontsize" },
        { kind: "block", type: "cells_bgcolor" },
        { kind: "block", type: "cells_bold" },
      ],
    },
    {
      kind: "category",
      name: "📦 リスト（配列）",
      colour: CAT_COLORS.array,
      contents: [
        { kind: "block", type: "array_set" },
        { kind: "block", type: "array_get" },
      ],
    },
    {
      kind: "category",
      name: "📑 シート操作",
      colour: CAT_COLORS.sheet,
      contents: [
        { kind: "block", type: "sheet_add" },
        { kind: "block", type: "sheet_select" },
      ],
    },
    {
      kind: "category",
      name: "📦 変数",
      colour: CAT_COLORS.var,
      contents: [
        { kind: "block", type: "var_set" },
        { kind: "block", type: "var_get" },
      ],
    },
    {
      kind: "category",
      name: "🔢 数・文字",
      colour: CAT_COLORS.value,
      contents: [
        { kind: "block", type: "value_number" },
        { kind: "block", type: "value_text" },
        { kind: "block", type: "value_math" },
      ],
    },
  ],
};
