/* ===== Blockly カスタムブロック定義 =====
 * 学習コア: セル操作・繰り返し・条件分岐・書式設定
 * 各ブロックは vba-generator.js で VBA に変換される
 */

// ----- カテゴリ色 -----
const CAT_COLORS = {
  cell: "#3498db",
  loop: "#e67e22",
  cond: "#9b59b6",
  fmt: "#1abc9c",
  value: "#5c6bc0",
};

// ===== セル操作 =====

Blockly.Blocks["cell_set_value"] = {
  init: function () {
    this.appendValueInput("VALUE")
      .appendField("📝 セル")
      .appendField(new Blockly.FieldTextInput("A1"), "CELL")
      .appendField("に値を入力");
    this.setPreviousStatement(true, null);
    this.setNextStatement(true, null);
    this.setColour(CAT_COLORS.cell);
    this.setTooltip("指定したセルに値を入力します");
  },
};

Blockly.Blocks["cell_get_value"] = {
  init: function () {
    this.appendDummyInput()
      .appendField("🔍 セル")
      .appendField(new Blockly.FieldTextInput("A1"), "CELL")
      .appendField("の値");
    this.setOutput(true, null);
    this.setColour(CAT_COLORS.cell);
    this.setTooltip("指定したセルの値を取得します");
  },
};

Blockly.Blocks["cell_copy"] = {
  init: function () {
    this.appendDummyInput()
      .appendField("📋 セル")
      .appendField(new Blockly.FieldTextInput("A1"), "FROM")
      .appendField("を")
      .appendField(new Blockly.FieldTextInput("B1"), "TO")
      .appendField("にコピー");
    this.setPreviousStatement(true, null);
    this.setNextStatement(true, null);
    this.setColour(CAT_COLORS.cell);
    this.setTooltip("セルの内容を別のセルにコピーします");
  },
};

Blockly.Blocks["cell_clear"] = {
  init: function () {
    this.appendDummyInput()
      .appendField("🗑 セル")
      .appendField(new Blockly.FieldTextInput("A1"), "CELL")
      .appendField("をクリア");
    this.setPreviousStatement(true, null);
    this.setNextStatement(true, null);
    this.setColour(CAT_COLORS.cell);
    this.setTooltip("セルの内容を消去します");
  },
};

// ===== 繰り返し =====

Blockly.Blocks["loop_repeat"] = {
  init: function () {
    this.appendDummyInput()
      .appendField("🔁")
      .appendField(new Blockly.FieldNumber(3, 1, 1000, 1), "TIMES")
      .appendField("回繰り返す");
    this.appendStatementInput("DO").appendField("実行");
    this.setPreviousStatement(true, null);
    this.setNextStatement(true, null);
    this.setColour(CAT_COLORS.loop);
    this.setTooltip("中のブロックを指定回数くり返します。変数 i が 1 から増えます");
  },
};

Blockly.Blocks["loop_range"] = {
  init: function () {
    this.appendDummyInput()
      .appendField("📊 行")
      .appendField(new Blockly.FieldNumber(1, 1, 1000, 1), "START")
      .appendField("から")
      .appendField(new Blockly.FieldNumber(5, 1, 1000, 1), "END")
      .appendField("までループ");
    this.appendStatementInput("DO").appendField("実行");
    this.setPreviousStatement(true, null);
    this.setNextStatement(true, null);
    this.setColour(CAT_COLORS.loop);
    this.setTooltip("変数 i が開始行から終了行まで動きながらくり返します");
  },
};

Blockly.Blocks["loop_index"] = {
  init: function () {
    this.appendDummyInput().appendField("カウンタ i");
    this.setOutput(true, "Number");
    this.setColour(CAT_COLORS.loop);
    this.setTooltip("現在の繰り返し回数 (i) を表します");
  },
};

// ===== 条件分岐 =====

Blockly.Blocks["cond_if"] = {
  init: function () {
    this.appendValueInput("CONDITION").appendField("❓ もし");
    this.appendStatementInput("THEN").appendField("なら");
    this.appendStatementInput("ELSE").appendField("そうでなければ");
    this.setPreviousStatement(true, null);
    this.setNextStatement(true, null);
    this.setColour(CAT_COLORS.cond);
    this.setTooltip("条件によって処理を分けます（そうでなければは空でもOK）");
  },
};

Blockly.Blocks["cond_compare"] = {
  init: function () {
    this.appendValueInput("A");
    this.appendDummyInput().appendField(
      new Blockly.FieldDropdown([
        ["=", "="],
        ["≠", "<>"],
        ["<", "<"],
        ["≤", "<="],
        [">", ">"],
        ["≥", ">="],
      ]),
      "OP"
    );
    this.appendValueInput("B");
    this.setInputsInline(true);
    this.setOutput(true, "Boolean");
    this.setColour(CAT_COLORS.cond);
    this.setTooltip("2つの値を比べます");
  },
};

// ===== 書式設定 =====

Blockly.Blocks["fmt_bgcolor"] = {
  init: function () {
    this.appendDummyInput()
      .appendField("🎨 セル")
      .appendField(new Blockly.FieldTextInput("A1"), "CELL")
      .appendField("の背景色")
      .appendField(
        new Blockly.FieldDropdown([
          ["🟥 赤", "RED"],
          ["🟩 緑", "GREEN"],
          ["🟦 青", "BLUE"],
          ["🟨 黄", "YELLOW"],
          ["⬜ 白", "WHITE"],
        ]),
        "COLOR"
      );
    this.setPreviousStatement(true, null);
    this.setNextStatement(true, null);
    this.setColour(CAT_COLORS.fmt);
    this.setTooltip("セルの背景色を変えます");
  },
};

Blockly.Blocks["fmt_bold"] = {
  init: function () {
    this.appendDummyInput()
      .appendField("𝗕 セル")
      .appendField(new Blockly.FieldTextInput("A1"), "CELL")
      .appendField("を太字に");
    this.setPreviousStatement(true, null);
    this.setNextStatement(true, null);
    this.setColour(CAT_COLORS.fmt);
    this.setTooltip("文字を太字にします");
  },
};

// ===== 値ブロック =====

Blockly.Blocks["value_number"] = {
  init: function () {
    this.appendDummyInput().appendField(new Blockly.FieldNumber(0), "NUM");
    this.setOutput(true, "Number");
    this.setColour(CAT_COLORS.value);
  },
};

Blockly.Blocks["value_text"] = {
  init: function () {
    this.appendDummyInput()
      .appendField('"')
      .appendField(new Blockly.FieldTextInput("テキスト"), "TEXT")
      .appendField('"');
    this.setOutput(true, "String");
    this.setColour(CAT_COLORS.value);
  },
};

Blockly.Blocks["value_math"] = {
  init: function () {
    this.appendValueInput("A");
    this.appendDummyInput().appendField(
      new Blockly.FieldDropdown([
        ["+", "+"],
        ["−", "-"],
        ["×", "*"],
        ["÷", "/"],
      ]),
      "OP"
    );
    this.appendValueInput("B");
    this.setInputsInline(true);
    this.setOutput(true, "Number");
    this.setColour(CAT_COLORS.value);
    this.setTooltip("計算をします");
  },
};

// ===== ツールボックス定義 =====
const TOOLBOX = {
  kind: "categoryToolbox",
  contents: [
    {
      kind: "category",
      name: "セル操作",
      colour: CAT_COLORS.cell,
      contents: [
        { kind: "block", type: "cell_set_value" },
        { kind: "block", type: "cell_get_value" },
        { kind: "block", type: "cell_copy" },
        { kind: "block", type: "cell_clear" },
      ],
    },
    {
      kind: "category",
      name: "繰り返し",
      colour: CAT_COLORS.loop,
      contents: [
        { kind: "block", type: "loop_repeat" },
        { kind: "block", type: "loop_range" },
        { kind: "block", type: "loop_index" },
      ],
    },
    {
      kind: "category",
      name: "条件分岐",
      colour: CAT_COLORS.cond,
      contents: [
        { kind: "block", type: "cond_if" },
        { kind: "block", type: "cond_compare" },
      ],
    },
    {
      kind: "category",
      name: "書式設定",
      colour: CAT_COLORS.fmt,
      contents: [
        { kind: "block", type: "fmt_bgcolor" },
        { kind: "block", type: "fmt_bold" },
      ],
    },
    {
      kind: "category",
      name: "値",
      colour: CAT_COLORS.value,
      contents: [
        { kind: "block", type: "value_number" },
        { kind: "block", type: "value_text" },
        { kind: "block", type: "value_math" },
      ],
    },
  ],
};
