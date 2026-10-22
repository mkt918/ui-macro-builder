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
    this.appendValueInput("ROW")
      .setCheck("Number")
      .appendField("🗑");
    this.appendDummyInput().appendField("行目をすべて消す");
    this.setInputsInline(true);
    this.setPreviousStatement(true, null);
    this.setNextStatement(true, null);
    this.setColour(CAT_COLORS.cell);
    this.setTooltip("指定した行のセルをすべて空にします");
  },
};

// ===== 繰り返し =====

Blockly.Blocks["loop_repeat"] = {
  init: function () {
    this.appendValueInput("TIMES")
      .setCheck("Number")
      .appendField("🔁");
    this.appendDummyInput().appendField("回くりかえす");
    this.appendStatementInput("DO").appendField("↓ やること");
    this.setInputsInline(true);
    this.setPreviousStatement(true, null);
    this.setNextStatement(true, null);
    this.setColour(CAT_COLORS.loop);
    this.setTooltip("中のブロックを指定回数くり返します。何回目かは「かうんた」で使えます");
  },
};

Blockly.Blocks["loop_range"] = {
  init: function () {
    this.appendValueInput("START")
      .setCheck("Number")
      .appendField("🔁");
    this.appendValueInput("END")
      .setCheck("Number")
      .appendField("から");
    this.appendDummyInput().appendField("まで1ずつくりかえす");
    this.appendStatementInput("DO").appendField("↓ やること");
    this.setInputsInline(true);
    this.setPreviousStatement(true, null);
    this.setNextStatement(true, null);
    this.setColour(CAT_COLORS.loop);
    this.setTooltip("開始から終了まで1ずつ増えながらくり返します");
  },
};

Blockly.Blocks["loop_for_step"] = {
  init: function () {
    this.appendValueInput("START")
      .setCheck("Number")
      .appendField("🔁 For");
    this.appendValueInput("END")
      .setCheck("Number")
      .appendField("から");
    this.appendValueInput("STEP")
      .setCheck("Number")
      .appendField("まで");
    this.appendDummyInput().appendField("ずつくりかえす");
    this.appendStatementInput("DO").appendField("↓ やること");
    this.setInputsInline(true);
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

// 「＋ / −」ボタンで項目（ElseIf・Case）を増やし減らせるブロックの共通部品。
// 初学者にも使えるよう、マウスでの小さな「ミューテーター窓」ではなく、ブロックの上のボタンで増減する
const BTN_SVG = (sign) =>
  "data:image/svg+xml;utf8," +
  encodeURIComponent(
    `<svg xmlns="http://www.w3.org/2000/svg" width="22" height="22"><circle cx="11" cy="11" r="10" fill="#ffffff" stroke="#444" stroke-width="1.5"/><path d="${
      sign === "+" ? "M6 11h10M11 6v10" : "M6 11h10"
    }" stroke="#222" stroke-width="2.2" stroke-linecap="round"/></svg>`
  );
const COUNT_MUTATION = {
  count_: 0,
  mutationToDom: function () {
    if (!this.count_) return null;
    const e = Blockly.utils.xml.createElement("mutation");
    e.setAttribute("count", String(this.count_));
    return e;
  },
  domToMutation: function (e) {
    this.setCount_(Number(e.getAttribute("count")) || 0);
  },
  saveExtraState: function () {
    return this.count_ ? { count: this.count_ } : null;
  },
  loadExtraState: function (s) {
    this.setCount_((s && s.count) || 0);
  },
  setCount_: function (n) {
    while (this.count_ < n) {
      this.addItem_(this.count_);
      this.count_++;
    }
    while (this.count_ > n) {
      this.count_--;
      this.removeItem_(this.count_);
    }
  },
  // ボタンで1つ増やす・減らす（元に戻すで戻せるよう、変更のイベントを記録する）
  changeCount_: function (delta) {
    if (this.isInFlyout) return;
    const n = Math.max(0, Math.min(12, this.count_ + delta));
    if (n === this.count_) return;
    const text = () => {
      const m = this.mutationToDom();
      return Blockly.Xml.domToText(m || Blockly.utils.xml.createElement("mutation"));
    };
    const before = text();
    Blockly.Events.setGroup(true);
    this.setCount_(n);
    Blockly.Events.fire(new Blockly.Events.BlockChange(this, "mutation", null, before, text()));
    Blockly.Events.setGroup(false);
  },
  appendCountButtons_: function (label) {
    this.appendDummyInput("BTNS")
      .appendField(new Blockly.FieldImage(BTN_SVG("+"), 22, 22, "＋", () => this.changeCount_(1)))
      .appendField(new Blockly.FieldImage(BTN_SVG("-"), 22, 22, "－", () => this.changeCount_(-1)))
      .appendField(label);
  },
};

Blockly.Blocks["cond_if"] = Object.assign({}, COUNT_MUTATION, {
  init: function () {
    this.count_ = 0;
    this.appendValueInput("CONDITION").appendField("❓ もし");
    this.appendStatementInput("THEN").appendField("なら → やること");
    this.appendStatementInput("ELSE").appendField("ぜんぶ ちがうなら → やること");
    this.appendCountButtons_("「ちがうけど もし」を ふやす／へらす（ElseIf）");
    this.setPreviousStatement(true, null);
    this.setNextStatement(true, null);
    this.setColour(CAT_COLORS.cond);
    this.setTooltip("条件が合うかどうかで処理を分けます。「ちがうなら」は空でもOK。＋を押すと「ちがうけど もし〜なら」（ElseIf）を足せます");
  },
  addItem_: function (i) {
    this.appendValueInput("ELSEIF_COND_" + i).appendField("ちがうけど もし");
    this.appendStatementInput("ELSEIF_DO_" + i).appendField("なら → やること");
    this.moveInputBefore("ELSEIF_COND_" + i, "ELSE");
    this.moveInputBefore("ELSEIF_DO_" + i, "ELSE");
  },
  removeItem_: function (i) {
    this.removeInput("ELSEIF_COND_" + i);
    this.removeInput("ELSEIF_DO_" + i);
  },
});

// 場合分け（Select Case）：「＋」で場合（Case）を足す
Blockly.Blocks["cond_select"] = Object.assign({}, COUNT_MUTATION, {
  init: function () {
    this.count_ = 0;
    this.appendValueInput("SUBJECT").appendField("🔀 場合分け Select Case");
    this.appendStatementInput("ELSE").appendField("どれでもないとき（Case Else） →");
    this.appendCountButtons_("「場合（Case）」を ふやす／へらす");
    this.setPreviousStatement(true, null);
    this.setNextStatement(true, null);
    this.setColour(CAT_COLORS.cond);
    this.setTooltip("1つの値で、いくつかに場合分けします。場合の欄には「1」「1, 2, 3」「60 To 79」「Is >= 80」のように書きます");
    this.addItem_(0);
    this.count_ = 1;
  },
  addItem_: function (i) {
    this.appendStatementInput("CASE_DO_" + i)
      .appendField("場合（Case）")
      .appendField(new Blockly.FieldTextInput("1"), "CASE_" + i)
      .appendField("のとき →");
    this.moveInputBefore("CASE_DO_" + i, "ELSE");
  },
  removeItem_: function (i) {
    this.removeInput("CASE_DO_" + i);
  },
  mutationToDom: function () {
    const e = Blockly.utils.xml.createElement("mutation");
    e.setAttribute("count", String(this.count_));
    return e;
  },
  saveExtraState: function () {
    return { count: this.count_ };
  },
  loadExtraState: function (s) {
    this.setCount_(s && s.count !== undefined ? s.count : 1);
  },
  domToMutation: function (e) {
    const c = e.getAttribute("count");
    this.setCount_(c === null ? 1 : Number(c) || 0);
  },
});

// かつ・または（And / Or）
Blockly.Blocks["cond_logic"] = {
  init: function () {
    this.appendValueInput("A");
    this.appendDummyInput().appendField(
      new Blockly.FieldDropdown([
        ["かつ（And）", "And"],
        ["または（Or）", "Or"],
      ]),
      "OP"
    );
    this.appendValueInput("B");
    this.setInputsInline(true);
    this.setOutput(true, "Boolean");
    this.setColour(CAT_COLORS.cond);
    this.setTooltip("かつ（And）：両方とも合っているとき。または（Or）：どちらか合っているとき");
  },
};

Blockly.Blocks["cond_not"] = {
  init: function () {
    this.appendValueInput("A").appendField("🚫 ではない（Not）");
    this.setInputsInline(true);
    this.setOutput(true, "Boolean");
    this.setColour(CAT_COLORS.cond);
    this.setTooltip("条件の合う・合わないを、ぎゃくにします");
  },
};

Blockly.Blocks["value_bool"] = {
  init: function () {
    this.appendDummyInput().appendField(
      new Blockly.FieldDropdown([
        ["✅ 正しい（True）", "True"],
        ["❌ まちがい（False）", "False"],
      ]),
      "BOOL"
    );
    this.setOutput(true, "Boolean");
    this.setColour(CAT_COLORS.cond);
    this.setTooltip("正しい（True）かまちがい（False）かを表す値です");
  },
};

Blockly.Blocks["prog_exit_sub"] = {
  init: function () {
    this.appendDummyInput().appendField("↩ ここで命令を終わる（Exit Sub）");
    this.setPreviousStatement(true, null);
    this.setNextStatement(true, null);
    this.setColour(CAT_COLORS.cond);
    this.setTooltip("この命令のまとまり（Sub）を、ここで終わりにします。もし〜なら の中でよく使います");
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
      .appendField("の文字サイズを");
    this.appendValueInput("SIZE").setCheck("Number");
    this.appendDummyInput().appendField("にする");
    this.setInputsInline(true);
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

// ===== 入出力 =====

Blockly.Blocks["io_inputbox"] = {
  init: function () {
    this.appendDummyInput()
      .appendField("💬「")
      .appendField(new Blockly.FieldTextInput("行数を入れてね"), "PROMPT")
      .appendField("」と聞いて答えをもらう　窓の題名")
      .appendField(new Blockly.FieldTextInput(""), "TITLE")
      .appendField("はじめの答え")
      .appendField(new Blockly.FieldTextInput(""), "DEFAULT");
    this.setOutput(true, null);
    this.setColour("#e91e63");
    this.setTooltip("ユーザーに文字や数字を入力してもらいます。答えは「文字」として返るので、計算するときは Val(…) で数にしよう。「窓の題名」「はじめの答え」は空でもOK（VBA の2番目・3番目の引数）");
  },
};

Blockly.Blocks["io_msgbox"] = {
  init: function () {
    this.appendValueInput("MSG").appendField("📢 表示する");
    this.setPreviousStatement(true, null);
    this.setNextStatement(true, null);
    this.setColour("#e91e63");
    this.setTooltip("ポップアップでメッセージを表示します");
  },
};

Blockly.Blocks["text_concat"] = {
  init: function () {
    this.appendValueInput("A").appendField("🔗");
    this.appendValueInput("B").appendField("と");
    this.appendDummyInput().appendField("をつなげる");
    this.setInputsInline(true);
    this.setOutput(true, null);
    this.setColour("#e91e63");
    this.setTooltip("2つの値を文字としてつなげます。入れ子にして3つ以上もつなげられます");
  },
};

Blockly.Blocks["value_date"] = {
  init: function () {
    this.appendDummyInput().appendField("📅 今日の日付");
    this.setOutput(true, null);
    this.setColour("#e91e63");
    this.setTooltip("今日の日付を返します");
  },
};

Blockly.Blocks["math_round"] = {
  init: function () {
    this.appendValueInput("VALUE").appendField("🔢 四捨五入（");
    this.appendValueInput("DIGITS").setCheck("Number").appendField("小数点");
    this.appendDummyInput().appendField("桁）");
    this.setInputsInline(true);
    this.setOutput(true, "Number");
    this.setColour("#e91e63");
    this.setTooltip("数を指定した小数点の桁で四捨五入します");
  },
};

// ===== 範囲・罫線 =====

Blockly.Blocks["loop_do_until"] = {
  init: function () {
    this.appendValueInput("CONDITION").appendField("🔁 くりかえす ⛔止まる条件：");
    this.appendStatementInput("DO").appendField("↓ やること");
    this.setInputsInline(false);
    this.setPreviousStatement(true, null);
    this.setNextStatement(true, null);
    this.setColour(CAT_COLORS.loop);
    this.setTooltip("条件が正しくなったら止まります。条件が False の間くり返し続けます");
  },
};

Blockly.Blocks["range_border"] = {
  init: function () {
    this.appendValueInput("R1").setCheck("Number").appendField("📏 行");
    this.appendValueInput("C1").setCheck("Number").appendField("列");
    this.appendValueInput("R2").setCheck("Number").appendField("〜 行");
    this.appendValueInput("C2").setCheck("Number").appendField("列");
    this.appendDummyInput().appendField("に罫線を引く");
    this.setInputsInline(true);
    this.setPreviousStatement(true, null);
    this.setNextStatement(true, null);
    this.setColour(CAT_COLORS.cell);
    this.setTooltip("指定した範囲すべてに枠線（罫線）を引きます");
  },
};

Blockly.Blocks["range_select"] = {
  init: function () {
    this.appendValueInput("R1").setCheck("Number").appendField("🖱 行");
    this.appendValueInput("C1").setCheck("Number").appendField("列");
    this.appendValueInput("R2").setCheck("Number").appendField("〜 行");
    this.appendValueInput("C2").setCheck("Number").appendField("列");
    this.appendDummyInput().appendField("を選ぶ");
    this.setInputsInline(true);
    this.setPreviousStatement(true, null);
    this.setNextStatement(true, null);
    this.setColour(CAT_COLORS.cell);
    this.setTooltip("指定した範囲を選択状態にします");
  },
};

// ===== 変数（ドロップダウン選択式）=====

// 予約名（ループカウンタ i / 配列 arr と同じ名前だと、VBAコード生成時に
// Dim の二重宣言でコンパイルエラーになる）。大文字小文字は区別しない。
const RESERVED_VAR_NAMES = new Set(["i", "j", "k", "arr"]);

// 変数名として使えるかチェックする。OKなら null、NGなら理由（日本語メッセージ）を返す。
// FieldVariable のフィールド検証（doClassValidation_）と、
// 「＋ 変数を作る」ボタンの両方から共通で使う（app.js から window.UMB_validateVarName で参照）。
function validateVarNameMessage(name) {
  if (!name || !name.trim()) return "変数名を入力してください";
  if (/[぀-ゟ]/.test(name)) return `変数名にひらがなは使えません: ${name}`;
  if (RESERVED_VAR_NAMES.has(name.toLowerCase())) {
    return `「${name}」はくり返しのカウンタ(i・j・k)や配列(arr)と同じ名前のため使えません`;
  }
  return null;
}
window.UMB_validateVarName = validateVarNameMessage;

// FieldVariable のサブクラス（ひらがな・予約名を拒否）
// 注意: これは「既存フィールドの値を書き換える」経路（ドロップダウンでの
// 名前変更など）だけを検証する。「＋ 変数を作る」ボタンからの新規作成は
// この検証を経由しないため、app.js 側の CREATE_VARIABLE_JP ボタン
// コールバックでも同じ validateVarNameMessage を使ってチェックしている。
class FieldVariableNoHiragana extends Blockly.FieldVariable {
  constructor(varName = "Gokei") {
    super(varName);
  }

  doClassValidation_(newValue) {
    if (validateVarNameMessage(newValue)) return null;
    return newValue;
  }
}

// 変数に値を入れる
Blockly.Blocks["var_set"] = {
  init: function () {
    this.appendValueInput("VALUE")
      .appendField("📦 変数")
      .appendField(new FieldVariableNoHiragana("Gokei"), "VAR")
      .appendField("に");
    this.appendDummyInput().appendField("を入れる");
    this.setInputsInline(true);
    this.setPreviousStatement(true, null);
    this.setNextStatement(true, null);
    this.setColour(CAT_COLORS.var);
    this.setTooltip("変数（名前のついた箱）に値を入れます");
  },
};

// 変数の中身を取り出す
Blockly.Blocks["var_get"] = {
  init: function () {
    this.appendDummyInput()
      .appendField("📦 変数")
      .appendField(new FieldVariableNoHiragana("Gokei"), "VAR")
      .appendField("の中身");
    this.setOutput(true, null);
    this.setColour(CAT_COLORS.var);
    this.setTooltip("変数に入っている値を取り出します");
  },
};

// 変数を◯ふやす（手軽な変更ブロック）
Blockly.Blocks["var_change"] = {
  init: function () {
    this.appendValueInput("DELTA")
      .appendField("📦 変数")
      .appendField(new FieldVariableNoHiragana("Gokei"), "VAR")
      .appendField("を");
    this.appendDummyInput().appendField("ふやす");
    this.setInputsInline(true);
    this.setPreviousStatement(true, null);
    this.setNextStatement(true, null);
    this.setColour(CAT_COLORS.var);
    this.setTooltip("変数の中身を指定した数だけ増やします（マイナスで減らす）");
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
        ["¥ わって整数（あまりは捨てる）", "\\"],
        ["÷のあまり（Mod）", "mod"],
        ["＾ べき乗（何回かける）", "^"],
      ]),
      "OP"
    );
    this.appendValueInput("B");
    this.setInputsInline(true);
    this.setOutput(true, "Number");
    this.setColour(CAT_COLORS.value);
    this.setTooltip("計算をします。7 ¥ 2 = 3（整数のわり算）、7 Mod 2 = 1（あまり）、7 ^ 2 = 49（7×7）。「÷のあまり」は偶数・奇数の判定にも使えます");
  },
};

// ===== ユーザーフォーム =====
CAT_COLORS.form = "#2c7be5";

// フォームの部品の名前を選ぶドロップダウン。
// 選べるのはフォームにある部品（app.js が window.UMB_formControls で渡す）だが、
// 一覧に無い名前（コードから戻したとき・部品の名前を変えたとき）も消さずに受け付ける
class FieldControlDropdown extends Blockly.FieldDropdown {
  constructor(typeFilter, fallback) {
    const self = { filter: typeFilter, fallback };
    super(function () {
      const list = (window.UMB_formControls ? window.UMB_formControls() : [])
        .filter((c) => !self.filter || self.filter.includes(c.type))
        .map((c) => [c.name, c.name]);
      const cur = this && typeof this.getValue === "function" ? this.getValue() : null;
      if (cur && !list.some((o) => o[1] === cur)) list.push([cur, cur]);
      if (!list.length) list.push([self.fallback, self.fallback]);
      return list;
    });
  }
  doClassValidation_(newValue) {
    return typeof newValue === "string" && newValue ? newValue : null;
  }
}

// ボタンをクリックしたとき（Private Sub CommandButton1_Click() 〜 End Sub）
Blockly.Blocks["form_button_click"] = {
  init: function () {
    this.appendDummyInput()
      .appendField("🖱")
      .appendField(new FieldControlDropdown(["CommandButton"], "CommandButton1"), "CONTROL")
      .appendField("をクリックしたとき");
    this.appendStatementInput("DO").appendField("↓ やること");
    this.appendDummyInput().appendField("（ここまで）");
    this.setColour(CAT_COLORS.form);
    this.setTooltip("フォームのボタンを押したときに動く命令です（VBA では Private Sub ボタン名_Click()）。中にブロックを入れよう");
  },
};

// フォームが開いたとき（Private Sub UserForm_Initialize()）
Blockly.Blocks["form_initialize"] = {
  init: function () {
    this.appendDummyInput().appendField("🪟 フォームが開いたとき");
    this.appendStatementInput("DO").appendField("↓ やること");
    this.appendDummyInput().appendField("（ここまで）");
    this.setColour(CAT_COLORS.form);
    this.setTooltip("フォームが開いたときに1回だけ動く命令です（VBA では Private Sub UserForm_Initialize()）。最初の準備に使います");
  },
};

const CONTROL_PROP_OPTIONS = [
  ["表示する文字（Caption）", "Caption"],
  ["入力された文字（Text）", "Text"],
  ["見える（Visible）", "Visible"],
  ["使える（Enabled）", "Enabled"],
];

// 部品のプロパティを変える（Label1.Caption = …）
Blockly.Blocks["ctl_set_prop"] = {
  init: function () {
    this.appendValueInput("VALUE")
      .appendField("🏷")
      .appendField(new FieldControlDropdown(null, "Label1"), "CONTROL")
      .appendField("の")
      .appendField(new Blockly.FieldDropdown(CONTROL_PROP_OPTIONS), "PROP")
      .appendField("を");
    this.appendDummyInput().appendField("にする");
    this.setInputsInline(true);
    this.setPreviousStatement(true, null);
    this.setNextStatement(true, null);
    this.setColour(CAT_COLORS.form);
    this.setTooltip("フォームの部品の表示を変えます。ラベルに文字を出すときは「表示する文字（Caption）」を使おう");
  },
};

// 部品のプロパティを読む（TextBox1.Text など）
Blockly.Blocks["ctl_get_prop"] = {
  init: function () {
    this.appendDummyInput()
      .appendField("🏷")
      .appendField(new FieldControlDropdown(null, "TextBox1"), "CONTROL")
      .appendField("の")
      .appendField(new Blockly.FieldDropdown([CONTROL_PROP_OPTIONS[1], CONTROL_PROP_OPTIONS[0]]), "PROP");
    this.setOutput(true, null);
    this.setColour(CAT_COLORS.form);
    this.setTooltip("フォームの部品の中身を取り出します。テキストボックスに入力された文字は「入力された文字（Text）」です");
  },
};

Blockly.Blocks["form_unload"] = {
  init: function () {
    this.appendDummyInput().appendField("🪟 フォームを閉じる（Unload Me）");
    this.setPreviousStatement(true, null);
    this.setNextStatement(true, null);
    this.setColour(CAT_COLORS.form);
    this.setTooltip("フォームを閉じます。「終了」ボタンでよく使います");
  },
};

Blockly.Blocks["prog_end"] = {
  init: function () {
    this.appendDummyInput().appendField("⏹ プログラムを終わる（End）");
    this.setPreviousStatement(true, null);
    this.setNextStatement(true, null);
    this.setColour(CAT_COLORS.form);
    this.setTooltip("プログラムをそこで終わりにします（フォームも閉じます）");
  },
};

// 文字を数にする（Val）
Blockly.Blocks["math_val"] = {
  init: function () {
    this.appendValueInput("VALUE").appendField("🔢 数にする Val(");
    this.appendDummyInput().appendField(")");
    this.setInputsInline(true);
    this.setOutput(true, "Number");
    this.setColour(CAT_COLORS.value);
    this.setTooltip("文字を数に変えます。InputBox やテキストボックスに入力された文字で計算するときに使います（数でない文字は 0 になります）");
  },
};

// ===== 変数の型・手続き・書式（教科書 第2節） =====
const VAR_TYPE_OPTIONS = [
  ["整数型（Integer）", "Integer"],
  ["長整数型（Long）", "Long"],
  ["単精度浮動小数点数型（Single）", "Single"],
  ["倍精度浮動小数点数型（Double）", "Double"],
  ["文字列型（String）", "String"],
  ["バリアント型（Variant）", "Variant"],
];

// 箱（変数）を用意する：Dim Tanka As Integer
Blockly.Blocks["var_dim"] = {
  init: function () {
    this.appendDummyInput()
      .appendField("📦 箱を用意する Dim")
      .appendField(new FieldVariableNoHiragana("Tanka"), "VAR")
      .appendField("As")
      .appendField(new Blockly.FieldDropdown(VAR_TYPE_OPTIONS), "TYPE");
    this.setPreviousStatement(true, null);
    this.setNextStatement(true, null);
    this.setColour(CAT_COLORS.var);
    this.setTooltip("変数（値を入れる箱）を、型を決めて用意します。整数型は -32768〜32767 の整数、文字列型は文字を入れる箱です");
  },
};

// 命令のまとまり（Sub）を作る
Blockly.Blocks["proc_sub"] = {
  init: function () {
    this.appendDummyInput()
      .appendField("📜 命令のまとまり Sub")
      .appendField(
        new Blockly.FieldTextInput("Keisan", (v) => (/^[A-Za-z][A-Za-z0-9_]{0,39}$/.test(v) ? v : null)),
        "NAME"
      );
    this.appendStatementInput("DO").appendField("↓ やること");
    this.appendDummyInput().appendField("（ここまで）");
    this.setColour("#6d4c41");
    this.setTooltip("いくつかの命令に名前を付けてまとめます（Sub）。「📞 Call」で呼び出して使います。名前は英字で始めよう");
  },
};

// まとめた命令を呼び出す：Call Keisan
Blockly.Blocks["proc_call"] = {
  init: function () {
    const field = new Blockly.FieldDropdown(function () {
      const ws = Blockly.getMainWorkspace && Blockly.getMainWorkspace();
      const names = ws ? ws.getTopBlocks(false).filter((b) => b.type === "proc_sub").map((b) => b.getFieldValue("NAME")) : [];
      const cur = this && typeof this.getValue === "function" ? this.getValue() : null;
      if (cur && !names.includes(cur)) names.push(cur);
      if (!names.length) names.push("Keisan");
      return names.map((n) => [n, n]);
    });
    field.doClassValidation_ = (v) => (typeof v === "string" && v ? v : null);
    this.appendDummyInput().appendField("📞 Call").appendField(field, "NAME").appendField("を呼び出す");
    this.setPreviousStatement(true, null);
    this.setNextStatement(true, null);
    this.setColour("#6d4c41");
    this.setTooltip("「📜 命令のまとまり Sub」で作った命令を、ここで動かします");
  },
};

// 書式を付けて文字にする：Format(Kin, "###,##0")
Blockly.Blocks["math_format"] = {
  init: function () {
    this.appendValueInput("VALUE").appendField("🔢 Format(");
    this.appendDummyInput()
      .appendField(",")
      .appendField(new Blockly.FieldTextInput("###,##0"), "FORMAT")
      .appendField(")");
    this.setInputsInline(true);
    this.setOutput(true, null);
    this.setColour(CAT_COLORS.value);
    this.setTooltip("数を決まった形の文字にします。例：###,##0 → 3けたごとに「,」（8000 → 8,000）、0.00 → 小数第2位まで");
  },
};

// 改行（vbCrLf）
Blockly.Blocks["value_newline"] = {
  init: function () {
    this.appendDummyInput().appendField("↵ 改行（vbCrLf）");
    this.setOutput(true, "String");
    this.setColour(CAT_COLORS.value);
    this.setTooltip("文字の途中で行をかえます。ラベルに何行も表示するときに、& でつなげて使います");
  },
};

// ===== ツールボックス定義 =====
// 数値シャドウ（普段は数字を直接入力、変数や計算ブロックで差し替え可能）
function numShadow(n) {
  return { shadow: { type: "value_number", fields: { NUM: n } } };
}

// カテゴリは教科書の順（フォーム → 入出力 → 変数 → 計算 → 条件 → くり返し → セル…）に並べる
const TOOLBOX_FORM = {
  kind: "category",
  name: "🪟 フォーム",
  colour: CAT_COLORS.form,
  contents: [
    { kind: "block", type: "form_button_click" },
    { kind: "block", type: "form_initialize" },
    {
      kind: "block",
      type: "ctl_set_prop",
      inputs: { VALUE: { shadow: { type: "value_text", fields: { TEXT: "こんにちは" } } } },
    },
    { kind: "block", type: "ctl_get_prop" },
    { kind: "block", type: "form_unload" },
    { kind: "block", type: "prog_end" },
  ],
};

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
        { kind: "block", type: "cell_clear_row", inputs: { ROW: numShadow(1) } },
        {
          kind: "block",
          type: "cells_set_value",
          inputs: { ROW: numShadow(1), COL: numShadow(1) },
        },
        {
          kind: "block",
          type: "cells_get_value",
          inputs: { ROW: numShadow(1), COL: numShadow(1) },
        },
        {
          kind: "block",
          type: "cells_clear",
          inputs: { ROW: numShadow(1), COL: numShadow(1) },
        },
        {
          kind: "block",
          type: "range_border",
          inputs: { R1: numShadow(1), C1: numShadow(1), R2: numShadow(5), C2: numShadow(3) },
        },
        {
          kind: "block",
          type: "range_select",
          inputs: { R1: numShadow(1), C1: numShadow(1), R2: numShadow(5), C2: numShadow(3) },
        },
      ],
    },
    {
      kind: "category",
      name: "🔁 くりかえし",
      colour: CAT_COLORS.loop,
      contents: [
        { kind: "block", type: "loop_repeat", inputs: { TIMES: numShadow(3) } },
        {
          kind: "block",
          type: "loop_range",
          inputs: { START: numShadow(1), END: numShadow(5) },
        },
        {
          kind: "block",
          type: "loop_for_step",
          inputs: { START: numShadow(1), END: numShadow(10), STEP: numShadow(2) },
        },
        { kind: "block", type: "loop_while" },
        { kind: "block", type: "loop_do_until" },
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
        { kind: "block", type: "cond_logic" },
        { kind: "block", type: "cond_not" },
        { kind: "block", type: "value_bool" },
        { kind: "block", type: "cond_select", extraState: { count: 2 } },
        { kind: "block", type: "prog_exit_sub" },
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
        { kind: "block", type: "fmt_fontsize", inputs: { SIZE: numShadow(14) } },
        {
          kind: "block",
          type: "cells_bgcolor",
          inputs: { ROW: numShadow(1), COL: numShadow(1) },
        },
        {
          kind: "block",
          type: "cells_bold",
          inputs: { ROW: numShadow(1), COL: numShadow(1) },
        },
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
      custom: "VARIABLE_JP",
    },
    {
      kind: "category",
      name: "💬 入出力",
      colour: "#e91e63",
      contents: [
        { kind: "block", type: "io_inputbox" },
        { kind: "block", type: "io_msgbox" },
        { kind: "block", type: "text_concat" },
        { kind: "block", type: "value_date" },
        { kind: "block", type: "math_round", inputs: { DIGITS: numShadow(1) } },
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
        { kind: "block", type: "math_val" },
        { kind: "block", type: "math_format" },
        { kind: "block", type: "value_newline" },
      ],
    },
    {
      kind: "category",
      name: "📜 命令のまとまり",
      colour: "#6d4c41",
      contents: [
        { kind: "block", type: "proc_sub" },
        { kind: "block", type: "proc_call" },
      ],
    },
  ],
};

// 教科書の順に並べ替える（名前で探すので、上の定義の順番は気にしなくてよい）
(function reorderToolbox() {
  const order = ["🪟 フォーム", "💬 入出力", "📦 変数", "🔢 数・文字", "📜 命令のまとまり", "❓ もし〜なら", "🔁 くりかえし", "📝 セル操作", "🎨 色・文字デザイン", "📦 リスト（配列）", "📑 シート操作"];
  const all = [TOOLBOX_FORM].concat(TOOLBOX.contents);
  TOOLBOX.contents = order.map((n) => all.find((c) => c.name === n)).filter(Boolean);
  // 「💬 入出力」の MsgBox には、最初から「こんにちは」の文字を入れておく
  const io = TOOLBOX.contents.find((c) => c.name === "💬 入出力");
  const msg = io && io.contents.find((b) => b.type === "io_msgbox");
  if (msg) msg.inputs = { MSG: { shadow: { type: "value_text", fields: { TEXT: "こんにちは" } } } };
})();
