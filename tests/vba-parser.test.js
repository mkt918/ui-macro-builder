// コード → ブロック変換（src/vba-parser.js）のテスト。実行方法: node --test tests/*.test.js
const test = require("node:test");
const assert = require("node:assert/strict");
const path = require("node:path");
const P = require(path.join(__dirname, "../src/vba-parser.js"));

const types = (r) => r.roots.map((x) => x.type);
// ブロックの木を平らにして、type の一覧にする
function flat(b, out = []) {
  if (!b) return out;
  out.push(b.type);
  Object.values(b.inputs || {}).forEach((i) => flat(i.block, out));
  if (b.next) flat(b.next.block, out);
  return out;
}

test("教科書 例題 4-2-03：イベントの手続きと型付き Dim がブロックになる", () => {
  const r = P.parseVBA(`Option Explicit
Private Sub CommandButton1_Click()
    Dim Tanka As Integer
    Dim Suu As Integer
    Dim Kin As Integer
    Tanka = 100
    Suu = Val(InputBox("数量を入力してください"))
    Kin = Tanka * Suu
    MsgBox ("金額は" & Kin & "円です")
End Sub`);
  assert.deepEqual(types(r), ["form_button_click"]);
  const all = flat(r.roots[0]);
  assert.equal(all.filter((t) => t === "var_dim").length, 3);
  assert.ok(all.includes("math_val"));
  assert.ok(all.includes("io_inputbox"));
  assert.ok(all.includes("io_msgbox"));
  assert.equal(r.roots[0].fields.CONTROL, "CommandButton1");
});

test("フォームの手続き・部品のプロパティ・Unload Me", () => {
  const r = P.parseVBA(`Private Sub UserForm_Initialize()
    Label1.Caption = ""
End Sub
Private Sub CommandButton1_Click()
    Me.Label1.Caption = "こんにちは " & TextBox1.Text
End Sub
Private Sub CommandButton2_Click()
    Unload Me
End Sub`);
  assert.deepEqual(types(r), ["form_initialize", "form_button_click", "form_button_click"]);
  const b1 = flat(r.roots[1]);
  assert.deepEqual(b1, ["form_button_click", "ctl_set_prop", "text_concat", "value_text", "ctl_get_prop"]);
  assert.deepEqual(flat(r.roots[2]), ["form_button_click", "form_unload"]);
});

test("Sub と Call・¥・^・Format・vbCrLf", () => {
  const r = P.parseVBA(`Sub MyMacro()
    Call Keisan
End Sub
Sub Keisan()
    Kotae = 7 ¥ 2 + 2 ^ 3
    Label1.Caption = Format(Kotae, "###,##0") & vbCrLf
End Sub`);
  assert.deepEqual(types(r), ["proc_call", "proc_sub"]);
  const all = flat(r.roots[1]);
  assert.ok(all.includes("math_format"));
  assert.ok(all.includes("value_newline"));
  const ops = [];
  const walk = (b) => {
    if (!b) return;
    if (b.type === "value_math") ops.push(b.fields.OP);
    Object.values(b.inputs || {}).forEach((i) => walk(i.block));
    if (b.next) walk(b.next.block);
  };
  walk(r.roots[1]);
  assert.deepEqual(ops.sort(), ["+", "\\", "^"].sort());
});

test("入れ子のくり返し（i と j）と、ブロックで表せないもの", () => {
  const ok = P.parseVBA(`Sub MyMacro()
    Dim i As Integer
    Dim j As Integer
    For i = 1 To 3
        For j = 1 To 2
            Range("B" & j).Value = j
        Next j
    Next i
End Sub`);
  assert.deepEqual(types(ok), ["loop_range"]);
  assert.throws(() => P.parseVBA("Sub MyMacro()\n For i = 1 To 3\n  For j = 1 To 2\n   Cells(i, j).Value = i\n  Next j\n Next i\nEnd Sub"), /外側/);
  assert.throws(() => P.parseVBA("Option Explicit\nDim Umei(5) As String\nPrivate Sub CommandButton1_Click()\n Label1.Caption = Umei(1)\nEnd Sub"), /Sub の外の Dim/);
  assert.throws(() => P.parseVBA("Sub MyMacro()\n Select Case x\n Case 1\n y = 1\n End Select\nEnd Sub"));
});

test("文字列の中の \"\" は \" 1文字", () => {
  const r = P.parseVBA(`Sub MyMacro()
    MsgBox "彼は""こんにちは""と言った"
End Sub`);
  assert.equal(r.roots[0].inputs.MSG.block.fields.TEXT, '彼は"こんにちは"と言った');
});
