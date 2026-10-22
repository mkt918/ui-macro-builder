// 実行エンジン（src/vba/*）のテスト。実行方法: node --test tests/
const test = require("node:test");
const assert = require("node:assert/strict");
const path = require("node:path");

const { runScript, Session } = require(path.join(__dirname, "../src/vba/session.js"));
const { parseProgram } = require(path.join(__dirname, "../src/vba/ast-parser.js"));
const { fixZenkaku } = require(path.join(__dirname, "../src/vba/lexer.js"));

const form = (...controls) => ({ name: "UserForm1", caption: "UserForm1", controls: controls.map((c) => (typeof c === "string" ? { name: c, type: typeFromName(c) } : c)) });
function typeFromName(n) {
  if (n.startsWith("Label")) return "Label";
  if (n.startsWith("TextBox")) return "TextBox";
  if (n.startsWith("ListBox")) return "ListBox";
  return "CommandButton";
}

// ---------- 教科書 p.91〜94 ----------
test("例題 4-2-03：InputBox に 5 と答えると「金額は500円です」", () => {
  const src = `Option Explicit
Private Sub CommandButton1_Click()
    Dim Tanka As Integer
    Dim Suu As Integer
    Dim Kin As Integer
    Tanka = 100
    Suu = Val(InputBox("数量を入力してください"))
    Kin = Tanka * Suu
    MsgBox ("金額は" & Kin & "円です")
End Sub`;
  const r = runScript(src, { form: form("CommandButton1"), actions: [{ showForm: true }, { click: "CommandButton1", inputs: ["5"] }] });
  assert.equal(r.error, null);
  assert.deepEqual(r.state.msgs, ["金額は500円です"]);
});

test("算術演算子の表（A=7 の単精度型）", () => {
  const src = `Sub Test()
    Dim A As Single, Kotae As Single
    A = 7
    Kotae = A + 2: Debug.Print Kotae
    Kotae = A - 2: Debug.Print Kotae
    Kotae = A * 2: Debug.Print Kotae
    Kotae = A / 2: Debug.Print Kotae
    Kotae = A ¥ 2: Debug.Print Kotae
    Kotae = A \\ 2: Debug.Print Kotae
    Kotae = A Mod 2: Debug.Print Kotae
    Kotae = A ^ 2: Debug.Print Kotae
End Sub`;
  const r = runScript(src, { actions: [{ run: "Test" }] });
  assert.equal(r.error, null);
  assert.deepEqual(r.state.debug, ["9", "5", "14", "3.5", "3", "3", "1", "49"]);
});

test("優先順位は教科書の表どおり（* / ¥ Mod は同じ順位で左から）", () => {
  const src = `Sub Test()
    Debug.Print 7 Mod 4 * 2
    Debug.Print 2 + 3 * 4 ^ 2
    Debug.Print -2 ^ 2
    Debug.Print (2 + 3) * 4
End Sub`;
  const r = runScript(src, { actions: [{ run: "Test" }] });
  assert.deepEqual(r.state.debug, ["6", "50", "-4", "20"]);
});

test("Val は先頭から数として読み、数でない文字で止める", () => {
  const src = `Sub Test()
    Debug.Print Val("12abc")
    Debug.Print Val("abc")
    Debug.Print Val(" 1 2 3")
    Debug.Print Val("３５")
    Debug.Print "1" & "2"
    Debug.Print Val("1") + Val("2")
End Sub`;
  const r = runScript(src, { actions: [{ run: "Test" }] });
  assert.deepEqual(r.state.debug, ["12", "0", "123", "35", "12", "3"]);
});

// ---------- 教科書 p.103〜108（例題 4-3-01） ----------
const SRC_4_3_01 = `Option Explicit
Private Sub CommandButton1_Click()
    Dim Omosa As Integer
    Dim Ryokin As Integer
    Omosa = Val(InputBox("重量を入力して下さい(g)", "重量入力", 0))
    If Omosa = 0 Then
        End
    End If
    If Omosa > 50 Then
        MsgBox("重量オーバーです")
        Exit Sub
    End If
    If Omosa <= 25 Then
        Ryokin = 84
    Else
        Ryokin = 94
    End If
    Label3.Caption = Omosa
    Label4.Caption = Ryokin
End Sub`;

test("例題 4-3-01：22→84、25→84、40→94、65→重量オーバー、0→終了", () => {
  const f = form("Label1", "Label2", "Label3", "Label4", "CommandButton1");
  const r = runScript(SRC_4_3_01, {
    form: f,
    actions: [
      { showForm: true },
      { click: "CommandButton1", inputs: ["22"] },
      { click: "CommandButton1", inputs: ["25"] },
      { click: "CommandButton1", inputs: ["40"] },
      { click: "CommandButton1", inputs: ["65"] },
      { click: "CommandButton1", inputs: ["0"] },
    ],
  });
  const s = r.session;
  const labelAfter = (segIdx) => s.modelAt(s.segments[segIdx].end - 1).controls.Label4.Caption;
  assert.equal(labelAfter(1), "84");
  assert.equal(labelAfter(2), "84");
  assert.equal(labelAfter(3), "94");
  assert.deepEqual(r.results[4].newMsgs, ["重量オーバーです"]);
  assert.equal(labelAfter(4), "94"); // Exit Sub なのでラベルは前のまま
  assert.equal(r.session.status, "ended");
  assert.equal(r.state.form.visible, false);
});

test("InputBox の第2・第3引数（タイトル・初期値）が届く", () => {
  const s = new Session(SRC_4_3_01, { form: form("Label3", "Label4", "CommandButton1") });
  s.showForm();
  const res = s.click("CommandButton1");
  assert.equal(res.status, "waiting");
  assert.equal(res.request.type, "inputbox");
  assert.equal(res.request.title, "重量入力");
  assert.equal(res.request.default, "0");
});

// ---------- 教科書 p.127〜131（例題 4-3-06） ----------
const URIAGE = "高橋　礼,4\r\n山口　桃子,5\r\n奥原　希美,2\r\n松友　美紀,7\r\n永原　和香,8\r\n松本　友子,10\r\n米本　小奈津,23\r\n田中　志保子,5\r\n東野　亜里沙,13\r\n栗原　歩美,20\r\n広田　花,9\r\n";

test("例題 4-3-06：uriage.csv を読んで、名前と金額をラベルに並べる", () => {
  const src = `Option Explicit
Private Sub CommandButton1_Click()
    Dim KouMei As String
    Dim Suu As Long
    Dim Kin As Long
    Label1.Caption = ""
    Label2.Caption = ""
    Open ThisWorkbook.Path & "¥uriage.csv" For Input As #1
    Do While Not EOF(1)
        Input #1, KouMei, Suu
        Kin = Suu * 2000
        Label1.Caption = Label1.Caption & KouMei & Chr(13) & Chr(10)
        Label2.Caption = Label2.Caption & Format(Kin, "###,##0") & Chr(13) & Chr(10)
    Loop
    Close #1
End Sub`;
  const r = runScript(src, {
    form: form("Label1", "Label2", "CommandButton1"),
    files: { "uriage.csv": URIAGE },
    actions: [{ showForm: true }, { click: "CommandButton1" }],
  });
  assert.equal(r.error, null);
  const names = r.state.controls.Label1.Caption.split("\r\n");
  assert.equal(names[0], "高橋　礼");
  assert.equal(names[10], "広田　花");
  const kin = r.state.controls.Label2.Caption.split("\r\n");
  assert.deepEqual(kin.slice(0, 11), ["8,000", "10,000", "4,000", "14,000", "16,000", "20,000", "46,000", "10,000", "26,000", "40,000", "18,000"]);
});

test("ファイルが無いときはエラー53（やさしい説明つき）", () => {
  const src = `Sub Test()
    Open "nai.csv" For Input As #1
End Sub`;
  const r = runScript(src, { actions: [{ run: "Test" }] });
  assert.equal(r.error.code, 53);
  assert.match(r.error.friendly, /nai\.csv/);
});

test("For Output と Print # で書き出したファイルを読み返せる", () => {
  const src = `Sub Test()
    Dim i As Integer, s As String
    Open "out.txt" For Output As #1
    For i = 1 To 3
        Print #1, "行" & i
    Next i
    Close #1
    Open "out.txt" For Input As #2
    Do Until EOF(2)
        Line Input #2, s
        Debug.Print s
    Loop
    Close #2
End Sub`;
  const r = runScript(src, { actions: [{ run: "Test" }] });
  assert.equal(r.error, null);
  assert.equal(r.state.files["out.txt"], "行1\r\n行2\r\n行3\r\n");
  assert.deepEqual(r.state.debug, ["行1", "行2", "行3"]);
});

// ---------- 教科書 p.141〜145（例題 4-4-01） ----------
test("例題 4-4-01：Initialize で配列に入れ、TextBox の番号の売場名を表示", () => {
  const src = `Option Explicit
Dim Umei(5) As String
Private Sub UserForm_Initialize()
    Umei(1) = "食料品"
    Umei(2) = "日用品"
    Umei(3) = "衣料品"
    Umei(4) = "家電"
    Umei(5) = "家具"
End Sub
Private Sub CommandButton1_Click()
    Dim Soe As Integer
    Soe = Val(TextBox1.Text)
    Label1.Caption = Umei(Soe)
End Sub`;
  const r = runScript(src, {
    form: form("TextBox1", "Label1", "CommandButton1"),
    actions: [{ showForm: true }, { text: { TextBox1: "1" }, click: "CommandButton1" }, { text: { TextBox1: "4" }, click: "CommandButton1" }],
  });
  assert.equal(r.error, null);
  assert.equal(r.state.controls.Label1.Caption, "家電");
  const s = r.session;
  assert.equal(s.modelAt(s.segments[2].end - 1).controls.Label1.Caption, "食料品");
  // 配列の箱：添字 0〜5（0 は使わない）
  assert.deepEqual(r.state.arrays.Umei.dims, [[0, 5]]);
  assert.equal(r.state.arrays.Umei.data["1"], "食料品");
});

test("配列の範囲外はエラー9（やさしい説明つき）", () => {
  const src = `Sub Test()
    Dim a(5) As Integer
    a(6) = 1
End Sub`;
  const r = runScript(src, { actions: [{ run: "Test" }] });
  assert.equal(r.error.code, 9);
  assert.match(r.error.friendly, /0〜5/);
  assert.equal(r.error.line, 3);
});

// ---------- 型・エラー ----------
test("Integer に入りきらないとオーバーフロー（エラー6）", () => {
  const src = `Sub Test()
    Dim a As Integer
    a = 40000
End Sub`;
  const r = runScript(src, { actions: [{ run: "Test" }] });
  assert.equal(r.error.code, 6);
  assert.match(r.error.hint, /Long/);
});

test("整数型への代入は銀行型の丸め", () => {
  const src = `Sub Test()
    Dim a As Integer
    a = 2.5: Debug.Print a
    a = 3.5: Debug.Print a
    a = -2.5: Debug.Print a
End Sub`;
  const r = runScript(src, { actions: [{ run: "Test" }] });
  assert.deepEqual(r.state.debug, ["2", "4", "-2"]);
});

test("Option Explicit で Dim していない変数はエラー", () => {
  const src = `Option Explicit
Sub Test()
    Dim Tanka As Integer
    Tanaka = 100
End Sub`;
  const r = runScript(src, { actions: [{ run: "Test" }] });
  assert.equal(r.error.kind, "compile");
  assert.match(r.error.friendly, /Tanaka/);
  assert.equal(r.error.line, 4);
});

test("空っぽの文字を計算に使うと型が一致しません（エラー13）", () => {
  const src = `Sub Test()
    Dim a As Integer
    a = InputBox("数を入れてください") + 1
End Sub`;
  const r = runScript(src, { actions: [{ run: "Test", inputs: [""] }] });
  assert.equal(r.error.code, 13);
  assert.match(r.error.hint, /Val/);
});

test("エラーが起きても、それまでの step は残る", () => {
  const src = `Sub Test()
    Range("A1").Value = 10
    Range("A2").Value = 1 / 0
End Sub`;
  const r = runScript(src, { actions: [{ run: "Test" }] });
  assert.equal(r.error.code, 11);
  assert.equal(r.error.line, 3);
  assert.equal(r.session.finalModel().cells.A1.value, 10);
});

test("全角の記号は ZenkakuError になり、fixZenkaku で直せる（文字列の中は直さない）", () => {
  const src = `Sub Test()
    Kotae ＝ ５ ＋ 1
    MsgBox "答えは＝" & Kotae
End Sub`;
  assert.throws(() => new Session(src), (e) => e.kind === "zenkaku" && e.zenkaku.length === 3 && e.line === 2);
  const fixed = fixZenkaku(src);
  assert.match(fixed, /Kotae = 5 \+ 1/);
  assert.match(fixed, /"答えは＝"/);
  const r = runScript(fixed, { actions: [{ run: "Test" }] });
  assert.deepEqual(r.state.msgs, ["答えは＝6"]);
});

test("文法エラー：End If の書き忘れ・Next の書き忘れ", () => {
  assert.throws(() => parseProgram("Sub A()\n  If x Then\n    y = 1\nEnd Sub"), (e) => e.line === 2 && /End If/.test(e.message));
  assert.throws(() => parseProgram("Sub A()\n  For i = 1 To 3\n    y = 1\nEnd Sub"), (e) => e.line === 2 && /Next/.test(e.message));
  assert.throws(() => parseProgram("x = 1"), (e) => /Sub の外/.test(e.friendly));
});

// ---------- 条件・繰り返し ----------
test("FizzBuzz（If ElseIf と Mod）", () => {
  const src = `Sub Test()
    Dim i As Integer
    For i = 1 To 15
        If i Mod 15 = 0 Then
            Cells(i, 1).Value = "FizzBuzz"
        ElseIf i Mod 3 = 0 Then
            Cells(i, 1).Value = "Fizz"
        ElseIf i Mod 5 = 0 Then
            Cells(i, 1).Value = "Buzz"
        Else
            Cells(i, 1).Value = i
        End If
    Next i
End Sub`;
  const r = runScript(src, { actions: [{ run: "Test" }] });
  const v = (n) => r.state.cells["A" + n].value;
  assert.equal(v(1), 1);
  assert.equal(v(3), "Fizz");
  assert.equal(v(5), "Buzz");
  assert.equal(v(15), "FizzBuzz");
});

test("Select Case（To・Is・カンマ区切り）と And / Or / Not", () => {
  const src = `Function Hyoka(ten As Integer) As String
    Select Case ten
        Case Is >= 80
            Hyoka = "A"
        Case 60 To 79
            Hyoka = "B"
        Case 0, 1, 2
            Hyoka = "ほぼ0"
        Case Else
            Hyoka = "C"
    End Select
End Function
Sub Test()
    Debug.Print Hyoka(90) & Hyoka(70) & Hyoka(1) & Hyoka(30)
    Debug.Print (3 > 2 And 1 > 2) & "/" & (3 > 2 Or 1 > 2) & "/" & Not (3 > 2)
End Sub`;
  const r = runScript(src, { actions: [{ run: "Test" }] });
  assert.equal(r.error, null);
  assert.deepEqual(r.state.debug, ["ABほぼ0C", "False/True/False"]);
});

test("二重ループで九九（カウンタ i と j）", () => {
  const src = `Sub Kuku()
    Dim i As Integer, j As Integer
    For i = 1 To 9
        For j = 1 To 9
            Cells(i, j).Value = i * j
        Next j
    Next i
End Sub`;
  const r = runScript(src, { actions: [{ run: "Kuku" }] });
  assert.equal(r.state.cells.I9.value, 81);
  assert.equal(r.state.cells.C7.value, 21);
});

test("Do Until / Loop While / Exit Do / Exit For", () => {
  const src = `Sub Test()
    Dim n As Integer, i As Integer
    n = 0
    Do Until n >= 5
        n = n + 2
    Loop
    Debug.Print n
    Do
        n = n - 1
        If n = 3 Then Exit Do
    Loop While n > 0
    Debug.Print n
    For i = 1 To 100
        If i = 4 Then Exit For
    Next i
    Debug.Print i
End Sub`;
  const r = runScript(src, { actions: [{ run: "Test" }] });
  assert.deepEqual(r.state.debug, ["6", "3", "4"]);
});

test("無限ループは止めて、やさしく知らせる", () => {
  const src = `Sub Test()
    Dim n As Integer
    Do While n < 10
    Loop
End Sub`;
  const r = runScript(src, { actions: [{ run: "Test" }] });
  assert.equal(r.error.kind, "limit");
  assert.match(r.error.friendly, /無限ループ/);
});

// ---------- 配列・並べ替え・ByRef ----------
test("バブルソート（配列を Sub に渡し、要素を ByRef で交換）", () => {
  const src = `Option Explicit
Dim Ten(1 To 6) As Integer
Sub Swap(a As Integer, b As Integer)
    Dim w As Integer
    w = a
    a = b
    b = w
End Sub
Sub Narabe(x() As Integer)
    Dim i As Integer, j As Integer
    For i = LBound(x) To UBound(x) - 1
        For j = LBound(x) To UBound(x) - i
            If x(j) > x(j + 1) Then Call Swap(x(j), x(j + 1))
        Next j
    Next i
End Sub
Sub Test()
    Dim k As Integer
    Ten(1) = 5: Ten(2) = 2: Ten(3) = 9: Ten(4) = 1: Ten(5) = 7: Ten(6) = 3
    Narabe Ten
    For k = 1 To 6
        Debug.Print Ten(k)
    Next k
End Sub`;
  const r = runScript(src, { actions: [{ run: "Test" }] });
  assert.equal(r.error, null);
  assert.deepEqual(r.state.debug, ["1", "2", "3", "5", "7", "9"]);
  // 最初の Swap が終わった時点で、呼び出し元（Narabe）の配列 x の表示が入れ替わっている
  const s = r.session;
  const firstSwapEnd = s.steps.findIndex((st) => st.procEvent === "end" && st.proc === "Swap");
  const m = s.modelAt(firstSwapEnd);
  assert.equal(m.proc, "Narabe");
  assert.equal(m.arrays.x.data["1"], 2);
  assert.equal(m.arrays.x.data["2"], 5);
  // Swap の途中でも、モジュールの配列 Ten の箱は入れ替わっていく
  const inSwap = s.steps.findIndex((st) => st.proc === "Swap" && st.scope === "var" && st.key === "b");
  assert.equal(s.modelAt(inSwap).arrays.Ten.data["2"], 5);
});

test("2次元配列", () => {
  const src = `Sub Test()
    Dim t(1 To 3, 1 To 4) As Integer
    Dim i As Integer, j As Integer
    For i = 1 To 3
        For j = 1 To 4
            t(i, j) = i * 10 + j
        Next j
    Next i
    Debug.Print t(2, 3) & "," & UBound(t, 2)
End Sub`;
  const r = runScript(src, { actions: [{ run: "Test" }] });
  assert.deepEqual(r.state.debug, ["23,4"]);
});

test("数当てゲーム：InputBox と MsgBox が本物の順番で交互に出る", () => {
  const src = `Sub Game()
    Dim Seikai As Integer, Kotae As Integer
    Seikai = 7
    Do
        Kotae = Val(InputBox("1〜10の数を当ててね"))
        If Kotae < Seikai Then
            MsgBox "もっと大きい"
        ElseIf Kotae > Seikai Then
            MsgBox "もっと小さい"
        End If
    Loop Until Kotae = Seikai
    MsgBox "あたり！"
End Sub`;
  const s = new Session(src);
  const order = [];
  let res = s.start("Game");
  const answers = ["3", "9", "7"];
  while (res.status === "waiting") {
    order.push(res.request.type === "inputbox" ? "Q" : res.request.prompt);
    res = res.request.type === "inputbox" ? s.resume(answers.shift()) : s.resume(undefined);
  }
  assert.deepEqual(order, ["Q", "もっと大きい", "Q", "もっと小さい", "Q", "あたり！"]);
});

// ---------- 今のブロックが生成するコード ----------
test("ブロックが生成する形のコード（Range・arr・シート・書式）が動く", () => {
  const src = `Sub MyMacro()
    Dim i As Integer
    Dim arr(1 To 100) As Variant
    Dim goukei As Variant
    goukei = 0
    For i = 1 To 5
        Range("A" & i).Value = i * 2
        arr(i) = i
        goukei = goukei + i
    Next i
    Range("A1").Interior.Color = RGB(231, 76, 60)
    Range("A2").Font.Bold = True
    Range(Cells(1, 1), Cells(2, 2)).Borders.LineStyle = 1
    Range("A1").Copy Range("C1")
    Worksheets.Add.Name = "Data"
    Range("B1").Value = WorksheetFunction.Round(2.345, 2)
    If (goukei Mod 2 = 0) Then
        MsgBox "偶数"
    Else
        MsgBox "奇数"
    End If
End Sub`;
  const r = runScript(src, { actions: [{ run: "MyMacro" }] });
  assert.equal(r.error, null);
  assert.equal(r.state.cells.A5.value, 10);
  assert.equal(r.state.cells.A1.bg, "#e74c3c");
  assert.equal(r.state.cells.A2.bold, true);
  assert.equal(r.state.cells.B2.border, true);
  assert.equal(r.state.cells.C1.value, 2);
  assert.equal(r.state.cells.B1.value, 2.35);
  assert.deepEqual(r.state.sheets, ["Sheet1", "Data"]);
  assert.equal(r.state.arrays.arr.data["3"], 3);
  assert.deepEqual(r.state.msgs, ["奇数"]);
});

// ---------- 記録（差分と全体の写し） ----------
test("差分から作った各時点の状態が、順番に再生したものと一致する", () => {
  const src = `Sub Test()
    Dim i As Integer
    For i = 1 To 700
        Cells(i Mod 20 + 1, 1).Value = i
    Next i
End Sub`;
  const r = runScript(src, { actions: [{ run: "Test" }] });
  const s = r.session;
  assert.ok(s.steps.length > 1400);
  assert.ok(s.steps.some((st) => st.kf), "全体の写し（kf）がある");
  const seq = [];
  for (let k = 0; k < s.steps.length; k += 97) seq.push([k, JSON.stringify(s.store.modelCopy(k))]);
  // 逆順・飛び飛びに取り出しても同じ
  for (const [k, json] of seq.reverse()) assert.equal(JSON.stringify(s.store.modelCopy(k)), json);
});

test("step には行番号と、分岐の結果が入る", () => {
  const r = runScript(SRC_4_3_01, { form: form("Label3", "Label4", "CommandButton1"), actions: [{ showForm: true }, { click: "CommandButton1", inputs: ["22"] }] });
  const steps = r.session.steps;
  const br = steps.filter((st) => st.scope === "branch").map((st) => [st.line, st.branch]);
  assert.deepEqual(br, [[6, false], [9, false], [13, true]]);
  const start = steps.find((st) => st.procEvent === "start");
  assert.equal(start.line, 2);
});

// ---------- M6・M7：探索・並べ替え（教科書後半の定番プログラム） ----------
const NUMS = "Dim a(1 To 8) As Integer\n    a(1) = 5: a(2) = 3: a(3) = 8: a(4) = 1: a(5) = 9: a(6) = 2: a(7) = 7: a(8) = 4";
const showAll = "For i = 1 To 8\n        s = s & a(i) & \",\"\n    Next i\n    Debug.Print s";

test("線形探索（見つかった位置と、比べた回数）", () => {
  const src = `Sub Test()
    Dim i As Integer, Pos As Integer, N As Integer
    ${NUMS}
    Pos = 0
    For i = 1 To 8
        N = N + 1
        If a(i) = 9 Then
            Pos = i
            Exit For
        End If
    Next i
    Debug.Print Pos & "番目 " & N & "回"
End Sub`;
  const r = runScript(src, { actions: [{ run: "Test" }] });
  assert.equal(r.error, null);
  assert.deepEqual(r.state.debug, ["5番目 5回"]);
});

test("二分探索（整列ずみの配列）", () => {
  const src = `Sub Test()
    Dim a(1 To 8) As Integer, Lo As Integer, Hi As Integer, Mid As Integer, Pos As Integer, N As Integer, i As Integer
    For i = 1 To 8
        a(i) = i * 10
    Next i
    Lo = 1: Hi = 8: Pos = 0
    Do While Lo <= Hi
        N = N + 1
        Mid = (Lo + Hi) \\ 2
        If a(Mid) = 70 Then
            Pos = Mid
            Exit Do
        ElseIf a(Mid) < 70 Then
            Lo = Mid + 1
        Else
            Hi = Mid - 1
        End If
    Loop
    Debug.Print Pos & "," & N
End Sub`;
  const r = runScript(src, { actions: [{ run: "Test" }] });
  assert.equal(r.error, null);
  assert.deepEqual(r.state.debug, ["7,3"]);
});

test("バブルソート・選択ソート・挿入ソートが、同じ結果になる", () => {
  const bubble = `Sub Test()
    Dim i As Integer, j As Integer, w As Integer, s As String
    ${NUMS}
    For i = 1 To 7
        For j = 1 To 8 - i
            If a(j) > a(j + 1) Then
                w = a(j): a(j) = a(j + 1): a(j + 1) = w
            End If
        Next j
    Next i
    ${showAll}
End Sub`;
  const select = `Sub Test()
    Dim i As Integer, j As Integer, m As Integer, w As Integer, s As String
    ${NUMS}
    For i = 1 To 7
        m = i
        For j = i + 1 To 8
            If a(j) < a(m) Then m = j
        Next j
        w = a(i): a(i) = a(m): a(m) = w
    Next i
    ${showAll}
End Sub`;
  const insert = `Sub Test()
    Dim i As Integer, j As Integer, w As Integer, s As String
    ${NUMS}
    For i = 2 To 8
        w = a(i)
        j = i - 1
        Do While j >= 1
            If a(j) <= w Then Exit Do
            a(j + 1) = a(j)
            j = j - 1
        Loop
        a(j + 1) = w
    Next i
    ${showAll}
End Sub`;
  const expect = ["1,2,3,4,5,7,8,9,"];
  for (const src of [bubble, select, insert]) {
    const r = runScript(src, { actions: [{ run: "Test" }] });
    assert.equal(r.error, null, r.error && r.error.message);
    assert.deepEqual(r.state.debug, expect);
  }
});

test("交換のとき、書きかわった2つの箱が step に記録される（表示で光らせる）", () => {
  const src = `Sub Test()
    Dim a(1 To 3) As Integer, w As Integer
    a(1) = 5: a(2) = 3
    w = a(1): a(1) = a(2): a(2) = w
End Sub`;
  const r = runScript(src, { actions: [{ run: "Test" }] });
  const changedKeys = r.session.steps
    .filter((st) => st.ops.some((o) => o[0][0] === "arrays" && o[0][2] === "data"))
    .map((st) => st.ops.filter((o) => o[0][0] === "arrays" && o[0][2] === "data").map((o) => o[0][3]).join(""));
  assert.deepEqual(changedKeys.slice(-2), ["1", "2"]);
});

test("サイコロ：Rnd のたねが同じなら、いつも同じ目（再生しても変わらない）", () => {
  const src = `Sub Test()
    Dim i As Integer
    For i = 1 To 5
        Debug.Print Int(Rnd * 6) + 1
    Next i
End Sub`;
  const a = runScript(src, { seed: 777, actions: [{ run: "Test" }] });
  const b = runScript(src, { seed: 777, actions: [{ run: "Test" }] });
  assert.deepEqual(a.state.debug, b.state.debug);
  assert.ok(a.state.debug.every((v) => Number(v) >= 1 && Number(v) <= 6));
});
