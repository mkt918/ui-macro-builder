/* ===== 教科書（第4章 手続き型言語のプログラミング）の例題・実習問題 =====
 * フィールドの意味は src/tasks/tasks.js の先頭のコメントを参照。
 * このファイルが最終的な TASKS（教科書の順 → セルを使った練習）を作る。
 * 画面・Blockly には依存しない（Node のテストからも読める）。
 */

// フォームの部品を短く書くためのヘルパー
const uf = (name, type, left, top, width, height, props) => ({
  name, type, left, top, width, height, props: Object.assign({ FontSize: 12 }, props || {}),
});
const ufLabel = (name, left, top, width, height, caption, props) =>
  uf(name, "Label", left, top, width, height, Object.assign({ Caption: caption }, props || {}));
const ufButton = (name, left, top, width, height, caption) => uf(name, "CommandButton", left, top, width, height, { Caption: caption });
const ufText = (name, left, top, width, height) => uf(name, "TextBox", left, top, width, height, { Text: "" });
const ufForm = (width, height, controls) => ({ name: "UserForm1", caption: "UserForm1", width, height, controls });

// 例題 4-3-06 の入力ファイル（uriage.csv）と、期待する出力
const URIAGE_ROWS = [
  ["高橋　礼", 4], ["山口　桃子", 5], ["奥原　希美", 2], ["松友　美紀", 7], ["永原　和香", 8], ["松本　友子", 10],
  ["米本　小奈津", 23], ["田中　志保子", 5], ["東野　亜里沙", 13], ["栗原　歩美", 20], ["広田　花", 9],
];
const URIAGE_CSV = URIAGE_ROWS.map((r) => `${r[0]},${r[1]}`).join("\r\n") + "\r\n";
const withCommas = (n) => String(n).replace(/\B(?=(\d{3})+(?!\d))/g, ",");
const URIAGE_NAMES = URIAGE_ROWS.map((r) => r[0] + "\r\n").join("");
const URIAGE_KIN = URIAGE_ROWS.map((r) => withCommas(r[1] * 2000) + "\r\n").join("");

const SHAIN = [
  ["青木武", 10], ["伊藤晶子", 25], ["加藤秀彦", 3], ["田中雅代", 14], ["山本正雄", 7],
];

const TEXTBOOK_TASKS = [
  /* ---------- 第2節 データの入出力と演算 ---------- */
  {
    id: "rei4-2-03",
    kind: "example",
    code: "4-2-03",
    page: 91,
    title: "例題 4-2-03　数量から金額を出そう",
    difficulty: 1,
    goal: "単価が100円の商品を販売するとき、販売数を入力すると金額を表示するプログラムを作成しよう。",
    spec: {
      input: "販売数　5",
      output: "金額は 500 円です",
      conditions: [
        "単価は 100 円とし、変数 Tanka（整数型）に格納する。",
        "InputBox 関数を利用して「数量を入力してください」と表示し、販売数を入力する。入力された値は、変数 Suu（整数型）に格納する。",
        "Suu（数量）× Tanka（単価）で計算した値を、Kin（金額）に格納する。",
        "MsgBox 関数を利用して、「金額は Kin 円です」と表示する。",
      ],
      inputExample: "InputBox に「5」と入力する",
      outputExample: "MsgBox に「金額は500円です」と表示される",
    },
    form: ufForm(300, 150, [ufButton("CommandButton1", 80, 50, 140, 36, "金額を出す")]),
    hints: [
      "ボタンを押したときの命令（CommandButton1_Click）の中に、①単価を決める ②数量を聞く ③かけ算する ④表示する の順で組もう。",
      "変数（値を入れる箱）を3つ用意します：Tanka（単価）・Suu（数量）・Kin（金額）、型は整数型（Integer）。InputBox で聞いた答えは「文字」なので、Val( ) で数にしてからかけ算します。",
      'Tanka = 100 ／ Suu = Val(InputBox("数量を入力してください")) ／ Kin = Tanka * Suu ／ MsgBox ("金額は" & Kin & "円です")',
    ],
    answer: `Option Explicit
Private Sub CommandButton1_Click()
    Dim Tanka As Integer
    Dim Suu As Integer
    Dim Kin As Integer
    Tanka = 100
    Suu = Val(InputBox("数量を入力してください"))
    Kin = Tanka * Suu
    MsgBox ("金額は" & Kin & "円です")
End Sub`,
    scenario: [
      { click: "CommandButton1", inputs: ["5"], expect: { msgs: ["金額は500円です"] } },
      { click: "CommandButton1", inputs: ["3"], expect: { msgs: ["金額は300円です"] } },
    ],
  },
  {
    id: "jis4-2-01",
    kind: "practice",
    code: "4-2-01",
    page: 90,
    title: "実習問題 4-2-01　ありがとうメッセージ",
    difficulty: 2,
    goal: "「商品名」と「購入数」をインプットボックスに入力すると、「『商品名』を『購入数』個購入してくださり、ありがとうございます」と表示するプログラムを作成しなさい。",
    spec: {
      input: "商品名　メロンパン　／　購入数　4",
      output: "メロンパンを4個購入してくださり、ありがとうございます",
      conditions: [
        "商品名は変数 Hinmei（文字列型）に格納する。",
        "購入数は変数 Suu（整数型）に格納する。",
        "InputBox 関数を利用して、「商品名を入力してください」「購入数を入力してください」と表示し、商品名と購入数を入力する。",
        "MsgBox 関数を利用して、「『商品名』を『購入数』個購入してくださり、ありがとうございます」と表示する。",
      ],
      inputExample: "1回目の InputBox に「メロンパン」、2回目の InputBox に「4」と入力する",
      outputExample: "MsgBox に「メロンパンを4個購入してくださり、ありがとうございます」と表示される",
    },
    form: ufForm(300, 150, [ufButton("CommandButton1", 80, 50, 140, 36, "購入")]),
    hints: [
      "InputBox を2回使います。1回目は商品名、2回目は購入数をたずねます。",
      "商品名は文字列型（String）の Hinmei に、購入数は整数型（Integer）の Suu に入れます。購入数は Val( ) で数にしよう。",
      'MsgBox (Hinmei & "を" & Suu & "個購入してくださり、ありがとうございます")',
    ],
    answer: `Option Explicit
Private Sub CommandButton1_Click()
    Dim Hinmei As String
    Dim Suu As Integer
    Hinmei = InputBox("商品名を入力してください")
    Suu = Val(InputBox("購入数を入力してください"))
    MsgBox (Hinmei & "を" & Suu & "個購入してくださり、ありがとうございます")
End Sub`,
    scenario: [
      { click: "CommandButton1", inputs: ["メロンパン", "4"], expect: { msgs: ["メロンパンを4個購入してくださり、ありがとうございます"] } },
    ],
  },
  {
    id: "jis4-2-02",
    kind: "practice",
    code: "4-2-02",
    page: 90,
    title: "実習問題 4-2-02　日本語と英語のメニュー",
    difficulty: 2,
    goal: "コマンドボタンをクリックすると、「本日のおすすめメニュー」と「USサーロインステーキセット」の部分が日本語と英語にそれぞれ切り替わるユーザーフォームを作成しなさい。",
    spec: {
      input: "なし（ボタンをクリックする）",
      output: "日本語：本日のおすすめメニュー／USサーロインステーキセット　英語：Today's Special Menu／US Sirloin Steak Set",
      conditions: [
        "日本語ボタンをクリックすると、日本語で表示される。",
        "English ボタンをクリックすると、英語で表示される。",
        "終了ボタンをクリックすると、ユーザーフォームを閉じる。",
      ],
      inputExample: "「日本語」ボタン、「English」ボタンをクリックする",
      outputExample: "Label1 と Label2 の文字が切り替わる。「終了」でフォームが閉じる",
    },
    form: ufForm(340, 220, [
      ufLabel("Label1", 30, 20, 280, 30, "（ここにメニュー）", { TextAlign: 2, BorderStyle: 1, FontSize: 14 }),
      ufLabel("Label2", 30, 66, 280, 30, "（ここにメニュー）", { TextAlign: 2, BorderStyle: 1, FontSize: 14 }),
      ufButton("CommandButton1", 18, 130, 90, 36, "日本語"),
      ufButton("CommandButton2", 124, 130, 90, 36, "English"),
      ufButton("CommandButton3", 230, 130, 90, 36, "終了"),
    ]),
    hints: [
      "ボタンが3つあるので、「ボタンをクリックしたとき」の命令も3つ必要です。",
      "日本語・English のボタンでは、Label1.Caption と Label2.Caption に文字を入れます。終了のボタンでは、フォームを閉じる命令（Unload Me）を使います。",
      'CommandButton1_Click の中：Label1.Caption = "本日のおすすめメニュー" ／ Label2.Caption = "USサーロインステーキセット"。CommandButton2 は英語、CommandButton3 は Unload Me。',
    ],
    answer: `Option Explicit
Private Sub CommandButton1_Click()
    Label1.Caption = "本日のおすすめメニュー"
    Label2.Caption = "USサーロインステーキセット"
End Sub
Private Sub CommandButton2_Click()
    Label1.Caption = "Today's Special Menu"
    Label2.Caption = "US Sirloin Steak Set"
End Sub
Private Sub CommandButton3_Click()
    Unload Me
End Sub`,
    scenario: [
      { click: "CommandButton1", expect: { controls: { Label1: { Caption: "本日のおすすめメニュー" }, Label2: { Caption: "USサーロインステーキセット" } } } },
      { click: "CommandButton2", expect: { controls: { Label1: { Caption: "Today's Special Menu" }, Label2: { Caption: "US Sirloin Steak Set" } } } },
      { click: "CommandButton1", expect: { controls: { Label1: { Caption: "本日のおすすめメニュー" } } } },
      { click: "CommandButton3", expect: { closed: true } },
    ],
  },

  /* ---------- 第3節 条件判定と繰り返し処理 ---------- */
  {
    id: "rei4-3-01",
    kind: "example",
    code: "4-3-01",
    page: 103,
    title: "例題 4-3-01　郵便料金を表示しよう",
    difficulty: 3,
    goal: "定形郵便物を送るとき、重量を入力すると郵便料金を表示するプログラムを作成しよう。",
    spec: {
      input: "重量　22、25、40、65、0",
      output: "料金　84、84、94、重量オーバーです、終了",
      conditions: [
        "CommandButton1 をクリックすると、重量を入力する画面（インプットボックス）を表示する。",
        "入力された重量を Label3 に表示する。",
        "重量が 0 の場合は、処理を終了する。",
        "重量が 50 をこえた場合は、「重量オーバーです」というエラーのメッセージを表示して、処理を抜ける。",
        "重量が 25 以下の場合は料金を 84、50 以下の場合は 94 として、Label4 に表示する。",
      ],
      inputExample: "InputBox（題名：重量入力）に「22」と入力する",
      outputExample: "Label3（重量）に 22、Label4（料金）に 84 と表示される",
    },
    form: ufForm(300, 210, [
      ufLabel("Label1", 30, 20, 70, 26, "重量", { FontSize: 16 }),
      ufLabel("Label2", 30, 70, 70, 26, "料金", { FontSize: 16 }),
      ufLabel("Label3", 120, 16, 130, 34, "", { SpecialEffect: 2, TextAlign: 1, FontSize: 16 }),
      ufLabel("Label4", 120, 66, 130, 34, "", { SpecialEffect: 2, TextAlign: 1, FontSize: 16 }),
      ufButton("CommandButton1", 90, 130, 120, 36, "開始"),
    ]),
    hints: [
      "まず InputBox で重量を聞き、そのあとで「もし〜なら」を3回使って調べます。順番が大切：0 → 50をこえる → 25以下かどうか。",
      "重量が 0 のときは End（プログラムをおわる）。50 をこえたときは MsgBox を出してから Exit Sub（この命令のまとまりをぬける）。それ以外は、25 以下なら 84、ちがうなら 94 を料金の変数に入れます。",
      "If Omosa = 0 Then End ／ If Omosa > 50 Then（MsgBox と Exit Sub）／ If Omosa <= 25 Then Ryokin = 84 Else Ryokin = 94 ／ 最後に Label3.Caption = Omosa、Label4.Caption = Ryokin",
    ],
    answer: `Option Explicit
Private Sub CommandButton1_Click()
    Dim Omosa As Integer
    Dim Ryokin As Integer
    Omosa = Val(InputBox("重量を入力して下さい(g)", "重量入力", 0))
    If Omosa = 0 Then
        End
    End If
    If Omosa > 50 Then
        MsgBox ("重量オーバーです")
        Exit Sub
    End If
    If Omosa <= 25 Then
        Ryokin = 84
    Else
        Ryokin = 94
    End If
    Label3.Caption = Omosa
    Label4.Caption = Ryokin
End Sub`,
    scenario: [
      { click: "CommandButton1", inputs: ["22"], expect: { controls: { Label3: { Caption: "22" }, Label4: { Caption: "84" } } } },
      { click: "CommandButton1", inputs: ["25"], expect: { controls: { Label3: { Caption: "25" }, Label4: { Caption: "84" } } } },
      { click: "CommandButton1", inputs: ["40"], expect: { controls: { Label3: { Caption: "40" }, Label4: { Caption: "94" } } } },
      // 50 をこえたときは、メッセージを出して処理を抜ける＝重量・料金の表示は前のまま
      { click: "CommandButton1", inputs: ["65"], expect: { msgs: ["重量オーバーです"], controls: { Label3: { Caption: "40" }, Label4: { Caption: "94" } } } },
      { click: "CommandButton1", inputs: ["0"], expect: { closed: true } },
    ],
  },
  {
    id: "rei4-3-06",
    kind: "example",
    code: "4-3-06",
    page: 127,
    title: "例題 4-3-06　CSV から購入者一覧を作ろう",
    difficulty: 4,
    goal: "お菓子売場で取り扱う贈答用商品（単価2,000円）の売上資料から「購入者資料」を作成したい。購入者名と購入数が記録された売上資料は、すでに CSV ファイルに入力済みである。この CSV ファイルを用いて、一覧を表示するプログラムを作成しよう。",
    spec: {
      input: "入力ファイル：uriage.csv（購入者名，数量）　例：高橋　礼，4",
      output: "購入者名と金額の一覧　例：高橋　礼　8,000",
      conditions: [
        "CommandButton1 をクリックすると uriage.csv を読みこみ、入力された数量から金額を計算する。購入商品の単価は 2,000 円とし、金額は 数量 × 単価 で計算する。",
        "Label1 に購入者名を、Label2 に金額を表示する。",
        "入力データの末尾の判定のため、EOF 関数を利用する。",
      ],
      inputExample: "右の「📁 ファイル」タブに入っている uriage.csv（11人分）",
      outputExample: "Label1 に11人の名前、Label2 に 8,000 から 18,000 までの金額が、それぞれ1行ずつ表示される",
    },
    files: { "uriage.csv": URIAGE_CSV },
    form: ufForm(300, 330, [
      ufButton("CommandButton1", 40, 10, 200, 28, "入力"),
      ufLabel("Label1", 18, 52, 120, 232, "", { BorderStyle: 1 }),
      ufLabel("Label2", 150, 52, 110, 232, "", { BorderStyle: 1, TextAlign: 3 }),
    ]),
    hints: [
      "ファイルを使う手順は、①開く（Open）→ ②読みこむ（Input）→ ③閉じる（Close）。データがなくなるまで②をくり返します。",
      'Open ThisWorkbook.Path & "¥uriage.csv" For Input As #1 で開き、Do While Not EOF(1) 〜 Loop で、最後までくり返します。くり返しの中で Input #1, KouMei, Suu と読みこみます。',
      'くり返しの中：Kin = Suu * 2000 ／ Label1.Caption = Label1.Caption & KouMei & Chr(13) & Chr(10) ／ Label2.Caption = Label2.Caption & Format(Kin, "###,##0") & Chr(13) & Chr(10)',
    ],
    answer: `Option Explicit
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
End Sub`,
    scenario: [
      { click: "CommandButton1", expect: { controls: { Label1: { Caption: URIAGE_NAMES }, Label2: { Caption: URIAGE_KIN } } } },
    ],
  },

  /* ---------- 第4節 配列の利用 ---------- */
  {
    id: "rei4-4-01",
    kind: "example",
    code: "4-4-01",
    page: 141,
    title: "例題 4-4-01　売場番号から売場名を表示しよう",
    difficulty: 3,
    goal: "入力された売場番号の売場名を表示するプログラムを作成しよう。",
    spec: {
      input: "売場番号　1、2、3、4、5",
      output: "売場名　食料品、日用品、衣料品、家電、家具",
      conditions: [
        "配列 Umei に、売場名を格納する。",
        "TextBox1 に売場番号を入力し、CommandButton1 をクリックすると、入力された売場番号に対応した売場名を Label1 に表示する。",
        "配列 Umei は売場番号を添字として使用している。",
      ],
      inputExample: "TextBox1 に「1」と入力して「表示」ボタンを押す",
      outputExample: "Label1 に「食料品」と表示される",
    },
    form: ufForm(300, 160, [
      ufText("TextBox1", 20, 26, 110, 26),
      ufLabel("Label1", 160, 26, 110, 26, ""),
      ufButton("CommandButton1", 90, 80, 110, 34, "表示"),
    ]),
    hints: [
      "配列は、番号（添字）のついた箱を並べたものです。売場名を入れる配列 Umei を用意して、フォームが開いたときに、1〜5番の箱へ売場名を入れておきます。",
      "Sub の外に Dim Umei(5) As String と書くと、2つの命令（フォームが開いたとき・ボタンを押したとき）から同じ配列が使えます。ボタンを押したときは、TextBox1.Text を Val で数にして、その番号の箱を Label1 に表示します。",
      'UserForm_Initialize：Umei(1) = "食料品" 〜 Umei(5) = "家具" ／ CommandButton1_Click：Soe = Val(TextBox1.Text) → Label1.Caption = Umei(Soe)',
    ],
    answer: `Option Explicit
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
End Sub`,
    scenario: [
      { text: { TextBox1: "1" }, click: "CommandButton1", expect: { controls: { Label1: { Caption: "食料品" } } } },
      { text: { TextBox1: "3" }, click: "CommandButton1", expect: { controls: { Label1: { Caption: "衣料品" } } } },
      { text: { TextBox1: "5" }, click: "CommandButton1", expect: { controls: { Label1: { Caption: "家具" } } } },
    ],
  },
  {
    id: "jis4-4-01",
    kind: "practice",
    code: "4-4-01",
    page: 145,
    title: "実習問題 4-4-01　社員の氏名と勤続年数",
    difficulty: 3,
    goal: "指定した社員番号に該当する社員の氏名と勤続年数を表示するプログラムを作成しなさい。",
    spec: {
      input: "社員番号　1、2、3、4、5",
      output: "氏名・勤続年数　青木武 10年、伊藤晶子 25年、加藤秀彦 3年、田中雅代 14年、山本正雄 7年",
      conditions: [
        "配列 Smei に氏名を、配列 Knen に勤続年数を格納する。",
        "TextBox1 に社員番号を入力し、CommandButton1 をクリックすると、Label1 に氏名を、Label2 に勤続年数を表示する。",
        "配列 Smei と配列 Knen は社員番号を添字として使用している。",
      ],
      inputExample: "TextBox1 に「1」と入力して「表示」ボタンを押す",
      outputExample: "Label1 に「青木武」、Label2 に「10年」と表示される",
    },
    form: ufForm(320, 170, [
      ufText("TextBox1", 20, 20, 110, 26),
      ufLabel("Label1", 160, 20, 130, 26, ""),
      ufLabel("Label2", 160, 60, 130, 26, ""),
      ufButton("CommandButton1", 20, 90, 110, 34, "表示"),
    ]),
    hints: [
      "例題 4-4-01 と同じ作り方です。配列を2つ（氏名用・勤続年数用）用意します。",
      "Sub の外に Dim Smei(5) As String と Dim Knen(5) As Integer を書き、UserForm_Initialize で、1〜5番の箱に氏名と勤続年数を入れておきます。",
      'Label2.Caption = Knen(Soe) & "年"（& でつなげると「10年」のように表示できます）',
    ],
    answer:
      "Option Explicit\nDim Smei(5) As String\nDim Knen(5) As Integer\nPrivate Sub UserForm_Initialize()\n" +
      SHAIN.map((r, i) => `    Smei(${i + 1}) = "${r[0]}"\n    Knen(${i + 1}) = ${r[1]}\n`).join("") +
      'End Sub\nPrivate Sub CommandButton1_Click()\n    Dim Soe As Integer\n    Soe = Val(TextBox1.Text)\n    Label1.Caption = Smei(Soe)\n    Label2.Caption = Knen(Soe) & "年"\nEnd Sub',
    scenario: SHAIN.map((r, i) => ({
      text: { TextBox1: String(i + 1) },
      click: "CommandButton1",
      expect: { controls: { Label1: { Caption: r[0] }, Label2: { Caption: r[1] + "年" } } },
    })),
  },
];

/* ===== 台本（scenario）による自動判定 =====
 *   runTaskScenario(task, source, { form, files, initialCells }) → { passed, failures: [文字], error }
 *   form は src/userform.js の toRuntimeForm() の形。省略すると task.form から作る。
 */
function runTaskScenario(task, source, opts) {
  opts = opts || {};
  const S = typeof VbaSession !== "undefined" ? VbaSession : require("../vba/session.js");
  const toRuntime = (f) => ({
    name: f.name,
    caption: f.caption,
    controls: f.controls.map((c) => ({
      name: c.name,
      type: c.type,
      props: Object.assign({}, c.props, { Left: c.left, Top: c.top, Width: c.width, Height: c.height }),
    })),
  });
  const form = opts.form || (task.form ? toRuntime(task.form) : null);
  const files = opts.files || task.files || {};
  const scenario = task.scenario || [];
  const actions = [{ showForm: true }].concat(scenario.map((s) => ({ text: s.text, inputs: s.inputs, click: s.click })));
  let run;
  try {
    run = S.runScript(source, { form, files, initialCells: opts.initialCells, seed: 1, actions });
  } catch (e) {
    return { passed: false, failures: [e.friendly || e.message], error: e };
  }
  const failures = [];
  const label = (k) => {
    const s = scenario[k];
    const parts = [];
    if (s.text) parts.push(Object.keys(s.text).map((n) => `${n} に「${s.text[n]}」`).join("、"));
    if (s.inputs && s.inputs.length) parts.push(`入力「${s.inputs.join("」「")}」`);
    return `${k + 1}回目（${parts.join("、") || s.click}）`;
  };
  if (run.error) {
    failures.push(`実行中にエラーが出ました：${run.error.friendly || run.error.message}`);
    return { passed: false, failures, error: run.error };
  }
  for (let k = 0; k < scenario.length; k++) {
    const res = run.results[k + 1];
    if (!res) {
      failures.push(`${label(k)}の前に、フォームが閉じてしまいました`);
      break;
    }
    if (res.status === "waiting") {
      failures.push(`${label(k)}：InputBox が、台本より多く聞かれています`);
      break;
    }
    const model = run.session.modelAt(res.stepsAfter - 1);
    const exp = scenario[k].expect || {};
    if (exp.msgs) {
      const got = res.newMsgs;
      if (JSON.stringify(got) !== JSON.stringify(exp.msgs)) {
        failures.push(
          `${label(k)}：メッセージは「${exp.msgs.join("」「")}」になるはずですが、${got.length ? "「" + got.join("」「") + "」" : "何も出ていません"}`
        );
      }
    } else if (res.newMsgs.length) {
      failures.push(`${label(k)}：出なくてよいメッセージ「${res.newMsgs[0]}」が出ています`);
    }
    if (exp.controls) {
      Object.keys(exp.controls).forEach((cn) => {
        const c = (model.controls || {})[cn];
        Object.keys(exp.controls[cn]).forEach((prop) => {
          const want = String(exp.controls[cn][prop]);
          const got = c ? String(c[prop] === undefined ? "" : c[prop]) : undefined;
          if (got !== want) {
            const show = (t) => (t === undefined ? "（部品がありません）" : "「" + t.replace(/\r?\n/g, "↵") + "」");
            failures.push(`${label(k)}：${cn} の ${prop} は ${show(want)} になるはずですが、${show(got)} です`);
          }
        });
      });
    }
    if (exp.closed !== undefined && (model.form.visible === false) !== !!exp.closed) {
      failures.push(`${label(k)}：フォームは${exp.closed ? "閉じる" : "開いたまま"}はずです`);
    }
  }
  return { passed: failures.length === 0, failures };
}

// 教科書の順に、そのあとにセルを使った練習
const TASKS = TEXTBOOK_TASKS.concat(typeof CELL_TASKS !== "undefined" ? CELL_TASKS : require("./tasks.js").CELL_TASKS);

if (typeof module !== "undefined" && module.exports) module.exports = { TASKS, TEXTBOOK_TASKS, runTaskScenario };
