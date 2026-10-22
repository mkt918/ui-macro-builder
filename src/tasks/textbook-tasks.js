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

/* ===== オリジナル課題：割り勘アプリをつくろう =====
 * 教科書の 4-2（変数・計算・InputBox / MsgBox）と 4-3 の演算子（¥ と Mod）を使う発展課題。
 *   kind "guide"     … 解説つき。lesson（読みもの）を読みながら、そのとおりに作る
 *   kind "challenge" … 自分で考えて作る。lesson は「考え方」だけ
 * lesson: [{ h: 見出し, p: 本文, code: コード例 }]（どれも省略できる）
 * 教科書の課題ではないので code / page は無く、series / step を表示に使う。
 */
const WARIKAN_SERIES = "割り勘アプリをつくろう";

// 結果を Label3・Label4 に出す課題（ステップ3・発展）のフォーム
const warikanLabelForm = (cap1, cap2) =>
  ufForm(300, 210, [
    ufLabel("Label1", 24, 20, 110, 26, cap1, { FontSize: 14 }),
    ufLabel("Label2", 24, 70, 110, 26, cap2, { FontSize: 14 }),
    ufLabel("Label3", 144, 16, 120, 34, "", { SpecialEffect: 2, TextAlign: 1, FontSize: 16 }),
    ufLabel("Label4", 144, 66, 120, 34, "", { SpecialEffect: 2, TextAlign: 1, FontSize: 16 }),
    ufButton("CommandButton1", 90, 130, 120, 36, "計算する"),
  ]);

// ---- 回転ずしアプリ（皿ボタン4つ・枚数と小計と合計の Label、そのあと割り勘）----
const SUSHI_SERIES = "回転ずしアプリをつくろう";
const SUSHI_PLATES = [100, 120, 150, 200];
const SUSHI_DIMS = SUSHI_PLATES.map((p) => `Dim Mai${p} As Long`).join("\n") + "\nDim Gokei As Long";
const SUSHI_SUBS = SUSHI_PLATES.map(
  (p, i) =>
    `Private Sub CommandButton${i + 1}_Click()\n    Mai${p} = Mai${p} + 1\n    Gokei = Gokei + ${p}\n` +
    `    Label${i + 1}.Caption = "${p}円皿" & Mai${p} & "枚"\n    Label${i + 5}.Caption = "小計" & Mai${p} * ${p} & "円"\n` +
    `    Label9.Caption = "合計" & Gokei & "円"\nEnd Sub`
).join("\n");
// withWari: true なら、割り勘のボタンと Label10・Label11 まで置いたフォーム（模範解答用）
function sushiForm(withWari) {
  const sunken = { SpecialEffect: 2, TextAlign: 1, FontSize: 12 };
  const controls = [];
  SUSHI_PLATES.forEach((p, i) => {
    const y = 16 + 50 * i;
    controls.push(ufButton(`CommandButton${i + 1}`, 20, y, 120, 36, `${p}円皿`));
    controls.push(ufLabel(`Label${i + 1}`, 155, y, 100, 34, "", sunken));
    controls.push(ufLabel(`Label${i + 5}`, 265, y, 130, 34, "", sunken));
  });
  controls.push(ufLabel("Label9", 155, 216, 240, 34, "", sunken));
  if (withWari) {
    controls.push(ufButton("CommandButton5", 20, 276, 120, 36, "割り勘"));
    controls.push(ufLabel("Label10", 155, 276, 100, 34, "", sunken));
    controls.push(ufLabel("Label11", 265, 276, 130, 34, "", sunken));
  }
  return ufForm(420, withWari ? 340 : 270, controls);
}

const ORIGINAL_TASKS = [
  {
    id: "warikan1",
    kind: "guide",
    series: WARIKAN_SERIES,
    step: "ステップ1",
    title: "割り勘アプリ ステップ1　1人あたりの金額を出そう",
    difficulty: 1,
    goal: "文化祭の打ち上げで、ジュース・ピザ・ケーキを注文した。3つの値段と人数を入力すると、合計と1人あたりの金額を表示するプログラムを作ろう。",
    lesson: [
      { p: "打ち上げの会計をするアプリを、少しずつ作っていきます。まずは「みんなの注文をたして、人数でわる」ところから。下の解説を読みながら、順番に打ってみよう。" },
      {
        h: "① 値段を聞いて、変数にしまう",
        p: "InputBox で聞いた答えは「文字」です。Val( ) をつけると「数」になります。数を入れる箱は Dim で用意します（Long は大きめの整数が入る型です）。",
        code: 'Dim Juice As Long\nJuice = Val(InputBox("ジュースの値段を入力してください"))',
      },
      { h: "② たし算をして、合計を出す", p: "たし算は「+」です。答えは変数 Gokei（合計）にしまいます。", code: 'Gokei = Juice + Pizza + Cake\nMsgBox ("合計は" & Gokei & "円です")' },
      { h: "③ 人数でわり算をする", p: "わり算は「/」です。", code: 'Hitori = Gokei / Ninzu\nMsgBox ("1人" & Hitori & "円です")' },
      { h: "考えてみよう", p: "1000円を3人でわったら、いくらになるかな？ 割り切れないときの分け方は、次のステップで考えます。" },
    ],
    spec: {
      input: "ジュース・ピザ・ケーキの値段、人数　300、1200、900、3",
      output: "合計は2400円です／1人800円です",
      conditions: [
        "値段は変数 Juice・Pizza・Cake に、人数は変数 Ninzu に入れる（型は Long）。",
        "InputBox 関数で、「ジュースの値段を入力してください」「ピザの値段を入力してください」「ケーキの値段を入力してください」「人数を入力してください」の順にたずねる。",
        "3つの値段をたした合計を、変数 Gokei に入れ、MsgBox で「合計は Gokei 円です」と表示する。",
        "Gokei を Ninzu でわった金額を、変数 Hitori に入れ、MsgBox で「1人 Hitori 円です」と表示する。",
      ],
      inputExample: "InputBox に「300」「1200」「900」「3」の順に入力する",
      outputExample: "MsgBox に「合計は2400円です」、続けて「1人800円です」と表示される",
    },
    form: ufForm(300, 150, [ufButton("CommandButton1", 80, 50, 140, 36, "計算する")]),
    hints: [
      "ボタンを押したときの命令の中に、①値段を3つと人数を聞く ②たし算で合計を出して表示 ③わり算で1人あたりを出して表示、の順で組もう。",
      "変数は6つ：Juice・Pizza・Cake・Ninzu・Gokei・Hitori（Long型）。InputBox の答えは Val( ) で数にします。合計は「+」、1人あたりは「/」で計算。",
      'Gokei = Juice + Pizza + Cake ／ MsgBox ("合計は" & Gokei & "円です") ／ Hitori = Gokei / Ninzu ／ MsgBox ("1人" & Hitori & "円です")',
    ],
    answer: `Option Explicit
Private Sub CommandButton1_Click()
    Dim Juice As Long
    Dim Pizza As Long
    Dim Cake As Long
    Dim Ninzu As Long
    Dim Gokei As Long
    Dim Hitori As Long
    Juice = Val(InputBox("ジュースの値段を入力してください"))
    Pizza = Val(InputBox("ピザの値段を入力してください"))
    Cake = Val(InputBox("ケーキの値段を入力してください"))
    Ninzu = Val(InputBox("人数を入力してください"))
    Gokei = Juice + Pizza + Cake
    MsgBox ("合計は" & Gokei & "円です")
    Hitori = Gokei / Ninzu
    MsgBox ("1人" & Hitori & "円です")
End Sub`,
    scenario: [
      { click: "CommandButton1", inputs: ["300", "1200", "900", "3"], expect: { msgs: ["合計は2400円です", "1人800円です"] } },
      { click: "CommandButton1", inputs: ["500", "900", "100", "5"], expect: { msgs: ["合計は1500円です", "1人300円です"] } },
    ],
  },
  {
    id: "warikan2",
    kind: "guide",
    series: WARIKAN_SERIES,
    step: "ステップ2",
    title: "割り勘アプリ ステップ2　あまりのお金を考えよう",
    difficulty: 1,
    goal: "合計金額と人数を入力すると、1人あたりの金額と、あまりのお金を表示するプログラムを作ろう。",
    lesson: [
      { h: "割り切れないとき", p: "1000円を3人で分けると、1人333円で、1円あまります。このあまりまで考えるのが「割り勘」です。" },
      { h: "「¥」で、分けた金額を出す", p: "「¥」は、わり算の答えの整数の部分を出します。", code: "1000 ¥ 3　→　333" },
      { h: "「Mod」で、あまりを出す", p: "「Mod」は、わり算のあまりを出します。", code: "1000 Mod 3　→　1" },
      { h: "たしかめ算", p: "333 × 3 + 1 = 1000 になれば、正しく分けられています。" },
      { h: "プログラムにすると", code: "Hitori = Gokei ¥ Ninzu\nAmari = Gokei Mod Ninzu" },
      { h: "ちゅうい", p: "「/」と「¥」はちがいます。「/」は小数まで出て、整数の箱に入れると四捨五入されるので、合計とずれることがあります。割り勘では「¥」を使いましょう。" },
    ],
    spec: {
      input: "合計金額、人数　1000、3",
      output: "1人333円です／あまりは1円です",
      conditions: [
        "合計金額は変数 Gokei に、人数は変数 Ninzu に入れる（型は Long）。",
        "InputBox 関数で、「合計金額を入力してください」「人数を入力してください」の順にたずねる。",
        "「¥」で求めた1人あたりの金額を変数 Hitori に入れ、MsgBox で「1人 Hitori 円です」と表示する。",
        "「Mod」で求めたあまりを変数 Amari に入れ、MsgBox で「あまりは Amari 円です」と表示する。",
      ],
      inputExample: "InputBox に「1000」「3」の順に入力する",
      outputExample: "MsgBox に「1人333円です」、続けて「あまりは1円です」と表示される",
    },
    form: ufForm(300, 150, [ufButton("CommandButton1", 80, 50, 140, 36, "計算する")]),
    hints: [
      "ステップ1のプログラムが土台です。①合計金額と人数を聞く ②「¥」で1人あたり ③「Mod」であまり ④MsgBox で2回表示、の順で組もう。",
      "変数は4つ：Gokei・Ninzu・Hitori・Amari（Long型）。わり算は「/」ではなく「¥」を使います。",
      'Hitori = Gokei ¥ Ninzu ／ Amari = Gokei Mod Ninzu ／ MsgBox ("1人" & Hitori & "円です") ／ MsgBox ("あまりは" & Amari & "円です")',
    ],
    answer: `Option Explicit
Private Sub CommandButton1_Click()
    Dim Gokei As Long
    Dim Ninzu As Long
    Dim Hitori As Long
    Dim Amari As Long
    Gokei = Val(InputBox("合計金額を入力してください"))
    Ninzu = Val(InputBox("人数を入力してください"))
    Hitori = Gokei ¥ Ninzu
    Amari = Gokei Mod Ninzu
    MsgBox ("1人" & Hitori & "円です")
    MsgBox ("あまりは" & Amari & "円です")
End Sub`,
    scenario: [
      { click: "CommandButton1", inputs: ["1000", "3"], expect: { msgs: ["1人333円です", "あまりは1円です"] } },
      { click: "CommandButton1", inputs: ["3000", "3"], expect: { msgs: ["1人1000円です", "あまりは0円です"] } },
      { click: "CommandButton1", inputs: ["4580", "6"], expect: { msgs: ["1人763円です", "あまりは2円です"] } },
      // 1100÷3＝366.67…：「/」だと四捨五入で 367 になってしまう
      { click: "CommandButton1", inputs: ["1100", "3"], expect: { msgs: ["1人366円です", "あまりは2円です"] } },
    ],
  },
  {
    id: "warikan3",
    kind: "challenge",
    series: WARIKAN_SERIES,
    step: "ステップ3",
    title: "割り勘アプリ ステップ3　割り勘アプリを完成させよう",
    difficulty: 2,
    goal: "合計金額と人数を入力すると、1人あたりの金額とあまりを、画面の Label に表示する「割り勘アプリ」を完成させよう。",
    lesson: [
      {
        h: "考え方",
        p: "ステップ2のプログラムが、そのまま土台になります。ちがうのは、答えを MsgBox ではなく Label に表示するところだけ。Label に文字を入れるときは「Label3.Caption = …」と書きます（教科書 4-2-02 と同じ）。",
      },
    ],
    spec: {
      input: "合計金額、人数　4580、6",
      output: "1人あたり 763円／あまり 2円",
      conditions: [
        "ボタンを押すと、InputBox で「合計金額を入力してください」「人数を入力してください」の順にたずねる。",
        "合計金額は Gokei、人数は Ninzu、1人あたりは Hitori、あまりは Amari に入れる（型は Long）。",
        "Hitori と Amari は、ステップ2で使った計算（¥ と Mod）で求める。",
        "Label3 に1人あたりの金額を、Label4 にあまりを、それぞれ「円」をつけて表示する。MsgBox は使わない。",
      ],
      inputExample: "InputBox に「4580」「6」の順に入力する",
      outputExample: "Label3 に「763円」、Label4 に「2円」と表示される",
    },
    form: warikanLabelForm("1人あたり", "あまり"),
    hints: [
      "ステップ2のプログラムをもとに、MsgBox の部分を Label への表示に書きかえよう。",
      'Label3.Caption に Hitori と「円」を、Label4.Caption に Amari と「円」を入れます。「円」は & でつなげます（例：Hitori & "円"）。',
      'Hitori = Gokei ¥ Ninzu ／ Amari = Gokei Mod Ninzu ／ Label3.Caption = Hitori & "円" ／ Label4.Caption = Amari & "円"',
    ],
    answer: `Option Explicit
Private Sub CommandButton1_Click()
    Dim Gokei As Long
    Dim Ninzu As Long
    Dim Hitori As Long
    Dim Amari As Long
    Gokei = Val(InputBox("合計金額を入力してください"))
    Ninzu = Val(InputBox("人数を入力してください"))
    Hitori = Gokei ¥ Ninzu
    Amari = Gokei Mod Ninzu
    Label3.Caption = Hitori & "円"
    Label4.Caption = Amari & "円"
End Sub`,
    scenario: [
      { click: "CommandButton1", inputs: ["3000", "3"], expect: { controls: { Label3: { Caption: "1000円" }, Label4: { Caption: "0円" } } } },
      { click: "CommandButton1", inputs: ["1000", "3"], expect: { controls: { Label3: { Caption: "333円" }, Label4: { Caption: "1円" } } } },
      { click: "CommandButton1", inputs: ["4580", "6"], expect: { controls: { Label3: { Caption: "763円" }, Label4: { Caption: "2円" } } } },
      { click: "CommandButton1", inputs: ["500", "7"], expect: { controls: { Label3: { Caption: "71円" }, Label4: { Caption: "3円" } } } },
      // 1100÷3＝366.67…：「/」だと四捨五入で 367 になってしまう
      { click: "CommandButton1", inputs: ["1100", "3"], expect: { controls: { Label3: { Caption: "366円" }, Label4: { Caption: "2円" } } } },
    ],
  },
  {
    id: "warikan4",
    kind: "challenge",
    series: WARIKAN_SERIES,
    step: "発展",
    continueFrom: "warikan3",
    title: "割り勘アプリ 発展　幹事とほかの人の金額を出そう",
    difficulty: 3,
    goal: "あまったお金は、幹事（かんじ）が引き受けることにしよう。ステップ3で作ったフォームの下に Label を4つ足して、幹事が出す金額と、ほかの人が出す金額も表示するアプリにしよう。",
    lesson: [
      {
        h: "考え方",
        p: "4580円を6人で分けると、1人763円であまりは2円。ほかの5人が763円ずつ出すなら、幹事は 763 + 2 = 765円 出せば、ぴったり4580円になります。\nたしかめ算：765 + 763 × 5 = 4580",
      },
      {
        h: "フォームに足すもの（ここが新しい！）",
        p: "ステップ3のフォームはそのまま使います。ボタンを下にずらして、あいたところに Label を4つ足そう。\nLabel5 … Caption を「幹事」にする（見出し）\nLabel6 … 幹事の金額を出す枠\nLabel7 … Caption を「ほかの人」にする（見出し）\nLabel8 … ほかの人の金額を出す枠\nLabel は作った順に、Label5・Label6・Label7・Label8 という名前になります。",
      },
    ],
    spec: {
      input: "合計金額、人数　4580、6",
      output: "1人あたり 763円／あまり 2円／幹事 765円／ほかの人 763円",
      conditions: [
        "ステップ3のプログラムとフォームを、そのまま使う。",
        "フォームの下に Label を4つ足す。Label5 の Caption は「幹事」、Label7 の Caption は「ほかの人」にする。",
        "Label6 に幹事が出す金額を、Label8 にほかの人が出す金額を、それぞれ「円」をつけて表示する。",
        "幹事の金額は、「1人あたり」に「あまり」をたした金額。変数 Kanji に入れる（型は Long）。ほかの人の金額は、「1人あたり」と同じ。",
      ],
      inputExample: "InputBox に「4580」「6」の順に入力する",
      outputExample: "Label3 に「763円」、Label4 に「2円」、Label6 に「765円」、Label8 に「763円」と表示される",
    },
    form: warikanLabelForm("1人あたり", "あまり"),
    answerForm: ufForm(300, 290, [
      ufLabel("Label1", 24, 20, 110, 26, "1人あたり", { FontSize: 14 }),
      ufLabel("Label2", 24, 70, 110, 26, "あまり", { FontSize: 14 }),
      ufLabel("Label5", 24, 120, 110, 26, "幹事", { FontSize: 14 }),
      ufLabel("Label7", 24, 170, 110, 26, "ほかの人", { FontSize: 14 }),
      ufLabel("Label3", 144, 16, 120, 34, "", { SpecialEffect: 2, TextAlign: 1, FontSize: 16 }),
      ufLabel("Label4", 144, 66, 120, 34, "", { SpecialEffect: 2, TextAlign: 1, FontSize: 16 }),
      ufLabel("Label6", 144, 116, 120, 34, "", { SpecialEffect: 2, TextAlign: 1, FontSize: 16 }),
      ufLabel("Label8", 144, 166, 120, 34, "", { SpecialEffect: 2, TextAlign: 1, FontSize: 16 }),
      ufButton("CommandButton1", 90, 226, 120, 36, "計算する"),
    ]),
    hints: [
      "ステップ3のプログラムに、続きを書こう。①フォームに Label を4つ足す ②幹事の金額を計算する ③Label6 と Label8 に表示する、の順です。",
      "幹事の金額 ＝ 1人あたり ＋ あまり。Kanji という変数を用意して、Kanji = Hitori + Amari と計算します。ほかの人の金額は Hitori と同じです。",
      'Kanji = Hitori + Amari ／ Label6.Caption = Kanji & "円" ／ Label8.Caption = Hitori & "円"（Label5 の Caption は「幹事」、Label7 の Caption は「ほかの人」）',
    ],
    answer: `Option Explicit
Private Sub CommandButton1_Click()
    Dim Gokei As Long
    Dim Ninzu As Long
    Dim Hitori As Long
    Dim Amari As Long
    Dim Kanji As Long
    Gokei = Val(InputBox("合計金額を入力してください"))
    Ninzu = Val(InputBox("人数を入力してください"))
    Hitori = Gokei ¥ Ninzu
    Amari = Gokei Mod Ninzu
    Kanji = Hitori + Amari
    Label3.Caption = Hitori & "円"
    Label4.Caption = Amari & "円"
    Label6.Caption = Kanji & "円"
    Label8.Caption = Hitori & "円"
End Sub`,
    scenario: [
      ["3000", "3", "1000円", "0円", "1000円", "1000円"],
      ["1000", "3", "333円", "1円", "334円", "333円"],
      ["4580", "6", "763円", "2円", "765円", "763円"],
      ["1100", "3", "366円", "2円", "368円", "366円"],
    ].map((r) => ({
      click: "CommandButton1",
      inputs: [r[0], r[1]],
      expect: {
        controls: {
          Label3: { Caption: r[2] },
          Label4: { Caption: r[3] },
          Label5: { Caption: "幹事" },
          Label6: { Caption: r[4] },
          Label7: { Caption: "ほかの人" },
          Label8: { Caption: r[5] },
        },
      },
    })),
  },

  /* ---------- 回転ずしアプリをつくろう ---------- */
  {
    id: "sushi1",
    kind: "guide",
    series: SUSHI_SERIES,
    seriesIcon: "🍣",
    step: "ステップ1",
    title: "回転ずしアプリ ステップ1　ボタンで皿を数えよう",
    difficulty: 3,
    goal: "回転ずしのお会計アプリを作ろう。100円皿・120円皿・150円皿・200円皿のボタンを押すと、お皿の枚数・小計・合計が表示されるようにしよう。",
    lesson: [
      { p: "回転ずしでは、食べたお皿の数でお会計が決まります。ボタンを1回押すと、お皿が1枚ふえるアプリを作ります。" },
      {
        h: "① 押すたびに数がふえる箱",
        p: "Sub の外に Dim を書くと、ボタンを押し終わっても、箱の中の数が消えません。押すたびに 1 ずつふやすには「Mai100 = Mai100 + 1」と書きます。",
        code: "Dim Mai100 As Long\nPrivate Sub CommandButton1_Click()\n    Mai100 = Mai100 + 1\nEnd Sub",
      },
      { h: "② 合計の箱もふやす", p: "合計の箱 Gokei にも、100円皿なら 100 をたしていきます。", code: "Gokei = Gokei + 100" },
      { h: "③ 小計を出す", p: "小計は「枚数 × 1枚の値段」です。", code: "Mai100 * 100" },
      {
        h: "④ Label に表示する",
        p: "Label に文字を入れるには Caption を使います。「&」で文字と数をつなげます。",
        code: 'Label1.Caption = "100円皿" & Mai100 & "枚"\nLabel5.Caption = "小計" & Mai100 * 100 & "円"\nLabel9.Caption = "合計" & Gokei & "円"',
      },
      { h: "⑤ ほかのお皿も同じ", p: "120円皿・150円皿・200円皿のボタンも、100円皿と同じ作りです。数字と Label の名前だけ変えます。" },
    ],
    spec: {
      input: "ボタンを押す（100円皿、120円皿、150円皿、200円皿）",
      output: "100円皿1枚／小計100円／合計100円",
      conditions: [
        "Sub の外で、枚数の箱 Mai100・Mai120・Mai150・Mai200 と、合計の箱 Gokei を用意する（型は Long）。",
        "100円皿のボタン（CommandButton1）を押すと、Mai100 を 1 ふやし、Gokei に 100 をたす。",
        "Label1 に「100円皿○枚」、Label5 に「小計○円」（枚数×100）、Label9 に「合計○円」を表示する。",
        "120円皿（CommandButton2：Label2・Label6）、150円皿（CommandButton3：Label3・Label7）、200円皿（CommandButton4：Label4・Label8）も同じように作る。合計は、どのボタンでも Label9 に出す。",
      ],
      inputExample: "100円皿のボタンを2回、120円皿のボタンを1回押す",
      outputExample: "Label1 に「100円皿2枚」、Label5 に「小計200円」、Label2 に「120円皿1枚」、Label6 に「小計120円」、Label9 に「合計320円」と表示される",
    },
    form: sushiForm(false),
    hints: [
      "まず 100円皿のボタン1つだけを完成させよう。できたら、コピーして数字と Label の名前を変えるだけです。",
      "Sub の外に Dim Mai100 As Long と Dim Gokei As Long を書きます。ボタンの中では ①Mai100 を 1 ふやす ②Gokei に 100 をたす ③Label を3つ書きかえる、の順です。",
      'Mai100 = Mai100 + 1 ／ Gokei = Gokei + 100 ／ Label1.Caption = "100円皿" & Mai100 & "枚" ／ Label5.Caption = "小計" & Mai100 * 100 & "円" ／ Label9.Caption = "合計" & Gokei & "円"',
    ],
    answer: `Option Explicit
${SUSHI_DIMS}
${SUSHI_SUBS}`,
    scenario: [
      { click: "CommandButton1", expect: { controls: { Label1: { Caption: "100円皿1枚" }, Label5: { Caption: "小計100円" }, Label9: { Caption: "合計100円" } } } },
      { click: "CommandButton1", expect: { controls: { Label1: { Caption: "100円皿2枚" }, Label5: { Caption: "小計200円" }, Label9: { Caption: "合計200円" } } } },
      { click: "CommandButton2", expect: { controls: { Label2: { Caption: "120円皿1枚" }, Label6: { Caption: "小計120円" }, Label9: { Caption: "合計320円" } } } },
      { click: "CommandButton3", expect: { controls: { Label3: { Caption: "150円皿1枚" }, Label7: { Caption: "小計150円" }, Label9: { Caption: "合計470円" } } } },
      { click: "CommandButton4", expect: { controls: { Label4: { Caption: "200円皿1枚" }, Label8: { Caption: "小計200円" }, Label9: { Caption: "合計670円" } } } },
      { click: "CommandButton1", expect: { controls: { Label1: { Caption: "100円皿3枚" }, Label5: { Caption: "小計300円" }, Label9: { Caption: "合計770円" } } } },
    ],
  },
  {
    id: "sushi2",
    kind: "challenge",
    series: SUSHI_SERIES,
    seriesIcon: "🍣",
    step: "ステップ2",
    continueFrom: "sushi1",
    title: "回転ずしアプリ ステップ2　人数で割り勘にしよう",
    difficulty: 3,
    goal: "回転ずしアプリに「割り勘」ボタンを足そう。押すと人数をたずねて、合計を人数でわった「1人あたりの金額」を表示するアプリにしよう。",
    lesson: [
      {
        h: "考え方",
        p: "ステップ1のアプリをそのまま使い、続きを足します。割り勘の計算は、割り勘アプリで使った「¥」（整数の割り算）です。",
      },
      {
        h: "フォームに足すもの（ここが新しい！）",
        p: "CommandButton5 … Caption を「割り勘」にする\nLabel10 … 人数を出す枠\nLabel11 … 1人あたりの金額を出す枠",
      },
    ],
    spec: {
      input: "お皿のボタンを押す。そのあと「割り勘」ボタンを押して、人数を入力する　例：3",
      output: "人数3人／1人あたり106円",
      conditions: [
        "ステップ1のプログラムとフォームを、そのまま使う。",
        "フォームの下に「割り勘」ボタン（CommandButton5）と、Label10・Label11 を足す。",
        "割り勘ボタンを押すと、InputBox で「人数を入力してください」とたずね、変数 Ninzu（型は Long）に入れる。",
        "Gokei を Ninzu で「¥」でわった金額を、変数 Hitori（型は Long）に入れる。",
        "Label10 に「人数○人」、Label11 に「1人あたり○円」を表示する。",
      ],
      inputExample: "100円皿1回、120円皿1回、100円皿もう1回を押して合計320円にし、「割り勘」を押して、InputBox に「3」と入力する",
      outputExample: "Label10 に「人数3人」、Label11 に「1人あたり106円」と表示される",
    },
    form: sushiForm(false),
    answerForm: sushiForm(true),
    hints: [
      "お皿のボタンの部分は、ステップ1のままで大丈夫。割り勘ボタンの命令を足すだけです。",
      "CommandButton5_Click の中で、①InputBox で人数を聞く ②Hitori = Gokei ¥ Ninzu ③Label10 と Label11 に表示する、の順に書きます。Gokei は Sub の外にあるので、そのまま使えます。",
      'Ninzu = Val(InputBox("人数を入力してください")) ／ Hitori = Gokei ¥ Ninzu ／ Label10.Caption = "人数" & Ninzu & "人" ／ Label11.Caption = "1人あたり" & Hitori & "円"',
    ],
    answer: `Option Explicit
${SUSHI_DIMS}
${SUSHI_SUBS}
Private Sub CommandButton5_Click()
    Dim Ninzu As Long
    Dim Hitori As Long
    Ninzu = Val(InputBox("人数を入力してください"))
    Hitori = Gokei ¥ Ninzu
    Label10.Caption = "人数" & Ninzu & "人"
    Label11.Caption = "1人あたり" & Hitori & "円"
End Sub`,
    scenario: [
      { click: "CommandButton1", expect: { controls: { Label1: { Caption: "100円皿1枚" }, Label9: { Caption: "合計100円" } } } },
      { click: "CommandButton2", expect: { controls: { Label2: { Caption: "120円皿1枚" }, Label6: { Caption: "小計120円" }, Label9: { Caption: "合計220円" } } } },
      { click: "CommandButton1", expect: { controls: { Label1: { Caption: "100円皿2枚" }, Label9: { Caption: "合計320円" } } } },
      { click: "CommandButton5", inputs: ["3"], expect: { controls: { Label10: { Caption: "人数3人" }, Label11: { Caption: "1人あたり106円" } } } },
      { click: "CommandButton4", expect: { controls: { Label4: { Caption: "200円皿1枚" }, Label9: { Caption: "合計520円" } } } },
      { click: "CommandButton5", inputs: ["4"], expect: { controls: { Label10: { Caption: "人数4人" }, Label11: { Caption: "1人あたり130円" } } } },
    ],
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
  const form = opts.form || ((task.answerForm || task.form) ? toRuntime(task.answerForm || task.form) : null);
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

// 教科書の順 → オリジナル課題（割り勘アプリ）→ セルを使った練習
const TASKS = TEXTBOOK_TASKS.concat(ORIGINAL_TASKS, typeof CELL_TASKS !== "undefined" ? CELL_TASKS : require("./tasks.js").CELL_TASKS);

if (typeof module !== "undefined" && module.exports) module.exports = { TASKS, TEXTBOOK_TASKS, ORIGINAL_TASKS, runTaskScenario };
