/* ===== 課題（クエスト）データベース =====
 * 新しいクエストを追加するときは、この配列に1つオブジェクトを足すだけ。
 * Claude に「こういうクエストを追加で」と頼めば、この形式で追記される。
 *
 * フィールド:
 *   id        : 一意のID（重複NG）
 *   title     : クエスト名
 *   difficulty: 難易度（1〜5）→ ★で表示
 *   goal      : 目標（複数行OK、\n で改行）
 *   hints     : ヒント3段階の配列（段階的に詳しく）
 *   answer    : 模範解答の VBA コード（答えを見るで表示）
 *   check     : (model) => boolean  クリア自動判定。
 *               model = { cells:{"A1":{value,bg,bold}}, arr:{1:v}, sheet, sheets }
 *               省略すると自動判定なし。
 */

// クリア判定用ヘルパー
function cellVal(model, addr) {
  const c = model.cells[addr];
  return c && c.value !== undefined ? c.value : "";
}
function cellBg(model, addr) {
  const c = model.cells[addr];
  return c ? c.bg : undefined;
}

const TASKS = [
  {
    id: "t01",
    title: "セルに名前を書こう",
    difficulty: 1,
    goal:
      "セル A1 に自分の名前を入力してみよう。\n\n" +
      "【使うブロック】\n" +
      "・📝 セルに値を入力\n" +
      "・\" \" テキスト",
    hints: [
      "「セル操作」カテゴリの『セルに値を入力』ブロックを使います。",
      "セルのところに A1 と入れて、右に『値』ブロックをつなげます。",
      "「値」カテゴリの \"テキスト\" ブロックをつなげて、名前を打ち込みましょう。",
    ],
    answer: 'Sub MyMacro()\n    Range("A1").Value = "山田太郎"\nEnd Sub',
    check: (m) => String(cellVal(m, "A1")).length > 0,
    goalPreview: [{ addr: "A1", value: "なまえ" }],
  },
  {
    id: "t02",
    title: "1から5までを縦に並べよう",
    difficulty: 2,
    goal:
      "繰り返しを使って、A1〜A5 に 1,2,3,4,5 を入れよう。\n\n" +
      "【ねらい】\n" +
      "ループ（くり返し）の感覚をつかむ。\n\n" +
      "【ヒント】カウンタ i が 1→2→3… と変わります。",
    hints: [
      "「繰り返し」の『○回繰り返す』ブロックを使い、5回にします。",
      "セルのアドレスに A{i} と書くと、i 行目のA列になります。",
      "入力する値には「繰り返し」カテゴリの『カウンタ i』ブロックを使いましょう。",
    ],
    answer:
      "Sub MyMacro()\n    Dim i As Integer\n    For i = 1 To 5\n        Range(\"A\" & i).Value = i\n    Next i\nEnd Sub",
    check: (m) =>
      [1, 2, 3, 4, 5].every((n) => String(cellVal(m, "A" + n)) === String(n)),
    goalPreview: [1, 2, 3, 4, 5].map((n) => ({ addr: "A" + n, value: n })),
  },
  {
    id: "t03",
    title: "偶数のセルだけ緑に塗ろう",
    difficulty: 3,
    goal:
      "1〜6行目をループし、行番号が偶数のときだけ B 列を緑に塗ろう。\n\n" +
      "【ねらい】\n" +
      "ループ + 条件分岐 の組み合わせ。\n\n" +
      "※偶数判定は (i ÷ 2 の余り) を考えますが、\n" +
      "まずは『i = 2,4,6 のとき』のように比較でもOK。",
    hints: [
      "『行1から6までループ』の中に『もし〜なら』を入れます。",
      "条件は『カウンタ i = 2』のような比較ブロックを使います。",
      "もし〜なら の中に『背景色を変える』を入れ、B{i} を緑にします。",
    ],
    answer:
      'Sub MyMacro()\n    Dim i As Integer\n    For i = 1 To 6\n        If i Mod 2 = 0 Then\n            Range("B" & i).Interior.Color = RGB(33, 115, 70)\n        End If\n    Next i\nEnd Sub',
    check: (m) =>
      [2, 4, 6].every((n) => cellBg(m, "B" + n)) &&
      [1, 3, 5].every((n) => !cellBg(m, "B" + n)),
    goalPreview: [2, 4, 6].map((n) => ({ addr: "B" + n, value: "", bg: "#217346" })),
  },
  {
    id: "t04",
    title: "配列に九九を貯めよう",
    difficulty: 3,
    goal:
      "配列を使って、1〜5番目に 5×1, 5×2 … 5×5（5,10,15,20,25）を入れよう。\n\n" +
      "【ねらい】\n" +
      "配列（はこの列）にデータをためる感覚をつかむ。\n\n" +
      "下の📦配列ビューに数字がたまっていきます。",
    hints: [
      "『5回繰り返す』の中に『配列の○番目に入れる』を入れます。",
      "配列の番目には『カウンタ i』を使います。",
      "入れる値は『i × 5』。値カテゴリの計算ブロックで i と 5 をかけ算します。",
    ],
    answer:
      "Sub MyMacro()\n    Dim i As Integer\n    Dim arr(1 To 100) As Variant\n    For i = 1 To 5\n        arr(i) = i * 5\n    Next i\nEnd Sub",
    check: (m) =>
      [1, 2, 3, 4, 5].every((n) => Number(m.arr[n]) === n * 5),
  },
  {
    id: "t05",
    title: "配列から表へ書き出そう",
    difficulty: 4,
    goal:
      "まず配列の1〜5番目に 10,20,30,40,50 を入れ、\n" +
      "次にその配列の中身を A1〜A5 に書き出そう。\n\n" +
      "【ねらい】\n" +
      "配列 → セルへの転記。2つのループに分けて考える練習。",
    hints: [
      "前半: ループで『配列の i 番目に i×10 を入れる』。",
      "後半: もう1つループを置き、『セル A{i} に値を入力』。",
      "入力する値に『配列の i 番目の値』ブロックを使います。",
    ],
    answer:
      "Sub MyMacro()\n    Dim i As Integer\n    Dim arr(1 To 100) As Variant\n    For i = 1 To 5\n        arr(i) = i * 10\n    Next i\n    For i = 1 To 5\n        Range(\"A\" & i).Value = arr(i)\n    Next i\nEnd Sub",
    check: (m) =>
      [1, 2, 3, 4, 5].every((n) => String(cellVal(m, "A" + n)) === String(n * 10)),
    goalPreview: [1, 2, 3, 4, 5].map((n) => ({ addr: "A" + n, value: n * 10 })),
  },
];
