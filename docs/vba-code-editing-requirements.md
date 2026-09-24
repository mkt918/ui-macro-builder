# VBAコード直打ち機能 — 要件定義

対象: 「🤖 VBAコード」タブでコードを直接編集できるようにし、ブロックと同期させる機能。
関連: [PLAN.md](../PLAN.md)、[vba-code-editing-tasks.md](vba-code-editing-tasks.md)

---

## 1. 背景・目的

現状、VBAコードタブは**表示専用**（ブロックから生成されたコードを読むだけ）。
生徒がコードを直接触って試行錯誤できるようにし、ブロック操作とコード編集を
どちらからでも学習を進められるようにする。

## 2. スコープ

### 2.1 含むもの

| # | 機能 | 内容 |
|---|------|------|
| R1 | コード直接編集 | VBAコードタブを編集可能なエディタにする |
| R2 | 自動インデント | For/If/Do 等のブロック構文で改行時に自動でインデントを調整 |
| R3 | 入力アシスト | `For` `While` `If` 等のキーワードからスケルトン（対応する `Next`/`Loop`/`End If` まで込み）を展開 |
| R4 | ブロック → コード同期 | 既存機能（そのまま維持）。ブロック編集は即座にコードへ反映 |
| R5 | コード → ブロック同期 | 手打ちしたコードを解析し、ブロックを再構築する |

### 2.2 含まないもの（明示的に対象外）

- VBAの完全な文法サポート（`Select Case` / `With` / `Function` 定義 / クラス等）
- キー入力ごとのリアルタイム構文解析（→ 2.4 参照）
- ブロックにない任意の変数型・演算子・組み込み関数

## 3. 対応する VBA 構文の範囲

**「現状のブロック語彙に厳密に限定する」**（ユーザー確認済み）。
具体的には以下のブロックが生成できる構文のみサポートする。

| カテゴリ | ブロック | 対応する VBA 構文 |
|---|---|---|
| セル操作 | cell_set_value / cell_get_value / cell_copy / cell_clear / cell_clear_row | `Range("A1").Value = ...` / `.Value` 読み取り / `.Copy` / `.ClearContents` / `Rows(n).ClearContents` |
| セル操作（動的） | 上記の `{i}` 形式 | `Range("A" & i)` |
| Cells指定 | cells_set_value / cells_get_value / cells_clear / cells_bgcolor / cells_bold | `Cells(r, c).Value` 等 |
| 繰り返し | loop_repeat / loop_range / loop_for_step / loop_while / loop_do_until | `For i = a To b [Step s]` / `Do While` / `Do Until` |
| 条件分岐 | cond_if / cond_compare / cond_is_even | `If ... Then ... [Else ...] End If` / 比較演算子 |
| 書式設定 | fmt_bgcolor / fmt_bold / fmt_fontsize / range_border / range_select | `Interior.Color = RGB(...)` / `Font.Bold` / `Font.Size` / `Borders.LineStyle` / `.Select` |
| 入出力 | io_inputbox / io_msgbox | `InputBox(...)` / `MsgBox ...` |
| 変数 | var_set / var_get / var_change | 代入 / 参照 / `x = x + n` パターン |
| 配列 | array_set / array_get | `arr(i) = ...` / `arr(i)` |
| シート | sheet_add / sheet_select | `Worksheets.Add.Name` / `Worksheets("x").Select` |
| 値・演算 | value_number / value_text / value_math / text_concat / math_round / value_date | リテラル・四則演算・`Mod`・文字列連結(`&`)・`WorksheetFunction.Round`・`Date` |

対応外の構文を検出した場合は **エラー表示のみ行い、ブロックは直前の状態を維持する**（後述 4.2）。

## 4. 同期の仕様

### 4.1 ブロック → コード（既存・維持）

ブロックが変更されるたびに即座にコード欄へ反映する（リアルタイム）。
ただし、コード編集中（後述の「コード側が編集済み」状態）は上書きしない。

### 4.2 コード → ブロック（新規）

**同期タイミング**: タブ切替／エディタからのフォーカス離脱時（ユーザー確認済み）。
キー入力のたびには解析・反映しない。

処理フロー:
1. コード欄で1文字でも変更 → 「コード編集済み」フラグを立てる（ブロック側からの上書きを止める）
2. コードタブを離れる、またはブロックタブに切り替える → 解析を実行
3. **解析成功** → ワークスペースをクリアして解析結果のブロックで再構築し、整列。「コード編集済み」フラグを解除
4. **解析失敗（対応外の構文・文法エラー）** → エラーメッセージ（行番号付き）を表示。**ブロックは直前の状態のまま変更しない**。「コード編集済み」フラグは維持（コードは消さない。生徒が直せるようにする）

### 4.3 変数の扱い

コード中で新しく登場した変数名は `workspace.createVariable()` で自動的にブロック側の変数として登録する。
ひらがなを含む変数名はブロックのバリデーションと同様にエラーとする。

## 5. エディタの要件

| # | 要件 |
|---|------|
| E1 | シンタックスハイライト（VBAキーワード・文字列・数値・コメント） |
| E2 | 自動インデント（4スペース、`For`/`Do`/`If...Then` の後で+1、`Next`/`Loop`/`End If`/`Else` で-1） |
| E3 | 入力アシスト（`For`/`Dim`/`While`/`Until`/`If`/`Ifelse` を打つと、キーワード入力に応じて例文候補（見出し＋一言説明）が自動でポップアップ表示される。矢印キー＋Enter/クリックで選ぶとその場で展開される。Tabキーでの直接展開（ポップアップを経由しない打ち方）も維持） |
| E4 | ダーク/ライトモード両対応（既存のCatppuccinパレットに合わせる） |
| E5 | 解析エラーはエディタ内の該当行が分かるように表示する |

## 6. 非機能要件

- ビルド不要（既存方針を踏襲。CDN経由のライブラリ追加はOK）
- パーサーはBlocklyに依存しない純粋なJSとし、Node単体でテスト可能にする
- 既存の6課題（`src/tasks/tasks.js`）の模範解答コードが全て解析できることを回帰テストとする

## 7. 既知の制約（ユーザー合意済み）

- 対応外構文（`Select Case` 等）は変換されず、エラー表示のみ
- 完全な双方向ラウンドトリップの視覚的忠実性は保証しない（生成されるVBAとしては等価）
  - `For i = 1 To N` は専用の `loop_repeat`（○回繰り返す）ではなく `loop_range`
    （開始・終わりを指定）として再構築される
  - `(x Mod 2 = 0)` は専用の `cond_is_even` ブロックではなく `cond_compare` +
    `value_math` の組み合わせとして再構築される（括弧の位置だけ変わるが、
    VBAの演算子優先順位により意味は同一）
