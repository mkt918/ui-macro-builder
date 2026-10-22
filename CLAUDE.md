# UI Macro Builder — プロジェクトメモ

授業用の Excel VBA ビジュアルプログラミング学習ツール。
詳細な計画は [PLAN.md](PLAN.md) を参照。

## 会話・作業のまとめ保存
このプロジェクトには会話ログ用の記録規約が無い（`PLAN.md`・`docs/vba-code-editing-*.md`は
設計・タスク管理用のドキュメント）。ユーザー共通skill `save`（`~/.claude/skills/save/`）で
「ここまでまとめて」等と頼まれた場合は、`.claude/summaries/YYYY-MM-DD_見出し.md` に
保存する（汎用フォーマット）。

## 大改修中（教科書準拠のユーザーフォーム対応）

計画は [docs/textbook-userform-plan.md](docs/textbook-userform-plan.md)（マイルストーン M0〜M8）。
対象は中学1年生を含むプログラミング初学者。画面の文言・エラーはやさしい言葉にする。

### 新しい実行エンジン `src/vba/`（DOM 非依存。ブラウザと Node の両方で動く）

- `lexer.js` — 字句解析。全角記号の検出（`zenkaku`）と `fixZenkaku`。`\`・`¥`・`￥` は整数除算
- `ast-parser.js` — 構文解析（AST）。演算子の優先順位は教科書 p.93 の表どおり（`* / ¥ Mod` は同順位）
- `runtime.js` — 実行エンジン（ジェネレータ）。MsgBox / InputBox で yield して一時停止する。
  状態の変化は step（差分 ops ＋ 200 step ごとの全体の写し kf）で記録し、`StepStore.modelAt(i)` で各時点を復元する
- `session.js` — 実行の管理（`start` / `resume` / `showForm` / `click` / `setText`、区間 `segments`）と、
  テスト・課題判定用の `runScript`
- テスト: `node --test tests/*.test.js`（教科書の例題 4-2-03・4-3-01・4-3-06・4-4-01 などを含む）

## 構成（ビルド不要・Blockly / CodeMirror は vendor/ に同梱）

- `index.html` — エントリーポイント
- `vendor/` — Blockly 11.2.1・CodeMirror 5.65.16（学校のネットワークで CDN が止められても動くように同梱）
- `src/blockly-setup.js` — カスタムブロック定義 + ツールボックス
- `src/vba-generator.js` — ブロック → VBA コード生成
- `src/vba-parser.js` — VBA コード → ブロック 変換（手打ちコードの解析）
- `src/excel-preview.js` — 実行の記録（step）を再生して仮想 Excel・変数の箱・配列の箱を描く（`ExcelView`）
- `src/vba-dialog.js` — InputBox / MsgBox のダイアログ（Windows の標準ダイアログと同じ見た目）
- `src/userform.js` — 🪟 ユーザーフォームのデザイナー兼実行画面（`UserFormPanel`）。フォームのデータは `umb_forms` に保存
- `src/tour.js` — はじめてガイド（初回だけ自動表示。「❓ 使い方」から再表示）
- 保存：ブラウザ（localStorage の `umb_blocks` / `umb_code` / `umb_forms` / `umb_initial`）に自動保存。
  「💾 保存」「📂 開く」で `.vbalearn`（JSON）ファイル、「📤 先生に送る」で同じ内容のリンク（`#share=`）
- `src/tasks/tasks.js` — 課題データ（`TASKS` 配列）
- `src/app.js` — 全体の結線

## VBAコード直打ち機能

「🤖 VBAコード」タブはCodeMirrorベースの編集可能なエディタで、ブロックと双方向に同期する。
詳細仕様は [docs/vba-code-editing-requirements.md](docs/vba-code-editing-requirements.md)、
実装状況は [docs/vba-code-editing-tasks.md](docs/vba-code-editing-tasks.md) を参照。

- **ブロック → コード**: ブロック変更のたびにリアルタイム反映
- **コード → ブロック**: タブ切替 / エディタからのフォーカス離脱時に `src/vba-parser.js` で解析し、ブロックにも並べる。
  書いたコードは書き換えない（実行はコードのまま）。対応外の構文（`Select Case` 等）を含むときは、
  ブロックタブに「ブロックでは表せない書き方がある」案内を出し、コードのまま実行する
- 対応するVBA構文は現行ブロック語彙に厳密に限定（新しいブロックを追加したら
  `src/vba-parser.js` の `SIMPLE_RULES` / `ExprParser` にも対応するパースルールを追加すること）

## ローカル起動

```
npx http-server -p 8778 -c-1
```
→ http://localhost:8778

## 課題を追加するときの依頼テンプレート

Claude に以下の形式で頼めば `src/tasks/tasks.js` の `TASKS` 配列に追記されます。

```
こういう課題を追加して：
- タイトル: （課題名）
- 難易度: （1〜5）
- 目標: （生徒に何をさせたいか。複数行OK）
- 使ってほしいブロック: （あれば）
- 答え: （イメージがあれば。無ければClaudeが作成）
```

ヒントは3段階（ふんわり → 具体的 → ほぼ答え）で Claude が自動生成します。
模範解答は VBA コード（`answer`）とブロック配置（`answerBlocks`）の両方を用意します。「答えを見る」でブロックがワークスペースに読み込まれます。

### `check` 関数を書くときの注意（重要）

`check(model)` は **`src/app.js` の `hasExecutableProgram()` が true のとき、
つまりブロックが実際に実行された結果の `model` が得られたときだけ**呼ばれる。
セルへの直接入力（ダブルクリック入力）だけを見て判定しているわけではないので、
`check` 自身はシンプルに「結果の状態」だけを見て書いてよい
（例: `cellVal(m, "A1") !== ""` のような書き方で問題ない。直接入力だけで
すり抜けられる心配はしなくてよい）。

ただし、**セルへの直接入力を課題の一部として要求する課題**（例: t06「まずA1に
数字を入力してから」）では、その入力用セルの値そのものは直接入力由来のままになる
点に注意。`check` はあくまで「ブロックを実行した後の最終状態」を見るので、
直接入力したセルの値と、ブロックが書き込んだセルの値の**組み合わせ**を条件にすること。

## ブロックを追加するとき

1. `src/blockly-setup.js` にブロック定義とツールボックス登録
2. `src/vba-generator.js` に VBA 生成関数（`vbaGenerator.forBlock["型名"]`）。
   くり返しのカウンタは `loopBody()` / `currentCounter()` を使う（入れ子で i→j→k）。文字列は `vbaStr()` で囲む
3. `src/vba-parser.js` に対応するパースルールを追加（コード直打ちからも
   同じブロックを作れるようにする。文なら `SIMPLE_RULES` に正規表現ルールを、
   値を返す式なら `ExprParser.parseIdentChain` に分岐を足す）

実行は生成した VBA コードを `src/vba/`（新エンジン）がそのまま動かすので、ブロックごとの実行ロジックは不要。
生成したコードの書き方を新エンジンがまだ知らない場合だけ、`src/vba/runtime.js` に足してテストを書く。

### 実行の流れ（app.js）

- 実行するのは常に VBAコードタブの中身（`currentSource()`）。ブロックが正（`codeOwner === "blocks"`）のときは
  ブロックから作ったコードが入る。生徒がコードを打つと `codeOwner = "code"` になり、書いたコードは書き換えない
- 編集中は `runPreview()`（MsgBox は自動で OK、InputBox は前の答えがあれば使う）。▶実行・1コマは `startRun()` →
  再生しながら、記録の最後で答え待ちなら `VbaDialog`（src/vba-dialog.js）で聞いて `session.resume()`
