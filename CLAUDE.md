# UI Macro Builder — プロジェクトメモ

授業用の Excel VBA ビジュアルプログラミング学習ツール。
詳細な計画は [PLAN.md](PLAN.md) を参照。

## 構成（ビルド不要・Blockly / CodeMirror CDN）

- `index.html` — エントリーポイント
- `src/blockly-setup.js` — カスタムブロック定義 + ツールボックス
- `src/vba-generator.js` — ブロック → VBA コード生成
- `src/vba-parser.js` — VBA コード → ブロック 変換（手打ちコードの解析）
- `src/excel-preview.js` — ブロックを解釈実行して仮想 Excel をアニメーション
- `src/tasks/tasks.js` — 課題データ（`TASKS` 配列）
- `src/app.js` — 全体の結線

## VBAコード直打ち機能

「🤖 VBAコード」タブはCodeMirrorベースの編集可能なエディタで、ブロックと双方向に同期する。
詳細仕様は [docs/vba-code-editing-requirements.md](docs/vba-code-editing-requirements.md)、
実装状況は [docs/vba-code-editing-tasks.md](docs/vba-code-editing-tasks.md) を参照。

- **ブロック → コード**: ブロック変更のたびにリアルタイム反映
- **コード → ブロック**: タブ切替 / エディタからのフォーカス離脱時に `src/vba-parser.js` で解析し反映。
  対応外の構文（`Select Case` 等、現行ブロック語彙に無いもの）はエラー表示のみでブロックは維持
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
2. `src/vba-generator.js` に VBA 生成関数（`vbaGenerator.forBlock["型名"]`）
3. `src/excel-preview.js` の `Interpreter.execBlock` / `evalValue` に実行ロジック
4. `src/vba-parser.js` に対応するパースルールを追加（コード直打ちからも
   同じブロックを作れるようにする。文なら `SIMPLE_RULES` に正規表現ルールを、
   値を返す式なら `ExprParser.parseIdentChain` に分岐を足す）

この4箇所をそろえると、プレビュー・コード生成・コード直打ちの全てが動きます。
