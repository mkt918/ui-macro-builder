# UI Macro Builder — プロジェクトメモ

授業用の Excel VBA ビジュアルプログラミング学習ツール。
詳細な計画は [PLAN.md](PLAN.md) を参照。

## 構成（ビルド不要・Blockly CDN）

- `index.html` — エントリーポイント
- `src/blockly-setup.js` — カスタムブロック定義 + ツールボックス
- `src/vba-generator.js` — ブロック → VBA コード生成
- `src/excel-preview.js` — ブロックを解釈実行して仮想 Excel をアニメーション
- `src/tasks/tasks.js` — 課題データ（`TASKS` 配列）
- `src/app.js` — 全体の結線

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

## ブロックを追加するとき

1. `src/blockly-setup.js` にブロック定義とツールボックス登録
2. `src/vba-generator.js` に VBA 生成関数（`vbaGenerator.forBlock["型名"]`）
3. `src/excel-preview.js` の `Interpreter.execBlock` / `evalValue` に実行ロジック

この3箇所をそろえると、プレビューとコード生成の両方が動きます。
