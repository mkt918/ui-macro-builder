# VBAコード直打ち機能 — 実装タスク

要件: [vba-code-editing-requirements.md](vba-code-editing-requirements.md)

進捗は上から順に更新する。チェックは実装 **かつ** 検証が済んだ時点で入れる。

---

## フェーズ1: パーサー（コード → ブロック変換の核）

- [x] `src/vba-parser.js` を新規作成（Blockly非依存の純粋JS）
  - [x] 字句解析（数値・文字列・識別子・演算子）
  - [x] 式パーサー（四則演算・`Mod`・文字列連結`&`・比較演算子・優先順位）
  - [x] `Range(...)` / `Cells(...)` / `arr(...)` / `InputBox(...)` / `WorksheetFunction.Round` / `Date` / `i` の読み取り式
  - [x] 文パーサー: `For`（`loop_range` / `loop_for_step`）
  - [x] 文パーサー: `Do While` / `Do Until`（`loop_while` / `loop_do_until`）
  - [x] 文パーサー: `If ... Then ... [Else] End If`（`cond_if`）
  - [x] 文パーサー: `MsgBox`
  - [x] 文パーサー: セル代入・書式・配列・変数・シート操作（正規表現ルールテーブル）
  - [x] 変数名プレースホルダ（`{__var:name}`）と `resolveVariables()`（workspace連携用、実行時のみ）
  - [x] RGB値 ⇔ 色名の逆引き
  - [x] エラー時に行番号つき `VbaParseError` を投げる
- [x] Node単体テスト: `src/tasks/tasks.js` の6課題の `answer`（VBA文字列）を全てパースし、
      既存の `answerBlocks` と構造的に一致することを確認
      → t01/t06 は完全一致。t02/t04/t05（`loop_range` vs `loop_repeat`）、
      t03（`cond_compare`+`value_math` vs `cond_is_even`）は既知の制約により
      ブロック構造は異なるが、生成されるVBAコードは意味的に等価（下記テストで確認）
- [x] Node単体テスト: 生成 → パース → 再生成のラウンドトリップで意味的に同じVBAコードになることを確認
      → 6課題中5課題は空白正規化後に完全一致。t03のみ括弧位置の違い
      （`(i Mod 2) = 0` vs `(i Mod 2 = 0)`）があるが、VBAの演算子優先順位
      （Mod > 比較）により意味的に同一
- [x] Node単体テスト: 意図的な構文エラー（`Select Case` 未対応構文、`Next`漏れ、
      ひらがな変数名、未対応オブジェクト）で `VbaParseError` が正しい行番号と
      ともに投げられることを確認。`End If`/`Next i` 等の大小文字ゆらぎは許容されることも確認

## フェーズ2: エディタ導入（表示のみ、まだ同期なし）

- [ ] CodeMirror 5 を CDN 経由で `index.html` に追加（core + vbscriptモード）
- [ ] `#code-output` の `<pre><code>` を CodeMirror インスタンスに置き換え
- [ ] 既存の `highlightVBA()` 呼び出しを CodeMirror の `setValue()` に置き換え
- [ ] ダーク/ライトモードでの配色をCatppuccinパレットに合わせて調整
- [ ] ブロック → コードの既存同期が壊れていないことを確認

## フェーズ3: 自動インデント・入力アシスト

- [ ] Enterキー時の自動インデントロジック実装
  （`For`/`Do`/`If...Then` の次行で+1、`Next`/`Loop`/`End If`/`Else` で-1）
- [ ] スニペットテーブル実装（`for` / `while` / `until` / `if` / `ifelse` → Tab で展開）
- [ ] キーボード操作の動作確認（実ブラウザ）

## フェーズ4: コード → ブロック同期の結線

- [ ] 「コード編集済み」フラグの実装（CodeMirror change イベントで立てる。
      プログラム側からの `setValue` では立てない）
- [ ] ブロック変更時、コード編集済みフラグが立っていなければ自動反映（既存動作の維持）
- [ ] タブ切替時 / エディタ blur 時に `parseVBA()` を実行し、成功したら
      `Blockly.serialization.workspaces.load()` でワークスペースへ反映
- [ ] 失敗時のエラー表示UI（行番号・メッセージ。エディタ内 or 専用バナー）
- [ ] 変数の自動作成（`resolveVariables()` と `workspace.createVariable()` の結線）
- [ ] `suppressSave` 等、既存の保存抑制フラグとの整合性を確認（無限ループ・二重保存防止）

## フェーズ5: 仕上げ・検証

- [ ] 6課題すべてで「答えを見る」→ ブロック表示 → コードタブ表示 → 手動編集 →
      ブロックタブに戻る、の一連が壊れないことを確認
- [ ] フリーモードでの動作確認
- [ ] 共有リンク機能（ブロックXMLベース）とコードエディタの状態の整合性確認
- [ ] README / CLAUDE.md への追記（コード編集機能の説明）
- [ ] git commit（検証ログを添えて）

---

## 未決事項（実装中に判断が必要になったら都度確認する）

- スニペットのキー割り当て（Tabトリガーで確定。他候補との衝突がないか実装時に確認）
- エラー表示の見た目（バナー1本 vs エディタ行ハイライト）は実装時に決める
