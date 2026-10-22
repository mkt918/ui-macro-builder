// 教科書の課題（src/tasks/textbook-tasks.js）のテスト。実行方法: node --test tests/*.test.js
const test = require("node:test");
const assert = require("node:assert/strict");
const path = require("node:path");
const { TASKS, TEXTBOOK_TASKS, ORIGINAL_TASKS, runTaskScenario } = require(path.join(__dirname, "../src/tasks/textbook-tasks.js"));
const P = require(path.join(__dirname, "../src/vba-parser.js"));

test("課題の id は重複せず、教科書の課題は必要な欄をそろえている", () => {
  const ids = TASKS.map((t) => t.id);
  assert.equal(new Set(ids).size, ids.length);
  for (const t of TEXTBOOK_TASKS) {
    assert.ok(t.kind === "example" || t.kind === "practice", t.id + " kind");
    assert.ok(t.code && t.page && t.title && t.goal, t.id + " 基本の欄");
    for (const k of ["input", "output", "inputExample", "outputExample"]) assert.ok(t.spec[k], `${t.id} spec.${k}`);
    assert.ok(t.spec.conditions.length >= 1, t.id + " 処理条件");
    assert.equal(t.hints.length, 3, t.id + " ヒントは3段階");
    assert.ok(t.form && t.form.controls.length, t.id + " フォーム");
    assert.ok(t.scenario && t.scenario.length, t.id + " scenario");
  }
});

test("オリジナル課題（割り勘アプリ）は必要な欄をそろえ、解説つき→挑戦の順に並ぶ", () => {
  assert.deepEqual(ORIGINAL_TASKS.map((t) => t.kind), ["guide", "guide", "challenge", "challenge"]);
  for (const t of ORIGINAL_TASKS) {
    assert.ok(t.series && t.step && t.title && t.goal, t.id + " 基本の欄");
    for (const k of ["input", "output", "inputExample", "outputExample"]) assert.ok(t.spec[k], `${t.id} spec.${k}`);
    assert.ok(t.spec.conditions.length >= 1, t.id + " 処理条件");
    assert.equal(t.hints.length, 3, t.id + " ヒントは3段階");
    assert.ok(t.lesson && t.lesson.length, t.id + " 解説・考え方");
    assert.ok(t.form && t.form.controls.length, t.id + " フォーム");
    assert.ok(t.scenario && t.scenario.length, t.id + " scenario");
  }
  // 教科書の課題のあと、セルを使った練習の前に入っている
  const ids = TASKS.map((t) => t.id);
  assert.ok(ids.indexOf("warikan1") > ids.indexOf("jis4-4-01"));
  assert.ok(ids.indexOf("warikan4") < ids.indexOf("t01"));
});

test("割り勘アプリ：「/」で書くとあまりが出ず、不合格になる（¥ と Mod が必要）", () => {
  const t = ORIGINAL_TASKS.find((x) => x.id === "warikan3");
  const wrong = t.answer.replace("Gokei ¥ Ninzu", "Gokei / Ninzu");
  const r = runTaskScenario(t, wrong);
  assert.equal(r.passed, false);
});

for (const t of TEXTBOOK_TASKS.concat(ORIGINAL_TASKS)) {
  test(`${t.title}：模範解答は、自分の課題に合格する`, () => {
    const r = runTaskScenario(t, t.answer);
    assert.deepEqual(r.failures, []);
    assert.equal(r.passed, true);
  });

  test(`${t.title}：模範解答は、ブロックに変換できるか、コード専用と判定される（エラーで落ちない）`, () => {
    try {
      const b = P.parseVBA(t.answer);
      assert.ok(b.roots.length >= 1);
    } catch (e) {
      assert.equal(e.name === "VbaParseError" || /ブロック|対応/.test(e.message), true, e.message);
    }
  });
}

test("まちがったプログラムは不合格になり、やさしい理由が出る", () => {
  const t = TEXTBOOK_TASKS.find((x) => x.id === "rei4-3-01");
  // 50 をこえたときの Exit Sub を忘れた
  const wrong = t.answer.replace("        Exit Sub\n", "");
  const r = runTaskScenario(t, wrong);
  assert.equal(r.passed, false);
  assert.match(r.failures.join("\n"), /4回目/);
  // 何も書いていない
  const empty = runTaskScenario(t, "Option Explicit\nPrivate Sub CommandButton1_Click()\nEnd Sub");
  assert.equal(empty.passed, false);
  assert.ok(empty.failures.length >= 3);
  // 文法まちがい
  const syntax = runTaskScenario(t, "Private Sub CommandButton1_Click()\n If x Then\nEnd Sub");
  assert.equal(syntax.passed, false);
  assert.match(syntax.failures[0], /End If/);
});

test("ボタンの命令（_Click）が無いと、フォームは開くが何も起きないので不合格", () => {
  const t = TEXTBOOK_TASKS.find((x) => x.id === "rei4-2-03");
  const r = runTaskScenario(t, "Option Explicit\nSub Hello()\nEnd Sub");
  assert.equal(r.passed, false);
});
