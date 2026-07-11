"""validation.py 引擎測試（無 Qt 依賴，直接用 .venv python 執行）。

用法：  .venv/Scripts/python.exe tests/test_validation.py
"""
import json
import os
import sys
import tempfile

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from json_data_manager import JsonDataManager
from validation import (ValidationEngine, SafeExpr, ExprError,
                        new_rule, eval_op)


# ──────────────────────────────────────────────────────────────────────────────
# fixtures
# ──────────────────────────────────────────────────────────────────────────────

SKILLS = [
    {
        "SkillID": 1, "Name": "火球", "Type": "Active", "MaxLevel": 10,
        "Operation": [
            {"SkillComponent": "Damage", "EffectDurationTime": ""},
            {"SkillComponent": "ContinueBuff", "EffectDurationTime": 5},
        ],
    },
    {
        "SkillID": 2, "Name": "灼燒", "Type": "Active", "MaxLevel": 5,
        "Operation": [
            {"SkillComponent": "ContinueBuff", "EffectDurationTime": ""},
        ],
    },
    {
        "SkillID": 3, "Name": "被動強化", "Type": "Passive", "MaxLevel": 0,
        "Operation": [],
    },
]


def make_manager():
    tmp = tempfile.mkdtemp(prefix="valtest_")
    jpath = os.path.join(tmp, "SkillData.json")
    with open(jpath, "w", encoding="utf-8") as f:
        json.dump(SKILLS, f, ensure_ascii=False)
    mgr = JsonDataManager(config_path=os.path.join(tmp, "config.json"))
    mgr.load_json(jpath)
    cfg = mgr.config["SkillData"]
    cfg["columns"]["SkillID"]["type"] = "int"
    cfg["columns"]["MaxLevel"]["type"] = "int"
    cfg["sub_tables"]["Operation"]["columns"]["EffectDurationTime"] = {"type": "float"}
    return mgr


def engine_with(mgr, *rules):
    mgr.config["SkillData"]["validations"] = [dict(r) for r in rules]
    eng = ValidationEngine(mgr)
    eng.reload()
    return eng


def sub_idx(mgr, component):
    sdf = mgr.sub_tables["SkillData.Operation"]
    return [i for i in sdf.index if sdf.at[i, "SkillComponent"] == component]


# ──────────────────────────────────────────────────────────────────────────────
# tests
# ──────────────────────────────────────────────────────────────────────────────

def test_eval_op_basics():
    assert eval_op("ContinueBuff", "eq", "ContinueBuff")
    assert eval_op("1.0", "eq", "1")                 # numeric-lenient eq
    assert not eval_op("", "not_empty")
    assert eval_op(None, "empty")
    assert eval_op(0, "not_empty")                   # 0 不是空
    assert eval_op("5", "gt", "3") and not eval_op("abc", "gt", "3")
    assert eval_op("7", "between", "5", "10")
    assert eval_op("B", "in_list", "A, B, C") and not eval_op("D", "in_list", "A,B,C")
    assert eval_op("Skill_01", "regex", r"^Skill_\d+$")


def test_builder_sub_rule_continuebuff():
    """ContinueBuff 必填 EffectDurationTime — 你在需求裡舉的例子。"""
    mgr = make_manager()
    rule = new_rule(scope="Operation")
    rule.update(name="ContinueBuff必填持續時間", color="#AA3355",
                when={"logic": "and", "conds": [
                    {"field": "SkillComponent", "op": "eq", "value": "ContinueBuff"}]},
                then=[{"field": "EffectDurationTime", "op": "not_empty"}])
    eng = engine_with(mgr, rule)

    bad_rows = [i for i in sub_idx(mgr, "ContinueBuff")
                if eng.row_has_violation("SkillData.Operation", i)]
    ok_rows = [i for i in sub_idx(mgr, "Damage")
               if eng.row_has_violation("SkillData.Operation", i)]
    assert len(bad_rows) == 1 and not ok_rows          # 只有技能2那列違規
    bad = bad_rows[0]
    # 上色欄位 = then 的欄位；顏色 = 規則色
    assert eng.cell_color("SkillData.Operation", bad, "EffectDurationTime") == "#AA3355"
    assert eng.cell_color("SkillData.Operation", bad, "SkillComponent") is None

    # 增量重驗：補上值後違規消失
    mgr.update_cell("SkillData.Operation", bad, "EffectDurationTime", "3")
    eng.on_cell_edited("SkillData.Operation", bad, "EffectDurationTime")
    assert not eng.row_has_violation("SkillData.Operation", bad)
    # 再清掉又出現
    mgr.update_cell("SkillData.Operation", bad, "EffectDurationTime", "")
    eng.on_cell_edited("SkillData.Operation", bad, "EffectDurationTime")
    assert eng.row_has_violation("SkillData.Operation", bad)


def test_builder_master_required():
    """when 留空 = 每列都檢查 then（必填規則）。"""
    mgr = make_manager()
    mgr.update_cell("SkillData", mgr.tables["SkillData"].index[1], "Name", "")
    rule = new_rule()
    rule.update(name="Name必填", then=[{"field": "Name", "op": "not_empty"}])
    eng = engine_with(mgr, rule)
    rows = eng.sheet_violation_rows("SkillData")
    assert rows == {mgr.tables["SkillData"].index[1]}


def test_sub_rule_reads_master():
    """子表規則參照 master.欄位：Passive 技能的 Operation 不可有 Damage。"""
    mgr = make_manager()
    # 給技能3(Passive) 加一列 Damage operation
    sdf = mgr.sub_tables["SkillData.Operation"]
    new_i = (max(sdf.index) + 1) if len(sdf.index) else 0
    sdf.loc[new_i] = {"SkillID": "3", "SkillComponent": "Damage",
                      "EffectDurationTime": ""}
    rule = new_rule(scope="Operation")
    rule.update(name="被動不可有Damage",
                when={"logic": "and", "conds": [
                    {"field": "master.Type", "op": "eq", "value": "Passive"}]},
                then=[{"field": "SkillComponent", "op": "ne", "value": "Damage"}],
                mark=["SkillComponent"])
    eng = engine_with(mgr, rule)
    assert eng.row_has_violation("SkillData.Operation", new_i)
    # 改母表 Type → 增量重驗連動子表
    m3 = mgr.tables["SkillData"].index[2]
    mgr.update_cell("SkillData", m3, "Type", "Active")
    eng.on_cell_edited("SkillData", m3, "Type")
    assert not eng.row_has_violation("SkillData.Operation", new_i)


def test_master_agg_rule():
    """母表聚合條件：Active 技能至少要有 1 列 Operation。"""
    mgr = make_manager()
    rule = new_rule()
    rule.update(name="Active需有Operation", severity="warn",
                when={"logic": "and", "conds": [
                    {"field": "Type", "op": "eq", "value": "Active"}]},
                then=[{"agg": {"sub": "Operation", "field": "", "op": "eq",
                               "value": "", "count_op": "ge", "count": 1}}],
                mark=["SkillID"])
    eng = engine_with(mgr, rule)
    assert not eng.sheet_violation_rows("SkillData")      # 技能1,2都有列
    # 把技能3改成 Active → 違規（它的 Operation 是空的）
    m3 = mgr.tables["SkillData"].index[2]
    mgr.update_cell("SkillData", m3, "Type", "Active")
    eng.on_cell_edited("SkillData", m3, "Type")
    assert eng.sheet_violation_rows("SkillData") == {m3}
    assert not eng.has_errors()                           # warn 不算 error


def test_expr_rule_and_any_sub():
    mgr = make_manager()
    r1 = new_rule(scope="Operation")
    r1.update(name="expr版ContinueBuff", mode="expr",
              expr='SkillComponent != "ContinueBuff" or not empty(EffectDurationTime)',
              mark=["EffectDurationTime"])
    r2 = new_rule()
    r2.update(name="Passive不可有Operation", mode="expr",
              expr='Type != "Passive" or count_sub("Operation") == 0',
              mark=["Type"])
    eng = engine_with(mgr, r1, r2)
    assert len(eng.sheet_violation_rows("SkillData.Operation")) == 1
    assert not eng.sheet_violation_rows("SkillData")
    # master.X 在 expr 裡
    r3 = new_rule(scope="Operation")
    r3.update(name="expr讀master", mode="expr",
              expr='master.MaxLevel >= 1')
    eng = engine_with(mgr, r3)
    # 技能3 MaxLevel=0，但它沒有 Operation 列 → 無違規
    assert not eng.sheet_violation_rows("SkillData.Operation")


def test_expr_safety():
    for bad in ('__import__("os")', 'open("x")', '[x for x in (1,2)]',
                'master.__class__', '(lambda: 1)()', 'Name.__len__()',
                'exec("1")'):
        try:
            SafeExpr(bad)
            raise AssertionError(f"應拒絕: {bad}")
        except ExprError:
            pass
    # 合法的都要能編
    for ok in ('A == 1', 'empty(B)', 'not empty(A) and (B > 3 or C in "xyz")',
               'any_sub("Op", \'X == 1\')', 'master.T == "a"',
               'num(A) != None', 'min(A, B) >= 0', 'A in [1, 2, 3]'):
        SafeExpr(ok)


def test_severity_and_disabled():
    mgr = make_manager()
    r1 = new_rule(scope="Operation")
    r1.update(name="err", severity="error",
              when={"logic": "and", "conds": [
                  {"field": "SkillComponent", "op": "eq", "value": "ContinueBuff"}]},
              then=[{"field": "EffectDurationTime", "op": "not_empty"}])
    r2 = new_rule()
    r2.update(name="disabled", enabled=False,
              then=[{"field": "Name", "op": "empty"}])   # 若啟用會全表違規
    eng = engine_with(mgr, r1, r2)
    assert eng.has_errors()
    assert not eng.sheet_violation_rows("SkillData")      # r2 停用
    errs, warns = eng.count_by_severity()
    assert errs == 1 and warns == 0
    s = eng.summary()
    assert len(s) == 1 and s[0]["pk_val"] == "2" and s[0]["cols"] == ["EffectDurationTime"]


def test_test_rule_adhoc():
    mgr = make_manager()
    eng = engine_with(mgr)                                # 無正式規則
    rule = new_rule(scope="Operation")
    rule.update(mode="expr", expr='not empty(EffectDurationTime)')
    cnt, samples, err = eng.test_rule("SkillData", rule)
    assert err is None and cnt == 2                       # 兩列空白
    cnt, samples, err = eng.test_rule("SkillData", {"mode": "expr", "expr": "1 +"})
    assert err is not None
    assert not eng.violations                             # 測試不污染違規映射


def test_pk_edit_relinks():
    """改母表主鍵 → 整表重驗（子表 FK 連動）。"""
    mgr = make_manager()
    rule = new_rule(scope="Operation")
    rule.update(name="讀master", mode="expr", expr='master.Type == "Active"')
    eng = engine_with(mgr, rule)
    assert not eng.sheet_violation_rows("SkillData.Operation")
    # 把技能2的主鍵改掉 → 它的 Operation 列變孤兒 → master.Type 是 None → 違規
    m2 = mgr.tables["SkillData"].index[1]
    mgr.update_cell("SkillData", m2, "SkillID", "99")
    eng.on_cell_edited("SkillData", m2, "SkillID")
    assert len(eng.sheet_violation_rows("SkillData.Operation")) == 1


ALL = [v for k, v in sorted(globals().items()) if k.startswith("test_")]

if __name__ == "__main__":
    failed = 0
    for fn in ALL:
        try:
            fn()
            print(f"  PASS  {fn.__name__}")
        except Exception as e:
            failed += 1
            import traceback
            print(f"  FAIL  {fn.__name__}: {e}")
            traceback.print_exc()
    print(f"\n{len(ALL) - failed}/{len(ALL)} passed")
    sys.exit(1 if failed else 0)
