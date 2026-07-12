"""field_binding.py 引擎測試（無 Qt）。

用法：  .venv/Scripts/python.exe tests/test_field_binding.py
"""
import json
import os
import sys
import tempfile

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from json_data_manager import JsonDataManager
from validation import new_rule

SKILLS = [
    {"SkillID": "S1", "Name": "火球", "Type": "Active",
     "Operation": [
         {"SkillComponent": "Damage", "EffectValue": 10, "TargetCount": 1,
          "InfluenceStatus": "", "EffectDurationTime": ""},
         {"SkillComponent": "PassiveBuff", "EffectValue": "", "TargetCount": "",
          "InfluenceStatus": "ATK", "EffectDurationTime": 5},
     ]},
    {"SkillID": "S2", "Name": "新技", "Type": "Active",
     "Operation": [
         {"SkillComponent": "NewThing", "EffectValue": 1, "TargetCount": 1,
          "InfluenceStatus": "", "EffectDurationTime": ""},
     ]},
]

BINDING = {
    "Operation": {
        "enabled": True,
        "driver": "SkillComponent",
        "groups": {
            "Damage":      ["EffectValue", "TargetCount"],
            "PassiveBuff": ["InfluenceStatus", "EffectDurationTime"],
        },
    }
}


def make_manager(binding=True):
    tmp = tempfile.mkdtemp(prefix="fb_")
    jpath = os.path.join(tmp, "SkillData.json")
    with open(jpath, "w", encoding="utf-8") as f:
        json.dump(SKILLS, f, ensure_ascii=False)
    mgr = JsonDataManager(config_path=os.path.join(tmp, "config.json"))
    mgr.load_json(jpath)
    if binding:
        mgr.config["SkillData"]["field_bindings"] = json.loads(json.dumps(BINDING))
    return mgr


def rows_by_component(mgr, comp):
    sdf = mgr.sub_tables["SkillData.Operation"]
    return [i for i in sdf.index if sdf.at[i, "SkillComponent"] == comp]


def test_relevant_fields():
    mgr = make_manager()
    dmg = rows_by_component(mgr, "Damage")[0]
    rel = mgr.binding.relevant_fields("SkillData", "Operation", dmg)
    # driver + FK + Damage group；PassiveBuff 專屬欄位不在
    assert "SkillComponent" in rel and "SkillID" in rel
    assert "EffectValue" in rel and "TargetCount" in rel
    assert "InfluenceStatus" not in rel and "EffectDurationTime" not in rel

    buf = rows_by_component(mgr, "PassiveBuff")[0]
    rel2 = mgr.binding.relevant_fields("SkillData", "Operation", buf)
    assert "InfluenceStatus" in rel2 and "EffectValue" not in rel2
    print("  PASS  test_relevant_fields")


def test_unconfigured_value_and_disabled():
    mgr = make_manager()
    new = rows_by_component(mgr, "NewThing")[0]
    # groups 沒設定的值 → None = 全相關
    assert mgr.binding.relevant_fields("SkillData", "Operation", new) is None
    assert mgr.binding.is_relevant("SkillData.Operation", new, "EffectValue")
    # 停用綁定 → 全相關
    mgr.config["SkillData"]["field_bindings"]["Operation"]["enabled"] = False
    dmg = rows_by_component(mgr, "Damage")[0]
    assert mgr.binding.relevant_fields("SkillData", "Operation", dmg) is None
    # 沒綁定的表 → 全相關
    assert mgr.binding.relevant_fields("SkillData", "", 0) is None
    print("  PASS  test_unconfigured_value_and_disabled")


def test_visible_columns_union():
    mgr = make_manager()
    idxs = rows_by_component(mgr, "Damage") + rows_by_component(mgr, "PassiveBuff")
    vis = mgr.binding.visible_columns("SkillData", "Operation", idxs)
    # 兩組聯集 → 全部綁定欄位都在
    for c in ("EffectValue", "TargetCount", "InfluenceStatus", "EffectDurationTime"):
        assert c in vis
    # 只有 Damage 列 → PassiveBuff 欄位整欄可藏
    vis2 = mgr.binding.visible_columns("SkillData", "Operation",
                                       rows_by_component(mgr, "Damage"))
    assert "InfluenceStatus" not in vis2 and "EffectValue" in vis2
    # 含未設定值的列 → None = 全顯示
    vis3 = mgr.binding.visible_columns("SkillData", "Operation",
                                       rows_by_component(mgr, "NewThing"))
    assert vis3 is None
    print("  PASS  test_visible_columns_union")


def test_shared_field_always_relevant():
    mgr = make_manager()
    # 沒被任何 group 提到的欄位（如子表沒列出的欄）→ 永遠相關
    sdf = mgr.sub_tables["SkillData.Operation"]
    dmg = rows_by_component(mgr, "Damage")[0]
    rel = mgr.binding.relevant_fields("SkillData", "Operation", dmg)
    mentioned = {"EffectValue", "TargetCount", "InfluenceStatus", "EffectDurationTime"}
    for c in sdf.columns:
        if c not in mentioned:
            assert c in rel, c
    print("  PASS  test_shared_field_always_relevant")


def test_validation_skips_irrelevant_marks():
    mgr = make_manager()
    # 規則：每列 EffectDurationTime 必填（對 Damage 列本來會違規）
    r = new_rule(scope="Operation")
    r.update(name="必填持續", then=[{"field": "EffectDurationTime", "op": "not_empty"}])
    mgr.config["SkillData"]["validations"] = [r]
    mgr.validator.reload()
    dmg = rows_by_component(mgr, "Damage")[0]
    new = rows_by_component(mgr, "NewThing")[0]
    # Damage 列：EffectDurationTime 不相關 → 不標violations
    assert not mgr.validator.row_has_violation("SkillData.Operation", dmg)
    # NewThing 列（未綁定→全相關）：照常違規
    assert mgr.validator.row_has_violation("SkillData.Operation", new)
    # 把 Damage 列 driver 改成 PassiveBuff → EffectDurationTime 變相關 → 增量重驗標出違規
    mgr.update_cell("SkillData.Operation", dmg, "SkillComponent", "PassiveBuff")
    assert mgr.validator.row_has_violation("SkillData.Operation", dmg)
    print("  PASS  test_validation_skips_irrelevant_marks")


if __name__ == "__main__":
    failed = 0
    for fn in (test_relevant_fields, test_unconfigured_value_and_disabled,
               test_visible_columns_union, test_shared_field_always_relevant,
               test_validation_skips_irrelevant_marks):
        try:
            fn()
        except Exception as e:
            failed += 1
            import traceback
            print(f"  FAIL  {fn.__name__}: {e}")
            traceback.print_exc()
    print(f"\n{5 - failed}/5 passed")
    sys.exit(1 if failed else 0)
