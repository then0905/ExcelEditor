"""驗證功能 UI 冒煙測試（offscreen Qt）。

用法：  QT_QPA_PLATFORM=offscreen .venv/Scripts/python.exe tests/test_validation_ui.py
"""
import json
import os
import sys
import tempfile

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from PySide6.QtWidgets import QApplication
from PySide6.QtCore import Qt

app = QApplication.instance() or QApplication([])

import main as M
from json_data_manager import JsonDataManager
from validation import new_rule

SKILLS = [
    {"SkillID": 1, "Name": "火球", "Type": "Active",
     "Operation": [{"SkillComponent": "Damage", "EffectDurationTime": ""},
                   {"SkillComponent": "ContinueBuff", "EffectDurationTime": 5}]},
    {"SkillID": 2, "Name": "灼燒", "Type": "Active",
     "Operation": [{"SkillComponent": "ContinueBuff", "EffectDurationTime": ""}]},
]


def make_manager_with_rule():
    tmp = tempfile.mkdtemp(prefix="valui_")
    jpath = os.path.join(tmp, "SkillData.json")
    with open(jpath, "w", encoding="utf-8") as f:
        json.dump(SKILLS, f, ensure_ascii=False)
    mgr = JsonDataManager(config_path=os.path.join(tmp, "config.json"))
    mgr.load_json(jpath)
    rule = new_rule(scope="Operation")
    rule.update(name="ContinueBuff必填持續時間", color="#AA3355",
                when={"logic": "and", "conds": [
                    {"field": "SkillComponent", "op": "eq", "value": "ContinueBuff"}]},
                then=[{"field": "EffectDurationTime", "op": "not_empty"}])
    mgr.config["SkillData"]["validations"] = [rule]
    mgr.validator.reload()
    return mgr


def test_subtable_model_colors():
    mgr = make_manager_with_rule()
    full = mgr.sub_tables["SkillData.Operation"]
    cols_cfg = mgr.config["SkillData"]["sub_tables"]["Operation"]["columns"]
    model = M.SubTableModel(full.copy(), cols_cfg, mgr, "SkillData.Operation")
    dur_c = list(full.columns).index("EffectDurationTime")
    comp_c = list(full.columns).index("SkillComponent")
    bad_r = [r for r in range(model.rowCount())
             if full.at[full.index[r], "SkillComponent"] == "ContinueBuff"
             and str(full.at[full.index[r], "EffectDurationTime"]).strip() == ""]
    assert len(bad_r) == 1
    r = bad_r[0]
    bg = model.data(model.index(r, dur_c), Qt.BackgroundRole)
    c = bg.color()
    assert (c.red(), c.green(), c.blue()) == (0xAA, 0x33, 0x55), c.name()
    tip = model.data(model.index(r, dur_c), Qt.ToolTipRole)
    assert tip and "ContinueBuff必填持續時間" in tip
    # 沒違規的格子照舊
    bg_ok = model.data(model.index(r, comp_c), Qt.BackgroundRole)
    assert bg_ok.color().name() != "#aa3355"
    # setData 修好 → 顏色消失
    model.setData(model.index(r, dur_c), "9")
    bg2 = model.data(model.index(r, dur_c), Qt.BackgroundRole)
    assert bg2.color().name() != "#aa3355"
    print("  PASS  test_subtable_model_colors")


def test_field_editor_styles():
    mgr = make_manager_with_rule()
    # 加一條母表規則：Name 必填
    r2 = new_rule()
    r2.update(name="Name必填", color="#2266DD",
              then=[{"field": "Name", "op": "not_empty"}])
    mgr.config["SkillData"]["validations"].append(r2)
    mgr.validator.reload()
    m1 = mgr.tables["SkillData"].index[0]
    mgr.update_cell("SkillData", m1, "Name", "")

    panel = M.FieldEditorWidget()
    panel.build_for(mgr.tables["SkillData"], mgr.config["SkillData"],
                    "SkillData", mgr)
    panel.load_row(mgr.tables["SkillData"].loc[m1], m1)
    w = panel._widgets["Name"]
    assert "rgba(34,102,221" in w.styleSheet().replace(" ", ""), w.styleSheet()
    assert "Name必填" in w.toolTip()
    # 修好 → refresh_validation 還原樣式
    mgr.update_cell("SkillData", m1, "Name", "火球")
    panel.refresh_validation()
    assert "rgba(34,102,221" not in w.styleSheet().replace(" ", "")
    print("  PASS  test_field_editor_styles")


def test_rules_dialog_roundtrip():
    mgr = make_manager_with_rule()
    dlg = M.ValidationRulesDialog(None, mgr, "SkillData")
    assert dlg._list.count() == 1
    assert dlg.f_name.text() == "ContinueBuff必填持續時間"
    assert dlg.f_scope.currentData() == "Operation"
    # when/then 條件列有載入
    assert len(dlg._cond_rows(dlg._when_lo)) == 1
    assert len(dlg._cond_rows(dlg._then_lo)) == 1
    # 改名 + roundtrip
    dlg.f_name.setText("改名了")
    dlg._save_form()
    assert dlg.rules[0]["name"] == "改名了"
    assert dlg.rules[0]["when"]["conds"][0]["value"] == "ContinueBuff"
    assert dlg.rules[0]["then"][0]["op"] == "not_empty"
    # 新增一條規則 → 進階模式 → 測試
    dlg._add_rule()
    assert dlg._list.count() == 2 and dlg._cur == 1
    dlg.f_mode_expr.setChecked(True)
    dlg.f_expr.setPlainText('not empty(Name)')
    dlg._test_rule()
    assert "全部通過" in dlg._test_lbl.text()
    dlg.f_expr.setPlainText('1 +')
    dlg._test_rule()
    assert "✗" in dlg._test_lbl.text()
    # 刪除
    dlg._del_rule()
    assert dlg._list.count() == 1
    # 套用回寫
    dlg._apply()
    assert dlg.rules[0]["name"] == "改名了"
    print("  PASS  test_rules_dialog_roundtrip")


def test_item_delegate_role():
    mgr = make_manager_with_rule()
    m2 = mgr.tables["SkillData"].index[1]     # 技能2 有違規子列
    sev = mgr.validator.record_violation_severity("SkillData", m2)
    assert sev == "error"
    m1 = mgr.tables["SkillData"].index[0]
    assert mgr.validator.record_violation_severity("SkillData", m1) is None
    print("  PASS  test_item_delegate_role")


if __name__ == "__main__":
    failed = 0
    for fn in (test_subtable_model_colors, test_field_editor_styles,
               test_rules_dialog_roundtrip, test_item_delegate_role):
        try:
            fn()
        except Exception as e:
            failed += 1
            import traceback
            print(f"  FAIL  {fn.__name__}: {e}")
            traceback.print_exc()
    print(f"\n{4 - failed}/4 passed")
    sys.exit(1 if failed else 0)
