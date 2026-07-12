"""欄位綁定 UI 測試（offscreen Qt）。

用法：  QT_QPA_PLATFORM=offscreen .venv/Scripts/python.exe tests/test_field_binding_ui.py
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

SKILLS = [
    {"SkillID": "S1", "Name": "火球", "Type": "Active",
     "Operation": [
         {"SkillComponent": "Damage", "EffectValue": 10,
          "InfluenceStatus": "", "EffectDurationTime": ""},
         {"SkillComponent": "Heal", "EffectValue": 5,
          "InfluenceStatus": "", "EffectDurationTime": ""},
     ]},
    {"SkillID": "S2", "Name": "護盾", "Type": "Passive",
     "Operation": [
         {"SkillComponent": "PassiveBuff", "EffectValue": "",
          "InfluenceStatus": "DEF", "EffectDurationTime": 9},
     ]},
]

SUB_BINDING = {
    "enabled": True, "driver": "SkillComponent",
    "groups": {"Damage": ["EffectValue"],
               "PassiveBuff": ["InfluenceStatus", "EffectDurationTime"]},
}
MASTER_BINDING = {
    "enabled": True, "driver": "Type",
    "groups": {"Active": ["Name"], "Passive": []},
}


def make_manager(sub=True, master=False):
    tmp = tempfile.mkdtemp(prefix="fbui_")
    jpath = os.path.join(tmp, "SkillData.json")
    with open(jpath, "w", encoding="utf-8") as f:
        json.dump(SKILLS, f, ensure_ascii=False)
    mgr = JsonDataManager(config_path=os.path.join(tmp, "config.json"))
    mgr.load_json(jpath)
    fb = {}
    if sub:
        fb["Operation"] = json.loads(json.dumps(SUB_BINDING))
    if master:
        fb[""] = json.loads(json.dumps(MASTER_BINDING))
    mgr.config["SkillData"]["field_bindings"] = fb
    return mgr


def test_model_lock_and_grey():
    mgr = make_manager()
    full = mgr.sub_tables["SkillData.Operation"]
    cols_cfg = mgr.config["SkillData"]["sub_tables"]["Operation"]["columns"]
    model = M.SubTableModel(full.copy(), cols_cfg, mgr, "SkillData.Operation")
    cols = list(full.columns)
    dmg_r = 0 if full.at[full.index[0], "SkillComponent"] == "Damage" else 1
    inf_c = cols.index("InfluenceStatus")
    eff_c = cols.index("EffectValue")
    # Damage 列：InfluenceStatus 不相關 → 不可編輯、灰底、tooltip 註明
    idx = model.index(dmg_r, inf_c)
    assert not (model.flags(idx) & Qt.ItemIsEditable)
    assert model.data(idx, Qt.BackgroundRole).color().name() == M._C["code"].lower()
    assert "無關" in model.data(idx, Qt.ToolTipRole)
    # 相關格照常可編輯
    idx2 = model.index(dmg_r, eff_c)
    assert model.flags(idx2) & Qt.ItemIsEditable
    print("  PASS  test_model_lock_and_grey")


def test_panel_column_hiding():
    mgr = make_manager()
    full = mgr.sub_tables["SkillData.Operation"]
    cols_cfg = mgr.config["SkillData"]["sub_tables"]["Operation"]["columns"]
    panel = M.SubTablePanel("SkillData.Operation", cols_cfg, mgr)
    # 只載入 Damage 那筆（S1）的列 → PassiveBuff 專屬欄整欄隱藏
    dmg_rows = full[full["SkillComponent"] == "Damage"]
    panel.reload(dmg_rows, cols_cfg)
    cols = list(dmg_rows.columns)
    assert panel._view.isColumnHidden(cols.index("InfluenceStatus"))
    assert panel._view.isColumnHidden(cols.index("EffectDurationTime"))
    assert not panel._view.isColumnHidden(cols.index("EffectValue"))
    assert not panel._view.isColumnHidden(cols.index("SkillComponent"))
    # 兩種列都在 → 聯集全顯示
    panel.reload(full, cols_cfg)
    assert not panel._view.isColumnHidden(cols.index("InfluenceStatus"))
    print("  PASS  test_panel_column_hiding")


def test_field_editor_hides_blocks():
    mgr = make_manager(sub=False, master=True)
    df = mgr.tables["SkillData"]
    panel = M.FieldEditorWidget()
    panel.build_for(df, mgr.config["SkillData"], "SkillData", mgr)
    m2 = df.index[1]                      # S2: Type=Passive → Name 不相關
    panel.load_row(df.loc[m2], m2)
    panel.refresh_binding()
    grp, sep = panel._grp_widgets["Name"]
    assert grp.isHidden()
    # Type（driver）與 SkillID（PK）永遠顯示
    assert not panel._grp_widgets["Type"][0].isHidden()
    assert not panel._grp_widgets["SkillID"][0].isHidden()
    # 切到 Active 列 → Name 顯示
    m1 = df.index[0]
    panel.load_row(df.loc[m1], m1)
    panel.refresh_binding()
    assert not panel._grp_widgets["Name"][0].isHidden()
    print("  PASS  test_field_editor_hides_blocks")


def _card(tab, value):
    return next(e for e in tab._cards
                if e["value_cb"].currentText() == value)


def test_binding_tab_roundtrip():
    mgr = make_manager()
    dlg = M.ValidationRulesDialog(None, mgr, "SkillData")
    tab = dlg._binding_tab
    # 切到 Operation scope → 只有 config 設定過的值有卡片（不自動鋪滿）
    i = tab.f_scope.findData("Operation")
    tab.f_scope.setCurrentIndex(i)
    assert tab.f_driver.currentData() == "SkillComponent"
    vals = {e["value_cb"].currentText() for e in tab._cards}
    assert vals == {"Damage", "PassiveBuff"}          # Heal 在資料裡但沒設定→無卡
    assert _card(tab, "Damage")["boxes"]["EffectValue"].isChecked()
    assert not _card(tab, "PassiveBuff")["boxes"]["EffectValue"].isChecked()
    # 值下拉候選：現有資料的值優先列出
    cands = tab._value_candidates("SkillComponent")
    assert set(cands) >= {"Damage", "Heal", "PassiveBuff"}
    # 改勾選：把 EffectValue 也綁到 PassiveBuff
    _card(tab, "PassiveBuff")["boxes"]["EffectValue"].setChecked(True)
    out = dlg.bindings()
    assert "EffectValue" in out["Operation"]["groups"]["PassiveBuff"]
    assert out["Operation"]["driver"] == "SkillComponent"
    # 母表沒綁定 → 不出現
    assert "" not in out
    print("  PASS  test_binding_tab_roundtrip")


def test_binding_tab_add_remove_card_and_unbind():
    mgr = make_manager()
    dlg = M.ValidationRulesDialog(None, mgr, "SkillData")
    tab = dlg._binding_tab
    tab.f_scope.setCurrentIndex(tab.f_scope.findData("Operation"))
    # ＋新增卡片 → 空卡片，沒選值前不寫入
    tab._add_card("", set())
    assert len(tab._cards) == 3
    out = dlg.bindings()
    assert set(out["Operation"]["groups"]) == {"Damage", "PassiveBuff"}
    # 選值＋打勾 → 寫入
    new = tab._cards[-1]
    new["value_cb"].setCurrentText("Heal")
    new["boxes"]["EffectValue"].setChecked(True)
    out = dlg.bindings()
    assert out["Operation"]["groups"]["Heal"] == ["EffectValue"]
    # 移除卡片（模擬 🗑）→ 該值恢復不隱藏
    tab._cards.remove(new)
    out = dlg.bindings()
    assert "Heal" not in out["Operation"]["groups"]
    # 驅動欄位改成（不綁定）→ 該表綁定移除
    tab.f_driver.setCurrentIndex(0)
    out = dlg.bindings()
    assert "Operation" not in out
    print("  PASS  test_binding_tab_add_remove_card_and_unbind")


if __name__ == "__main__":
    tests = (test_model_lock_and_grey, test_panel_column_hiding,
             test_field_editor_hides_blocks, test_binding_tab_roundtrip,
             test_binding_tab_add_remove_card_and_unbind)
    failed = 0
    for fn in tests:
        try:
            fn()
        except Exception as e:
            failed += 1
            import traceback
            print(f"  FAIL  {fn.__name__}: {e}")
            traceback.print_exc()
    print(f"\n{len(tests) - failed}/{len(tests)} passed")
    sys.exit(1 if failed else 0)
