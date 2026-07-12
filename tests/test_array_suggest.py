"""陣列欄位建議值快速選填測試（offscreen Qt）。

用法：  QT_QPA_PLATFORM=offscreen .venv/Scripts/python.exe tests/test_array_suggest.py
"""
import json
import os
import sys
import tempfile

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from PySide6.QtWidgets import QApplication, QPushButton

app = QApplication.instance() or QApplication([])

import main as M
from json_data_manager import JsonDataManager

SKILLS = [
    {"SkillID": "S1", "Name": "火球", "Type": "Active", "Tags": ["fire", "aoe"],
     "Operation": [
         {"SkillComponent": "Damage", "InfluenceStatus": ["ATK", "DEF"]},
         {"SkillComponent": "Damage", "InfluenceStatus": ["ATK", "SPD"]},
     ]},
    {"SkillID": "S2", "Name": "灼燒", "Type": "Active", "Tags": ["burn"],
     "Operation": []},
    {"SkillID": "S3", "Name": "護盾", "Type": "Passive", "Tags": ["shield"],
     "Operation": [
         {"SkillComponent": "PassiveBuff", "InfluenceStatus": ["HP"]},
     ]},
]


def make_manager():
    tmp = tempfile.mkdtemp(prefix="arrsug_")
    jpath = os.path.join(tmp, "SkillData.json")
    with open(jpath, "w", encoding="utf-8") as f:
        json.dump(SKILLS, f, ensure_ascii=False)
    mgr = JsonDataManager(config_path=os.path.join(tmp, "config.json"))
    mgr.load_json(jpath)
    sub_cols = mgr.config["SkillData"]["sub_tables"]["Operation"]["columns"]
    sub_cols["InfluenceStatus"] = {"type": "array", "suggest_from": "SkillComponent"}
    mgr.config["SkillData"]["columns"]["Tags"] = {"type": "array",
                                                  "suggest_from": "Type"}
    return mgr


def test_get_array_suggestions():
    mgr = make_manager()
    sdf = mgr.sub_tables["SkillData.Operation"]
    # 依 context 過濾＋拆 token 去重排序
    out = M._get_array_suggestions(sdf, "InfluenceStatus", "SkillComponent", "Damage")
    assert out == ["ATK", "DEF", "SPD"], out
    out2 = M._get_array_suggestions(sdf, "InfluenceStatus", "SkillComponent", "PassiveBuff")
    assert out2 == ["HP"], out2
    # 沒 context → 全表 token
    out3 = M._get_array_suggestions(sdf, "InfluenceStatus", "", "")
    assert out3 == ["ATK", "DEF", "HP", "SPD"], out3
    print("  PASS  test_get_array_suggestions")


def test_dialog_suggestion_block():
    dlg = M.ArrayEditDialog("ATK", suggestions=["DEF", "SPD"], source_note="來源：測試")
    # 建議按鈕存在
    btns = [b for b in dlg.findChildren(QPushButton) if b.text().startswith("＋ ")]
    assert {b.text() for b in btns} == {"＋ DEF", "＋ SPD"}
    # 點擊建議 → 加入 chips
    next(b for b in btns if b.text() == "＋ DEF").click()
    assert dlg.value() == "ATK, DEF"
    # 沒建議 → 不出現區塊
    dlg2 = M.ArrayEditDialog("ATK")
    btns2 = [b for b in dlg2.findChildren(QPushButton) if b.text().startswith("＋ ")]
    assert not btns2
    print("  PASS  test_dialog_suggestion_block")


def test_delegate_wiring():
    mgr = make_manager()
    cols_cfg = mgr.config["SkillData"]["sub_tables"]["Operation"]["columns"]
    panel = M.SubTablePanel("SkillData.Operation", cols_cfg, mgr)
    sdf = mgr.sub_tables["SkillData.Operation"]
    panel.reload(sdf, cols_cfg)
    c = list(sdf.columns).index("InfluenceStatus")
    d = panel._view.itemDelegateForColumn(c)
    assert isinstance(d, M.ArrayDelegate)
    assert d._this_col == "InfluenceStatus" and d._context_col == "SkillComponent"
    assert d._df_provider() is mgr.sub_tables["SkillData.Operation"]
    print("  PASS  test_delegate_wiring")


def test_master_field_editor_suggestions():
    mgr = make_manager()
    df = mgr.tables["SkillData"]
    panel = M.FieldEditorWidget()
    panel.build_for(df, mgr.config["SkillData"], "SkillData", mgr)
    assert "Tags" in panel._array_sug
    # S1（Active）→ 建議＝所有 Active 列的 Tags token
    m1 = df.index[0]
    panel.load_row(df.loc[m1], m1)
    panel.refresh_array_suggestions()
    sug = panel._array_sug["Tags"]
    texts = set()
    for i in range(sug["flow"].count()):
        texts.add(sug["flow"].itemAt(i).widget().text())
    assert texts == {"＋ aoe", "＋ burn", "＋ fire"}, texts
    assert "Type＝Active" in sug["cap"].text()
    # 點建議 → 加入 chips
    for i in range(sug["flow"].count()):
        b = sug["flow"].itemAt(i).widget()
        if b.text() == "＋ burn":
            b.click()
    assert "burn" in panel._widgets["Tags"].value()
    # 換到 Passive 列 → 建議跟著換
    m3 = df.index[2]
    panel.load_row(df.loc[m3], m3)
    panel.refresh_array_suggestions()
    texts = {sug["flow"].itemAt(i).widget().text()
             for i in range(sug["flow"].count())}
    assert texts == {"＋ shield"}, texts
    print("  PASS  test_master_field_editor_suggestions")


if __name__ == "__main__":
    tests = (test_get_array_suggestions, test_dialog_suggestion_block,
             test_delegate_wiring, test_master_field_editor_suggestions)
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
