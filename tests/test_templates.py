"""範本列功能測試（offscreen Qt，實際建構 TableEditor）。

用法：  QT_QPA_PLATFORM=offscreen .venv/Scripts/python.exe tests/test_templates.py
"""
import json
import os
import sys
import tempfile

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from PySide6.QtWidgets import QApplication

app = QApplication.instance() or QApplication([])

import main as M
from json_data_manager import JsonDataManager

SKILLS = [
    {"SkillID": "S1", "Name": "火球", "Type": "Active",
     "Operation": [{"SkillComponent": "Damage", "EffectValue": 10},
                   {"SkillComponent": "ContinueBuff", "EffectValue": 3}]},
    {"SkillID": "S2", "Name": "灼燒", "Type": "Active",
     "Operation": [{"SkillComponent": "ContinueBuff", "EffectValue": 1}]},
]


def make_editor():
    tmp = tempfile.mkdtemp(prefix="tpl_")
    jpath = os.path.join(tmp, "SkillData.json")
    with open(jpath, "w", encoding="utf-8") as f:
        json.dump(SKILLS, f, ensure_ascii=False)
    mgr = JsonDataManager(config_path=os.path.join(tmp, "config.json"))
    mgr.load_json(jpath)
    ed = M.TableEditor("SkillData", mgr)
    return ed, mgr


def test_set_and_list_template():
    ed, mgr = make_editor()
    # 直接寫 config（set_template 走 QInputDialog，邏輯等價）
    ed._templates().append({"name": "標準攻擊技", "pk": "S1"})
    mgr.save_config()
    assert ed._template_pks() == {"S1"}
    # config 有存
    cfg = mgr.config["SkillData"]
    assert cfg["templates"][0]["name"] == "標準攻擊技"
    # 卡片顯示 ⭐
    ed.current_cls_val = M._ALL_GROUPS
    ed._load_item_list()
    texts = [ed._card_list.item(i).data(M.ItemCardDelegate.R_PK)
             for i in range(ed._card_list.count())]
    assert any(t.startswith("⭐") for t in texts), texts
    print("  PASS  test_set_and_list_template")


def test_create_from_template():
    ed, mgr = make_editor()
    tpl = {"name": "標準攻擊技", "pk": "S1"}
    ed._templates().append(tpl)
    ed.current_cls_val = M._ALL_GROUPS
    ed._load_cls_list(); ed._load_item_list()

    ed._create_from_template(tpl, "S9")
    df = mgr.tables["SkillData"]
    assert "S9" in df["SkillID"].astype(str).values
    new_row = df[df["SkillID"].astype(str) == "S9"].iloc[0]
    assert new_row["Name"] == "火球" and new_row["Type"] == "Active"   # 母列複製
    # 插在 S1 同分類群後面
    idx_s1 = df[df["SkillID"] == "S1"].index[0]
    idx_s9 = df[df["SkillID"] == "S9"].index[0]
    assert idx_s9 > idx_s1

    # 子表列深拷貝＋FK 換新
    sdf = mgr.sub_tables["SkillData.Operation"]
    s9_rows = sdf[sdf["SkillID"].astype(str) == "S9"]
    assert len(s9_rows) == 2                                            # S1 有 2 列
    assert set(s9_rows["SkillComponent"]) == {"Damage", "ContinueBuff"}
    # 原範本列不受影響
    assert len(sdf[sdf["SkillID"].astype(str) == "S1"]) == 2
    assert mgr.dirty
    # 新項目被選中
    assert str(df.at[ed.current_master_idx, "SkillID"]) == "S9"
    print("  PASS  test_create_from_template")


def test_remove_template_and_dangling():
    ed, mgr = make_editor()
    ed._templates().append({"name": "t", "pk": "S1"})
    ed.remove_template("S1")
    assert ed._templates() == []
    # dangling：pk 不存在時 _create_from_template 不做事
    tpl = {"name": "ghost", "pk": "NOPE"}
    n = len(mgr.tables["SkillData"])
    ed._create_from_template(tpl, "S99")
    assert len(mgr.tables["SkillData"]) == n
    print("  PASS  test_remove_template_and_dangling")


if __name__ == "__main__":
    failed = 0
    for fn in (test_set_and_list_template, test_create_from_template,
               test_remove_template_and_dangling):
        try:
            fn()
        except Exception as e:
            failed += 1
            import traceback
            print(f"  FAIL  {fn.__name__}: {e}")
            traceback.print_exc()
    print(f"\n{3 - failed}/3 passed")
    sys.exit(1 if failed else 0)
