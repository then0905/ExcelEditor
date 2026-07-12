"""config 路徑沿用測試——模擬換電腦/搬資料夾後開同名檔（無 Qt）。

用法：  .venv/Scripts/python.exe tests/test_config_adopt.py
"""
import json
import os
import shutil
import sys
import tempfile

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from json_data_manager import JsonDataManager

SKILLS = [{"SkillID": "S1", "Name": "火球", "MaxLevel": 10,
           "Operation": [{"SkillComponent": "Damage"}]}]


def setup_machine_a():
    """機器A：載入並客製配置後存檔，回傳 (config路徑, 資料檔路徑)。"""
    tmp = tempfile.mkdtemp(prefix="cfgadopt_")
    dir_a = os.path.join(tmp, "machineA"); os.makedirs(dir_a)
    jpath = os.path.join(dir_a, "SkillData.json")
    with open(jpath, "w", encoding="utf-8") as f:
        json.dump(SKILLS, f, ensure_ascii=False)
    cfg_path = os.path.join(tmp, "config.json")
    mgr = JsonDataManager(config_path=cfg_path)
    mgr.load_json(jpath)
    mgr.config["SkillData"]["columns"]["MaxLevel"] = {"type": "int", "note": "上限"}
    mgr.config["SkillData"]["validations"] = [{"name": "測試規則"}]
    mgr.save_config()
    return tmp, cfg_path, jpath


def test_adopt_on_new_path():
    tmp, cfg_path, jpath_a = setup_machine_a()
    # 「機器B」：同一份 config，但資料檔在不同路徑
    dir_b = os.path.join(tmp, "machineB"); os.makedirs(dir_b)
    jpath_b = os.path.join(dir_b, "SkillData.json")
    shutil.copy(jpath_a, jpath_b)

    mgr = JsonDataManager(config_path=cfg_path)
    mgr.load_json(jpath_b)
    # 配置沿用成功：欄位型別/note/驗證規則都在
    assert mgr.adopted_config_from == os.path.normpath(jpath_a)
    assert mgr.config["SkillData"]["columns"]["MaxLevel"]["type"] == "int"
    assert mgr.config["SkillData"]["columns"]["MaxLevel"]["note"] == "上限"
    assert mgr.config["SkillData"]["validations"][0]["name"] == "測試規則"
    # 是複製不是搬移：舊 entry 還在，且已寫回 config.json
    with open(cfg_path, encoding="utf-8") as f:
        raw = json.load(f)
    assert os.path.normpath(jpath_a) in raw
    assert os.path.normpath(jpath_b) in raw
    print("  PASS  test_adopt_on_new_path")


def test_exact_match_no_adopt():
    tmp, cfg_path, jpath_a = setup_machine_a()
    mgr = JsonDataManager(config_path=cfg_path)
    mgr.load_json(jpath_a)                      # 路徑一致 → 不觸發沿用
    assert mgr.adopted_config_from is None
    assert mgr.config["SkillData"]["columns"]["MaxLevel"]["type"] == "int"
    print("  PASS  test_exact_match_no_adopt")


def test_multiple_candidates_prefers_recent():
    tmp, cfg_path, jpath_a = setup_machine_a()
    # 再造一個同名檔（機器A2）也載入過 → config 裡有兩個 SkillData.json entry
    dir_a2 = os.path.join(tmp, "machineA2"); os.makedirs(dir_a2)
    jpath_a2 = os.path.join(dir_a2, "SkillData.json")
    with open(jpath_a2, "w", encoding="utf-8") as f:
        json.dump(SKILLS, f, ensure_ascii=False)
    mgr = JsonDataManager(config_path=cfg_path)
    mgr.load_json(jpath_a2)
    mgr.config["SkillData"]["columns"]["MaxLevel"] = {"type": "float"}
    mgr.save_config()                            # A2 是最近開啟的

    dir_b = os.path.join(tmp, "machineB"); os.makedirs(dir_b)
    jpath_b = os.path.join(dir_b, "SkillData.json")
    shutil.copy(jpath_a, jpath_b)
    mgr2 = JsonDataManager(config_path=cfg_path)
    mgr2.load_json(jpath_b)
    # 兩個候選 → 沿用最近開啟的 A2（float）
    assert mgr2.adopted_config_from == os.path.normpath(jpath_a2)
    assert mgr2.config["SkillData"]["columns"]["MaxLevel"]["type"] == "float"
    print("  PASS  test_multiple_candidates_prefers_recent")


def test_no_candidate_defaults():
    tmp, cfg_path, jpath_a = setup_machine_a()
    other = os.path.join(tmp, "Other.json")
    with open(other, "w", encoding="utf-8") as f:
        json.dump(SKILLS, f, ensure_ascii=False)
    mgr = JsonDataManager(config_path=cfg_path)
    mgr.load_json(other)                         # 沒有同名 entry → 走預設
    assert mgr.adopted_config_from is None
    assert mgr.config["Other"]["columns"]["MaxLevel"]["type"] in ("int", "string")
    print("  PASS  test_no_candidate_defaults")


if __name__ == "__main__":
    tests = (test_adopt_on_new_path, test_exact_match_no_adopt,
             test_multiple_candidates_prefers_recent, test_no_candidate_defaults)
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
