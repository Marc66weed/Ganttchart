import os
import sys
import pandas as pd

def run_ci():
    print("🚀 [CI 流程] 啟動自動化檢查...")

    # Step 1: 檔案完整性檢查
    # 請確保這行字串與你資料夾中的檔名完全一致
    required_files = ["gantt_generater.py", "tasks_gantt.xlsx", "requirements.txt"]
    for f in required_files:
        if os.path.exists(f):
            print(f"✅ 找到關鍵檔案: {f}")
        else:
            print(f"❌ 錯誤: 缺少檔案 {f}，整合失敗！")
            sys.exit(1)

    # Step 2: 數據來源檢查 (Smoke Test)
    try:
        df = pd.read_excel("tasks_gantt.xlsx")
        print(f"✅ 成功讀取 Excel，共 {len(df)} 筆任務。")
    except Exception as e:
        print(f"❌ 錯誤: 無法讀取 Excel 檔案內容: {e}")
        sys.exit(1)

    print("\n🎊 [CI 成功] 目前環境與代碼基礎穩固，可以開始迭代！")

if __name__ == "__main__":
    run_ci()