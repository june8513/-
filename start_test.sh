#!/bin/bash

# 定義原始專案的 Venv 路徑 (借用原本的環境，省去重新安裝的時間)
VENV_PYTHON="/home/june/-/venv/bin/python"

# 確保在正確的目錄
cd /home/june/mr_test

echo "=================================================="
echo "正在啟動 測試用伺服器 (Test Server)"
echo "--------------------------------------------------"
echo "網站位置: http://192.168.6.137:8001"
echo "使用資料庫: /home/june/mr_test/db.sqlite3 (獨立資料庫)"
echo "程式碼位置: /home/june/mr_test/"
echo "--------------------------------------------------"
echo "請按 Ctrl+C 停止伺服器"
echo "=================================================="

# 使用原始環境的 Python 執行測試專案的 manage.py
$VENV_PYTHON manage.py runserver 0.0.0.0:8001
