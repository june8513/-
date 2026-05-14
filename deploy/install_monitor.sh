#!/bin/bash

# 檢查是否以 root 權限執行
if [ "$EUID" -ne 0 ]; then 
  echo "請使用 sudo 執行此腳本"
  echo "Usage: sudo ./install_monitor.sh"
  exit 1
fi

echo "--- 安裝自動上傳監控 (Auto Upload Monitor) ---"

# 複製檔案
echo "Copying service and timer files..."
cp /home/june/material-requisition/deploy/mr_monitor.service /etc/systemd/system/
cp /home/june/material-requisition/deploy/mr_monitor.timer /etc/systemd/system/

# 重載 systemd
echo "Reloading systemd daemon..."
systemctl daemon-reload

# 啟用 Timer (注意：是 enable timer，不是 service)
echo "Enabling and starting timer..."
systemctl enable mr_monitor.timer
systemctl start mr_monitor.timer

# 顯示狀態
echo "Checking timer status..."
systemctl status mr_monitor.timer --no-pager
echo "Checking service status (execution check)..."
systemctl status mr_monitor.service --no-pager

echo "--- 安裝完成 ---"
echo "現在系統每 1 分鐘會自動檢查 auto_upload 資料夾。"
