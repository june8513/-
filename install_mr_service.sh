#!/bin/bash

# 檢查是否以 root 權限執行
if [ "$EUID" -ne 0 ]; then 
  echo "請使用 sudo 執行此腳本 (Please run as root)"
  echo "Usage: sudo ./install_mr_service.sh"
  exit 1
fi

echo "--- 安裝 Material Requisition 服務 ---"

# 複製服務檔案
echo "Copying service file to /etc/systemd/system/..."
cp /home/june/-/material_requisition.service /etc/systemd/system/

# 重載 systemd
echo "Reloading systemd daemon..."
systemctl daemon-reload

# 啟用開機自動啟動
echo "Enabling material_requisition service on boot..."
systemctl enable material_requisition

# 立即啟動服務
echo "Starting material_requisition service now..."
systemctl restart material_requisition

# 顯示狀態
echo "Checking service status..."
systemctl status material_requisition --no-pager

echo "--- 安裝完成 (Installation Complete) ---"
