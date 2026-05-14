#!/bin/bash

# 設定路徑
PROJECT_DIR="/home/june/material-requisition"
BACKUP_DIR="${PROJECT_DIR}/db_backups"
DB_FILE="${PROJECT_DIR}/db.sqlite3"

# 確保備份資料夾存在
mkdir -p "${BACKUP_DIR}"

# 取得當前時間戳記 (例如: 20260514_130000)
TIMESTAMP=$(date +"%Y%m%d_%H%M%S")
BACKUP_FILE="${BACKUP_DIR}/db_backup_${TIMESTAMP}.sqlite3"

# 複製並備份資料庫
if [ -f "$DB_FILE" ]; then
    cp "$DB_FILE" "$BACKUP_FILE"
fi

# 清除超過 10 天的舊備份檔案
find "${BACKUP_DIR}" -name "db_backup_*.sqlite3" -type f -mtime +10 -delete
