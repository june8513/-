# 投料點分類常數
PROCESS_CATEGORIES = [
    ('組件', '組件'),
    ('電裝', '電裝'),
    ('機械', '機械'),
    ('鑄件', '鑄件'),
    ('系統', '系統'),
    ('護蓋', '護蓋'),
    ('刀庫', '刀庫'),
    ('出貨', '出貨'),
    ('軟體研發部', '軟體研發部'),
    ('其他', '其他'),
]

# 投料點分類名稱列表
PROCESS_CATEGORY_NAMES = [name for name, _ in PROCESS_CATEGORIES]

# 投料點分類顏色映射（用於前端顯示）
PROCESS_CATEGORY_COLORS = {
    '組件': '#3B82F6',  # blue
    '電裝': '#22C55E',  # green
    '機械': '#F97316',  # orange
    '鑄件': '#A855F7',  # purple
    '系統': '#14B8A6',  # teal
    '護蓋': '#EF4444',  # red
    '刀庫': '#6366F1',  # indigo
    '出貨': '#F59E0B',  # amber
    '軟體研發部': '#EC4899',  # pink
    '其他': '#64748B',  # slate/gray
}

# 群組名稱常數
GROUP_NAMES = {
    'ADMIN': '管理員',
    'APPLICANT_SUPERVISOR': '申請人員主管',
    'APPLICANT': '申請人員',
    'DISPATCHER_SUPERVISOR': '撥料人員主管',
    'DISPATCHER': '撥料人員',
}
