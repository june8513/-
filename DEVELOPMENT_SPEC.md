# 撥料申請系統 — 開發規範 (DEVELOPMENT_SPEC)

本文件定義了本專案的開發標準與架構規範，所有新增功能與程式碼異動均須遵守。

---

## 一、 架構規範 (Architecture)

### 1.1 App 職責劃分
*   **`work_orders`**: 負責工單 (WorkOrder)、機型 (MachineModel)、投料點規則及物料主檔。包含 SAP 資料同步與 Excel 原始資料匯入。
*   **`requisitions`**: 負責撥料申請單 (Requisition) 的生命週期，包含申請、撥料、欠料彙整及簽收邏輯。
*   **`inventory`**: 負責庫存更新、盤點及盤點差異管理。
*   **`core`**: 負責首頁路由、使用者角色判斷及全站共用 Context。
*   **`common`**: 存放全站共用的常數 (Constants)、權限檢查 (Permissions) 及通用工具函數。

### 1.2 App 內部目錄結構
每個 App 應遵循以下標準結構：
```text
apps/<app_name>/
├── models.py          # Model 定義 (單一檔案 ≤ 300 行)
├── services/          # 業務邏輯服務層 (不包含 HTTP request/response)
├── views/             # HTTP 處理視圖 (單一檔案 ≤ 400 行)
├── urls.py            # 路由定義
├── forms.py           # Django 表單
├── management/        # CLI 指令 (Management Commands)
└── templates/         # 模板文件
```

---

## 二、 程式碼實作規範 (Coding Standards)

### 2.1 邏輯分離 (Service Layer)
*   **View 的職責**：僅負責解析 Request、驗證權限、呼叫 Service 並回傳 Response。
*   **Service 的職責**：處理所有資料運算、外部整合、複雜的資料庫異動。**業務邏輯嚴禁寫在 View 中。**

### 2.2 規模限制
*   **檔案行數**：View 檔案不超過 400 行，Model 檔案不超過 300 行。
*   **函數行數**：單一函數原則上不超過 80 行，超過則應進行功能拆分。

### 2.3 命名規範
*   **檔案名稱**：使用 `snake_case` (例如 `material_data_views.py`)。
*   **類別名稱**：使用 `PascalCase` (例如 `RequisitionService`)。
*   **變數/函數**：使用 `snake_case`。

---

## 三、 新功能開發流程 (Workflow)

遵循以下步驟進行開發：
1.  **確認 App**：判斷功能屬於哪個業務領域。
2.  **實作 Service**：在對應 App 的 `services/` 下寫入業務邏輯。
3.  **實作 View**：建立視圖並呼叫 Service。
4.  **建立 Template**：實作前端介面。
5.  **配置 URL**：在 `urls.py` 中註冊路由。
6.  **提交代碼**：遵循 Git 提交規範。

---

## 四、 Git 提交規範 (Commit Message)

格式：`<type>(<scope>): <subject>`

*   **type**:
    *   `feat`: 新功能
    *   `fix`: 修補錯誤
    *   `refactor`: 重構（非新增功能也非修錯）
    *   `docs`: 文件變更
    *   `chore`: 建置程序或輔助工具的變更
*   **scope**: 影響的 App 名稱（例如 `requisitions`, `work_orders`）。

**範例**：
*   `feat(requisitions): 增加撥料單批次簽收功能`
*   `refactor(work_orders): 將 SAP 同步邏輯抽離至 service 層`

---

## 五、 維護性規範

*   **禁止使用散落腳本**：所有的資料維護或同步任務，必須寫成 Django Management Command，禁止直接在根目錄放置 `.py` 腳本。
*   **清理機制**：定期清理 `scratch/` 及過期的 debug 檔案，保持工作目錄整潔。
