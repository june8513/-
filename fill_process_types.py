import pandas as pd
import os

def main():
    file_path = r"C:\Users\gfpap\june\output.xlsx"
    output_path = r"C:\Users\gfpap\june\output_updated.xlsx"

    if not os.path.exists(file_path):
        print(f"找不到檔案: {file_path}")
        return

    try:
        # Read all sheets (sheet_name=None returns a dict of DataFrames)
        print("正在讀取所有分頁...")
        sheets_dict = pd.read_excel(file_path, sheet_name=None, dtype=str)
    except Exception as e:
        print(f"讀取 Excel 失敗: {e}")
        return

    # Combine all sheets into one DataFrame
    df_list = []
    for sheet_name, sheet_df in sheets_dict.items():
        sheet_df.columns = sheet_df.columns.str.strip()
        sheet_df['Original_Sheet'] = sheet_name
        df_list.append(sheet_df)
    
    if not df_list:
        print("Excel 檔案中沒有資料。")
        return

    df = pd.concat(df_list, ignore_index=True)
    
    print("-" * 40)
    print("合併後的 Excel 欄位:")
    for i, col in enumerate(df.columns):
        print(f"[{i}] {col}")
    print("-" * 40)
    
    # Mapping logic
    col_map = {
        'material': ['物料', '拇'], 
        'model': ['機型', '璈'],
        'process': ['投料點', '暺']
    }
    
    cols = {}
    for key, candidates in col_map.items():
        found = False
        for c in candidates:
            if c in df.columns:
                cols[key] = c
                found = True
                break
        if not found:
            if len(df.columns) >= 3:
                if key == 'material': cols[key] = df.columns[0]
                elif key == 'model': cols[key] = df.columns[1]
                elif key == 'process': cols[key] = df.columns[2]
            else:
                print(f"無法識別欄位 '{key}'。")
                return

    print(f"系統自動識別: 物料='{cols['material']}', 機型='{cols['model']}', 投料點='{cols['process']}'")

    # Ask user for helper column
    helper_col_input = input(f"請輸入要顯示的輔助資訊欄位編號 (例如品名/說明) [預設為 1: {df.columns[1]}]: ").strip()
    if not helper_col_input:
        helper_col = df.columns[1] if len(df.columns) > 1 else df.columns[0]
    else:
        try:
            helper_col = df.columns[int(helper_col_input)]
        except:
            print("無效編號，使用預設值。")
            helper_col = df.columns[1] if len(df.columns) > 1 else df.columns[0]
            
    print(f"輔助資訊欄位: {helper_col}")

    options = {
        '1': '機械',
        '2': '系統',
        '3': '電裝',
        '4': '鑄件',
        '5': '護蓋',
        '6': '刀庫',
        '7': '出貨',
        '8': '組件',
    }

    # Filter empty rows
    empty_mask = df[cols['process']].isna() | (df[cols['process']].astype(str).str.strip() == '') | (df[cols['process']].astype(str).str.lower() == 'nan')
    
    target_indices = df[empty_mask].index.tolist()
    total_empty = len(target_indices)
    
    print(f"\n發現 {total_empty} 筆資料缺少投料點 (來自所有分頁)。\n")

    if total_empty == 0:
        print("沒有需要處理的資料。")
        return

    processed_count = 0
    
    try:
        for idx in target_indices:
            row = df.loc[idx]
            material = row[cols['material']]
            model = row[cols['model']]
            helper_val = row[helper_col]
            sheet_origin = row.get('Original_Sheet', 'Unknown')
            
            print("-" * 40)
            print(f"進度: {processed_count + 1}/{total_empty}")
            print(f"來源分頁: {sheet_origin}")
            print(f"物料: {material}")
            print(f"機型: {model}")
            print(f"{helper_col}: {helper_val}")
            print("-" * 40)
            print("請選擇投料點:")
            for key, val in options.items():
                print(f"[{key}] {val}")
            print("[s] 跳過 (Skip)")
            print("[q] 儲存並離開 (Quit)")
            
            while True:
                choice = input("輸入選項: ").strip().lower()
                
                if choice == 'q':
                    print("正在儲存並離開...")
                    df.to_excel(output_path, index=False)
                    print(f"已儲存至: {output_path}")
                    return
                
                if choice == 's':
                    print("已跳過。")
                    break
                
                if choice in options:
                    selected_val = options[choice]
                    df.at[idx, cols['process']] = selected_val
                    print(f"已設定為: {selected_val}")
                    processed_count += 1
                    break
                
                print("無效的選項，請重新輸入。")
                
    except KeyboardInterrupt:
        print("\n使用者中斷，正在儲存...")
    
    df.to_excel(output_path, index=False)
    print(f"\n處理完成！共更新 {processed_count} 筆資料。")
    print(f"所有分頁已合併並儲存至: {output_path}")

if __name__ == "__main__":
    main()
