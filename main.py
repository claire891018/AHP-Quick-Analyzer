import os
import pandas as pd
from src.excel_processer import list_sheets, read_excel, export_to_excel
from src.ahp_matrix import build_ahp_matrix
from src.weights_calculator import calculate_weights, calculate_geometric_mean
from src.consistency_checker import calculate_consistency_ratio

def main():
    file_path = os.path.join('data', 'comparison_data.xlsx')
    output_path = os.path.join('data', 'analysis_results.xlsx')

    sheets = list_sheets(file_path)
    if not sheets:
        print("無法找到任何工作表")
        return

    print("可用的工作表:")
    for idx, sheet in enumerate(sheets):
        print(f"{idx + 1}. {sheet}")

    while True:
        try:
            choice = int(input("請選擇要分析的工作表 (輸入編號): ")) - 1
            if 0 <= choice < len(sheets):
                sheet_name = sheets[choice]
                break
            else:
                print("請輸入有效的編號範圍")
        except ValueError:
            print("請輸入有效的數字")

    df = read_excel(file_path, sheet_name)
    if df is None:
        return

    start_row = int(input("請輸入要開始的行數: ")) - 1
    end_row = int(input("請輸入要結束的行數: "))

    selected_rows = df.iloc[start_row:end_row]

    results = []

    for index, row in selected_rows.iterrows():
        try:
            matrix = build_ahp_matrix(row)
            weights = calculate_weights(matrix)
            geometric_mean = calculate_geometric_mean(matrix)
            ci, cr = calculate_consistency_ratio(matrix, weights)
            results.append({
                'Row': index + 1,
                'Matrix': matrix.tolist(),
                'Weights': weights.tolist(),
                'Geometric Mean': geometric_mean.tolist(),
                'CI': ci,
                'CR': cr
            })
        except Exception as e:
            print(f"第 {index + 1} 行資料分析失敗：{e}")

    export_to_excel(results, output_path)
    print(f"分析結果已匯出到 {output_path}")

if __name__ == "__main__":
    main()