import numpy as np
import pandas as pd

def convert_to_numerical(row):
    """將比例字串轉為數值，並檢查資料格式"""
    numerical_data = []
    for item in row:
        item = str(item)

        if pd.isna(item) or item.strip() == "" or item.lower() == 'nan':
            print(f"警告：發現空值或無效資料 '{item}'，已跳過。")
            continue

        try:
            ratio = item.split(":")
            if len(ratio) != 2:
                raise ValueError("比例格式錯誤")
            numerical_data.append(float(ratio[0]) / float(ratio[1]))
        except (ValueError, IndexError) as e:
            print(f"資料格式錯誤：{item}，錯誤訊息：{e}")
            continue

    return numerical_data

def build_ahp_matrix(row):
    """從一行資料建立成對比較矩陣"""
    numerical_data = convert_to_numerical(row)
    num_elements = 4  # 4 個元素
    comparison_matrix = np.ones((num_elements, num_elements))
    index = 0

    for i in range(num_elements):
        for j in range(i + 1, num_elements):
            comparison_matrix[i, j] = numerical_data[index]
            comparison_matrix[j, i] = 1 / numerical_data[index]
            index += 1
    print(comparison_matrix)
    return comparison_matrix
