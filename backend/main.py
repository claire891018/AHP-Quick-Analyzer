from fastapi import FastAPI, File, UploadFile, HTTPException
from fastapi.responses import FileResponse, JSONResponse
from io import BytesIO
import tempfile
import pandas as pd
from src.excel_processer import list_sheets, read_excel, export_to_excel
from src.ahp_matrix import build_ahp_matrix
from src.weights_calculator import calculate_weights, calculate_geometric_mean
from src.consistency_checker import calculate_consistency_ratio

app = FastAPI()

def convert_to_bytes(file: UploadFile) -> BytesIO:
    """將 UploadFile 轉為 BytesIO"""
    content = file.file.read()
    file_bytes = BytesIO(content)
    file_bytes.seek(0)
    return file_bytes

@app.post("/list-sheets/")
async def get_sheets(file: UploadFile = File(...)):
    try:
        file_bytes = convert_to_bytes(file)
        sheets = list_sheets(file_bytes)
        return JSONResponse(content={"sheets": sheets})
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"無法讀取 Excel：{e}")

@app.post("/analyze/")
async def analyze_excel(
    file: UploadFile = File(...), 
    sheet_name: str = "Sheet1", 
    start_row: int = 1, 
    end_row: int = 5
):
    try:
        file_bytes = convert_to_bytes(file)
        df = read_excel(file_bytes, sheet_name)

        selected_rows = df.iloc[start_row - 1:end_row]
        results = []

        for index, row in selected_rows.iterrows():
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

        # 使用臨時檔案來存儲 Excel
        with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as temp_file:
            output_path = temp_file.name
            export_to_excel(results, output_path)

        return FileResponse(
            output_path, 
            filename="analysis_results.xlsx", 
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"分析失敗：{e}")
