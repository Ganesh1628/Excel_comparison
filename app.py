# backend/main.py
import os
from fastapi import FastAPI, File, UploadFile, Form
from fastapi.responses import FileResponse, HTMLResponse
from fastapi.staticfiles import StaticFiles
from tempfile import NamedTemporaryFile
import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Font
from starlette.middleware.cors import CORSMiddleware

app = FastAPI()

# Serve static HTML page
app.mount("/static", StaticFiles(directory="static"), name="static")

# CORS for frontend usage
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

@app.get("/", response_class=HTMLResponse)
async def get_home():
    with open("static/index.html", "r", encoding="utf-8") as f:
        return f.read()


@app.post("/process")
async def process_excel(
    input_file: UploadFile = File(...),
    output_file: UploadFile = File(...)
):
    compare_columns = ['intellimindz(SERP)', 'Proxy  (SERP)', 'online job support(SERP)']
    arrow_colors = {"↑": "008000", "↓": "FF0000", "→": "808080"}

    # Save uploaded files
    input_path = f"/tmp/temp_{input_file.filename}"
    output_path = f"/tmp/temp_{output_file.filename}"

    with open(input_path, "wb") as f:
        f.write(await input_file.read())
    with open(output_path, "wb") as f:
        f.write(await output_file.read())
        
        result_file = "/tmp/SERP_Comparison_Result.xlsx"
    with pd.ExcelWriter(result_file, engine='openpyxl') as writer:
        sheet_names = pd.ExcelFile(input_path).sheet_names
        for sheet in sheet_names:
            try:
                df_week2 = pd.read_excel(input_path, sheet_name=sheet)
                df_week3 = pd.read_excel(output_path, sheet_name=sheet)
                comparison_df = df_week3[['S.No', 'Technology']].copy()
                difference_df = df_week3[['S.No', 'Technology']].copy()
                for col in compare_columns:
                    if col not in df_week2.columns or col not in df_week3.columns:
                        continue
                    results, diffs = [], []
                    for prev, curr in zip(df_week2[col], df_week3[col]):
                        if pd.isna(prev) or pd.isna(curr):
                            results.append("→ NA"); diffs.append("NA"); continue
                        if curr < prev:
                            results.append(f"↑ {curr}")
                        elif curr > prev:
                            results.append(f"↓ {curr}")
                        else:
                            results.append(f"→ {curr}")
                        diffs.append(abs(curr - prev))
                    comparison_df[col] = results
                    difference_df[col] = diffs
                comparison_df.to_excel(writer, sheet_name=f"{sheet}_Comparison", index=False)
                difference_df.to_excel(writer, sheet_name=f"{sheet}_Difference", index=False)
            except Exception as e:
                print(f"Error in sheet {sheet}: {e}")

    # Apply color formatting
    wb = load_workbook(result_file)
    for sheet in wb.sheetnames:
        if not sheet.endswith("_Comparison"):
            continue
        ws = wb[sheet]
        for col in range(3, ws.max_column + 1):
            for row in range(2, ws.max_row + 1):
                cell = ws.cell(row=row, column=col)
                value = cell.value
                if isinstance(value, str) and len(value) > 1:
                    arrow = value[0]
                    if arrow in arrow_colors:
                        cell.font = Font(color=arrow_colors[arrow], bold=True)
    wb.save(result_file)

    return FileResponse(result_file, filename="Processed_Output.xlsx")
