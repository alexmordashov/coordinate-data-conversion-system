from fastapi import FastAPI, UploadFile, File, HTTPException
from fastapi.responses import StreamingResponse
from io import BytesIO
import pandas as pd
import numpy as np
from typing import List
import matplotlib.pyplot as plt
import seaborn as sns
from datetime import datetime
from sympy import symbols, parse_expr, latex, Matrix
import json

app = FastAPI(
    title='Excel to Markdown API',
    description='API для обработки Excel файлов и генерации отчётов в формате Markdown',
    version='1.0.0'
)

@app.get('/')
def read_root():
    return {
        "message": "Excel to Markdown работает",
        "endpoints": {
            "/process-excel/": "Загрузка и обработка Excel-файла с генерацией отчёта в формате Markdown"
        }
    }

@app.post("/process-excel/")
async def process_excel(file: UploadFile = File(...)):
    if not file.filename.endswith(('.xlsx', '.xls', '.csv')):
        raise HTTPException(
            status_code=400,
            detail="Поддерживаются только файлы Excel (.xlsx, .xls, .csv)"
        )
    try:
        contents = await file.read()
        buffer = BytesIO(contents)
        if file.filename.endswith('.csv'):
            df = pd.read_csv(buffer)
        else:
            df = pd.read_excel(buffer, engine='openpyxl')
        with open('parameters.json', 'r', encoding='utf-8') as f:
            parameters = json.load(f)
        param = parameters['СК-42']
        transformed_df = calculate(df.copy(), param)
        #df = calculate(df)
        if df is None:
            raise HTTPException(status_code=500, detail="calculate вернул None")
        report = generate_markdown_report(df, transformed_df, param, 'СК-42', 'ГСК-2011')
        #report = generate_markdown_report(df)
        output = BytesIO(report.encode())
        #output.seek(0)
        filename = f"report_{datetime.now().strftime('%Y%m%d_%H%M%S')}.md"
        return StreamingResponse(
            output,
            media_type="text/markdown",
            headers={"Content-Disposition": f"attachment; filename={filename}"}
        )
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Ошибка обработки файла: {str(e)}")

def generate_markdown_report(df, trans_df, p, start_, end_) -> str:
    x, y, z = symbols('X Y Z')
    dx, dy, dz = symbols(r'\Delta\ X, \Delta\ Y, \Delta\ Z')
    wx, wy, wz = symbols(r'\omega_x, \omega_y, \omega_z')
    m = symbols('m')
    vector_b = Matrix(symbols('X Y Z'))
    vector_a = Matrix([x, y, z])
    trans_vector = Matrix([dx, dy, dz])
    rotation_matrix = Matrix([
        [1, wz, -wy],
        [-wz, 1, wx],
        [wy, -wx, 1]
    ])

    formula = f"{latex(vector_b)} = (1 + m) {latex(rotation_matrix)} {latex(vector_a)} + {latex(trans_vector)}"

    report_content = f"# Отчёт по преобразованию координат\n\n"
    report_content += f"## 1. Введение\n\n"
    report_content += "В этом отчёте представлены результаты преобразований координат.\n\n"
    report_content += f"## 2. Параметры ввода\n\n"
    report_content += f"- **Исходная таблица данных**: {len(df)} точек\n"
    report_content += f"- **Начальная система**: {start_}\n"
    report_content += f"- **Конечная система**: {end_}\n"
    report_content += f"- **Параметры**: {p}\n\n"
    report_content += f"## 3. Общая формула перехода между выбранными системами\n\n$$ {formula} $$\n\n"
    report_content += f"## 4. Таблицы с координатами до и после преобразования\n\n"
    report_content += "Координаты до преобразований\n\n"
    report_content += "| Имя | Начальная X | Начальная Y | Начальная Z |\n"
    report_content += "| --- | --- | --- | --- |\n"

    for i in range(len(df)):
      original = df.iloc[i]
      report_content += f"| {original['Name']} | {original['X']} | {original['Y']} | {original['Z']} |\n"

    report_content += "\nКоординаты после преобразований\n\n"
    report_content += "| Имя | Конечная X | Конечная Y | Конечная Z |\n"
    report_content += "| --- | --- | --- | --- |\n"

    for i in range(len(df)):
      new = trans_df.iloc[i]
      report_content += f"| {new['Name']} | {new['X_new']} | {new['Y_new']} | {new['Z_new']} |\n"

    report_content += f"## 5. Вывод\n\n"
    report_content += "Процесс преобразования координат был успешно выполнен, с результатами, представленными выше."

    return report_content

def calculate(df, p):
    m_corr = 1 + (p['m'] * 1e-6)
    df['X_new'] = df['X'] * m_corr + p['dx']
    df['Y_new'] = df['Y'] * m_corr + p['dy']
    df['Z_new'] = df['Z'] * m_corr + p['dz']
    return df
