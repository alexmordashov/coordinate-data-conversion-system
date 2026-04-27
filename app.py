import streamlit as st
import pandas as pd
import requests
from io import BytesIO
import os
import re
from urllib.parse import urljoin
import tempfile
import pypandoc

st.set_page_config(
    page_title="Excel Analyzer",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

BACKEND_URL = "http://127.0.0.1:8000"

def check_api_status():
    try:
        response = requests.get(BACKEND_URL, timeout=10)
        return response.status_code == 200
    except Exception:
        return False

def process_excel(file):
    url = urljoin(BACKEND_URL, "/process-excel/")
    files = {"file": file}
    try:
        response = requests.post(url, files=files)
        if response.status_code == 200:
            return response.content.decode('utf-8')
        else:
            st.error(f"Ошибка при обработке файла: {response.text}")
            return None
    except requests.RequestException as e:
        st.error(f"Ошибка соединения с API: {str(e)}")
        return None

def main():
    st.title("📊 Анализатор Excel файлов")
    st.markdown("""Этот инструмент позволяет загрузить Excel-файл и получить аналитический отчёт в формате Markdown. Просто загрузите файл и нажмите кнопку "Анализировать".""")
    if not check_api_status():
        st.error("⚠️ Не удалось подключиться к API. Пожалуйста, проверьте соединение или попробуйте позже.")
        return

    uploaded_file = st.file_uploader("Выберите Excel файл", type=['xlsx', 'xls', 'csv'])

    if uploaded_file is not None:
        try:
            if uploaded_file.name.endswith('.csv'):
                df = pd.read_csv(uploaded_file)
            else:
                df = pd.read_excel(uploaded_file, engine='openpyxl')
            st.subheader("Предварительный просмотр данных")
            st.dataframe(df)
            st.subheader("Базовая информация")
            col1, col2, col3 = st.columns(3)
            col1.metric("Строки", df.shape[0])
            col2.metric("Столбцы", df.shape[1])
            col3.metric("Пропущенные значения", df.isna().sum().sum())
            uploaded_file.seek(0)
            if st.button("Преобразовать"):
                with st.spinner("Обрабатываем данные..."):
                    markdown_report = process_excel(uploaded_file)
                if markdown_report:
                    st.success("Отчёт успешно создан!")
                    st.subheader("Отчёт")
                    st.markdown(markdown_report)
                    
                    with tempfile.NamedTemporaryFile(suffix=".docx", delete=False) as tmp_file:
                        tmp_path = tmp_file.name

                    try:
                        pypandoc.convert_text(
                            markdown_report,
                            to='docx',
                            format='md',
                            outputfile=tmp_path
                        )

                        with open(tmp_path, 'rb') as f:
                            docx_bytes = f.read()
                        
                        if os.path.exists(tmp_path):
                            os.remove(tmp_path)
                        
                        st.download_button(
                            label="📄 Скачать отчёт (docx)",
                            data=docx_bytes,
                            file_name="report.docx",
                            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                        )
                    except Exception as e:
                        st.warning(f"Не удалось создать DOCX: {e}. Убедитесь, что Pandoc установлен.")
                        
                    st.download_button(
                        label="📄 Скачать отчёт (md)",
                        data=markdown_report,
                        file_name="report.md",
                        mime="text/markdown"
                    )
        except Exception as e:
            st.error(f"Ошибка при чтении файла: {str(e)}")

if __name__ == "__main__":
    main()
