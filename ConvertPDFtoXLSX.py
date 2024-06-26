import os
import pandas as pd
import pdfplumber

# Defina o caminho do arquivo PDF
pdf_path = r"C:\Convert-PDF\v3.pdf"
# Defina o caminho do arquivo Excel a ser salvo
excel_path = r"C:\Convert-PDF\v5.xlsx"

def extract_tables_from_pdf(pdf_path):
    tables = []
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            extracted_tables = page.extract_tables()
            for table in extracted_tables:
                if table:
                    tables.append(table)
    return tables

# Extrair tabelas do PDF
tables = extract_tables_from_pdf(pdf_path)

# Converter as tabelas extraídas em DataFrames
dataframes = [pd.DataFrame(table[1:], columns=table[0]) for table in tables if table]

# Combinar todos os DataFrames em um único DataFrame
combined_df = pd.concat(dataframes, ignore_index=True)

# Salvar o DataFrame combinado em um arquivo Excel
combined_df.to_excel(excel_path, index=False)

print(f"Arquivo Excel salvo em: {excel_path}")

# pip install pdfplumber pandas openpyxl