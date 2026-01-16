"""
Gera um arquivo Excel de exemplo com colunas: date, product, quantity, price
"""
import pandas as pd
from datetime import datetime
data = [
    ("2025-01-05", "Camiseta", 3, 49.90),
    ("2025-01-15", "Tênis", 1, 299.90),
    ("2025-02-02", "Camiseta", 2, 49.90),
    ("2025-02-20", "Bermuda", 4, 79.90),
    ("2025-03-01", "Tênis", 2, 299.90),
    ("2025-03-18", "Meia", 10, 9.90),
    ("2025-04-03", "Camiseta", 5, 49.90),
    ("2025-04-22", "Boné", 3, 39.90),
    ("2025-05-10", "Bermuda", 2, 79.90),
    ("2025-05-30", "Tênis", 1, 299.90),
]
df = pd.DataFrame(data, columns=["date", "product", "quantity", "price"])
df["date"] = pd.to_datetime(df["date"])
# salva como Excel (.xlsx)
out = r"c:\Users\Usuario\Projetos Python\excel_to_pdf\sample_sales.xlsx"
df.to_excel(out, index=False, engine="openpyxl")
print(f"Arquivo de exemplo criado: {out}")