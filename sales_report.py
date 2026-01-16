"""
Script: sales_report.py
Descrição:
- Lê um arquivo Excel de vendas (colunas esperadas: date, product, quantity, price) ou (date, product, sales)
- Calcula total de vendas por produto e por mês
- Gera gráficos com matplotlib
- Gera um PDF com os resumos e os gráficos usando reportlab

Uso:
python sales_report.py caminho/para/vendas.xlsx caminho/para/relatorio.pdf
"""
import sys
import io
from datetime import datetime
import tempfile

import pandas as pd
import matplotlib.pyplot as plt

from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.lib.units import mm
from reportlab.platypus import (
    SimpleDocTemplate,
    Paragraph,
    Spacer,
    Image,
    Table,
    TableStyle,
)
from reportlab.lib.styles import getSampleStyleSheet

# -----------------------------
# Funções de processamento
# -----------------------------

def read_sales_excel(path):
    """
    Lê a planilha Excel em path e retorna um DataFrame.
    Tenta converter a coluna de data para datetime.
    """
    # pandas usa openpyxl como engine para .xlsx
    df = pd.read_excel(path, engine="openpyxl")
    # normaliza nomes de colunas para minúsculas
    df.columns = [c.strip().lower() for c in df.columns]

    # assegura que exista coluna de data
    date_cols = [c for c in df.columns if "date" in c or "data" in c]
    if date_cols:
        df['date'] = pd.to_datetime(df[date_cols[0]])
    else:
        raise ValueError("Não foi encontrada coluna de data (nome contendo 'date' ou 'data').")

    # detecta vendas (sales) ou calcula a partir de quantidade*preço
    if 'sales' in df.columns or 'valor' in df.columns:
        sales_col = 'sales' if 'sales' in df.columns else 'valor'
        df['sales'] = pd.to_numeric(df[sales_col], errors='coerce').fillna(0)
    else:
        # procura quantity/quantity-like e price/price-like
        qty_cols = [c for c in df.columns if 'qty' in c or 'quantity' in c or 'quantidade' in c]
        price_cols = [c for c in df.columns if 'price' in c or 'preco' in c or 'valor_unit' in c]
        if qty_cols and price_cols:
            df['quantity'] = pd.to_numeric(df[qty_cols[0]], errors='coerce').fillna(0)
            df['price'] = pd.to_numeric(df[price_cols[0]], errors='coerce').fillna(0)
            df['sales'] = df['quantity'] * df['price']
        else:
            # se não conseguir calcular, tenta usar uma coluna chamada 'amount' ou similar
            other = [c for c in df.columns if c in ('amount','valor_total','total')]
            if other:
                df['sales'] = pd.to_numeric(df[other[0]], errors='coerce').fillna(0)
            else:
                raise ValueError("Não foi possível identificar colunas de vendas. Espere 'sales' ou 'quantity'+'price'.")

    # assegura coluna product
    prod_cols = [c for c in df.columns if 'product' in c or 'produto' in c or 'item' in c]
    if prod_cols:
        df['product'] = df[prod_cols[0]].astype(str)
    else:
        raise ValueError("Não foi encontrada coluna de produto (nome contendo 'product' ou 'produto').")

    return df

def summarize_by_product(df):
    """
    Agrupa por produto e soma vendas.
    Retorna DataFrame ordenado decrescentemente por vendas.
    """
    grp = df.groupby('product', dropna=False)['sales'].sum().reset_index()
    grp = grp.sort_values('sales', ascending=False).reset_index(drop=True)
    return grp

def summarize_by_month(df, date_col='date'):
    """
    Agrupa por ano-mês (YYYY-MM) e soma vendas.
    Retorna DataFrame com coluna 'month' (datetime de primeiro dia do mês) e 'sales'.
    """
    df['month'] = pd.to_datetime(df[date_col]).dt.to_period('M').dt.to_timestamp()
    grp = df.groupby('month')['sales'].sum().reset_index().sort_values('month')
    return grp

# -----------------------------
# Funções de plotagem
# -----------------------------

def plot_top_products(df_products, top_n=10):
    """
    Gera um gráfico de barras dos top N produtos por vendas.
    Retorna objeto BytesIO com PNG.
    """
    top = df_products.head(top_n).iloc[::-1]  # reverte para barras horizontais do menor para o maior
    fig, ax = plt.subplots(figsize=(8, max(3, 0.5 * len(top))))
    ax.barh(top['product'], top['sales'], color='#7c5cff', alpha=0.9)
    ax.set_title(f'Top {min(top_n, len(top))} Produtos por Vendas')
    ax.set_xlabel('Vendas (unidade monetária)')
    ax.grid(axis='x', linestyle='--', alpha=0.3)
    plt.tight_layout()

    buf = io.BytesIO()
    fig.savefig(buf, format='png', dpi=150)
    plt.close(fig)
    buf.seek(0)
    return buf

def plot_monthly_sales(df_monthly):
    """
    Gera um gráfico de linhas (ou barras) com vendas por mês.
    Retorna objeto BytesIO com PNG.
    """
    fig, ax = plt.subplots(figsize=(10, 4))
    ax.plot(df_monthly['month'], df_monthly['sales'], marker='o', color='#00aaff')
    ax.set_title('Vendas por Mês')
    ax.set_ylabel('Vendas')
    ax.set_xlabel('Mês')
    ax.grid(alpha=0.25)
    fig.autofmt_xdate(rotation=45)
    plt.tight_layout()

    buf = io.BytesIO()
    fig.savefig(buf, format='png', dpi=150)
    plt.close(fig)
    buf.seek(0)
    return buf

# -----------------------------
# Função para gerar PDF
# -----------------------------

def create_pdf_report(output_path, df_products, df_monthly, charts_buffers, metadata=None):
    """
    Monta o PDF usando reportlab.platypus.
    charts_buffers: dict com {'top_products': BytesIO, 'monthly': BytesIO}
    metadata: dict opcional com informações (autor, título)
    """
    doc = SimpleDocTemplate(output_path, pagesize=A4, rightMargin=20*mm, leftMargin=20*mm, topMargin=20*mm, bottomMargin=20*mm)
    styles = getSampleStyleSheet()
    story = []

    # Título
    title = metadata.get('title') if metadata and 'title' in metadata else 'Relatório de Vendas'
    story.append(Paragraph(f'<b>{title}</b>', styles['Title']))
    subtitle_text = metadata.get('subtitle', f'Gerado em {datetime.now().strftime("%Y-%m-%d %H:%M:%S")}')
    story.append(Paragraph(subtitle_text, styles['Normal']))
    story.append(Spacer(1, 8))

    # Sumário rápido
    total_sales = df_products['sales'].sum()
    total_products = len(df_products)
    story.append(Paragraph(f'<b>Resumo rápido</b>', styles['Heading2']))
    story.append(Paragraph(f'Total de vendas: <b>{total_sales:,.2f}</b>', styles['Normal']))
    story.append(Paragraph(f'Produtos distintos: <b>{total_products}</b>', styles['Normal']))
    story.append(Spacer(1, 10))

    # Inserir gráfico de top produtos
    story.append(Paragraph('<b>Top produtos</b>', styles['Heading2']))
    img_top = Image(charts_buffers['top_products'], width=160*mm, height=90*mm)
    story.append(img_top)
    story.append(Spacer(1, 8))

    # Tabela: top produtos (mostra top 20)
    story.append(Paragraph('<b>Vendas por produto (Top 20)</b>', styles['Heading3']))
    top20 = df_products.head(20)
    table_data = [['Produto', 'Vendas']]
    for _, row in top20.iterrows():
        table_data.append([row['product'], f"{row['sales']:,.2f}"])
    tbl = Table(table_data, colWidths=[110*mm, 40*mm])
    tbl.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,0), colors.HexColor('#f0f0f0')),
        ('TEXTCOLOR', (0,0), (-1,0), colors.black),
        ('ALIGN', (1,1), (-1,-1), 'RIGHT'),
        ('GRID', (0,0), (-1,-1), 0.25, colors.grey),
        ('FONTNAME', (0,0), (-1,0), 'Helvetica-Bold'),
    ]))
    story.append(tbl)
    story.append(Spacer(1, 12))

    # Inserir gráfico mensal
    story.append(Paragraph('<b>Vendas por mês</b>', styles['Heading2']))
    img_month = Image(charts_buffers['monthly'], width=160*mm, height=70*mm)
    story.append(img_month)
    story.append(Spacer(1, 12))

    # Tabela: vendas por mês
    story.append(Paragraph('<b>Vendas por mês</b>', styles['Heading3']))
    table_data = [['Mês', 'Vendas']]
    for _, row in df_monthly.iterrows():
        month_label = pd.to_datetime(row['month']).strftime('%Y-%m')
        table_data.append([month_label, f"{row['sales']:,.2f}"])
    tbl2 = Table(table_data, colWidths=[110*mm, 40*mm])
    tbl2.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,0), colors.HexColor('#f0f0f0')),
        ('TEXTCOLOR', (0,0), (-1,0), colors.black),
        ('ALIGN', (1,1), (-1,-1), 'RIGHT'),
        ('GRID', (0,0), (-1,-1), 0.25, colors.grey),
        ('FONTNAME', (0,0), (-1,0), 'Helvetica-Bold'),
    ]))
    story.append(tbl2)

    # Constrói o PDF
    doc.build(story)

# -----------------------------
# Fluxo principal
# -----------------------------

def main(argv):
    if len(argv) < 3:
        print("Uso: python sales_report.py caminho/para/vendas.xlsx caminho/para/relatorio.pdf")
        sys.exit(1)
    excel_path = argv[1]
    pdf_path = argv[2]

    # 1) Ler Excel
    df = read_sales_excel(excel_path)

    # 2) Gerar resumos
    df_products = summarize_by_product(df)
    df_monthly = summarize_by_month(df)

    # 3) Gerar gráficos (em memória)
    buf_top = plot_top_products(df_products, top_n=10)
    buf_month = plot_monthly_sales(df_monthly)

    charts = {
        'top_products': buf_top,
        'monthly': buf_month,
    }

    # 4) Criar PDF
    metadata = {'title': 'Relatório de Vendas', 'subtitle': f'Fonte: {excel_path}'}
    create_pdf_report(pdf_path, df_products, df_monthly, charts, metadata=metadata)

    print(f"Relatório gerado: {pdf_path}")

if __name__ == '__main__':
    main(sys.argv)