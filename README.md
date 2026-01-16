# Sales Report — Excel → PDF (pandas, matplotlib, reportlab)

Resumo rápido
------------
Pipeline simples e profissional para transformar dados de vendas em um relatório PDF contendo:
- resumo por produto e por mês;
- tabelas e gráficos gerados automaticamente;
- output pronto para stakeholders.

Por que fiz este projeto?
--------------------------------------------------
- Demonstra ETL básico com pandas: limpeza, agregação e transformação de dados.  
- Produz visualizações (matplotlib) e artefatos finais (PDF) para tomada de decisão.  
- Mostra automação de fluxo (scripts + PowerShell) e preocupação com reprodutibilidade (venv, requirements).  
- Código comentado e modular — fácil de estender para pipelines reais (APIs, dashboards, CI).

Tecnologias
-----------
- Python 3.x  
- pandas, openpyxl — leitura/transformação de Excel  
- matplotlib — gráficos  
- reportlab — criação de PDF profissional

Destaques do que foi implementado
--------------------------------
- Leitura flexível de Excel (colunas típicas: date/data, product/produto, quantity/preço ou sales).  
- Resumo de vendas por produto (Top-N) e por mês.  
- Gráficos gerados em memória e inseridos no PDF.  
- Relatório com título, sumário rápido, gráficos e tabelas (top produtos e por mês).  
- Scripts auxiliares: criação de planilha de exemplo e PowerShell para automatizar execução.

Pré-requisitos (Windows)
------------------------
- Python 3.8+ instalado  
- PowerShell (Windows)  
- Recomenda-se Git para versionamento

Instalação e execução (rápido)
------------------------------
Abra PowerShell na pasta do projeto (exemplo):

```powershell
cd "C:\Users\Usuario\Projetos Python\excel_to_pdf"
```

1) Criar e ativar ambiente virtual
```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
# se houver erro de execução no PowerShell:
# Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser
```

2) Instalar dependências
```powershell
pip install --upgrade pip
pip install -r requirements.txt
```

3) (Opcional) Gerar planilha de exemplo
```powershell
python create_sample_excel.py
```

4) Gerar o relatório PDF
```powershell
python sales_report.py sample_sales.xlsx relatorio_vendas.pdf
```

5) Abrir o PDF (Windows)
```powershell
Start-Process .\relatorio_vendas.pdf
```

Atalho (PowerShell)
-------------------
Existe um script `run_report.ps1` que cria o venv, instala dependências, gera o sample (se necessário) e produz o PDF. Basta executar:
```powershell
.\run_report.ps1
```

Estrutura do repositório
------------------------
- sales_report.py         — script principal (comentado, modular)  
- create_sample_excel.py  — gera `sample_sales.xlsx` de exemplo  
- sample_sales.xlsx       — (opcional) planilha de teste gerada  
- requirements.txt        — dependências do projeto  
- run_report.ps1          — script PowerShell para automatizar execução  
- .gitignore              — ignorados (venv, outputs, etc.)

Futuras melhorias que irei fazer
-------------------------------------------------
- Adicionar suporte a múltiplas abas/formatos (CSV, Parquet).  
- Gerar relatórios por região/cliente; combinar com geodados.  
- Adicionar testes unitários para a lógica de agregação (pytest).  
- Automatizar via GitHub Actions: rodar e gerar artefato PDF a cada push.  
- Reescrever geração de PDF em templates (Jinja2 → HTML → PDF) para layouts ricos.
