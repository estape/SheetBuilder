/**
 * View.js - Camada de Apresentação do SheetBuilder
 * Gerenciamento de interface, estados visuais, gabaritos financeiros executivos e renderização de tabelas.
 */

const PRESETS = {
  dre_gerencial: `# Gabarito: Demonstração do Resultado do Exercício (DRE Gerencial Completa)
# Padrão: Controladoria & Consultoria Executiva (pt-BR)
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

wb = openpyxl.Workbook()
ws = wb.active
ws.title = "DRE_Gerencial"
ws.views.sheetView[0].showGridLines = True

# Paleta Executiva (Azul Marinho Corporativo)
HEADER_FILL = PatternFill(start_color="1B365D", end_color="1B365D", fill_type="solid")
SUBTOTAL_FILL = PatternFill(start_color="EAEFF7", end_color="EAEFF7", fill_type="solid")
HIGHLIGHT_FILL = PatternFill(start_color="D1E2F2", end_color="D1E2F2", fill_type="solid")
ZEBRA_FILL = PatternFill(start_color="F8FAFC", end_color="F8FAFC", fill_type="solid")

FONT_HEADER = Font(name="Segoe UI", size=10, bold=True, color="FFFFFF")
FONT_TITLE = Font(name="Segoe UI", size=13, bold=True, color="1B365D")
FONT_BOLD = Font(name="Segoe UI", size=10, bold=True, color="0F172A")
FONT_REGULAR = Font(name="Segoe UI", size=10, color="334155")

THIN_BORDER = Border(
    left=Side(style='thin', color='E2E8F0'),
    right=Side(style='thin', color='E2E8F0'),
    top=Side(style='thin', color='E2E8F0'),
    bottom=Side(style='thin', color='E2E8F0')
)
DOUBLE_BOTTOM_BORDER = Border(
    top=Side(style='thin', color='94A3B8'),
    bottom=Side(style='double', color='0F172A'),
    left=Side(style='thin', color='E2E8F0'),
    right=Side(style='thin', color='E2E8F0')
)

# Título do Relatório
ws.merge_cells("A1:G1")
title_cell = ws["A1"]
title_cell.value = "DEMONSTRAÇÃO DO RESULTADO DO EXERCÍCIO (DRE GERENCIAL)"
title_cell.font = FONT_TITLE
title_cell.alignment = Alignment(horizontal="left", vertical="center")
ws.row_dimensions[1].height = 30

headers = ["Estrutura de Contas", "2024 Realizado", "2025 Orçado", "2025 Realizado", "Análise Vertical (AV %)", "Variação (R$)", "Desvio (%)"]
for col_idx, h in enumerate(headers, 1):
    c = ws.cell(row=2, column=col_idx, value=h)
    c.fill = HEADER_FILL
    c.font = FONT_HEADER
    c.alignment = Alignment(horizontal="center" if col_idx > 1 else "left", vertical="center")
    c.border = THIN_BORDER
ws.row_dimensions[2].height = 24

# Linhas da DRE: [Descrição, 2024, 2025 Orç, 2025 Real, Tipo]
rows = [
    ["1. RECEITA BRUTA DE VENDAS E SERVIÇOS", 14500000, 16800000, 17250000, "data"],
    ["  (-) Deduções da Receita e Impostos s/ Vendas", -2175000, -2520000, -2587500, "data"],
    ["(=) RECEITA OPERACIONAL LÍQUIDA (ROL)", "=B3+B4", "=C3+C4", "=D3+D4", "subtotal"],
    ["  (-) Custo das Vendas e Serviços (CPV/CSP)", -6525000, -7560000, -7762500, "data"],
    ["(=) LUCRO BRUTO (MARGEM BRUTA)", "=B5+B6", "=C5+C6", "=D5+D6", "subtotal"],
    ["  (-) Despesas Comerciais e Marketing (CAC)", -1305000, -1512000, -1580000, "data"],
    ["  (-) Despesas Gerais e Administrativas (SG&A)", -1885000, -2016000, -2050000, "data"],
    ["(=) EBITDA / LAJIDA (Geração Operacional)", "=B7+B8+B9", "=C7+C8+C9", "=D7+D8+D9", "subtotal"],
    ["  (-) Depreciação e Amortização", -420000, -450000, -460000, "data"],
    ["(=) RESULTADO OPERACIONAL (EBIT / LAJIR)", "=B10+B11", "=C10+C11", "=D10+D11", "subtotal"],
    ["  (+/-) Resultado Financeiro Líquido", -280000, -310000, -295000, "data"],
    ["(=) RESULTADO ANTES DO IR/CSLL (LAIR)", "=B12+B13", "=C12+C13", "=D12+D13", "subtotal"],
    ["  (-) Provisão para IRPJ e CSLL", -650000, -720000, -745000, "data"],
    ["(=) LUCRO LÍQUIDO DO EXERCÍCIO", "=B14+B15", "=C14+C15", "=D14+D15", "total"]
]

for r_offset, r_data in enumerate(rows, 3):
    desc, v24, v25_b, v25_a, r_type = r_data
    ws.row_dimensions[r_offset].height = 20
    
    ws.cell(row=r_offset, column=1, value=desc)
    ws.cell(row=r_offset, column=2, value=v24)
    ws.cell(row=r_offset, column=3, value=v25_b)
    ws.cell(row=r_offset, column=4, value=v25_a)
    ws.cell(row=r_offset, column=5, value=f"=D{r_offset}/$D$5")
    ws.cell(row=r_offset, column=6, value=f"=D{r_offset}-C{r_offset}")
    ws.cell(row=r_offset, column=7, value=f"=IF(C{r_offset}<>0, (D{r_offset}-C{r_offset})/ABS(C{r_offset}), 0)")

    for col_i in range(1, 8):
        cell = ws.cell(row=r_offset, column=col_i)
        cell.border = THIN_BORDER
        if r_type == "total":
            cell.font = Font(name="Segoe UI", size=10, bold=True, color="1B365D")
            cell.fill = HIGHLIGHT_FILL
            cell.border = DOUBLE_BOTTOM_BORDER
        elif r_type == "subtotal":
            cell.font = FONT_BOLD
            cell.fill = SUBTOTAL_FILL
        else:
            cell.font = FONT_REGULAR
            if r_offset % 2 == 1:
                cell.fill = ZEBRA_FILL
                
        if 2 <= col_i <= 4 or col_i == 6:
            cell.number_format = 'R$ #,##0;[Red]-R$ #,##0;"-"'
            cell.alignment = Alignment(horizontal="right")
        elif col_i in (5, 7):
            cell.number_format = '0.0%'
            cell.alignment = Alignment(horizontal="right")

for col in ws.columns:
    max_len = max(len(str(cell.value or '')) for cell in col)
    ws.column_dimensions[get_column_letter(col[0].column)].width = max(max_len + 3, 14)
ws.column_dimensions["A"].width = 46

ws.freeze_panes = "A3"
wb.save("DRE_Gerencial_Executiva.xlsx")
print("Demonstração do Resultado do Exercício gerada com sucesso: DRE_Gerencial_Executiva.xlsx")
`,

  painel_kpis: `# Gabarito: Painel Executivo de Indicadores Financeiros (KPIs & Health Check)
# Padrão: C-Level, Conselho de Administração e M&A (pt-BR)
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

wb = openpyxl.Workbook()
ws = wb.active
ws.title = "Painel_Indicadores"
ws.views.sheetView[0].showGridLines = True

PRIMARY_FILL = PatternFill(start_color="0F766E", end_color="0F766E", fill_type="solid") # Verde Petróleo
ZEBRA_FILL = PatternFill(start_color="F8FAFC", end_color="F8FAFC", fill_type="solid")

FONT_HEADER = Font(name="Segoe UI", size=10, bold=True, color="FFFFFF")
FONT_TITLE = Font(name="Segoe UI", size=13, bold=True, color="0F766E")
FONT_REGULAR = Font(name="Segoe UI", size=10, color="1E293B")
FONT_BOLD = Font(name="Segoe UI", size=10, bold=True, color="0F172A")

THIN_BORDER = Border(
    left=Side(style='thin', color='CBD5E1'),
    right=Side(style='thin', color='CBD5E1'),
    top=Side(style='thin', color='CBD5E1'),
    bottom=Side(style='thin', color='CBD5E1')
)

ws.merge_cells("A1:F1")
ws["A1"] = "PAINEL EXECUTIVO DE INDICADORES FINANCEIROS (KPIs)"
ws["A1"].font = FONT_TITLE
ws["A1"].alignment = Alignment(horizontal="left", vertical="center")
ws.row_dimensions[1].height = 30

headers = ["Grupo / Categoria", "Indicador Financeiro", "Meta / Benchmark", "Realizado 2024", "Realizado 2025", "Avaliação de Desempenho"]
for col_idx, h in enumerate(headers, 1):
    c = ws.cell(row=2, column=col_idx, value=h)
    c.fill = PRIMARY_FILL
    c.font = FONT_HEADER
    c.alignment = Alignment(horizontal="center" if col_idx > 2 else "left", vertical="center")
    c.border = THIN_BORDER
ws.row_dimensions[2].height = 24

kpis = [
    ["Rentabilidade", "Margem Bruta (Lucro Bruto / ROL)", 0.40, 0.421, 0.435, "Superou a Meta", "pct"],
    ["Rentabilidade", "Margem EBITDA (EBITDA / ROL)", 0.18, 0.165, 0.198, "Superou a Meta", "pct"],
    ["Rentabilidade", "Margem Líquida (Lucro Líq. / ROL)", 0.08, 0.072, 0.091, "Dentro do Esperado", "pct"],
    ["Retorno", "ROE (Retorno sobre o Patrimônio Líquido)", 0.22, 0.195, 0.248, "Excelente Retorno", "pct"],
    ["Retorno", "ROIC (Retorno sobre Capital Investido)", 0.16, 0.142, 0.175, "Criação de Valor", "pct"],
    ["Liquidez", "Liquidez Corrente (Ativo Circ. / Passivo Circ.)", 1.40, 1.28, 1.52, "Saudável", "dec"],
    ["Liquidez", "Liquidez Seca ((AC - Estoques) / PC)", 1.00, 0.95, 1.12, "Saudável", "dec"],
    ["Endividamento", "Alavancagem (Dívida Líquida / EBITDA)", 2.00, 2.35, 1.78, "Desalavancagem Positiva", "mult"],
    ["Endividamento", "Cobertura de Juros (EBITDA / Juros)", 3.50, 3.20, 4.10, "Confortável", "mult"],
    ["Ciclo Financeiro", "Prazo Médio de Recebimento (PMR em dias)", 45, 52, 43, "Melhoria de Cobrança", "int"],
    ["Ciclo Financeiro", "Prazo Médio de Pagamento (PMP em dias)", 60, 55, 62, "Alongamento Favorável", "int"],
    ["Ciclo Financeiro", "Ciclo de Conversão de Caixa (CCC em dias)", 30, 41, 28, "Otimização de Capital", "int"]
]

for row_idx, data in enumerate(kpis, 3):
    cat, name, target, v24, v25, status, fmt = data
    ws.row_dimensions[row_idx].height = 20
    
    ws.cell(row=row_idx, column=1, value=cat)
    ws.cell(row=row_idx, column=2, value=name)
    ws.cell(row=row_idx, column=3, value=target)
    ws.cell(row=row_idx, column=4, value=v24)
    ws.cell(row=row_idx, column=5, value=v25)
    ws.cell(row=row_idx, column=6, value=status)

    for col_i in range(1, 7):
        cell = ws.cell(row=row_idx, column=col_i)
        cell.font = FONT_REGULAR
        cell.border = THIN_BORDER
        if row_idx % 2 == 1:
            cell.fill = ZEBRA_FILL
        
        if col_i in (3, 4, 5):
            cell.alignment = Alignment(horizontal="right")
            if fmt == "pct":
                cell.number_format = "0.0%"
            elif fmt == "dec":
                cell.number_format = "0.00"
            elif fmt == "mult":
                cell.number_format = '0.0"x"'
            elif fmt == "int":
                cell.number_format = '#,##0" dias"'
        elif col_i == 6:
            cell.alignment = Alignment(horizontal="center")
            cell.font = FONT_BOLD

for col in ws.columns:
    max_len = max(len(str(cell.value or '')) for cell in col)
    ws.column_dimensions[get_column_letter(col[0].column)].width = max(max_len + 3, 14)
ws.column_dimensions["A"].width = 18
ws.column_dimensions["B"].width = 46

ws.freeze_panes = "A3"
wb.save("Painel_Indicadores_Financeiros.xlsx")
print("Painel de Indicadores gerado com sucesso: Painel_Indicadores_Financeiros.xlsx")
`,

  fluxo_caixa: `# Gabarito: Fluxo de Caixa Direto e Gestão de Tesouraria
# Padrão: Gestão de Caixa & FP&A Corporativo
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

wb = openpyxl.Workbook()
ws = wb.active
ws.title = "Fluxo_de_Caixa"
ws.views.sheetView[0].showGridLines = True

HEADER_FILL = PatternFill(start_color="065F46", end_color="065F46", fill_type="solid")
SUBTOTAL_FILL = PatternFill(start_color="ECFDF5", end_color="ECFDF5", fill_type="solid")
ZEBRA_FILL = PatternFill(start_color="F8FAFC", end_color="F8FAFC", fill_type="solid")

FONT_HEADER = Font(name="Segoe UI", size=10, bold=True, color="FFFFFF")
FONT_TITLE = Font(name="Segoe UI", size=13, bold=True, color="065F46")
FONT_BOLD = Font(name="Segoe UI", size=10, bold=True, color="065F46")
FONT_REGULAR = Font(name="Segoe UI", size=10, color="1E293B")

THIN_BORDER = Border(
    left=Side(style='thin', color='E2E8F0'),
    right=Side(style='thin', color='E2E8F0'),
    top=Side(style='thin', color='E2E8F0'),
    bottom=Side(style='thin', color='E2E8F0')
)

ws.merge_cells("A1:D1")
ws["A1"] = "RELATÓRIO CONSOLIDADO DE FLUXO DE CAIXA (DFC DIRETA)"
ws["A1"].font = FONT_TITLE
ws["A1"].alignment = Alignment(horizontal="left", vertical="center")
ws.row_dimensions[1].height = 30

headers = ["Estrutura de Movimentação de Tesouraria", "Trimestre 1 (R$)", "Trimestre 2 (R$)", "Total Acumulado (R$)"]
for col_idx, h in enumerate(headers, 1):
    c = ws.cell(row=2, column=col_idx, value=h)
    c.fill = HEADER_FILL
    c.font = FONT_HEADER
    c.alignment = Alignment(horizontal="center" if col_idx > 1 else "left", vertical="center")
    c.border = THIN_BORDER
ws.row_dimensions[2].height = 24

fluxo = [
    ["1. SALDO INICIAL DE DISPONIBILIDADES", 1250000, 2410000, 1250000, "subtotal"],
    ["(+) Recebimento de Clientes (Vendas À Vista / Boletos)", 4200000, 4850000, "=B4+C4", "data"],
    ["(+) Outras Receitas Operacionais e Rendimentos", 65000, 72000, "=B5+C5", "data"],
    ["(-) Pagamento a Fornecedores de Insumos", -1850000, -2100000, "=B6+C6", "data"],
    ["(-) Folha de Pagamento, Pró-Labore e Encargos", -980000, -1050000, "=B7+C7", "data"],
    ["(-) Tributos e Impostos Operacionais", -350000, -410000, "=B8+C8", "data"],
    ["(-) Despesas Operacionais Fixas (Aluguel, Luz, TI)", -210000, -225000, "=B9+C9", "data"],
    ["(=) GERAÇÃO OPERACIONAL LÍQUIDA DE CAIXA", "=SUM(B4:B9)", "=SUM(C4:C9)", "=SUM(D4:D9)", "subtotal"],
    ["(-) Investimentos em Equipamentos e TI (CAPEX)", -180000, -95000, "=B11+C11", "data"],
    ["(+) Captação / (-) Amortização de Financiamentos", -420000, -435000, "=B12+C12", "data"],
    ["(=) SALDO FINAL DE CAIXA E EQUIVALENTES", "=B3+B10+B11+B12", "=C3+C10+C11+C12", "=D3+D10+D11+D12", "total"]
]

for row_idx, data in enumerate(fluxo, 3):
    desc, q1, q2, tot, r_type = data
    ws.row_dimensions[row_idx].height = 20
    ws.cell(row=row_idx, column=1, value=desc)
    ws.cell(row=row_idx, column=2, value=q1)
    ws.cell(row=row_idx, column=3, value=q2)
    ws.cell(row=row_idx, column=4, value=tot)

    for col_i in range(1, 5):
        cell = ws.cell(row=row_idx, column=col_i)
        cell.border = THIN_BORDER
        if r_type in ("subtotal", "total"):
            cell.font = FONT_BOLD
            cell.fill = SUBTOTAL_FILL
        else:
            cell.font = FONT_REGULAR
            if row_idx % 2 == 1:
                cell.fill = ZEBRA_FILL
        
        if col_i > 1:
            cell.number_format = 'R$ #,##0;[Red]-R$ #,##0;"-"'
            cell.alignment = Alignment(horizontal="right")

for col in ws.columns:
    max_len = max(len(str(cell.value or '')) for cell in col)
    ws.column_dimensions[get_column_letter(col[0].column)].width = max(max_len + 3, 16)
ws.column_dimensions["A"].width = 46

ws.freeze_panes = "A3"
wb.save("Fluxo_de_Caixa_Gerencial.xlsx")
print("Fluxo de Caixa Gerencial gerado com sucesso: Fluxo_de_Caixa_Gerencial.xlsx")
`,

  contas_receber: `# Gabarito: Controle de Faturamento e Duplicatas a Receber
# Padrão: Gestão de Contas a Receber, Inadimplência e Tesouraria
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

wb = openpyxl.Workbook()
ws = wb.active
ws.title = "Contas_a_Receber"
ws.views.sheetView[0].showGridLines = True

HEADER_FILL = PatternFill(start_color="1E293B", end_color="1E293B", fill_type="solid")
ZEBRA_FILL = PatternFill(start_color="F8FAFC", end_color="F8FAFC", fill_type="solid")

FONT_HEADER = Font(name="Segoe UI", size=10, bold=True, color="FFFFFF")
FONT_TITLE = Font(name="Segoe UI", size=13, bold=True, color="1E293B")
FONT_REGULAR = Font(name="Segoe UI", size=10, color="1E293B")

THIN_BORDER = Border(
    left=Side(style='thin', color='CBD5E1'),
    right=Side(style='thin', color='CBD5E1'),
    top=Side(style='thin', color='CBD5E1'),
    bottom=Side(style='thin', color='CBD5E1')
)

ws.merge_cells("A1:H1")
ws["A1"] = "RELATÓRIO DE FATURAMENTO E CONTAS A RECEBER"
ws["A1"].font = FONT_TITLE
ws["A1"].alignment = Alignment(horizontal="left", vertical="center")
ws.row_dimensions[1].height = 30

headers = ["Duplicata / NF", "Emissão", "Vencimento", "Cliente / Razão Social", "Linha de Negócio", "Valor Bruto (R$)", "Tributos Retidos (R$)", "Valor Líquido (R$)", "Status de Liquidação"]
for col_idx, h in enumerate(headers, 1):
    c = ws.cell(row=2, column=col_idx, value=h)
    c.fill = HEADER_FILL
    c.font = FONT_HEADER
    c.alignment = Alignment(horizontal="center" if col_idx not in (4, 5) else "left", vertical="center")
    c.border = THIN_BORDER
ws.row_dimensions[2].height = 24

duplicatas = [
    ["NF-2026-081", "2026-10-01", "2026-10-31", "Empresa Alpha Ltda", "Consultoria Estratégica", 45000.00, 4250.00, "=F3-G3", "Liquidado"],
    ["NF-2026-082", "2026-10-05", "2026-11-04", "Grupo Beta S/A", "Auditoria de Processos", 128000.00, 12096.00, "=F4-G4", "A Vencer"],
    ["NF-2026-083", "2026-10-12", "2026-11-11", "Comércio Gama Eireli", "Treinamento Gerencial", 28500.00, 2693.25, "=F5-G5", "A Vencer"],
    ["NF-2026-084", "2026-10-15", "2026-11-14", "Indústria Delta Corp", "Reestruturação Financeira", 89000.00, 8410.50, "=F6-G6", "Em Cobrança"],
    ["NF-2026-085", "2026-10-18", "2026-11-17", "Logística Ômega Ltda", "Mapeamento de Custos", 54000.00, 5103.00, "=F7-G7", "Liquidado"],
    ["NF-2026-086", "2026-10-22", "2026-11-21", "Rede Épsilon Varejo", "Consultoria Tributária", 61000.00, 5764.50, "=F8-G8", "A Vencer"],
    ["NF-2026-087", "2026-10-25", "2026-11-24", "Holding Sigma Participações", "Valuation e Due Diligence", 145000.00, 13702.50, "=F9-G9", "A Vencer"]
]

for row_idx, row in enumerate(duplicatas, 3):
    ws.row_dimensions[row_idx].height = 20
    for col_idx, val in enumerate(row, 1):
        cell = ws.cell(row=row_idx, column=col_idx, value=val)
        cell.font = FONT_REGULAR
        cell.border = THIN_BORDER
        if row_idx % 2 == 1:
            cell.fill = ZEBRA_FILL
        
        if 6 <= col_idx <= 8:
            cell.number_format = 'R$ #,##0.00'
            cell.alignment = Alignment(horizontal="right")
        elif col_idx in (1, 2, 3, 9):
            cell.alignment = Alignment(horizontal="center")

for col in ws.columns:
    max_len = max(len(str(c.value or '')) for c in col)
    ws.column_dimensions[get_column_letter(col[0].column)].width = max(max_len + 3, 14)

ws.freeze_panes = "A3"
wb.save("Contas_a_Receber_Consolidado.xlsx")
print("Relatório de Contas a Receber gerado com sucesso: Contas_a_Receber_Consolidado.xlsx")
`,

  genai_gemini: `# Modelo de Automação: Consulta ao Google AI Studio via SDK Python
# Gera estrutura analítica de projetos e orçamentos financeiros

import os
import json
import openpyxl
from google import genai
from google.genai import types

def gerar_planilha():
    api_key = os.environ.get("GEMINI_API_KEY")
    if not api_key:
        print("[AVISO] Chave GEMINI_API_KEY não configurada. Simulando dados analíticos...")
        dados = [
            {"Projeto": "Auditoria de Controladoria", "Responsável": "Carlos Alberto", "Orcamento": 45000, "Realizado": 42000, "Status": "Concluído"},
            {"Projeto": "Implantação de ERP Financeiro", "Responsável": "Beatriz Mendes", "Orcamento": 120000, "Realizado": 115000, "Status": "Em Andamento"},
            {"Projeto": "Reestruturação de Passivos", "Responsável": "Roberto Souza", "Orcamento": 85000, "Realizado": 79000, "Status": "Em Andamento"},
            {"Projeto": "Planejamento Tributário 2027", "Responsável": "Juliana Prado", "Orcamento": 35000, "Realizado": 12000, "Status": "Iniciado"}
        ]
    else:
        client = genai.Client(api_key=api_key)
        prompt = "Gere uma lista estruturada de 4 projetos de consultoria com Projeto, Responsável, Orcamento, Realizado e Status em formato JSON."
        response = client.models.generate_content(
            model="gemini-2.5-flash",
            contents=prompt,
            config=types.GenerateContentConfig(response_mime_type="application/json")
        )
        dados = json.loads(response.text)

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Projetos_Consultoria"

    cabecalhos = list(dados[0].keys())
    ws.append(cabecalhos)
    for linha in dados:
        ws.append(list(linha.values()))

    arquivo_saida = "Relatorio_Projetos_Consultoria.xlsx"
    wb.save(arquivo_saida)
    print(f"Planilha executiva gerada com êxito: {arquivo_saida}")

if __name__ == "__main__":
    gerar_planilha()
`
};

class View {
  constructor() {
    this.codeInput = document.getElementById("code-input");
    this.lineNumbers = document.getElementById("line-numbers");
    this.presetSelect = document.getElementById("preset-select");
    this.pasteBtn = document.getElementById("paste-btn");
    this.clearBtn = document.getElementById("clear-btn");
    this.quickPresetBtn = document.getElementById("quick-preset-btn");

    this.geminiKeyInput = document.getElementById("gemini-key-input");
    this.executeBtn = document.getElementById("execute-btn");
    this.btnExecIcon = document.getElementById("btn-exec-icon");
    this.btnExecText = document.getElementById("btn-exec-text");
    this.execStatusBadge = document.getElementById("exec-status-badge");

    this.tabPreviewBtn = document.getElementById("tab-preview-btn");
    this.tabConsoleBtn = document.getElementById("tab-console-btn");
    this.previewContent = document.getElementById("preview-content");
    this.consoleContent = document.getElementById("console-content");
    this.previewCountPill = document.getElementById("preview-count-pill");
    this.terminalPulse = document.getElementById("terminal-pulse");

    this.sheetMetaBar = document.getElementById("sheet-meta-bar");
    this.previewFilename = document.getElementById("preview-filename");
    this.previewFilesize = document.getElementById("preview-filesize");
    this.sheetTabsContainer = document.getElementById("sheet-tabs-container");
    this.previewFilter = document.getElementById("preview-filter");

    this.previewTable = document.getElementById("preview-table");
    this.previewThead = document.getElementById("preview-thead");
    this.previewTbody = document.getElementById("preview-tbody");
    this.previewEmptyState = document.getElementById("preview-empty-state");

    this.terminalWrapper = document.getElementById("terminal-wrapper");
    this.terminalOutput = document.getElementById("terminal-output");
    this.clearConsoleBtn = document.getElementById("clear-console-btn");

    this.currentFileLabel = document.getElementById("current-file-label");
    this.saveExcelBtn = document.getElementById("save-excel-btn");
    this.openExcelBtn = document.getElementById("open-excel-btn");
    this.openFolderBtn = document.getElementById("open-folder-btn");

    this.envBadge = document.getElementById("env-badge");
    this.envText = document.getElementById("env-text");

    this.themeSystemBtn = document.getElementById("theme-system-btn");
    this.themeLightBtn = document.getElementById("theme-light-btn");
    this.themeDarkBtn = document.getElementById("theme-dark-btn");

    this.toast = document.getElementById("toast");
    this.toastMessage = document.getElementById("toast-message");
    this.toastTimeout = null;

    this.currentSheetData = null;
    this.allRows = [];

    this.initViewEvents();
  }

  initViewEvents() {
    this.codeInput.addEventListener("input", () => this.updateLineNumbers());
    this.codeInput.addEventListener("scroll", () => {
      this.lineNumbers.scrollTop = this.codeInput.scrollTop;
    });

    this.codeInput.addEventListener("keydown", (e) => {
      if (e.key === "Tab") {
        e.preventDefault();
        const start = this.codeInput.selectionStart;
        const end = this.codeInput.selectionEnd;
        this.codeInput.value = this.codeInput.value.substring(0, start) + "    " + this.codeInput.value.substring(end);
        this.codeInput.selectionStart = this.codeInput.selectionEnd = start + 4;
        this.updateLineNumbers();
      }
    });

    this.tabPreviewBtn.addEventListener("click", () => this.switchTab("preview"));
    this.tabConsoleBtn.addEventListener("click", () => this.switchTab("console"));

    this.clearConsoleBtn.addEventListener("click", () => {
      this.terminalOutput.textContent = "";
    });

    this.previewFilter.addEventListener("input", (e) => {
      this.filterTableRows(e.target.value);
    });

    this.themeSystemBtn.addEventListener("click", () => this.setTheme("system"));
    this.themeLightBtn.addEventListener("click", () => this.setTheme("light"));
    this.themeDarkBtn.addEventListener("click", () => this.setTheme("dark"));

    try {
      window.matchMedia("(prefers-color-scheme: dark)").addEventListener("change", () => {
        const currentPref = localStorage.getItem("sheetbuilder_theme");
        if (currentPref === "system" || !currentPref) {
          this.setTheme("system");
        }
      });
    } catch (e) {}

    const initialTheme = localStorage.getItem("sheetbuilder_theme") || "system";
    this.setTheme(initialTheme);

    this.updateLineNumbers();
  }

  updateLineNumbers() {
    const lines = this.codeInput.value.split("\n").length;
    let numbers = "";
    for (let i = 1; i <= Math.max(lines, 1); i++) {
      numbers += i + "\n";
    }
    this.lineNumbers.textContent = numbers;
  }

  switchTab(tabName) {
    if (tabName === "preview") {
      this.tabPreviewBtn.classList.add("active");
      this.tabPreviewBtn.setAttribute("aria-selected", "true");
      this.tabConsoleBtn.classList.remove("active");
      this.tabConsoleBtn.setAttribute("aria-selected", "false");

      this.previewContent.classList.add("active");
      this.consoleContent.classList.remove("active");
    } else {
      this.tabConsoleBtn.classList.add("active");
      this.tabConsoleBtn.setAttribute("aria-selected", "true");
      this.tabPreviewBtn.classList.remove("active");
      this.tabPreviewBtn.setAttribute("aria-selected", "false");

      this.consoleContent.classList.add("active");
      this.previewContent.classList.remove("active");
    }
  }

  setTheme(theme) {
    const root = document.documentElement;
    this.themeSystemBtn.classList.remove("active");
    this.themeLightBtn.classList.remove("active");
    this.themeDarkBtn.classList.remove("active");
    this.themeSystemBtn.setAttribute("aria-pressed", "false");
    this.themeLightBtn.setAttribute("aria-pressed", "false");
    this.themeDarkBtn.setAttribute("aria-pressed", "false");

    if (theme === "light") {
      root.setAttribute("data-theme", "light");
      root.style.colorScheme = "light";
      this.themeLightBtn.classList.add("active");
      this.themeLightBtn.setAttribute("aria-pressed", "true");
      try { localStorage.setItem("sheetbuilder_theme", "light"); } catch (e) {}
    } else if (theme === "dark") {
      root.setAttribute("data-theme", "dark");
      root.style.colorScheme = "dark";
      this.themeDarkBtn.classList.add("active");
      this.themeDarkBtn.setAttribute("aria-pressed", "true");
      try { localStorage.setItem("sheetbuilder_theme", "dark"); } catch (e) {}
    } else {
      root.removeAttribute("data-theme");
      root.style.colorScheme = "light dark";
      this.themeSystemBtn.classList.add("active");
      this.themeSystemBtn.setAttribute("aria-pressed", "true");
      try { localStorage.setItem("sheetbuilder_theme", "system"); } catch (e) {}
    }
  }

  setExecutingState(isExecuting) {
    if (isExecuting) {
      this.executeBtn.disabled = true;
      this.btnExecIcon.innerHTML = `<circle cx="12" cy="12" r="10" stroke="currentColor" stroke-dasharray="32" stroke-dashoffset="16" fill="none"><animateTransform attributeName="transform" type="rotate" from="0 12 12" to="360 12 12" dur="0.8s" repeatCount="indefinite"/></circle>`;
      this.btnExecText.textContent = "Gerando Planilha...";
      this.execStatusBadge.className = "status-badge";
      this.execStatusBadge.textContent = "Processando Modelo...";
      this.execStatusBadge.classList.remove("hidden");
      this.terminalPulse.classList.remove("hidden");
    } else {
      this.executeBtn.disabled = false;
      this.btnExecIcon.innerHTML = `<polygon points="5 3 19 12 5 21 5 3"></polygon>`;
      this.btnExecText.textContent = "Gerar Planilha Executiva";
      this.terminalPulse.classList.add("hidden");
    }
  }

  appendLog(line, isError) {
    const p = document.createElement("div");
    p.textContent = line;
    if (isError) {
      p.className = "term-line-err";
    } else if (line.includes("[SUCESSO]") || line.includes("[SUCCESS]")) {
      p.className = "term-line-succ";
    } else if (line.includes("[INFORMAÇÃO]") || line.includes("[INFO]")) {
      p.className = "term-line-info";
    }
    this.terminalOutput.appendChild(p);
    this.terminalWrapper.scrollTop = this.terminalWrapper.scrollHeight;
  }

  renderPreview(sheetData, onSwitchSheet) {
    this.currentSheetData = sheetData;
    if (!sheetData || !sheetData.headers || sheetData.headers.length === 0) {
      this.previewEmptyState.classList.remove("hidden");
      this.sheetMetaBar.classList.add("hidden");
      this.previewTable.classList.add("hidden");
      this.previewCountPill.classList.add("hidden");
      return;
    }

    this.previewEmptyState.classList.add("hidden");
    this.sheetMetaBar.classList.remove("hidden");
    this.previewTable.classList.remove("hidden");

    this.previewFilename.textContent = sheetData.fileName || "Planilha_Gerada.xlsx";
    this.previewFilesize.textContent = sheetData.fileSizeFormatted || "";
    this.previewCountPill.textContent = `${sheetData.totalRows} linhas formatadas`;
    this.previewCountPill.classList.remove("hidden");

    this.sheetTabsContainer.innerHTML = "";
    if (sheetData.sheetNames && sheetData.sheetNames.length > 1) {
      sheetData.sheetNames.forEach(sName => {
        const chip = document.createElement("button");
        chip.className = `sheet-tab-chip ${sName === sheetData.sheetName ? "active" : ""}`;
        chip.textContent = sName;
        chip.addEventListener("click", () => {
          if (onSwitchSheet) onSwitchSheet(sName);
        });
        this.sheetTabsContainer.appendChild(chip);
      });
    }

    let theadHtml = "<tr>";
    sheetData.headers.forEach(h => {
      theadHtml += `<th scope="col">${this.escapeHtml(h)}</th>`;
    });
    theadHtml += "</tr>";
    this.previewThead.innerHTML = theadHtml;

    this.allRows = sheetData.rows || [];
    this.renderTableRows(this.allRows);

    this.switchTab("preview");
  }

  renderTableRows(rows) {
    let tbodyHtml = "";
    rows.forEach(r => {
      const isTotal = r.some(cell => typeof cell === "string" && (cell.includes("(=)") || cell.toLowerCase().includes("total") || cell.toLowerCase().includes("lucro líquido")));
      tbodyHtml += `<tr class="${isTotal ? "rule-double-total" : ""}">`;
      r.forEach((cell, idx) => {
        const strVal = String(cell || "");
        const isNumeric = idx > 0 && strVal.match(/[0-9]/) && (strVal.includes("R$") || strVal.includes("%") || !isNaN(Number(strVal.replace(/[^0-9.-]/g, ""))));
        const alignStyle = isNumeric ? 'style="text-align: right;"' : '';
        tbodyHtml += `<td ${alignStyle} title="${this.escapeHtml(cell)}">${this.escapeHtml(cell)}</td>`;
      });
      tbodyHtml += "</tr>";
    });
    this.previewTbody.innerHTML = tbodyHtml;
  }

  filterTableRows(query) {
    if (!this.allRows) return;
    const q = query.trim().toLowerCase();
    if (!q) {
      this.renderTableRows(this.allRows);
      return;
    }

    const filtered = this.allRows.filter(row => {
      return row.some(cell => String(cell).toLowerCase().includes(q));
    });
    this.renderTableRows(filtered);
  }

  updateFileActions(filePath) {
    if (filePath) {
      this.currentFileLabel.textContent = filePath;
      this.currentFileLabel.title = filePath;
      this.saveExcelBtn.disabled = false;
      this.openExcelBtn.disabled = false;
      this.openFolderBtn.disabled = false;
    } else {
      this.currentFileLabel.textContent = "Nenhuma pasta de trabalho gerada";
      this.currentFileLabel.title = "";
      this.saveExcelBtn.disabled = true;
      this.openExcelBtn.disabled = true;
      this.openFolderBtn.disabled = true;
    }
  }

  showToast(message, isError = false) {
    if (this.toastTimeout) clearTimeout(this.toastTimeout);
    this.toastMessage.textContent = message;
    this.toast.style.borderColor = isError ? "var(--danger)" : "var(--brand-excel)";
    this.toast.classList.remove("hidden");

    this.toastTimeout = setTimeout(() => {
      this.toast.classList.add("hidden");
    }, 4000);
  }

  escapeHtml(str) {
    if (str === null || str === undefined) return "";
    return String(str)
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;")
      .replace(/"/g, "&quot;")
      .replace(/'/g, "&#039;");
  }
}
