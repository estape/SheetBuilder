# SheetBuilder — Automação de Planilhas Financeiras (Google AI Studio)

Aplicativo desktop moderno para Windows 11 desenvolvido em **.NET 10** e **Photino.NET 4.0**, projetado sob medida para **Gestores Financeiros, Controllers e Consultores de Negócios** executarem modelos em Python gerados no **Google AI Studio** e criarem relatórios gerenciais e demonstrações financeiras em Excel (`.xlsx`) no padrão executivo.

---

## Destaques do Sistema

- **Projetado para Gestão Financeira & Consultoria**: Terminologia corporativa e gabaritos pré-configurados prontos para reuniões de diretoria e conselho:
  - **DRE Gerencial Completa**: Receita Bruta, Deduções, ROL, Margem Bruta, Despesas SG&A, EBITDA/LAJIDA, EBIT, Resultado Financeiro e Lucro Líquido com Análise Vertical (AV %) e variações.
  - **Painel Executivo de Indicadores (KPIs & Valuation)**: Margem EBITDA, Margem Líquida, ROE, ROIC, Liquidez Corrente/Seca, Alavancagem Financeira (Dívida Líquida / EBITDA) e Ciclos Financeiros (PMR, PMP, CCC).
  - **Fluxo de Caixa Direto & Tesouraria**: Geração Operacional, CAPEX e Financiamento.
  - **Controle de Faturamento & Duplicatas a Receber**: Gestão de recebíveis, prazos médios e tributos retidos.
- **Fluxo Simples em 3 Passos**:
  1. No **Google AI Studio** (no navegador), solicite ou personalize a sua planilha financeira.
  2. No SheetBuilder, clique em **Colar Código**.
  3. Clique em **⚡ Gerar Planilha Executiva** para auditar a tabela no Painel de Conferência e salvar na pasta **Documentos** ou abrir no **Microsoft Excel**.
- **Arquitetura MVC Estrita**: Separação clara entre Camada de Apresentação (View), Controladores de Negócio (Controller) e Modelos/Serviços de Execução (Model).
- **Processamento Python Local e Isolado**: Execução segura em sandbox temporário no Windows 11, com registro de atividades em tempo real.
- **Integração Nativa com o Windows 11**:
  - Diálogo de salvamento padrão (`SaveFileDialog`) pré-configurado na pasta **Documentos** do usuário.
  - Abertura imediata no **Microsoft Excel** com 1 clique.
  - Suporte completo a Modo Claro e Modo Escuro acompanhando as preferências do Windows.

---

## Estrutura do Projeto

```
SheetBuilder/
├── SheetBuilder.sln
├── SheetBuilder.csproj             # Projeto WinExe .NET 10 Windows
├── Program.cs                      # Inicialização da janela PhotinoWindow e thread STA
├── Models/
│   ├── AppSettings.cs              # Preferências e caminho padrão da pasta Documentos
│   ├── PythonEnvironmentInfo.cs    # Diagnóstico do motor analítico e bibliotecas
│   ├── ScriptExecutionRequest.cs   # Parâmetros de execução do modelo analítico
│   ├── ScriptExecutionResult.cs    # Resultado e caminhos dos arquivos gerados
│   └── SheetPreviewData.cs         # Estrutura dos dados para o Painel de Conferência
├── Services/
│   ├── PythonRunnerService.cs      # Motor de processamento Python isolado e logs
│   └── ExcelService.cs             # Inspetor ClosedXML e síntese de relatórios
├── Controllers/
│   ├── EnvironmentController.cs    # Diagnósticos do ambiente e gerenciador de pacotes
│   ├── SheetController.cs          # Ciclo de vida da geração e conferência de planilhas
│   ├── DialogController.cs         # Diálogo Salvar Como e acionador do Microsoft Excel
│   └── MessageControllerRouter.cs  # Roteador assíncrono de comunicação IPC
├── wwwroot/                        # Camada de Apresentação (View)
│   ├── index.html                  # Interface executiva semântica e acessível (pt-BR)
│   ├── css/
│   │   └── app.css                 # Estilos modernos (light-dark(), tabelas responsivas)
│   └── js/
│       ├── view.js                 # Gabaritos de consultoria (DRE, KPIs, Fluxo) e UI
│       └── app.js                  # Orquestrador de requisições e ações do usuário
└── tests/
    └── SheetBuilder.Tests/         # Testes automatizados xUnit (5 testes aprovados)
```

---

## Como Executar

1. Abra o terminal na pasta do projeto:
   ```powershell
   dotnet run
   ```

2. Selecione um gabarito financeiro no menu suspenso ou cole o código gerado no Google AI Studio.

3. Clique em **"Gerar Planilha Executiva"**.

4. Audite os dados no **Painel de Conferência**.

5. Clique em **"Salvar Pasta de Trabalho..."** (salva diretamente em Documentos) ou **"Abrir no Microsoft Excel"**.

---

## Execução dos Testes Automatizados

```powershell
dotnet test
```
