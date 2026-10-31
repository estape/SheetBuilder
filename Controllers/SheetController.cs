using SheetBuilder.Models;
using SheetBuilder.Services;

namespace SheetBuilder.Controllers;

public class SheetController
{
    private readonly PythonRunnerService _pythonRunnerService;
    private readonly ExcelService _excelService;

    // Cache the most recently generated or inspected excel path
    public string? CurrentExcelFilePath { get; private set; }

    public SheetController(PythonRunnerService pythonRunnerService, ExcelService excelService)
    {
        _pythonRunnerService = pythonRunnerService;
        _excelService = excelService;
    }

    public async Task<ScriptExecutionResult> ExecuteScriptAsync(
        ScriptExecutionRequest request, 
        Action<string, bool> onLogReceived)
    {
        var result = await _pythonRunnerService.ExecuteScriptAsync(request, onLogReceived);

        string? excelPath = result.GeneratedExcelPath;

        // If no file was directly generated on disk by the Python script,
        // check if the script output contains JSON or tabular data
        if (string.IsNullOrWhiteSpace(excelPath))
        {
            var workDir = request.WorkingDirectory ?? Path.Combine(Path.GetTempPath(), "SheetBuilder");
            excelPath = _excelService.SynthesizeExcelFromOutput(result.OutputLines, workDir);
            if (!string.IsNullOrWhiteSpace(excelPath))
            {
                result.GeneratedExcelPath = excelPath;
                onLogReceived($"[INFORMAÇÃO] Planilha executiva estruturada a partir dos dados do modelo: {Path.GetFileName(excelPath)}", false);
            }
        }

        if (!string.IsNullOrWhiteSpace(excelPath) && File.Exists(excelPath))
        {
            CurrentExcelFilePath = excelPath;
            try
            {
                result.SheetData = _excelService.InspectExcelFile(excelPath);
                onLogReceived($"[INFORMAÇÃO] Planilha carregada com sucesso para conferência: '{result.SheetData.SheetName}' ({result.SheetData.TotalRows} linhas de dados, {result.SheetData.TotalColumns} colunas)", false);
            }
            catch (Exception ex)
            {
                onLogReceived($"[ALERTA] Não foi possível estruturar a pré-visualização da planilha: {ex.Message}", true);
            }
        }
        else if (result.IsSuccess)
        {
            onLogReceived("[AVISO] O modelo concluiu a execução, mas nenhuma pasta de trabalho (.xlsx) foi gerada no diretório.", true);
        }

        return result;
    }

    public SheetPreviewData? GetSheetPreview(string filePath, string? sheetName = null)
    {
        if (!File.Exists(filePath)) return null;
        CurrentExcelFilePath = filePath;
        return _excelService.InspectExcelFile(filePath, sheetName);
    }
}
