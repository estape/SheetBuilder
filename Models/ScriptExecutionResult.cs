namespace SheetBuilder.Models;

public class ScriptExecutionResult
{
    public bool IsSuccess { get; set; }
    public int ExitCode { get; set; }
    public long ExecutionTimeMs { get; set; }
    public List<string> OutputLines { get; set; } = new();
    public List<string> ErrorLines { get; set; } = new();
    public string? GeneratedExcelPath { get; set; }
    public SheetPreviewData? SheetData { get; set; }
    public string? ErrorMessage { get; set; }
}
