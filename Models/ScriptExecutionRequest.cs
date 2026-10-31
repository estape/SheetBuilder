namespace SheetBuilder.Models;

public class ScriptExecutionRequest
{
    public string ScriptCode { get; set; } = string.Empty;
    public string? WorkingDirectory { get; set; }
    public int TimeoutSeconds { get; set; } = 90;
    public Dictionary<string, string> EnvironmentVariables { get; set; } = new();
}
