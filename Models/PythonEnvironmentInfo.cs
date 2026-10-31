namespace SheetBuilder.Models;

public class PythonEnvironmentInfo
{
    public bool IsInstalled { get; set; }
    public string ExecutablePath { get; set; } = string.Empty;
    public string VersionString { get; set; } = string.Empty;
    public bool HasOpenPyXL { get; set; }
    public bool HasPandas { get; set; }
    public bool HasXlsxWriter { get; set; }
    public bool HasGoogleGenAI { get; set; }
    public string DefaultDocumentsFolder { get; set; } = string.Empty;
}
