namespace SheetBuilder.Models;

public class SheetPreviewData
{
    public string SheetName { get; set; } = string.Empty;
    public List<string> SheetNames { get; set; } = new();
    public List<string> Headers { get; set; } = new();
    public List<List<string>> Rows { get; set; } = new();
    public int TotalRows { get; set; }
    public int TotalColumns { get; set; }
    public string FileSizeFormatted { get; set; } = string.Empty;
    public string FilePath { get; set; } = string.Empty;
    public string FileName { get; set; } = string.Empty;
}
