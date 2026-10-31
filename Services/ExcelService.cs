using System.Text.Json;
using ClosedXML.Excel;
using SheetBuilder.Models;

namespace SheetBuilder.Services;

public class ExcelService
{
    public SheetPreviewData InspectExcelFile(string filePath, string? targetSheetName = null, int maxPreviewRows = 250)
    {
        if (!File.Exists(filePath))
            throw new FileNotFoundException("Excel file not found", filePath);

        var preview = new SheetPreviewData
        {
            FilePath = filePath,
            FileName = Path.GetFileName(filePath)
        };

        var fileInfo = new FileInfo(filePath);
        preview.FileSizeFormatted = FormatBytes(fileInfo.Length);

        using var workbook = new XLWorkbook(filePath);
        foreach (var ws in workbook.Worksheets)
        {
            preview.SheetNames.Add(ws.Name);
        }

        var worksheet = string.IsNullOrWhiteSpace(targetSheetName)
            ? workbook.Worksheets.FirstOrDefault()
            : workbook.Worksheet(targetSheetName);

        if (worksheet == null)
        {
            preview.SheetName = "Empty";
            return preview;
        }

        preview.SheetName = worksheet.Name;
        var range = worksheet.RangeUsed();
        if (range == null)
        {
            return preview;
        }

        var firstRow = range.FirstRow();
        var lastRow = range.LastRow();
        var firstCol = range.FirstColumn();
        var lastCol = range.LastColumn();

        int totalRows = lastRow.RowNumber() - firstRow.RowNumber() + 1;
        int totalCols = lastCol.ColumnNumber() - firstCol.ColumnNumber() + 1;

        preview.TotalRows = Math.Max(0, totalRows - 1); // Excluding header
        preview.TotalColumns = totalCols;

        // Read headers
        for (int c = firstCol.ColumnNumber(); c <= lastCol.ColumnNumber(); c++)
        {
            var headerVal = worksheet.Cell(firstRow.RowNumber(), c).GetFormattedString();
            if (string.IsNullOrWhiteSpace(headerVal))
                headerVal = $"Column {c}";
            preview.Headers.Add(headerVal);
        }

        // Read data rows
        int startDataRow = firstRow.RowNumber() + 1;
        int endDataRow = Math.Min(lastRow.RowNumber(), startDataRow + maxPreviewRows - 1);

        for (int r = startDataRow; r <= endDataRow; r++)
        {
            var rowList = new List<string>();
            for (int c = firstCol.ColumnNumber(); c <= lastCol.ColumnNumber(); c++)
            {
                var val = worksheet.Cell(r, c).GetFormattedString();
                rowList.Add(val);
            }
            preview.Rows.Add(rowList);
        }

        return preview;
    }

    public string? SynthesizeExcelFromOutput(List<string> outputLines, string workDir)
    {
        var rawText = string.Join(Environment.NewLine, outputLines).Trim();
        if (string.IsNullOrWhiteSpace(rawText)) return null;

        // Try extracting JSON
        var jsonText = ExtractJson(rawText);
        if (!string.IsNullOrWhiteSpace(jsonText))
        {
            try
            {
                using var doc = JsonDocument.Parse(jsonText);
                var root = doc.RootElement;
                JsonElement rowsArray = default;

                if (root.ValueKind == JsonValueKind.Array)
                {
                    rowsArray = root;
                }
                else if (root.ValueKind == JsonValueKind.Object)
                {
                    foreach (var prop in root.EnumerateObject())
                    {
                        if (prop.Value.ValueKind == JsonValueKind.Array)
                        {
                            rowsArray = prop.Value;
                            break;
                        }
                    }
                }

                if (rowsArray.ValueKind == JsonValueKind.Array && rowsArray.GetArrayLength() > 0)
                {
                    var outputPath = Path.Combine(workDir, $"Generated_Sheet_{DateTime.Now:yyyyMMdd_HHmmss}.xlsx");
                    BuildExcelFromJsonArray(rowsArray, outputPath);
                    return outputPath;
                }
            }
            catch
            {
                // Not valid JSON array, fallback to CSV parsing
            }
        }

        // Try CSV / TSV parsing
        if (rawText.Contains(',') || rawText.Contains('\t') || rawText.Contains('|'))
        {
            try
            {
                var outputPath = Path.Combine(workDir, $"Generated_Sheet_{DateTime.Now:yyyyMMdd_HHmmss}.xlsx");
                if (BuildExcelFromDelimitedText(rawText, outputPath))
                    return outputPath;
            }
            catch { }
        }

        return null;
    }

    private static void BuildExcelFromJsonArray(JsonElement rowsArray, string outputPath)
    {
        using var workbook = new XLWorkbook();
        var worksheet = workbook.Worksheets.Add("Dados_Extraidos");

        // Collect distinct headers
        var headers = new List<string>();
        foreach (var item in rowsArray.EnumerateArray())
        {
            if (item.ValueKind == JsonValueKind.Object)
            {
                foreach (var prop in item.EnumerateObject())
                {
                    if (!headers.Contains(prop.Name))
                        headers.Add(prop.Name);
                }
            }
        }

        if (headers.Count == 0) return;

        // Write Header Row
        for (int col = 0; col < headers.Count; col++)
        {
            var cell = worksheet.Cell(1, col + 1);
            cell.Value = headers[col];
            cell.Style.Font.Bold = true;
            cell.Style.Fill.BackgroundColor = XLColor.FromHtml("#1E3A8A"); // Navy Blue
            cell.Style.Font.FontColor = XLColor.White;
            cell.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
        }

        // Write Data Rows
        int currentRow = 2;
        foreach (var item in rowsArray.EnumerateArray())
        {
            for (int col = 0; col < headers.Count; col++)
            {
                var cell = worksheet.Cell(currentRow, col + 1);
                var headerName = headers[col];

                if (item.TryGetProperty(headerName, out var propVal))
                {
                    switch (propVal.ValueKind)
                    {
                        case JsonValueKind.Number:
                            if (propVal.TryGetInt64(out var intVal))
                                cell.Value = intVal;
                            else if (propVal.TryGetDouble(out var dblVal))
                                cell.Value = dblVal;
                            break;
                        case JsonValueKind.True:
                        case JsonValueKind.False:
                            cell.Value = propVal.GetBoolean();
                            break;
                        case JsonValueKind.String:
                            cell.Value = propVal.GetString();
                            break;
                        case JsonValueKind.Null:
                            cell.Value = "";
                            break;
                        default:
                            cell.Value = propVal.ToString();
                            break;
                    }
                }

                // Zebra striping
                if (currentRow % 2 == 1)
                {
                    cell.Style.Fill.BackgroundColor = XLColor.FromHtml("#F8FAFC");
                }
            }
            currentRow++;
        }

        if (currentRow > 2)
        {
            var tableRange = worksheet.Range(1, 1, currentRow - 1, headers.Count);
            tableRange.CreateTable();
        }

        worksheet.Columns().AdjustToContents();
        worksheet.SheetView.FreezeRows(1);

        workbook.SaveAs(outputPath);
    }

    private static bool BuildExcelFromDelimitedText(string text, string outputPath)
    {
        var rawLines = text.Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries)
            .Select(l => l.Trim())
            .Where(l => !string.IsNullOrWhiteSpace(l))
            .ToList();

        if (rawLines.Count < 2) return false;

        // Check if markdown table
        char delimiter = ',';
        bool isMarkdown = false;

        if (rawLines[0].StartsWith("|") && rawLines[0].EndsWith("|"))
        {
            isMarkdown = true;
            delimiter = '|';
        }
        else if (rawLines[0].Contains('\t'))
        {
            delimiter = '\t';
        }

        var parsedRows = new List<List<string>>();
        foreach (var line in rawLines)
        {
            if (isMarkdown && line.Contains("---")) continue; // Skip markdown separator row

            var parts = isMarkdown
                ? line.Trim('|').Split('|').Select(p => p.Trim()).ToList()
                : line.Split(delimiter).Select(p => p.Trim(' ', '"')).ToList();

            if (parts.Count > 0)
                parsedRows.Add(parts);
        }

        if (parsedRows.Count < 2) return false;

        using var workbook = new XLWorkbook();
        var worksheet = workbook.Worksheets.Add("Tabela_Consolidada");

        var headers = parsedRows[0];
        for (int col = 0; col < headers.Count; col++)
        {
            var cell = worksheet.Cell(1, col + 1);
            cell.Value = headers[col];
            cell.Style.Font.Bold = true;
            cell.Style.Fill.BackgroundColor = XLColor.FromHtml("#1E3A8A");
            cell.Style.Font.FontColor = XLColor.White;
        }

        for (int r = 1; r < parsedRows.Count; r++)
        {
            var rowData = parsedRows[r];
            int excelRow = r + 1;
            for (int col = 0; col < headers.Count; col++)
            {
                var val = col < rowData.Count ? rowData[col] : "";
                var cell = worksheet.Cell(excelRow, col + 1);

                if (double.TryParse(val, out var num))
                    cell.Value = num;
                else if (bool.TryParse(val, out var b))
                    cell.Value = b;
                else
                    cell.Value = val;

                if (excelRow % 2 == 1)
                    cell.Style.Fill.BackgroundColor = XLColor.FromHtml("#F8FAFC");
            }
        }

        worksheet.Columns().AdjustToContents();
        worksheet.SheetView.FreezeRows(1);
        workbook.SaveAs(outputPath);
        return true;
    }

    private static string? ExtractJson(string text)
    {
        // Strip markdown fences
        var cleaned = text;
        int codeStart = cleaned.IndexOf("```json", StringComparison.OrdinalIgnoreCase);
        if (codeStart >= 0)
        {
            int contentStart = codeStart + 7;
            int codeEnd = cleaned.IndexOf("```", contentStart, StringComparison.Ordinal);
            if (codeEnd > contentStart)
            {
                return cleaned.Substring(contentStart, codeEnd - contentStart).Trim();
            }
        }

        int bracketStart = cleaned.IndexOf('[');
        int braceStart = cleaned.IndexOf('{');

        int start = -1;
        if (bracketStart >= 0 && braceStart >= 0)
            start = Math.Min(bracketStart, braceStart);
        else if (bracketStart >= 0)
            start = bracketStart;
        else if (braceStart >= 0)
            start = braceStart;

        if (start < 0) return null;

        int bracketEnd = cleaned.LastIndexOf(']');
        int braceEnd = cleaned.LastIndexOf('}');
        int end = Math.Max(bracketEnd, braceEnd);

        if (end > start)
        {
            return cleaned.Substring(start, end - start + 1).Trim();
        }

        return null;
    }

    private static string FormatBytes(long bytes)
    {
        string[] suffixes = { "B", "KB", "MB", "GB" };
        int i = 0;
        double d = bytes;
        while (d >= 1024 && i < suffixes.Length - 1)
        {
            d /= 1024;
            i++;
        }
        return $"{d:0.##} {suffixes[i]}";
    }
}
