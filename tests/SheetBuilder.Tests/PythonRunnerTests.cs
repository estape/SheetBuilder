using SheetBuilder.Models;
using SheetBuilder.Services;

namespace SheetBuilder.Tests;

public class PythonRunnerTests
{
    [Fact]
    public async Task GetEnvironmentInfoAsync_DetectsPythonSuccessfully()
    {
        var service = new PythonRunnerService();
        var env = await service.GetEnvironmentInfoAsync();

        Assert.True(env.IsInstalled, "Python should be detected on the machine.");
        Assert.False(string.IsNullOrWhiteSpace(env.VersionString));
        Assert.True(env.HasOpenPyXL, "openpyxl should be detected.");
    }

    [Fact]
    public async Task ExecuteScriptAsync_CreatesXlsxAndCapturesIt()
    {
        var service = new PythonRunnerService();
        var tempDir = Path.Combine(Path.GetTempPath(), "SheetBuilder_PyTest_" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(tempDir);

        var pythonCode = @"
import openpyxl

wb = openpyxl.Workbook()
ws = wb.active
ws.title = 'TestSheet'
ws.append(['City', 'Population', 'Country'])
ws.append(['Tokyo', 37400000, 'Japan'])
ws.append(['Delhi', 30290000, 'India'])
wb.save('cities.xlsx')
print('Script finished successfully')
";

        var req = new ScriptExecutionRequest
        {
            ScriptCode = pythonCode,
            WorkingDirectory = tempDir,
            TimeoutSeconds = 30
        };

        var logs = new List<string>();

        try
        {
            var result = await service.ExecuteScriptAsync(req, (line, isErr) => logs.Add(line));

            Assert.True(result.IsSuccess);
            Assert.Equal(0, result.ExitCode);
            Assert.NotNull(result.GeneratedExcelPath);
            Assert.True(File.Exists(result.GeneratedExcelPath));
            Assert.EndsWith("cities.xlsx", result.GeneratedExcelPath, StringComparison.OrdinalIgnoreCase);

            var excelService = new ExcelService();
            var preview = excelService.InspectExcelFile(result.GeneratedExcelPath);
            Assert.Equal("TestSheet", preview.SheetName);
            Assert.Equal(3, preview.TotalColumns);
            Assert.Equal(2, preview.TotalRows);
            Assert.Equal("Tokyo", preview.Rows[0][0]);
        }
        finally
        {
            if (Directory.Exists(tempDir))
                Directory.Delete(tempDir, true);
        }
    }
}
