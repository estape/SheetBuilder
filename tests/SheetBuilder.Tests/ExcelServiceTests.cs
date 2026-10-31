using SheetBuilder.Services;
using SheetBuilder.Models;

namespace SheetBuilder.Tests;

public class ExcelServiceTests
{
    [Fact]
    public void SynthesizeExcelFromOutput_WithJsonArray_GeneratesValidExcelAndInspectsIt()
    {
        // Arrange
        var service = new ExcelService();
        var tempDir = Path.Combine(Path.GetTempPath(), "SheetBuilder_Tests_" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(tempDir);

        var outputLines = new List<string>
        {
            "```json",
            "[",
            "  {\"Product\": \"Laptop Pro\", \"Category\": \"Hardware\", \"Price\": 1299.99, \"InStock\": true},",
            "  {\"Product\": \"Wireless Mouse\", \"Category\": \"Accessories\", \"Price\": 49.50, \"InStock\": false}",
            "]",
            "```"
        };

        try
        {
            // Act
            var excelPath = service.SynthesizeExcelFromOutput(outputLines, tempDir);

            // Assert
            Assert.NotNull(excelPath);
            Assert.True(File.Exists(excelPath));

            // Inspect
            var preview = service.InspectExcelFile(excelPath);
            Assert.NotNull(preview);
            Assert.Equal(4, preview.TotalColumns);
            Assert.Equal(2, preview.TotalRows);
            Assert.Contains("Product", preview.Headers);
            Assert.Contains("Price", preview.Headers);
            Assert.Equal(2, preview.Rows.Count);
            Assert.Equal("Laptop Pro", preview.Rows[0][0]);
        }
        finally
        {
            if (Directory.Exists(tempDir))
                Directory.Delete(tempDir, true);
        }
    }

    [Fact]
    public void SynthesizeExcelFromOutput_WithMarkdownTable_GeneratesValidExcel()
    {
        // Arrange
        var service = new ExcelService();
        var tempDir = Path.Combine(Path.GetTempPath(), "SheetBuilder_Tests_" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(tempDir);

        var outputLines = new List<string>
        {
            "| Employee ID | Name | Department |",
            "|---|---|---|",
            "| EMP-001 | Alice Johnson | Engineering |",
            "| EMP-002 | Bob Miller | Marketing |"
        };

        try
        {
            // Act
            var excelPath = service.SynthesizeExcelFromOutput(outputLines, tempDir);

            // Assert
            Assert.NotNull(excelPath);
            Assert.True(File.Exists(excelPath));

            var preview = service.InspectExcelFile(excelPath);
            Assert.Equal(3, preview.TotalColumns);
            Assert.Equal(2, preview.TotalRows);
            Assert.Equal("Alice Johnson", preview.Rows[0][1]);
        }
        finally
        {
            if (Directory.Exists(tempDir))
                Directory.Delete(tempDir, true);
        }
    }

    [Fact]
    public void SynthesizeExcelFromOutput_WithBrazilianFinancialTable_InspectsSuccessfully()
    {
        // Arrange
        var service = new ExcelService();
        var tempDir = Path.Combine(Path.GetTempPath(), "SheetBuilder_Tests_" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(tempDir);

        var outputLines = new List<string>
        {
            "| Estrutura de Contas | Orçado (R$) | Realizado (R$) | AV (%) |",
            "|---|---|---|---|",
            "| 1. Receita Bruta | 16800000 | 17250000 | 117.7% |",
            "| 2. (-) Deduções e Impostos | -2520000 | -2587500 | -17.7% |",
            "| (=) EBITDA / LAJIDA | 3192000 | 3270000 | 22.3% |"
        };

        try
        {
            // Act
            var excelPath = service.SynthesizeExcelFromOutput(outputLines, tempDir);

            // Assert
            Assert.NotNull(excelPath);
            Assert.True(File.Exists(excelPath));

            var preview = service.InspectExcelFile(excelPath);
            Assert.Equal(4, preview.TotalColumns);
            Assert.Equal(3, preview.TotalRows);
            Assert.Contains("Estrutura de Contas", preview.Headers);
            Assert.Contains("Realizado (R$)", preview.Headers);
            Assert.Equal("1. Receita Bruta", preview.Rows[0][0]);
        }
        finally
        {
            if (Directory.Exists(tempDir))
                Directory.Delete(tempDir, true);
        }
    }
}
