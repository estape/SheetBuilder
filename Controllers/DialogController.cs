using System.Diagnostics;
using SheetBuilder.Models;

namespace SheetBuilder.Controllers;

public class DialogController
{
    public Task<string?> SaveExcelFileDialogAsync(string currentFilePath)
    {
        var tcs = new TaskCompletionSource<string?>();

        var thread = new Thread(() =>
        {
            try
            {
                using var dialog = new SaveFileDialog
                {
                    Title = "Salvar Pasta de Trabalho do Excel Como",
                    InitialDirectory = AppSettings.DefaultSaveFolder,
                    FileName = Path.GetFileName(currentFilePath),
                    Filter = "Pasta de Trabalho do Excel (*.xlsx)|*.xlsx|Todos os Arquivos (*.*)|*.*",
                    DefaultExt = "xlsx",
                    AddExtension = true,
                    OverwritePrompt = true
                };

                if (dialog.ShowDialog() == DialogResult.OK)
                {
                    var dest = dialog.FileName;
                    File.Copy(currentFilePath, dest, true);
                    tcs.SetResult(dest);
                }
                else
                {
                    tcs.SetResult(null); // User cancelled
                }
            }
            catch (Exception ex)
            {
                tcs.SetException(ex);
            }
        });

        thread.SetApartmentState(ApartmentState.STA);
        thread.Start();

        return tcs.Task;
    }

    public bool OpenInExcel(string filePath)
    {
        if (!File.Exists(filePath)) return false;

        try
        {
            var psi = new ProcessStartInfo
            {
                FileName = filePath,
                UseShellExecute = true
            };
            Process.Start(psi);
            return true;
        }
        catch
        {
            return false;
        }
    }

    public bool OpenContainingFolder(string filePath)
    {
        if (!File.Exists(filePath)) return false;

        try
        {
            var psi = new ProcessStartInfo
            {
                FileName = "explorer.exe",
                Arguments = $"/select,\"{filePath}\"",
                UseShellExecute = true
            };
            Process.Start(psi);
            return true;
        }
        catch
        {
            return false;
        }
    }
}
