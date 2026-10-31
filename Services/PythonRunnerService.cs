using System.Diagnostics;
using System.Text;
using SheetBuilder.Models;

namespace SheetBuilder.Services;

public class PythonRunnerService
{
    private string _pythonExecutable = "python";

    public PythonRunnerService()
    {
        _pythonExecutable = DetectPythonExecutable();
    }

    public string PythonExecutable => _pythonExecutable;

    public static string DetectPythonExecutable()
    {
        // Try 'python', 'py', or full path if found in typical Windows locations
        string[] candidates = { "python", "py" };
        foreach (var cmd in candidates)
        {
            try
            {
                var psi = new ProcessStartInfo
                {
                    FileName = cmd,
                    Arguments = "--version",
                    RedirectStandardOutput = true,
                    RedirectStandardError = true,
                    UseShellExecute = false,
                    CreateNoWindow = true
                };
                using var p = Process.Start(psi);
                if (p != null)
                {
                    p.WaitForExit(3000);
                    if (p.ExitCode == 0) return cmd;
                }
            }
            catch
            {
                // Continue searching
            }
        }
        return "python";
    }

    public async Task<PythonEnvironmentInfo> GetEnvironmentInfoAsync()
    {
        var info = new PythonEnvironmentInfo
        {
            DefaultDocumentsFolder = AppSettings.DefaultSaveFolder
        };

        try
        {
            var psi = new ProcessStartInfo
            {
                FileName = _pythonExecutable,
                Arguments = "-c \"import sys; print(sys.version); print(sys.executable)\"",
                RedirectStandardOutput = true,
                RedirectStandardError = true,
                UseShellExecute = false,
                CreateNoWindow = true
            };

            using var process = Process.Start(psi);
            if (process != null)
            {
                var stdout = await process.StandardOutput.ReadToEndAsync();
                await process.WaitForExitAsync();

                if (process.ExitCode == 0)
                {
                    info.IsInstalled = true;
                    var lines = stdout.Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries);
                    if (lines.Length > 0) info.VersionString = lines[0].Trim();
                    if (lines.Length > 1) info.ExecutablePath = lines[1].Trim();
                }
            }
        }
        catch
        {
            info.IsInstalled = false;
        }

        if (info.IsInstalled)
        {
            // Check packages
            try
            {
                var pkgPsi = new ProcessStartInfo
                {
                    FileName = _pythonExecutable,
                    Arguments = "-c \"import openpyxl; print('openpyxl:ok')\"",
                    RedirectStandardOutput = true,
                    UseShellExecute = false,
                    CreateNoWindow = true
                };
                using var p = Process.Start(pkgPsi);
                if (p != null)
                {
                    var outText = await p.StandardOutput.ReadToEndAsync();
                    await p.WaitForExitAsync();
                    info.HasOpenPyXL = outText.Contains("openpyxl:ok");
                }
            }
            catch { }

            try
            {
                var pkgPsi = new ProcessStartInfo
                {
                    FileName = _pythonExecutable,
                    Arguments = "-c \"import pandas; print('pandas:ok')\"",
                    RedirectStandardOutput = true,
                    UseShellExecute = false,
                    CreateNoWindow = true
                };
                using var p = Process.Start(pkgPsi);
                if (p != null)
                {
                    var outText = await p.StandardOutput.ReadToEndAsync();
                    await p.WaitForExitAsync();
                    info.HasPandas = outText.Contains("pandas:ok");
                }
            }
            catch { }

            try
            {
                var pkgPsi = new ProcessStartInfo
                {
                    FileName = _pythonExecutable,
                    Arguments = "-c \"import xlsxwriter; print('xlsxwriter:ok')\"",
                    RedirectStandardOutput = true,
                    UseShellExecute = false,
                    CreateNoWindow = true
                };
                using var p = Process.Start(pkgPsi);
                if (p != null)
                {
                    var outText = await p.StandardOutput.ReadToEndAsync();
                    await p.WaitForExitAsync();
                    info.HasXlsxWriter = outText.Contains("xlsxwriter:ok");
                }
            }
            catch { }

            try
            {
                var pkgPsi = new ProcessStartInfo
                {
                    FileName = _pythonExecutable,
                    Arguments = "-c \"import google.genai; print('genai:ok')\"",
                    RedirectStandardOutput = true,
                    UseShellExecute = false,
                    CreateNoWindow = true
                };
                using var p = Process.Start(pkgPsi);
                if (p != null)
                {
                    var outText = await p.StandardOutput.ReadToEndAsync();
                    await p.WaitForExitAsync();
                    info.HasGoogleGenAI = outText.Contains("genai:ok");
                }
            }
            catch { }
        }

        return info;
    }

    public async Task<bool> InstallPackageAsync(string packageName, Action<string> onLog)
    {
        try
        {
            var psi = new ProcessStartInfo
            {
                FileName = _pythonExecutable,
                Arguments = $"-m pip install {packageName}",
                RedirectStandardOutput = true,
                RedirectStandardError = true,
                UseShellExecute = false,
                CreateNoWindow = true
            };

            using var process = Process.Start(psi);
            if (process == null) return false;

            process.OutputDataReceived += (_, e) =>
            {
                if (e.Data != null) onLog(e.Data);
            };
            process.ErrorDataReceived += (_, e) =>
            {
                if (e.Data != null) onLog($"[ERR] {e.Data}");
            };

            process.BeginOutputReadLine();
            process.BeginErrorReadLine();

            await process.WaitForExitAsync();
            return process.ExitCode == 0;
        }
        catch (Exception ex)
        {
            onLog($"Failed to install {packageName}: {ex.Message}");
            return false;
        }
    }

    public async Task<ScriptExecutionResult> ExecuteScriptAsync(
        ScriptExecutionRequest request, 
        Action<string, bool> onLogReceived)
    {
        var sw = Stopwatch.StartNew();
        var result = new ScriptExecutionResult();

        // 1. Prepare isolated workspace directory
        var workDir = request.WorkingDirectory;
        if (string.IsNullOrWhiteSpace(workDir) || !Directory.Exists(workDir))
        {
            workDir = Path.Combine(Path.GetTempPath(), "SheetBuilder", $"run_{Guid.NewGuid():N}");
            Directory.CreateDirectory(workDir);
        }

        var scriptFile = Path.Combine(workDir, "generate_sheets.py");
        await File.WriteAllTextAsync(scriptFile, request.ScriptCode, Encoding.UTF8);

        // Record initial files in workDir
        var initialFiles = Directory.GetFiles(workDir, "*.*", SearchOption.AllDirectories)
            .ToHashSet(StringComparer.OrdinalIgnoreCase);

        onLogReceived($"[INFORMAÇÃO] Diretório de trabalho seguro: {workDir}", false);
        onLogReceived($"[INFORMAÇÃO] Inicializando motor analítico: {_pythonExecutable} generate_sheets.py", false);

        var outputLines = new List<string>();
        var errorLines = new List<string>();

        try
        {
            var psi = new ProcessStartInfo
            {
                FileName = _pythonExecutable,
                Arguments = $"\"{scriptFile}\"",
                WorkingDirectory = workDir,
                RedirectStandardOutput = true,
                RedirectStandardError = true,
                UseShellExecute = false,
                CreateNoWindow = true,
                StandardOutputEncoding = Encoding.UTF8,
                StandardErrorEncoding = Encoding.UTF8
            };

            // Pass environment variables (e.g. GEMINI_API_KEY if user provided one)
            foreach (var kv in request.EnvironmentVariables)
            {
                psi.EnvironmentVariables[kv.Key] = kv.Value;
            }

            using var process = new Process { StartInfo = psi };

            process.OutputDataReceived += (_, e) =>
            {
                if (e.Data != null)
                {
                    outputLines.Add(e.Data);
                    onLogReceived(e.Data, false);
                }
            };

            process.ErrorDataReceived += (_, e) =>
            {
                if (e.Data != null)
                {
                    errorLines.Add(e.Data);
                    onLogReceived(e.Data, true);
                }
            };

            process.Start();
            process.BeginOutputReadLine();
            process.BeginErrorReadLine();

            using var cts = new CancellationTokenSource(TimeSpan.FromSeconds(request.TimeoutSeconds));
            await process.WaitForExitAsync(cts.Token);

            sw.Stop();
            result.ExitCode = process.ExitCode;
            result.ExecutionTimeMs = sw.ElapsedMilliseconds;
            result.OutputLines = outputLines;
            result.ErrorLines = errorLines;
            result.IsSuccess = process.ExitCode == 0;

            if (result.IsSuccess)
            {
                onLogReceived($"[SUCESSO] Processamento concluído com êxito em {sw.ElapsedMilliseconds}ms (Validação Aprovada)", false);
            }
            else
            {
                onLogReceived($"[ALERTA] Processamento finalizado com ocorrências (Código: {process.ExitCode})", true);
                result.ErrorMessage = string.Join(Environment.NewLine, errorLines);
            }

            // 2. Scan for newly created .xlsx / .xls files in the work directory
            var currentFiles = Directory.GetFiles(workDir, "*.*", SearchOption.AllDirectories);
            var newExcelFiles = currentFiles
                .Where(f => !initialFiles.Contains(f) && 
                            (f.EndsWith(".xlsx", StringComparison.OrdinalIgnoreCase) || 
                             f.EndsWith(".xls", StringComparison.OrdinalIgnoreCase)))
                .ToList();

            if (newExcelFiles.Count > 0)
            {
                // Take the newest generated excel file
                var generated = newExcelFiles.OrderByDescending(File.GetLastWriteTimeUtc).First();
                result.GeneratedExcelPath = generated;
                onLogReceived($"[INFORMAÇÃO] Pasta de trabalho Excel gerada detectada: {Path.GetFileName(generated)}", false);
            }
            else
            {
                // Check if any existing excel file was updated
                var anyExcel = currentFiles
                    .Where(f => f.EndsWith(".xlsx", StringComparison.OrdinalIgnoreCase) || 
                                f.EndsWith(".xls", StringComparison.OrdinalIgnoreCase))
                    .OrderByDescending(File.GetLastWriteTimeUtc)
                    .FirstOrDefault();

                if (anyExcel != null)
                {
                    result.GeneratedExcelPath = anyExcel;
                    onLogReceived($"[INFORMAÇÃO] Pasta de trabalho identificada no diretório: {Path.GetFileName(anyExcel)}", false);
                }
            }

            return result;
        }
        catch (OperationCanceledException)
        {
            sw.Stop();
            result.IsSuccess = false;
            result.ErrorMessage = $"O tempo limite de processamento de {request.TimeoutSeconds} segundos foi excedido.";
            onLogReceived($"[ERRO] {result.ErrorMessage}", true);
            return result;
        }
        catch (Exception ex)
        {
            sw.Stop();
            result.IsSuccess = false;
            result.ErrorMessage = ex.Message;
            onLogReceived($"[ERRO] Falha ao executar o motor analítico: {ex.Message}", true);
            return result;
        }
    }
}
