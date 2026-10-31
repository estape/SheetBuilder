using System.Text.Json;
using Photino.NET;
using SheetBuilder.Models;
using SheetBuilder.Services;

namespace SheetBuilder.Controllers;

public class MessageControllerRouter
{
    private readonly SheetController _sheetController;
    private readonly DialogController _dialogController;
    private readonly EnvironmentController _envController;

    private static readonly JsonSerializerOptions JsonOpts = new()
    {
        PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
        WriteIndented = false
    };

    public MessageControllerRouter()
    {
        var pythonService = new PythonRunnerService();
        var excelService = new ExcelService();

        _sheetController = new SheetController(pythonService, excelService);
        _dialogController = new DialogController();
        _envController = new EnvironmentController(pythonService);
    }

    public async Task HandleMessageAsync(string rawMessage, PhotinoWindow window)
    {
        string? action = null;
        string? msgId = null;

        try
        {
            using var doc = JsonDocument.Parse(rawMessage);
            var root = doc.RootElement;

            action = root.GetProperty("action").GetString();
            if (root.TryGetProperty("id", out var idProp))
                msgId = idProp.GetString();

            var payload = root.TryGetProperty("payload", out var pProp) ? pProp : default;

            switch (action)
            {
                case "get_env":
                {
                    var env = await _envController.GetEnvironmentStatusAsync();
                    SendResponse(window, action, msgId, true, env);
                    break;
                }

                case "install_package":
                {
                    var pkg = payload.GetProperty("package").GetString() ?? "";
                    var ok = await _envController.InstallPackageAsync(pkg, log =>
                    {
                        SendEvent(window, "log", msgId, new { message = log, isError = false });
                    });
                    var env = await _envController.GetEnvironmentStatusAsync();
                    SendResponse(window, action, msgId, ok, new { success = ok, env });
                    break;
                }

                case "execute_script":
                {
                    var code = payload.GetProperty("code").GetString() ?? "";
                    var timeout = payload.TryGetProperty("timeout", out var tProp) ? tProp.GetInt32() : 90;

                    var envVars = new Dictionary<string, string>();
                    if (payload.TryGetProperty("envVars", out var envProp) && envProp.ValueKind == JsonValueKind.Object)
                    {
                        foreach (var prop in envProp.EnumerateObject())
                        {
                            envVars[prop.Name] = prop.Value.GetString() ?? "";
                        }
                    }

                    var req = new ScriptExecutionRequest
                    {
                        ScriptCode = code,
                        TimeoutSeconds = timeout,
                        EnvironmentVariables = envVars
                    };

                    var execResult = await _sheetController.ExecuteScriptAsync(req, (line, isError) =>
                    {
                        SendEvent(window, "log", msgId, new { message = line, isError });
                    });

                    SendResponse(window, action, msgId, execResult.IsSuccess, execResult);
                    break;
                }

                case "get_sheet_preview":
                {
                    var filePath = payload.GetProperty("filePath").GetString() ?? _sheetController.CurrentExcelFilePath ?? "";
                    var sheetName = payload.TryGetProperty("sheetName", out var sProp) ? sProp.GetString() : null;

                    var preview = _sheetController.GetSheetPreview(filePath, sheetName);
                    SendResponse(window, action, msgId, preview != null, preview);
                    break;
                }

                case "save_excel_dialog":
                {
                    var filePath = payload.TryGetProperty("filePath", out var fProp) 
                        ? fProp.GetString() 
                        : _sheetController.CurrentExcelFilePath;

                    if (string.IsNullOrWhiteSpace(filePath) || !File.Exists(filePath))
                    {
                        SendResponse(window, action, msgId, false, new { error = "No Excel file exists to save yet." });
                        return;
                    }

                    var savedPath = await _dialogController.SaveExcelFileDialogAsync(filePath);
                    SendResponse(window, action, msgId, savedPath != null, new { savedPath, cancelled = savedPath == null });
                    break;
                }

                case "open_in_excel":
                {
                    var filePath = payload.TryGetProperty("filePath", out var fProp) 
                        ? fProp.GetString() 
                        : _sheetController.CurrentExcelFilePath;

                    if (string.IsNullOrWhiteSpace(filePath) || !File.Exists(filePath))
                    {
                        SendResponse(window, action, msgId, false, new { error = "No Excel file available to open." });
                        return;
                    }

                    var opened = _dialogController.OpenInExcel(filePath);
                    SendResponse(window, action, msgId, opened, new { opened, filePath });
                    break;
                }

                case "open_folder":
                {
                    var filePath = payload.TryGetProperty("filePath", out var fProp) 
                        ? fProp.GetString() 
                        : _sheetController.CurrentExcelFilePath;

                    if (string.IsNullOrWhiteSpace(filePath) || !File.Exists(filePath))
                    {
                        SendResponse(window, action, msgId, false, new { error = "No file available." });
                        return;
                    }

                    var opened = _dialogController.OpenContainingFolder(filePath);
                    SendResponse(window, action, msgId, opened, new { opened, filePath });
                    break;
                }

                default:
                    SendResponse(window, action ?? "unknown", msgId, false, new { error = $"Unrecognized action: {action}" });
                    break;
            }
        }
        catch (Exception ex)
        {
            SendResponse(window, action ?? "error", msgId, false, new { error = ex.Message, stack = ex.StackTrace });
        }
    }

    private static void SendResponse(PhotinoWindow window, string action, string? id, bool success, object? payload)
    {
        var msg = JsonSerializer.Serialize(new
        {
            action,
            id,
            success,
            payload
        }, JsonOpts);

        window.SendWebMessage(msg);
    }

    private static void SendEvent(PhotinoWindow window, string eventName, string? id, object payload)
    {
        var msg = JsonSerializer.Serialize(new
        {
            @event = eventName,
            id,
            payload
        }, JsonOpts);

        window.SendWebMessage(msg);
    }
}
