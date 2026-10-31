using SheetBuilder.Models;
using SheetBuilder.Services;

namespace SheetBuilder.Controllers;

public class EnvironmentController
{
    private readonly PythonRunnerService _pythonRunnerService;

    public EnvironmentController(PythonRunnerService pythonRunnerService)
    {
        _pythonRunnerService = pythonRunnerService;
    }

    public async Task<PythonEnvironmentInfo> GetEnvironmentStatusAsync()
    {
        return await _pythonRunnerService.GetEnvironmentInfoAsync();
    }

    public async Task<bool> InstallPackageAsync(string packageName, Action<string> onLog)
    {
        return await _pythonRunnerService.InstallPackageAsync(packageName, onLog);
    }
}
