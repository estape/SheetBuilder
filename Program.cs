using System.Drawing;
using Photino.NET;
using SheetBuilder.Controllers;

namespace SheetBuilder;

internal class Program
{
    [STAThread]
    static void Main(string[] args)
    {
        var router = new MessageControllerRouter();

        var window = new PhotinoWindow()
            .SetTitle("SheetBuilder — Automação de Planilhas Financeiras (Google AI Studio)")
            .SetUseOsDefaultSize(false)
            .SetSize(1320, 860)
            .SetMinSize(960, 620)
            .Center()
            .SetResizable(true)
#if DEBUG
            .SetDevToolsEnabled(true)
            .SetContextMenuEnabled(true)
#else
            .SetDevToolsEnabled(false)
            .SetContextMenuEnabled(false)
#endif
            .RegisterWebMessageReceivedHandler((sender, rawMessage) =>
            {
                if (sender is not PhotinoWindow targetWindow) return;

                // Offload to background task to keep native OS UI thread responsive
                _ = Task.Run(async () =>
                {
                    await router.HandleMessageAsync(rawMessage, targetWindow);
                });
            })
            .Load("wwwroot/index.html");

        window.WaitForClose();
    }
}
