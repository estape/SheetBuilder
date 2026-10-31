namespace SheetBuilder.Models;

public class AppSettings
{
    public static string DefaultSaveFolder => 
        Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);

    public string PythonPath { get; set; } = "python";
    public string LastSaveFolder { get; set; } = DefaultSaveFolder;
    public bool AutoOpenExcel { get; set; } = false;
    public string Theme { get; set; } = "system"; // system, light, dark
}
