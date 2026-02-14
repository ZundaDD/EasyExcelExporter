namespace ExcelExporter;

[Serializable]
public class ConfigItem
{
    public string Namespace { get; set; } = "AutoGen";

    public string ScriptExportPath { get; set; } = "..\\..\\Generated\\Script";

    public string ResourceExportPath { get; set; } = "..\\..\\Generated\\Resource";

    public HashSet<string> IncludeLists { get; set; } = ["System.Collections.Generic"];
}

[Serializable]
public class Config
{
    public Dictionary<string, ConfigItem> Items = new();
}
