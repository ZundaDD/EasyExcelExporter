using MiniExcelLibs;
using Newtonsoft.Json;
using System.Text;

namespace ExcelExporter;

public class Program
{
    private static readonly string config_path;

    static Program()
    {
        config_path = Path.Combine(AppContext.BaseDirectory, "config.json");
    }

    private static List<string> enumCheck = ["int", "float", "bool", "string"];

    private static Config config = null!;

    private static void ReadConfig() => config = JsonConvert.DeserializeObject<Config>(File.ReadAllText(config_path)) ?? new();

    private static void SaveConfig() => File.WriteAllText(config_path, JsonConvert.SerializeObject(config, Formatting.Indented));

    private static void WriteC(string str, ConsoleColor color)
    {
        var color_c = Console.ForegroundColor;
        Console.ForegroundColor = color;
        Console.Write(str);
        Console.ForegroundColor = color_c;
    }

    public static void Main(string[] args)
    {
        if (Path.Exists(config_path)) ReadConfig();
        if (config == null) config = new();

        //扫描新xlsx
        foreach (var file in Directory.GetFiles("."))
        {
            if (!file.EndsWith(".xlsx") || file.StartsWith("~$")) continue;

            var standard_key = Path.GetFileNameWithoutExtension(file);
            WriteC($"New xlsx scanned: {standard_key}.xlsx", ConsoleColor.Green);
            if (!config.Items.ContainsKey(standard_key)) config.Items[standard_key] = new();
        }

        SaveConfig();

        if (args.Length < 1)
        {
            WriteC("Press Any Key to Exit", ConsoleColor.DarkBlue);
            Console.ReadKey();
            return;
        }

        if (!config.Items.ContainsKey(args[0]))
        {
            WriteC($"No excel {args[0]},please check file name\n", ConsoleColor.Red);
            Console.Write("Press Enter to Exit:");
            Console.ReadLine();
            return;
        }

        var filePath = $"{args[0]}.xlsx";

        WriteC("[Exporting scripts]\n", ConsoleColor.Green);
        ExportScript(config.Items[args[0]], filePath);

        WriteC("[Exporting jsons]\n", ConsoleColor.Blue);
        ExportResource(config.Items[args[0]], filePath);

        WriteC("Press Any Key to Exit", ConsoleColor.DarkBlue);
        Console.ReadKey();
    }

    public static void ExportResource(ConfigItem config, string filePath)
    {
        var fileName = Path.GetFileNameWithoutExtension(filePath);

        // 覆盖路径
        var exportPath = Path.Combine(config.ResourceExportPath, fileName);
        if (Path.Exists(exportPath)) Directory.Delete(exportPath, true);
        Directory.CreateDirectory(exportPath);

        ExportExcelRows(filePath, exportPath);
    }


    public static void ExportScript(ConfigItem config, string filePath)
    {
        var exportPath = $"{config.ScriptExportPath}";
        if (!Path.Exists(exportPath)) Directory.CreateDirectory(exportPath);

        var structure = ReadStructure(filePath);
        var fileName = Path.GetFileNameWithoutExtension(filePath);

        var runtimeCode = GenerateRuntimeCode(fileName, config.Namespace, structure, config.IncludeLists.ToList());
        var runtimePath = Path.Combine(config.ScriptExportPath, $"{fileName}.cs");
        SaveCode(runtimePath, runtimeCode);
    }

    public static void ExportExcelRows(string path, string outputDir)
    {
        using var stream = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);

        var rows = stream.Query(useHeaderRow: true);
        using var enumerator = rows.GetEnumerator();

        if (enumerator.MoveNext()) { }
        if (enumerator.MoveNext()) { }
        var types = ((IDictionary<string, object>)enumerator.Current);


        while (enumerator.MoveNext())
        {
            var rowData = (IDictionary<string, object>)enumerator.Current;
            var rowDict = new Dictionary<string, object>();
            bool isEmptyRow = true;

            foreach (var kvp in rowData)
            {
                string? val = kvp.Value?.ToString()?.Trim();

                if (!string.IsNullOrEmpty(val))
                {
                    object parsedValue = ParseValue(val, types[kvp.Key].ToString());
                    rowDict[kvp.Key] = parsedValue;
                    isEmptyRow = false;
                }
            }

            if (isEmptyRow) continue;

            if (rowDict.TryGetValue("Id", out object? ido) && ido is string id && !string.IsNullOrWhiteSpace(id))
            {
                string json = JsonConvert.SerializeObject(rowDict, Formatting.Indented);
                string outputPath = Path.Combine(outputDir, $"{id}.json");
                File.WriteAllText(outputPath, json);

                Console.Write("Generated at ");
                WriteC(outputPath, ConsoleColor.Yellow);
                Console.Write("\n");
            }
        }
    }

    private static object ParseValue(string rawValue, string type)
    {
        if (string.IsNullOrEmpty(rawValue)) return null!;

        if (type.StartsWith("List<") && type.EndsWith(">"))
        {
            string innerType = type.Substring(5, type.Length - 6);

            return rawValue.Split('|')
                           .Select(s => s.Trim())
                           .Where(s => !string.IsNullOrEmpty(s))
                           .Select(s => ParseValue(s, innerType))
                           .ToList();
        }

        switch (type.ToLower())
        {
            case "int":
                return int.TryParse(rawValue, out int intVal) ? intVal : 0;
            case "float":
                return float.TryParse(rawValue, out float floatVal) ? floatVal : 0f;
            case "bool":
                return bool.TryParse(rawValue, out bool boolVal) && boolVal;
            case "string":
            default:
                return rawValue;
        }
    }

    private static string GenerateRuntimeCode(string className, string @namespace, List<(string name, string type)> sources, List<string> includes)
    {
        StringBuilder sb = new();

        StringBuilder includeBuilder = new();
        foreach (var include in includes) includeBuilder.Append($"using {include};\n");

        sb.AppendLine($@"/*
此文件根据Excel文件自动生成，请不要手动修改。
This file is auto-generated from Excel. Do not modify manually.
*/
{includeBuilder}

namespace {@namespace};

public partial class {className}
{{");

        foreach (var col in sources)
        {
            var colName = col.name;
            if(!enumCheck.Contains( col.type) && !col.type.StartsWith("List<"))
                sb.AppendLine("\t[JsonConverter(typeof(StringEnumConverter))]\n");
            sb.AppendLine($"\tpublic {col.type} {colName} {{ get; set; }}\n");
        }

        sb.Append("}");
        return sb.ToString();
    }

    private static void SaveCode(string filePath, string code)
    {
        File.WriteAllText(filePath, code);

        Console.Write("Generated at ");
        WriteC(filePath, ConsoleColor.Yellow);
        Console.Write("\n");
    }

    private static List<(string name, string type)> ReadStructure(string filePath)
    {
        var result = new List<(string Name, string Type)>();
        try
        {
            using var stream = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);

            var rows = stream.Query(useHeaderRow: false);
            using var enumerator = rows.GetEnumerator();

            if (!enumerator.MoveNext()) return null!;

            var row1Data = enumerator.Current as IDictionary<string, object>;
            if (row1Data == null) return null!;

            var fieldNames = row1Data.Values
                .Select(v => v?.ToString()?.Trim())
                .ToList();

            if (!enumerator.MoveNext()) return null!;
            if (!enumerator.MoveNext()) return null!;

            var row2Data = enumerator.Current as IDictionary<string, object>;
            if (row2Data == null) return null!;

            var fieldTypes = row2Data.Values
                .Select(v => v?.ToString()?.Trim())
                .ToList();

            for (int i = 0; i < fieldNames.Count; i++)
            {
                string name = fieldNames[i];
                string type = i < fieldTypes.Count ? fieldTypes[i] : null;

                if (string.IsNullOrEmpty(name) || string.IsNullOrEmpty(type)) continue;

                result.Add((name, type));
            }
        }
        catch (Exception ex) { WriteC($"Excel Read Error: {ex.Message}\n", ConsoleColor.Red); }

        return result;
    }
}

