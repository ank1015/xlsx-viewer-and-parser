using OpenXml.Parsers;
using OpenXml.Parsers.Xlsx;
using System.Text.Json;
using System.Text.Json.Serialization;

if (args.Length == 0 || args[0].Equals("--help", StringComparison.OrdinalIgnoreCase) || args[0].Equals("-h", StringComparison.OrdinalIgnoreCase))
{
    PrintUsage();
    return 0;
}

try
{
    if (args[0].Equals("xlsx", StringComparison.OrdinalIgnoreCase))
    {
        if (args.Length < 2)
        {
            PrintUsage();
            return 1;
        }

        var result = XlsxWorkbookV6Parser.Parse(args[1]);
        Console.WriteLine(JsonSerializer.Serialize(result, JsonOptions()));
        return 0;
    }

    if (args[0].Equals("xlsx-assets", StringComparison.OrdinalIgnoreCase))
    {
        if (args.Length < 3)
        {
            PrintUsage();
            return 1;
        }

        Directory.CreateDirectory(args[2]);
        var result = XlsxWorkbookV6Parser.Parse(args[1], args[2]);
        var workbookJsonPath = Path.Combine(args[2], "workbook.json");
        File.WriteAllText(workbookJsonPath, JsonSerializer.Serialize(result, JsonOptions()));
        Console.WriteLine(JsonSerializer.Serialize(new
        {
            schemaVersion = result.SchemaVersion,
            workbookJsonPath,
            charts = result.Workbook.Sheets.Sum(sheet => sheet.Charts.Count),
            images = result.Workbook.Sheets.Sum(sheet => sheet.Images.Count),
            warnings = result.Warnings.Count
        }, JsonOptions()));
        return 0;
    }

    if (args[0].Equals("xlsx-v5", StringComparison.OrdinalIgnoreCase))
    {
        if (args.Length < 2)
        {
            PrintUsage();
            return 1;
        }

        var result = XlsxWorkbookV5Parser.Parse(args[1]);
        Console.WriteLine(JsonSerializer.Serialize(result, JsonOptions()));
        return 0;
    }

    if (args[0].Equals("xlsx-v4", StringComparison.OrdinalIgnoreCase))
    {
        if (args.Length < 2)
        {
            PrintUsage();
            return 1;
        }

        var result = XlsxWorkbookV4Parser.Parse(args[1]);
        Console.WriteLine(JsonSerializer.Serialize(result, JsonOptions()));
        return 0;
    }

    if (args[0].Equals("xlsx-v3", StringComparison.OrdinalIgnoreCase))
    {
        if (args.Length < 2)
        {
            PrintUsage();
            return 1;
        }

        var result = XlsxWorkbookV3Parser.Parse(args[1]);
        Console.WriteLine(JsonSerializer.Serialize(result, JsonOptions()));
        return 0;
    }

    if (args[0].Equals("xlsx-v2", StringComparison.OrdinalIgnoreCase))
    {
        if (args.Length < 2)
        {
            PrintUsage();
            return 1;
        }

        var result = XlsxWorkbookV2Parser.Parse(args[1]);
        Console.WriteLine(JsonSerializer.Serialize(result, JsonOptions()));
        return 0;
    }

    if (args[0].Equals("xlsx-v1", StringComparison.OrdinalIgnoreCase))
    {
        if (args.Length < 2)
        {
            PrintUsage();
            return 1;
        }

        var result = XlsxWorkbookV1Parser.Parse(args[1]);
        Console.WriteLine(JsonSerializer.Serialize(result, JsonOptions()));
        return 0;
    }

    if (args[0].Equals("inspect", StringComparison.OrdinalIgnoreCase))
    {
        if (args.Length < 2)
        {
            PrintUsage();
            return 1;
        }

        return PrintInspection(args[1]);
    }

    if (args.Length == 1)
    {
        return PrintInspection(args[0]);
    }

    PrintUsage();
    return 1;
}
catch (Exception ex)
{
    Console.Error.WriteLine($"{ex.GetType().Name}: {ex.Message}");
    return 1;
}

static int PrintInspection(string filePath)
{
    var result = OpenXmlInspector.Inspect(filePath);

    Console.WriteLine($"File: {result.FilePath}");
    Console.WriteLine($"Kind: {result.Kind}");
    Console.WriteLine();

    Console.WriteLine("Summary");
    foreach (var item in result.Summary)
    {
        Console.WriteLine($"- {item.Key}: {item.Value}");
    }

    if (result.PreviewRows.Count > 0)
    {
        Console.WriteLine();
        Console.WriteLine("Preview");
        foreach (var row in result.PreviewRows)
        {
            Console.WriteLine($"- {row}");
        }
    }

    return result.Kind == OpenXmlFileKind.Unknown ? 2 : 0;
}

static JsonSerializerOptions JsonOptions()
{
    var options = new JsonSerializerOptions(JsonSerializerDefaults.Web)
    {
        WriteIndented = true
    };

    options.Converters.Add(new JsonStringEnumConverter(JsonNamingPolicy.CamelCase));
    return options;
}

static void PrintUsage()
{
    Console.WriteLine("Usage:");
    Console.WriteLine("  heysnap-xlsxl xlsx <file.xlsx>");
    Console.WriteLine("  heysnap-xlsxl xlsx-assets <file.xlsx> <output-dir>");
    Console.WriteLine("  heysnap-xlsxl xlsx-v5 <file.xlsx>");
    Console.WriteLine("  heysnap-xlsxl xlsx-v4 <file.xlsx>");
    Console.WriteLine("  heysnap-xlsxl xlsx-v3 <file.xlsx>");
    Console.WriteLine("  heysnap-xlsxl xlsx-v2 <file.xlsx>");
    Console.WriteLine("  heysnap-xlsxl xlsx-v1 <file.xlsx>");
    Console.WriteLine("  heysnap-xlsxl inspect <file.docx|file.xlsx|file.pptx>");
    Console.WriteLine();
    Console.WriteLine("Development:");
    Console.WriteLine("  dotnet run --project src/OpenXml.Cli -- xlsx <file.xlsx>");
}
