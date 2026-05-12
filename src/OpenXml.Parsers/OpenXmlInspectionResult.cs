namespace OpenXml.Parsers;

public sealed record OpenXmlInspectionResult(
    string FilePath,
    OpenXmlFileKind Kind,
    IReadOnlyDictionary<string, string> Summary,
    IReadOnlyList<string> PreviewRows);
