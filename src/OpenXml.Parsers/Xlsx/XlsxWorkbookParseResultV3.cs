namespace OpenXml.Parsers.Xlsx;

public sealed record XlsxWorkbookParseResultV3(
    string SchemaVersion,
    XlsxSourceInfoV1 Source,
    XlsxWorkbookV3 Workbook,
    IReadOnlyList<string> Warnings)
{
    public const string CurrentSchemaVersion = "xlsx-workbook/v3";
}
