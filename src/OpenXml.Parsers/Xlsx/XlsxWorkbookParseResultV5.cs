namespace OpenXml.Parsers.Xlsx;

public sealed record XlsxWorkbookParseResultV5(
    string SchemaVersion,
    XlsxSourceInfoV1 Source,
    XlsxWorkbookV5 Workbook,
    IReadOnlyList<string> Warnings)
{
    public const string CurrentSchemaVersion = "xlsx-workbook/v5";
}
