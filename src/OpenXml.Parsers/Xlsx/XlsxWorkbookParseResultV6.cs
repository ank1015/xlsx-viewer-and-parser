namespace OpenXml.Parsers.Xlsx;

public sealed record XlsxWorkbookParseResultV6(
    string SchemaVersion,
    XlsxSourceInfoV1 Source,
    XlsxWorkbookV6 Workbook,
    IReadOnlyList<string> Warnings)
{
    public const string CurrentSchemaVersion = "xlsx-workbook/v6";
}
