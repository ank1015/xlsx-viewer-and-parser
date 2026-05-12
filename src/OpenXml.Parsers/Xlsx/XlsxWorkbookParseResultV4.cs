namespace OpenXml.Parsers.Xlsx;

public sealed record XlsxWorkbookParseResultV4(
    string SchemaVersion,
    XlsxSourceInfoV1 Source,
    XlsxWorkbookV4 Workbook,
    IReadOnlyList<string> Warnings)
{
    public const string CurrentSchemaVersion = "xlsx-workbook/v4";
}
