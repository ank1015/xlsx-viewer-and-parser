namespace OpenXml.Parsers.Xlsx;

public sealed record XlsxWorkbookParseResultV1(
    string SchemaVersion,
    XlsxSourceInfoV1 Source,
    XlsxWorkbookV1 Workbook,
    IReadOnlyList<string> Warnings)
{
    public const string CurrentSchemaVersion = "xlsx-workbook/v1";
}
