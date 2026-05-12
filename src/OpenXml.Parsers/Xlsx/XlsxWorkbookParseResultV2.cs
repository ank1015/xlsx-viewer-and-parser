namespace OpenXml.Parsers.Xlsx;

public sealed record XlsxWorkbookParseResultV2(
    string SchemaVersion,
    XlsxSourceInfoV1 Source,
    XlsxWorkbookV2 Workbook,
    IReadOnlyList<string> Warnings)
{
    public const string CurrentSchemaVersion = "xlsx-workbook/v2";
}
