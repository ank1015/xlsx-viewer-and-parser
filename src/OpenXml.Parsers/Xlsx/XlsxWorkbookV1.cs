namespace OpenXml.Parsers.Xlsx;

public sealed record XlsxWorkbookV1(
    IReadOnlyList<XlsxSheetV1> Sheets);
