namespace OpenXml.Parsers.Xlsx;

public sealed record XlsxWorkbookV2(
    IReadOnlyList<XlsxSheetV2> Sheets);
