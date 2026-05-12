namespace OpenXml.Parsers.Xlsx;

public sealed record XlsxWorkbookV4(
    IReadOnlyList<XlsxSheetV4> Sheets,
    XlsxStylesV3 Styles);
