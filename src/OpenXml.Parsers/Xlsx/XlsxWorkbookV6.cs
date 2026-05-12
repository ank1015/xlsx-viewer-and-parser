namespace OpenXml.Parsers.Xlsx;

public sealed record XlsxWorkbookV6(
    IReadOnlyList<XlsxSheetV6> Sheets,
    XlsxStylesV3 Styles);
