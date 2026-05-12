namespace OpenXml.Parsers.Xlsx;

public sealed record XlsxWorkbookV3(
    IReadOnlyList<XlsxSheetV3> Sheets,
    XlsxStylesV3 Styles);
