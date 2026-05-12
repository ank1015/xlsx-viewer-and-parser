namespace OpenXml.Parsers.Xlsx;

public sealed record XlsxWorkbookV5(
    IReadOnlyList<XlsxSheetV5> Sheets,
    XlsxStylesV3 Styles);
