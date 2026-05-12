namespace OpenXml.Parsers.Xlsx;

public sealed record XlsxBorderV3(
    uint Index,
    XlsxBorderSideV3 Left,
    XlsxBorderSideV3 Right,
    XlsxBorderSideV3 Top,
    XlsxBorderSideV3 Bottom,
    XlsxBorderSideV3 Diagonal);
