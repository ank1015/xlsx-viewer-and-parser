namespace OpenXml.Parsers.Xlsx;

public sealed record XlsxStylesV3(
    IReadOnlyList<XlsxNumberFormatV3> NumberFormats,
    IReadOnlyList<XlsxThemeColorV3> ThemeColors,
    IReadOnlyList<XlsxFontV3> Fonts,
    IReadOnlyList<XlsxFillV3> Fills,
    IReadOnlyList<XlsxBorderV3> Borders,
    IReadOnlyList<XlsxCellFormatV3> CellFormats);
