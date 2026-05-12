namespace OpenXml.Parsers.Xlsx;

public sealed record XlsxFillV3(
    uint Index,
    string? PatternType,
    XlsxColorV3? ForegroundColor,
    XlsxColorV3? BackgroundColor);
