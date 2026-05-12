namespace OpenXml.Parsers.Xlsx;

public sealed record XlsxFontV3(
    uint Index,
    string? Name,
    double? Size,
    bool Bold,
    bool Italic,
    bool Underline,
    bool Strike,
    XlsxColorV3? Color);
