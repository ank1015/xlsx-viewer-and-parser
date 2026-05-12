namespace OpenXml.Parsers.Xlsx;

public sealed record XlsxColorV3(
    string? Rgb,
    uint? Indexed,
    uint? Theme,
    double? Tint,
    bool Auto,
    string? ResolvedRgb);
