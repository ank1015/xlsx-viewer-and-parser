namespace OpenXml.Parsers.Xlsx;

public sealed record XlsxRowV4(
    uint Index,
    double? Height,
    bool CustomHeight,
    bool Hidden,
    uint? StyleIndex,
    byte? OutlineLevel,
    bool Collapsed);
