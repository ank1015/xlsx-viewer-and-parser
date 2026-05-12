namespace OpenXml.Parsers.Xlsx;

public sealed record XlsxColumnV4(
    uint Min,
    uint Max,
    double? Width,
    bool CustomWidth,
    bool BestFit,
    bool Hidden,
    uint? StyleIndex,
    byte? OutlineLevel,
    bool Collapsed);
