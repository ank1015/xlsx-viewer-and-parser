namespace OpenXml.Parsers.Xlsx;

public sealed record XlsxChartV6(
    string Id,
    string RelationshipId,
    string? Type,
    string? BarDirection,
    string? BarGrouping,
    string? Title,
    XlsxDrawingAnchorV6? Anchor,
    IReadOnlyList<XlsxChartSeriesV6> Series,
    string? AssetPath);
