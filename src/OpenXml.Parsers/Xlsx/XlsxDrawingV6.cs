namespace OpenXml.Parsers.Xlsx;

public sealed record XlsxDrawingV6(
    string RelationshipId,
    IReadOnlyList<XlsxChartV6> Charts,
    IReadOnlyList<XlsxImageV6> Images);
