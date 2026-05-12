namespace OpenXml.Parsers.Xlsx;

public sealed record XlsxAutoFilterV5(
    string? Reference,
    IReadOnlyList<XlsxFilterColumnV5> FilterColumns);
