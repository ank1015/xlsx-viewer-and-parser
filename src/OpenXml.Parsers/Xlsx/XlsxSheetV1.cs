namespace OpenXml.Parsers.Xlsx;

public sealed record XlsxSheetV1(
    string Id,
    string Name,
    int Index,
    string? Dimension,
    IReadOnlyDictionary<string, XlsxCellV1> Cells);
