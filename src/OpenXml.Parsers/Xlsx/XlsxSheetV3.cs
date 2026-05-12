namespace OpenXml.Parsers.Xlsx;

public sealed record XlsxSheetV3(
    string Id,
    string Name,
    int Index,
    string? Dimension,
    IReadOnlyDictionary<string, XlsxCellV3> Cells);
