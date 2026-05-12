namespace OpenXml.Parsers.Xlsx;

public sealed record XlsxSheetV2(
    string Id,
    string Name,
    int Index,
    string? Dimension,
    IReadOnlyDictionary<string, XlsxCellV2> Cells);
