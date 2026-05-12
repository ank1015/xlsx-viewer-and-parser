namespace OpenXml.Parsers.Xlsx;

public sealed record XlsxSheetV6(
    string Id,
    string Name,
    int Index,
    string? Dimension,
    IReadOnlyDictionary<string, XlsxCellV3> Cells,
    IReadOnlyList<XlsxMergedCellV4> MergedCells,
    IReadOnlyList<XlsxRowV4> Rows,
    IReadOnlyList<XlsxColumnV4> Columns,
    IReadOnlyList<XlsxTableV5> Tables,
    IReadOnlyList<XlsxDrawingV6> Drawings,
    IReadOnlyList<XlsxChartV6> Charts,
    IReadOnlyList<XlsxImageV6> Images);
