namespace OpenXml.Parsers.Xlsx;

public sealed record XlsxDrawingAnchorMarkerV6(
    int Column,
    long ColumnOffsetEmu,
    uint Row,
    long RowOffsetEmu);
