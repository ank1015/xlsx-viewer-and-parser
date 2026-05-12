namespace OpenXml.Parsers.Xlsx;

public sealed record XlsxFilterColumnV5(
    uint ColumnId,
    bool HiddenButton,
    bool ShowButton,
    IReadOnlyList<string> FilterTypes);
