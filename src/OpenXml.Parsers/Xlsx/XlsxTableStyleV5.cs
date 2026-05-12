namespace OpenXml.Parsers.Xlsx;

public sealed record XlsxTableStyleV5(
    string? Name,
    bool ShowFirstColumn,
    bool ShowLastColumn,
    bool ShowRowStripes,
    bool ShowColumnStripes);
