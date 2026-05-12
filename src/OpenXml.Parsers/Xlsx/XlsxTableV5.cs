namespace OpenXml.Parsers.Xlsx;

public sealed record XlsxTableV5(
    string RelationshipId,
    uint? Id,
    string? Name,
    string? DisplayName,
    string? Reference,
    string? StartAddress,
    string? EndAddress,
    uint? StartRow,
    int? StartColumn,
    uint? EndRow,
    int? EndColumn,
    uint HeaderRowCount,
    uint TotalsRowCount,
    bool InsertRow,
    bool Published,
    IReadOnlyList<XlsxTableColumnV5> Columns,
    XlsxAutoFilterV5? AutoFilter,
    XlsxTableStyleV5? Style);
