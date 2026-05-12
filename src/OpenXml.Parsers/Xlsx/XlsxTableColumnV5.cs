namespace OpenXml.Parsers.Xlsx;

public sealed record XlsxTableColumnV5(
    uint Id,
    string Name,
    string? TotalsRowLabel,
    string? TotalsRowFunction,
    string? CalculatedColumnFormula,
    string? TotalsRowFormula);
