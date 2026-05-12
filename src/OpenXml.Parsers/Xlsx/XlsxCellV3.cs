namespace OpenXml.Parsers.Xlsx;

public sealed record XlsxCellV3(
    string Address,
    uint Row,
    int Column,
    XlsxCellValueKind Type,
    string? RawValue,
    object? Value,
    string DisplayValue,
    uint? StyleIndex,
    XlsxCellFormatV3? Style,
    XlsxFormulaV2? Formula);
