namespace OpenXml.Parsers.Xlsx;

public sealed record XlsxCellV1(
    string Address,
    uint Row,
    int Column,
    XlsxCellValueKind Type,
    string? RawValue,
    object? Value,
    string DisplayValue);
