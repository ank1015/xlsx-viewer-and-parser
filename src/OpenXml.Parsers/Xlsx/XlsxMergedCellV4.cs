namespace OpenXml.Parsers.Xlsx;

public sealed record XlsxMergedCellV4(
    string Reference,
    string StartAddress,
    string EndAddress,
    uint StartRow,
    int StartColumn,
    uint EndRow,
    int EndColumn);
