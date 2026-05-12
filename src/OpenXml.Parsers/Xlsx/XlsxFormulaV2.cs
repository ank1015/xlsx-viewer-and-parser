namespace OpenXml.Parsers.Xlsx;

public sealed record XlsxFormulaV2(
    string? Text,
    string? ResolvedText,
    XlsxFormulaKind Kind,
    uint? SharedIndex,
    string? Reference,
    string? BaseAddress,
    string? BaseText,
    string? CachedRawValue,
    object? CachedValue,
    string CachedDisplayValue,
    XlsxCellValueKind CachedValueType);
