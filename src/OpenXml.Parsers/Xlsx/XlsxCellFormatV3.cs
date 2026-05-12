namespace OpenXml.Parsers.Xlsx;

public sealed record XlsxCellFormatV3(
    uint Index,
    uint? NumberFormatId,
    string? NumberFormatCode,
    bool IsDateFormat,
    uint? FontId,
    uint? FillId,
    uint? BorderId,
    bool ApplyNumberFormat,
    bool ApplyFont,
    bool ApplyFill,
    bool ApplyBorder,
    bool ApplyAlignment,
    bool WrapText);
