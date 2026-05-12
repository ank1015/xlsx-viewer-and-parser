namespace OpenXml.Parsers.Xlsx;

public sealed record XlsxNumberFormatV3(
    uint Id,
    string FormatCode,
    bool IsBuiltIn,
    bool IsDateFormat);
