namespace OpenXml.Parsers.Xlsx;

public sealed record XlsxChartSeriesV6(
    string? Name,
    string? CategoriesRange,
    string? ValuesRange,
    IReadOnlyList<string> CachedCategories,
    IReadOnlyList<double> CachedValues,
    XlsxColorV3? Color);
