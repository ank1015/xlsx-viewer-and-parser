namespace OpenXml.Parsers.Xlsx;

public sealed record XlsxSourceInfoV1(
    string FileName,
    string FullPath,
    long FileSizeBytes);
