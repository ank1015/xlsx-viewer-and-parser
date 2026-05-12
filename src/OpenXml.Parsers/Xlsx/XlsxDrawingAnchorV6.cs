namespace OpenXml.Parsers.Xlsx;

public sealed record XlsxDrawingAnchorV6(
    string Kind,
    XlsxDrawingAnchorMarkerV6? From,
    XlsxDrawingAnchorMarkerV6? To,
    long? XEmu,
    long? YEmu,
    long? CxEmu,
    long? CyEmu);
