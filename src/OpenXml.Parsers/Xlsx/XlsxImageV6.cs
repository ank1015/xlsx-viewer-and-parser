namespace OpenXml.Parsers.Xlsx;

public sealed record XlsxImageV6(
    string Id,
    string RelationshipId,
    string? Name,
    string? Description,
    string ContentType,
    string? FileName,
    string? AssetPath,
    XlsxDrawingAnchorV6? Anchor);
