#pragma warning disable IDE0130
namespace ShapeCrawler;

/// <summary>
///     Office document types supported for embedding as OLE objects.
/// </summary>
public enum OleObjectType
{
    /// <summary>
    ///     A Word document (.docx).
    /// </summary>
    WordDocument,

    /// <summary>
    ///     An Excel workbook (.xlsx).
    /// </summary>
    ExcelWorkbook,

    /// <summary>
    ///     A PowerPoint presentation (.pptx).
    /// </summary>
    PowerPointPresentation
}
