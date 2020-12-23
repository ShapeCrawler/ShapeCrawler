using System.Diagnostics.CodeAnalysis;

namespace ShapeCrawler.Enums
{
    /// <summary>
    /// Main shape content type.
    /// </summary>
    [SuppressMessage("ReSharper", "InconsistentNaming")]
    public enum ShapeContentType
    {
        /// <summary>
        /// Chart.
        /// </summary>
        Chart,

        /// <summary>
        /// Element which has grouped elements.
        /// </summary>
        Group,

        /// <summary>
        /// Picture.
        /// </summary>
        Picture,

        /// <summary>
        /// AutoShape.
        /// </summary>
        AutoShape,

        /// <summary>
        /// Table.
        /// </summary>
        Table,

        /// <summary>
        /// OLE Object.
        /// </summary>
        OLEObject
    }
}
