using System.Diagnostics.CodeAnalysis;

namespace SlideXML.Enums
{
    /// <summary>
    /// Element type.
    /// </summary>
    [SuppressMessage("ReSharper", "InconsistentNaming")]
    public enum ElementType
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
