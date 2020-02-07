using System.Diagnostics.CodeAnalysis;

namespace SlideXML.Enums
{
    /// <summary>
    /// Enumerations of shape type.
    /// </summary>
    [SuppressMessage("ReSharper", "InconsistentNaming")]
    public enum ElementType
    {
        Chart,
        Group,
        Picture,
        AutoShape,
        Table,
        OLEObject,
        Placeholder
    }
}
