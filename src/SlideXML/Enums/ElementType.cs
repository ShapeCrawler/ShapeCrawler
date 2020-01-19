using System.Diagnostics.CodeAnalysis;

namespace SlideXML.Enums
{
    /// <summary>
    /// Represents the type of a element.
    /// </summary>
    [SuppressMessage("ReSharper", "InconsistentNaming")]
    public enum ElementType
    {
        Chart,
        Group,
        Picture,
        Shape,
        Table,
        OLEObject
    }
}
