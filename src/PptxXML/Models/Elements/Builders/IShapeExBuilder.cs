using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;

namespace PptxXML.Models.Elements.Builders
{
    /// <summary>
    /// Defines method to create <see cref="ShapeEx"/> instance.
    /// </summary>
    public interface IShapeExBuilder
    {
        ShapeEx Build(OpenXmlCompositeElement compositeElement, SlidePart sldPart);
    }
}
