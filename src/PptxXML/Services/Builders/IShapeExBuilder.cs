using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using PptxXML.Models.Elements;
using PptxXML.Models.Settings;

namespace PptxXML.Services.Builders
{
    /// <summary>
    /// Defines method to create <see cref="ShapeEx"/> instance.
    /// </summary>
    public interface IShapeExBuilder
    {
        ShapeEx Build(OpenXmlCompositeElement compositeElement, SlidePart sldPart, ShapeSettings spSettings);
    }
}
