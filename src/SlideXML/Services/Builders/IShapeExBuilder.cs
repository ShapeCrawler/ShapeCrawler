using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using SlideXML.Models.Elements;
using SlideXML.Models.Settings;

namespace SlideXML.Services.Builders
{
    /// <summary>
    /// Defines method to create <see cref="ShapeEx"/> instance.
    /// </summary>
    public interface IShapeExBuilder
    {
        ShapeEx Build(OpenXmlCompositeElement compositeElement, SlidePart sldPart, ElementSettings spSettings);
    }
}
