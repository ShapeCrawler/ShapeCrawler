using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using SlideXML.Models.Elements;
using SlideXML.Models.Settings;

namespace SlideXML.Services
{
    /// <summary>
    /// Provides APIs to create shape tree's elements.
    /// </summary>
    public interface IElementFactory
    {
        ShapeSL CreateShape(ElementCandidate ec, IPreSettings preSettings);

        ShapeSL CreateGroupShape(OpenXmlCompositeElement compositeElement, IPreSettings preSettings);
    }
}